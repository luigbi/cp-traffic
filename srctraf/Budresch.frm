VERSION 5.00
Begin VB.Form BudResch 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4275
   ClientLeft      =   1260
   ClientTop       =   1980
   ClientWidth     =   8145
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
   ScaleHeight     =   4275
   ScaleWidth      =   8145
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Left            =   30
      Picture         =   "Budresch.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.VScrollBar vbcVehicle 
      Height          =   2670
      LargeChange     =   11
      Left            =   7575
      Min             =   1
      TabIndex        =   17
      Top             =   315
      Value           =   1
      Width           =   270
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   75
      Left            =   7860
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   15
      Width           =   45
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
      Left            =   750
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1095
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.ListBox lbcDemo 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   750
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1335
      Visible         =   0   'False
      Width           =   1365
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
      Left            =   1695
      Picture         =   "Budresch.frx":030A
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1095
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcVehicle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   180
      Picture         =   "Budresch.frx":0404
      ScaleHeight     =   2730
      ScaleWidth      =   7365
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   300
      Width           =   7365
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   15
         Top             =   345
         Visible         =   0   'False
         Width           =   7365
      End
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
      Height          =   60
      Left            =   1320
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   3
      Top             =   3705
      Width           =   120
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
      Left            =   750
      ScaleHeight     =   90
      ScaleWidth      =   135
      TabIndex        =   9
      Top             =   3765
      Width           =   135
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6810
      TabIndex        =   14
      Top             =   3750
      Visible         =   0   'False
      Width           =   1035
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
      Height          =   3435
      Left            =   150
      ScaleHeight     =   3375
      ScaleWidth      =   7710
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   7770
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
         TabIndex        =   12
         Top             =   3000
         Width           =   795
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   5340
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0FFFF&
         X1              =   120
         X2              =   5340
         Y1              =   2865
         Y2              =   2865
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
         TabIndex        =   11
         Top             =   3045
         Width           =   1560
      End
      Begin VB.Label lacTotal 
         Alignment       =   1  'Right Justify
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
         Left            =   5415
         TabIndex        =   10
         Top             =   2820
         Width           =   2010
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4455
      TabIndex        =   13
      Top             =   3825
      Width           =   945
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
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
      Height          =   285
      Left            =   2820
      TabIndex        =   2
      Top             =   3825
      Width           =   945
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   195
      Top             =   3840
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "BudResch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Budresch.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BudResch.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
Dim tmCtrls(0 To 7)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current event name Box
Dim imRowNo As Integer
'Vehicle
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
Dim tmBudResch() As BUDRESEARCH
Dim tmDemoCode() As SORTCODE
Dim smDemoCodeTag As String
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imLbcArrowSetting As Integer
Dim imBypassFocus As Integer
Dim imTabDirection As Integer
Dim imSettingValue As Integer
Dim imUpdateAllowed As Integer    'User can update records

Const VEHICLEINDEX = 1
Const DEMOINDEX = 2
Const RATEAUDINDEX = 3
Const CPPCPMINDEX = 4
Const AVAILSINDEX = 5
Const PCTSELLOUTINDEX = 6
Const DOLLARINDEX = 7
Private Sub cmcApply_Click()
    Dim slIndex As String
    Dim ilLoop As Integer
    Dim ilBvf As Integer
    Dim slVehicle As String
    Dim ilWk As Integer
    Dim slStr As String
    Dim slRound As String
    Dim llVTotal As Long
    Dim ilRet As Integer
    Dim ilCount As Integer
    Dim ilNoWks As Integer
    Dim ilYear As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    slRound = edcRound.Text
    If slRound = "" Then
        slRound = "1"
    End If
    ilRet = MsgBox("This will Replace the Budget Values, Ok to Proceed", vbYesNo + vbQuestion, "Research")
    If ilRet = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilYear = 0
    'For ilLoop = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
    For ilLoop = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
        If tgBvfRec(ilLoop).tBvf.iYear > 0 Then
            ilYear = tgBvfRec(ilLoop).tBvf.iYear
            Exit For
        End If
    Next ilLoop
    If tgSpf.sRUseCorpCal = "Y" Then
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            If tgMCof(ilLoop).iYear = ilYear Then
                'gUnpackDate tgMCof(ilLoop).iStartDate(0, 1), tgMCof(ilLoop).iStartDate(1, 1), slStartDate
                'gUnpackDate tgMCof(ilLoop).iEndDate(0, 12), tgMCof(ilLoop).iEndDate(1, 12), slEndDate
                gUnpackDate tgMCof(ilLoop).iStartDate(0, 0), tgMCof(ilLoop).iStartDate(1, 0), slStartDate
                gUnpackDate tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), slEndDate
                ilNoWks = (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 7 + 1
                Exit For
            End If
        Next ilLoop
    Else
        slStartDate = "1/15/" & Trim$(Str$(ilYear))
        slStartDate = gObtainYearStartDate(0, slStartDate)
        slEndDate = "12/15/" & Trim$(Str$(ilYear))
        slEndDate = gObtainYearEndDate(0, slEndDate)
        ilNoWks = (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 7 + 1
    End If

    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        slVehicle = lbcVehicle.List(ilLoop)
        llVTotal = 0
        'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                'For ilWk = LBound(tgBvfRec(ilBvf).tBvf.lGross) To UBound(tgBvfRec(ilBvf).tBvf.lGross) Step 1
                For ilWk = 1 To 53 Step 1
                    If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                        llVTotal = llVTotal + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                    End If
                Next ilWk
            End If
        Next ilBvf
        'gdcVehicle.Col = DOLLARINDEX
        'gdcVehicle.Row = ilLoop + 1
        'slStr = gdcVehicle.Text
        slStr = gLongToStrDec(tmBudResch(ilLoop).lDollars, 0)
        ilCount = 0
        If llVTotal > 0 Then
            slIndex = gDivStr(slStr & ".0000", gLongToStrDec(llVTotal, 0))
        Else
            slIndex = ".0000"
            'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
            For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                    ilCount = ilCount + 1
                End If
            Next ilBvf
        End If
        'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                'For ilWk = LBound(tgBvfRec(ilBvf).tBvf.lGross) To UBound(tgBvfRec(ilBvf).tBvf.lGross) Step 1
                For ilWk = 1 To ilNoWks Step 1
                    If llVTotal > 0 Then
                        slStr = Trim$(Str$(tgBvfRec(ilBvf).tBvf.lGross(ilWk)))
                        slStr = gMulStr(slStr, slIndex)
                        slStr = gRoundStr(slStr, slRound, 0)
                        tgBvfRec(ilBvf).tBvf.lGross(ilWk) = gStrDecToLong(slStr, 0)
                    Else
                        If ilCount > 0 Then
                            tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tmBudResch(ilLoop).lDollars / (ilCount * ilNoWks)
                        End If
                    End If
                Next ilWk
            End If
        Next ilBvf
        'Correct Balance if off
        llVTotal = 0
        'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                'For ilWk = LBound(tgBvfRec(ilBvf).tBvf.lGross) To UBound(tgBvfRec(ilBvf).tBvf.lGross) Step 1
                For ilWk = 1 To 53 Step 1
                    If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                        llVTotal = llVTotal + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                    End If
                Next ilWk
            End If
        Next ilBvf
        If (llVTotal <> tmBudResch(ilLoop).lDollars) And (slRound = "1") Then
            For ilWk = 1 To 53 Step 1
                'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
                For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                    If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                        If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                            If llVTotal > tmBudResch(ilLoop).lDollars Then
                                tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tgBvfRec(ilBvf).tBvf.lGross(ilWk) - 1
                                llVTotal = llVTotal - 1
                            ElseIf llVTotal < tmBudResch(ilLoop).lDollars Then
                                tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tgBvfRec(ilBvf).tBvf.lGross(ilWk) + 1
                                llVTotal = llVTotal + 1
                            End If
                            If llVTotal = tmBudResch(ilLoop).lDollars Then
                                Exit For
                            End If
                        End If
                    End If
                Next ilBvf
                If llVTotal = tmBudResch(ilLoop).lDollars Then
                    Exit For
                End If
            Next ilWk
        End If
        'Do
        '    tmVefSrchKey.iCode = tmBudResch(ilLoop).iVefCode
        '    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        '    If ilRet <> BTRV_ERR_NONE Then
        '        Exit Do
        '    End If
        '    tmVef.iMnfDemo = tmBudResch(ilLoop).iMnfDemo
        '    tmVef.lRateAud = tmBudResch(ilLoop).lRateAud
        '    tmVef.lCPPCPM = tmBudResch(ilLoop).lCPPCPM
        '    tmVef.lYearAvails = tmBudResch(ilLoop).lAvails
        '    tmVef.iPctSellout = tmBudResch(ilLoop).iPctSellout
        '    ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
        'Loop While ilRet = BTRV_ERR_CONFLICT
    Next ilLoop
    'mComputeTotal
    igBDReturn = 1
    Screen.MousePointer = vbDefault
    mTerminate
End Sub
Private Sub cmcApply_GotFocus()
    mSetShow imBoxNo
    mSaveRec
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    mSaveRec
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case DEMOINDEX  'Demo
            lbcDemo.Visible = Not lbcDemo.Visible
        Case RATEAUDINDEX  'Avg Rating or Audience
        Case CPPCPMINDEX
        Case AVAILSINDEX
        Case PCTSELLOUTINDEX
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_Change()
    Dim ilRet As Integer
    Dim slStr As String

    Select Case imBoxNo
        Case DEMOINDEX  'Demo
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcDemo, imBSMode, slStr)
            If ilRet = 1 Then
                lbcDemo.ListIndex = 0
            End If
            imLbcArrowSetting = False
        Case RATEAUDINDEX  'Avg Rating or Audience
        Case CPPCPMINDEX
        Case AVAILSINDEX
        Case PCTSELLOUTINDEX
    End Select
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case DEMOINDEX  'Demo
            'imComboBoxIndex = lbcDemo.ListIndex
        Case RATEAUDINDEX  'Avg Rating or Audience
        Case CPPCPMINDEX
        Case AVAILSINDEX
        Case PCTSELLOUTINDEX
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
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imBoxNo
        Case DEMOINDEX  'Demo
            ilKey = KeyAscii
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case RATEAUDINDEX  'Avg Rating or Audience
            ilKey = KeyAscii
            If Not mDropDownKeyPress(ilKey, False) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case CPPCPMINDEX
            ilKey = KeyAscii
            If Not mDropDownKeyPress(ilKey, False) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case AVAILSINDEX
            If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
                If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
                    imBSMode = True 'Force deletion of character prior to selected text
                End If
            End If
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case PCTSELLOUTINDEX
            If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
                If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
                    imBSMode = True 'Force deletion of character prior to selected text
                End If
            End If
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case imBoxNo
        Case DEMOINDEX  'Demo
            If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                gProcessArrowKey Shift, KeyCode, lbcDemo, imLbcArrowSetting
            End If
        Case RATEAUDINDEX  'Avg Rating or Audience
        Case CPPCPMINDEX
        Case AVAILSINDEX
        Case PCTSELLOUTINDEX
    End Select
End Sub
Private Sub edcRound_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
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
    BudResch.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmBudResch
    Erase tmDemoCode
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    
    Set BudResch = Nothing   'Remove data segment

End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcDemo_Click()
    gProcessLbcClick lbcDemo, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcDemo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mComputeDollar                  *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute Dollars for Vehicle    *
'*                                                     *
'*******************************************************
Private Function mComputeDollar(llCPPCPM As Long, llAvails As Long, llRateAud As Long, ilPctSellout As Integer) As Long
    Dim llDollars As Long
    If (CDbl(llCPPCPM) * CDbl(llAvails) * CDbl(llRateAud) * CDbl(ilPctSellout)) / CDbl(1000000) < 2000000000 Then
        llDollars = CLng((CDbl(llCPPCPM) * CDbl(llAvails) * CDbl(llRateAud) * CDbl(ilPctSellout)) / CDbl(1000000))
    Else
        MsgBox "Total dollars exceeded $2 billion, total set to zero. Please adjust values entered", vbInformation + vbOKOnly
        llDollars = 0
    End If
    mComputeDollar = llDollars
End Function
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
Private Sub mComputeTotal()
    Dim slStr As String
    Dim ilLoop As Integer
    lgTotal = 0
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        lgTotal = lgTotal + tmBudResch(ilLoop).lDollars
    Next ilLoop
    slStr = Trim$(Str$(lgTotal))
    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
    lacTotal.Caption = "(Total: " & slStr & ")"
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDemoPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Demo List box if      *
'*                      required                       *
'*                                                     *
'*******************************************************
Private Sub mDemoPop()
'
'   mDemoPop
'   Where:
'
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcDemo.ListIndex
    If ilIndex > 1 Then
        slName = lbcDemo.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilFilter(0) = CHARFILTER
    slFilter(0) = "D"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(Vehicle, lbcDemo, lbcDemoCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(BudResch, lbcDemo, tmDemoCode(), smDemoCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mDemoPopErr
        gCPErrorMsg ilRet, "mDemoPop (gIMoveListBox)", BudResch
        On Error GoTo 0
        lbcDemo.AddItem "[None]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 1, lbcDemo
            If gLastFound(lbcDemo) > 0 Then
                lbcDemo.ListIndex = gLastFound(lbcDemo)
            Else
                lbcDemo.ListIndex = -1
            End If
        Else
            lbcDemo.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mDemoPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDropDownKeyPress               *
'*                                                     *
'*             Created:5/11/94       By:D. Hannifan    *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                      in transaction section         *
'*******************************************************
Private Function mDropDownKeyPress(KeyAscii As Integer, ilNegAllowed As Integer) As Integer
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcDropDown.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                mDropDownKeyPress = False
                Exit Function
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If ilNegAllowed Then
        If (KeyAscii = KEYNEG) And ((Len(edcDropDown.Text) = 0) Or (Len(edcDropDown.Text) = edcDropDown.SelLength)) Then
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) And (KeyAscii <> KEYNEG) Then
                Beep
                mDropDownKeyPress = False
                Exit Function
            End If
        Else
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                Beep
                mDropDownKeyPress = False
                Exit Function
            End If
        End If
    Else
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
            Beep
            mDropDownKeyPress = False
            Exit Function
        End If
    End If
    slStr = edcDropDown.Text
    slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
    If gCompAbsNumberStr(slStr, "999999999.99") > 0 Then
        Beep
        mDropDownKeyPress = False
        Exit Function
    End If
    'If KeyAscii <> KEYBACKSPACE Then
    '    Select Case Chr(KeyAscii)
    '        Case "7"
    '            ilRowNo = 1
    '            ilColNo = 1
    '        Case "8"
    '            ilRowNo = 1
    '            ilColNo = 2
    '        Case "9"
    '            ilRowNo = 1
    '            ilColNo = 3
    '        Case "4"
    '            ilRowNo = 2
    '            ilColNo = 1
    '        Case "5"
    '            ilRowNo = 2
    '            ilColNo = 2
    '        Case "6"
    '            ilRowNo = 2
    '            ilColNo = 3
    '        Case "1"
    '            ilRowNo = 3
    '            ilColNo = 1
    '        Case "2"
    '            ilRowNo = 3
    '            ilColNo = 2
    '        Case "3"
    '            ilRowNo = 3
    '            ilColNo = 3
    '        Case "0"
    '            ilRowNo = 4
    '            ilColNo = 1
    '        Case "00"   'Not possible
    '            ilRowNo = 4
    '            ilColNo = 2
    '        Case "."
    '            ilRowNo = 4
    '            ilColNo = 3
    '        Case "-"
    '            ilRowNo = 0
    '    End Select
    '    If ilRowNo > 0 Then
    '        flX = fgPadMinX + (ilColNo - 1) * fgPadDeltaX
    '        flY = fgPadMinY + (ilRowNo - 1) * fgPadDeltaY
    '        imcNumOutLine.Move flX - 15, flY - 15
    '        imcNumOutLine.Visible = True
    '    Else
    '        imcNumOutLine.Visible = False
    '    End If
    'Else
    '    imcNumOutLine.Visible = False
    'End If
    mDropDownKeyPress = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    Dim ilPos As Integer
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If
    If (imRowNo < vbcVehicle.Value) Or (imRowNo >= vbcVehicle.Value + vbcVehicle.LargeChange + 1) Then
        mSetShow ilBoxNo
        Exit Sub
    End If
    lacFrame.Move 0, tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcVehicle.Value) * (fgBoxGridH + 15) - 30
    lacFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcSpec.Top + tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcVehicle.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True


    Select Case ilBoxNo
        Case VEHICLEINDEX   'Not allowed to be changed
        Case DEMOINDEX  'Demo
            lbcDemo.Height = gListBoxHeight(lbcDemo.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 8
            gMoveTableCtrl pbcVehicle, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcVehicle.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If imRowNo - vbcVehicle.Value <= vbcVehicle.LargeChange \ 2 Then
                lbcDemo.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcDemo.Move edcDropDown.Left, edcDropDown.Top - lbcDemo.Height
            End If
            imChgMode = True
            slStr = Trim$(tmBudResch(imRowNo - 1).sDemoName)
            If slStr <> "" Then
                gFindMatch slStr, 1, lbcDemo
                If gLastFound(lbcDemo) > 0 Then
                    lbcDemo.ListIndex = gLastFound(lbcDemo)
                Else
                    lbcDemo.ListIndex = 0
                End If
            Else
                lbcDemo.ListIndex = 0
            End If
            If lbcDemo.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcDemo.List(lbcDemo.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case RATEAUDINDEX  'Avg Rating or Audience
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 7
            gMoveTableCtrl pbcVehicle, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcVehicle.Value) * (fgBoxGridH + 15)
            slStr = gLongToStrDec(tmBudResch(imRowNo - 1).lRateAud, 2)
            ilPos = InStr(slStr, ".00")
            If ilPos > 0 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case CPPCPMINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 6
            gMoveTableCtrl pbcVehicle, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcVehicle.Value) * (fgBoxGridH + 15)
            slStr = gLongToStrDec(tmBudResch(imRowNo - 1).lCPPCPM, 2)
            ilPos = InStr(slStr, ".00")
            If ilPos > 0 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case AVAILSINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 7
            gMoveTableCtrl pbcVehicle, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcVehicle.Value) * (fgBoxGridH + 15)
            slStr = gLongToStrDec(tmBudResch(imRowNo - 1).lAvails, 0)
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case PCTSELLOUTINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 3
            gMoveTableCtrl pbcVehicle, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcVehicle.Value) * (fgBoxGridH + 15)
            slStr = gIntToStrDec(tmBudResch(imRowNo - 1).iPctSellout, 0)
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus

        Case DOLLARINDEX
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
    Dim ilRet As Integer
    imTerminate = False
    imFirstActivate = True
    igBDReturn = 0

    igLBBvfRec = 1  'Match definition in Budget

    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    BudResch.Height = cmcApply.Top + 5 * cmcApply.Height / 3
    gCenterStdAlone BudResch
    mInitBox
    hmVef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", Budget
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    imBypassFocus = False
    imLbcArrowSetting = False
    imSettingValue = False
    imTabDirection = 0
    imRowNo = -1
    imBoxNo = -1
    'BudResch.Show
    Screen.MousePointer = vbHourglass
    mDemoPop
    If imTerminate Then
        Exit Sub
    End If
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    'slStr = Trim$(Str$(lgTotal))
    'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
    'lacTotal.Caption = "(Current Total: " & slStr & ")"
    mComputeTotal
    edcRound.Text = "1"
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    gCenterModalForm BudResch
    'plcScreen.Caption = "Research for " & sgBAName
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
'*             Created:5/17/93       By:D. LeVine      *
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
    flTextHeight = pbcVehicle.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcSpec.Move 135, 255, pbcVehicle.Width + fgPanelAdj + vbcVehicle.Width
    pbcVehicle.Move plcSpec.Left + fgBevelX, plcSpec.Top + fgBevelY
    vbcVehicle.Move pbcVehicle.Left + pbcVehicle.Width - 15, pbcVehicle.Top
    'Vehcile
    gSetCtrl tmCtrls(VEHICLEINDEX), 30, 375, 2265, fgBoxGridH
    'Demo
    gSetCtrl tmCtrls(DEMOINDEX), 2325, tmCtrls(VEHICLEINDEX).fBoxY, 720, fgBoxGridH
    'Rating or Audience
    gSetCtrl tmCtrls(RATEAUDINDEX), 3060, tmCtrls(VEHICLEINDEX).fBoxY, 810, fgBoxGridH
    'CPP or CPM
    gSetCtrl tmCtrls(CPPCPMINDEX), 3885, tmCtrls(VEHICLEINDEX).fBoxY, 810, fgBoxGridH
    'Avails
    gSetCtrl tmCtrls(AVAILSINDEX), 4710, tmCtrls(VEHICLEINDEX).fBoxY, 810, fgBoxGridH
    '% Sellout
    gSetCtrl tmCtrls(PCTSELLOUTINDEX), 5535, tmCtrls(VEHICLEINDEX).fBoxY, 525, fgBoxGridH
    'Dollars
    gSetCtrl tmCtrls(DOLLARINDEX), 6075, tmCtrls(VEHICLEINDEX).fBoxY, 1260, fgBoxGridH
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
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slName As String
    Dim ilBvf As Integer
    Dim slVehicle As String

    ReDim tmBudResch(0 To 0) As BUDRESEARCH
    'For ilLoop = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
    For ilLoop = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
        slName = Trim$(tgBvfRec(ilLoop).sVehicle)
        gFindMatch slName, 0, lbcVehicle
        If gLastFound(lbcVehicle) < 0 Then
            lbcVehicle.AddItem slName
        End If
    Next ilLoop
    ReDim tmBudResch(0 To lbcVehicle.ListCount) As BUDRESEARCH
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        slVehicle = lbcVehicle.List(ilLoop)
        'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                tmBudResch(ilLoop).iVefCode = tgBvfRec(ilBvf).tBvf.iVefCode
                tmVefSrchKey.iCode = tgBvfRec(ilBvf).tBvf.iVefCode
                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                tmBudResch(ilLoop).iMnfDemo = 0
                tmBudResch(ilLoop).sDemoName = ""
                tmBudResch(ilLoop).lRateAud = 0
                tmBudResch(ilLoop).lCPPCPM = 0
                tmBudResch(ilLoop).lAvails = 0
                tmBudResch(ilLoop).iPctSellout = 0
                tmBudResch(ilLoop).lDollars = 0
                If tmVef.iMnfDemo > 0 Then
                    For ilIndex = 0 To UBound(tmDemoCode) - 1 Step 1 'lbcDemoCode.ListCount - 1 Step 1
                        slNameCode = tmDemoCode(ilIndex).sKey   'lbcDemoCode.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = tmVef.iMnfDemo Then
                            ilRet = gParseItem(slNameCode, 1, "\", tmBudResch(ilLoop).sDemoName)
                            tmBudResch(ilLoop).iMnfDemo = tmVef.iMnfDemo
                            tmBudResch(ilLoop).lRateAud = tmVef.lRateAud
                            tmBudResch(ilLoop).lCPPCPM = tmVef.lCPPCPM
                            tmBudResch(ilLoop).lAvails = tmVef.lYearAvails
                            tmBudResch(ilLoop).iPctSellout = tmVef.iPctSellout
                            tmBudResch(ilLoop).lDollars = mComputeDollar(tmVef.lCPPCPM, tmVef.lYearAvails, tmVef.lRateAud, tmVef.iPctSellout)
                            Exit For
                        End If
                    Next ilIndex
                End If
                'Compute Dollars
            End If
        Next ilBvf
    Next ilLoop
    If lbcVehicle.ListCount <= vbcVehicle.LargeChange + 1 Then
        vbcVehicle.Max = vbcVehicle.Min
    Else
        vbcVehicle.Max = lbcVehicle.ListCount - vbcVehicle.LargeChange
    End If
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan     *
'*                                                     *
'*            Comments: Save vehicle values            *
'*                                                     *
'*******************************************************
Private Sub mSaveRec()
    Dim ilRet As Integer
    If imRowNo < 0 Then
        Exit Sub
    End If
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    Do
        tmVefSrchKey.iCode = tmBudResch(imRowNo - 1).iVefCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        tmVef.iMnfDemo = tmBudResch(imRowNo - 1).iMnfDemo
        tmVef.lRateAud = tmBudResch(imRowNo - 1).lRateAud
        tmVef.lCPPCPM = tmBudResch(imRowNo - 1).lCPPCPM
        tmVef.lYearAvails = tmBudResch(imRowNo - 1).lAvails
        tmVef.iPctSellout = tmBudResch(imRowNo - 1).iPctSellout
        ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, "mSaveRec (btrUpdate Or btrGetEqual)", BudResch
    On Error GoTo 0
    ilRet = gBinarySearchVef(tmVef.iCode)
    If ilRet <> -1 Then
        tgMVef(ilRet) = tmVef
    End If
    '11/26/17
    gFileChgdUpdate "vef.btr", False
    
    Exit Sub
mSaveRecErr:
    imTerminate = True
    On Error GoTo 0
    Resume Next
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
    cmcApply.Enabled = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
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
    If (imRowNo < vbcVehicle.Value) Or (imRowNo >= vbcVehicle.Value + vbcVehicle.LargeChange + 1) Then
        Exit Sub
    End If

    Select Case ilBoxNo
        Case VEHICLEINDEX   'Not allowed to be changed
        Case DEMOINDEX  'Demo
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case RATEAUDINDEX  'Avg Rating or Audience
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case CPPCPMINDEX
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case AVAILSINDEX
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PCTSELLOUTINDEX
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case DOLLARINDEX   'Not allowed to be changed
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
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
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    pbcArrow.Visible = False
    lacFrame.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo
        Case VEHICLEINDEX   'Not allowed to be changed
        Case DEMOINDEX  'Demo
            lbcDemo.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcDemo.ListIndex > 0 Then
                slNameCode = tmDemoCode(lbcDemo.ListIndex - 1).sKey 'lbcDemoCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmBudResch(imRowNo - 1).iMnfDemo = Val(slCode)
                ilRet = gParseItem(slNameCode, 1, "\", tmBudResch(imRowNo - 1).sDemoName)
            Else
                tmBudResch(imRowNo - 1).iMnfDemo = 0
                tmBudResch(imRowNo - 1).sDemoName = ""
            End If
        Case RATEAUDINDEX  'Avg Rating or Audience
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            tmBudResch(imRowNo - 1).lRateAud = gStrDecToLong(slStr, 2)
        Case CPPCPMINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            tmBudResch(imRowNo - 1).lCPPCPM = gStrDecToLong(slStr, 2)
        Case AVAILSINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            tmBudResch(imRowNo - 1).lAvails = gStrDecToLong(slStr, 0)
        Case PCTSELLOUTINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            tmBudResch(imRowNo - 1).iPctSellout = gStrDecToLong(slStr, 0)
        Case DOLLARINDEX   'Not allowed to be changed
    End Select
    tmBudResch(imRowNo - 1).lDollars = mComputeDollar(tmBudResch(imRowNo - 1).lCPPCPM, tmBudResch(imRowNo - 1).lAvails, tmBudResch(imRowNo - 1).lRateAud, tmBudResch(imRowNo - 1).iPctSellout)
    mComputeTotal
    mSetCommands
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
    Unload BudResch
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    mSaveRec
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            imSettingValue = True
            vbcVehicle.Value = 1
            imSettingValue = False
            imRowNo = 1
            ilBox = DEMOINDEX
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case DEMOINDEX 'Name (first control within header)
            mSetShow imBoxNo
            mSaveRec
            If imRowNo <= 1 Then
                If cmcApply.Enabled Then
                    imBoxNo = -1
                    imRowNo = -1
                    cmcApply.SetFocus
                    Exit Sub
                End If
                imBoxNo = -1
                imRowNo = -1
                cmcCancel.SetFocus
            Else
                ilBox = PCTSELLOUTINDEX
                imRowNo = imRowNo - 1
                If imRowNo < vbcVehicle.Value Then
                    imSettingValue = True
                    vbcVehicle.Value = vbcVehicle.Value - 1
                    imSettingValue = False
                End If
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            imRowNo = UBound(tmBudResch)
            imSettingValue = True
            If imRowNo <= vbcVehicle.LargeChange + 1 Then
                vbcVehicle.Value = 1
            Else
                vbcVehicle.Value = imRowNo - vbcVehicle.LargeChange - 1
            End If
            imSettingValue = False
            ilBox = PCTSELLOUTINDEX
        Case PCTSELLOUTINDEX  'Last control within header
            mSetShow imBoxNo
            mSaveRec
            If imRowNo >= UBound(tmBudResch) Then     'UBound(tgBvfRec) Then
                If cmcApply.Enabled Then
                    cmcApply.SetFocus
                Else
                    cmcCancel.SetFocus
                End If
                Exit Sub
            End If
            imRowNo = imRowNo + 1
            If imRowNo > vbcVehicle.Value + vbcVehicle.LargeChange - 1 Then
                imSettingValue = True
                vbcVehicle.Value = vbcVehicle.Value + 1
                imSettingValue = False
            End If
            ilBox = DEMOINDEX
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcVehicle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    ilCompRow = vbcVehicle.LargeChange + 1
    If UBound(tmBudResch) >= ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tmBudResch)
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcVehicle.Value - 1
                    If ilRowNo > UBound(tmBudResch) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If (ilBox = VEHICLEINDEX) Or (ilBox = DOLLARINDEX) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    mSetShow imBoxNo
                    If ilRowNo <> imRowNo Then
                        mSaveRec
                    End If
                    imRowNo = ilRow + vbcVehicle.Value - 1
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mSetFocus imBoxNo
End Sub
Private Sub pbcVehicle_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim ilPos As Integer
    ilStartRow = vbcVehicle.Value '+ 1  'Top location
    ilEndRow = vbcVehicle.Value + vbcVehicle.LargeChange ' + 1
    If ilEndRow - 1 >= UBound(tmBudResch) Then
        ilEndRow = UBound(tmBudResch)
    End If
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            'If ilBox <> TOTALINDEX Then
            '    gPaintArea pbcProj, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
            'Else
            '    gPaintArea pbcProj, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
            'End If
            pbcVehicle.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcVehicle.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            Select Case ilBox
                Case VEHICLEINDEX
                    slStr = lbcVehicle.List(ilRow - 1)
                Case DEMOINDEX  'Demo
                    slStr = Trim$(tmBudResch(ilRow - 1).sDemoName)
                Case RATEAUDINDEX  'Avg Rating or Audience
                    slStr = gLongToStrDec(tmBudResch(ilRow - 1).lRateAud, 2)
                    ilPos = InStr(slStr, ".00")
                    If ilPos > 0 Then
                        slStr = Left$(slStr, ilPos - 1)
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    Else
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    End If
                Case CPPCPMINDEX
                    slStr = gLongToStrDec(tmBudResch(ilRow - 1).lCPPCPM, 2)
                    ilPos = InStr(slStr, ".00")
                    If ilPos > 0 Then
                        slStr = Left$(slStr, ilPos - 1)
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    Else
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    End If
                Case AVAILSINDEX
                    slStr = gLongToStrDec(tmBudResch(ilRow - 1).lAvails, 0)
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                Case PCTSELLOUTINDEX
                    slStr = gIntToStrDec(tmBudResch(ilRow - 1).iPctSellout, 0)
                Case DOLLARINDEX
                    slStr = gLongToStrDec(tmBudResch(ilRow - 1).lDollars, 0)
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            End Select
            gSetShow pbcVehicle, slStr, tmCtrls(ilBox)
            pbcVehicle.Print tmCtrls(ilBox).sShow
        Next ilBox
    Next ilRow
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcSpec_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub vbcVehicle_Change()
    If imSettingValue Then
        pbcVehicle.Cls
        pbcVehicle_Paint
        imSettingValue = False
    Else
        mSetShow imBoxNo
        pbcVehicle.Cls
        pbcVehicle_Paint
        mEnableBox imBoxNo
    End If
End Sub
Private Sub vbcVehicle_GotFocus()
    mSetShow imBoxNo
    mSaveRec
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Research for " & sgBAName
End Sub
