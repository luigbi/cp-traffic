VERSION 5.00
Begin VB.Form Bud12Mo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4545
   ClientLeft      =   225
   ClientTop       =   3180
   ClientWidth     =   8880
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
   ScaleHeight     =   4545
   ScaleWidth      =   8880
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
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   60
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5000
      Width           =   60
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
      Left            =   5640
      Picture         =   "Bud12mo.frx":0000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1470
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
      Left            =   4695
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1470
      Visible         =   0   'False
      Width           =   930
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
      Left            =   3510
      ScaleHeight     =   90
      ScaleWidth      =   135
      TabIndex        =   8
      Top             =   2775
      Width           =   135
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
      Left            =   4080
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   3
      Top             =   2715
      Width           =   120
   End
   Begin VB.PictureBox pbcDollars 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   3390
      Picture         =   "Bud12mo.frx":00FA
      ScaleHeight     =   1635
      ScaleWidth      =   5160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   5160
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   4050
   End
   Begin VB.PictureBox plcDollars 
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
      Height          =   1755
      Left            =   3345
      ScaleHeight     =   1695
      ScaleWidth      =   5220
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   915
      Width           =   5280
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   3870
      TabIndex        =   9
      Top             =   4080
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
      Left            =   150
      ScaleHeight     =   3450
      ScaleWidth      =   2880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   315
      Width           =   2940
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   3390
         Left            =   30
         TabIndex        =   2
         Top             =   45
         Width           =   2820
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   255
      Top             =   4005
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "Bud12Mo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Bud12mo.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Bud12Mo.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Trend input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imYear As Integer
Dim imStartIndex As Integer
Dim imBoxNo As Integer   'Current event name Box

'Dim tmCtrls(1 To 12)  As FIELDAREA
Dim tmCtrls(0 To 12)  As FIELDAREA
Dim tmLBCtrls As Integer

'Dim tmQCtrls(1 To 5)  As FIELDAREA
Dim tmQCtrls(0 To 5)  As FIELDAREA
Dim imLBQCtrls As Integer

'Dim tmMnthInfo(1 To 12) As MNTHINFO
Dim tmMnthInfo(0 To 12) As MNTHINFO
Dim imLBMnthInfo As Integer

'Dim lmSave(1 To 12) As Long
Dim lmSave(0 To 12) As Long
Dim imLBSave As Integer

'Dim smShow(1 To 12) As String
Dim smShow(0 To 12) As String
Dim imLBShow As Integer


Dim imBvfIndex() As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)

Dim imUpdateAllowed As Integer

Const DOLLAR1INDEX = 1
Const DOLLAR2INDEX = 2
Const DOLLAR3INDEX = 3
Const DOLLAR4INDEX = 4
Const DOLLAR5INDEX = 5
Const DOLLAR6INDEX = 6
Const DOLLAR7INDEX = 7
Const DOLLAR8INDEX = 8
Const DOLLAR9INDEX = 9
Const DOLLAR10INDEX = 10
Const DOLLAR11INDEX = 11
Const DOLLAR12INDEX = 12
Const Q1INDEX = 1
Const Q2INDEX = 2
Const Q3INDEX = 3
Const Q4INDEX = 4
Const TOTALINDEX = 5
Private Sub cmcDone_Click()
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub edcDropDown_GotFocus()
    gCtrlGotFocus edcDropDown
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim slComp As String

    Select Case imBoxNo
        Case DOLLAR1INDEX To DOLLAR12INDEX
            'If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            '    Beep
            '    KeyAscii = 0
            '    Exit Sub
            'End If
            If (KeyAscii = KEYNEG) And ((Len(edcDropDown.Text) = 0) Or (Len(edcDropDown.Text) = edcDropDown.SelLength)) Then
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYNEG) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            Else
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            slComp = "99999999"
            'If gCompNumberStr(slStr, slComp) > 0 Then
            If gCompAbsNumberStr(slStr, slComp) > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
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
    Bud12Mo.Refresh
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
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imBvfIndex
    Set Bud12Mo = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcVehicle_Click()
    Dim slStr As String
    'Determine Row and obtain value
    If Budget!rbcSort(0).Value Then 'Vehicle Total
        mComputeTotals
    ElseIf Budget!rbcSort(1).Value Then 'Office Total
        mComputeTotals
    ElseIf Budget!rbcSort(2).Value Then 'Vehicle within Office
        slStr = lbcVehicle.List(lbcVehicle.ListIndex)
        If Left$(slStr, 2) <> "  " Then
            tmcDelay.Enabled = True
            Exit Sub
        End If
        mComputeTotals
    ElseIf Budget!rbcSort(3).Value Then 'Office within vehicle
        slStr = lbcVehicle.List(lbcVehicle.ListIndex)
        If Left$(slStr, 2) <> "  " Then
            tmcDelay.Enabled = True
            Exit Sub
        End If
        mComputeTotals
    End If
End Sub
Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mComputeQTotal                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute Quarter budget totals  *
'*                                                     *
'*******************************************************
Private Sub mComputeQTotals()
    Dim ilMnth As Integer
    Dim ilWk As Integer
    Dim ilBvf As Integer
    Dim ilQBox As Integer
    Dim llQTotal As Long
    Dim llGTotal As Long
    Dim slStr As String
    llGTotal = 0
    For ilQBox = 1 To 4 Step 1
        llQTotal = 0
        For ilMnth = 3 * (ilQBox - 1) + 1 To 3 * (ilQBox - 1) + 3 Step 1
            For ilWk = tmMnthInfo(ilMnth).iStartWkNo To tmMnthInfo(ilMnth).iEndWkNo Step 1
                For ilBvf = LBound(imBvfIndex) To UBound(imBvfIndex) - 1 Step 1
                    llQTotal = llQTotal + tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk)
                Next ilBvf
            Next ilWk
        Next ilMnth
        llGTotal = llGTotal + llQTotal
        slStr = gLongToStrDec(llQTotal, 0)
        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
        gSetShow pbcDollars, slStr, tmQCtrls(ilQBox)
    Next ilQBox
    slStr = gLongToStrDec(llGTotal, 0)
    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
    gSetShow pbcDollars, slStr, tmQCtrls(5)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mComputeTotal                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute budget totals          *
'*                                                     *
'*******************************************************
Private Sub mComputeTotals()
    Dim slVehicle As String
    Dim slOffice As String
    Dim ilLoop As Integer
    Dim ilBvf As Integer
    Dim ilMnth As Integer
    Dim ilWk As Integer
    Dim slStr As String
    If lbcVehicle.ListIndex < 0 Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    pbcDollars.Cls
    If Budget!rbcSort(0).Value Then 'Vehicle
        slVehicle = lbcVehicle.List(lbcVehicle.ListIndex)
        For ilMnth = 1 To 12 Step 1
            lmSave(ilMnth) = 0
        Next ilMnth
        ReDim imBvfIndex(0 To 0) As Integer
        'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                imBvfIndex(UBound(imBvfIndex)) = ilBvf
                ReDim Preserve imBvfIndex(0 To UBound(imBvfIndex) + 1) As Integer
                For ilMnth = 1 To 12 Step 1
                    For ilWk = tmMnthInfo(ilMnth).iStartWkNo To tmMnthInfo(ilMnth).iEndWkNo Step 1
                        If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                            lmSave(ilMnth) = lmSave(ilMnth) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                        End If
                    Next ilWk
                Next ilMnth
            End If
        Next ilBvf
        For ilMnth = 1 To 12 Step 1
            slStr = Trim$(Str$(lmSave(ilMnth)))
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            gSetShow pbcDollars, slStr, tmCtrls(ilMnth)
            smShow(ilMnth) = tmCtrls(ilMnth).sShow
        Next ilMnth
    ElseIf Budget!rbcSort(1).Value Then    'office
        slOffice = lbcVehicle.List(lbcVehicle.ListIndex)
        For ilMnth = 1 To 12 Step 1
            lmSave(ilMnth) = 0
        Next ilMnth
        ReDim imBvfIndex(0 To 0) As Integer
        'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            If StrComp(slOffice, Trim$(tgBvfRec(ilBvf).SOffice), 1) = 0 Then
                imBvfIndex(UBound(imBvfIndex)) = ilBvf
                ReDim Preserve imBvfIndex(0 To UBound(imBvfIndex) + 1) As Integer
                For ilMnth = 1 To 12 Step 1
                    For ilWk = tmMnthInfo(ilMnth).iStartWkNo To tmMnthInfo(ilMnth).iEndWkNo Step 1
                        If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                            lmSave(ilMnth) = lmSave(ilMnth) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                        End If
                    Next ilWk
                Next ilMnth
            End If
        Next ilBvf
        For ilMnth = 1 To 12 Step 1
            slStr = Trim$(Str$(lmSave(ilMnth)))
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            gSetShow pbcDollars, slStr, tmCtrls(ilMnth)
            smShow(ilMnth) = tmCtrls(ilMnth).sShow
        Next ilMnth
    ElseIf Budget!rbcSort(2).Value Then    'Vehicle within office
        slVehicle = Trim$(lbcVehicle.List(lbcVehicle.ListIndex))
        For ilLoop = lbcVehicle.ListIndex - 1 To 0 Step -1
            slStr = lbcVehicle.List(ilLoop)
            If Left$(slStr, 2) <> "  " Then
                slOffice = slStr
                Exit For
            End If
        Next ilLoop
        For ilMnth = 1 To 12 Step 1
            lmSave(ilMnth) = 0
        Next ilMnth
        'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            If (StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0) And (StrComp(slOffice, Trim$(tgBvfRec(ilBvf).SOffice), 1) = 0) Then
                ReDim imBvfIndex(0 To 1) As Integer
                imBvfIndex(0) = ilBvf
                For ilMnth = 1 To 12 Step 1
                    For ilWk = tmMnthInfo(ilMnth).iStartWkNo To tmMnthInfo(ilMnth).iEndWkNo Step 1
                        If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                            lmSave(ilMnth) = lmSave(ilMnth) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                        End If
                    Next ilWk
                Next ilMnth
                Exit For
            End If
        Next ilBvf
        For ilMnth = 1 To 12 Step 1
            slStr = Trim$(Str$(lmSave(ilMnth)))
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            gSetShow pbcDollars, slStr, tmCtrls(ilMnth)
            smShow(ilMnth) = tmCtrls(ilMnth).sShow
        Next ilMnth
    ElseIf Budget!rbcSort(3).Value Then    'office within vehicle
        slOffice = Trim$(lbcVehicle.List(lbcVehicle.ListIndex))
        For ilLoop = lbcVehicle.ListIndex - 1 To 0 Step -1
            slStr = lbcVehicle.List(ilLoop)
            If Left$(slStr, 2) <> "  " Then
                slVehicle = slStr
                Exit For
            End If
        Next ilLoop
        For ilMnth = 1 To 12 Step 1
            lmSave(ilMnth) = 0
        Next ilMnth
        'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            If (StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0) And (StrComp(slOffice, Trim$(tgBvfRec(ilBvf).SOffice), 1) = 0) Then
                ReDim imBvfIndex(0 To 1) As Integer
                imBvfIndex(0) = ilBvf
                For ilMnth = 1 To 12 Step 1
                    For ilWk = tmMnthInfo(ilMnth).iStartWkNo To tmMnthInfo(ilMnth).iEndWkNo Step 1
                        If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                            lmSave(ilMnth) = lmSave(ilMnth) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                        End If
                    Next ilWk
                Next ilMnth
                Exit For
            End If
        Next ilBvf
        For ilMnth = 1 To 12 Step 1
            slStr = Trim$(Str$(lmSave(ilMnth)))
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            gSetShow pbcDollars, slStr, tmCtrls(ilMnth)
            smShow(ilMnth) = tmCtrls(ilMnth).sShow
        Next ilMnth
    End If
    mComputeQTotals
    pbcDollars_Paint
    Screen.MousePointer = vbDefault
End Sub
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
    If ilBoxNo < tmLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If


    Select Case ilBoxNo
        Case DOLLAR1INDEX To DOLLAR12INDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcDollars, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcDropDown.Text = Trim$(Str$(lmSave(ilBoxNo)))
            edcDropDown.Visible = True  'Set visibility
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
    Dim slDate As String
    Dim ilLoop As Integer
    Dim ilMnth As Integer
    Dim ilFirstFlag  As Integer
    imTerminate = False
    imFirstActivate = True
    igBDReturn = 0
    Screen.MousePointer = vbHourglass
    
    tmLBCtrls = 1
    imLBMnthInfo = 1
    imLBSave = 1
    imLBQCtrls = 1
    
    igLBBvfRec = 1  'Match definition in Budget
    
    Bud12Mo.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    mInitBox
    gCenterStdAlone Bud12Mo
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    imBoxNo = -1
    imTabDirection = 0  'Left to right movement
    'Bud12Mo.Show
    Screen.MousePointer = vbHourglass
    imStartIndex = -1
    'imYear = tgBvfRec(LBound(tgBvfRec)).tBvf.iYear
    imYear = tgBvfRec(igLBBvfRec).tBvf.iYear
    If tgSpf.sRUseCorpCal <> "Y" Then
        'Standard Months
        For ilMnth = 1 To 12 Step 1
            slDate = Trim$(Str$(ilMnth)) & "/15/" & Trim$(Str$(imYear))
            tmMnthInfo(ilMnth).sName = gMonthName(slDate)
            slDate = gObtainStartStd(slDate)
            gObtainWkNo 0, slDate, tmMnthInfo(ilMnth).iStartWkNo, ilFirstFlag
            slDate = gObtainEndStd(slDate)
            gObtainWkNo 0, slDate, tmMnthInfo(ilMnth).iEndWkNo, ilFirstFlag
        Next ilMnth
    Else
        'Corporate Months
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            If tgMCof(ilLoop).iYear = imYear Then
                For ilMnth = tgMCof(ilLoop).iStartMnthNo To 12 Step 1
                    slDate = Trim$(Str$(ilMnth)) & "/15/" & Trim$(Str$(imYear - 1))
                    tmMnthInfo(ilMnth - tgMCof(ilLoop).iStartMnthNo + 1).sName = gMonthName(slDate)
                    slDate = gObtainStartCorp(slDate, False)
                    gObtainWkNo 5, slDate, tmMnthInfo(ilMnth - tgMCof(ilLoop).iStartMnthNo + 1).iStartWkNo, ilFirstFlag
                    slDate = gObtainEndCorp(slDate, False)
                    gObtainWkNo 5, slDate, tmMnthInfo(ilMnth - tgMCof(ilLoop).iStartMnthNo + 1).iEndWkNo, ilFirstFlag
                Next ilMnth
                For ilMnth = 1 To tgMCof(ilLoop).iStartMnthNo - 1 Step 1
                    slDate = Trim$(Str$(ilMnth)) & "/15/" & Trim$(Str$(imYear))
                    tmMnthInfo((13 - tgMCof(ilLoop).iStartMnthNo) + ilMnth).sName = gMonthName(slDate)
                    slDate = gObtainStartCorp(slDate, False)
                    gObtainWkNo 5, slDate, tmMnthInfo((13 - tgMCof(ilLoop).iStartMnthNo) + ilMnth).iStartWkNo, ilFirstFlag
                    slDate = gObtainEndCorp(slDate, False)
                    gObtainWkNo 5, slDate, tmMnthInfo((13 - tgMCof(ilLoop).iStartMnthNo) + ilMnth).iEndWkNo, ilFirstFlag
                Next ilMnth
                Exit For
            End If
        Next ilLoop
    End If
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    'plcScreen.Caption = "12 Months Budget for " & sgBAName
'    gCenterModalForm Bud12Mo
    Screen.MousePointer = vbDefault
    Exit Sub

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
    flTextHeight = pbcDollars.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcDollars.Move 3345, 915, pbcDollars.Width + fgPanelAdj
    pbcDollars.Move plcDollars.Left + fgBevelX, plcDollars.Top + fgBevelY
    'Dollars
    gSetCtrl tmCtrls(DOLLAR1INDEX), 30, 30, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmCtrls(DOLLAR2INDEX), 1305, 30, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmCtrls(DOLLAR3INDEX), 2580, 30, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmCtrls(DOLLAR4INDEX), 30, 375, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmCtrls(DOLLAR5INDEX), 1305, 375, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmCtrls(DOLLAR6INDEX), 2580, 375, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmCtrls(DOLLAR7INDEX), 30, 720, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmCtrls(DOLLAR8INDEX), 1305, 720, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmCtrls(DOLLAR9INDEX), 2580, 720, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmCtrls(DOLLAR10INDEX), 30, 1065, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmCtrls(DOLLAR11INDEX), 1305, 1065, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmCtrls(DOLLAR12INDEX), 2580, 1065, 1260, fgBoxStH

    'Dollars
    gSetCtrl tmQCtrls(Q1INDEX), 3885, 30, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmQCtrls(Q2INDEX), 3885, 375, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmQCtrls(Q3INDEX), 3885, 720, 1260, fgBoxStH
    'Dollars
    gSetCtrl tmQCtrls(Q4INDEX), 3885, 1065, 1260, fgBoxStH
    'Total
    gSetCtrl tmQCtrls(TOTALINDEX), 3885, 1440, 1260, fgBoxGridH
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
    Dim slPOffice As String
    Dim slPVehicle As String
    Dim ilRowNo As Integer

    slPOffice = ""
    slPVehicle = ""
    'For ilRowNo = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
    For ilRowNo = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
        If slPOffice = "" Then
            'Add row to smSave and smShow
            If Budget!rbcSort(1).Value Or Budget!rbcSort(2).Value Then    'Vehicle within office
                lbcVehicle.AddItem Trim$(tgBvfRec(ilRowNo).SOffice)
            ElseIf Budget!rbcSort(0).Value Or Budget!rbcSort(3).Value Then
                lbcVehicle.AddItem Trim$(tgBvfRec(ilRowNo).sVehicle)
            End If
            If Budget!rbcSort(2).Value Or Budget!rbcSort(3).Value Then
                If Budget!rbcSort(2).Value Then    'Vehicle within office
                    lbcVehicle.AddItem "  " & Trim$(tgBvfRec(ilRowNo).sVehicle)
                ElseIf Budget!rbcSort(3).Value Then
                    lbcVehicle.AddItem "  " & Trim$(tgBvfRec(ilRowNo).SOffice)
                End If
            End If
            slPOffice = Trim$(tgBvfRec(ilRowNo).SOffice)
            slPVehicle = Trim$(tgBvfRec(ilRowNo).sVehicle)
        Else
            If Budget!rbcSort(1).Value Then
                If StrComp(slPOffice, Trim$(tgBvfRec(ilRowNo).SOffice), 1) <> 0 Then
                    lbcVehicle.AddItem Trim$(tgBvfRec(ilRowNo).SOffice)
                End If
                slPOffice = Trim$(tgBvfRec(ilRowNo).SOffice)
            ElseIf Budget!rbcSort(2).Value Then    'Vehicle within office
                If StrComp(slPOffice, Trim$(tgBvfRec(ilRowNo).SOffice), 1) <> 0 Then
                    lbcVehicle.AddItem Trim$(tgBvfRec(ilRowNo).SOffice)
                End If
                lbcVehicle.AddItem "  " & Trim$(tgBvfRec(ilRowNo).sVehicle)
                slPOffice = Trim$(tgBvfRec(ilRowNo).SOffice)
            ElseIf Budget!rbcSort(0).Value Then
                If StrComp(slPVehicle, Trim$(tgBvfRec(ilRowNo).sVehicle), 1) <> 0 Then
                    lbcVehicle.AddItem Trim$(tgBvfRec(ilRowNo).sVehicle)
                End If
                slPVehicle = Trim$(tgBvfRec(ilRowNo).sVehicle)
            ElseIf Budget!rbcSort(3).Value Then
                If StrComp(slPVehicle, Trim$(tgBvfRec(ilRowNo).sVehicle), 1) <> 0 Then
                    lbcVehicle.AddItem Trim$(tgBvfRec(ilRowNo).sVehicle)
                End If
                lbcVehicle.AddItem "  " & Trim$(tgBvfRec(ilRowNo).SOffice)
                slPVehicle = Trim$(tgBvfRec(ilRowNo).sVehicle)
            End If
        End If
        If imStartIndex = -1 Then
            If Budget!rbcSort(0).Value Then
                If sgBvfVehName = Trim$(tgBvfRec(ilRowNo).sVehicle) Then
                    imStartIndex = lbcVehicle.ListCount - 1
                End If
            ElseIf Budget!rbcSort(1).Value Then
                If sgBvfOffName = Trim$(tgBvfRec(ilRowNo).SOffice) Then
                    imStartIndex = lbcVehicle.ListCount - 1
                End If
            ElseIf Budget!rbcSort(2).Value Then    'Vehicle within office
                If sgBvfOffName = Trim$(tgBvfRec(ilRowNo).SOffice) Then
                    If sgBvfVehName = Trim$(tgBvfRec(ilRowNo).sVehicle) Then
                        imStartIndex = lbcVehicle.ListCount - 1
                    End If
                End If
            ElseIf Budget!rbcSort(3).Value Then
                If sgBvfVehName = Trim$(tgBvfRec(ilRowNo).sVehicle) Then
                    If sgBvfOffName = Trim$(tgBvfRec(ilRowNo).SOffice) Then
                        imStartIndex = lbcVehicle.ListCount - 1
                    End If
                End If
            End If
        End If
    Next ilRowNo
    If imStartIndex <> -1 Then
        lbcVehicle.ListIndex = imStartIndex
    End If
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
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < tmLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case DOLLAR1INDEX To DOLLAR12INDEX 'Vehicle
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
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim llDollar As Long
    Dim llTDollar As Long
    Dim ilBvf As Integer
    Dim ilWk As Integer
    Dim llAvgDollar As Long
    Dim llTAvgDollar As Long
    Dim slDollar As String
    Dim ilRet As Integer
    Dim llUpper As Long

    On Error GoTo mSetShowErr
    ilRet = 0
    llUpper = LBound(imBvfIndex)
    If ilRet <> 0 Then
        Exit Sub
    End If

    If (ilBoxNo < tmLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    If ilRet <> 0 Then
        Exit Sub
    End If
    On Error GoTo 0

    Select Case ilBoxNo 'Branch on box type (control)
        Case DOLLAR1INDEX To DOLLAR12INDEX
            edcDropDown.Visible = False
            slDollar = edcDropDown.Text
            llDollar = Val(slDollar)
            If lmSave(ilBoxNo) <> llDollar Then
                'Set new values into fields
                lmSave(ilBoxNo) = llDollar
                slStr = Trim$(Str$(lmSave(ilBoxNo)))
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                gSetShow pbcDollars, slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo) = tmCtrls(ilBoxNo).sShow

                llTAvgDollar = 0
                For ilWk = tmMnthInfo(ilBoxNo).iStartWkNo To tmMnthInfo(ilBoxNo).iEndWkNo Step 1
                    llTDollar = 0
                    For ilBvf = LBound(imBvfIndex) To UBound(imBvfIndex) - 1 Step 1
                        llTDollar = llTDollar + tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk)
                    Next ilBvf
                    For ilBvf = LBound(imBvfIndex) To UBound(imBvfIndex) - 1 Step 1
                        llAvgDollar = llDollar / (tmMnthInfo(ilBoxNo).iEndWkNo - tmMnthInfo(ilBoxNo).iStartWkNo + 1)
                        If llTDollar = 0 Then
                            tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) = llAvgDollar / (UBound(imBvfIndex))
                            If (tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) <= 0) And (llAvgDollar > 0) Then
                                tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) = 1
                            ElseIf (tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) >= 0) And (llAvgDollar < 0) Then
                                tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) = -1
                            End If
                        Else
                            If tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) <> 0 Then
                                tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) = CLng((CDbl(llAvgDollar) * CDbl(tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk))) / CDbl(llTDollar))
                                If (tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) <= 0) And (llAvgDollar > 0) Then
                                    tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) = 1
                                ElseIf (tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) >= 0) And (llAvgDollar < 0) Then
                                    tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) = -1
                                End If
                            End If
                        End If
                        llTAvgDollar = llTAvgDollar + tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk)
                    Next ilBvf
                Next ilWk

                If llTAvgDollar <> llDollar Then
                    For ilWk = tmMnthInfo(ilBoxNo).iStartWkNo To tmMnthInfo(ilBoxNo).iEndWkNo Step 1
                        For ilBvf = LBound(imBvfIndex) To UBound(imBvfIndex) - 1 Step 1
                            If tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) <> 0 Then
                                If llTAvgDollar > llDollar Then
                                    tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) = tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) - 1
                                    llTAvgDollar = llTAvgDollar - 1
                                ElseIf llTAvgDollar < llDollar Then
                                    tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) = tgBvfRec(imBvfIndex(ilBvf)).tBvf.lGross(ilWk) + 1
                                    llTAvgDollar = llTAvgDollar + 1
                                End If
                                If llTAvgDollar = llDollar Then
                                    Exit For
                                End If
                            End If
                        Next ilBvf
                        If llTAvgDollar = llDollar Then
                            Exit For
                        End If
                    Next ilWk
                End If
                pbcDollars.Cls
                mComputeQTotals
                pbcDollars_Paint
            End If
    End Select
    mSetCommands
    Exit Sub
mSetShowErr:
    ilRet = 1
    Resume Next
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
    Unload Bud12Mo
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub pbcDollars_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = tmLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub
Private Sub pbcDollars_Paint()
    Dim ilBox As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    llColor = pbcDollars.ForeColor
    slFontName = pbcDollars.FontName
    flFontSize = pbcDollars.FontSize
    pbcDollars.ForeColor = BLUE
    pbcDollars.FontBold = False
    pbcDollars.FontSize = 7
    pbcDollars.FontName = "Arial"
    pbcDollars.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    For ilBox = 1 To 12 Step 1
        pbcDollars.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX - 15
        pbcDollars.CurrentY = tmCtrls(ilBox).fBoxY - 15 '+ fgBoxInsetY
        pbcDollars.Print tmMnthInfo(ilBox).sName
    Next ilBox
    pbcDollars.FontSize = flFontSize
    pbcDollars.FontName = slFontName
    pbcDollars.FontSize = flFontSize
    pbcDollars.ForeColor = llColor
    pbcDollars.FontBold = True
    For ilBox = 1 To 12 Step 1
        pbcDollars.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcDollars.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcDollars.Print smShow(ilBox)
    Next ilBox
    For ilBox = 1 To 4 Step 1
        pbcDollars.CurrentX = tmQCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcDollars.CurrentY = tmQCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcDollars.Print tmQCtrls(ilBox).sShow
    Next ilBox
    pbcDollars.CurrentX = tmQCtrls(5).fBoxX + fgBoxInsetX
    pbcDollars.CurrentY = tmQCtrls(5).fBoxY - 30
    'pbcDollars.Print tmQCtrls(5).sShow
    pbcDollars.Print tmQCtrls(4).sShow
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If lbcVehicle.ListIndex < 0 Then
        cmcDone.SetFocus
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            ilBox = DOLLAR1INDEX
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case DOLLAR1INDEX 'Name (first control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If lbcVehicle.ListIndex <= 0 Then
                cmcDone.SetFocus
                Exit Sub
            End If
            lbcVehicle.ListIndex = lbcVehicle.ListIndex - 1
            ilBox = DOLLAR1INDEX
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
    If lbcVehicle.ListIndex < 0 Then
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            ilBox = DOLLAR12INDEX
        Case DOLLAR12INDEX 'Last control within header
            mSetShow imBoxNo
            imBoxNo = -1
            If lbcVehicle.ListIndex >= lbcVehicle.ListCount - 1 Then
                cmcDone.SetFocus
                Exit Sub
            End If
            lbcVehicle.ListIndex = lbcVehicle.ListIndex + 1
            ilBox = DOLLAR1INDEX
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcDelay_Timer()
    tmcDelay.Enabled = False
    lbcVehicle.ListIndex = lbcVehicle.ListIndex + 1
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "12 Months Budget for " & sgBAName
End Sub
