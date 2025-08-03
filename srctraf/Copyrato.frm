VERSION 5.00
Begin VB.Form CopyRato 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5385
   ClientLeft      =   1110
   ClientTop       =   1500
   ClientWidth     =   7620
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5385
   ScaleWidth      =   7620
   Begin VB.VScrollBar vbcInstructions 
      Height          =   3525
      LargeChange     =   15
      Left            =   7170
      TabIndex        =   6
      Top             =   615
      Width           =   240
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3825
      MaxLength       =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox pbcInstructions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   3510
      Left            =   270
      Picture         =   "Copyrato.frx":0000
      ScaleHeight     =   3510
      ScaleWidth      =   6900
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   615
      Width           =   6900
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   30
      ScaleHeight     =   165
      ScaleWidth      =   135
      TabIndex        =   5
      Top             =   4215
      Width           =   135
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   270
      Width           =   105
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   30
      ScaleHeight     =   165
      ScaleWidth      =   105
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5070
      Width           =   105
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
      Left            =   60
      ScaleHeight     =   240
      ScaleWidth      =   2865
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   2865
   End
   Begin VB.CommandButton cmcGenerate 
      Appearance      =   0  'Flat
      Caption         =   "&Generate"
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
      Left            =   2730
      TabIndex        =   7
      Top             =   5010
      Width           =   945
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
      Left            =   3975
      TabIndex        =   8
      Top             =   5010
      Width           =   945
   End
   Begin VB.PictureBox plcInstructions 
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
      Height          =   4620
      Left            =   225
      ScaleHeight     =   4560
      ScaleWidth      =   7185
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   225
      Width           =   7245
      Begin VB.PictureBox plcSort 
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
         Height          =   255
         Left            =   105
         ScaleHeight     =   255
         ScaleWidth      =   3780
         TabIndex        =   11
         Top             =   90
         Width           =   3780
         Begin VB.OptionButton rbcSort 
            Caption         =   "Top"
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
            Height          =   195
            Index           =   0
            Left            =   1635
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   15
            Width           =   615
         End
         Begin VB.OptionButton rbcSort 
            Caption         =   "Bottom"
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
            Height          =   195
            Index           =   1
            Left            =   2250
            TabIndex        =   13
            Top             =   15
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.PictureBox plcSpots 
         BackColor       =   &H00BFFFFF&
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   3720
         ScaleHeight     =   180
         ScaleWidth      =   735
         TabIndex        =   20
         Top             =   4275
         Width           =   795
      End
      Begin VB.PictureBox plcSpots 
         BackColor       =   &H00BFFFFF&
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   6045
         ScaleHeight     =   180
         ScaleWidth      =   735
         TabIndex        =   9
         Top             =   3975
         Width           =   795
      End
      Begin VB.PictureBox plcSpots 
         BackColor       =   &H00BFFFFF&
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   4905
         ScaleHeight     =   180
         ScaleWidth      =   735
         TabIndex        =   18
         Top             =   3975
         Width           =   795
      End
      Begin VB.PictureBox plcSpots 
         BackColor       =   &H00BFFFFF&
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   3720
         ScaleHeight     =   180
         ScaleWidth      =   735
         TabIndex        =   17
         Top             =   3975
         Width           =   795
      End
      Begin VB.PictureBox plcUse 
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
         Height          =   255
         Left            =   4260
         ScaleHeight     =   255
         ScaleWidth      =   2730
         TabIndex        =   14
         Top             =   105
         Width           =   2730
         Begin VB.OptionButton rbcUse 
            Caption         =   "Altered"
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
            Height          =   195
            Index           =   1
            Left            =   1695
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   15
            Width           =   915
         End
         Begin VB.OptionButton rbcUse 
            Caption         =   "Reduced"
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
            Height          =   195
            Index           =   0
            Left            =   510
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   15
            Width           =   1110
         End
      End
      Begin VB.Label lbc4Wk 
         Appearance      =   0  'Flat
         Caption         =   "#  Spots in Rotation (max 4 weeks)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   960
         TabIndex        =   19
         Top             =   4290
         Width           =   2640
      End
      Begin VB.Label lacSpotsReq 
         Appearance      =   0  'Flat
         Caption         =   "Spots required for Full Rotation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1245
         TabIndex        =   10
         Top             =   3990
         Width           =   2265
      End
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6300
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4830
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6675
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4980
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6525
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4935
      Visible         =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "CopyRato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Copyrato.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CopyRato.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Copy Ratio input screen code
Option Explicit
Option Compare Text
Dim smMatchCode As String
Dim smInstruction() As String
Dim smCode() As String
Dim smInputInstCode As String   '# in 1st 4weeks|Count|Instruction\Code|Instuction\Code|...|Instruction\Code|
Dim smGenerateCode() As String  'Codes to be generated
Dim smWorkArea() As String
Dim imFirstActivate As Integer
'Program library dates Field Areas
Dim tmCtrls(0 To 7)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current Media Box
Dim imRowNo As Integer      'Current row number in Program area (start at 0)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imBypassFocus As Integer
Dim smShow() As String  'Values shown in date/time area
Dim smSave() As String  'Values saved (1=Start time; 2=Start date; 3=End date)
Dim imUpdateAllowed As Integer
Dim smSpotsValue(0 To 3) As String

Const LBONE = 1

Const INSTRUCTIONINDEX = 1    'Instruction control/field
Const INPUTRATIOINDEX = 2          'Instruction Input Ratio control/field
Const INPUTPERCENTINDEX = 3    'Instruction Input % control/field
Const REDUCEDRATIOINDEX = 4          'Instruction Reduced Ratio control/field
Const REDUCEDPERCENTINDEX = 5    'Instruction Reduced Ratio control/field
Const ALTEREDRATIOINDEX = 6          'Instruction Altered Ratio control/field
Const ALTEREDPERCENTINDEX = 7    'Instruction Altered Ratio control/field
Private Sub cmcCancel_Click()
    igCopyInvCallSource = CALLCANCELLED
    smMatchCode = "0"
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcGenerate_Click()
    Dim ilLoop As Integer
    'Create new instruction pattern and store into slMatchCode
    If Not mGeneratePattern() Then
        Exit Sub
    End If
    igCopyInvCallSource = CALLDONE
    smMatchCode = Trim$(str$(UBound(smGenerateCode)))   'Count\Inst Code\Inst Code\...\Inst Code\
    If rbcSort(0).Value Then
        For ilLoop = LBONE To UBound(smGenerateCode) Step 1
            smMatchCode = smMatchCode & "\" & smGenerateCode(ilLoop)
        Next ilLoop
    Else
        'Reverse order
        For ilLoop = UBound(smGenerateCode) To LBONE Step -1
            smMatchCode = smMatchCode & "\" & smGenerateCode(ilLoop)
        Next ilLoop
    End If
    mTerminate
End Sub
Private Sub cmcGenerate_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub edcDropDown_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(COPYJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    CopyRato.Refresh
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
'*******************************************************
'*                                                     *
'*      Procedure Name:mComputeValue                   *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute rdeuced and Altered    *
'*                      Values                         *
'*                                                     *
'*******************************************************
Private Sub mComputeValues()
    Dim ilAllDefined As Integer
    Dim ilLoop As Integer
    Dim ilField As Integer
    Dim slStr As String
    Dim llSmallest As Long
    Dim ilDivisible As Integer
    Dim llLCD As Long
    Dim llTotal As Long
    Screen.MousePointer = vbHourglass
    pbcInstructions.Cls
    For ilField = 3 To 7 Step 1
        For ilLoop = LBONE To UBound(smSave, 2) Step 1
            smSave(ilField, ilLoop) = ""
            slStr = smSave(ilField, ilLoop)
            gSetShow pbcInstructions, slStr, tmCtrls(ilField)
            smShow(ilField, ilLoop) = tmCtrls(ilField).sShow
        Next ilLoop
    Next ilField
    ilAllDefined = True
    llSmallest = -1
    For ilLoop = LBONE To UBound(smSave, 2) Step 1
        If (Len(Trim$(smSave(INPUTRATIOINDEX, ilLoop))) = 0) Or (Val(smSave(INPUTRATIOINDEX, ilLoop)) = 0) Then
            ilAllDefined = False
        Else
            If llSmallest = -1 Then
                llSmallest = Val(smSave(INPUTRATIOINDEX, ilLoop))
            Else
                If Val(smSave(INPUTRATIOINDEX, ilLoop)) < llSmallest Then
                    llSmallest = Val(smSave(INPUTRATIOINDEX, ilLoop))
                End If
            End If
        End If
    Next ilLoop
    If ilAllDefined Then
        llTotal = 0
        For ilLoop = LBONE To UBound(smSave, 2) Step 1
            llTotal = llTotal + Val(smSave(INPUTRATIOINDEX, ilLoop))
        Next ilLoop
        For ilLoop = LBONE To UBound(smSave, 2) Step 1
            smSave(INPUTPERCENTINDEX, ilLoop) = gDivStr(gMulStr(smSave(INPUTRATIOINDEX, ilLoop), "100.0"), str$(llTotal))
        Next ilLoop
        'plcSpots(0).Caption = Trim$(Str$(llTotal))
        smSpotsValue(0) = Trim$(str$(llTotal))
        plcSpots(0).Cls
        plcSpots(0).CurrentX = 0
        plcSpots(0).CurrentY = 0
        plcSpots(0).Print Trim$(str$(llTotal))
        'Compute reduced- Look for largest common denominator
        llLCD = llSmallest
        Do
            ilDivisible = True
            For ilLoop = LBONE To UBound(smSave, 2) Step 1
                If (Val(smSave(INPUTRATIOINDEX, ilLoop)) Mod llLCD) <> 0 Then
                    ilDivisible = False
                    llLCD = llLCD - 1
                    Exit For
                End If
            Next ilLoop
        Loop While (Not ilDivisible) And (llLCD > 1)
        If ilDivisible Then
            For ilLoop = LBONE To UBound(smSave, 2) Step 1
                smSave(REDUCEDRATIOINDEX, ilLoop) = gDivStr(smSave(INPUTRATIOINDEX, ilLoop), str$(llLCD))
            Next ilLoop
            rbcUse(0).Value = True
        Else
            For ilLoop = LBONE To UBound(smSave, 2) Step 1
                smSave(REDUCEDRATIOINDEX, ilLoop) = smSave(INPUTRATIOINDEX, ilLoop)
            Next ilLoop
            rbcUse(1).Value = True
        End If
        'Compute percent
        llTotal = 0
        For ilLoop = LBONE To UBound(smSave, 2) Step 1
            llTotal = llTotal + Val(smSave(REDUCEDRATIOINDEX, ilLoop))
        Next ilLoop
        For ilLoop = LBONE To UBound(smSave, 2) Step 1
            smSave(REDUCEDPERCENTINDEX, ilLoop) = gDivStr(gMulStr(smSave(REDUCEDRATIOINDEX, ilLoop), "100.0"), str$(llTotal))
        Next ilLoop
        'plcSpots(1).Caption = Trim$(Str$(llTotal))
        smSpotsValue(1) = Trim$(str$(llTotal))
        plcSpots(1).Cls
        plcSpots(1).CurrentX = 0
        plcSpots(1).CurrentY = 0
        plcSpots(1).Print Trim$(str$(llTotal))
        'Compute Altered
        For ilLoop = LBONE To UBound(smSave, 2) Step 1
            smSave(ALTEREDRATIOINDEX, ilLoop) = gDivStr(smSave(INPUTRATIOINDEX, ilLoop), str$(llSmallest))
        Next ilLoop
        llTotal = 0
        For ilLoop = LBONE To UBound(smSave, 2) Step 1
            llTotal = llTotal + Val(smSave(ALTEREDRATIOINDEX, ilLoop))
        Next ilLoop
        For ilLoop = LBONE To UBound(smSave, 2) Step 1
            smSave(ALTEREDPERCENTINDEX, ilLoop) = gDivStr(gMulStr(smSave(ALTEREDRATIOINDEX, ilLoop), "100.0"), str$(llTotal))
        Next ilLoop
        'plcSpots(2).Caption = Trim$(Str$(llTotal))
        smSpotsValue(2) = Trim$(str$(llTotal))
        plcSpots(2).Cls
        plcSpots(2).CurrentX = 0
        plcSpots(2).CurrentY = 0
        plcSpots(2).Print Trim$(str$(llTotal))
        For ilField = 3 To 7 Step 1
            For ilLoop = LBONE To UBound(smSave, 2) Step 1
                slStr = smSave(ilField, ilLoop)
                gSetShow pbcInstructions, slStr, tmCtrls(ilField)
                smShow(ilField, ilLoop) = tmCtrls(ilField).sShow
            Next ilLoop
        Next ilField
    End If
    Screen.MousePointer = vbDefault
    mSetCommands
    pbcInstructions_Paint
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    If (imRowNo < vbcInstructions.Value) Or (imRowNo >= vbcInstructions.Value + vbcInstructions.LargeChange + 1) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case INPUTRATIOINDEX 'Input Ratio
            edcDropDown.Width = tmCtrls(INPUTRATIOINDEX).fBoxW
            edcDropDown.MaxLength = 3
            gMoveTableCtrl pbcInstructions, edcDropDown, tmCtrls(INPUTRATIOINDEX).fBoxX, tmCtrls(INPUTRATIOINDEX).fBoxY + (imRowNo - vbcInstructions.Value) * (fgBoxGridH + 15)
            edcDropDown.Text = smSave(2, imRowNo)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGeneratePattern                *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Generate copy pattern from     *
'*                      input                          *
'*            Note: Step Description                   *
'*                    1  Create empty cells (total # of*
'*                       instructions to be gen.)      *
'*                    2  Find instruction with smallest*
'*                       distribution count            *
'*                    3  Evenly distribute instruction *
'*                       into empty cells              *
'*                                                     *
'*******************************************************
Private Function mGeneratePattern() As Integer
    Dim ilSmallest As Integer
    Dim ilSmallestIndex As Integer
    Dim ilTotal As Integer
    Dim ilLoop As Integer
    Dim ilNoUsed As Integer
    Dim ilDone As Integer
    Dim ilNoBeforeInc As Integer
    Dim ilField As Integer
    Dim ilInsert As Integer
    Dim ilPlaced As Integer
    Dim ilAdded As Integer
    Dim ilAddIndex As Integer
    Dim ilFromIndex As Integer
    Dim ilRes As Integer
    If rbcUse(0).Value Then
        ilField = REDUCEDRATIOINDEX
    Else
        ilField = ALTEREDRATIOINDEX
    End If
    ilTotal = 0
    For ilLoop = LBONE To UBound(smSave, 2) Step 1
        ilTotal = ilTotal + Val(smSave(ilField, ilLoop))
    Next ilLoop
    If ilTotal <= 0 Then
        ilRes = MsgBox("Define Ratio or Press Cancel", vbOKOnly + vbExclamation, "Warning")
        mGeneratePattern = False
        Exit Function
    End If
    mGeneratePattern = True
    ReDim smGenerateCode(0 To ilTotal) As String
    ReDim smWorkArea(0 To ilTotal) As String
    For ilLoop = LBONE To UBound(smGenerateCode) Step 1
        smGenerateCode(ilLoop) = ""
    Next ilLoop
    ilDone = False
    Do
        ilSmallest = -1
        For ilLoop = LBONE To UBound(smSave, 2) Step 1
            If smCode(ilLoop) <> "" Then
                If ilSmallest = -1 Then
                    ilSmallest = Val(smSave(ilField, ilLoop))
                    ilSmallestIndex = ilLoop
                Else
                    If Val(smSave(ilField, ilLoop)) < ilSmallest Then
                        ilSmallest = Val(smSave(ilField, ilLoop))
                        ilSmallestIndex = ilLoop
                    End If
                End If
            End If
        Next ilLoop
        If ilSmallest = -1 Then
            ilDone = True
            Exit Do
        End If
        ilNoUsed = 0
        For ilLoop = LBONE To UBound(smGenerateCode) Step 1
            If smGenerateCode(ilLoop) <> "" Then
                ilNoUsed = ilNoUsed + 1
            End If
        Next ilLoop
        If ilNoUsed = 0 Then
            ilInsert = ilSmallest
            For ilLoop = LBONE To ilSmallest Step 1
                smGenerateCode(ilLoop) = smCode(ilSmallestIndex)
            Next ilLoop
            smCode(ilSmallestIndex) = ""
        Else
            ilInsert = ilSmallest \ ilNoUsed
            ilNoBeforeInc = ilNoUsed - (ilSmallest Mod ilNoUsed)
            ilAddIndex = LBONE  '1
            ilFromIndex = LBONE '1
            ilPlaced = 0
            ilAdded = 0
            Do
                If ilFromIndex <= ilNoUsed Then
                    smWorkArea(ilAddIndex) = smGenerateCode(ilFromIndex)
                    ilFromIndex = ilFromIndex + 1
                    ilAddIndex = ilAddIndex + 1
                End If
                If ilAdded <= ilSmallest Then
                    For ilLoop = LBONE To ilInsert Step 1
                        smWorkArea(ilAddIndex) = smCode(ilSmallestIndex)
                        ilAddIndex = ilAddIndex + 1
                        ilAdded = ilAdded + 1
                        If ilAdded >= ilSmallest Then
                            smCode(ilSmallestIndex) = ""
                            Exit For
                        End If
                    Next ilLoop
                End If
                ilPlaced = ilPlaced + 1
                If ilPlaced = ilNoBeforeInc Then
                    ilInsert = ilInsert + 1
                End If
            Loop While (ilFromIndex <= ilNoUsed) And (ilAdded < ilSmallest)
            For ilLoop = LBONE To UBound(smGenerateCode) Step 1
                smGenerateCode(ilLoop) = smWorkArea(ilLoop)
            Next ilLoop
        End If
    Loop While Not ilDone
End Function
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
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imFirstActivate = True
    smInputInstCode = sgDoneMsg
    imTerminate = False
    imBypassFocus = False
    imSettingValue = False
    imBoxNo = -1 'Initialize current Box to N/A
    imChgMode = False
    imBSMode = False
    smSpotsValue(0) = ""
    smSpotsValue(1) = ""
    smSpotsValue(2) = ""
    smSpotsValue(3) = ""
    mInitBox
    mXFerRecToCtrl
    CopyRato.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone CopyRato
    'CopyRato.Show
    'gCenterModalForm CopyRato
    rbcSort(1).Value = True
    'Traffic!plcHelp.Caption = ""
    Screen.MousePointer = vbDefault
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
    flTextHeight = pbcInstructions.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcInstructions.Move 210, 225, pbcInstructions.Width + vbcInstructions.Width + fgPanelAdj   ', pbcInstructions.Height + fgPanelAdj
    pbcInstructions.Move plcInstructions.Left + fgBevelX    ', plcDates.Top + fgBevelY
    vbcInstructions.Move pbcInstructions.Left + pbcInstructions.Width + 15, pbcInstructions.Top
    'pbcArrow.Move plcDates.Left - pbcArrow.Width - 15   'Set arrow
    'Instruction
    gSetCtrl tmCtrls(INSTRUCTIONINDEX), 30, 375, 3435, fgBoxGridH
    tmCtrls(INSTRUCTIONINDEX).iReq = False
    'Input
    gSetCtrl tmCtrls(INPUTRATIOINDEX), 3555, tmCtrls(INSTRUCTIONINDEX).fBoxY, 510, fgBoxGridH
    tmCtrls(INPUTRATIOINDEX).iReq = True
    gSetCtrl tmCtrls(INPUTPERCENTINDEX), 4080, tmCtrls(INSTRUCTIONINDEX).fBoxY, 510, fgBoxGridH
    tmCtrls(INPUTPERCENTINDEX).iReq = False
    'Reduced
    gSetCtrl tmCtrls(REDUCEDRATIOINDEX), 4695, tmCtrls(INSTRUCTIONINDEX).fBoxY, 510, fgBoxGridH
    tmCtrls(REDUCEDRATIOINDEX).iReq = False
    gSetCtrl tmCtrls(REDUCEDPERCENTINDEX), 5220, tmCtrls(INSTRUCTIONINDEX).fBoxY, 510, fgBoxGridH
    tmCtrls(REDUCEDPERCENTINDEX).iReq = False
    'Altered
    gSetCtrl tmCtrls(ALTEREDRATIOINDEX), 5835, tmCtrls(INSTRUCTIONINDEX).fBoxY, 510, fgBoxGridH
    tmCtrls(ALTEREDRATIOINDEX).iReq = False
    gSetCtrl tmCtrls(ALTEREDPERCENTINDEX), 6360, tmCtrls(INSTRUCTIONINDEX).fBoxY, 510, fgBoxGridH
    tmCtrls(ALTEREDPERCENTINDEX).iReq = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set button status              *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
    Dim ilAllDefined As Integer
    Dim ilLoop As Integer
    If rbcUse(0).Value Or rbcUse(1).Value Then
        ilAllDefined = True
        For ilLoop = LBONE To UBound(smSave, 2) Step 1
            If Len(Trim$(smSave(INPUTRATIOINDEX, ilLoop))) = 0 Then
                ilAllDefined = False
                Exit For
            End If
        Next ilLoop
        If ilAllDefined Then
            cmcGenerate.Enabled = True
        Else
            cmcGenerate.Enabled = False
        End If
    Else
        cmcGenerate.Enabled = False
    End If
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    If (imRowNo < vbcInstructions.Value) Or (imRowNo >= vbcInstructions.Value + vbcInstructions.LargeChange + 1) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case INPUTRATIOINDEX 'Start time
            If (edcDropDown.Visible) And (edcDropDown.Enabled) Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case Else
            pbcClickFocus.SetFocus
    End Select
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    If (imRowNo < vbcInstructions.Value) Or (imRowNo >= vbcInstructions.Value + vbcInstructions.LargeChange + 1) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case INPUTRATIOINDEX
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smSave(2, imRowNo) = slStr
            gSetShow pbcInstructions, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            mComputeValues
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
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
    sgDoneMsg = smMatchCode
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload CopyRato
    igManUnload = NO
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
    Dim slStr As String
    Dim ilCount As Integer
    Dim ilRet As Integer
    Dim ilField As Integer
    ilRet = gParseItem(smInputInstCode, 1, "|", slStr)    '# in first 4 weeks
    'plcSpots(3).Caption = slStr
    smSpotsValue(3) = Trim$(slStr)
    plcSpots(3).Cls
    plcSpots(3).CurrentX = 0
    plcSpots(3).CurrentY = 0
    plcSpots(3).Print slStr
    ilRet = gParseItem(smInputInstCode, 2, "|", slStr)    'Count
    ilCount = Val(slStr)
    ReDim smInstruction(0 To ilCount) As String
    ReDim smCode(0 To ilCount) As String
    For ilLoop = LBONE To ilCount Step 1
        ilRet = gParseItem(smInputInstCode, ilLoop + 2, "|", slStr)  'Count
        ilRet = gParseItem(slStr, 1, "\", smInstruction(ilLoop))  'Instruction
        ilRet = gParseItem(slStr, 2, "\", smCode(ilLoop))  'Instruction
    Next ilLoop
    ReDim smSave(0 To 7, 0 To ilCount) As String
    ReDim smShow(0 To 7, 0 To ilCount) As String
    For ilLoop = LBONE To ilCount Step 1
        smSave(1, ilLoop) = smInstruction(ilLoop)
        slStr = smSave(1, ilLoop)
        gSetShow pbcInstructions, slStr, tmCtrls(INSTRUCTIONINDEX)
        smShow(INSTRUCTIONINDEX, ilLoop) = tmCtrls(INSTRUCTIONINDEX).sShow
    Next ilLoop
    For ilField = 2 To 7 Step 1
        For ilLoop = LBONE To ilCount Step 1
            smSave(ilField, ilLoop) = ""
            slStr = smSave(ilField, ilLoop)
            gSetShow pbcInstructions, slStr, tmCtrls(ilField)
            smShow(ilField, ilLoop) = tmCtrls(ilField).sShow
        Next ilLoop
    Next ilField
    vbcInstructions.Min = LBONE 'LBound(smShow, 2)
    If UBound(smShow, 2) <= vbcInstructions.LargeChange Then
        vbcInstructions.Max = LBONE 'LBound(smShow, 2)
    Else
        vbcInstructions.Max = UBound(smShow, 2) - vbcInstructions.LargeChange
    End If
    vbcInstructions.Value = vbcInstructions.Min
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase smInstruction
    Erase smCode
    Erase smSave
    Erase smShow
    
    Set CopyRato = Nothing   'Remove data segment
    
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcInstructions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    ilCompRow = vbcInstructions.LargeChange + 1
    If UBound(smSave, 2) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smSave, 2)
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    If ilBox <> INPUTRATIOINDEX Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    ilRowNo = ilRow + vbcInstructions.Value - 1
                    mSetShow imBoxNo
                    imRowNo = ilRowNo
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mSetFocus imBoxNo
End Sub
Private Sub pbcInstructions_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer

    ilStartRow = vbcInstructions.Value  'Top location
    ilEndRow = vbcInstructions.Value + vbcInstructions.LargeChange
    If ilEndRow > UBound(smSave, 2) Then
        ilEndRow = UBound(smSave, 2)
    End If
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            pbcInstructions.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcInstructions.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            pbcInstructions.Print smShow(ilBox, ilRow)
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    mSetShow imBoxNo
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imSettingValue = True
                vbcInstructions.Value = vbcInstructions.Min
                imSettingValue = False
                If UBound(smSave, 2) <= vbcInstructions.LargeChange Then  'was <=
                    vbcInstructions.Max = LBONE 'LBound(smSave, 2)
                Else
                    vbcInstructions.Max = UBound(smSave, 2) - vbcInstructions.LargeChange
                End If
                imRowNo = 1
                imSettingValue = True
                vbcInstructions.Value = vbcInstructions.Min
                imSettingValue = False
                ilBox = 1
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            Case INSTRUCTIONINDEX 'Time (first control within header)
                ilBox = ALTEREDPERCENTINDEX
                If imRowNo <= 1 Then
                    imBoxNo = -1
                    imRowNo = -1
                    If cmcGenerate.Enabled Then
                        cmcGenerate.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                imRowNo = imRowNo - 1
                If imRowNo < vbcInstructions.Value Then
                    imSettingValue = True
                    vbcInstructions.Value = vbcInstructions.Value - 1
                    imSettingValue = False
                End If
                ilFound = False
            Case INPUTPERCENTINDEX
                ilBox = INPUTRATIOINDEX
                ilFound = True
            Case Else
                ilBox = ilBox - 1
                ilFound = False
        End Select
    Loop While Not ilFound
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    ilBox = imBoxNo
    mSetShow imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imRowNo = UBound(smSave, 2)
                imSettingValue = True
                If imRowNo - 1 <= vbcInstructions.LargeChange Then
                    vbcInstructions.Value = vbcInstructions.Min
                Else
                    vbcInstructions.Value = imRowNo - vbcInstructions.LargeChange - 1
                End If
                imSettingValue = False
                ilBox = INPUTRATIOINDEX
                ilFound = True
            Case ALTEREDPERCENTINDEX
                imRowNo = imRowNo + 1
                If imRowNo > UBound(smSave, 2) Then
                    imBoxNo = -1
                    imRowNo = -1
                    If cmcGenerate.Enabled Then
                        cmcGenerate.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                If imRowNo > vbcInstructions.Value + vbcInstructions.LargeChange Then
                    imSettingValue = True
                    vbcInstructions.Value = vbcInstructions.Value + 1
                    imSettingValue = False
                End If
                ilBox = INSTRUCTIONINDEX
                ilFound = False
            Case INSTRUCTIONINDEX
                ilBox = INPUTRATIOINDEX
                ilFound = True
            Case Else 'Last control within header
                ilBox = ilBox + 1
                ilFound = False
        End Select
    Loop While Not ilFound
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub plcInstructions_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcSpots_Paint(Index As Integer)
    plcSpots(Index).Cls
    plcSpots(Index).CurrentX = 0
    plcSpots(Index).CurrentY = 0
    plcSpots(Index).Print smSpotsValue(Index)

End Sub

Private Sub vbcInstructions_Change()
    If imSettingValue Then
        pbcInstructions.Cls
        pbcInstructions_Paint
        imSettingValue = False
    Else
        mSetShow imBoxNo
        pbcInstructions.Cls
        pbcInstructions_Paint
        mEnableBox imBoxNo
    End If
End Sub
Private Sub vbcInstructions_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Copy Ratios"
End Sub
Private Sub plcSort_Paint()
    plcSort.CurrentX = 0
    plcSort.CurrentY = 0
    plcSort.Print "Sort Least Used To"
End Sub
Private Sub plcUse_Paint()
    plcUse.CurrentX = 0
    plcUse.CurrentY = 0
    plcUse.Print "Use"
End Sub
