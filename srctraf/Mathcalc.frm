VERSION 5.00
Begin VB.Form MathCalc 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculator"
   ClientHeight    =   3420
   ClientLeft      =   2130
   ClientTop       =   1665
   ClientWidth     =   1800
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
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3420
   ScaleWidth      =   1800
   Begin VB.PictureBox plcShow 
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   1740
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1800
      Begin VB.Label plcResult 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   1680
      End
   End
   Begin VB.PictureBox plcCalculator 
      ForeColor       =   &H00000000&
      Height          =   3045
      Left            =   0
      ScaleHeight     =   2985
      ScaleWidth      =   1740
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   405
      Width           =   1800
      Begin VB.PictureBox pbcCover 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   75
         Left            =   0
         ScaleHeight     =   75
         ScaleWidth      =   1725
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   945
         Width           =   1725
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
         Left            =   75
         ScaleHeight     =   165
         ScaleWidth      =   90
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2685
         Width           =   90
      End
      Begin VB.CommandButton cmcDone 
         Appearance      =   0  'Flat
         Caption         =   "&Done"
         Height          =   285
         HelpContextID   =   1
         Left            =   480
         TabIndex        =   3
         Top             =   2655
         Width           =   795
      End
      Begin VB.PictureBox pbcCalculator 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1680
         Left            =   30
         Picture         =   "Mathcalc.frx":0000
         ScaleHeight     =   1680
         ScaleWidth      =   1680
         TabIndex        =   2
         Top             =   945
         Width           =   1680
         Begin VB.Image imcOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   345
            Top             =   345
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "Mathcalc.frx":1A0E
            Top             =   840
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Label plcImmediate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   45
         TabIndex        =   11
         Top             =   645
         Width           =   1635
      End
      Begin VB.Label plcSecond 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   60
         TabIndex        =   10
         Top             =   345
         Width           =   1620
      End
      Begin VB.Label plcFirst 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   60
         TabIndex        =   9
         Top             =   75
         Width           =   1620
      End
      Begin VB.Image imcBorder 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   585
         Left            =   30
         Top             =   60
         Width           =   1695
      End
      Begin VB.Line lncTotal 
         BorderWidth     =   2
         X1              =   60
         X2              =   1725
         Y1              =   645
         Y2              =   645
      End
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
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
      Height          =   285
      Left            =   915
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3210
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
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
      Height          =   285
      Left            =   615
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3210
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
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
      Height          =   285
      Left            =   360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3210
      Visible         =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "MathCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Mathcalc.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: MathCalc.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Calculator input screen code (same as Calc exceopt:
'   Form Active; DeActivate; KeyUp events and Calc. changed to MathCalc
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim fmDeltaX As Single
Dim fmDeltaY As Single
Dim fmX As Single
Dim fmY As Single
Dim fmMaxY As Single
Dim fmMaxX As Single
Dim fmMinY As Single
Dim fmMinX As Single
Dim imNoRows As Integer
Dim imNoCols As Integer
Dim smFirstTerm As String
Dim imFirstType As Integer  '-1=undefined; 0=Decimal; 1=Time; 2= Length
Dim smSecondTerm As String
Dim imSecondType As Integer  '-1=undefined; 0=Decimal; 1=Time; 2= Length
Dim smImmediate As String
Dim imImmediateType As Integer  '-1=undefined; 0=Decimal; 1=Time; 2= Length
Dim imOperator As Integer  '0=/;1=*;2=-;3=+
Dim imResultType As Integer  '-1=undefined; 0=Decimal; 1=Time; 2= Length
Dim imClearInput As Integer 'True- when first key entered- clear input
Dim imFirstFocus As Integer
Dim imTerminate As Integer

Private Sub cmcDone_Click()
    fgCalcLeft = MathCalc.Left
    fgCalcTop = MathCalc.Top
'    igCalcActive = 0 'False
    Screen.MousePointer = vbDefault
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Exit Sub
    End If
    imFirstActivate = False
 '   igCalcActive = 1 'True
'    If (Forms.Count = igNoMinForms + igCalActive + igCalcActive) And Not igShowPicture Then 'Show basic 10 if only this form to showing
'        Traffic!pbcMsgArea.SetFocus
'    End If
'    DoEvents    'Process events so pending keys are not sent to this
                'form when keypreview turn on
'    Calc.KeyPreview = True  'To get Alt J and Alt L keys
    MathCalc.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub Form_Deactivate()
'    Calc.KeyPreview = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    mTerminate
    Exit Sub
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    If ((Shift And vbAltMask) > 0) And (KeyCode = 74) Then    'J=74
'        Calc.KeyPreview = False
'        Traffic!gpcBasicWnd.Value = True   'Button up and unload
'    End If
'    If ((Shift And vbAltMask) > 0) And (KeyCode = 76) Then    'L=76
'        Calc.KeyPreview = False
'        Traffic!gpcAuxWnd.Value = True   'Button up and unload
'    End If

End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set MathCalc = Nothing   'Remove data segment
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDoMath                         *
'*                                                     *
'*             Created:10/27/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Proform the math operation     *
'*                                                     *
'*******************************************************
Private Sub mDoMath(slStore As String, plcCtrl As Control, ilType As Integer)
'
'   mDoMath slStore, plcCtrl, ilType
'   Where:
'       slStore (O)- Variable to store results into
'       plcCtrl (O)- Control to show results within
'       ilType (O)- Result type (0=number, 1=time, 2=length)
'
    Dim clFirst As Currency
    Dim clSecond As Currency
    If (imFirstType <= 0) And (imSecondType <= 0) Then
        Select Case imOperator
            Case 0  '/
                If Val(smSecondTerm) <> 0 Then
                    slStore = gDivStr(smFirstTerm, smSecondTerm)
                    plcCtrl.Caption = slStore
                    ilType = 0
                Else
                    slStore = ""
                    ilType = -1
                    plcCtrl.Caption = "Error"
                    Exit Sub
                End If
            Case 1  '*
                slStore = gMulStr(smFirstTerm, smSecondTerm)
                plcCtrl.Caption = slStore
                ilType = 0
            Case 2  '-
                slStore = gSubStr(smFirstTerm, smSecondTerm)
                plcCtrl.Caption = slStore
                ilType = 0
            Case 3  '+
                slStore = gAddStr(smFirstTerm, smSecondTerm)
                plcCtrl.Caption = slStore
                ilType = 0
        End Select
    Else
        If imFirstType <= 0 Then
            clFirst = Val(smFirstTerm)
        ElseIf imFirstType = 1 Then
            clFirst = gTimeToCurrency(smFirstTerm, False)
        Else
            clFirst = gLengthToCurrency(smFirstTerm)
        End If
        If imSecondType <= 0 Then
            clSecond = Val(smSecondTerm)
        ElseIf imSecondType = 1 Then
            clSecond = gTimeToCurrency(smSecondTerm, False)
        Else
            clSecond = gLengthToCurrency(smSecondTerm)
        End If
        Select Case imOperator
            Case 0  '/
                If clSecond <> 0 Then
                    clFirst = clFirst / clSecond
                Else
                    slStore = ""
                    ilType = -1
                    plcCtrl.Caption = "Error"
                    Exit Sub
                End If
            Case 1  '*
                clFirst = clFirst * clSecond
            Case 2  '-
                clFirst = clFirst - clSecond
            Case 3  '+
                clFirst = clFirst + clSecond
        End Select
        If (imFirstType = 1) And (imSecondType = 1) Then
            slStore = gCurrencyToLength(clFirst)
            plcCtrl.Caption = slStore
            ilType = 2
        ElseIf (imFirstType = 1) Or (imSecondType = 1) Then
            slStore = gCurrencyToTime(clFirst)
            plcCtrl.Caption = slStore
            ilType = 1
        Else
            slStore = gCurrencyToLength(clFirst)
            plcCtrl.Caption = slStore
            ilType = 2
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:10/27/93      By:D. LeVine      *
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
    Dim illoop As Integer
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    imFirstFocus = True
    imNoRows = 6
    imNoCols = 5
    fmDeltaX = 330
    fmDeltaY = 255
    fmMinY = 90
    fmMinX = 15
    fmMaxY = fmMinY
    For illoop = 1 To imNoRows - 1 Step 1
        fmMaxY = fmMaxY + fmDeltaY
    Next illoop
    fmMaxX = fmMinX
    For illoop = 1 To imNoCols - 1 Step 1
        fmMaxX = fmMaxX + fmDeltaX
    Next illoop
    fmX = -1
    fmY = -1
    smFirstTerm = ""
    imFirstType = -1
    smSecondTerm = ""
    imSecondType = -1
    smImmediate = ""
    imImmediateType = -1
    imOperator = -1
    imResultType = -1
    imClearInput = True
    imcInv.Visible = False
    plcFirst.Caption = ""
    plcSecond.Caption = ""
    plcResult.Caption = ""
    MathCalc.Move fgCalcLeft, fgCalcTop
    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer

    slCommand = sgCommandStr    'Command$
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone MathCalc, slStr, ilTestSystem
    'If Not igStdAloneMode Then
        ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get user name
        If slStr <> "" Then
            fgCalcLeft = Val(slStr)
        End If
        ilRet = gParseItem(slCommand, 4, "\", slStr)    'Get user name
        If slStr <> "" Then
            fgCalcTop = Val(slStr)
        End If
   'End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mProcessInput                   *
'*                                                     *
'*             Created:10/27/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process input key              *
'*                                                     *
'*******************************************************
Private Sub mProcessInput(ilRowNo As Integer, ilColNo As Integer)
'
'   mProcessInput
'   Where:
'       ilRowNo (I)- Row number of key pressed (1-6)
'       ilColNo (I)- Column number of key pressed (1-5)
'
    Dim ilPos As Integer
    Dim slStr As String
    Dim clFirst As Currency
    Dim clSecond As Currency
    If imClearInput Then
        plcFirst.Caption = ""
        plcSecond.Caption = ""
        plcImmediate.Caption = ""
        imClearInput = False
    End If
    fmX = fmMinX + (ilColNo - 1) * fmDeltaX
    fmY = fmMinY + (ilRowNo - 1) * fmDeltaY
    imcOutline.Move fmX - 15, fmY - 15
    Select Case ilRowNo
        Case 1  'Row number
            Select Case ilColNo
                Case 1  'Move to result window
                    plcResult.Caption = smImmediate
                    Clipboard.SetText plcResult.Caption
                    imResultType = imImmediateType
                Case 2  'Hour
                    If imOperator = -1 Then 'On first term
                        If imFirstType = -1 Then    'If undefined, then OK
                            plcFirst.Caption = plcFirst.Caption & "h"
                            imFirstType = 2 'Length
                        Else
                            Beep
                        End If
                    Else    'Second term
                        If imSecondType = -1 Then
                            plcSecond.Caption = plcSecond.Caption & "h"
                            imSecondType = 2 'Length
                        Else
                            Beep
                        End If
                    End If
                Case 3
                    If imOperator = -1 Then 'On first term
                        If (imFirstType = -1) Or (imFirstType = 2) Then
                            If (InStr(plcFirst.Caption, "m") = 0) And (InStr(plcFirst.Caption, "s") = 0) Then
                                plcFirst.Caption = plcFirst.Caption & "m"
                                imFirstType = 2 'Length
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    Else    'Second term
                        If (imSecondType = -1) Or (imSecondType = 2) Then
                            If (InStr(plcSecond.Caption, "m") = 0) And (InStr(plcSecond.Caption, "s") = 0) Then
                                plcSecond.Caption = plcSecond.Caption & "m"
                                imSecondType = 2 'Length
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    End If
                Case 4
                    If imOperator = -1 Then 'On first term
                        If (imFirstType = -1) Or (imFirstType = 2) Then
                            If InStr(plcFirst.Caption, "s") = 0 Then
                                plcFirst.Caption = plcFirst.Caption & "s"
                                imFirstType = 2 'Length
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    Else    'Second term
                        If (imSecondType = -1) Or (imSecondType = 2) Then
                            If InStr(plcSecond.Caption, "s") = 0 Then
                                plcSecond.Caption = plcSecond.Caption & "s"
                                imSecondType = 2 'Length
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    End If
                Case 5  '/
                    If (imOperator = -1) Or (Len(plcSecond.Caption) <= 2) Then
                        imOperator = 0  'divide
                        plcSecond.Caption = "/ "
                        If Len(plcFirst.Caption) > 0 Then
                            smFirstTerm = plcFirst.Caption
                        Else
                            smFirstTerm = smImmediate  'chain
                            imFirstType = imImmediateType
                            plcFirst.Caption = smFirstTerm
                        End If
                    Else
                        If Len(plcSecond.Caption) > 2 Then
                            smSecondTerm = Mid$(plcSecond.Caption, 3)
                            mDoMath smFirstTerm, plcFirst, imFirstType
                            smImmediate = smFirstTerm
                            imImmediateType = imFirstType
                            plcImmediate.Caption = ""
                            imOperator = 0
                            plcSecond.Caption = "/ "
                            smSecondTerm = ""
                            imSecondType = -1
                        Else
                            Beep
                        End If
                    End If
            End Select
        Case 2  'Row number
            Select Case ilColNo
                Case 1  '+ result
                    If (imImmediateType <= 0) And (imResultType <= 0) Then
                        slStr = plcResult.Caption
                        plcResult.Caption = gAddStr(slStr, smImmediate)
                        Clipboard.SetText plcResult.Caption
                        imResultType = 0
                    Else
                        If imImmediateType <= 0 Then
                            clFirst = Val(smImmediate)
                        ElseIf imImmediateType = 1 Then
                            clFirst = gTimeToCurrency(smImmediate, False)
                        Else
                            clFirst = gLengthToCurrency(smImmediate)
                        End If
                        slStr = plcResult.Caption
                        If imResultType <= 0 Then
                            clSecond = Val(slStr)
                        ElseIf imResultType = 1 Then
                            clSecond = gTimeToCurrency(slStr, False)
                        Else
                            clSecond = gLengthToCurrency(slStr)
                        End If
                        clFirst = clFirst + clSecond
                        If (imImmediateType = 1) Or (imResultType = 1) Then
                            plcResult.Caption = gCurrencyToTime(clFirst)
                            imResultType = 1
                        Else
                            plcResult.Caption = gCurrencyToLength(clFirst)
                            imResultType = 2
                        End If
                        Clipboard.SetText plcResult.Caption
                    End If
                Case 2
                    If imOperator = -1 Then
                        plcFirst.Caption = plcFirst.Caption & "7"
                    Else
                        plcSecond.Caption = plcSecond.Caption & "7"
                    End If
                Case 3
                    If imOperator = -1 Then
                        plcFirst.Caption = plcFirst.Caption & "8"
                    Else
                        plcSecond.Caption = plcSecond.Caption & "8"
                    End If
                Case 4
                    If imOperator = -1 Then
                        plcFirst.Caption = plcFirst.Caption & "9"
                    Else
                        plcSecond.Caption = plcSecond.Caption & "9"
                    End If
                Case 5  '*
                    If (imOperator = -1) Or (Len(plcSecond.Caption) <= 2) Then
                        imOperator = 1  'multiply
                        plcSecond.Caption = "* "
                        If Len(plcFirst.Caption) > 0 Then
                            smFirstTerm = plcFirst.Caption
                        Else
                            smFirstTerm = smImmediate  'chain
                            imFirstType = imImmediateType
                            plcFirst.Caption = smFirstTerm
                        End If
                    Else
                        If Len(plcSecond.Caption) > 2 Then
                            smSecondTerm = Mid$(plcSecond.Caption, 3)
                            mDoMath smFirstTerm, plcFirst, imFirstType
                            smImmediate = smFirstTerm
                            imImmediateType = imFirstType
                            plcImmediate.Caption = ""
                            imOperator = 1
                            plcSecond.Caption = "* "
                            smSecondTerm = ""
                            imSecondType = -1
                        Else
                            Beep
                        End If
                    End If
            End Select
        Case 3  'Row number
            Select Case ilColNo
                Case 1
                    If imOperator = -1 Then
                        ilPos = InStr(plcFirst.Caption, "-")
                        If ilPos = 1 Then   'Remove negative sign
                            plcFirst.Caption = Mid(plcFirst.Caption, 2)
                        Else
                            plcFirst.Caption = "-" & plcFirst.Caption
                        End If
                    Else
                        ilPos = InStr(3, plcSecond.Caption, "-")
                        If ilPos <> 0 Then   'Remove negative sign
                            plcSecond.Caption = Left$(plcSecond.Caption, ilPos - 1) & Mid$(plcSecond.Caption, ilPos + 1)
                        Else
                            plcSecond.Caption = Left$(plcSecond.Caption, 2) & "-" & Mid$(plcSecond.Caption, 3)
                        End If
                    End If
                Case 2
                    If imOperator = -1 Then
                        plcFirst.Caption = plcFirst.Caption & "4"
                    Else
                        plcSecond.Caption = plcSecond.Caption & "4"
                    End If
                Case 3
                    If imOperator = -1 Then
                        plcFirst.Caption = plcFirst.Caption & "5"
                    Else
                        plcSecond.Caption = plcSecond.Caption & "5"
                    End If
                Case 4
                    If imOperator = -1 Then
                        plcFirst.Caption = plcFirst.Caption & "6"
                    Else
                        plcSecond.Caption = plcSecond.Caption & "6"
                    End If
                Case 5  '-
                    If (imOperator = -1) Or (Len(plcSecond.Caption) <= 2) Then
                        imOperator = 2  'subtraction
                        plcSecond.Caption = "- " & plcSecond.Caption
                        If Len(plcFirst.Caption) > 0 Then
                            smFirstTerm = plcFirst.Caption
                        Else
                            smFirstTerm = smImmediate  'chain
                            imFirstType = imImmediateType
                            plcFirst.Caption = smFirstTerm
                        End If
                    Else
                        If Len(plcSecond.Caption) > 2 Then
                            smSecondTerm = Mid$(plcSecond.Caption, 3)
                            mDoMath smFirstTerm, plcFirst, imFirstType
                            smImmediate = smFirstTerm
                            imImmediateType = imFirstType
                            plcImmediate.Caption = ""
                            imOperator = 2
                            plcSecond.Caption = "- "
                            smSecondTerm = ""
                            imSecondType = -1
                        Else
                            Beep
                        End If
                    End If
            End Select
        Case 4  'Row number
            Select Case ilColNo
                Case 1
                    If imOperator = -1 Then
                        If Len(plcFirst.Caption) > 0 Then
                            plcFirst.Caption = Left$(plcFirst.Caption, Len(plcFirst.Caption) - 1)
                        End If
                    Else
                        If Len(plcSecond.Caption) > 0 Then
                            plcSecond.Caption = Left$(plcSecond.Caption, Len(plcSecond.Caption) - 1)
                        End If
                    End If
                Case 2
                    If imOperator = -1 Then
                        plcFirst.Caption = plcFirst.Caption & "1"
                    Else
                        plcSecond.Caption = plcSecond.Caption & "1"
                    End If
                Case 3
                    If imOperator = -1 Then
                        plcFirst.Caption = plcFirst.Caption & "2"
                    Else
                        plcSecond.Caption = plcSecond.Caption & "2"
                    End If
                Case 4
                    If imOperator = -1 Then
                        plcFirst.Caption = plcFirst.Caption & "3"
                    Else
                        plcSecond.Caption = plcSecond.Caption & "3"
                    End If
                Case 5  '+
                    If (imOperator = -1) Or (Len(plcSecond.Caption) <= 2) Then
                        imOperator = 3  'addition
                        plcSecond.Caption = "+ "
                        If Len(plcFirst.Caption) > 0 Then
                            smFirstTerm = plcFirst.Caption
                        Else
                            smFirstTerm = smImmediate  'chain
                            imFirstType = imImmediateType
                            plcFirst.Caption = smFirstTerm
                        End If
                    Else
                        If Len(plcSecond.Caption) > 2 Then
                            smSecondTerm = Mid$(plcSecond.Caption, 3)
                            mDoMath smFirstTerm, plcFirst, imFirstType
                            smImmediate = smFirstTerm
                            imImmediateType = imFirstType
                            plcImmediate.Caption = ""
                            imOperator = 3
                            plcSecond.Caption = "+ "
                            smSecondTerm = ""
                            imSecondType = -1
                        Else
                            Beep
                        End If
                    End If
            End Select
        Case 5  'Row number
            Select Case ilColNo
                Case 1
                    smImmediate = ""
                    imImmediateType = -1
                    plcFirst.Caption = ""
                    plcSecond.Caption = ""
                    plcImmediate.Caption = ""
                    imOperator = -1
                    smFirstTerm = ""
                    imFirstType = -1
                    smSecondTerm = ""
                    imSecondType = -1
                Case 2
                    If imOperator = -1 Then
                        plcFirst.Caption = plcFirst.Caption & "0"
                    Else
                        plcSecond.Caption = plcSecond.Caption & "0"
                    End If
                Case 3
                    If imOperator = -1 Then
                        plcFirst.Caption = plcFirst.Caption & "00"
                    Else
                        plcSecond.Caption = plcSecond.Caption & "00"
                    End If
                Case 4
                    If imOperator = -1 Then 'On first term
                        If imFirstType <= 0 Then
                            If InStr(plcFirst.Caption, ".") = 0 Then
                                plcFirst.Caption = plcFirst.Caption & "."
                                imFirstType = 0
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    Else    'Second term
                        If imSecondType <= 0 Then
                            If InStr(plcSecond.Caption, ".") = 0 Then
                                plcSecond.Caption = plcSecond.Caption & "."
                                imSecondType = 0
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    End If
                Case 5  '=
                    If imOperator = -1 Then
                        Beep
                        Exit Sub
                    End If
                    smSecondTerm = Mid$(plcSecond.Caption, 3)
                    mDoMath smImmediate, plcImmediate, imImmediateType
                    imOperator = -1
                    smFirstTerm = ""
                    imFirstType = -1
                    smSecondTerm = ""
                    imSecondType = -1
                    imClearInput = True
            End Select
        Case 6  'Row Number
            Select Case ilColNo
                Case 1
                    If imOperator = -1 Then
                        plcFirst.Caption = ""
                    Else
                        Select Case imOperator
                            Case 0
                                plcSecond.Caption = "/ "
                            Case 1
                                plcSecond.Caption = "* "
                            Case 2
                                plcSecond.Caption = "- "
                            Case 3
                                plcSecond.Caption = "+ "
                        End Select
                    End If
                Case 2
                    If imOperator = -1 Then 'On first term
                        If (imFirstType = -1) Or (imFirstType = 1) Then
                            If InStr(plcFirst.Caption, "M") = 0 Then
                                plcFirst.Caption = plcFirst.Caption & ":"
                                imFirstType = 1 'Time
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    Else    'Second term
                        If (imSecondType = -1) Or (imSecondType = 1) Then
                            If InStr(plcSecond.Caption, "M") = 0 Then
                                plcSecond.Caption = plcSecond.Caption & ":"
                                imSecondType = 1 'Time
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    End If
                Case 3
                    If imOperator = -1 Then 'On first term
                        If (imFirstType = -1) Or (imFirstType = 1) Then
                            If InStr(plcFirst.Caption, "M") = 0 Then
                                plcFirst.Caption = plcFirst.Caption & "AM"
                                imFirstType = 1 'Time
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    Else    'Second term
                        If (imSecondType = -1) Or (imSecondType = 1) Then
                            If InStr(plcSecond.Caption, "M") = 0 Then
                                plcSecond.Caption = plcSecond.Caption & "AM"
                                imSecondType = 1 'Time
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    End If
                Case 4
                    If imOperator = -1 Then 'On first term
                        If (imFirstType = -1) Or (imFirstType = 1) Then
                            If InStr(plcFirst.Caption, "M") = 0 Then
                                plcFirst.Caption = plcFirst.Caption & "PM"
                                imFirstType = 1 'Time
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    Else    'Second term
                        If (imSecondType = -1) Or (imSecondType = 1) Then
                            If InStr(plcSecond.Caption, "M") = 0 Then
                                plcSecond.Caption = plcSecond.Caption & "PM"
                                imSecondType = 1 'Time
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    End If
                Case 5
            End Select
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
    sgDoneMsg = "Done" & "|" & Trim$(Str$(CInt(fgCalcLeft))) & "|" & Trim$(Str$(CInt(fgCalcTop)))
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload MathCalc
    igManUnload = NO
End Sub
Private Sub pbcCalculator_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
    If (fmX = -1) And (fmY = -1) Then
        fmY = fmMinY
        fmX = fmMinX
        imcOutline.Move fmX - 15, fmY - 15
    End If
    imcOutline.Visible = True
End Sub
Private Sub pbcCalculator_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    i = 0   '46=delete
End Sub
Private Sub pbcCalculator_KeyPress(KeyAscii As Integer)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Select Case UCase(Chr(KeyAscii))
        Case "H"
            ilRowNo = 1
            ilColNo = 2
            mProcessInput ilRowNo, ilColNo
        Case "M"
            ilRowNo = 1
            ilColNo = 3
            mProcessInput ilRowNo, ilColNo
        Case "S"
            ilRowNo = 1
            ilColNo = 4
            mProcessInput ilRowNo, ilColNo
        Case "7"
            ilRowNo = 2
            ilColNo = 2
            mProcessInput ilRowNo, ilColNo
        Case "8"
            ilRowNo = 2
            ilColNo = 3
            mProcessInput ilRowNo, ilColNo
        Case "9"
            ilRowNo = 2
            ilColNo = 4
            mProcessInput ilRowNo, ilColNo
        Case "4"
            ilRowNo = 3
            ilColNo = 2
            mProcessInput ilRowNo, ilColNo
        Case "5"
            ilRowNo = 3
            ilColNo = 3
            mProcessInput ilRowNo, ilColNo
        Case "6"
            ilRowNo = 3
            ilColNo = 4
            mProcessInput ilRowNo, ilColNo
        Case "1"
            ilRowNo = 4
            ilColNo = 2
            mProcessInput ilRowNo, ilColNo
        Case "2"
            ilRowNo = 4
            ilColNo = 3
            mProcessInput ilRowNo, ilColNo
        Case "3"
            ilRowNo = 4
            ilColNo = 4
            mProcessInput ilRowNo, ilColNo
        Case "0"
            ilRowNo = 5
            ilColNo = 2
            mProcessInput ilRowNo, ilColNo
        Case "00"   'Not possible
            ilRowNo = 5
            ilColNo = 3
            mProcessInput ilRowNo, ilColNo
        Case "."
            ilRowNo = 5
            ilColNo = 4
            mProcessInput ilRowNo, ilColNo
        Case ":"
            ilRowNo = 6
            ilColNo = 2
            mProcessInput ilRowNo, ilColNo
        Case "A"
            ilRowNo = 6
            ilColNo = 3
            mProcessInput ilRowNo, ilColNo
        Case "P"
            ilRowNo = 6
            ilColNo = 4
            mProcessInput ilRowNo, ilColNo
        Case "/"
            ilRowNo = 1
            ilColNo = 5
            mProcessInput ilRowNo, ilColNo
        Case "*"
            ilRowNo = 2
            ilColNo = 5
            mProcessInput ilRowNo, ilColNo
        Case "-"
            ilRowNo = 3
            ilColNo = 5
            mProcessInput ilRowNo, ilColNo
        Case "+"
            ilRowNo = 4
            ilColNo = 5
            mProcessInput ilRowNo, ilColNo
        Case "="
            ilRowNo = 5
            ilColNo = 5
            mProcessInput ilRowNo, ilColNo
        Case " "
            flY = fmMinY
            For ilRowNo = 1 To imNoRows Step 1
                If (fmY + 5 >= flY) And (fmY + 5 <= flY + fmDeltaY) Then
                    flX = fmMinX
                    For ilColNo = 1 To imNoCols Step 1
                        If (fmX + 5 >= flX) And (fmX + 5 <= flX + fmDeltaX) Then
                            mProcessInput ilRowNo, ilColNo
                            Exit Sub
                        End If
                        flX = flX + fmDeltaX
                    Next ilColNo
                End If
                flY = flY + fmDeltaY
            Next ilRowNo
    End Select
    Select Case KeyAscii
        Case 8 'Backspace
            ilRowNo = 3
            ilColNo = 1
            mProcessInput ilRowNo, ilColNo
    End Select
End Sub
Private Sub pbcCalculator_KeyUp(KeyCode As Integer, Shift As Integer)

    If ((Shift And vbCtrlMask) > 0) And (KeyCode = KEYINSERT) Then    'Move to clipboard
        Clipboard.SetText plcImmediate.Caption, vbCFText
    End If
    If KeyCode = KEYUP Then
        fmY = fmY - fmDeltaY
        If fmY < fmMinY Then
            fmY = fmMaxY
        End If
        imcOutline.Move fmX - 15, fmY - 15
    ElseIf KeyCode = KeyDown Then
        fmY = fmY + fmDeltaY
        If fmY > fmMaxY Then
            fmY = fmMinY
        End If
        imcOutline.Move fmX - 15, fmY - 15
    ElseIf KeyCode = KEYLEFT Then
        fmX = fmX - fmDeltaX
        If fmX < fmMinX Then
            fmX = fmMaxX
        End If
        imcOutline.Move fmX - 15, fmY - 15
    ElseIf KeyCode = KEYRIGHT Then
        fmX = fmX + fmDeltaX
        If fmX > fmMaxX Then
            fmX = fmMinX
        End If
        imcOutline.Move fmX - 15, fmY - 15
    End If
End Sub
Private Sub pbcCalculator_LostFocus()
    imcOutline.Visible = False
End Sub
Private Sub pbcCalculator_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    imcInv.Visible = False
    flY = fmMinY
    For ilRowNo = 1 To imNoRows Step 1
        If (Y >= flY) And (Y <= flY + fmDeltaY) Then
            flX = fmMinX
            For ilColNo = 1 To imNoCols Step 1
                If (X >= flX) And (X <= flX + fmDeltaX) Then
                    imcInv.Move flX, flY
                    imcInv.Visible = True
                    imcOutline.Move flX - 15, flY - 15
                    fmX = flX
                    fmY = flY
                    Exit Sub
                End If
                flX = flX + fmDeltaX
            Next ilColNo
        End If
        flY = flY + fmDeltaY
    Next ilRowNo
End Sub
Private Sub pbcCalculator_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    imcInv.Visible = False
    flY = fmMinY
    For ilRowNo = 1 To imNoRows Step 1
        If (Y >= flY) And (Y <= flY + fmDeltaY) Then
            flX = fmMinX
            For ilColNo = 1 To imNoCols Step 1
                If (X >= flX) And (X <= flX + fmDeltaX) Then
                    mProcessInput ilRowNo, ilColNo
                    Exit Sub
                End If
                flX = flX + fmDeltaX
            Next ilColNo
        End If
        flY = flY + fmDeltaY
    Next ilRowNo
End Sub
Private Sub pbcClickFocus_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
    imcOutline.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcCalculator_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcFirst_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcImmediate_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcResult_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSecond_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcShow_Click()
    pbcClickFocus.SetFocus
End Sub
