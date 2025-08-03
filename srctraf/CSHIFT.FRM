VERSION 5.00
Begin VB.Form CShift 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4860
   ClientLeft      =   240
   ClientTop       =   1770
   ClientWidth     =   9690
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
   ScaleHeight     =   4860
   ScaleWidth      =   9690
   Begin VB.OptionButton rbcDates 
      Caption         =   "Remove Weeks"
      Height          =   240
      Index           =   1
      Left            =   3840
      TabIndex        =   12
      Top             =   30
      Width           =   1740
   End
   Begin VB.OptionButton rbcDates 
      Caption         =   "Shift Dates"
      Height          =   240
      Index           =   0
      Left            =   2505
      TabIndex        =   11
      Top             =   30
      Width           =   1260
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
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4650
      Width           =   120
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   195
      ScaleHeight     =   240
      ScaleWidth      =   2175
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   2175
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   3555
      TabIndex        =   5
      Top             =   4500
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4965
      TabIndex        =   0
      Top             =   4500
      Width           =   1050
   End
   Begin VB.PictureBox plcInv 
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
      Height          =   4020
      Left            =   90
      ScaleHeight     =   3960
      ScaleWidth      =   9345
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   375
      Visible         =   0   'False
      Width           =   9405
      Begin VB.ListBox lbcDates 
         Height          =   2790
         ItemData        =   "CSHIFT.frx":0000
         Left            =   8040
         List            =   "CSHIFT.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   13
         Top             =   750
         Width           =   1155
      End
      Begin VB.PictureBox pbcLbcLines 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2715
         Left            =   150
         ScaleHeight     =   2715
         ScaleWidth      =   9030
         TabIndex        =   10
         Top             =   765
         Width           =   9030
      End
      Begin VB.HScrollBar hbcShift 
         Height          =   240
         Index           =   0
         Left            =   8310
         Max             =   200
         TabIndex        =   7
         Top             =   240
         Value           =   50
         Width           =   540
      End
      Begin VB.CheckBox ckcAll 
         Caption         =   "All Lines"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   135
         TabIndex        =   4
         Top             =   3570
         Width           =   1110
      End
      Begin VB.ListBox lbcLines 
         Appearance      =   0  'Flat
         Height          =   2760
         ItemData        =   "CSHIFT.frx":0004
         Left            =   135
         List            =   "CSHIFT.frx":0006
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   750
         Width           =   9045
      End
      Begin VB.Label lacTitle 
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
         Height          =   225
         Left            =   135
         TabIndex        =   9
         Top             =   525
         Width           =   9030
      End
      Begin VB.Label lacShift 
         Appearance      =   0  'Flat
         Caption         =   "Week"
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
         Index           =   0
         Left            =   8340
         TabIndex        =   8
         Top             =   45
         Width           =   570
      End
   End
End
Attribute VB_Name = "CShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CSHIFT.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CShift.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract terminate screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcLines)
Dim lmCntrStartDate As Long
Dim lmCntrEndDate As Long
Dim imWkValue As Integer
'Dim imListField(1 To 7) As Integer
Dim imListField(0 To 7) As Integer
Dim imLBCtrls As Integer

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
        If ilValue Then
            'Remove test as only shift for proposals
            'For ilLoop = lbcLines.ListCount - 1 To 0 Step -1
            '    If mOKToSelect(ilLoop) Then
            '        lbcLines.Selected(ilLoop) = ilValue
            '    End If
            'Next ilLoop
            If lbcLines.ListCount > 0 Then
                llRg = CLng(lbcLines.ListCount - 1) * &H10000 + 0
                llRet = SendMessageByNum(lbcLines.hWnd, LB_SELITEMRANGE, True, llRg)
            End If
        Else
'            For ilLoop = lbcLines.ListCount - 1 To 0 Step -1
'                lbcLines.Selected(ilLoop) = ilValue
'            Next ilLoop
            If lbcLines.ListCount > 0 Then
                llRg = CLng(lbcLines.ListCount - 1) * &H10000 + 0
                llRet = SendMessageByNum(lbcLines.hWnd, LB_SELITEMRANGE, False, llRg)
            End If
        End If
        pbcLbcLines_Paint
        imAllClicked = False
    End If
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub
Private Sub cmcCancel_Click()
    igTerminateReturn = 0   'Cancel selected
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    Dim ilLoop As Integer
    Dim ilCff As Integer
    Dim slStartDate As String
    Dim llDate As Long
    Dim slDate As String
    Dim slLine As String
    Dim ilNoDaysChg As Integer
    Dim ilRet As Integer
    Dim ilLineNo As Integer
    Dim slLineNo As String
    Dim ilClf As Integer
    Dim ilHClf As Integer
    Dim ilPkg As Integer
    Dim ilIndex As Integer
    Dim slLineType As String
    Dim slWDFlag As String
    Dim ilWk As Integer
    Dim ilDay As Integer
    Dim llSpots As Long
    Dim blWeekFound As Boolean
    Dim ilDays(0 To 6) As Integer
    
    'Shift Dates
    If lbcLines.SelCount <= 0 Then
        If (rbcDates(0).Enabled = False) Or (rbcDates(1).Enabled = False) Then
            ilRet = MsgBox("No Line Selected, Continue with Done and Exit Screen?", vbQuestion + vbYesNo, "Done")
            If ilRet = vbNo Then
                Exit Sub
            End If
        End If
        igTerminateReturn = 0
        mTerminate
        Exit Sub
    End If
    If rbcDates(0).Value Then   'Shift dates
        For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
            If Not tgClfCntr(lbcLines.ItemData(ilLoop)).iCancel Then
                tgClfCntr(lbcLines.ItemData(ilLoop)).iResearch = False
                ilNoDaysChg = 0
                slLine = lbcLines.List(ilLoop)
                'Get Dates and test if different
                ilRet = gParseItem(slLine, 5, "|", slDate)
                'slDate = Mid$(slLine, imListFieldChar(4), imListFieldChar(5) - imListFieldChar(4))
                If ilRet = CP_MSG_NONE Then
                    ilRet = gParseItem(slDate, 1, "-", slStartDate)
                    If ilRet = CP_MSG_NONE Then
                        ilRet = gParseItem(slLine, 6, "|", slDate)
                        'slDate = Mid$(slLine, imListFieldChar(5), imListFieldChar(6) - imListFieldChar(5))
                        ilNoDaysChg = gDateValue(slDate) - gDateValue(slStartDate)
                    End If
                End If
                If ilNoDaysChg <> 0 Then
                    ilRet = gParseItem(slLine, 1, "|", slLineNo)
                    ilLineNo = Val(slLineNo)
                    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                        If tgClfCntr(ilClf).ClfRec.iLine = ilLineNo Then
                            ilCff = tgClfCntr(ilClf).iFirstCff
                            Do While ilCff <> -1
                                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                    gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llDate
                                    slDate = Format$(llDate + ilNoDaysChg, "m/d/yy")
                                    gPackDate slDate, tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1)
                                    tgCffCntr(ilCff).lStartDate = llDate + ilNoDaysChg
                                    gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llDate
                                    slDate = Format$(llDate + ilNoDaysChg, "m/d/yy")
                                    gPackDate slDate, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
                                    tgCffCntr(ilCff).lEndDate = llDate + ilNoDaysChg
                                End If
                                ilCff = tgCffCntr(ilCff).iNextCff
                            Loop
                            slLineType = mGetLineType(ilClf + 1)
                            If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                                'For ilHClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                '    slLineType = mGetLineType(ilHClf + 1)
                                '    If (slLineType = "H") And (tgClfCntr(ilHClf).ClfRec.iPkLineNo = tgClfCntr(ilClf).ClfRec.iLine) Then
                                '        ilCff = tgClfCntr(ilHClf).iFirstCff
                                '        Do While ilCff <> -1
                                '            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                '                gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llDate
                                '                slDate = Format$(llDate + ilNoDaysChg, "m/d/yy")
                                '                gPackDate slDate, tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1)
                                '                tgCffCntr(ilCff).lStartDate = llDate + ilNoDaysChg
                                '                gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llDate
                                '                slDate = Format$(llDate + ilNoDaysChg, "m/d/yy")
                                '                gPackDate slDate, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
                                '                tgCffCntr(ilCff).lEndDate = llDate + ilNoDaysChg
                                '            End If
                                '            ilCff = tgCffCntr(ilCff).iNextCff
                                '        Loop
                                '    End If
                                'Next ilHClf
                                For ilPkg = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                                    If tgPackLn(ilPkg).iLineNo = ilLineNo Then
                                        ilIndex = tgPackLn(ilPkg).iIndex
                                        Do While ilIndex >= 0
                                            For ilHClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                                If tgPackLn(ilIndex).iLineNo = tgClfCntr(ilHClf).ClfRec.iLine Then
                                                    If Not tgClfCntr(ilClf).iCancel Then
                                                        ilCff = tgClfCntr(ilHClf).iFirstCff
                                                        Do While ilCff <> -1
                                                            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                                                gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llDate
                                                                slDate = Format$(llDate + ilNoDaysChg, "m/d/yy")
                                                                gPackDate slDate, tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1)
                                                                tgCffCntr(ilCff).lStartDate = llDate + ilNoDaysChg
                                                                gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llDate
                                                                slDate = Format$(llDate + ilNoDaysChg, "m/d/yy")
                                                                gPackDate slDate, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
                                                                tgCffCntr(ilCff).lEndDate = llDate + ilNoDaysChg
                                                            End If
                                                            ilCff = tgCffCntr(ilCff).iNextCff
                                                        Loop
                                                    End If
                                                    Exit For
                                                End If
                                            Next ilHClf
                                            ilIndex = tgPackLn(ilIndex).iIndex
                                        Loop
                                    End If
                                Next ilPkg
                                
                            End If
                            Exit For
                        End If
                    Next ilClf
                End If
            End If
        Next ilLoop
    ElseIf rbcDates(1).Value Then 'remove weeks
        For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
            If (Not tgClfCntr(lbcLines.ItemData(ilLoop)).iCancel) And (lbcLines.Selected(ilLoop)) Then
                tgClfCntr(lbcLines.ItemData(ilLoop)).iResearch = False
                ilClf = lbcLines.ItemData(ilLoop)
                slLineType = mGetLineType(ilClf + 1)
                If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                    ilLineNo = tgClfCntr(ilClf).ClfRec.iLine
                    'Remove weeks from hidden lines
                    'For ilHClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                    '    slLineType = mGetLineType(ilHClf + 1)
                    '    If (slLineType = "H") And (tgClfCntr(ilHClf).ClfRec.iPkLineNo = tgClfCntr(ilClf).ClfRec.iLine) Then
                    For ilPkg = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                        If tgPackLn(ilPkg).iLineNo = ilLineNo Then
                            ilIndex = tgPackLn(ilPkg).iIndex
                            Do While ilIndex >= 0
                                For ilHClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                    If tgPackLn(ilIndex).iLineNo = tgClfCntr(ilHClf).ClfRec.iLine Then
                                        For ilWk = 0 To lbcDates.ListCount - 1 Step 1
                                            If lbcDates.Selected(ilWk) Then
                                                tgClfCntr(ilHClf).iResearch = False
                                                llDate = gDateValue(lbcDates.List(ilWk))
                                                llSpots = 0
                                                ilCff = tgClfCntr(ilHClf).iFirstCff
                                                blWeekFound = False
                                                If ilCff <> -1 Then
                                                    Do While ilCff <> -1
                                                        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                                            If (llDate >= tgCffCntr(ilCff).lStartDate) And (llDate <= tgCffCntr(ilCff).lEndDate) Then
                                                                slWDFlag = tgCffCntr(ilCff).CffRec.sDyWk
                                                                If slWDFlag = "D" Then
                                                                    For ilDay = 0 To 6 Step 1
                                                                        ilDays(ilDay) = 0
                                                                    Next ilDay
                                                                Else
                                                                    For ilDay = 0 To 6 Step 1
                                                                        If (tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilDay) = "1") Then
                                                                            ilDays(ilDay) = True
                                                                        Else
                                                                            ilDays(ilDay) = False
                                                                        End If
                                                                    Next ilDay
                                                                End If
                                                                blWeekFound = True
                                                                Exit Do
                                                            End If
                                                        End If
                                                        ilCff = tgCffCntr(ilCff).iNextCff
                                                    Loop
                                                    If blWeekFound Then
                                                        mAddWeek imLastCffUsed, ilHClf + 1, llDate, slWDFlag, llSpots, ilDays()
                                                        tgClfCntr(ilHClf).sLineChgd = "Y"
                                                        ilCff = tgClfCntr(ilHClf).iFirstCff
                                                        Do While ilCff <> -1
                                                            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                                                If (llDate >= tgCffCntr(ilCff).lStartDate) And (llDate <= tgCffCntr(ilCff).lEndDate) Then
                                                                    tgCffCntr(ilCff).CffRec.sPriceType = ""
                                                                    tgCffCntr(ilCff).CffRec.lActPrice = 0
                                                                    Exit Do
                                                                End If
                                                            End If
                                                            ilCff = tgCffCntr(ilCff).iNextCff
                                                        Loop
                                                    End If
                                                End If
                                            End If
                                        Next ilWk
                                        Exit For
                                    End If
                                Next ilHClf
                                ilIndex = tgPackLn(ilIndex).iIndex
                            Loop
                        End If
                    Next ilPkg
                    '    End If
                    'Next ilHClf
                End If
                'Remove week from line
                For ilWk = 0 To lbcDates.ListCount - 1 Step 1
                    If lbcDates.Selected(ilWk) Then
                        llDate = gDateValue(lbcDates.List(ilWk))
                        llSpots = 0
                        blWeekFound = False
                        ilCff = tgClfCntr(ilClf).iFirstCff
                        If ilCff <> -1 Then
                            Do While ilCff <> -1
                                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                    If (llDate >= tgCffCntr(ilCff).lStartDate) And (llDate <= tgCffCntr(ilCff).lEndDate) Then
                                        slWDFlag = tgCffCntr(ilCff).CffRec.sDyWk
                                        If slWDFlag = "D" Then
                                            For ilDay = 0 To 6 Step 1
                                                ilDays(ilDay) = 0
                                            Next ilDay
                                        Else
                                            For ilDay = 0 To 6 Step 1
                                                If (tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilDay) = "1") Then
                                                    ilDays(ilDay) = True
                                                Else
                                                    ilDays(ilDay) = False
                                                End If
                                            Next ilDay
                                        End If
                                        blWeekFound = True
                                        Exit Do
                                    End If
                                End If
                                ilCff = tgCffCntr(ilCff).iNextCff
                            Loop
                            If blWeekFound Then
                                mAddWeek imLastCffUsed, ilClf + 1, llDate, slWDFlag, llSpots, ilDays()
                                tgClfCntr(ilClf).sLineChgd = "Y"
                                ilCff = tgClfCntr(ilClf).iFirstCff
                                Do While ilCff <> -1
                                    If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                        If (llDate >= tgCffCntr(ilCff).lStartDate) And (llDate <= tgCffCntr(ilCff).lEndDate) Then
                                            tgCffCntr(ilCff).CffRec.sPriceType = ""
                                            tgCffCntr(ilCff).CffRec.lActPrice = 0
                                            Exit Do
                                        End If
                                    End If
                                    ilCff = tgCffCntr(ilCff).iNextCff
                                Loop
                            End If
                        End If
                    End If
                Next ilWk
            End If
        Next ilLoop
    End If
    If rbcDates(0).Value Then
        igTerminateReturn = 1   'Done selected
    Else
        igTerminateReturn = 2   'Done selected
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
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

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub hbcShift_Change(Index As Integer)
    Dim ilLoop As Integer
    Dim slLine As String
    Dim slDate As String
    Dim slField1 As String
    Dim slField2 As String
    Dim slField3 As String
    Dim slField4 As String
    Dim slField5 As String
    Dim ilRet As Integer
    rbcDates(1).Enabled = False
    If Index = 0 Then
        If imWkValue > hbcShift(0).Value Then
            'Reduce by one week
            For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
                If lbcLines.Selected(ilLoop) And (Not tgClfCntr(lbcLines.ItemData(ilLoop)).iCancel) Then
                    slLine = lbcLines.List(ilLoop)
                    ilRet = gParseItem(slLine, 6, "|", slDate)
                    'slDate = Trim$(Mid$(slLine, imListFieldChar(5), imListFieldChar(6) - imListFieldChar(5)))
                    slDate = gDecOneWeek(slDate)
                    ilRet = gParseItem(slLine, 1, "|", slField1)
                    'slField1 = Left$(slLine, imListFieldChar(1) - 1)
                    ilRet = gParseItem(slLine, 2, "|", slField2)
                    'slField2 = Trim$(Mid$(slLine, imListFieldChar(1), imListFieldChar(2) - imListFieldChar(1)))
                    ilRet = gParseItem(slLine, 3, "|", slField3)
                    'slField3 = Trim$(Mid$(slLine, imListFieldChar(2), imListFieldChar(3) - imListFieldChar(2)))
                    ilRet = gParseItem(slLine, 4, "|", slField4)
                    'slField4 = Trim$(Mid$(slLine, imListFieldChar(3), imListFieldChar(4) - imListFieldChar(3)))
                    ilRet = gParseItem(slLine, 5, "|", slField5)
                    'slField5 = Trim$(Mid$(slLine, imListFieldChar(4), imListFieldChar(5) - imListFieldChar(4)))
                    'lbcLines.List(ilLoop) = gAlignStringByPixel(slField1 & "|" & slField2 & "|" & slField3 & "|" & slField4 & "|" & slField5 & "|" & slDate, "|", imListField(), imListFieldChar())
                    lbcLines.List(ilLoop) = slField1 & "|" & slField2 & "|" & slField3 & "|" & slField4 & "|" & slField5 & "|" & slDate
                End If
            Next ilLoop
        ElseIf imWkValue < hbcShift(0).Value Then
            'Increase by one week
            For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
                If lbcLines.Selected(ilLoop) And (Not tgClfCntr(lbcLines.ItemData(ilLoop)).iCancel) Then
                    slLine = lbcLines.List(ilLoop)
                    ilRet = gParseItem(slLine, 6, "|", slDate)
                    'slDate = Trim$(Mid$(slLine, imListFieldChar(5), imListFieldChar(6) - imListFieldChar(5)))
                    slDate = gIncOneWeek(slDate)
                    ilRet = gParseItem(slLine, 1, "|", slField1)
                    ilRet = gParseItem(slLine, 2, "|", slField2)
                    ilRet = gParseItem(slLine, 3, "|", slField3)
                    ilRet = gParseItem(slLine, 4, "|", slField4)
                    ilRet = gParseItem(slLine, 5, "|", slField5)
                    lbcLines.List(ilLoop) = slField1 & "|" & slField2 & "|" & slField3 & "|" & slField4 & "|" & slField5 & "|" & slDate
                    'slField1 = Left$(slLine, imListFieldChar(1) - 1)
                    'slField2 = Trim$(Mid$(slLine, imListFieldChar(1), imListFieldChar(2) - imListFieldChar(1)))
                    'slField3 = Trim$(Mid$(slLine, imListFieldChar(2), imListFieldChar(3) - imListFieldChar(2)))
                    'slField4 = Trim$(Mid$(slLine, imListFieldChar(3), imListFieldChar(4) - imListFieldChar(3)))
                    'slField5 = Trim$(Mid$(slLine, imListFieldChar(4), imListFieldChar(5) - imListFieldChar(4)))
                    'lbcLines.List(ilLoop) = gAlignStringByPixel(slField1 & "|" & slField2 & "|" & slField3 & "|" & slField4 & "|" & slField5 & "|" & slDate, "|", imListField(), imListFieldChar())
                End If
            Next ilLoop
        End If
        imWkValue = hbcShift(0).Value
    Else
    End If
    pbcLbcLines_Paint
End Sub
Private Sub lacTitle_Click()
    Dim ilLoop As Integer
    ilLoop = 0
End Sub

Private Sub lbcDates_Click()
    rbcDates(0).Enabled = False
End Sub

Private Sub lbcLines_Click()
    Dim slLineType As String
    Dim ilLoop As Integer
    Dim ilClf As Integer
    Dim ilLp As Integer
    Dim ilIndex As Integer
    Dim ilLnIndex As Integer
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
        'For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
        '    If lbcLines.Selected(ilLoop) Then
        '        If Not mOKToSelect(ilLoop) Then
        '            Beep
        '            lbcLines.Selected(ilLoop) = False
        '        End If
        '    End If
        'Next ilLoop
        'For now (2/2/99) disallow hidden lines to be selected--
        'Reason-  if hidden shifted and Not package, then need to
        '         add weeks to package when done pressed- extra code that
        '         might never be used- we will wait until asked
        '         (Dick and Darlene)
        'If package, then select all lines within package
        For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
            If lbcLines.Selected(ilLoop) Then
                ilLnIndex = lbcLines.ItemData(ilLoop)
                slLineType = mGetLineType(ilLnIndex + 1)
                If (slLineType = "H") Then
                    Beep
                    imAllClicked = True
                    lbcLines.Selected(ilLoop) = False
                    imAllClicked = False
                End If
            End If
        Next ilLoop
        For ilClf = 0 To lbcLines.ListCount - 1 Step 1
            If lbcLines.Selected(ilClf) Then
                slLineType = mGetLineType(lbcLines.ItemData(ilClf) + 1)
                If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                    'For ilClf = 0 To lbcLines.ListCount - 1 Step 1
                        'If tgClfCntr(ilClf).ClfRec.iPkLineNo = tgClfCntr(ilLoop).ClfRec.iLine Then
                        '    lbcLines.Selected(ilClf) = True
                        'End If
                    'Next ilClf
                    For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                        If tgClfCntr(lbcLines.ItemData(ilClf)).ClfRec.iLine = tgPackLn(ilLoop).iLineNo Then
                            ilIndex = tgPackLn(ilLoop).iIndex
                            Do While ilIndex >= 0
                                For ilLp = 0 To lbcLines.ListCount - 1 Step 1
                                    ilLnIndex = lbcLines.ItemData(ilLp)
                                    If tgPackLn(ilIndex).iLineNo = tgClfCntr(ilLnIndex).ClfRec.iLine Then
                                        If Not lbcLines.Selected(ilLp) Then
                                            imAllClicked = True
                                            lbcLines.Selected(ilLp) = True
                                            imAllClicked = False
                                        End If
                                        Exit For
                                    End If
                                Next ilLp
                                ilIndex = tgPackLn(ilIndex).iIndex
                            Loop
                            Exit For
                        End If
                    Next ilLoop
                End If
            End If
        Next ilClf
        pbcLbcLines_Paint
    End If
End Sub
Private Sub lbcLines_GotFocus()
    gCtrlGotFocus lbcLines
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
    Dim ilTest As Integer
    Dim slStr As String
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilCff As Integer
    Dim slLineType As String
    Dim ilVef As Integer
    Dim ilSport As Integer
    Dim ilPkg As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llDate As Long
    Dim blAddLine As Boolean
    
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    imLBCtrls = 1
    imListField(1) = 15
    imListField(2) = 5 * igAlignCharWidth
    imListField(3) = 35 * igAlignCharWidth
    imListField(4) = 62 * igAlignCharWidth
    imListField(5) = 65 * igAlignCharWidth
    imListField(6) = 90 * igAlignCharWidth
    imListField(7) = 134 * igAlignCharWidth
    'If igShiftDates Then
    '    lacShift(0).Visible = True
    '    hbcShift(0).Visible = True
    '    ckcAll.Visible = True
    '    lacTitle.Caption = "Line  Vehicle                                  Daypart                                   Current Dates                 New Start"
    'Else
    '    lacShift(0).Visible = False
    '    hbcShift(0).Visible = False
    '    ckcAll.Visible = False
    '    lacTitle.Caption = "Line  Vehicle                                  Daypart                                   Current Dates"
    'End If
    imSetAll = True
    imAllClicked = False
    gUnpackDate tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), slStartDate
    lmCntrStartDate = gDateValue(slStartDate)
    gUnpackDate tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), slEndDate
    lmCntrEndDate = gDateValue(slEndDate)
    For llDate = gDateValue(slStartDate) To gDateValue(slEndDate) Step 7
        If llDate >= lmFirstSchMonDate Then
            lbcDates.AddItem Format(llDate, "m/d/yy")
        End If
    Next llDate
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        slLineType = mGetLineType(ilClf + 1)
        ilSport = False
        If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
            ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
            If ilVef <> -1 Then
                If tgMVef(ilVef).sType = "G" Then
                    ilSport = True
                End If
            End If
            If slLineType = "H" Then
                For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                    If (tgPackLn(ilLoop).sType = "O") Or (tgPackLn(ilLoop).sType = "A") Or (tgPackLn(ilLoop).sType = "E") Then
                        ilIndex = tgPackLn(ilLoop).iIndex
                        Do While ilIndex >= 0
                            If tgPackLn(ilIndex).iLineNo = tgClfCntr(ilClf).ClfRec.iLine Then
                                For ilPkg = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                    If tgPackLn(ilLoop).iLineNo = tgClfCntr(ilPkg).ClfRec.iLine Then
                                        If mGameInPkg(ilPkg + 1) Then
                                            ilSport = True
                                        End If
                                        Exit For
                                    End If
                                Next ilPkg
                                If ilSport = True Then
                                    Exit Do
                                End If
                            End If
                            ilIndex = tgPackLn(ilIndex).iIndex
                        Loop
                    End If
                Next ilLoop
            End If
        Else
            If mGameInPkg(ilClf + 1) Then
                ilSport = True
            End If
        End If
        If Not ilSport Then
            slStr = Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine))
            If tgClfCntr(ilClf).iCancel Then
                slStr = slStr & "| |" & "Canceled before Start"
            Else
                slRecCode = Trim$(str$(tgClfCntr(ilClf).ClfRec.iVefCode))
                slLineType = mGetLineType(ilClf + 1)
                If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                    If slLineType = "H" Then
                        blAddLine = False
                    Else
                        blAddLine = True
                    End If
                    For ilTest = 0 To UBound(tmVehicleCode) - 1 Step 1  'Contract!lbcVehicle.ListCount - 1 Step 1
                        slNameCode = tmVehicleCode(ilTest).sKey    'Contract!lbcVehicle.List(ilTest)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        On Error GoTo mInitErr
                        gCPErrorMsg ilRet, "mInit (gParseItem field 2)", CShift
                        On Error GoTo 0
                        If slRecCode = slCode Then
                            ilRet = gParseItem(slNameCode, 1, "\", slName)
                            On Error GoTo mInitErr
                            gCPErrorMsg ilRet, "mInit (gParseItem field 2)", CShift
                            On Error GoTo 0
                            ilRet = gParseItem(slName, 3, "|", slName)
                            On Error GoTo mInitErr
                            gCPErrorMsg ilRet, "mInit (gParseItem field 2)", CShift
                            On Error GoTo 0
                            slStr = slStr & "|" & slName
                            Exit For
                        End If
                    Next ilTest
                Else
                    blAddLine = True
                    If (tgClfCntr(ilClf).ClfRec.iVefCode = 0) Then
                        slStr = slStr & "|" & Trim$(smLnSave(1, ilClf + 1))
                    Else
                        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If tgClfCntr(ilClf).ClfRec.iVefCode = tgMVef(ilVef).iCode Then
                            ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
                            If ilVef <> -1 Then
                                slStr = slStr & "|" & Trim$(tgMVef(ilVef).sName)
                        '        Exit For
                            End If
                        'Next ilVef
                    End If
                End If
                slStr = slStr & "|" & mMakeDPName(ilClf + 1)  '
                slStr = slStr & "|" & Trim$(str$(tgClfCntr(ilClf).ClfRec.iLen))
                If tgClfCntr(ilClf).iFirstCff <> -1 Then
                    slStartDate = ""
                    slEndDate = ""
                    ilCff = tgClfCntr(ilClf).iFirstCff
                    Do While ilCff <> -1
                        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                            If slStartDate = "" Then
                                gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate
                            End If
                            gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate
                        End If
                        ilCff = tgCffCntr(ilCff).iNextCff
                    Loop
                    slStr = slStr & "|" & slStartDate & "-" & slEndDate
                    slStr = slStr & "|" & slStartDate
                End If
            End If
            'lbcLines.AddItem gAlignStringByPixel(slStr, "|", imListField(), imListFieldChar())
            If blAddLine Then
                lbcLines.AddItem slStr
                lbcLines.ItemData(lbcLines.NewIndex) = ilClf
            End If
        End If
    Next ilClf
    imWkValue = hbcShift(0).Value
    'imMnValue = hbcShift(1).Value
    CShift.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterModalForm CShift
    pbcLbcLines_Paint
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Unload CShift
    Set CShift = Nothing   'Remove data segment
End Sub

Private Sub lbcLines_Scroll()
    pbcLbcLines_Paint
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcLbcLines_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLinesEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 5) As String
    Dim llFgColor As Long
    Dim ilFieldIndex As Integer
    
    ilLinesEnd = lbcLines.TopIndex + lbcLines.Height \ fgListHtArial825
    If ilLinesEnd > lbcLines.ListCount Then
        ilLinesEnd = lbcLines.ListCount
    End If
    'If lbcLines.ListCount <= lbcLines.Height \ fgListHtArial825 Then
    '    llWidth = lbcLines.Width - 30
    'Else
    '    llWidth = lbcLines.Width - igScrollBarWidth - 30
    'End If
    'pbcLbcLines.Width = llWidth
    pbcLbcLines.Cls
    llFgColor = pbcLbcLines.ForeColor
    For ilLoop = lbcLines.TopIndex To ilLinesEnd - 1 Step 1
        pbcLbcLines.ForeColor = llFgColor
        If lbcLines.MultiSelect = 0 Then
            If lbcLines.ListIndex = ilLoop Then
                gPaintArea pbcLbcLines, CSng(0), CSng((ilLoop - lbcLines.TopIndex) * fgListHtArial825), CSng(pbcLbcLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLines.ForeColor = vbWhite
            End If
        Else
            If lbcLines.Selected(ilLoop) Then
                gPaintArea pbcLbcLines, CSng(0), CSng((ilLoop - lbcLines.TopIndex) * fgListHtArial825), CSng(pbcLbcLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLines.ForeColor = vbWhite
            End If
        End If
        slStr = lbcLines.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcLines.CurrentX = imListField(ilField)
            pbcLbcLines.CurrentY = (ilLoop - lbcLines.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcLines, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcLines.Print slStr
        Next ilField
        pbcLbcLines.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub plcInv_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Contract Lines "
End Sub

Private Sub mSetControls()
    Dim ilLinesEnd As Integer
    Dim llWidth As Long
    Dim blScrollShown As Boolean
    
    plcInv.Visible = False
    lbcLines.Width = plcInv.Width - 2 * 150
    lbcDates.Visible = False
    If lbcLines.ListCount <= lbcLines.Height \ fgListHtArial825 Then
        llWidth = lbcLines.Width - 30
        blScrollShown = False
    Else
        llWidth = lbcLines.Width - igScrollBarWidth - 30
        blScrollShown = True
    End If
    If rbcDates(0).Value Then
        If igShiftDates Then
            lacShift(0).Visible = True
            hbcShift(0).Visible = True
            ckcAll.Visible = True
            lacTitle.Caption = "Line  Vehicle                                  Daypart                                   Current Dates                 New Start"
            pbcLbcLines.Width = llWidth
        Else
            lacShift(0).Visible = False
            hbcShift(0).Visible = False
            ckcAll.Visible = False
            lacTitle.Caption = "Line  Vehicle                                  Daypart                                   Current Dates"
            'pbcLbcLines.Width = llWidth
            pbcLbcLines.Width = llWidth - lbcDates.Width - 240
        End If
    Else
        lbcDates.Left = lbcLines.Left + lbcLines.Width - lbcDates.Width
        lacShift(0).Visible = False
        hbcShift(0).Visible = False
        ckcAll.Visible = True
        lacTitle.Caption = "Line  Vehicle                                  Daypart                                   Current Dates                 Remove"
        pbcLbcLines.Width = llWidth - lbcDates.Width - 240
        lbcDates.Visible = True
    End If
    If blScrollShown Then
        lbcLines.Width = pbcLbcLines.Width + igScrollBarWidth + 30
    Else
        lbcLines.Width = pbcLbcLines.Width + 30
    End If
    plcInv.Visible = True
End Sub

Private Sub rbcDates_Click(Index As Integer)
    If rbcDates(Index).Value Then
        mSetControls
    End If
End Sub
