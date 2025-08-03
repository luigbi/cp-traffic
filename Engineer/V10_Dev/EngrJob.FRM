VERSION 5.00
Begin VB.Form EngrJob 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3435
   ClientLeft      =   135
   ClientTop       =   135
   ClientWidth     =   7965
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3435
   ScaleWidth      =   7965
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   6945
      Top             =   1350
   End
   Begin VB.PictureBox plcWindow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3435
      Left            =   0
      ScaleHeight     =   3435
      ScaleWidth      =   7965
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7965
      Begin VB.PictureBox pbcHor 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   3255
         Left            =   90
         Picture         =   "EngrJob.frx":0000
         ScaleHeight     =   3195
         ScaleWidth      =   7740
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   90
         Width           =   7800
         Begin VB.Label lacPressMsg 
            BackStyle       =   0  'Transparent
            Caption         =   "Press F3 or Click Logo for Lists"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   3960
            TabIndex        =   13
            Top             =   2880
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.Label lacJobs 
            BackStyle       =   0  'Transparent
            Caption         =   "Jobs"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7110
            TabIndex        =   12
            Top             =   30
            Width           =   570
         End
         Begin VB.Label imcOutline 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   570
            Left            =   6135
            TabIndex        =   11
            Top             =   405
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Schedules"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   2370
            TabIndex        =   7
            Top             =   2370
            Width           =   2085
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Templates"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   4230
            TabIndex        =   5
            Top             =   1635
            Width           =   1890
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Libraries"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   825
            TabIndex        =   3
            Top             =   1635
            Width           =   1890
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Time Finder"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   2160
            TabIndex        =   1
            Top             =   900
            Width           =   2655
         End
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Templates"
         Height          =   570
         Index           =   2
         Left            =   4350
         TabIndex        =   6
         Tag             =   "Add, Change, Cancel or Display contracts"
         Top             =   1635
         Width           =   1980
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Schedules"
         Height          =   570
         Index           =   3
         Left            =   2430
         TabIndex        =   8
         Tag             =   "Add, change, or delete all Program material for logs"
         Top             =   2355
         Width           =   2205
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Libraries"
         Height          =   570
         Index           =   1
         Left            =   915
         TabIndex        =   4
         Tag             =   "Create or modify an Agency Proposal"
         Top             =   1605
         Width           =   1980
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Rat&e Cards"
         Height          =   570
         Index           =   0
         Left            =   2100
         TabIndex        =   2
         Tag             =   "Create new Rate cards or modify existing ones"
         Top             =   915
         Width           =   2955
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   165
         ScaleHeight     =   165
         ScaleWidth      =   75
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3855
         Width           =   75
      End
      Begin VB.Image cmcStartList 
         Height          =   1365
         Left            =   6300
         Top             =   1965
         Width           =   1560
      End
      Begin VB.Image cmcCSLogo 
         Height          =   675
         Left            =   180
         Top             =   120
         Width           =   4605
      End
   End
End
Attribute VB_Name = "EngrJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EngrJob.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Basic 10 task buttons screen code
Option Explicit
Option Compare Text
                                    
Dim imTaskIndex As Integer
Dim imPrevTaskIndex As Integer


Const RESOURCEINDEX = 0
Const LIBRARYINDEX = 1
Const TEMPLATEINDEX = 2
Const SCHEDULEINDEX = 3


Private Sub mSetCtrlEnable(ilCtrlIndex As Integer, ilTaskNo As Integer)
    If igJobStatus(ilTaskNo) = 0 Then
        cmcTask(ilCtrlIndex).Enabled = False
        lacTask(ilCtrlIndex).Enabled = False
    Else
        cmcTask(ilCtrlIndex).Enabled = True
    End If
End Sub


Public Sub mSetCtrlProperties(ilCtrlIndex As Integer, ilTaskNo As Integer)
    If igJobStatus(ilTaskNo) = 0 Then
        cmcTask(ilCtrlIndex).ToolTipText = "Disallowed"
    Else
        If igJobStatus(ilTaskNo) = 1 Then
            cmcTask(ilCtrlIndex).ToolTipText = "View Only"
            imPrevTaskIndex = ilCtrlIndex
        Else
            cmcTask(ilCtrlIndex).ToolTipText = "Updates Allowed"
        End If
    End If
End Sub

Private Sub cmcCSLogo_Click()
    Dim ilAns As Integer
    Dim llNow As Long
    Dim slStr As String
    Dim slNow As String
    Dim ilLoop As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llDate As Long
    Dim slKey As String
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilValue As Integer
    Dim ilNxtValue As Integer
    Dim slNxtValue As String
    Dim slInValue As String
    Dim slDays As String
    Dim ilPattern As Integer
    Dim ilRet As Integer
    Dim slDate As String
    
'    If igGGFlag = 1 Then
'        ilAns = MsgBox(sgCSVersion, 0, "Counterpoint Software")
'    Else
'        hmSaf = CBtrvTable(TWOHANDLES)
'        ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'        If ilRet = BTRV_ERR_NONE Then
'            imSafRecLen = Len(tmSaf)  'btrRecordLength(hlUrf)  'Get and save record length
'            ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'            If ilRet = BTRV_ERR_NONE Then
'                ilField1 = Asc(tmSaf.sName)
'                slStr = Mid$(tmSaf.sName, 2, 5)
'                If Left$(slStr, 1) <> "9" Then
'                    llDate = Val(slStr)
'                Else
'                    llDate = gDateValue(Format$(gNow(), "ddddd"))
'                End If
'                llNow = gDateValue(Format$(gNow(), "ddddd"))
'                ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
'                sgGenMsg = sgCSVersion
'                If llNow < llDate Then
'                    sgGenMsg = sgGenMsg & "," & Str$(llDate - llNow)
'                End If
'                slStr = Mid$(tmSaf.sName, 7, 4)
'                'If slStr = "0000" Then
'                    'ilValue = Int(10000 * Rnd(-ilValue) + 1)
'                    ilMin = Asc("A")
'                    ilMax = Asc("Z")
'                    slStr = ""
'                    For ilLoop = 0 To 3 Step 1
'                        Randomize
'                        slStr = slStr & Trim$(Chr$(Int((ilMax - ilMin + 1) * Rnd + ilMin)))
'                    Next ilLoop
'                'End If
'                For ilLoop = 0 To 3 Step 1
'                    ilValue = ilValue + (ilLoop + 1) * Asc(Mid$(slStr, ilLoop + 1, 1))
'                Next ilLoop
'                slNxtValue = ""
'                For ilLoop = 0 To 3 Step 1
'                    ilValue = Int(9000 * Rnd(-ilValue) + 1000)
'                    slNxtValue = slNxtValue & Trim$(Chr$(Asc("A") + ilValue Mod 25))
'                Next ilLoop
'                sgGenMsg = sgGenMsg & sgCR & sgLF & "Key Code: " & slStr
'                sgCMCTitle(0) = "Ok"
'                sgCMCTitle(1) = "Cancel"
'                sgCMCTitle(2) = ""
'                sgCMCTitle(3) = ""
'                igDefCMC = 1
'                igEditBox = 1
'                sgEditValue = ""
'                On Error Resume Next
'                GenMsg.Show vbModal
'                If (igAnsCMC = 0) And (Len(sgEditValue) = 9) Then    'Ok, test value in sgEditValue
'                    'decode input
'                    ilPattern = ((Asc(sgEditValue) - Asc("A")) Mod 4) + 1
'                    Select Case ilPattern
'                        Case 1  'v1d1v2d2v3d3v4d4
'                            slInValue = Mid$(sgEditValue, 2, 1) & Mid$(sgEditValue, 4, 1) & Mid$(sgEditValue, 6, 1) & Mid$(sgEditValue, 8, 1)
'                            slDays = Mid$(sgEditValue, 3, 1) & Mid$(sgEditValue, 5, 1) & Mid$(sgEditValue, 7, 1) & Mid$(sgEditValue, 9, 1)
'                        Case 2  'v1d4v2d3v3d2v4d1
'                            slInValue = Mid$(sgEditValue, 2, 1) & Mid$(sgEditValue, 4, 1) & Mid$(sgEditValue, 6, 1) & Mid$(sgEditValue, 8, 1)
'                            slDays = Mid$(sgEditValue, 9, 1) & Mid$(sgEditValue, 7, 1) & Mid$(sgEditValue, 5, 1) & Mid$(sgEditValue, 3, 1)
'                        Case 3  'd4v4d3v3d2v2d1v1
'                            slInValue = Mid$(sgEditValue, 9, 1) & Mid$(sgEditValue, 7, 1) & Mid$(sgEditValue, 5, 1) & Mid$(sgEditValue, 3, 1)
'                            slDays = Mid$(sgEditValue, 8, 1) & Mid$(sgEditValue, 6, 1) & Mid$(sgEditValue, 4, 1) & Mid$(sgEditValue, 2, 1)
'                        Case 4  'd1v4d2v3d3v2d4v1
'                            slInValue = Mid$(sgEditValue, 9, 1) & Mid$(sgEditValue, 7, 1) & Mid$(sgEditValue, 5, 1) & Mid$(sgEditValue, 3, 1)
'                            slDays = Mid$(sgEditValue, 2, 1) & Mid$(sgEditValue, 4, 1) & Mid$(sgEditValue, 6, 1) & Mid$(sgEditValue, 8, 1)
'                    End Select
'                    slDate = Trim$(Str$(Asc(Mid$(slDays, 1, 1)) - Asc("A") + 1))
'                    slDate = slDate & "/" & Trim$(Str$(Asc(Mid$(slDays, 2, 1)) - Asc("A"))) & Trim$(Str$(Asc(Mid$(slDays, 3, 1)) - Asc("A")))
'                    If Len(Trim$(Str$(Asc(Mid$(slDays, 4, 1)) - Asc("A")))) = 1 Then
'                        slDate = slDate & "/0" & Trim$(Str$(Asc(Mid$(slDays, 4, 1)) - Asc("A")))
'                    Else
'                        slDate = slDate & "/" & Trim$(Str$(Asc(Mid$(slDays, 4, 1)) - Asc("A")))
'                    End If
'                    If (slNxtValue = slInValue) And (gValidDate(slDate)) Then
'                        llNow = gDateValue(Format$(gNow(), "ddddd"))
'                        If llNow + 730 < gDateValue(slDate) Then
'                            slStr = Chr(1)
'                            'slStr = slStr & "99999" & slInValue
'                            slStr = slStr & "9"
'                            For ilLoop = 0 To 3 Step 1
'                                slStr = slStr & Trim$(Str$(Int((9 - 1 + 1) * Rnd + 1)))
'                            Next ilLoop
'                            slStr = slStr & slInValue
'                            slStr = slStr & Chr(1)
'                            slStr = slStr & "9"
'                            For ilLoop = 0 To 3 Step 1
'                                slStr = slStr & Trim$(Str$(Int((9 - 1 + 1) * Rnd + 1)))
'                            Next ilLoop
'                            tmSaf.sName = slStr
'                            ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
'                        ElseIf llNow < gDateValue(slDate) Then
'                            slStr = Chr(0)
'                            llNow = gDateValue(slDate)
'                            slNow = Trim$(Str$(llNow))
'                            Do While Len(slNow) < 5
'                                slNow = "0" & slNow
'                            Loop
'                            slStr = slStr & slNow & slInValue
'                            slStr = slStr & Chr(1)
'                            tmSaf.sName = slStr & Mid$(tmSaf.sName, 12, 5)
'                            ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
'                        End If
'                    End If
'                End If
'            End If
'            btrDestroy hmSaf
'        End If
'    End If
End Sub

Private Sub cmcStartList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift Then
        Unload EngrJob
    Else
        If igListVisible Then
            EngrList.SetFocus
        Else
            EngrList.Show vbModeless
        End If
    End If
End Sub

Private Sub cmcTask_Click(Index As Integer)
    Dim ilRet As Integer
    If imPrevTaskIndex <> -1 Then
        imPrevTaskIndex = -1
    End If
    If tmcClick.Enabled = False Then
        imTaskIndex = Index
        tmcClick.Enabled = True
    End If
End Sub
Private Sub cmcTask_GotFocus(Index As Integer)
    gCtrlGotFocus cmcTask(Index)
End Sub

Private Sub cmcTask_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    imcOutline.Move cmcTask(Index).Left - 90, cmcTask(Index).Top - 90, cmcTask(Index).Width, cmcTask(Index).Height
    cmcTask(Index).ToolTipText = ""
    If imPrevTaskIndex <> -1 Then
        imPrevTaskIndex = -1
    End If
    Select Case Index
        Case RESOURCEJOB
            mSetCtrlProperties Index, RESOURCEJOB
        Case LIBRARYJOB
            mSetCtrlProperties Index, LIBRARYJOB
        Case TEMPLATEJOB
            mSetCtrlProperties Index, TEMPLATEJOB
        Case SCHEDULEJOB
            mSetCtrlProperties Index, SCHEDULEJOB
    End Select
    imcOutline.Visible = True
End Sub

Private Sub Form_Activate()
    Dim ilLoop As Integer
    mSetControls
    imTaskIndex = 100
    tmcClick.Enabled = True
    igJobVisible = True
'    If igGGFlag = 0 Then
'        pbcHor.Enabled = False
'        For ilLoop = RATECARDSJOB To REPORTSJOB - 1 Step 1
'            cmcTask(ilLoop).Enabled = False
'            lacTask(ilLoop).Enabled = False
'            lacTask(ilLoop).ForeColor = vbBlack
'        Next ilLoop
'        'cmcReport.Enabled = False
'    End If
    Me.ZOrder 0 'Send to front
    pbcHor.Visible = False
    'DoEvents
    pbcHor.Visible = True
    
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Load()
    Dim ilLoop As Integer
    If igListVisible Then
        Unload EngrList
    End If
    igJobVisible = True
    imTaskIndex = -1
    imPrevTaskIndex = -1
    plcWindow.Left = 0
    plcWindow.Top = 0
    EngrJob.Height = plcWindow.Height + fgPanelAdj ' \ 2 + 30
    EngrJob.Width = plcWindow.Width + fgPanelAdj ' \ 2 + 30
    If lgCurrBPP <= 8 Then
        pbcHor.Picture = EngrMain!pbc256Job
    End If
    gCenterForm EngrJob
End Sub

Private Sub lacChange_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    igJobVisible = False
End Sub

Private Sub pbcClickFocus_GotFocus()
    imcOutline.Visible = False
    If imPrevTaskIndex <> -1 Then
        imPrevTaskIndex = -1
    End If
End Sub

Private Sub plcBasic10_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcJob_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    imcOutline.Visible = False
    If imPrevTaskIndex <> -1 Then
        imPrevTaskIndex = -1
    End If
End Sub

Private Sub tmcClick_Timer()
    
    
    tmcClick.Enabled = False
    Select Case imTaskIndex
        Case RESOURCEJOB
            imcOutline.Visible = False
            On Error Resume Next
            If igJobShowing(RESOURCEJOB) Then
                EngrRS.SetFocus
            Else
                EngrRS.Show vbModeless  'This is required to activate form if previousely activated
            End If
        Case LIBRARYJOB
            imcOutline.Visible = False
            On Error Resume Next
            If igJobShowing(LIBRARYJOB) = 1 Then
                EngrLib.SetFocus
            ElseIf igJobShowing(LIBRARYJOB) = 2 Then
                EngrLibDef.SetFocus
            Else
                EngrLib.Show vbModeless  'This is required to activate form if previousely activated
            End If
        Case TEMPLATEJOB
            imcOutline.Visible = False
            On Error Resume Next
            If igJobShowing(TEMPLATEJOB) = 1 Then
                EngrTemp.SetFocus
            ElseIf igJobShowing(TEMPLATEJOB) = 2 Then
                EngrTempDef.SetFocus
            Else
                EngrTemp.Show vbModeless  'This is required to activate form if previousely activated
            End If
        Case SCHEDULEJOB
            imcOutline.Visible = False
            On Error Resume Next
            If igJobShowing(SCHEDULEJOB) = 1 Then
                EngrSchd.SetFocus
            ElseIf igJobShowing(SCHEDULEJOB) = 2 Then
                EngrSchdDef.SetFocus
            ElseIf igJobShowing(SCHEDULEJOB) = 3 Then
                EngrAsAirCompare.SetFocus
            ElseIf igJobShowing(SCHEDULEJOB) = 4 Then
                EngrImportAsAir.SetFocus
            Else
                EngrSchd.Show vbModeless  'This is required to activate form if previousely activated
            End If
    End Select
    imcOutline.Visible = False
    imTaskIndex = -1
End Sub
Private Sub mSetControls()
    Dim ilCount As Integer
    mSetCtrlEnable RESOURCEINDEX, RESOURCEJOB
    mSetCtrlEnable LIBRARYINDEX, LIBRARYJOB
    mSetCtrlEnable TEMPLATEINDEX, TEMPLATEJOB
    mSetCtrlEnable SCHEDULEINDEX, SCHEDULEJOB
    ilCount = gGetCheckStatus()
    If ilCount > 0 Then
        lacTask(SCHEDULEJOB).ForeColor = vbRed
        EngrMain!imcTask(SCHEDULEJOB).ForeColor = vbRed
    Else
        lacTask(SCHEDULEJOB).ForeColor = vbButtonText
        EngrMain!imcTask(SCHEDULEJOB).ForeColor = vbButtonText
    End If
End Sub
