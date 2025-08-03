VERSION 5.00
Begin VB.Form Sports 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3720
   ClientLeft      =   135
   ClientTop       =   135
   ClientWidth     =   7770
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
   ScaleHeight     =   3720
   ScaleWidth      =   7770
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   405
      Top             =   2820
   End
   Begin VB.PictureBox plcSports 
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
      Height          =   3720
      Left            =   0
      ScaleHeight     =   3720
      ScaleWidth      =   7770
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7770
      Begin VB.PictureBox pbcHor 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   3510
         Index           =   0
         Left            =   90
         Picture         =   "Sports.frx":0000
         ScaleHeight     =   3450
         ScaleWidth      =   7515
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   7575
         Begin VB.Timer tmcTerminate 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   1020
            Top             =   2805
         End
         Begin VB.Label imcOutline 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   570
            Index           =   0
            Left            =   5505
            TabIndex        =   9
            Top             =   255
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "Reports"
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
            Left            =   3210
            TabIndex        =   5
            Top             =   2130
            Width           =   1260
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "Add Multimedia Inventory"
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
            Left            =   1905
            TabIndex        =   3
            Top             =   1470
            Width           =   3915
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Proposal"
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
            Left            =   2715
            TabIndex        =   1
            Top             =   810
            Width           =   2280
         End
      End
      Begin VB.PictureBox pbcHor 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   3510
         Index           =   2
         Left            =   90
         Picture         =   "Sports.frx":6B03
         ScaleHeight     =   3450
         ScaleWidth      =   7515
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   120
         Width           =   7575
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Exit"
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
            Index           =   5
            Left            =   3270
            TabIndex        =   17
            Top             =   2430
            Width           =   990
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post KAAA-AM Invoices"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   3120
            TabIndex        =   16
            Top             =   1440
            Width           =   1320
         End
         Begin VB.Label imcOutline 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   570
            Index           =   2
            Left            =   5085
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.PictureBox pbcHor 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   3510
         Index           =   1
         Left            =   90
         Picture         =   "Sports.frx":D299
         ScaleHeight     =   3450
         ScaleWidth      =   7515
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   7575
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Live Log"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   3075
            TabIndex        =   12
            Top             =   1470
            Width           =   1320
         End
         Begin VB.Label imcOutline 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   570
            Index           =   1
            Left            =   5085
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   165
         ScaleHeight     =   165
         ScaleWidth      =   75
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3855
         Width           =   75
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Enter &Proposals"
         Height          =   570
         Index           =   0
         Left            =   2625
         TabIndex        =   2
         Tag             =   "Create new Rate cards or modify existing ones"
         Top             =   825
         Width           =   2520
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Add Multimedia Inventory"
         Height          =   570
         Index           =   1
         Left            =   1860
         TabIndex        =   4
         Tag             =   "Create or modify an Agency Proposal"
         Top             =   1500
         Width           =   4080
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "&Reports"
         Height          =   570
         Index           =   2
         Left            =   3030
         TabIndex        =   6
         Tag             =   "Add, Change, Cancel or Display contracts"
         Top             =   2175
         Width           =   1710
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "&Live Log"
         Height          =   570
         Index           =   3
         Left            =   3045
         TabIndex        =   13
         Tag             =   "Add, Change, Cancel or Display contracts"
         Top             =   1500
         Width           =   1710
      End
   End
End
Attribute VB_Name = "Sports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Sports.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Sports.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Basic 10 task buttons screen code
Option Explicit
Option Compare Text

Dim imTaskIndex As Integer
Dim imPrevTaskIndex As Integer

Dim tmSaf As SAF
Dim hmSaf As Integer
Dim imSafRecLen As Integer
Dim cmt_rst As ADODB.Recordset
 

Private Sub cmcTask_Click(Index As Integer)
    If imPrevTaskIndex <> -1 Then
        'lacTask(imPrevTaskIndex).ForeColor = vbBlack
        imPrevTaskIndex = -1
    End If
    If tmcClick.Enabled = False Then
        imTaskIndex = Index
        tmcClick.Enabled = True
    End If
End Sub
Private Sub cmcTask_GotFocus(Index As Integer)
    gCtrlGotFocus cmcTask(Index)
    Select Case Index + 200
    End Select
End Sub

Private Sub cmcTask_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (igSportsSystem = 1) And (Index <= 2) Then
        imcOutline(0).Move cmcTask(Index).Left - 90, cmcTask(Index).Top - 90, cmcTask(Index).Width, cmcTask(Index).Height
        cmcTask(Index).ToolTipText = ""
        If imPrevTaskIndex <> -1 Then
            'lacTask(imPrevTaskIndex).ForeColor = vbBlack
            imPrevTaskIndex = -1
        End If
        Select Case Index + 200
            Case SPORTPROPOSAL
                If tgSpf.sGUsePropSys = "Y" Then
                    If igWinStatus(PROPOSALSJOB) = 0 Then
                        cmcTask(Index).ToolTipText = "Disallowed"
                    Else
                        If igWinStatus(PROPOSALSJOB) = 1 Then
                            cmcTask(Index).ToolTipText = "View Only"
                            imPrevTaskIndex = Index
                            'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                        Else
                            cmcTask(Index).ToolTipText = "Updates Allowed"
                        End If
                    End If
                Else
                    cmcTask(Index).ToolTipText = "Disallowed"
                End If
            Case SPORTINVENTORY
                If igWinStatus(PROGRAMMINGJOB) = 0 Then
                    cmcTask(Index).ToolTipText = "Disallowed"
                Else
                    If igWinStatus(PROGRAMMINGJOB) = 1 Then
                        cmcTask(Index).ToolTipText = "View Only"
                        imPrevTaskIndex = Index
                        'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                    Else
                        cmcTask(Index).ToolTipText = "Updates Allowed"
                    End If
                End If
        End Select
        imcOutline(0).Visible = True
    ElseIf (igSportsSystem = 2) And (Index = 3) Then
        lacTask(3).Caption = "Live Log"
        cmcTask(3).Move lacTask(3).Left - 30, lacTask(3).Top + 30, lacTask(3).Width + 180
        imcOutline(1).Move cmcTask(3).Left - 90, cmcTask(3).Top - 90, cmcTask(3).Width, cmcTask(3).Height
        cmcTask(3).ToolTipText = "Update Allowed"
        If imPrevTaskIndex <> -1 Then
            'lacTask(imPrevTaskIndex).ForeColor = vbBlack
            imPrevTaskIndex = -1
        End If
        imcOutline(1).Visible = True
    ElseIf (igSportsSystem = 3) Then
        'lacTask(4).Caption = "Post " & sgStationCallLetters & " Invoices"
        'cmcTask(3).Move lacTask(4).Left - 30, lacTask(4).Top + 30, lacTask(4).Width + 180
        If (Index = 3) Then
            imcOutline(2).Move cmcTask(3).Left - 90, cmcTask(3).Top - 90, cmcTask(3).Width, cmcTask(3).Height
            cmcTask(3).ToolTipText = "Update Allowed"
        ElseIf (Index = 2) Then
            imcOutline(2).Move cmcTask(2).Left - 90, cmcTask(2).Top - 90, cmcTask(2).Width, cmcTask(2).Height
            cmcTask(2).ToolTipText = "Update Allowed"
        End If
        If imPrevTaskIndex <> -1 Then
            'lacTask(imPrevTaskIndex).ForeColor = vbBlack
            imPrevTaskIndex = -1
        End If
        imcOutline(2).Visible = True
    End If
End Sub

Private Sub Form_Activate()
    Dim ilLoop As Integer
    'igSportsShowing = True
    igBasic10Showing = True
    gShowPartBranner
    mSetControls
    If igSportsSystem = 1 Then
        imTaskIndex = 100
    Else
        imTaskIndex = -1
    End If
    tmcClick.Enabled = True
    Me.KeyPreview = True
    Me.ZOrder 0 'Send to front
    If (igSportsSystem = 2) Then
        pbcHor(2).Visible = False
        pbcHor(1).Visible = False
        pbcHor(0).Visible = False
    ElseIf (igSportsSystem = 3) Then
        pbcHor(2).Visible = False
        pbcHor(1).Visible = False
        pbcHor(0).Visible = False
    Else
        pbcHor(2).Visible = False
        pbcHor(0).Visible = False
        pbcHor(1).Visible = False
    End If
    DoEvents
    If igSportsSystem = 2 Then
        pbcHor(1).Visible = True
        pbcHor(0).Visible = False
        pbcHor(2).Visible = False
        For ilLoop = SPORTPROPOSAL To SPORTINVENTORY Step 1
            cmcTask(ilLoop - 200).Enabled = False
            lacTask(ilLoop - 200).Enabled = False
            lacTask(ilLoop - 200).ForeColor = vbBlack
        Next ilLoop
        cmcTask(2).Enabled = False
        lacTask(2).Enabled = False
        lacTask(2).ForeColor = vbBlack
        lacTask(4).Enabled = False
        lacTask(5).Enabled = False
        cmcTask(3).Enabled = True
        lacTask(3).Enabled = True
        lacTask(3).ForeColor = vbBlack
        imcOutline(0).Visible = False
        lacTask(3).Caption = "Live Log"
        cmcTask(3).Move lacTask(3).Left - 30, lacTask(3).Top + 30, lacTask(3).Width + 180
        imcOutline(1).Move cmcTask(3).Left - 90, cmcTask(3).Top - 90, cmcTask(3).Width, cmcTask(3).Height
    ElseIf igSportsSystem = 3 Then
        pbcHor(2).Visible = True
        pbcHor(0).Visible = False
        pbcHor(1).Visible = False
        For ilLoop = SPORTPROPOSAL To SPORTINVENTORY Step 1
            cmcTask(ilLoop - 200).Enabled = False
            lacTask(ilLoop - 200).Enabled = False
            lacTask(ilLoop - 200).ForeColor = vbBlack
        Next ilLoop
        cmcTask(2).Enabled = False
        lacTask(2).Enabled = False
        lacTask(2).ForeColor = vbBlack
        cmcTask(3).Enabled = True
        lacTask(3).Enabled = False
        lacTask(4).Enabled = True
        lacTask(4).ForeColor = vbBlack
        lacTask(5).Enabled = True
        lacTask(5).ForeColor = vbBlack
        imcOutline(2).Visible = False
        lacTask(4).Caption = "Post " & sgStationCallLetters & " Invoices"
        cmcTask(3).Move lacTask(4).Left - 30, lacTask(4).Top + 30, lacTask(4).Width + 180
        cmcTask(2).Caption = "Exit"
        cmcTask(2).Enabled = True
        cmcTask(2).Move lacTask(5).Left - 30, lacTask(5).Top + 30, lacTask(5).Width + 180
        imcOutline(2).Move cmcTask(3).Left - 90, cmcTask(3).Top - 90, cmcTask(3).Width, cmcTask(3).Height
    Else
        pbcHor(0).Visible = True
        pbcHor(1).Visible = False
        pbcHor(2).Visible = False
        For ilLoop = SPORTPROPOSAL To SPORTINVENTORY Step 1
            cmcTask(ilLoop - 200).Enabled = True
            lacTask(ilLoop - 200).Enabled = True
        Next ilLoop
        cmcTask(2).Enabled = True
        lacTask(2).Enabled = True
        cmcTask(3).Enabled = False
        lacTask(3).Enabled = False
        lacTask(4).Enabled = False
        lacTask(5).Enabled = False
        lacTask(3).ForeColor = vbBlack
        imcOutline(1).Visible = False
    End If
    If igGGFlag = 0 Then
        pbcHor(0).Enabled = False
        pbcHor(1).Enabled = False
        pbcHor(2).Enabled = False
        For ilLoop = SPORTPROPOSAL To SPORTINVENTORY Step 1
            cmcTask(ilLoop - 200).Enabled = False
            lacTask(ilLoop - 200).Enabled = False
            lacTask(ilLoop - 200).ForeColor = vbBlack
        Next ilLoop
        'cmcReport.Enabled = False
        cmcTask(3).Enabled = False
        lacTask(3).Enabled = False
        lacTask(3).ForeColor = vbBlack
        If (igGGFlag = 0) And (igRptGGFlag = 0) Then
            cmcTask(2).Enabled = False
            lacTask(2).Enabled = False
            lacTask(2).ForeColor = vbBlack
        End If
    End If
    DoEvents

End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
'        gFunctionKeyBranch KeyCode
'    'ElseIf KeyCode = KEYF5 Then
'    '    AuxWnd.Show vbModeless
'    End If
'    If (KeyCode = vbKey8) And (Shift = 4) Then
'        igBasicOrAux = -1
'        'igSportsShowing = False
'        igBasic10Showing = True
'        Unload Sports
'    End If
End Sub

Private Sub Form_Load()
    If igBasicOrAux = 2 Then
        Unload AuxWnd
    End If
    'igSportsShowing = True
    igBasic10Showing = True
    igBasicOrAux = 1    'Indicate that basic 10 loaded
    imTaskIndex = -1
    imPrevTaskIndex = -1
    'Sports.Caption = "Jobs"
    'cmcCSLogo.Picture = IconTraf!imcCSLogo.Picture
    plcSports.Left = 0
    plcSports.Top = 0
    Sports.Height = plcSports.Height + fgPanelAdj ' \ 2 + 30
    Sports.Width = plcSports.Width + fgPanelAdj ' \ 2 + 30
    'If lgCurrBPP <= 8 Then
    '    pbcHor.Picture = Traffic!pbc256Job
    'End If
    'cmcCSLogo.Picture = IconTraf!imcCSLogo.Picture
    gCenterForm Sports
End Sub




Private Sub Form_Unload(Cancel As Integer)
    'igSportsShowing = False
    igBasic10Showing = False
End Sub

Private Sub pbcClickFocus_GotFocus()
    imcOutline(0).Visible = False
    imcOutline(1).Visible = False
    imcOutline(2).Visible = False
    If imPrevTaskIndex <> -1 Then
        'lacTask(imPrevTaskIndex).ForeColor = vbBlack
        imPrevTaskIndex = -1
    End If
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcSports_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcSports_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline(0).Visible = False
    imcOutline(1).Visible = False
    imcOutline(2).Visible = False
    If imPrevTaskIndex <> -1 Then
        'lacTask(imPrevTaskIndex).ForeColor = vbBlack
        imPrevTaskIndex = -1
    End If
End Sub

Private Sub tmcClick_Timer()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slCommand                                                                             *
'******************************************************************************************


    Dim slStr As String        'General string
    Dim ilLoop As Integer
    Dim ilRptType As Integer
    Dim ilAskMkt As Integer
    Dim dlShellRet As Double
    Dim slCommandStr As String
    Dim ilPos As Integer
    Dim slDate As String
    Dim ilRet As Integer

    tmcClick.Enabled = False
    gShowHelpMess tgJobHelp(), -1   'Blank help
    mCheckGG
    If igSportsSystem = 1 Then
        '*** Market Code
        If (tgSpf.sMktBase = "Y") And ((imTaskIndex = SPORTPROPOSAL - 200)) Then
            ilAskMkt = True
            If imTaskIndex = SPORTPROPOSAL - 200 Then
                If igJobShowing(PROPOSALSJOB) Then
                    ilAskMkt = False
                End If
            End If
            If ilAskMkt Then
                LogMkt.Show vbModal
                DoEvents
                If UBound(igMktCode) <= 0 Then
                    plcSports.Visible = False
                    plcSports.Visible = True
                    imTaskIndex = -1
                    Exit Sub
                End If
                If (imTaskIndex = SPORTPROPOSAL - 200) Then
                    ReDim igCntrMktCode(0 To UBound(igMktCode)) As Integer
                    For ilLoop = 0 To UBound(igMktCode) Step 1
                        igCntrMktCode(ilLoop) = igMktCode(ilLoop)
                    Next ilLoop
                End If
            End If
        End If
        '*** Market Code
        Select Case imTaskIndex
            Case SPORTPROPOSAL - 200
                imcOutline(0).Visible = False
                If (igGGFlag <> 0) Then
                    On Error Resume Next
                    If igJobShowing(CONTRACTSJOB) Then
                        Contract.SetFocus
                    Else
                        igAlertCntrStatus = -1
                        sgCntrScreen = "Proposals"
                        Contract.Show vbModeless  'This is required to activate form if previousely activated
                    End If
                End If
            Case SPORTINVENTORY - 200
                imcOutline(0).Visible = False
                If (igGGFlag <> 0) Then
        '            Traffic!mnuRemoteUser.Enabled = False
                    On Error Resume Next
                    igGameSchdVefCode = -1
                    If igJobShowing(PROGRAMMINGJOB) Then
                        GameInv.SetFocus
                    Else
                        GameInv.Show vbModeless  'This is required to activate form if previousely activated
                    End If
                End If
            Case SPORTREPORT - 200
                imcOutline(0).Visible = False
                ilRptType = 0
                If (igGGFlag <> 0) Or (igRptGGFlag <> 0) Then
                    igRptCallType = -2  'CONTRACTSJOB
                    If igTestSystem Then
                        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(ilRptType))
                    Else
                        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(ilRptType))
                    End If
                    sgCommandStr = slStr
                    sgSpfStamp = "~"    'Force read
                    gSpfRead
                    'gSetUsingTraffic
                    On Error Resume Next
                    RptList.Show vbModal    'vbModeless  'This is required to activate form if previousely activated
                End If
        End Select
    ElseIf igSportsSystem = 3 Then
        If imTaskIndex = 3 Then
            imcOutline(2).Visible = False
            If (igGGFlag <> 0) Then
                igAnsCMC = 0
                'SQLQuery = "SELECT * FROM CMT Where cmtType = " & "'C'"
                SQLQuery = "Select * From cxf_Hdr_Ln_Comments Where cxfCode  = " & tgSaf(0).lCitationComment
                'Set cmt_rst = cnn.Execute(SQLQuery)
                Set cmt_rst = gSQLSelectCall(SQLQuery)
                If Not cmt_rst.EOF Then
                    slStr = gStripChr0(cmt_rst!cxfComment)
                Else
                    slStr = "By pressing 'I Agree', I hereby acknowledge and attest that the information I am about to enter is accurate and true, and that the commercials shown did air, unedited, at the indicated times."
                End If
                Do
                    '4/30/19: Moved above so that a default citation can be used
                    ''slStr = Trim$(cmt_rst!cmtPart1 & cmt_rst!cmtPart2 & cmt_rst!cmtPart3 & cmt_rst!cmtPart4)
                    'slStr = gStripChr0(cmt_rst!cxfComment)
                    sgGenMsg = slStr
                    sgCMCTitle(0) = "I Agree"
                    sgCMCTitle(1) = "Disagree"
                    sgCMCTitle(2) = ""
                    sgCMCTitle(3) = ""
                    igDefCMC = 0
                    igEditBox = 1
                    sgEditValue = ""
                    GenMsg.Show vbModal
                    DoEvents
                Loop While (sgEditValue = "") And (igAnsCMC = 0)
                If igAnsCMC = 0 Then
                    sgRemoteUserName = sgEditValue
                    'PostManualVeh.Show vbModal
                    'If igManualPostVefCode > 0 Then
                    '    PostManualCntr.Show vbModal
                    'End If
                    gShellToPostManual Sports, "RM/" & sgRemoteUserName & "SCL/" & sgStationCallLetters & "SPW/" & sgStationPW & " FROM/SPORTSSYSTEM"
                End If
            End If
        ElseIf imTaskIndex = 2 Then
            tmcTerminate.Enabled = True
        End If
    Else
        If imTaskIndex = 3 Then
            If (igGGFlag <> 0) Then
                'LiveLog.Show vbModal
                If igTestSystem Then
                    slCommandStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
                Else
                    slCommandStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
                End If
                If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
                    ilPos = InStr(1, slCommandStr, "Guide", vbTextCompare)
                    If ilPos > 0 Then
                        slCommandStr = Left(slCommandStr, ilPos - 1) & "CSI" & Mid(slCommandStr, ilPos + 5)
                    End If
                End If
                'Dan M 9/20/10 problems in v57 reports.exe running GetCsiName
                'slDate = Trim$(gGetCSIName("SYSDate"))
                slDate = gCSIGetName()
                If slDate <> "" Then
                    'use slDate when writing to file later
                    slDate = " /D:" & slDate
                    'slCommandStr = slCommandStr & " /D:" & slDate
                    slCommandStr = slCommandStr & slDate
                End If
                slCommandStr = slCommandStr & " /ULF:" & lgUlfCode
                On Error GoTo LoadErr
                ilRet = 0
                AppActivate "LiveLog"
                If ilRet = 1 Then
                    'ilShell = Shell(sgExePath & "Reports.Exe " & slCommandStr, vbNormalFocus)
                    dlShellRet = Shell(sgExePath & "LiveLog.Exe " & slCommandStr, vbNormalFocus)
                End If
                tmcTerminate.Enabled = True
                Exit Sub
            End If
        End If
    End If
    imcOutline(0).Visible = False
    imcOutline(1).Visible = False
    imcOutline(2).Visible = False
    imTaskIndex = -1
    Exit Sub
LoadErr:
    ilRet = 1
    Resume Next
    Exit Sub
End Sub
Private Sub plcSports_Paint()
    'plcSports.CurrentX = cmcCSLogo.Left + (3 * cmcCSLogo.Width) / 2
    'plcSports.CurrentY = cmcCSLogo.Top + cmcCSLogo.Height / 3
    'plcSports.Print "Counterpoint Software ®"
End Sub

Private Sub mSetControls()
    Dim ilLoop As Integer

    If igSportsSystem = 1 Then
        For ilLoop = SPORTPROPOSAL To SPORTINVENTORY Step 1
            If ilLoop = SPORTPROPOSAL Then
                If tgSpf.sGUsePropSys = "Y" Then
                    If igWinStatus(PROPOSALSJOB) = 0 Then
                        cmcTask(ilLoop - 200).Enabled = False
                        lacTask(ilLoop - 200).Enabled = False
                    Else
                        cmcTask(0).Enabled = True
                    End If
                Else
                    cmcTask(ilLoop - 200).Enabled = False
                    lacTask(ilLoop - 200).Enabled = False
                End If
            ElseIf ilLoop = SPORTINVENTORY Then
                If igWinStatus(PROGRAMMINGJOB) = 0 Then
                    cmcTask(ilLoop - 200).Enabled = False
                    lacTask(ilLoop - 200).Enabled = False
                Else
                    cmcTask(ilLoop - 200).Enabled = True
                End If
            End If
        Next ilLoop
        cmcTask(3).Enabled = False
        lacTask(3).Enabled = False
        lacTask(4).Enabled = False
        lacTask(5).Enabled = False
    ElseIf igSportsSystem = 3 Then
        cmcTask(0).Enabled = False
        lacTask(0).Enabled = False
        cmcTask(1).Enabled = False
        lacTask(1).Enabled = False
        cmcTask(2).Enabled = False
        lacTask(2).Enabled = False
        cmcTask(3).Enabled = True
        lacTask(3).Enabled = True
    Else
        cmcTask(0).Enabled = False
        lacTask(0).Enabled = False
        cmcTask(1).Enabled = False
        lacTask(1).Enabled = False
        cmcTask(2).Enabled = False
        lacTask(2).Enabled = False
        cmcTask(3).Enabled = True
        lacTask(3).Enabled = True
        lacTask(4).Enabled = False
        lacTask(5).Enabled = False
    End If
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    
    Unload Sports
    Set Sports = Nothing   'Remove data segment
    Unload Traffic
End Sub

Private Sub mCheckGG()
    Dim ilRet As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llNow As Long
    Dim llDate As Long
    Dim slStr As String
    Dim ilLoop As Integer
    
    On Error Resume Next
    
    'If imLastHourGGChecked = Hour(Now) Then
    '    Exit Sub
    'End If
    'imLastHourGGChecked = Hour(Now)
    
    hmSaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    imSafRecLen = Len(tmSaf)
    ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    ilField1 = Asc(tmSaf.sName)
    slStr = Mid$(tmSaf.sName, 2, 5)
    llDate = Val(slStr)
    ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
    llNow = gDateValue(Format$(Now, "m/d/yy"))
    If (ilField1 = 0) And (ilField2 = 1) Then
        If llDate <= llNow Then
            ilField2 = 0
        End If
    End If
    If (ilField1 = 0) And (ilField2 = 0) Then
        igGGFlag = 0
    End If
    gSetRptGGFlag tmSaf
    btrDestroy hmSaf
End Sub
