VERSION 5.00
Begin VB.Form Calendar 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar"
   ClientHeight    =   3060
   ClientLeft      =   660
   ClientTop       =   3405
   ClientWidth     =   2505
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   2505
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   3060
      Left            =   0
      ScaleHeight     =   3060
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   0
      Width           =   2505
      Begin VB.VScrollBar sbcYear 
         Height          =   285
         Left            =   1830
         Max             =   2
         TabIndex        =   19
         Top             =   315
         Value           =   1
         Width           =   150
      End
      Begin VB.VScrollBar sbcMonth 
         Height          =   285
         Left            =   885
         Max             =   2
         TabIndex        =   18
         Top             =   315
         Value           =   1
         Width           =   150
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
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   75
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   285
         Width           =   75
      End
      Begin VB.CommandButton cmcAcross 
         Appearance      =   0  'Flat
         Caption         =   "D&own"
         Height          =   285
         HelpContextID   =   1
         Left            =   90
         TabIndex        =   13
         Top             =   2655
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmcDone 
         Appearance      =   0  'Flat
         Caption         =   "&Done"
         Height          =   285
         HelpContextID   =   1
         Left            =   840
         TabIndex        =   10
         Top             =   2655
         Width           =   795
      End
      Begin VB.CheckBox ckcYear 
         Caption         =   "&Year"
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
         Left            =   1725
         TabIndex        =   9
         Top             =   75
         Width           =   720
      End
      Begin VB.PictureBox pbcControls 
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
         Height          =   495
         Left            =   165
         ScaleHeight     =   495
         ScaleWidth      =   2250
         TabIndex        =   4
         Top             =   2145
         Width           =   2250
         Begin VB.OptionButton rbcMonth 
            Caption         =   "Co&rp"
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
            Left            =   1260
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton rbcMonth 
            Caption         =   "Jul&ian"
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
            Left            =   240
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   825
         End
         Begin VB.OptionButton rbcMonth 
            Caption         =   "&Cal"
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
            Left            =   630
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton rbcMonth 
            Caption         =   "&Std"
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
            Left            =   15
            TabIndex        =   5
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton rbcMonth 
            Caption         =   "J&ulian"
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
            Left            =   1245
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   825
         End
         Begin VB.Image imcJulianUp 
            Appearance      =   0  'Flat
            Height          =   135
            Left            =   2070
            Picture         =   "Calendar.frx":0000
            Top             =   45
            Width           =   165
         End
         Begin VB.Image imcJulianDn 
            Appearance      =   0  'Flat
            Height          =   135
            Left            =   1065
            Picture         =   "Calendar.frx":00CA
            Top             =   285
            Width           =   165
         End
      End
      Begin VB.TextBox edcMonth 
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
         Left            =   525
         MaxLength       =   2
         TabIndex        =   1
         Top             =   315
         Width           =   360
      End
      Begin VB.TextBox edcYear 
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
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   2
         Top             =   315
         Width           =   630
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Index           =   0
         Left            =   330
         Picture         =   "Calendar.frx":0194
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   660
         Width           =   1875
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
         Left            =   1800
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2655
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
         Left            =   1875
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2610
         Visible         =   0   'False
         Width           =   525
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
         Left            =   1590
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2730
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lacMonth 
         Alignment       =   2  'Center
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
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   8
         Top             =   255
         Visible         =   0   'False
         Width           =   1860
      End
   End
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Calendar.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Calendar.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Calendar input screen code
Option Explicit
Option Compare Text
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imSettingCD As Integer     'True- don't display
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imMonth As Integer
Dim imYear As Integer
Dim imFirstActivate As Integer




Private Sub ckcYear_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcYear.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    If Value Then
        If Not igCalByYear Then
            fgCalLeft = Calendar.Left
            fgCalTop = Calendar.Top
        End If
        igCalByYear = True
        mSetCalendar
    Else
        igCalByYear = False
        mSetCalendar
    End If
End Sub
Private Sub cmcAcross_Click()
    If igCalAcross Then
        cmcAcross.Caption = "Acr&oss"
        igCalAcross = False
    Else
        cmcAcross.Caption = "D&own"
        igCalAcross = True
    End If
    mSetCalendar
End Sub
Private Sub cmcDone_Click()
    igCalMonth = Val(edcMonth.Text)
    igCalYear = Val(edcYear.Text)
    If Not igCalByYear Then
        fgCalLeft = Calendar.Left
        fgCalTop = Calendar.Top
    End If
    igCalActive = 0 'False
    mTerminate
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcMonth_Change()
    If Not igCalByYear Then
        pbcCalendar_Paint (0)
    End If
End Sub
Private Sub edcMonth_GotFocus()
    gCtrlGotFocus ActiveControl
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
End Sub
Private Sub edcMonth_KeyPress(KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcYear_Change()
    Dim ilLoop As Integer
    If igCalByYear Then
        For ilLoop = 0 To 11 Step 1
            pbcCalendar_Paint (ilLoop)
        Next ilLoop
    Else
        pbcCalendar_Paint (0)
    End If
End Sub
Private Sub edcYear_GotFocus()
    gCtrlGotFocus ActiveControl
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
End Sub
Private Sub edcYear_KeyPress(KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        'Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    igCalActive = 1 'True
    'gShowBranner
'    If (Forms.Count = igNoMinForms + igCalActive + igCalcActive) And Not igShowPicture Then 'Show basic 10 if only this form to showing
'        Traffic!pbcMsgArea.SetFocus
'    End If
    DoEvents    'Process events so pending keys are not sent to this
                'form when keypreview turn on
    'Calendar.KeyPreview = True  'To get Alt J and Alt L keys
    If igCalByYear Then
        If ckcYear.Value = vbChecked Then
            mSetCalendar
        Else
            ckcYear.Value = vbChecked
        End If
    Else
        pbcCalendar(0).Cls
        If ckcYear.Value = vbChecked Then
            ckcYear.Value = vbUnchecked
        Else
            mSetCalendar
        End If
    End If
    Calendar.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub Form_Deactivate()
    'Calendar.KeyPreview = False
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'If ((Shift And vbAltMask) > 0) And (KeyCode = 74) Then    'J=74
    '    Calendar.KeyPreview = False
    '    Traffic!gpcBasicWnd.Value = True   'Button up and unload
    'End If
    'If ((Shift And vbAltMask) > 0) And (KeyCode = 76) Then    'L=76
    '    Calendar.KeyPreview = False
    '    Traffic!gpcAuxWnd.Value = True   'Button up and unload
    'End If
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Calendar = Nothing   'Remove data segment
End Sub
Private Sub imcJulianDn_Click()
    rbcMonth(3).Value = True
End Sub
Private Sub imcJulianUp_Click()
    rbcMonth(2).Value = True
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
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imFirstActivate = True
    'mParseCmmdLine
    imFirstFocus = True
    For ilLoop = 1 To 11 Step 1
        Load pbcCalendar(ilLoop)
        Load lacMonth(ilLoop)
    Next ilLoop
    For ilLoop = 0 To 11 Step 1
        Select Case ilLoop
            Case 0
                lacMonth(ilLoop).Caption = "January"
            Case 1
                lacMonth(ilLoop).Caption = "February"
            Case 2
                lacMonth(ilLoop).Caption = "March"
            Case 3
                lacMonth(ilLoop).Caption = "April"
            Case 4
                lacMonth(ilLoop).Caption = "May"
            Case 5
                lacMonth(ilLoop).Caption = "June"
            Case 6
                lacMonth(ilLoop).Caption = "July"
            Case 7
                lacMonth(ilLoop).Caption = "August"
            Case 8
                lacMonth(ilLoop).Caption = "September"
            Case 9
                lacMonth(ilLoop).Caption = "October"
            Case 10
                lacMonth(ilLoop).Caption = "November"
            Case 11
                lacMonth(ilLoop).Caption = "December"
        End Select
    Next ilLoop
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    If edcMonth.Text = "" Then
        imSettingCD = True
        edcMonth.Text = Trim$(str$(igCalMonth))
        imSettingCD = True
        edcYear.Text = Trim$(str$(igCalYear))
    End If
    imSettingCD = False
    If Not igCalByYear Then 'Set calendar location as rbcMonth will save location
        Calendar.Move fgCalLeft, fgCalTop
    End If
    ilRet = gObtainCorpCal()
    If tgSpf.sRUseCorpCal <> "Y" Then
        rbcMonth(4).Enabled = False
        If igCalMonthType = 4 Then
            igCalMonthType = 0
        End If
    Else
        rbcMonth(4).Enabled = True
    End If
    rbcMonth(igCalMonthType).Value = True
    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCalendar                    *
'*                                                     *
'*             Created:10/27/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mSetCalendar()
'
'   mSetCalendar
'   Where:
'
    Dim ilLoop As Integer
    plcCalendar.Move 0, 0
    If igCalByYear Then
        Calendar.Caption = ""
        edcMonth.Visible = False
        sbcMonth.Visible = False
        pbcCalendar(0).Move 330, 480
        If igCalAcross Then
            For ilLoop = 1 To 3 Step 1
                pbcCalendar(ilLoop).Move pbcCalendar(ilLoop - 1).Left + pbcCalendar(0).Width + 30, pbcCalendar(0).Top
            Next ilLoop
            pbcCalendar(4).Move pbcCalendar(0).Left, pbcCalendar(0).Top + pbcCalendar(0).Height + 30 + lacMonth(0).Height
            For ilLoop = 5 To 7 Step 1
                pbcCalendar(ilLoop).Move pbcCalendar(ilLoop - 1).Left + pbcCalendar(0).Width + 30, pbcCalendar(4).Top
            Next ilLoop
            pbcCalendar(8).Move pbcCalendar(0).Left, pbcCalendar(4).Top + pbcCalendar(0).Height + 30 + lacMonth(0).Height
            For ilLoop = 9 To 11 Step 1
                pbcCalendar(ilLoop).Move pbcCalendar(ilLoop - 1).Left + pbcCalendar(0).Width + 30, pbcCalendar(8).Top
            Next ilLoop
            edcYear.Move pbcCalendar(1).Left + pbcCalendar(1).Width - edcYear.Width \ 2 - sbcYear.Width \ 2, 150
            sbcYear.Move edcYear.Left + edcYear.Width - 15, edcYear.Top
            pbcControls.Move pbcCalendar(0).Left, pbcCalendar(8).Top + pbcCalendar(8).Height + 60, 4065, 285
            rbcMonth(3).Move imcJulianUp.Left + imcJulianUp.Width + 75, 0
            imcJulianDn.Move rbcMonth(3).Left + rbcMonth(3).Width + 15, rbcMonth(3).Top + 45
            rbcMonth(4).Move rbcMonth(3).Left + 1020, rbcMonth(3).Top
            cmcAcross.Move pbcCalendar(3).Left + pbcCalendar(3).Width \ 2 - cmcAcross.Width \ 2, pbcControls.Top
            cmcDone.Move cmcAcross.Left - (3 * cmcDone.Width) \ 2, pbcControls.Top
        Else
            For ilLoop = 3 To 11 Step 3
                pbcCalendar(ilLoop).Move pbcCalendar(ilLoop - 3).Left + pbcCalendar(0).Width + 30, pbcCalendar(0).Top
            Next ilLoop
            pbcCalendar(1).Move pbcCalendar(0).Left, pbcCalendar(0).Top + pbcCalendar(0).Height + 30 + lacMonth(0).Height
            For ilLoop = 4 To 11 Step 3
                pbcCalendar(ilLoop).Move pbcCalendar(ilLoop - 3).Left + pbcCalendar(0).Width + 30, pbcCalendar(1).Top
            Next ilLoop
            pbcCalendar(2).Move pbcCalendar(0).Left, pbcCalendar(4).Top + pbcCalendar(0).Height + 30 + lacMonth(0).Height
            For ilLoop = 5 To 11 Step 3
                pbcCalendar(ilLoop).Move pbcCalendar(ilLoop - 3).Left + pbcCalendar(0).Width + 30, pbcCalendar(2).Top
            Next ilLoop
            edcYear.Move pbcCalendar(3).Left + pbcCalendar(3).Width - edcYear.Width \ 2 - sbcYear.Width \ 2, 150
            sbcYear.Move edcYear.Left + edcYear.Width - 15, edcYear.Top
            pbcControls.Move pbcCalendar(0).Left, pbcCalendar(2).Top + pbcCalendar(2).Height + 60, 4095, 285
            ckcYear.Move pbcCalendar(3).Left + pbcCalendar(3).Width - ckcYear.Width \ 2, pbcControls.Top + 15
            rbcMonth(3).Move imcJulianUp.Left + imcJulianUp.Width + 75, 0
            imcJulianDn.Move rbcMonth(3).Left + rbcMonth(3).Width + 15, rbcMonth(3).Top + 45
            rbcMonth(4).Move rbcMonth(3).Left + 1020, rbcMonth(3).Top
            cmcAcross.Move pbcCalendar(9).Left + pbcCalendar(9).Width \ 2 - cmcAcross.Width \ 2, pbcControls.Top
            cmcDone.Move cmcAcross.Left - (3 * cmcDone.Width) \ 2, pbcControls.Top
        End If
        For ilLoop = 0 To 11 Step 1
            lacMonth(ilLoop).Move pbcCalendar(ilLoop).Left, pbcCalendar(ilLoop).Top - lacMonth(0).Height - 15
        Next ilLoop
        If igCalAcross Then
            plcCalendar.Move 0, 0, pbcCalendar(3).Left + pbcCalendar(3).Width + 300, pbcControls.Top + pbcControls.Height + 120
        Else
            plcCalendar.Move 0, 0, pbcCalendar(9).Left + pbcCalendar(9).Width + 300, cmcDone.Top + cmcDone.Height + 120
        End If
        ckcYear.Move plcCalendar.Left + plcCalendar.Width - ckcYear.Width - 60, plcCalendar.Top + 75
        cmcAcross.Visible = True
        DoEvents
        For ilLoop = 0 To 11 Step 1
            pbcCalendar_Paint (ilLoop)
        Next ilLoop
        For ilLoop = 0 To 11 Step 1
            pbcCalendar(ilLoop).Visible = True
            lacMonth(ilLoop).Visible = True
        Next ilLoop
        Calendar.Move 100, 200, plcCalendar.Width + 115, plcCalendar.Height + 115
    Else 'Month
        Calendar.Caption = "Calendar"
        For ilLoop = 0 To 11 Step 1
            pbcCalendar(ilLoop).Visible = False
        Next ilLoop
        For ilLoop = 0 To 11 Step 1
            lacMonth(ilLoop).Visible = False
        Next ilLoop
        cmcAcross.Visible = False
        plcCalendar.Move 0, 0, 2505, 3060
        ckcYear.Move plcCalendar.Left + plcCalendar.Width - ckcYear.Width - 60, plcCalendar.Top + 75
        edcMonth.Move 510, 315
        edcMonth.Visible = True
        sbcMonth.Move 855, 315
        sbcMonth.Visible = True
        edcYear.Move 1170, 315
        sbcYear.Move 1770, 315
        pbcCalendar(0).Move 330, 660
        pbcControls.Move 180, 2145, 2250, 480
        rbcMonth(3).Move 240, 240
        imcJulianDn.Move rbcMonth(3).Left + rbcMonth(3).Width + 15, rbcMonth(3).Top + 45
        rbcMonth(4).Move rbcMonth(3).Left + 1020, rbcMonth(3).Top
        cmcDone.Move 810, 2655
        pbcCalendar(0).Visible = True
        pbcCalendar_Paint (0)
        Calendar.Move fgCalLeft, fgCalTop, plcCalendar.Width + 115, plcCalendar.Height + 435    '115
    End If
    Screen.MousePointer = vbDefault
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
    mUrfUpdateCal Calendar, tgUrf()
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Calendar
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mUrfUpdateCal                   *
'*                                                     *
'*             Created:5/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update User calendar/calculator *
'*                     related fields in the records   *
'*                                                     *
'*******************************************************
Private Sub mUrfUpdateCal(frm As Form, tlUrf() As URF)
'
'   mUrfUpdateCal MainForm, tlUrf()
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       tlUrf (O)- the updated user records
'                   Note: tlUrf must be defined as Dim tlUrf() as URF
'
    Dim ilRecLen As Integer     'URF record length
    Dim hlUrf As Integer        'User Option file handle
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim tlUrfSet As URF    'Position to record so it can be updated
    Dim tlSrchKey As INTKEY0    'URF key record image
    hlUrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mUrfUpdateCalErr
    gBtrvErrorMsg ilRet, "mUrfUpdateCal (btrOpen):" & "Urf.Btr", frm
    On Error GoTo 0
    On Error GoTo gUrfNoDefinedErr
    ilRecLen = Len(tlUrf(0))  'btrRecordLength(hlUrf)  'Get and save record length
    If tlUrf(0).iCode <= 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    For ilLoop = LBound(tlUrf) To UBound(tlUrf) Step 1
        tlSrchKey.iCode = tlUrf(ilLoop).iCode
        ilRet = btrGetEqual(hlUrf, tlUrfSet, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        gUrfDecrypt tlUrfSet
        On Error GoTo mUrfUpdateCalErr
        gBtrvErrorMsg ilRet, "mUrfUpdateCal (btrGetEqual)", frm
        On Error GoTo 0
        If Not igCalByYear Then
            tlUrf(ilLoop).sClnMoYr = "M"    'Calendar by month
        Else
            tlUrf(ilLoop).sClnMoYr = "Y"    'Calendar by month
        End If
        Select Case igCalMonthType
            Case 0
                tlUrf(ilLoop).sClnType = "S"    'Calendar type-standard
            Case 1
                tlUrf(ilLoop).sClnType = "R"    'Calendar type-calendar
            Case 2
                tlUrf(ilLoop).sClnType = "U"    'Calendar type-julian +
            Case 3
                tlUrf(ilLoop).sClnType = "D"    'Calendar type-julian -
            Case 4
                tlUrf(ilLoop).sClnType = "C"    'Calendar type-corp
        End Select
        If igCalAcross Then
            tlUrf(ilLoop).sClnLayout = "A"  'Calendar- Across
        Else
            tlUrf(ilLoop).sClnLayout = "D"  'Calendar- Down
        End If
        tlUrf(ilLoop).iClnLeft = fgCalLeft    'Calendar left
        tlUrf(ilLoop).iClnTop = fgCalTop 'Calendar top
        tlUrf(ilLoop).iClcLeft = fgCalcLeft
        tlUrf(ilLoop).iClcTop = fgCalcTop
        gUrfEncrypt tlUrf(ilLoop)
        ilRet = btrUpdate(hlUrf, tlUrf(ilLoop), ilRecLen)
        On Error GoTo mUrfUpdateCalErr
        gBtrvErrorMsg ilRet, "mUrfUpdateCal (btrUpdate)", frm
        On Error GoTo 0
        gUrfDecrypt tlUrf(ilLoop)
    Next ilLoop
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    Exit Sub
mUrfUpdateCalErr:
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    'Error ERRORCODEBASE
    Exit Sub
gUrfNoDefinedErr:
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    Exit Sub
End Sub
Private Sub pbcCalendar_Paint(Index As Integer)
    Dim ilCalType As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilAdjMonth As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If imSettingCD Then
        imSettingCD = False
        Exit Sub
    End If
    On Error GoTo pbcCalendarErr
    ilYear = Val(edcYear.Text)
    ilMonth = Val(edcMonth.Text)
    imMonth = ilMonth
    imYear = ilYear
    ilCalType = igCalMonthType
    If ilCalType = 4 Then
        ilCalType = 5
    End If
    If igCalByYear Then
        If ilCalType <> 5 Then
            Select Case Index
                Case 0
                    lacMonth(Index).Caption = "January"
                Case 1
                    lacMonth(Index).Caption = "February"
                Case 2
                    lacMonth(Index).Caption = "March"
                Case 3
                    lacMonth(Index).Caption = "April"
                Case 4
                    lacMonth(Index).Caption = "May"
                Case 5
                    lacMonth(Index).Caption = "June"
                Case 6
                    lacMonth(Index).Caption = "July"
                Case 7
                    lacMonth(Index).Caption = "August"
                Case 8
                    lacMonth(Index).Caption = "September"
                Case 9
                    lacMonth(Index).Caption = "October"
                Case 10
                    lacMonth(Index).Caption = "November"
                Case 11
                    lacMonth(Index).Caption = "December"
            End Select
            gPaintCalendar Index + 1, ilYear, ilCalType, pbcCalendar(Index), tmCDCtrls(), llStartDate, llEndDate
        Else
            ilFound = False
            For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
                If tgMCof(ilLoop).iYear = ilYear Then
                    ilAdjMonth = tgMCof(ilLoop).iStartMnthNo + Index
                    If ilAdjMonth > 12 Then
                        ilAdjMonth = ilAdjMonth - 12
                    End If
                    Select Case ilAdjMonth - 1
                        Case 0
                            lacMonth(Index).Caption = "January"
                        Case 1
                            lacMonth(Index).Caption = "February"
                        Case 2
                            lacMonth(Index).Caption = "March"
                        Case 3
                            lacMonth(Index).Caption = "April"
                        Case 4
                            lacMonth(Index).Caption = "May"
                        Case 5
                            lacMonth(Index).Caption = "June"
                        Case 6
                            lacMonth(Index).Caption = "July"
                        Case 7
                            lacMonth(Index).Caption = "August"
                        Case 8
                            lacMonth(Index).Caption = "September"
                        Case 9
                            lacMonth(Index).Caption = "October"
                        Case 10
                            lacMonth(Index).Caption = "November"
                        Case 11
                            lacMonth(Index).Caption = "December"
                    End Select
                    gPaintCalendar ilAdjMonth, ilYear, ilCalType, pbcCalendar(Index), tmCDCtrls(), llStartDate, llEndDate
                    ilFound = True
                End If
            Next ilLoop
            If Not ilFound Then
                pbcCalendar(Index).Cls
            End If
        End If
    Else
        gPaintCalendar ilMonth, ilYear, ilCalType, pbcCalendar(Index), tmCDCtrls(), llStartDate, llEndDate
    End If
    Exit Sub
pbcCalendarErr:
    pbcCalendar(Index).Cls
    Exit Sub
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcCalendar_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcMonth_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcMonth(Index).Value
    'End of coded added
    If Value Then
        If Not igCalByYear Then 'save location as mSetCalendar will reset location
            fgCalLeft = Calendar.Left
            fgCalTop = Calendar.Top
        End If
        igCalMonthType = Index
        mSetCalendar
    End If
End Sub

Private Sub sbcMonth_Change()
    Dim slStr As String
    If sbcMonth.Value = 1 Then
        Exit Sub
    End If
    If sbcMonth.Value = 2 Then
        imMonth = Val(edcMonth.Text) - 1
        If imMonth <= 0 Then
            imMonth = 12
            imSettingCD = True
            imYear = sbcYear.Value - 1
            sbcYear.Value = 2
        End If
        slStr = Trim$(str$(imMonth))
        Do While Len(slStr) < 2
            slStr = "0" & slStr
        Loop
        edcMonth.Text = slStr
    Else
        imMonth = Val(edcMonth.Text) + 1
        If imMonth > 12 Then
            imMonth = 1
            imSettingCD = True
            imYear = sbcYear.Value + 1
            sbcYear.Value = 0
        End If
        slStr = Trim$(str$(imMonth))
        Do While Len(slStr) < 2
            slStr = "0" & slStr
        Loop
        edcMonth.Text = slStr
    End If
    sbcMonth.Value = 1
End Sub

Private Sub plcCalendar_Paint()
    plcCalendar.CurrentX = 0
    plcCalendar.CurrentY = 0
    plcCalendar.Print " Calendar"
End Sub

Private Sub sbcYear_Change()
    Dim slStr As String

    If sbcYear.Value = 1 Then
        Exit Sub
    End If
    If sbcYear.Value = 2 Then
        imYear = Val(edcYear.Text) - 1
        slStr = Trim$(str$(imYear))
        edcYear.Text = slStr
    Else
        imYear = Val(edcYear.Text) + 1
        slStr = Trim$(str$(imYear))
        edcYear.Text = slStr
    End If
    sbcYear.Value = 1
End Sub
