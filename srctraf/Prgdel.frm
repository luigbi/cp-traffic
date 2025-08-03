VERSION 5.00
Begin VB.Form PrgDel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3015
   ClientLeft      =   1065
   ClientTop       =   3135
   ClientWidth     =   6480
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
   ScaleHeight     =   3015
   ScaleWidth      =   6480
   Begin VB.PictureBox plcCalendar 
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
      Height          =   1770
      Left            =   2685
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Prgdel.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   23
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
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
         Height          =   210
         Left            =   315
         TabIndex        =   20
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3450
      TabIndex        =   18
      Top             =   2640
      Width           =   1050
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
      Left            =   60
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1965
      Width           =   120
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   3345
      TabIndex        =   0
      Top             =   15
      Width           =   3345
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Remove"
      Height          =   285
      Left            =   1845
      TabIndex        =   17
      Top             =   2640
      Width           =   1050
   End
   Begin VB.PictureBox plcDelete 
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
      Height          =   1005
      Left            =   135
      ScaleHeight     =   945
      ScaleWidth      =   6135
      TabIndex        =   1
      Top             =   345
      Width           =   6195
      Begin VB.TextBox edcDate 
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
         Index           =   1
         Left            =   3540
         MaxLength       =   10
         TabIndex        =   6
         Top             =   210
         Width           =   930
      End
      Begin VB.CommandButton cmcDate 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   4470
         Picture         =   "Prgdel.frx":2E1A
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   210
         Width           =   195
      End
      Begin VB.TextBox edcDate 
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
         Index           =   0
         Left            =   1245
         MaxLength       =   10
         TabIndex        =   3
         Top             =   210
         Width           =   930
      End
      Begin VB.CommandButton cmcDate 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   2175
         Picture         =   "Prgdel.frx":2F14
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   195
      End
      Begin VB.PictureBox plcDays 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   6030
         TabIndex        =   8
         Top             =   660
         Width           =   6030
         Begin VB.CheckBox ckcDay 
            Caption         =   "S"
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
            Index           =   6
            Left            =   5460
            TabIndex        =   15
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox ckcDay 
            Caption         =   "Sa"
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
            Index           =   5
            Left            =   4890
            TabIndex        =   14
            Top             =   0
            Width           =   570
         End
         Begin VB.CheckBox ckcDay 
            Caption         =   "Fr"
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
            Index           =   4
            Left            =   4320
            TabIndex        =   13
            Top             =   0
            Width           =   570
         End
         Begin VB.CheckBox ckcDay 
            Caption         =   "Th"
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
            Index           =   3
            Left            =   3750
            TabIndex        =   12
            Top             =   15
            Width           =   570
         End
         Begin VB.CheckBox ckcDay 
            Caption         =   "W"
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
            Index           =   2
            Left            =   3180
            TabIndex        =   11
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox ckcDay 
            Caption         =   "Tu"
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
            Index           =   1
            Left            =   2610
            TabIndex        =   10
            Top             =   0
            Width           =   570
         End
         Begin VB.CheckBox ckcDay 
            Caption         =   "Mo"
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
            Left            =   2040
            TabIndex        =   9
            Top             =   0
            Width           =   585
         End
      End
      Begin VB.Label lacTo 
         Appearance      =   0  'Flat
         Caption         =   "Delete to"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2610
         TabIndex        =   5
         Top             =   210
         Width           =   1110
      End
      Begin VB.Label lacFrom 
         Appearance      =   0  'Flat
         Caption         =   "Delete from"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   135
         TabIndex        =   2
         Top             =   210
         Width           =   1110
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   2535
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacMsg 
      Appearance      =   0  'Flat
      Caption         =   "This operation can't be undone except by adding library back into days after scheduling"
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
      Left            =   60
      TabIndex        =   16
      Top             =   1395
      Width           =   6375
   End
End
Attribute VB_Name = "PrgDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Prgdel.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CCancel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract revision number increment screen code
Option Explicit
Option Compare Text
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imWhichEdc As Integer   '0=edcFrom; 1=edcTo
Dim imBypassFocus As Integer
Dim imBSMode As Integer
Dim smScreenCaption As String
Dim imFirstActivate As Integer


Private Sub ckcDay_GotFocus(Index As Integer)
    plcCalendar.Visible = False
End Sub
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDate(imWhichEdc).SelStart = 0
    edcDate(imWhichEdc).SelLength = Len(edcDate(imWhichEdc).Text)
    edcDate(imWhichEdc).SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDate(imWhichEdc).SelStart = 0
    edcDate(imWhichEdc).SelLength = Len(edcDate(imWhichEdc).Text)
    edcDate(imWhichEdc).SetFocus
End Sub
Private Sub cmcCancel_Click()
    igRemLibReturn = 0
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcDate_Click(Index As Integer)
    imWhichEdc = Index
    plcCalendar.Visible = Not plcCalendar.Visible
    edcDate(imWhichEdc).SelStart = 0
    edcDate(imWhichEdc).SelLength = Len(edcDate(imWhichEdc).Text)
    edcDate(imWhichEdc).SetFocus
End Sub
Private Sub cmcDate_GotFocus(Index As Integer)
    If Index <> imWhichEdc Then
        plcCalendar.Visible = False
    End If
    imWhichEdc = Index
    plcCalendar.Move plcDelete.Left + edcDate(Index).Left, plcDelete.Top + edcDate(Index).Top + edcDate(Index).Height
End Sub
Private Sub cmcDone_Click()
    Dim slStr As String
    Dim ilDay As Integer
    Dim ilOneYes As Integer
    Dim ilLoop As Integer
    Dim ilRes As Integer
    If Trim$(edcDate(0).Text) = "" Then
        ilRes = MsgBox("Delete From must be specified", vbOKOnly + vbExclamation, "Incomplete")
        edcDate(0).SetFocus
        Exit Sub
    End If
    slStr = edcDate(0).Text
    If Not gValidDate(slStr) Then
        ilRes = MsgBox("Delete From must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
        edcDate(0).SetFocus
        Exit Sub
    End If
    slStr = edcDate(1).Text
    If (Trim$(slStr) <> "") And (StrComp(slStr, "TFN", 1) <> 0) Then
        If Not gValidDate(slStr) Then
            ilRes = MsgBox("Delete To must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            edcDate(1).SetFocus
            Exit Sub
        End If
    End If
    'If Selling or Airing or conventional with feed(then M-F and/or Sa and/or Su must be selected)
    If (sgVefTypeViaPrg = "S") Or (sgVefTypeViaPrg = "A") Or (sgVefTypeViaPrg = "CF") Then
        ilOneYes = False
        For ilLoop = 0 To 4 Step 1
            If ckcDay(ilLoop).Value = vbChecked Then
                ilOneYes = True
                Exit For
            End If
        Next ilLoop
        If ilOneYes Then
            For ilLoop = 0 To 4 Step 1
                If ckcDay(ilLoop).Value = vbUnchecked Then
                    ilRes = MsgBox("Mo-Fr must be specified since one is selected", vbOKOnly + vbExclamation, "Incomplete")
                    ckcDay(0).SetFocus
                    Exit Sub
                End If
            Next ilLoop
        End If
    End If
    igRemLibReturn = 1
    tgRPrg(0).sStartDate = edcDate(0).Text
    If (Trim$(edcDate(1).Text) = "") Or (StrComp(edcDate(1).Text, "TFN", 1) = 0) Then
        tgRPrg(0).sEndDate = ""
    Else
        tgRPrg(0).sEndDate = edcDate(1).Text
    End If
    For ilDay = 0 To 6 Step 1
        If ckcDay(ilDay).Value = vbChecked Then
            tgRPrg(0).iDay(ilDay) = 1
        Else
            tgRPrg(0).iDay(ilDay) = 0
        End If
    Next ilDay
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub edcDate_Change(Index As Integer)
    Dim slStr As String
    slStr = edcDate(Index).Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    lacDate.Visible = True
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcDate_GotFocus(Index As Integer)
    imWhichEdc = Index
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
    plcCalendar.Move plcDelete.Left + edcDate(Index).Left, plcDelete.Top + edcDate(Index).Top + edcDate(Index).Height
End Sub
Private Sub edcDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDate(Index).SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    If Index = 0 Then
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Else
        If (Len(edcDate(Index).Text) = edcDate(Index).SelLength) And (igViewType = 0) And (igLibType = 0) Then
            If (KeyAscii = Asc("T")) Or (KeyAscii = Asc("t")) Then
                edcDate(Index).Text = "TFN"
                edcDate(Index).SelStart = 0
                edcDate(Index).SelLength = 3
                KeyAscii = 0
                Exit Sub
            End If
        End If
        'Filter characters (allow only BackSpace, numbers 0 thru 9
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Private Sub edcDate_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If Index = 0 Then
            If (Shift And vbAltMask) > 0 Then
                plcCalendar.Visible = Not plcCalendar.Visible
            Else
                slDate = edcDate(Index).Text
                If gValidDate(slDate) Then
                    If KeyCode = KEYUP Then 'Up arrow
                        slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                    Else
                        slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                    End If
                    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                    edcDate(Index).Text = slDate
                End If
            End If
        Else
            If (Shift And vbAltMask) > 0 Then
                plcCalendar.Visible = Not plcCalendar.Visible
            Else
                slDate = edcDate(Index).Text
                If gValidDate(slDate) Then
                    If KeyCode = KEYUP Then 'Up arrow
                        slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                    Else
                        slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                    End If
                    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                    edcDate(Index).Text = slDate
                End If
            End If
        End If
    End If
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        gFunctionKeyBranch KeyCode
    End If

End Sub

Private Sub Form_Load()
    mInit
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = edcDate(imWhichEdc).Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub
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
    Dim ilLoop As Integer
    Dim slStr As String
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imFirstActivate = True
    imBypassFocus = False
    smScreenCaption = "Removing Library- " & Trim$(sgRemLibName) & ":" & tgRPrg(0).sStartTime
    imCalType = 0   'Standard
    imBSMode = False
    mInitBox
    edcDate(0).Text = tgRPrg(0).sStartDate
    If Trim$(tgRPrg(0).sEndDate) = "" Then
        edcDate(1).Text = "TFN"
        slStr = tgRPrg(0).sStartDate
        gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
        pbcCalendar_Paint
    Else
        edcDate(1).Text = tgRPrg(0).sEndDate
    End If
    For ilLoop = 0 To 6 Step 1
        If tgRPrg(0).iDay(ilLoop) = 1 Then
            ckcDay(ilLoop).Value = vbChecked
        Else
            ckcDay(ilLoop).Value = vbUnchecked
        End If
    Next ilLoop
    PrgDel.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone PrgDel
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
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
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
    Unload PrgDel
    Set PrgDel = Nothing   'Remove data segment
End Sub
Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcDate(imWhichEdc).Text = Format$(llDate, "m/d/yy")
                edcDate(imWhichEdc).SelStart = 0
                edcDate(imWhichEdc).SelLength = Len(edcDate(imWhichEdc).Text)
                imBypassFocus = True
                edcDate(imWhichEdc).SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcDate(imWhichEdc).SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub
Private Sub plcDays_Paint()
    plcDays.CurrentX = 0
    plcDays.CurrentY = 0
    plcDays.Print "Days to Remove Library"
End Sub
