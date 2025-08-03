VERSION 5.00
Begin VB.Form PostManualVeh 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4545
   ClientLeft      =   1170
   ClientTop       =   2115
   ClientWidth     =   9150
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4545
   ScaleWidth      =   9150
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7395
      Top             =   4215
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   2895
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "PostManualVeh.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   13
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   315
         TabIndex        =   10
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3585
      Width           =   75
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
      Left            =   4830
      TabIndex        =   7
      Top             =   4230
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Post Counts"
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
      Left            =   2970
      TabIndex        =   6
      Top             =   4230
      Width           =   1290
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
      ScaleWidth      =   4680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   4680
   End
   Begin VB.PictureBox plcDates 
      ForeColor       =   &H00000000&
      Height          =   3720
      Left            =   180
      ScaleHeight     =   3660
      ScaleWidth      =   8685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   8745
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         ItemData        =   "PostManualVeh.frx":2E1A
         Left            =   4935
         List            =   "PostManualVeh.frx":2E1C
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   285
         Width           =   3645
      End
      Begin VB.TextBox edcDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   2685
         MaxLength       =   10
         TabIndex        =   3
         Top             =   300
         Width           =   1455
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
         Left            =   4155
         Picture         =   "PostManualVeh.frx":2E1E
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Width           =   195
      End
      Begin VB.Label lacStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Start Date of Invoice Month"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   2475
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   345
      Top             =   4125
      Width           =   360
   End
End
Attribute VB_Name = "PostManualVeh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PostManualVeh.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PostManualVeh.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Salesperson Model library dates input screen code
Option Explicit
Option Compare Text

'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imBypassFocus As Integer
Dim lmNowDate As Long
Dim smPassCommands As String
Dim bmAutoMode As Boolean

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer


'Dim hmIihf As Integer
'Dim tmIihf As IIHF        'CFF record image
'Dim tmIihfSrchKey0 As LONGKEY0    'CFF key record image
'Dim tmIihfSrchKey1 As IIHFKEY1    'CFF key record image
'Dim tmIihfSrchKey2 As IIHFKEY2    'CFF key record image
'Dim tmIihfSrchKey3 As IIHFKEY3    'CFF key record image
'Dim imIihfRecLen As Integer        'CFF record length
'
'Dim hmIidf As Integer
'Dim tmIidf As IIDF        'CFF record image
'Dim tmIidfSrchKey0 As LONGKEY0    'CFF key record image
'Dim tmIidfSrchKey1 As LONGKEY0    'CFF key record image
'Dim tmIidfSrchKey2 As LONGKEY0    'CFF key record image
'Dim imIidfRecLen As Integer

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDate.SelStart = 0
    edcDate.SelLength = Len(edcDate.Text)
    edcDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDate.SelStart = 0
    edcDate.SelLength = Len(edcDate.Text)
    edcDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    igManualPostVefCode = 0
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcDate.SelStart = 0
    edcDate.SelLength = Len(edcDate.Text)
    edcDate.SetFocus
End Sub
Private Sub cmcDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    Dim ilRet As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim slLastDate As String
    slDate = Trim$(edcDate.Text)
    If slDate = "" Then
        ilRet = MsgBox("Start month must be specified", vbOKOnly + vbExclamation, "Incomplete")
        edcDate.SetFocus
        Exit Sub
    Else
        If Not gValidDate(slDate) Then
            ilRet = MsgBox("Start month must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            edcDate.SetFocus
            Exit Sub
        End If
    End If
    slDate = gObtainStartStd(slDate)
    llDate = gDateValue(gObtainEndStd(slDate))
    If llDate >= lmNowDate Then
        slDate = gObtainStartStd(DateAdd("d", -1, slDate))
        ilRet = MsgBox("Start month must be on or prior to " & Format(slDate, "m/d/yy"), vbOKOnly + vbExclamation, "Incomplete")
        edcDate.SetFocus
        Exit Sub
    End If
    edcDate.Text = slDate
    sgManualPostDate = slDate
    igManualPostVefCode = -1
    Screen.MousePointer = vbHourglass
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            igManualPostVefCode = lbcVehicle.ItemData(ilLoop)
            Exit For
        End If
    Next ilLoop
    Screen.MousePointer = vbDefault
    If igSportsSystem = 3 Then
        gAddOrUpdateLrf "A", igManualPostVefCode, sgManualPostDate
        PostManualCntr.Show vbModal
        gAddOrUpdateLrf "U", igManualPostVefCode, sgManualPostDate
    ElseIf igManualPostVefCode <= 0 Then
        mTerminate
    Else
        PostManualCntr.Show vbModal
        plcDates.Visible = False
        DoEvents
        plcDates.Visible = True
    End If
End Sub
Private Sub cmcDone_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    plcCalendar.Visible = False
End Sub
Private Sub edcDate_Change()
    Dim slStr As String
    slStr = edcDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    lacDate.Visible = True
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcDate_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcDate.Text = slDate
            End If
        End If
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcDate.Text = slDate
            End If
        End If
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
    End If
End Sub


Private Sub edcDate_LostFocus()
    Dim slDate As String
    slDate = edcDate.Text
    If Not gValidDate(slDate) Then
        Exit Sub
    End If
    slDate = gObtainStartStd(slDate)
    edcDate.Text = slDate
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
    If imTerminate Then
        tmcStart.Enabled = True
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    ilRet = btrClose(hgAuf)
    btrDestroy hgAuf
    ilRet = btrClose(hgUlf)
    btrDestroy hgUlf
    gCloseVpf
    '4/2/11: Add setting and call.
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    
    'ilRet = btrClose(hmIidf)
    'btrDestroy hmIidf
    'ilRet = btrClose(hmIihf)
    'btrDestroy hmIihf
    Set PostManualVeh = Nothing   'Remove data segment
    igManUnload = NO
    gEraseGlobalVar True
    'Dan M 9/21/10 delete text file passing the date from traffic--just in case.
    mEraseOrChangeDateFile
    End
End Sub

Private Sub imcHelp_Click()
    plcCalendar.Visible = False
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcVehicle_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
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
    slStr = edcDate.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(Str$(Day(llDate)))
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
    Dim slStr As String
    Dim ilRet As Integer
    Dim slDate As String
    Screen.MousePointer = vbHourglass
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    imLBCDCtrls = 1
    imFirstActivate = True
    imTerminate = False
    igModReturn = 0
    imBypassFocus = False
    imSettingValue = False
    imChgMode = False
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox

    'imIidfRecLen = Len(tmIidf)
    'hmIidf = CBtrvTable(TWOHANDLES)
    'ilRet = btrOpen(hmIidf, "", sgDBPath & "Iidf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'On Error GoTo mInitErr
    'gBtrvErrorMsg ilRet, "mInit (btrOpen: Iidf.Btr)", PostManualVeh
    'On Error GoTo 0

    'imIihfRecLen = Len(tmIihf)
    'hmIihf = CBtrvTable(TWOHANDLES)
    'ilRet = btrOpen(hmIihf, "", sgDBPath & "Iihf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'On Error GoTo mInitErr
    'gBtrvErrorMsg ilRet, "mInit (btrOpen: Iihf.Btr)", PostManualVeh
    'On Error GoTo 0


    PostManualVeh.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone PostManualVeh
    plcCalendar.Move plcDates.Left + edcDate.Left, plcDates.Top + edcDate.Top + edcDate.Height
    slStr = Format$(gNow(), "m/d/yy")
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slDate
    slStr = gIncOneDay(slDate)
    slStr = gObtainStartStd(slStr)    'Get tuesday of the current week
    edcDate.Text = slStr
    ilRet = gObtainAdvt()
    ilRet = gObtainStations()
    ilRet = mVefPop()

    'gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    'pbcCalendar_Paint   'mBoxCalDate called within paint
    'lacDate.Visible = False
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    slDate = gObtainEndStd(slDate)
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
    Unload PostManualVeh
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVefPop                         *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mVefPop() As Integer
    Dim ilRet As Integer
    Dim ilShtt As Integer
    Dim ilVef As Integer
    Dim ilVff As Integer
    Dim slDate As String
'    ilRet = gPopUserVehicleBox(PostManualVeh, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH + DORMANTVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
'    If ilRet <> CP_MSG_NOPOPREQ Then
'        On Error GoTo mVehPopErr
'        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", PostManualVeh
'        On Error GoTo 0
'    End If
    mVefPop = True
    lbcVehicle.Clear
    slDate = edcDate.Text
    If Not gValidDate(slDate) Then
        Exit Function
    End If
    slDate = gObtainStartStd(slDate)
    For ilShtt = LBound(tgVehicleStations) To UBound(tgVehicleStations) - 1 Step 1
        If tgVehicleStations(ilShtt).iShttVefCode > 0 Then
            ilVef = gBinarySearchVef(tgVehicleStations(ilShtt).iShttVefCode)
            If ilVef <> -1 Then
                If gIsImportInvoicedSpots(tgVehicleStations(ilShtt).iShttVefCode) Then
                    'tmIihfSrchKey1.iVefCode = tgVehicleStations(ilShtt).iShttVefCode
                    'gPackDate slDate, tmIihfSrchKey1.iInvStartDate(0), tmIihfSrchKey1.iInvStartDate(1)
                    'ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    'Add to list
                    '    lbcVehicle.AddItem (tgMVef(ilVef).sName)
                    '    lbcVehicle.ItemData(lbcVehicle.NewIndex) = tgMVef(ilVef).iCode
                    'Else
                    '    If (Trim$(tmIihf.sSourceForm) = "T") Or (Trim$(tmIihf.sSourceForm) = "C") Then
                    If igSportsSystem = 3 Then
                        If UCase$(Trim$(tgMVef(ilVef).sName)) = sgStationCallLetters Then
                            lbcVehicle.AddItem (tgMVef(ilVef).sName)
                            lbcVehicle.ItemData(lbcVehicle.NewIndex) = tgMVef(ilVef).iCode
                        Else
                            ilVff = gBinarySearchVff(tgMVef(ilVef).iCode)
                            If ilVff <> -1 Then
                                If UCase$(Trim$(tgVff(ilVff).sStationPassword)) = sgStationPW Then
                                    lbcVehicle.AddItem (tgMVef(ilVef).sName)
                                    lbcVehicle.ItemData(lbcVehicle.NewIndex) = tgMVef(ilVef).iCode
                                End If
                            End If
                        End If
                    Else
                        lbcVehicle.AddItem (tgMVef(ilVef).sName)
                        lbcVehicle.ItemData(lbcVehicle.NewIndex) = tgMVef(ilVef).iCode
                    End If
                    '    End If
                    'End If
                End If
            End If
        End If
    Next ilShtt
    If lbcVehicle.ListCount = 1 Then
        lbcVehicle.ListIndex = 0
    End If
    If (igSportsSystem = 3) And (lbcVehicle.ListCount > 1) Then
        For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
            If UCase$(Trim$(lbcVehicle.List(ilVef))) = sgStationCallLetters Then
                lbcVehicle.ListIndex = ilVef
                Exit For
            End If
        Next ilVef
    End If
    Exit Function
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    Dim slDate As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(Str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                slDate = Format$(llDate, "m/d/yy")
                slDate = gObtainStartStd(slDate)
                edcDate.Text = slDate
                edcDate.SelStart = 0
                edcDate.SelLength = Len(edcDate.Text)
                imBypassFocus = True
                edcDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(Str$(imCalMonth)) & "/15/" & Trim$(Str$(imCalYear))
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
Private Sub plcDates_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Manual Posting"
End Sub

Private Sub tmcStart_Timer()
    
    tmcStart.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
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
Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slStartIn As String
    Dim ilPos As Integer
    Dim ilLen As Integer
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer

    sgCommandStr = Command$
    slStartIn = CurDir$
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    igShowVersionNo = 0
    If (InStr(1, slStartIn, "Prod", vbTextCompare) = 0) And (InStr(1, slStartIn, "Test", vbTextCompare) = 0) Then
        igShowVersionNo = 1
        If InStr(1, sgCommandStr, "Debug", vbTextCompare) > 0 Then
            igShowVersionNo = 2
        End If
    End If
    slCommand = sgCommandStr    'Command$
    smPassCommands = slCommand
    lgCurrHRes = GetDeviceCaps(Traffic!pbcList.hdc, HORZRES)
    lgCurrVRes = GetDeviceCaps(Traffic!pbcList.hdc, VERTRES)
    lgCurrBPP = GetDeviceCaps(Traffic!pbcList.hdc, BITSPIXEL)
    mTestPervasive
    
    DoEvents                    'try to prevent open 3012 error on auf
    'hgAuf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    hgAuf = CBtrvTable(ONEHANDLE) 'CBtrvObj()

    ilRet = btrOpen(hgAuf, "", sgDBPath & "AUF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "Unable to Open Alert File, Error = " & Str$(ilRet), vbOKOnly + vbInformation, "Warning"
    End If
    hgUlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hgUlf, "", sgDBPath & "ULF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "Unable to Open User Log File, Error = " & Str$(ilRet), vbOKOnly + vbInformation, "Warning"
    End If
    gOpenVpf
    
    '4/2/11: Add setting of value
    lgUlfCode = 0
    
    'examples of auto export command line
    'Auto-Matrix Section-Matrix Std TNet
    'Auto-Efficio Revenue Section-Efficio Cal TNet
    If (Trim$(sgCommandStr) = "") Or (Trim$(sgCommandStr) = "/UserInput") Then
        bmAutoMode = False
        Signon.Show vbModal
        If igExitTraffic Then
            imTerminate = True
            Exit Sub
        End If
        slStr = sgUserName
        'dan M 12/01/10 redundant
        'gCallNetReporter CsiReportCall.StartReports
        'Sleep 1000
        sgCallAppName = "Traffic"
        igRptCallType = -2
        If Not igTestSystem Then
            smPassCommands = "Traffic^Prod\" & slStr & "\-2\0"
        Else
            smPassCommands = "Traffic^Test\" & slStr & "\-2\0"
        End If
    Else
        igBkgdProg = 0
        bmAutoMode = False
        igSportsSystem = 0
        If InStr(1, sgCommandStr, "FROM/SPORTSSYSTEM") > 0 Then
            igSportsSystem = 3
            ilPos1 = InStr(1, sgCommandStr, "SCL/")
            ilPos2 = InStr(1, sgCommandStr, "RM/")
            If ilPos2 > 0 Then
                ilPos2 = ilPos2 + 3
                sgRemoteUserName = Mid$(sgCommandStr, ilPos2, ilPos1 - ilPos2)
            End If
            ilPos1 = InStr(1, sgCommandStr, "SPW")
            ilPos2 = InStr(1, sgCommandStr, "SCL/")
            If ilPos2 > 0 Then
                ilPos2 = ilPos2 + 4
                sgStationCallLetters = UCase$(Trim$(Mid$(sgCommandStr, ilPos2, ilPos1 - ilPos2)))
            End If
            ilPos1 = InStr(1, sgCommandStr, "FROM/SPORTSSYSTEM")
            ilPos2 = InStr(1, sgCommandStr, "SPW/")
            If ilPos2 > 0 Then
                ilPos2 = ilPos2 + 4
                sgStationPW = UCase$(Trim$(Mid$(sgCommandStr, ilPos2, ilPos1 - ilPos2)))
            End If
        End If
        sgCallAppName = "Traffic"
        igRptCallType = 0
        sgUrfStamp = "~" 'Clear time stamp incase same name
        If igSportsSystem = 3 Then
            sgUserName = sgStationCallLetters
        Else
            sgUserName = "Guide"
        End If
        gUrfRead Signon, sgUserName, True, tgUrf(), False  'Obtain user records
        mGetUlfCode
        DoEvents
        gInitStdAlone
        mCheckForDate
        '4/2/11: Add setting and call.  Note: The call in _Load will be ignored
        ilRet = gObtainSAF()
        igLogActivityStatus = 32123
        gUserActivityLog "L", "PostManualVeh.Frm"
        Exit Sub
    End If
    'End If
    DoEvents
'    gInitStdAlone ResearchList, slStr, igTestSystem
    gInitStdAlone
    mCheckForDate
    '4/2/11: Add setting and call.  Note: The call in _Load will be ignored
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "PostManualVeh.Frm"
    'ilRet = gParseItem(slCommand, 3, "\", slStr)
    'igRptCallType = Val(slStr)
    
    'debug
    'MsgBox "igExporttype = " & str$(igExportType) & " smExportName = " & smExportName & " sgExportIniSectionName=" & sgExportIniSectionName, vbOKOnly
End Sub
Private Sub mEraseOrChangeDateFile()
    Dim oMyFileObj As FileSystemObject
    Dim MyFile As TextStream
    Dim slFullPath As String
    
        '5676 remove hard coded c:
    'slFullPath = "C:\csi\ReportPasser.txt"
    slFullPath = sgRootDrive & "csi\ReportPasser.txt"
    Set oMyFileObj = New FileSystemObject
    If oMyFileObj.FILEEXISTS(slFullPath) Then
On Error GoTo FIXFILE
        oMyFileObj.DeleteFile slFullPath, True
    End If
Cleanup:
    Set oMyFileObj = Nothing
    Exit Sub
FIXFILE:
    'couldn't delete file...probably open, so erase the value
    If oMyFileObj.FILEEXISTS(slFullPath) Then
        Set MyFile = oMyFileObj.OpenTextFile(slFullPath, ForWriting, False)
        MyFile.WriteLine ("")
        MyFile.Close
        Set MyFile = Nothing
        GoTo Cleanup
    End If

End Sub

Private Sub mTestPervasive()
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim hlSpf As Integer
    Dim tlSpf As SPF

    gInitGlobalVar
    hlSpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    ilRecLen = Len(tlSpf)
    ilRet = btrGetFirst(hlSpf, tlSpf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    btrDestroy hlSpf
End Sub
Private Sub mCheckForDate()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slDate As String
    Dim slSetDate As String
    Dim ilRet As Integer
    
    ilPos = InStr(1, sgCommandStr, "/D:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3))
        Else
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gValidDate(slDate) Then
            slDate = gAdjYear(slDate)
            slSetDate = slDate
        End If
    End If
    If Trim$(slSetDate) = "" Then
        If (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0) Or (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) > 0) Then
            slSetDate = "12/15/2015"
            slDate = slSetDate
        End If
    End If
    If Trim$(slSetDate) <> "" Then
        'Dan M 9/20/10 problems with gGetCSIName("SYSDate") in v57 reports.exe... change to global variable
     '   ilRet = csiSetName("SYSDate", slDate)
        ilRet = gCsiSetName(slDate)
    End If
    
End Sub
Private Sub mGetUlfCode()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    ilPos = InStr(1, sgCommandStr, "/ULF:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5)))
        Else
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5, ilSpace - ilPos - 3)))
        End If
    End If
End Sub
