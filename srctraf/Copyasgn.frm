VERSION 5.00
Begin VB.Form CopyAsgn 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5490
   ClientLeft      =   1770
   ClientTop       =   1320
   ClientWidth     =   5985
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
   ScaleHeight     =   5490
   ScaleWidth      =   5985
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
      Left            =   420
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2550
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
      Left            =   270
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2505
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
      Left            =   435
      TabIndex        =   21
      Top             =   4590
      Visible         =   0   'False
      Width           =   255
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
      Left            =   30
      ScaleHeight     =   165
      ScaleWidth      =   105
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4740
      Width           =   105
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   1290
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "Copyasgn.frx":0000
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "Copyasgn.frx":0CBE
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image imcTmeOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
      End
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
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   15
      Top             =   1485
      Width           =   60
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
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   15
      TabIndex        =   3
      Top             =   240
      Width           =   15
   End
   Begin VB.PictureBox pbcSelections 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   840
      ScaleHeight     =   210
      ScaleWidth      =   1395
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   420
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1020
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
      Left            =   1440
      Picture         =   "Copyasgn.frx":0FC8
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   195
   End
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
      Left            =   3540
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3090
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Copyasgn.frx":10C2
         ScaleHeight     =   1410
         ScaleWidth      =   1830
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   255
         Width           =   1860
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
            TabIndex        =   14
            Top             =   405
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
         Height          =   210
         Left            =   45
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   45
         Width           =   255
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
         Height          =   210
         Left            =   1635
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   45
         Width           =   255
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
         Height          =   195
         Left            =   315
         TabIndex        =   11
         Top             =   45
         Width           =   1275
      End
   End
   Begin VB.PictureBox pbcAsgn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1065
      Left            =   180
      Picture         =   "Copyasgn.frx":3EDC
      ScaleHeight     =   1065
      ScaleWidth      =   5670
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3165
      Width           =   5670
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1140
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1140
   End
   Begin VB.CommandButton cmcGenerate 
      Appearance      =   0  'Flat
      Caption         =   "&Assign"
      Height          =   285
      Left            =   1830
      TabIndex        =   16
      Top             =   5100
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3345
      TabIndex        =   17
      Top             =   5100
      Width           =   945
   End
   Begin VB.PictureBox plcAsgn 
      ForeColor       =   &H00000000&
      Height          =   1185
      Left            =   120
      ScaleHeight     =   1125
      ScaleWidth      =   5715
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3105
      Width           =   5775
   End
   Begin VB.PictureBox plcVehicle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   495
      ScaleHeight     =   2655
      ScaleWidth      =   4995
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   5055
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   165
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   390
         Width           =   4755
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   165
      Top             =   4995
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "CopyAsgn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Copyasgn.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'************************************************************
' File Name: CopyAsgn.Frm
'
' Release: 1.0
'               Created: ?          By: D. LeVine
'
' Description:
'   This file contains the Copy assign input screen code
'************************************************************
Option Explicit
Option Compare Text
'Btrieve files
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer      'VEF record length
Dim hmVef As Integer            'Vehicle file handle
Dim tmVpf As VPF                'VPF record image
Dim tmVpfSrchKey As VPFKEY0     'VPF key 0 image
Dim imVpfRecLen As Integer      'VPF record length
Dim hmVpf As Integer            'Vehicle Option file handle
Dim tmLcf As LCF                'LCF record image
Dim tmLcfSrchKey As LCFKEY0     'LCF key 0 image
Dim imLcfRecLen As Integer      'LCF record length
Dim hmLcf As Integer            'Log Calendar file handle
'CopyAsgn flags
Dim imFirstActivate As Integer
Dim imTerminate As Integer      'True=terminate  False = OK
Dim imChgMode As Integer        'True=value changed
'CopyAsgn modular variables
Dim imNoSelected As Integer
Dim imPFAllowed As Integer    'True=Preliminary/final allowed; False=only final
Dim imVehCode As Integer        'vehicle code
Dim smDate As String            'now +1
Dim smNowDate As String
Dim smDefaultDate As String     'default start date
Dim smDefaultTime As String     'default start time
Dim smSave(0 To 4) As String    'edc save strings (Index: 1=Start Date; 2=# of days; 3=Start Time; 4=End Time)
Dim imSave(0 To 2) As Integer   '1=On Air (0) or Alternate(1); 2=Prelimary(0) or Final(1)
Dim imLbcArrowSetting As Integer 'True = Processing arrow key- retain current list box visible
                                'False= Make list box invisible
Dim imFirstFocus As Integer
Dim imVefSelected() As Integer  'Indicator if vehicle selected- used to test that
                                'only vehicles with same date are selected
'Calendar variables
Dim tmCDCtrls(0 To 7) As FIELDAREA  'Field area image
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer        'Month of displayed calendar
Dim imCalMonth As Integer       'Year of displayed calendar
Dim lmCalStartDate As Long      'Start date of displayed calendar
Dim lmCalEndDate As Long        'End date of displayed calendar
Dim imCalType As Integer        'Calendar type
Dim imBSMode As Integer         'Backspace flag
Dim imBypassFocus As Integer    'Bypass gotfocus
'Tabs
Dim tmCtrls(0 To 6)  As FIELDAREA   'Field area image
Dim imLBCtrls As Integer
Dim imBoxNo As Integer          'Current Media Box
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)

Dim imUpdateAllowed As Integer    'User can update records

Dim bmFirstCallToVpfFind As Boolean

'Constants
Const SDATEINDEX = 1            'Start Date control/field
Const NODAYSINDEX = 2           'Number of days control/field
Const STIMEINDEX = 3            'Start time control/index
Const ENDTIMEINDEX = 4          'End time control/field
Const ONAIRINDEX = 5            'On air control/index
Const PFCOPYINDEX = 6            'Final or preliminary copy control/index
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1

End Sub
Private Sub cmcDropDown_Click()

    Select Case imBoxNo
        Case SDATEINDEX  'Start date
            plcCalendar.Visible = Not plcCalendar.Visible
        Case NODAYSINDEX 'End date
            plcCalendar.Visible = Not plcCalendar.Visible
        Case STIMEINDEX   'Start time
            plcTme.Visible = Not plcTme.Visible
        Case ENDTIMEINDEX  'End Time
            plcTme.Visible = Not plcTme.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcGenerate_Click()
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVehCode As Integer
    Dim ilVehIndex As Integer
    Dim slEndDate As String
    Dim ilRet As Integer
    Dim ilType As Integer
    Dim ilPos As Integer
    Dim slStr As String
    Dim slNextDate As String
    Dim ilUpdateDate As Integer
    Dim slDays As String
    Dim slCopyDate As String
    Dim llDate As Long
    Dim clTime1 As Currency     'Start time value
    Dim clTime2 As Currency     'End Time Value
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            slNameCode = tgCAVehicle(ilLoop).sKey    'lbcVehCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVehCode = Val(slCode)
            slStr = lbcVehicle.List(ilLoop)
            ilPos = InStr(slStr, ": ")
            If ilPos > 0 Then
                slStr = Mid$(slStr, ilPos + 2)
                ilPos = InStr(slStr, "(N")
                If ilPos > 0 Then
                    slCopyDate = Left$(slStr, ilPos - 1)
                    If smSave(1) = "" Then
                        smSave(1) = slCopyDate
                    End If
                    Exit For
                Else
                    slCopyDate = smSave(1)
                    'smSave(1) = ""
                End If
            Else
                slCopyDate = smSave(1)
                'smSave(1) = ""
            End If
        End If
    Next ilLoop
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        If mTestFields(ilLoop, ALLMANDEFINED + SHOWMSG) = NO Then
            Beep
            imBoxNo = ilLoop
            mEnableBox imBoxNo
            Exit Sub
        End If
    Next ilLoop
    'If (gDateValue(smSave(1)) <= gDateValue(smNowDate)) Then
    '    Beep
    '    ilRet = MsgBox("Date specified must be after Today's Date", vbOkOnly + vbExclamation, "Date Check")
    '    imBoxNo = SDATEINDEX
    '    mEnableBox imBoxNo
    '    Exit Sub
    'End If
    If (gDateValue(smNowDate) + 30) < gDateValue(smSave(1)) Then
        Beep
        slDays = Trim$(str$(gDateValue(smSave(1)) - gDateValue(smNowDate)))
        ilRet = MsgBox("Start Date is " & slDays & " days after today's date, Ok to Continue with Assigning Copy", vbYesNo + vbQuestion + vbDefaultButton2, "Date Check")
        If ilRet <> vbYes Then
            imBoxNo = SDATEINDEX
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    If (gDateValue(smSave(1)) < gDateValue(smNowDate)) Then
        Beep
        ilRet = MsgBox("Date specified prior to Today's Date, Ok to Continue with Assigning Copy", vbYesNo + vbQuestion + vbDefaultButton2, "Date Check")
        If ilRet <> vbYes Then
            imBoxNo = SDATEINDEX
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    'If imNoSelected = 1 Then
    '    slNameCode = lbcVehCode.List(lbcVehicle.ListIndex)
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '    ilVehCode = Val(slCode)
    '    ilVehIndex = gVpfFind(CopyAsgn, ilVehCode)
    '    gUnpackDate tgVpf(ilVehIndex).iLLastDateCpyAsgn(0), tgVpf(ilVehIndex).iLLastDateCpyAsgn(1), slCopyDate
    '    If slCopyDate <> "" Then
    '        slCopyDate = gIncOneDay(slCopyDate)
            'Determine next valid date to assign copy
    '        tmLcfSrchKey.sType = "O"
    '        tmLcfSrchKey.sStatus = "C"
    '        tmLcfSrchKey.iVefCode = ilVehCode
    '        gPackDate slCopyDate, tmLcfSrchKey.iLogDate(0), tmLcfSrchKey.iLogDate(1)
    '        tmLcfSrchKey.iSeqNo = 0
    '        ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    '        If (ilRet = BTRV_ERR_NONE) And (tmLcf.sType = "O") And (tmLcf.sStatus = "C") And (tmLcf.iVefCode = ilVehCode) Then
    '            gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slCopyDate
    '        Else
    '            'No dates in the future
    '            'gUnpackDate tgVpf(ilVehIndex).iLLastDateCpyAsgn(0), tgVpf(ilVehIndex).iLLastDateCpyAsgn(1), slCopyDate
    '        End If
            If gDateValue(smSave(1)) < gDateValue(slCopyDate) Then
                'Assigning to dates in past
                Beep
                ilRet = MsgBox(smSave(1) & " has been assigned copy previously, this will assign new and revised copy only. Ok to Continue with Assigning Copy", vbYesNo + vbQuestion + vbDefaultButton2, "Date Check")
                If ilRet <> vbYes Then
                    imBoxNo = SDATEINDEX
                    mEnableBox imBoxNo
                    Exit Sub
                End If
            ElseIf gDateValue(smSave(1)) > gDateValue(slCopyDate) Then
                'Assigning to dates in the future
                Beep
                ilRet = MsgBox("Assigning to " & smSave(1) & " will cause the agencies copy pattern to be assigned out of order. Ok to Continue with Assigning Copy", vbYesNo + vbQuestion + vbDefaultButton2, "Date Check")
                If ilRet <> vbYes Then
                    imBoxNo = SDATEINDEX
                    mEnableBox imBoxNo
                    Exit Sub
                End If
            End If
        'End If
    'End If
    clTime1 = gTimeToCurrency(smSave(3), False)
    clTime2 = gTimeToCurrency(smSave(4), True)
    If (clTime2 < clTime1) Then
        Beep
        MsgBox "End Time earlier than Start Time.", vbExclamation, "Invalid Time"
        imBoxNo = ENDTIMEINDEX
        mEnableBox imBoxNo
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If (Not imPFAllowed) And (imSave(2) = -1) Then
        imSave(2) = 1
    End If
    If imSave(1) = 1 Then
        ilType = 1
    Else
        ilType = 0
    End If
    'If imNoSelected > 1 Then
    '    smSave(3) = "12M"
    '    smSave(4) = "12M"
    'End If
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            slNameCode = tgCAVehicle(ilLoop).sKey    'lbcVehCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVehCode = Val(slCode)
            If imNoSelected > 1 Then
                slStr = lbcVehicle.List(ilLoop)
                ilPos = InStr(slStr, ": ")
                If ilPos > 0 Then
                    slStr = Mid$(slStr, ilPos + 2)
                    ilPos = InStr(slStr, "(N")
                    If ilPos > 0 Then
                        slNextDate = Left$(slStr, ilPos - 1)
                    Else
                        slNextDate = ""
                    End If
                Else
                    slNextDate = ""
                End If
            Else
                slStr = lbcVehicle.List(ilLoop)
                ilPos = InStr(slStr, ": ")
                If ilPos > 0 Then
                    slStr = Mid$(slStr, ilPos + 2)
                    ilPos = InStr(slStr, "(N")
                    If ilPos > 0 Then
                        slNextDate = Left$(slStr, ilPos - 1)
                    Else
                        slNextDate = ""
                    End If
                Else
                    slNextDate = ""
                End If
            End If
            If smSave(1) <> "" Then
                slEndDate = Format$(gDateValue(smSave(1)) + Val(smSave(2)) - 1, "m/d/yy")
                ilRet = gAssignCopyToSpots(ilType, ilVehCode, imSave(2), smSave(1), slEndDate, smSave(3), smSave(4))

                If ilRet Then
                    ilUpdateDate = True 'Update for any date last long as it is after current last copy assign date
                                        'and 12m-12m selected
                Else
                    ilUpdateDate = False
                End If
                'ilUpdateDate = False
                'If slNextDate <> "" Then
                '    If gDateValue(smSave(1)) = gDateValue(slNextDate) Then
                '        ilUpdateDate = True
                '    End If
                'Else    'No date defined
                '    ilUpdateDate = True
                'End If
                'If not 12m-12m don't update dates
                If (gTimeToCurrency(smSave(3), False) <> gTimeToCurrency("12AM", False)) Or (gTimeToCurrency(smSave(4), False) <> gTimeToCurrency("12AM", False)) Then
                    ilUpdateDate = False
                End If
                If ilUpdateDate Then
                    If bmFirstCallToVpfFind Then
                        ilVehIndex = gVpfFind(CopyAsgn, ilVehCode)
                        bmFirstCallToVpfFind = False
                    Else
                        ilVehIndex = gVpfFindIndex(ilVehCode)
                    End If
                    Do
                        tmVpfSrchKey.iVefKCode = ilVehCode
                        ilRet = btrGetEqual(hmVpf, tgVpf(ilVehIndex), imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit Do
                        End If
                        gUnpackDateLong tgVpf(ilVehIndex).iLLastDateCpyAsgn(0), tgVpf(ilVehIndex).iLLastDateCpyAsgn(1), llDate
                        If gDateValue(slEndDate) <= llDate Then
                            Exit Do
                        End If
                        gPackDate slEndDate, tgVpf(ilVehIndex).iLLastDateCpyAsgn(0), tgVpf(ilVehIndex).iLLastDateCpyAsgn(1)
                        ilRet = btrUpdate(hmVpf, tgVpf(ilVehIndex), imVpfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                End If
            End If
        End If
    Next ilLoop
    '11/26/17
    gFileChgdUpdate "vpf.btr", False
    lbcVehicle.Clear    'Initialize List boxes
    'lbcVehCode.Clear
    sgCAVehicleTag = ""
    ReDim tgCAVehicle(0 To 0) As SORTCODE
    mPopulate           'Populate vehicle list boxes
    cmcCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmcGenerate_GotFocus()
    mSetShow imBoxNo  'Process last field if user moused here
    imBoxNo = -1
    gCtrlGotFocus ActiveControl
    'For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
    '    If mTestFields(ilLoop, ALLMANDEFINED + SHOWMSG) = NO Then
    '        Beep
    '        imBoxNo = ilLoop
    '        mEnableBox imBoxNo
    '        Exit Sub
    '    End If
    'Next ilLoop
    'If imNoSelected = 1 Then
    '    If Not gValidDate(smSave(1)) Then
    '        Beep
    '        imBoxNo = SDATEINDEX
    '        mEnableBox imBoxNo
    '        Exit Sub
    '    End If
    '    If (gDateValue(smSave(1)) < gDateValue(smNowDate)) Then
    '        Beep
    '        MsgBox "Date must be on or after " & smNowDate, vbExclamation, "Invalid Date"
    '        imBoxNo = SDATEINDEX
    '        mEnableBox imBoxNo
    '        Exit Sub
    '    End If
        'If tgSpf.sSMove = "Y" Then
        '    'Test that date is prior to last log date
        '    slEarliestLLD = ""
        '    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        '        If lbcVehicle.Selected(ilLoop) Then
        '            slNameCode = lbcVehCode.List(ilLoop)
        '            ilRet = gParseItem(slNameCode, 2, "\", slCode)
        '            ilVehCode = Val(slCode)
        '            ilVpfIndex = gVpfFind(CopyAsgn, ilVehCode)
        '            gUnpackDate tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), slLogDate
        '            If slLogDate <> "" Then
        '                If slEarliestLLD = "" Then
        '                    slEarliestLLD = slLogDate
        '                Else
        '                    If gDateValue(slLogDate) < gDateValue(slEarliestLLD) Then
        '                        slEarliestLLD = slLogDate
        '                    End If
        '                End If
        '            End If
        '        End If
        '    Next ilLoop
        '    If slEarliestLLD <> "" Then
        '        If gDateValue(smSave(1)) + Val(smSave(2)) - 1 > gDateValue(slEarliestLLD) Then
        '            Beep
        '            MsgBox "Date must be prior to or on " & slEarliestLLD, vbExclamation, "Invalid Date"
        '            imBoxNo = SDATEINDEX
        '            mEnableBox imBoxNo
        '            Exit Sub
        '        End If
        '    End If
        'End If
        'clTime1 = gTimeToCurrency(smSave(3), False)
        'clTime2 = gTimeToCurrency(smSave(4), True)
        'If (clTime2 < clTime1) Then
        '    Beep
        '    MsgBox "End Time earlier than Start Time.", vbExclamation, "Invalid Time"
        '    imBoxNo = ENDTIMEINDEX
        '    mEnableBox imBoxNo
        '    Exit Sub
        'End If

    'End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Select Case imBoxNo
        Case SDATEINDEX    'Start date
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                Exit Sub
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case NODAYSINDEX   'End Date
    End Select
    imLbcArrowSetting = False
End Sub
Private Sub edcDropDown_GotFocus()
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
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imBoxNo
        Case STIMEINDEX    'Start time
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                ilFound = False
                For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                    If KeyAscii = igLegalTime(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        Case ENDTIMEINDEX  'End Time
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                ilFound = False
                For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                    If KeyAscii = igLegalTime(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        Case SDATEINDEX   'Start date
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case NODAYSINDEX  'End Date
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case STIMEINDEX  'start time
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case ENDTIMEINDEX  'end time
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case SDATEINDEX    'start date
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case NODAYSINDEX  'end date
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNo
            Case SDATEINDEX   'Start date
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case NODAYSINDEX   'End Date
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)

        End Select
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
    DoEvents    'Process events so pending keys are not sent to this
                'form when keypreview turn on
    'CopyAsgn.KeyPreview = True  'To get Alt J and Alt L keys
    pbcSelections.Enabled = True
    pbcSTab.Enabled = True
    pbcTab.Enabled = True
    If (igWinStatus(COPYJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    CopyAsgn.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub Form_Deactivate()
    CopyAsgn.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
    End If

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    'Close btrieve files
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef

    'Delete arrays
    Erase tmCDCtrls
    Erase tmCtrls
    
    Set CopyAsgn = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcVehicle_Click()
    Dim ilLoop As Integer
    Dim slCommDate As String
    Dim slNameDate As String
    Dim slDate As String
    Dim ilNotMatching As Integer
    Dim ilPos As Integer
    Dim ilNoSelected As Integer
    Dim slStr As String
    'Vehicles with same date can only be picked
    If Not imChgMode Then
        imChgMode = True
        ilNoSelected = imNoSelected
        imNoSelected = 0
        slCommDate = "|"
        '11/9/96- install matching code back in
        ilNotMatching = False
        For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
            If lbcVehicle.Selected(ilLoop) Then
                slNameDate = lbcVehicle.List(ilLoop)
                ilPos = InStr(slNameDate, ": ")
                If ilPos > 0 Then
                    slStr = Mid$(slNameDate, ilPos + 2)
                    ilPos = InStr(slStr, "(N")
                    If ilPos > 0 Then
                        slDate = Left$(slStr, ilPos - 1)
                    Else
                        slDate = ""
                    End If
                Else
                    slDate = ""
                End If
                If slCommDate = "|" Then
                    slCommDate = slDate
                Else
                    If slCommDate = "" Then
                        If slDate <> "" Then
                            ilNotMatching = True
                            Exit For
                        End If
                    Else
                        If slDate = "" Then
                            ilNotMatching = True
                            Exit For
                        Else
                            If gDateValue(slCommDate) <> gDateValue(slDate) Then
                                ilNotMatching = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            Else
                imVefSelected(ilLoop) = False
            End If
        Next ilLoop
        If ilNotMatching Then
            For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                If lbcVehicle.Selected(ilLoop) Then
                    If Not imVefSelected(ilLoop) Then
                        lbcVehicle.Selected(ilLoop) = False
                    End If
                End If
            Next ilLoop
        Else
            For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                If lbcVehicle.Selected(ilLoop) Then
                    imNoSelected = imNoSelected + 1
                End If
            Next ilLoop
            'If imNoSelected > 1 Then    'Check that all have dates to assign copy
            '    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
            '        If lbcVehicle.Selected(ilLoop) Then
            '            ilPos = InStr(lbcVehicle.List(ilLoop), ": ")
            '            If ilPos = 0 Then
            '                Beep
            '                imNoSelected = 0
            '                For ilVeh = 0 To lbcVehicle.ListCount - 1 Step 1
            '                    lbcVehicle.Selected(ilVeh) = imVefSelected(ilVeh)
            '                    If imVefSelected(ilVeh) Then
            '                        imNoSelected = imNoSelected + 1
            '                    End If
            '                Next ilVeh
            '                imChgMode = False
            '                mSetCommands
            '                Exit Sub
            '            End If
            '        End If
            '    Next ilLoop
            'End If
            For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                imVefSelected(ilLoop) = lbcVehicle.Selected(ilLoop)
            Next ilLoop
            If imNoSelected > 1 Then    'Clear fields which are not allowed
                'smSave(1) = ""  '"Date Shown With Vehicle"
                'gSetShow pbcAsgn, smSave(1), tmCtrls(SDATEINDEX)
                ''smSave(3) = "12M"  'Start Time
                ''gSetShow pbcAsgn, smSave(3), tmCtrls(STIMEINDEX)
                ''smSave(4) = "12M"  'End Time
                ''gSetShow pbcAsgn, smSave(4), tmCtrls(ENDTIMEINDEX)
                'pbcAsgn.Cls
                'pbcAsgn_Paint
            ElseIf (imNoSelected = 1) And (ilNoSelected > 1) Then
                'smSave(1) = ""
                'gSetShow pbcAsgn, smSave(1), tmCtrls(SDATEINDEX)
                ''smSave(3) = ""  'Start Time
                ''gSetShow pbcAsgn, smSave(3), tmCtrls(STIMEINDEX)
                ''smSave(4) = ""  'End Time
                ''gSetShow pbcAsgn, smSave(4), tmCtrls(ENDTIMEINDEX)
                'pbcAsgn.Cls
                'pbcAsgn_Paint
            End If
        End If
        imChgMode = False
        mSetCommands
    End If
End Sub
Private Sub lbcVehicle_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
    End If
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
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

    slStr = edcDropDown.Text
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
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
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
    Dim ilLoop As Integer       'For loop control parameter
    Dim slStr As String         'Parse string
    Dim slCopyDate As String
    Dim ilPos As Integer
    Dim slNameDate As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case SDATEINDEX 'Start date
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcAsgn, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If smSave(1) = "" Then     'initialize
                ''Only one vehicle selected
                'slNameCode = lbcVehCode.List(lbcVehicle.ListIndex)
                'ilRet = gParseItem(slNameCode, 2, "\", slCode)
                'ilVehCode = Val(slCode)
                'ilVehIndex = gVpfFind(CopyAsgn, ilVehCode)
                'gUnpackDate tgVpf(ilVehIndex).iLLastDateCpyAsgn(0), tgVpf(ilVehIndex).iLLastDateCpyAsgn(1), slCopyDate
                'If slCopyDate <> "" Then
                '    slCopyDate = gIncOneDay(slCopyDate)
                '    'Determine next valid date to assign copy
                '    tmLcfSrchKey.sType = "O"
                '    tmLcfSrchKey.sStatus = "C"
                '    tmLcfSrchKey.iVefCode = ilVehCode
                '    gPackDate slCopyDate, tmLcfSrchKey.iLogDate(0), tmLcfSrchKey.iLogDate(1)
                '    tmLcfSrchKey.iSeqNo = 0
                '    ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                '    If (ilRet = BTRV_ERR_NONE) And (tmLcf.sType = "O") And (tmLcf.sStatus = "C") And (tmLcf.iVefCode = ilVehCode) Then
                '        gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slCopyDate
                '    Else
                '        'No dates in the future
                '        gUnpackDate tgVpf(ilVehIndex).iLLastDateCpyAsgn(0), tgVpf(ilVehIndex).iLLastDateCpyAsgn(1), slCopyDate
                '    End If
                slCopyDate = ""
                For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                    If lbcVehicle.Selected(ilLoop) Then
                        slNameDate = lbcVehicle.List(ilLoop)
                        ilPos = InStr(slNameDate, ": ")
                        If ilPos > 0 Then
                            slStr = Mid$(slNameDate, ilPos + 2)
                            ilPos = InStr(slStr, "(N")
                            If ilPos > 0 Then
                                slCopyDate = Left$(slStr, ilPos - 1)
                            Else
                                slCopyDate = ""
                            End If
                        Else
                            slCopyDate = ""
                        End If
                    End If
                Next ilLoop
                If slCopyDate = "" Then
                    slCopyDate = Format$(gNow(), "m/d/yy") 'Correctly format current date
                    slCopyDate = gIncOneDay(slCopyDate) 'Default date
                End If
                smSave(1) = slCopyDate
            End If
            pbcCalendar_Paint
            edcDropDown.Text = smSave(1)
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            pbcCalendar.Visible = True
            edcDropDown.SetFocus
        Case NODAYSINDEX 'End Date
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 3
            gMoveFormCtrl pbcAsgn, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            If smSave(2) = "" Then
                smSave(2) = "1"
            End If
            edcDropDown.Text = smSave(2)
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case STIMEINDEX 'Start time
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcAsgn, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If smSave(3) = "" Then  'default to midnight
                smSave(3) = "12M"
            End If
            edcDropDown.Text = smSave(3)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ENDTIMEINDEX 'End Time
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcAsgn, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If smSave(4) = "" Then  'default to midnight
                smSave(4) = "12M"
            End If
            edcDropDown.Text = smSave(4)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ONAIRINDEX  'Prelimary or final copy
            If imSave(1) = -1 Then
                imSave(1) = 0       'On Air  default
                tmCtrls(ilBoxNo).iChg = True
            End If
            pbcSelections.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAsgn, pbcSelections, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcSelections_Paint
            pbcSelections.Visible = True
            pbcSelections.SetFocus
        Case PFCOPYINDEX  'Prelimary or final copy
            If imSave(2) = -1 Then
                imSave(2) = 0       'Preliminary  default
                tmCtrls(ilBoxNo).iChg = True
            End If
            pbcSelections.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcAsgn, pbcSelections, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcSelections_Paint
            pbcSelections.Visible = True
            pbcSelections.SetFocus
    End Select
    mSetCommands   'Check mandatory fields and set controls
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Initialize module              *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilCount As Integer      'general counter
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    bmFirstCallToVpfFind = True
    'mParseCmmdLine
    mInitBox
    CopyAsgn.Height = cmcGenerate.Top + 5 * cmcGenerate.Height / 3
    gCenterStdAlone CopyAsgn
    'CopyAsgn.Show
    Screen.MousePointer = vbHourglass
    'Initialize variables
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imLBCtrls = 1
    imLBCDCtrls = 1
    sgCAVehicleTag = ""
    imNoSelected = 0
    imFirstFocus = True
    imTerminate = False         'terminate if true
    imBypassFocus = False       'don't bypass focus on any control
    imChgMode = False           'no change made
    imBSMode = False            'back space key
    imCalType = 0               'Standard type
    imBoxNo = -1                'Initialize current Box to N/A
    imTabDirection = 0          'Left to right movement
    imLbcArrowSetting = False   'List box invisible
    smDate = Format$(gNow(), "m/d/yy") 'Correctly format current date
    smNowDate = smDate
    smDate = gIncOneDay(smDate) 'Default date
    smDefaultTime = "12M"       'set default time to 12 midnight
    imVehCode = -1              'Invalidate vehicle code
    imPFAllowed = False         'Only final
    smDefaultDate = smDate      'temporarily set default date to now +1

    'Initialize save arrays
    For ilCount = LBound(smSave) To UBound(smSave) Step 1
        smSave(ilCount) = ""
    Next ilCount
    For ilCount = LBound(imSave) To UBound(imSave) Step 1
        imSave(ilCount) = -1
    Next ilCount
    'Open btrieve files
    imVefRecLen = Len(tmVef)    'Save VEF record length
    hmVef = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: VEF.BTR)", CopyAsgn
    On Error GoTo 0
    imVpfRecLen = Len(tmVpf)    'Save VPF record length
    hmVpf = CBtrvTable(TWOHANDLES)          'Save VEF handle
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: VPF.BTR)", CopyAsgn
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)    'Save LCF record length
    hmLcf = CBtrvTable(ONEHANDLE)          'Save LCF handle
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: LCF.BTR)", CopyAsgn
    On Error GoTo 0


    'Initialize positioning and show form

    lbcVehicle.Clear    'Initialize List boxes
    'lbcVehCode.Clear
    ReDim tgCAVehicle(0 To 0) As SORTCODE
    mPopulate           'Populate vehicle list boxes
    If imTerminate Then
        Exit Sub
    End If

    mSetCommands        'Set Commands
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
'*            Modified:5/4/94       By:D. Hannifan    *
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
    Dim ilLoop As Integer

    On Error GoTo mInitBoxErr
    flTextHeight = pbcAsgn.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcAsgn.Move 120, 3105, pbcAsgn.Width + fgPanelAdj, pbcAsgn.Height + fgPanelAdj
    pbcAsgn.Move plcAsgn.Left + fgBevelX, plcAsgn.Top + fgBevelY
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop

    'Control Fields
    'Start Date
    gSetCtrl tmCtrls(SDATEINDEX), 30, 30, 2805, fgBoxStH
    'End date or # of Weeks
    gSetCtrl tmCtrls(NODAYSINDEX), 2850, tmCtrls(SDATEINDEX).fBoxY, 2805, fgBoxStH
    'Start time
    gSetCtrl tmCtrls(STIMEINDEX), 30, tmCtrls(SDATEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    'End Time
    gSetCtrl tmCtrls(ENDTIMEINDEX), 2850, tmCtrls(STIMEINDEX).fBoxY, 2805, fgBoxStH
    'On air or alternate
    gSetCtrl tmCtrls(ONAIRINDEX), 30, tmCtrls(STIMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    'Preliminary or Final Copy
    gSetCtrl tmCtrls(PFCOPYINDEX), 2850, tmCtrls(ONAIRINDEX).fBoxY, 2805, fgBoxStH
    If Not imPFAllowed Then
        tmCtrls(PFCOPYINDEX).iReq = False
    End If
    Exit Sub
mInitBoxErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Populate Vehicle and time zone *
'*                      list boxes                     *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer            'return status
    Dim ilVehCode As Integer
    Dim llFilter As Long    'btrieve filter
    Dim ilVehIndex As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slCopyDate As String
    Dim slLastDate As String
    'Populate vehicle list box
    llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH ' Selling and conventional vehicles
    'ilRet = gPopUserVehicleBox(CopyAsgn, ilFilter, lbcVehicle, lbcVehCode)
    ilRet = gPopUserVehicleBox(CopyAsgn, llFilter, lbcVehicle, tgCAVehicle(), sgCAVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopUserVehicleBox)", CopyAsgn
        On Error GoTo 0
        'Add last copy date
        ReDim imVefSelected(0 To lbcVehicle.ListCount - 1) As Integer
        For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
            imVefSelected(ilLoop) = False
            slNameCode = tgCAVehicle(ilLoop).sKey    'lbcVehCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVehCode = Val(slCode)
            If bmFirstCallToVpfFind Then
                ilVehIndex = gVpfFind(CopyAsgn, ilVehCode)
                bmFirstCallToVpfFind = False
            Else
                ilVehIndex = gVpfFindIndex(ilVehCode)
            End If
            gUnpackDate tgVpf(ilVehIndex).iLLastDateCpyAsgn(0), tgVpf(ilVehIndex).iLLastDateCpyAsgn(1), slCopyDate
            If slCopyDate <> "" Then
                slLastDate = slCopyDate
                slCopyDate = gIncOneDay(slCopyDate)
                'Determine next valid date to assign copy
                tmLcfSrchKey.iType = 0
                tmLcfSrchKey.sStatus = "C"
                tmLcfSrchKey.iVefCode = ilVehCode
                gPackDate slCopyDate, tmLcfSrchKey.iLogDate(0), tmLcfSrchKey.iLogDate(1)
                tmLcfSrchKey.iSeqNo = 0
                ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If (ilRet = BTRV_ERR_NONE) And (tmLcf.iType = 0) And (tmLcf.sStatus = "C") And (tmLcf.iVefCode = ilVehCode) Then
                    gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slCopyDate
                    lbcVehicle.List(ilLoop) = lbcVehicle.List(ilLoop) & ": " & slCopyDate & "(Next), " & slLastDate & "(Last)"
                    If gDateValue(smDefaultDate) > gDateValue(slCopyDate) Then
                        smDefaultDate = slCopyDate
                    End If
                Else
                    lbcVehicle.List(ilLoop) = lbcVehicle.List(ilLoop) & ": " & slLastDate & "(Last)"
                End If
            Else
                smDefaultDate = Format$(gNow(), "m/d/yy") 'Correctly format current date
                smDefaultDate = gIncOneDay(smDefaultDate) 'Default date
            End If
        Next ilLoop
    End If

    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:4/27/94       By:D. Hannifan    *
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
    'Check all mandatory control fields
    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = NO Then
        cmcGenerate.Enabled = False
        Exit Sub
    End If
    If imUpdateAllowed Then
        cmcGenerate.Enabled = True
    Else
        cmcGenerate.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Set focus specified control    *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If


    Select Case ilBoxNo 'Branch on box type (control)
        Case SDATEINDEX 'Start Date
            If edcDropDown.Enabled And edcDropDown.Visible Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case NODAYSINDEX 'End Date
            If edcDropDown.Enabled And edcDropDown.Visible Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case STIMEINDEX 'Start Time
            If edcDropDown.Enabled And edcDropDown.Visible Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case ENDTIMEINDEX 'End Time
            If edcDropDown.Enabled And edcDropDown.Visible Then
                edcDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case ONAIRINDEX   'Preliminary or final
            If pbcSelections.Enabled And pbcSelections.Visible Then
                pbcSelections.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case PFCOPYINDEX   'Preliminary or final
            If pbcSelections.Enabled And pbcSelections.Visible Then
                pbcSelections.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
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
    Dim slStr As String     'show string
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case SDATEINDEX 'Start Date
            edcDropDown.Visible = False  'Set visibility
            cmcDropDown.Visible = False
            plcCalendar.Visible = False
            slStr = edcDropDown.Text
            smSave(1) = slStr
            gSetShow pbcAsgn, slStr, tmCtrls(ilBoxNo)
        Case NODAYSINDEX 'End Date
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smSave(2) = slStr
            gSetShow pbcAsgn, slStr, tmCtrls(ilBoxNo)
       Case STIMEINDEX 'Start Time
            edcDropDown.Visible = False  'Set visibility
            cmcDropDown.Visible = False
            plcTme.Visible = False
            slStr = edcDropDown.Text
            slStr = gFormatTime(slStr, "A", "2")
            smSave(3) = slStr
            gSetShow pbcAsgn, slStr, tmCtrls(ilBoxNo)
       Case ENDTIMEINDEX 'End Time
            edcDropDown.Visible = False  'Set visibility
            cmcDropDown.Visible = False
            plcTme.Visible = False
            slStr = edcDropDown.Text
            slStr = gFormatTime(slStr, "A", "2")
            smSave(4) = slStr
            gSetShow pbcAsgn, slStr, tmCtrls(ilBoxNo)
       Case ONAIRINDEX 'Preliminary or final copy
            pbcSelections.Visible = False  'Set visibility
            If imSave(1) = 0 Then
                slStr = "On Air"
            ElseIf imSave(1) = 1 Then
                slStr = "Alternate"
            Else
                slStr = ""
            End If
            gSetShow pbcAsgn, slStr, tmCtrls(ilBoxNo)
       Case PFCOPYINDEX 'Preliminary or final copy
            pbcSelections.Visible = False  'Set visibility
            If imSave(2) = 0 Then
                slStr = "Preliminary Copy"
            ElseIf imSave(2) = 1 Then
                slStr = "Final Copy"
            Else
                slStr = ""
            End If
            gSetShow pbcAsgn, slStr, tmCtrls(ilBoxNo)
    End Select
    mSetCommands
    cmcCancel.Caption = "&Cancel"
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'


    imTerminate = False
    sgDoneMsg = Trim$(str$(igCopyAsgnCallSource)) & "\" & "Done"
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload CopyAsgn
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:4/27/94       By:D. Hannifan    *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim slStr As String     'Text string
    If (ilCtrlNo = SDATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imNoSelected = 1 Then
            If gFieldDefinedStr(smSave(1), "", "Start date must be specified", tmCtrls(SDATEINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = SDATEINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
            If Not gValidDate(smSave(1)) Then  'invalid date
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = NODAYSINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSave(2), "", "Number of days must be specified", tmCtrls(NODAYSINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NODAYSINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = STIMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        'If imNoSelected = 1 Then
            If gFieldDefinedStr(smSave(3), "", "Start time must be specified", tmCtrls(STIMEINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = STIMEINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
            If (smSave(3) = "") Then
                mTestFields = NO
                Exit Function
            End If
            If Not gValidTime(smSave(3)) Then
                mTestFields = NO
                Exit Function
            End If
        'End If
    End If
    If (ilCtrlNo = ENDTIMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        'If imNoSelected = 1 Then
            If gFieldDefinedStr(smSave(4), "", "Valid End Time must be specified", tmCtrls(ENDTIMEINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = ENDTIMEINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
            If smSave(4) = "" Then
                mTestFields = NO
                Exit Function
            End If
            If Not gValidTime(smSave(4)) Then
                mTestFields = NO
                Exit Function
            End If
        'End If
    End If
    If (ilCtrlNo = ONAIRINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imSave(1) = 0 Then
            slStr = "On Air"
        ElseIf imSave(1) = 1 Then
            slStr = "Alternate"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "On Air Or Alternate must be specified", tmCtrls(PFCOPYINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ONAIRINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PFCOPYINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        'If (Not imPFAllowed) And (imSave(2) = -1) Then
        '    imSave(2) = 1
        'End If
        If imPFAllowed Then
            If imSave(2) = 0 Then
                slStr = "Preliminary Copy"
            ElseIf imSave(2) = 1 Then
                slStr = "Final Copy"
            Else
                slStr = ""
            End If
            If gFieldDefinedStr(slStr, "", "Copy assignment must be specified", tmCtrls(PFCOPYINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = PFCOPYINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    mTestFields = YES
End Function
Private Sub pbcAsgn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    If imNoSelected <= 0 Then
        lbcVehicle.SetFocus
        Exit Sub
    End If
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                If (Not imPFAllowed) And (ilBox = PFCOPYINDEX) Then
                    Beep
                    pbcClickFocus.SetFocus
                    Exit Sub
                End If
                If (imNoSelected <= 0) Then
                    Beep
                    pbcClickFocus.SetFocus
                    Exit Sub
                End If
                'If (imNoSelected <> 1) And (ilBox = SDATEINDEX) Then
                '    Beep
                '    pbcClickFocus.SetFocus
                '    Exit Sub
                'End If
                'If (imNoSelected <> 1) And (ilBox = STIMEINDEX) Then
                '    Beep
                '    pbcClickFocus.SetFocus
                '    Exit Sub
                'End If
                'If (imNoSelected <> 1) And (ilBox = ENDTIMEINDEX) Then
                '    Beep
                '    pbcClickFocus.SetFocus
                '    Exit Sub
                'End If
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub
Private Sub pbcAsgn_Paint()
    Dim ilBox As Integer
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        pbcAsgn.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcAsgn.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcAsgn.Print tmCtrls(ilBox).sShow
    Next ilBox
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
                edcDropDown.Text = Format$(llDate, "m/d/yy")
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                imBypassFocus = True
                edcDropDown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate

    edcDropDown.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcSelections_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcSelections_KeyPress(KeyAscii As Integer)
    If imBoxNo = ONAIRINDEX Then   'On air
        If KeyAscii = Asc("O") Or (KeyAscii = Asc("o")) Then
            If (imSave(1) <> 0) Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(1) = 0
            pbcSelections_Paint
        ElseIf KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
            If (imSave(1) <> 1) Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(1) = 1
            pbcSelections_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If (imSave(1) = 0) Then
                tmCtrls(imBoxNo).iChg = True
                imSave(1) = 1
                pbcSelections_Paint
            ElseIf (imSave(1) = 1) Then
                tmCtrls(imBoxNo).iChg = True
                imSave(1) = 0
                pbcSelections_Paint
            End If
        End If
    End If
    If imBoxNo = PFCOPYINDEX Then   'Preliminary or final copy
        If KeyAscii = Asc("P") Or (KeyAscii = Asc("p")) Then
            If (imSave(2) <> 0) Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(2) = 0
            pbcSelections_Paint
        ElseIf KeyAscii = Asc("F") Or (KeyAscii = Asc("f")) Then
            If (imSave(2) <> 1) Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(2) = 1
            pbcSelections_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If (imSave(2) = 0) Then
                tmCtrls(imBoxNo).iChg = True
                imSave(2) = 1
                pbcSelections_Paint
            ElseIf (imSave(2) = 1) Then
                tmCtrls(imBoxNo).iChg = True
                imSave(2) = 0
                pbcSelections_Paint
            End If
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcSelections_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = ONAIRINDEX Then     'On air
        If (imSave(1) = 0) Then
            tmCtrls(imBoxNo).iChg = True
            imSave(1) = 1
            pbcSelections_Paint
        ElseIf (imSave(1) = 1) Then
            tmCtrls(imBoxNo).iChg = True
            imSave(1) = 0
            pbcSelections_Paint
        End If
    End If
    If imBoxNo = PFCOPYINDEX Then     'Preliminary or final copy
        If (imSave(2) = 0) Then
            tmCtrls(imBoxNo).iChg = True
            imSave(2) = 1
            pbcSelections_Paint
        ElseIf (imSave(2) = 1) Then
            tmCtrls(imBoxNo).iChg = True
            imSave(2) = 0
            pbcSelections_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcSelections_Paint()
    pbcSelections.Cls
    pbcSelections.CurrentX = fgBoxInsetX
    pbcSelections.CurrentY = 0 'fgBoxInsetY
    If imBoxNo = ONAIRINDEX Then     'On air
        If imSave(1) = 0 Then
            pbcSelections.Print "On Air"
        ElseIf imSave(1) = 1 Then
            pbcSelections.Print "Alternate"
        Else
            pbcSelections.Print "   "
        End If
    End If
    If imBoxNo = PFCOPYINDEX Then     'Preliminary or final copy
        If imSave(2) = 0 Then
            pbcSelections.Print "Preliminary Copy"
        ElseIf imSave(2) = 1 Then
            pbcSelections.Print "Final Copy"
        Else
            pbcSelections.Print "   "
        End If
    End If
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer    'control index
    Dim ilFound As Integer  'loop exit flag
    Dim slDate As String    'date string
    Dim slTime As String    'time string
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    imTabDirection = -1  'Set-right to left
    If imNoSelected <= 0 Then
        lbcVehicle.SetFocus
        Exit Sub
    End If
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1              'Initial
                imTabDirection = 0  'Set-Left to right
                'If imNoSelected = 1 Then
                    ilBox = imLBCtrls
                'Else
                '    ilBox = NODAYSINDEX
                'End If
                mSetCommands
            Case SDATEINDEX       'Start date
                slDate = edcDropDown.Text
                If gValidDate(slDate) Then
                    ilBox = ilBox - 1
                Else                      'Invalid date
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
                mSetShow imBoxNo
                imBoxNo = -1
                lbcVehicle.SetFocus
                Exit Sub
            Case NODAYSINDEX     'End Date
                slDate = edcDropDown.Text
                If slDate = "" Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
                'If imNoSelected = 1 Then
                    ilBox = ilBox - 1
                'Else
                '    lbcVehicle.SetFocus
                '    Exit Sub
                'End If
            Case STIMEINDEX    'Start Time
                If edcDropDown.Text = "" Then  'invalid time
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
                slTime = edcDropDown.Text
                If gValidTime(slTime) Then
                    ilBox = ilBox - 1
                Else             'Invalid time
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Case ENDTIMEINDEX      'End Time
                If edcDropDown.Text = "" Then 'Invalid time
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
                slTime = edcDropDown.Text
                If gValidTime(slTime) Then 'valid time
                    ilBox = ilBox - 1
                Else        'Invalid time
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Case ONAIRINDEX
                'If imNoSelected = 1 Then
                    ilBox = ENDTIMEINDEX
                'Else
                '    ilBox = NODAYSINDEX
                'End If
            Case Else
                ilBox = ilBox - 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer    'local Control box counter
    Dim ilFound As Integer  'redirect focus flag
    Dim slDate As String    'date string
    Dim slTime As String    'time string
    Dim slStr As String
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    imTabDirection = 0  'Set-Left to right
    If imNoSelected <= 0 Then
        lbcVehicle.SetFocus
        Exit Sub
    End If
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1          'Initial
                imTabDirection = -1  'Set-Right to left
                'If preliminary not allowed
                If imPFAllowed Then
                    ilBox = UBound(tmCtrls)
                Else
                    If imSave(2) = -1 Then
                        imSave(2) = 1
                        slStr = "Final Copy"
                        gSetShow pbcAsgn, slStr, tmCtrls(PFCOPYINDEX)
                        gPaintArea pbcAsgn, tmCtrls(PFCOPYINDEX).fBoxX, tmCtrls(PFCOPYINDEX).fBoxY + fgOffset - 15, tmCtrls(PFCOPYINDEX).fBoxW - 15, fgBoxGridH, WHITE
                        pbcAsgn_Paint
                    End If
                    ilBox = ONAIRINDEX
                End If
            Case ENDTIMEINDEX      'End Time
                If edcDropDown.Text = "" Then 'Invalid time
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
                slTime = edcDropDown.Text
                If gValidTime(slTime) Then 'valid time
                    ilBox = ilBox + 1
                Else        'Invalid time
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Case STIMEINDEX      'Start Time
                If edcDropDown.Text = "" Then  'invalid time
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
                slTime = edcDropDown.Text
                If gValidTime(slTime) Then
                    ilBox = ilBox + 1
                Else             'Invalid time
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Case SDATEINDEX      'Start Date
                slDate = edcDropDown.Text
                If gValidDate(slDate) Then
                    ilBox = ilBox + 1
                Else                      'Invalid date
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Case NODAYSINDEX      'End Date
                slDate = edcDropDown.Text
                If slDate = "" Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
                'If imNoSelected <> 1 Then
                '    ilBox = ONAIRINDEX
                'Else
                    ilBox = ilBox + 1
                'End If
            Case ONAIRINDEX
                'If not allowing Preliminary- set to Final
                If Not imPFAllowed Then
                    If imSave(2) = -1 Then
                        imSave(2) = 1
                        slStr = "Final Copy"
                        gSetShow pbcAsgn, slStr, tmCtrls(PFCOPYINDEX)
                        gPaintArea pbcAsgn, tmCtrls(PFCOPYINDEX).fBoxX, tmCtrls(PFCOPYINDEX).fBoxY + fgOffset - 15, tmCtrls(PFCOPYINDEX).fBoxW - 15, fgBoxGridH, WHITE
                        pbcAsgn_Paint
                    End If
                    ilFound = False
                End If
                ilBox = ilBox + 1   'Bypass PFCOPYINDEX
            Case PFCOPYINDEX
                mSetShow imBoxNo
                imBoxNo = -1
                If cmcGenerate.Enabled Then
                    cmcGenerate.SetFocus
                Else
                    cmcCancel.SetFocus
                End If
                Exit Sub
            Case Else
                ilBox = ilBox + 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    Select Case imBoxNo
                        Case STIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                        Case ENDTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                    End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub plcAsgn_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Assign Copy"
End Sub
Private Sub plcVehicle_Paint()
    plcVehicle.CurrentX = 0
    plcVehicle.CurrentY = 0
    plcVehicle.Print "Vehicles"
End Sub
