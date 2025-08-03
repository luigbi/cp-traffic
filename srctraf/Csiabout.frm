VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form CSIAbout 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5085
   ClientLeft      =   360
   ClientTop       =   3330
   ClientWidth     =   8895
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
   Icon            =   "Csiabout.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5085
   ScaleWidth      =   8895
   Visible         =   0   'False
   Begin VB.PictureBox plcAbout 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5070
      Left            =   15
      ScaleHeight     =   5070
      ScaleWidth      =   8850
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   -15
      Width           =   8850
      Begin VB.CommandButton cmcSymb 
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
         Height          =   645
         Left            =   375
         Picture         =   "Csiabout.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1455
         Width           =   645
      End
      Begin VB.ListBox lbcInfo 
         Appearance      =   0  'Flat
         Height          =   870
         Index           =   1
         Left            =   300
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   3360
         Width           =   6045
      End
      Begin VB.PictureBox pbcAVI 
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
         Height          =   2235
         Left            =   6525
         ScaleHeight     =   2235
         ScaleWidth      =   1965
         TabIndex        =   10
         Top             =   2220
         Width           =   1965
      End
      Begin VB.ListBox lbcInfo 
         Appearance      =   0  'Flat
         Height          =   870
         Index           =   0
         Left            =   300
         TabIndex        =   0
         Top             =   2235
         Width           =   6045
      End
      Begin VB.CommandButton cmcOk 
         Appearance      =   0  'Flat
         Caption         =   "&Ok"
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
         Left            =   3840
         TabIndex        =   1
         Top             =   4560
         Width           =   1125
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
         Left            =   165
         ScaleHeight     =   165
         ScaleWidth      =   105
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2580
         Width           =   105
      End
      Begin VB.CommandButton cmcCSLogo 
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
         Height          =   555
         Left            =   375
         Picture         =   "Csiabout.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   750
         Width           =   585
      End
      Begin VB.PictureBox plcName 
         BackColor       =   &H00FFFFFF&
         Height          =   1005
         Left            =   1200
         ScaleHeight     =   945
         ScaleWidth      =   7095
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   420
         Width           =   7155
         Begin VB.Label lacMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Traffic System Version xx.xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   6945
         End
         Begin VB.Label lacMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Copyright© 1993-2002 Counterpoint Software ®"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   16
            Top             =   330
            Width           =   6945
         End
         Begin VB.Label lacMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "All Rights Reserved"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   0
            TabIndex        =   15
            Top             =   660
            Width           =   6945
         End
      End
      Begin VB.PictureBox plcStation 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   630
         Left            =   1200
         ScaleHeight     =   570
         ScaleWidth      =   7095
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1470
         Width           =   7155
         Begin VB.Label lacMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "This product is Licensed to:"
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
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   6945
         End
         Begin VB.Label lacMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   15
            TabIndex        =   13
            Top             =   210
            Width           =   6945
         End
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
      Left            =   195
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
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
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   15
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
      Left            =   420
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   525
   End
   Begin MCI.MMControl mmcAbout 
      Height          =   405
      Left            =   8430
      TabIndex        =   18
      Top             =   4620
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   0
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "AVIVideo"
      FileName        =   ""
   End
End
Attribute VB_Name = "CSIAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Csiabout.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CSIAbout.Frm
'
' Release: 1.0
'
' Description:
'   This file contains About code
Option Explicit
Option Compare Text
'Program library dates Field Areas
'General Files
Dim tmSpf As SPF        'SPF record image
Dim hmSpf As Integer    'Site preference handle
Dim imSpfRecLen As Integer        'SPF record length
Dim tmVef As VEF        'VEF record image
Dim hmVef As Integer    'Vehicle file handle
Dim imVefRecLen As Integer        'VEF record length
Dim tmVpf As VPF        'VPF record image
Dim tmVpfSrchKey As VPFKEY0    'VPF key record image
Dim hmVpf As Integer    'Vehicle preference file handle
Dim imVpfRecLen As Integer        'VPF record length
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imShowingCSI As Integer
'Dates
Dim smNowDate As String
Dim lmNowDate As Long
Dim imFirstActivate As Integer

Private Sub cmcCSLogo_Click()
    If imShowingCSI Then
        imShowingCSI = False
        mmcAbout.Command = "Stop"
        mmcAbout.Command = "Close"
        mmcAbout.Visible = False
    Else
        imShowingCSI = True
        mmcAbout.Visible = True
        mmcAbout.hWndDisplay = pbcAVI.hWnd
        mmcAbout.Command = "Close"
        mmcAbout.DeviceType = "AVIVideo"
        mmcAbout.fileName = sgExePath & "CSIAbout.Avi"
        mmcAbout.Command = "Open"
        mmcAbout.Command = "Play"
    End If
End Sub
Private Sub cmcOk_Click()
    mTerminate
End Sub
Private Sub cmcSymb_Click()
    'mmcSymb.Command = "Close"
    'mmcSymb.DeviceType = "WaveAudio"
    'mmcSymb.fileName = sgLogoPath & "CustSymb.Wav"
    'mmcSymb.Command = "Open"
    'mmcSymb.Command = "Play"
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner
    Me.KeyPreview = True
    CSIAbout.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    If App.PrevInstance Then
        End
    End If
    mInit
    If imTerminate Then
        mTerminate
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    mmcAbout.Command = "Stop"
    mmcAbout.Command = "Close"
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    btrExtClear hmSpf   'Clear any previous extend operation
    ilRet = btrClose(hmSpf)
    btrDestroy hmSpf
    btrDestroy hmVpf
    
    Set CSIAbout = Nothing   'Remove data segment
    
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
    Dim ilRet As Integer
    Dim slDate As String
    Dim slWinDir As String
    Dim llWinLen As Long
    Dim slWindowsDir As String * MAX_PATH

    lacMsg(1).Caption = "Copyrighted © 1993-" & Year(Now) & " Counterpoint Software, Inc. ®"

    imTerminate = False
    imFirstActivate = True

    'Screen.MousePointer = vbHourGlass
    'mParseCmmdLine
    CSIAbout.Height = plcAbout.Height + fgPanelAdj
    CSIAbout.Width = plcAbout.Width + fgPanelAdj
    gCenterStdAlone CSIAbout
    'CSIAbout.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    cmcCSLogo.Picture = IconTraf!imcCSLogo.Picture
    cmcSymb.Picture = Traffic!pbcSymb.Picture
    'ilPos = InStr(sgCSVersion, "created")
    'If ilPos > 0 Then
    '    lacMsg(0).Caption = "Traffic System " & Left$(sgCSVersion, ilPos - 1) '& ", O.S. # " & fgWinVersion
    'Else
        lacMsg(0).Caption = "Traffic System " & sgCSVersion '& ", O.S. # " & fgWinVersion
    'End If
    'CSIAbout.Show
    imFirstFocus = True
    imShowingCSI = False
    smNowDate = Format$(Now, "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    hmSpf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Spf.Btr)", CSIAbout
    On Error GoTo 0
    imSpfRecLen = Len(tmSpf)
    'Until multi-user version of btrieve installed- test field
    ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    hmVef = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", CSIAbout
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVpf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vpf.Btr)", CSIAbout
    On Error GoTo 0
    imVpfRecLen = Len(tmVpf)

    lacMsg(4).Caption = Trim$(tmSpf.sGClient)
    lbcInfo(0).AddItem "DDF Info: " & sgDDFDateInfo
    gUnpackDate tmSpf.iBLastStdMnth(0), tmSpf.iBLastStdMnth(1), slDate
    lbcInfo(0).AddItem "Last Standard Month Billing Date " & slDate
    gUnpackDate tmSpf.iBLastCalMnth(0), tmSpf.iBLastCalMnth(1), slDate
    lbcInfo(0).AddItem "Last Calendar Month Billing Date " & slDate
    gUnpackDate tmSpf.iRLastPay(0), tmSpf.iRLastPay(1), slDate
    lbcInfo(0).AddItem "Date of Last Payment Received " & slDate
    gUnpackDate tmSpf.iRCreditDate(0), tmSpf.iRCreditDate(1), slDate
    lbcInfo(0).AddItem "Date Advertisers & Agencies Unbilled Values Computed " & slDate
    ilRet = btrGetFirst(hmVef, tmVef, imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If tmVef.sType <> "P" Then  'Bypass package vehicles
            tmVpfSrchKey.iVefKCode = tmVef.iCode
            ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDate tmVpf.iLLD(0), tmVpf.iLLD(1), slDate
                lbcInfo(1).AddItem Trim$(tmVef.sName) & ": Last Log Date " & slDate
                gUnpackDate tmVpf.iLLastDateCpyAsgn(0), tmVpf.iLLastDateCpyAsgn(1), slDate
                lbcInfo(1).AddItem Trim$(tmVef.sName) & ": Last Date Copy Assigned " & slDate
            End If
        End If
        ilRet = btrGetNext(hmVef, tmVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    llWinLen = GetWindowsDirectory(slWindowsDir, MAX_PATH)
    slWinDir = Left$(slWindowsDir, llWinLen)
    lbcInfo(1).AddItem "Base Windows Folder: " & slWinDir, lbcInfo(1).ListCount
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub mmcAbout_Done(NotifyCode As Integer)
    mmcAbout.Command = "Prev"
    mmcAbout.Command = "Play"
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
    'Unload Traffic
    Unload CSIAbout
    igManUnload = NO
End Sub
Private Sub plcAbout_Paint()
    plcAbout.CurrentX = 0
    plcAbout.CurrentY = 0
    plcAbout.Print "Counterpoint Software ®"
End Sub
Private Sub plcName_Paint()
    plcName.CurrentX = 0
    plcName.CurrentY = 0
    plcName.Print "Panel3D1"
End Sub
