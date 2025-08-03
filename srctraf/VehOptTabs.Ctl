VERSION 5.00
Begin VB.UserControl VehOptTabs 
   Appearance      =   0  'Flat
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10455
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3240
   ScaleWidth      =   10455
   Begin VB.PictureBox plcAffLog 
      Height          =   2535
      Left            =   705
      ScaleHeight     =   2475
      ScaleWidth      =   8880
      TabIndex        =   0
      Top             =   510
      Width           =   8940
      Begin VB.CheckBox ckcWebInfo 
         Caption         =   "Cart #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   6585
         TabIndex        =   5
         Top             =   75
         Width           =   1080
      End
      Begin VB.CheckBox ckcAffLog 
         Caption         =   "Suppress Rotation Comments on Electronic Affidavit System (Web)"
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
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   2085
         Width           =   6270
      End
      Begin VB.TextBox edcLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3540
         MaxLength       =   2
         TabIndex        =   13
         Top             =   1665
         Width           =   405
      End
      Begin VB.CheckBox ckcWebInfo 
         Caption         =   "Export to Web after Airing"
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
         Index           =   4
         Left            =   3540
         TabIndex        =   7
         Top             =   450
         Width           =   3240
      End
      Begin VB.CheckBox ckcAffLog 
         Caption         =   "Replacements"
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
         Index           =   1
         Left            =   3570
         TabIndex        =   12
         Top             =   1290
         Width           =   1650
      End
      Begin VB.CheckBox ckcAffLog 
         Caption         =   "MGs"
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
         Index           =   0
         Left            =   2535
         TabIndex        =   11
         Top             =   1290
         Width           =   1020
      End
      Begin VB.TextBox edcLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1860
         MaxLength       =   250
         TabIndex        =   9
         Top             =   795
         Width           =   6750
      End
      Begin VB.CheckBox ckcWebInfo 
         Caption         =   "Suppress Log/Audio on Web Site"
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
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   450
         Width           =   3240
      End
      Begin VB.CheckBox ckcWebInfo 
         Caption         =   "Avail Names"
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
         Left            =   1770
         TabIndex        =   2
         Top             =   75
         Width           =   1515
      End
      Begin VB.CheckBox ckcWebInfo 
         Caption         =   "Feed Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3360
         TabIndex        =   3
         Top             =   75
         Width           =   1260
      End
      Begin VB.CheckBox ckcWebInfo 
         Caption         =   "Spot Summary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   4785
         TabIndex        =   4
         Top             =   75
         Width           =   1680
      End
      Begin VB.Label lacLog 
         Appearance      =   0  'Flat
         Caption         =   "+/- Live Compliant Window in Minutes"
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
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1725
         Width           =   3135
      End
      Begin VB.Label lacAffLog 
         Caption         =   "For Affiliate Logs, Allow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1275
         Width           =   2385
      End
      Begin VB.Label lacLog 
         Appearance      =   0  'Flat
         Caption         =   "Audio FTP Address"
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
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1785
      End
      Begin VB.Label lacLog 
         Caption         =   "On Web Log Show"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   1665
      End
   End
End
Attribute VB_Name = "VehOptTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: VehOptTabs.ctl
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text

Public Event SetSave(slType As String)

Dim imUpdateAllowed As Integer
Dim imIgnoreClickEvent As Integer
Dim imFirstActivate As Integer
Dim imTerminate As Integer



Public Sub Action(ilType As Integer, ilState As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilIndex As Integer
    Select Case ilType
        Case 1  'Clear Focus
        Case 2  'Init function
            'Test if unloading control
            ilRet = 0
            On Error GoTo UserControlErr:
            ilIndex = VehOpt!tbcSelection.SelectedItem.Index - 1
            If ilRet = 0 Then
                Form_Load
                Form_Activate
                Select Case ilIndex
                    Case 0  'General,...
                    Case 1  'Sales
                    Case 2  'Schedule
                    Case 3  'Sports
                    Case 4  'PSA/Promo
                    Case 5  'Log
                    Case 6  'Producter
                    Case 7  'Export
                    Case 8  '6  'Accounting
                    Case 9  '7  'Barter
                    Case 10  '8  'Participant
                    Case 11 '9  'Great Plains G/L 'allowed packages 5/7/08
                    Case 12 'Affiliate Log
                        plcAffLog.Visible = True
                        plcAffLog.Enabled = ilState
                End Select
            End If
        Case 3  'terminate function
        Case 4  'Clear
            Screen.MousePointer = vbDefault
        Case 5  'Save
            plcAffLog.Visible = False
        Case 6  'Features turned on/off depending on answers on other fields
            'If ilState = 0 Then
            '    frcRepNet.Visible = False
            'ElseIf ilState = 1 Then
            '    frcRepNet.Visible = True
            '    lacDBID.Visible = True
            '    edcDBID.Visible = True
            '    frcNet.Visible = False
            'Else
            '    frcRepNet.Visible = True
            '    lacDBID.Visible = True
            '    edcDBID.Visible = True
            '    frcNet.Visible = True
            'End If
        Case 7
            If ilState = 0 Then
                imIgnoreClickEvent = True
            Else
                imIgnoreClickEvent = False
            End If
    End Select
    Exit Sub
UserControlErr:
    ilRet = 1
    Resume Next
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Exit Sub
    End If
    imFirstActivate = False
    imUpdateAllowed = igUpdateAllowed
    'If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
    If Not imUpdateAllowed Then
    Else
    End If
    gShowBranner imUpdateAllowed
End Sub

Private Sub Form_Click()
End Sub

Private Sub Form_Deactivate()
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub
Public Property Get WebInfo(ilIndex As Integer) As Integer
    WebInfo = ckcWebInfo(ilIndex).Value
End Property
Public Property Let WebInfo(ilIndex As Integer, ilValue As Integer)
    ckcWebInfo(ilIndex).Value = ilValue
End Property
Public Property Get FTPInfo() As String
    FTPInfo = edcLog(0).Text
End Property
Public Property Let FTPInfo(slText As String)
    edcLog(0).Text = slText
End Property
Public Property Get LiveWindow() As String
    LiveWindow = edcLog(1).Text
End Property
Public Property Let LiveWindow(slText As String)
    edcLog(1).Text = slText
End Property
Public Property Get AffLog(ilIndex As Integer) As Integer
    AffLog = ckcAffLog(ilIndex).Value
End Property
Public Property Let AffLog(ilIndex As Integer, ilValue As Integer)
    ckcAffLog(ilIndex).Value = ilValue
End Property
Public Property Let Enabled(ilState As Integer)
    UserControl.Enabled = ilState
    PropertyChanged "Enabled"
End Property

Private Sub mInit()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         slNameCode                    slName                    *
'*  slCode                        ilLoop                        slDaypart                 *
'*  slLineNo                      slStr                                                   *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mInitErr                                                                              *
'******************************************************************************************

'
'   mInit
'   Where:
'

    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    mInitBox
    plcAffLog.BorderStyle = 0
    plcAffLog.Visible = False
    plcAffLog.Move 0, 0

    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
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
End Sub

Private Sub ckcAffLog_Click(Index As Integer)
    RaiseEvent SetSave("AffLog")
End Sub

Private Sub ckcWebInfo_Click(Index As Integer)
    RaiseEvent SetSave("WebInfo")
End Sub

Private Sub edcLog_Change(Index As Integer)
    If Index = 0 Then
        RaiseEvent SetSave("FTPInfo")
    ElseIf Index = 1 Then
        RaiseEvent SetSave("LiveWindow")
    End If
End Sub

Private Sub edcLog_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Public Property Let AllowMGEnabled(ilState As Integer)
    ckcAffLog(0).Enabled = ilState
End Property
Public Property Get AllowMGEnabled() As Integer
    AllowMGEnabled = ckcAffLog(0).Enabled
End Property

Public Property Get AllowReplEnabled() As Integer
    AllowReplEnabled = ckcAffLog(1).Enabled
End Property

Public Property Let AllowReplEnabled(ilState As Integer)
    ckcAffLog(1).Enabled = ilState
End Property

Public Property Get SuppressRotCommEnabled() As Integer
    SuppressRotCommEnabled = ckcAffLog(2).Enabled
End Property

Public Property Let SuppressRotCommEnabled(ilState As Integer)
    ckcAffLog(2).Enabled = ilState
End Property

