VERSION 5.00
Begin VB.Form LlAddSpt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5070
   ClientLeft      =   2640
   ClientTop       =   2610
   ClientWidth     =   10740
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
   ScaleHeight     =   5070
   ScaleWidth      =   10740
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5655
      TabIndex        =   10
      Top             =   4620
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
      Left            =   45
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4845
      Width           =   120
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Add"
      Height          =   285
      Left            =   4065
      TabIndex        =   9
      Top             =   4620
      Width           =   1050
   End
   Begin VB.PictureBox plcAddSpot 
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
      Height          =   4035
      Left            =   105
      ScaleHeight     =   3975
      ScaleWidth      =   10350
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   10410
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Only Active Copy"
         Height          =   225
         Left            =   4425
         TabIndex        =   7
         Top             =   825
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox ckcAdvt 
         Caption         =   "Only Advertisers in Live Log"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   825
         Value           =   1  'Checked
         Width           =   3435
      End
      Begin VB.ListBox lbcCopy 
         Height          =   2790
         ItemData        =   "LLAddSpt.frx":0000
         Left            =   4425
         List            =   "LLAddSpt.frx":0002
         MultiSelect     =   1  'Simple
         TabIndex        =   8
         Top             =   1110
         Width           =   5805
      End
      Begin VB.ListBox lbcAdvt 
         Height          =   2790
         ItemData        =   "LLAddSpt.frx":0004
         Left            =   120
         List            =   "LLAddSpt.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1110
         Width           =   4020
      End
      Begin VB.TextBox edcTime 
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
         Height          =   285
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   2
         Top             =   75
         Width           =   1545
      End
      Begin VB.Label lacCopy 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Copy"
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
         Left            =   4440
         TabIndex        =   6
         Top             =   555
         Width           =   5535
      End
      Begin VB.Label lacAdvt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Advertiser"
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
         Left            =   150
         TabIndex        =   3
         Top             =   555
         Width           =   3900
      End
      Begin VB.Label lacTime 
         Appearance      =   0  'Flat
         Caption         =   "Aired Time"
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
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.Label lacScreen 
      Caption         =   "Add Bonus Spot"
      Height          =   195
      Left            =   75
      TabIndex        =   12
      Top             =   30
      Width           =   1515
   End
End
Attribute VB_Name = "LlAddSpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of LLAddSpt.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: LlAddSpt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract number screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer

Dim tmCopyCode() As SORTCODE
Dim smCopyCodeTag As String


Private Sub ckcAdvt_Click()
    mPopAdvt
End Sub

Private Sub ckcCopy_Click()
    mPopCopy
End Sub

Private Sub cmcCancel_Click()
    igTerminateReturn = 0   '0=Canceled
    sgPassValue = ""
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String

    slStr = edcTime.Text
    If (Trim$(slStr) = "") Or (Not gValidTime(slStr)) Then
       Beep
       edcTime.SetFocus
       Exit Sub
    End If
    sgAddTime = slStr
    If lbcAdvt.ListIndex < 0 Then
        Beep
        lbcAdvt.SetFocus
        Exit Sub
    End If
    If lbcCopy.ListIndex < 0 Then
        Beep
        lbcAdvt.SetFocus
        Exit Sub
    End If
    igAddAvtCode = Val(lbcAdvt.ItemData(lbcAdvt.ListIndex))
    slNameCode = tmCopyCode(lbcCopy.ListIndex).sKey
    slStr = gParseItem(slNameCode, 2, "\", slCode)
    lgAddCopyCode = Val(slCode)
    igTerminateReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         slStr                                                   *
'******************************************************************************************


    gCtrlGotFocus cmcDone
End Sub
Private Sub edcTime_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcTime_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer

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
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilAdf As Integer

    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    LlAddSpt.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'LlAddSpt.Left = LiveLog.Left + LiveLog.Width / 2 - LlAddSpt.Width / 2
    'LlAddSpt.Top = LiveLog.Top + LiveLog.Height / 2 - LlAddSpt.Height / 2
    gCenterStdAlone LlAddSpt
    edcTime.Text = Format$(gNow(), "h:mm:ssAM/PM")
    mPopAdvt
    Screen.MousePointer = vbDefault
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

    Unload LlAddSpt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tmCopyCode
    
    Set LlAddSpt = Nothing   'Remove data segment

End Sub

Private Sub lbcAdvt_Click()
    mPopCopy
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub



Private Sub mPopAdvt()
    Dim ilAdf As Integer
    Dim blAddAdvt As Boolean
    Dim ilCheck As Integer
    
    Screen.MousePointer = vbHourglass
    lbcAdvt.Clear
    For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
            '3/31/15: Don't include dormant advertisers
        If (Trim$(tgCommAdf(ilAdf).sState) <> "D") Then
            blAddAdvt = True    'vbTrue
            If ckcAdvt.Value = vbChecked Then
                blAddAdvt = False
                For ilCheck = 0 To UBound(igEventAdfCode) - 1 Step 1
                    If igEventAdfCode(ilCheck) = tgCommAdf(ilAdf).iCode Then
                        blAddAdvt = True
                        Exit For
                    End If
                Next ilCheck
            End If
            If blAddAdvt Then
                If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
                    lbcAdvt.AddItem Trim$(tgCommAdf(ilAdf).sName) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID)
                Else
                    lbcAdvt.AddItem Trim$(tgCommAdf(ilAdf).sName)
                End If
                lbcAdvt.ItemData(lbcAdvt.NewIndex) = tgCommAdf(ilAdf).iCode
            End If
        End If
    Next ilAdf
    Screen.MousePointer = vbDefault
End Sub

Public Sub mPopCopy()
    Dim ilAdvtCode As Integer
    Dim ilISCIProd As Integer
    Dim ilRet As Integer
    Dim ilInvType As Integer

    Screen.MousePointer = vbHourglass
    If lbcAdvt.ListIndex >= 0 Then
        ilAdvtCode = Val(lbcAdvt.ItemData(lbcAdvt.ListIndex))
        ilISCIProd = 2 + &H100  'Lengt Cart ISCI Product
        lbcCopy.Clear
        smCopyCodeTag = ""
        ilInvType = 0   'Active only
        If ckcCopy.Value = vbUnchecked Then
            ilInvType = 6   'Active and Purge
        End If
        ilRet = gPopCopyForAdvtBox(LiveLog, ilAdvtCode, ilISCIProd, ilInvType, lbcCopy, tmCopyCode(), smCopyCodeTag)
    Else
        lbcCopy.Clear
    End If
    Screen.MousePointer = vbDefault
End Sub
