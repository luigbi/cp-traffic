VERSION 5.00
Begin VB.Form PrgDupl 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5400
   ClientLeft      =   1215
   ClientTop       =   1470
   ClientWidth     =   6915
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
   ScaleHeight     =   5400
   ScaleWidth      =   6915
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3750
      TabIndex        =   6
      Top             =   4980
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
      Left            =   -30
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   2
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
      ScaleWidth      =   1740
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1740
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   4980
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
      Height          =   4380
      Left            =   135
      ScaleHeight     =   4320
      ScaleWidth      =   6570
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   6630
      Begin VB.PictureBox pbcLibType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   5505
         ScaleHeight     =   180
         ScaleWidth      =   1020
         TabIndex        =   8
         Top             =   4065
         Width           =   1050
      End
      Begin VB.CheckBox ckcShowVersion 
         Caption         =   "Show All Versions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3825
         TabIndex        =   7
         Top             =   4065
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   4020
         Left            =   135
         TabIndex        =   3
         Top             =   165
         Width           =   3375
      End
      Begin VB.ListBox lbcLib 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   3810
         Left            =   3855
         TabIndex        =   4
         Top             =   165
         Width           =   2625
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   4965
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "PrgDupl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Prgdupl.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PrgDupl.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Program Duplicate screen code
Option Explicit
Option Compare Text
Dim imBypassFocus As Integer
Dim imBSMode As Integer
Dim imFirstActivate As Integer
Dim imTerminate As Integer
Dim imLibType As Integer
Dim tmLibName() As SORTCODE
Dim smLibNameTag As String
Dim tmVehNameCode() As SORTCODE
Dim smVehNameCodeTag As String
Private Sub ckcShowVersion_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcShowVersion.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    mLibPop
End Sub
Private Sub cmcCancel_Click()
    igPrgDupl = False
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slCode As String
    Dim slName As String
    Dim ilRes As Integer
    If lbcLib.ListIndex < 0 Then
        ilRes = MsgBox("Select Library or Press Cancel", vbOKOnly + vbExclamation, "Warning")
        Exit Sub
    End If
    igPrgDupl = True
    slNameCode = tmLibName(lbcLib.ListIndex).sKey  'lbcLibName.List(lbcLib.ListIndex)
    ilRet = gParseItem(slNameCode, 1, "|", slName)
    sgRemLibName = slName
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    lgLibLength = Val(slCode)
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
    pbcLibType_Paint
    Me.KeyPreview = True
    PrgDupl.Refresh
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

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tmLibName
    Erase tmVehNameCode
    Set PrgDupl = Nothing   'Remove data segment

End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcLib_Click()
    If lbcLib.ListIndex >= 0 Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
    End If
End Sub
Private Sub lbcVehicle_Click()
    mLibPop
    If lbcLib.ListIndex >= 0 Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
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
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imBypassFocus = False
    imTerminate = False
    'plcScreen.Caption = "Removing Library- " & Trim$(sgRemLibName) & ":" & tgRPrg(0).sStartTime
    imBSMode = False
    mInitBox
    PrgDupl.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone PrgDupl
    imLibType = igLibType
    mVehPop
    ckcShowVersion.Value = Program!ckcShowVersion.Value
    'If igVehIndexViaPrg >= 0 Then
    '    lbcVehicle.ListIndex = igVehIndexViaPrg
    'End If
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
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLibPop                         *
'*                                                     *
'*             Created:10/11/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection library *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mLibPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slType As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilVer As Integer
    Screen.MousePointer = vbHourglass  'Wait
    lbcLib.Clear
    'lbcLibName.Clear
    'lbcLibName.Tag = ""
    ReDim tmLibName(0 To 0) As SORTCODE
    smLibNameTag = ""
    If (lbcVehicle.ListIndex < 0) Or (lbcVehicle.ListIndex > UBound(tmVehNameCode) - 1) Then 'lbcVehNameCode.ListCount - 1) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    slNameCode = tmVehNameCode(lbcVehicle.ListIndex).sKey  'lbcVehNameCode.List(lbcVehicle.ListIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mLibPopErr
    gCPErrorMsg ilRet, "mLibPop (gParseItem field 2: Vehicle)", PrgDupl
    On Error GoTo 0
    ilVefCode = Val(slCode)
    If imLibType = 3 Then 'Std Format
        slType = "F"
    ElseIf igLibType = 2 Then 'Sports
        slType = "P"
    ElseIf igLibType = 1 Then 'Special
        slType = "S"
    Else    'Regular
        slType = "R"
    End If
    'If ckcShowVersion.Value Then
    '    ilVer = ALLLIBFRONT
    'Else
        ilVer = LATESTLIB
    'End If
    'ilRet = gPopProgLibBox(PrgDupl, ilVer, slType, ilVefCode, lbcLib, lbcLibName)
    ilRet = gPopProgLibBox(PrgDupl, ilVer, slType, ilVefCode, lbcLib, tmLibName(), smLibNameTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLibPopErr
        gCPErrorMsg ilRet, "mLibPop (gPopProgLibBox: Library)", PrgDupl
        On Error GoTo 0
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mLibPopErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
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
    Unload PrgDupl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer

    'ilRet = gPopUserVehicleBox(PrgDupl, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH, lbcVehicle, lbcVehNameCode)
    ilRet = gPopUserVehicleBox(PrgDupl, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH, lbcVehicle, tmVehNameCode(), smVehNameCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", PrgDupl
        On Error GoTo 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcLibType_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(" ") Then
        If imLibType = 0 Then
            imLibType = 1
        ElseIf imLibType = 1 Then
            imLibType = 2
        ElseIf imLibType = 2 Then
            imLibType = 3
        Else
            imLibType = 0
        End If
        pbcLibType_Paint
        mLibPop
    End If
End Sub
Private Sub pbcLibType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imLibType = 0 Then
        imLibType = 1
    ElseIf imLibType = 1 Then
        imLibType = 2
    ElseIf imLibType = 2 Then
        imLibType = 3
    Else
        imLibType = 0
    End If
    pbcLibType_Paint
    mLibPop
End Sub
Private Sub pbcLibType_Paint()
    pbcLibType.Cls
    pbcLibType.CurrentX = fgBoxInsetX
    pbcLibType.CurrentY = -15 'fgBoxInsetY
    If imLibType = 0 Then
        pbcLibType.Print "Regulars"
    ElseIf imLibType = 1 Then
        pbcLibType.Print "Specials"
    ElseIf imLibType = 2 Then
        pbcLibType.Print "Sports"
    ElseIf imLibType = 3 Then
        pbcLibType.Print "Std Formats"
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Duplicate Library"
End Sub
