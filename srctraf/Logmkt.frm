VERSION 5.00
Begin VB.Form LogMkt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4620
   ClientLeft      =   630
   ClientTop       =   2550
   ClientWidth     =   4185
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4620
   ScaleWidth      =   4185
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Markets"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   225
      TabIndex        =   5
      Top             =   3915
      Width           =   1350
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2580
      TabIndex        =   2
      Top             =   4215
      Width           =   945
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
      Top             =   -15
      Width           =   1140
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "C&ontinue"
      Height          =   285
      Left            =   780
      TabIndex        =   1
      Top             =   4215
      Width           =   945
   End
   Begin VB.PictureBox plcMkt 
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
      Height          =   3525
      Left            =   210
      ScaleHeight     =   3465
      ScaleWidth      =   3660
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   3720
      Begin VB.ListBox lbcMkt 
         Appearance      =   0  'Flat
         Height          =   3390
         Left            =   30
         MultiSelect     =   2  'Extended
         TabIndex        =   4
         Top             =   30
         Width           =   3600
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   285
      Top             =   4200
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "LogMkt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Logmkt.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: LogMkt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Log Check screen code
Option Explicit
Option Compare Text
'Btrieve files
Dim tmVehGp3Code() As SORTCODE
Dim smVehGp3CodeTag As String
'Program library dates Field Areas
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imAllClicked As Integer
Dim imSetAll As Integer


Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilValue As Integer
    Dim llRet As Long
    Dim llRg As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        If lbcMkt.ListCount > 0 Then
            llRg = CLng(lbcMkt.ListCount - 1) * &H10000 + 0
            'llRet = SendMessageByNum(lbcLines.hwnd, &H400 + 28, ilValue, llRg)
            llRet = SendMessageByNum(lbcMkt.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        End If
        imAllClicked = False
    End If
End Sub

Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Screen.MousePointer = vbHourglass
    For ilLoop = 0 To lbcMkt.ListCount - 1 Step 1
        If lbcMkt.Selected(ilLoop) Then
            slNameCode = tmVehGp3Code(ilLoop).sKey    'lbcVehCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            igMktCode(UBound(igMktCode)) = Val(slCode)
            ReDim Preserve igMktCode(0 To UBound(igMktCode) + 1) As Integer
        End If
    Next ilLoop
    Screen.MousePointer = vbDefault
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub Form_Activate()
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String

    If lbcMkt.ListCount = 1 Then
        slNameCode = tmVehGp3Code(0).sKey    'lbcVehCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        igMktCode(UBound(igMktCode)) = Val(slCode)
        ReDim Preserve igMktCode(0 To UBound(igMktCode) + 1) As Integer
        mTerminate
        Exit Sub
    End If
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner
    Me.KeyPreview = True
    LogMkt.Refresh
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
    Set LogMkt = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    LogMkt.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone LogMkt
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    imAllClicked = False
    imSetAll = True
    ReDim igMktCode(0 To 0) As Integer
    'Market
    mPopulate
    'LogMkt.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    'pbcPrinting.Move (LogMkt.Width - pbcPrinting.Width) / 2, (LogMkt.Height - pbcPrinting.Height) / 2
'    gCenterModalForm LogMkt
    Screen.MousePointer = vbDefault
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Vehicle Group Code    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mVehGpPop
'   Where:
'
    Dim ilRet As Integer
    ilRet = gPopMnfPlusFieldsBox(LogMkt, lbcMkt, tmVehGp3Code(), smVehGp3CodeTag, "H3")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopMnfPlusFieldsBox)", LogMkt
        On Error GoTo 0
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
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
    'Close btrieve files
    igManUnload = YES
    Unload LogMkt
    igManUnload = NO
End Sub

Private Sub lbcMkt_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
    End If
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Markets"
End Sub
