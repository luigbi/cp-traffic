VERSION 5.00
Begin VB.Form POModel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3540
   ClientLeft      =   2310
   ClientTop       =   1980
   ClientWidth     =   4740
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
   ScaleHeight     =   3540
   ScaleWidth      =   4740
   Begin VB.ListBox lbcSelectionCode 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "POModel.frx":0000
      Left            =   3180
      List            =   "POModel.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2475
      TabIndex        =   5
      Top             =   3195
      Width           =   945
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
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   1770
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   1920
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Left            =   1245
      TabIndex        =   3
      Top             =   3195
      Width           =   945
   End
   Begin VB.PictureBox plcModel 
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
      Height          =   2670
      Left            =   150
      ScaleHeight     =   2610
      ScaleWidth      =   4260
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   4320
      Begin VB.ListBox lbcSelection 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   4200
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   3165
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "POModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of POModel.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: POModel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Private Sub cmcCancel_Click()
    igPOCodeXFre = 0
    igPOReturn = 0
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilParseNo As Integer
    Dim slStr As String

    igPOCodeXFre = 0
    igPOReturn = 1
    If lbcSelection.ListIndex >= 0 Then
        slNameCode = lbcSelectionCode.List(lbcSelection.ListIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        igPOCodeXFre = Val(slCode)
        ilParseNo = 2
        ilRet = gParseItem(lbcSelection.List(lbcSelection.ListIndex), ilParseNo, "/", slStr)
        Do While ilRet = 0 'Then 'parse delimeter found
            slStr = Trim$(slStr)
            ilRet = gParseItem(slStr, 1, " ", slStr)
            slStr = Trim$(slStr)
            If slStr = "Direct" Then 'Direct advertiser
                igPOReturn = 2
                Exit Do
            End If
            ilParseNo = ilParseNo + 1
            ilRet = gParseItem(lbcSelection.List(lbcSelection.ListIndex), ilParseNo, "/", slStr)
        Loop

    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(COLLECTIONSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        lbcSelection.Enabled = False
    Else
        lbcSelection.Enabled = True
    End If
'    gShowBranner
    Me.KeyPreview = True
    POModel.Refresh
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
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set POModel = Nothing   'Remove data segment
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
    POModel.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone POModel
    'POModel.Show
    Screen.MousePointer = vbHourglass
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    lbcSelection.ListIndex = -1
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    gCenterModalForm POModel
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
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status

    imPopReqd = False
    'ilRet = gPopUserVehicleBox(Vehicle, ilFilter, cbcSelect, Traffic!lbcVehicle)
    ilRet = gPopAgyCollectBox(POModel, "A", lbcSelection, lbcSelectionCode)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopAgyCollectBox)", POModel
        On Error GoTo 0
        imPopReqd = True
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Unload POModel
    igManUnload = NO
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
    plcScreen.Print "Transfer PO to:"
End Sub
