VERSION 5.00
Begin VB.Form RSModel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5160
   ClientLeft      =   2310
   ClientTop       =   1980
   ClientWidth     =   4710
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
   ScaleHeight     =   5160
   ScaleWidth      =   4710
   Begin VB.CheckBox ckcImpressions 
      Caption         =   "Impressions"
      Height          =   210
      Left            =   165
      TabIndex        =   5
      Top             =   4275
      Width           =   2025
   End
   Begin VB.TextBox edcPercentChg 
      Height          =   315
      Left            =   2310
      TabIndex        =   4
      Text            =   "0"
      Top             =   3915
      Width           =   840
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2655
      TabIndex        =   8
      Top             =   4710
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
      TabIndex        =   7
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
      ScaleWidth      =   4380
      TabIndex        =   0
      Top             =   -15
      Width           =   4380
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Left            =   1095
      TabIndex        =   6
      Top             =   4710
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
      Height          =   3525
      Left            =   150
      ScaleHeight     =   3465
      ScaleWidth      =   4230
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   4290
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   3390
         ItemData        =   "Rsmodel.frx":0000
         Left            =   60
         List            =   "Rsmodel.frx":0002
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   45
         Visible         =   0   'False
         Width           =   4155
      End
      Begin VB.ListBox lbcBookName 
         Appearance      =   0  'Flat
         Height          =   3390
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Visible         =   0   'False
         Width           =   4155
      End
   End
   Begin VB.Label lacPercentChg 
      Caption         =   "+/- Audience % Adjust"
      Height          =   240
      Left            =   165
      TabIndex        =   3
      Top             =   3945
      Width           =   1875
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   555
      Top             =   4680
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "RSModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rsmodel.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RSModel.Frm
'
' Release: 1.0
'
' Description:
'    This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim tmBookNameCode() As SORTCODE
Dim smBookNameCodeTag As String
Dim imForm As Integer   '16=16 buckets; 18 = 18 buckets
Public dnf_rst As ADODB.Recordset

Private Sub ckcImpressions_Click()
    bgResearchByImpressions = False
    If tgSaf(0).sHideDemoOnBR = "Y" Then
        If ckcImpressions.Value = vbChecked Then
            bgResearchByImpressions = True
        End If
    End If
End Sub

Private Sub cmcCancel_Click()
    igDnfModel = 0
    igReturn = 0
    bgResearchByImpressions = False
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    
    igDnfModel = 0
    sgPercentChg = edcPercentChg.Text
    bgResearchByImpressions = False
    If igResearchModelMethod = 0 Then
        If lbcBookName.ListIndex >= 1 Then
            slNameCode = tmBookNameCode(lbcBookName.ListIndex - 1).sKey    'lbcBookNameCode.List(lbcBookName.ListIndex - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            igDnfModel = Val(slCode)
            igReturn = 1
        Else
            igReturn = 0
        End If
        bgResearchByImpressions = False
        If tgSaf(0).sHideDemoOnBR = "Y" Then
            If ckcImpressions.Value = vbChecked Then
                bgResearchByImpressions = True
            End If
        End If
    Else
        igReturn = 0
        For ilIndex = 0 To lbcVehicle.ListCount - 1 Step 1
            If lbcVehicle.Selected(ilIndex) Then
                tgResearchAdjustVehicle(lbcVehicle.ItemData(ilIndex)).bAdjust = True
                igReturn = 1
            Else
                tgResearchAdjustVehicle(lbcVehicle.ItemData(ilIndex)).bAdjust = False
            End If
        Next ilIndex
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
    If (igWinStatus(RESEARCHLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        lbcBookName.Enabled = False
        lbcVehicle.Enabled = False
    Else
        lbcBookName.Enabled = True
        lbcVehicle.Enabled = True
    End If
'    gShowBranner
    Me.KeyPreview = True
    RSModel.Refresh
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
    On Error Resume Next
    
    dnf_rst.Close
    Erase tmBookNameCode
    Set RSModel = Nothing   'Remove data segment

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
    imForm = igDnfModel
    RSModel.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone RSModel
    'RSModel.Show
    Screen.MousePointer = vbHourglass
    If igResearchModelMethod = 0 Then
        If tgSaf(0).sHideDemoOnBR <> "Y" Then
            ckcImpressions.Visible = False
        End If
        lbcBookName.Visible = True
        mPopBookNames
        If imTerminate Then
            Exit Sub
        End If
        lbcBookName.ListIndex = 0
    Else
        lbcVehicle.Visible = True
        mPopVehicle
        ckcImpressions.Visible = False
    End If
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    gCenterModalForm RSModel
    Screen.MousePointer = vbDefault
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopBookNames                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopBookNames()
'
'   mPopBookNames
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim ilSort As Integer
    Dim ilShow As Integer
    Dim ilVefCode As Integer

    imPopReqd = False
    ilVefCode = 0
    ilSort = 1  'sort by date, then book name
    ilShow = 1  'show book name, then date
    'ilRet = gPopBookNameBox(RSModel, ilVefCode, ilSort, ilShow, lbcBookName, lbcBookNameCode)
    ilRet = gPopBookNameBox(RSModel, imForm, 0, ilVefCode, ilSort, ilShow, lbcBookName, tmBookNameCode(), smBookNameCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopBookNamesErr
        gCPErrorMsg ilRet, "mPopBookNames (gPopBookNameBox)", RSModel
        On Error GoTo 0
        lbcBookName.AddItem "[None]", 0  'Force as first item on list
        imPopReqd = True
    End If
    Exit Sub
mPopBookNamesErr:
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
    Unload RSModel
    igManUnload = NO
End Sub

Private Sub lbcBookName_Click()
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slSQLQuery As String
    
    If igResearchModelMethod = 0 Then
        If lbcBookName.ListIndex = 0 Then
            ckcImpressions.Enabled = True
        ElseIf lbcBookName.ListIndex >= 1 Then
            slNameCode = tmBookNameCode(lbcBookName.ListIndex - 1).sKey    'lbcBookNameCode.List(lbcBookName.ListIndex - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            slSQLQuery = "Select dnfSource From DNF_Demo_Rsrch_Names Where dnfCode = '" & slCode & "'"
            Set dnf_rst = gSQLSelectCall(slSQLQuery)
            If Not dnf_rst.EOF Then
                If dnf_rst!dnfSource <> "I" Then
                    ckcImpressions.Value = vbUnchecked
                Else
                    ckcImpressions.Value = vbChecked
                End If
            End If
            ckcImpressions.Enabled = False
        End If
    End If
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
    If igResearchModelMethod = 0 Then
        plcScreen.Print "Model From Research Book"
    Else
        plcScreen.Print "Adjust Research Audience by Vehicle"
    End If
End Sub

Private Sub mPopVehicle()
    Dim ilIndex As Integer
    lbcVehicle.Clear
    For ilIndex = 0 To UBound(tgResearchAdjustVehicle) - 1 Step 1
        lbcVehicle.AddItem Trim$(tgResearchAdjustVehicle(ilIndex).sVehicleName)
        lbcVehicle.ItemData(lbcVehicle.NewIndex) = ilIndex
    Next ilIndex
End Sub
