VERSION 5.00
Begin VB.Form BkoutPool 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5760
   ClientLeft      =   2640
   ClientTop       =   2610
   ClientWidth     =   10095
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
   ScaleHeight     =   5760
   ScaleWidth      =   10095
   Begin VB.PictureBox pbcDeactiveToActive 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   180
      Left            =   6465
      Picture         =   "BkoutPool.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3135
      Width           =   180
   End
   Begin VB.PictureBox pbcActiveToDeactive 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   180
      Left            =   7035
      Picture         =   "BkoutPool.frx":00DA
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2085
      Width           =   180
   End
   Begin VB.CommandButton cmcActiveToDeactive 
      Appearance      =   0  'Flat
      Caption         =   "M&ove   "
      Height          =   300
      Left            =   6345
      TabIndex        =   20
      Top             =   2055
      Width           =   945
   End
   Begin VB.CommandButton cmcDeactiveToActive 
      Appearance      =   0  'Flat
      Caption         =   "    Mo&ve"
      Height          =   300
      Left            =   6345
      TabIndex        =   19
      Top             =   3105
      Width           =   945
   End
   Begin VB.PictureBox pbcAvailToActive 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   180
      Left            =   3315
      Picture         =   "BkoutPool.frx":01B4
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2115
      Width           =   180
   End
   Begin VB.PictureBox pbcActiveToAvail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   180
      Left            =   2745
      Picture         =   "BkoutPool.frx":028E
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3165
      Width           =   180
   End
   Begin VB.CommandButton cmcAvailToActive 
      Appearance      =   0  'Flat
      Caption         =   "M&ove   "
      Height          =   300
      Left            =   2625
      TabIndex        =   15
      Top             =   2055
      Width           =   945
   End
   Begin VB.CommandButton cmcActiveToAvail 
      Appearance      =   0  'Flat
      Caption         =   "    Mo&ve"
      Height          =   300
      Left            =   2625
      TabIndex        =   14
      Top             =   3105
      Width           =   945
   End
   Begin VB.PictureBox plcDeactivated 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   7560
      ScaleHeight     =   4110
      ScaleWidth      =   2265
      TabIndex        =   10
      Top             =   795
      Width           =   2325
      Begin VB.ListBox lbcDeactivated 
         Appearance      =   0  'Flat
         Height          =   4020
         ItemData        =   "BkoutPool.frx":0368
         Left            =   30
         List            =   "BkoutPool.frx":036A
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   45
         Width           =   2175
      End
   End
   Begin VB.PictureBox plcActive 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   3795
      ScaleHeight     =   4110
      ScaleWidth      =   2250
      TabIndex        =   8
      Top             =   780
      Width           =   2310
      Begin VB.ListBox lbcActive 
         Appearance      =   0  'Flat
         Height          =   4020
         ItemData        =   "BkoutPool.frx":036C
         Left            =   30
         List            =   "BkoutPool.frx":036E
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   45
         Width           =   2175
      End
   End
   Begin VB.PictureBox plcAvailable 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   120
      ScaleHeight     =   4080
      ScaleWidth      =   2250
      TabIndex        =   6
      Top             =   795
      Width           =   2310
      Begin VB.ListBox lbcAvailable 
         Appearance      =   0  'Flat
         Height          =   4020
         ItemData        =   "BkoutPool.frx":0370
         Left            =   30
         List            =   "BkoutPool.frx":0372
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   30
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5685
      TabIndex        =   2
      Top             =   5340
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1515
      Width           =   120
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   75
      ScaleHeight     =   270
      ScaleWidth      =   3195
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   3195
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   3495
      TabIndex        =   1
      Top             =   5340
      Width           =   1050
   End
   Begin VB.Label lacDeavtivated 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Deactivated Advertisers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   7635
      TabIndex        =   16
      Top             =   450
      Width           =   2175
   End
   Begin VB.Label lacAvailable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Available Advertisers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   135
      TabIndex        =   5
      Top             =   450
      Width           =   2235
   End
   Begin VB.Label lacActive 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Active Advertisers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   3720
      TabIndex        =   4
      Top             =   450
      Width           =   2625
   End
End
Attribute VB_Name = "BkoutPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of BkoutPool.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BkoutPool.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract number screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imUpdateAllowed As Integer
Dim bmChgd As Boolean
Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor


Dim rst_adf As ADODB.Recordset


Private Sub cmcActiveToAvail_Click()
    mMoveName lbcActive, lbcAvailable
End Sub

Private Sub cmcActiveToDeactive_Click()
    mMoveName lbcActive, lbcDeactivated
End Sub

Private Sub cmcAvailToActive_Click()
    mMoveName lbcAvailable, lbcActive
End Sub

Private Sub cmcCancel_Click()
    igTerminateReturn = 0   '0=Canceled
    sgPassValue = ""
    mTerminate
End Sub

Private Sub cmcDeactiveToActive_Click()
    mMoveName lbcDeactivated, lbcActive
End Sub

Private Sub cmcDone_Click()
    Dim slStr As String
    Dim slActive As String
    Dim slUnassigned As String
    Dim slDeactivated As String
    Dim slItemData As String
    Dim llItemData As Long
    Dim slSQLQuery As String
    Dim llRet As Long
    Dim ilLoop As Integer
    Dim ilRes As Integer
    
    If lbcActive.ListCount = 0 Then
        bmChgd = False
    End If
    If lbcActive.ListCount < 4 Then
        MsgBox "To avoid Back-To-Back Spots airing with blackout commercial, set more than three Advertisers as Active", vbOK, "Update Stopped"
        Exit Sub
    End If
    If bmChgd Then
        ilRes = MsgBox("Save Changes?", vbYesNoCancel + vbQuestion, "Update")
        If ilRes = vbCancel Then
            Exit Sub
        End If
    
        If ilRes = vbYes Then
            slSQLQuery = "Update adf_Advertisers Set adfBkoutPoolStatus = 'N'"
            llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            slActive = ""
            slUnassigned = ""
            For ilLoop = 0 To lbcActive.ListCount - 1 Step 1
                llItemData = Val(lbcActive.ItemData(ilLoop))
                If llItemData > 50000 Then
                    llItemData = llItemData - 50000
                    If slUnassigned = "" Then
                        slUnassigned = str(llItemData)
                    Else
                        slUnassigned = slUnassigned & "," & str(llItemData)
                    End If
                Else
                    If slActive = "" Then
                        slActive = str(llItemData)
                    Else
                        slActive = slActive & "," & str(llItemData)
                    End If
                End If
            Next ilLoop
            If slUnassigned <> "" Then
                slSQLQuery = "Update adf_Advertisers Set adfBkoutPoolStatus = 'U'"
                slSQLQuery = slSQLQuery & " Where adfCode In (" & slUnassigned & ")"
                llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            End If
            If slActive <> "" Then
                slSQLQuery = "Update adf_Advertisers Set adfBkoutPoolStatus = 'A'"
                slSQLQuery = slSQLQuery & " Where adfCode In (" & slActive & ")"
                llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            End If
            slDeactivated = ""
            For ilLoop = 0 To lbcDeactivated.ListCount - 1 Step 1
                llItemData = Val(lbcDeactivated.ItemData(ilLoop))
                If slDeactivated = "" Then
                    slDeactivated = str(llItemData)
                Else
                    slDeactivated = slDeactivated & "," & str(llItemData)
                End If
            Next ilLoop
            If slDeactivated <> "" Then
                slSQLQuery = "Update adf_Advertisers Set adfBkoutPoolStatus = 'D'"
                slSQLQuery = slSQLQuery & " Where adfCode In (" & slDeactivated & ")"
                llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            End If
            gFileChgdUpdate "adf.btr", False
        End If
    End If
    igTerminateReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    Dim ilRet As Integer
    Dim slStr As String

    gCtrlGotFocus cmcDone
End Sub


Private Sub Form_Activate()
    If (igWinStatus(COPYJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
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
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        'fmAdjFactorW = (((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        fmAdjFactorW = (((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        If fmAdjFactorW < 1# Then
            fmAdjFactorW = 1#
        Else
            'Me.Width = ((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
            Me.Width = ((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        End If
        'fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        'Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
        fmAdjFactorH = 1#
    End If
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
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    bmChgd = False
    plcAvailable.Width = (Me.Width - 2 * cmcAvailToActive.Width - 7 * 120) / 3 'fmAdjFactorW * plcAvailable.Width
    lbcAvailable.Width = plcAvailable.Width - 120    'fmAdjFactorW * lbcAvailable.Width
    lacAvailable.Width = plcAvailable.Width
    lacAvailable.Left = plcAvailable.Left
    cmcAvailToActive.Left = plcAvailable.Left + plcAvailable.Width + 120
    pbcAvailToActive.Left = cmcAvailToActive.Left + cmcAvailToActive.Width - (3 * pbcAvailToActive.Width) / 2
    pbcAvailToActive.Top = cmcAvailToActive.Top + 60
    cmcActiveToAvail.Left = cmcAvailToActive.Left
    pbcActiveToAvail.Left = cmcActiveToAvail.Left + pbcActiveToAvail.Width
    pbcActiveToAvail.Top = cmcActiveToAvail.Top + 60
    plcActive.Width = plcAvailable.Width    'fmAdjFactorW * plcActive.Width
    lbcActive.Width = lbcAvailable.Width    'fmAdjFactorW * lbcActive.Width
    plcActive.Left = cmcAvailToActive.Left + cmcAvailToActive.Width + 120
    lacActive.Width = plcActive.Width
    lacActive.Left = plcActive.Left
    cmcActiveToDeactive.Left = plcActive.Left + plcActive.Width + 120
    pbcActiveToDeactive.Left = cmcActiveToDeactive.Left + cmcActiveToDeactive.Width - (3 * pbcActiveToDeactive.Width) / 2
    pbcActiveToDeactive.Top = cmcActiveToDeactive.Top + 60
    cmcDeactiveToActive.Left = cmcActiveToDeactive.Left
    pbcDeactiveToActive.Left = cmcDeactiveToActive.Left + pbcDeactiveToActive.Width
    pbcDeactiveToActive.Top = cmcDeactiveToActive.Top + 60
    plcDeactivated.Width = plcAvailable.Width    'fmAdjFactorW * plcDeactivated.Width
    lbcDeactivated.Width = lbcAvailable.Width    'fmAdjFactorW * lbcDeactivated.Width
    plcDeactivated.Left = cmcActiveToDeactive.Left + cmcActiveToDeactive.Width + 120
    lacDeavtivated.Width = plcDeactivated.Width
    lacDeavtivated.Left = plcDeactivated.Left
    
    cmcDone.Left = Me.Width / 2 - 2 * cmcDone.Width
    cmcCancel.Left = Me.Width / 2 + cmcCancel.Width
    mPopAdvt
    BkoutPool.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterModalForm BkoutPool
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
    Unload BkoutPool
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set BkoutPool = Nothing   'Remove data segment
End Sub

Private Sub lbcActive_Click()
    mSetCommands
End Sub

Private Sub lbcAvailable_Click()
    mSetCommands
End Sub

Private Sub lbcDeactivated_Click()
    mSetCommands
End Sub

Private Sub pbcActiveToAvail_Click()
    mMoveName lbcActive, lbcAvailable
End Sub

Private Sub pbcActiveToDeactive_Click()
    mMoveName lbcActive, lbcDeactivated
End Sub

Private Sub pbcAvailToActive_Click()
    mMoveName lbcAvailable, lbcActive
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcDeactiveToActive_Click()
    mMoveName lbcDeactivated, lbcActive
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Blackout Advertiser Pool"
End Sub

Private Sub mPopAdvt()
    Dim slSQLQuery As String
    lbcAvailable.Clear
    slSQLQuery = "Select adfCode, adfName from Adf_Advertisers"
    slSQLQuery = slSQLQuery & " Where adfState <> 'D'"
    slSQLQuery = slSQLQuery & " And (adfBkoutPoolStatus = 'N'"
    slSQLQuery = slSQLQuery & " Or adfBkoutPoolStatus = ''" & ")"
    slSQLQuery = slSQLQuery & " Order by adfName"
    Set rst_adf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_adf.EOF
        lbcAvailable.AddItem Trim$(rst_adf!adfName)
        lbcAvailable.ItemData(lbcAvailable.NewIndex) = rst_adf!adfCode + 50000
        rst_adf.MoveNext
    Loop
    
    lbcActive.Clear
    slSQLQuery = "Select adfCode, adfName, adfBkoutPoolStatus from Adf_Advertisers"
    slSQLQuery = slSQLQuery & " Where "
    slSQLQuery = slSQLQuery & " (adfBkoutPoolStatus = 'A'"
    slSQLQuery = slSQLQuery & " Or adfBkoutPoolStatus = 'U'" & ")"
    slSQLQuery = slSQLQuery & " Order by adfName"
    Set rst_adf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_adf.EOF
        lbcActive.AddItem Trim$(rst_adf!adfName)
        If rst_adf!adfBkoutPoolStatus <> "U" Then
            lbcActive.ItemData(lbcActive.NewIndex) = rst_adf!adfCode
        Else
            lbcActive.ItemData(lbcActive.NewIndex) = rst_adf!adfCode + 50000
        End If
        rst_adf.MoveNext
    Loop
    
    lbcDeactivated.Clear
    slSQLQuery = "Select adfCode, adfName from Adf_Advertisers"
    slSQLQuery = slSQLQuery & " Where "
    slSQLQuery = slSQLQuery & " adfBkoutPoolStatus = 'D'"
    slSQLQuery = slSQLQuery & " Order by adfName"
    Set rst_adf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_adf.EOF
        lbcDeactivated.AddItem Trim$(rst_adf!adfName)
        lbcDeactivated.ItemData(lbcDeactivated.NewIndex) = rst_adf!adfCode
        rst_adf.MoveNext
    Loop
    
End Sub

Private Sub mSetCommands()
    cmcAvailToActive.Enabled = False
    pbcAvailToActive.Enabled = False
    
    cmcActiveToAvail.Enabled = False
    pbcActiveToAvail.Enabled = False
    
    cmcActiveToDeactive.Enabled = False
    pbcActiveToDeactive.Enabled = False
    
    cmcDeactiveToActive.Enabled = False
    pbcDeactiveToActive.Enabled = False
    
    If lbcAvailable.SelCount > 0 Then
        cmcAvailToActive.Enabled = True
        pbcAvailToActive.Enabled = True
    End If
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If lbcActive.SelCount > 0 Then
        If lbcActive.ItemData(lbcActive.ListIndex) > 50000 Then
            cmcActiveToAvail.Enabled = True
            pbcActiveToAvail.Enabled = True
        Else
            cmcActiveToDeactive.Enabled = True
            pbcActiveToDeactive.Enabled = True
        End If
    End If
    
    If lbcDeactivated.SelCount > 0 Then
        cmcDeactiveToActive.Enabled = True
        pbcDeactiveToActive.Enabled = True
    End If
End Sub

Private Sub mMoveName(lbcFrom As ListBox, lbcTo As ListBox)
    Dim slName As String
    Dim slItemData As String
    
    If lbcFrom.ListIndex >= 0 Then
        bmChgd = True
        slName = lbcFrom.List(lbcFrom.ListIndex)
        slItemData = lbcFrom.ItemData(lbcFrom.ListIndex)
        lbcFrom.RemoveItem lbcFrom.ListIndex
        lbcTo.AddItem slName
        lbcTo.ItemData(lbcTo.NewIndex) = slItemData
    End If
    mSetCommands


End Sub
