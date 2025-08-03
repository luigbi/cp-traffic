VERSION 5.00
Begin VB.Form PLogMkt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4455
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
   ScaleHeight     =   4455
   ScaleWidth      =   4185
   Begin VB.Frame frcPost 
      Caption         =   "Post by"
      Height          =   3360
      Left            =   180
      TabIndex        =   2
      Top             =   495
      Width           =   3735
      Begin VB.OptionButton rbcPost 
         Caption         =   "Rep Spot Times by Vehicle"
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   255
         TabIndex        =   9
         Top             =   2820
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.OptionButton rbcPost 
         Caption         =   "Rep Spot Times by Network Name"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   255
         TabIndex        =   8
         Top             =   2445
         Visible         =   0   'False
         Width           =   3150
      End
      Begin VB.OptionButton rbcPost 
         Caption         =   "Cluster"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   255
         TabIndex        =   7
         Top             =   2070
         Width           =   1395
      End
      Begin VB.OptionButton rbcPost 
         Caption         =   "Received"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   255
         TabIndex        =   6
         Top             =   1665
         Width           =   1425
      End
      Begin VB.OptionButton rbcPost 
         Caption         =   "Rep Counts by Advertiser"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   255
         TabIndex        =   5
         Top             =   1275
         Width           =   2565
      End
      Begin VB.OptionButton rbcPost 
         Caption         =   "Rep Counts by Vehicle"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   255
         TabIndex        =   4
         Top             =   870
         Width           =   2445
      End
      Begin VB.OptionButton rbcPost 
         Caption         =   "Air Time"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   255
         TabIndex        =   3
         Top             =   510
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   4035
      Width           =   945
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Post"
      Height          =   285
      Left            =   870
      TabIndex        =   0
      Top             =   4035
      Width           =   945
   End
   Begin VB.Label lacScreen 
      Caption         =   "Post Log Selection"
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   45
      Width           =   1980
   End
End
Attribute VB_Name = "PLogMkt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PLogmkt.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PLogMkt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Log Check screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer




Private Sub cmcCancel_Click()
    igPostType = -1
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    Screen.MousePointer = vbHourglass
    If rbcPost(0).Value Then
        igPostType = 0
    ElseIf rbcPost(1).Value Then
        igPostType = 1
    ElseIf rbcPost(2).Value Then
        igPostType = 2
    ElseIf rbcPost(3).Value Then
        igPostType = 3
    ElseIf rbcPost(4).Value Then
        igPostType = 4
    ElseIf rbcPost(5).Value Then
        igPostType = 5
    ElseIf rbcPost(6).Value Then
        igPostType = 6
    End If
    Screen.MousePointer = vbDefault
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
'    gShowBranner
    Me.KeyPreview = True
    PLogMkt.Refresh
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
    Dim ilLoop As Integer
    Dim ilMkt As Integer
    Dim blVisible(0 To 6) As Boolean
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    PLogMkt.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone PLogMkt
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    'PLogMkt.Show
    Screen.MousePointer = vbHourglass
    ilRet = gObtainVef()
    ilRet = gObtainMnfForType("H3", sgMktMnfStamp, tgMktMnf())
    For ilLoop = 0 To 6 Step 1
        rbcPost(ilLoop).Visible = False
        blVisible(ilLoop) = False
    Next ilLoop
    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Then
        '    For ilMkt = LBound(tgMktMnf) To UBound(tgMktMnf) - 1 Step 1
        '        If tgMktMnf(ilMkt).iCode = tgMVef(ilLoop).iMnfVehGp3Mkt Then
        '            If Trim$(tgMktMnf(ilMkt).sRPU) = "Y" Then
        '                rbcPost(4).Enabled = True
        '            Else
                        rbcPost(0).Visible = True
                        rbcPost(0).Enabled = True
                        blVisible(0) = True
        '            End If
        '            Exit For
        '        End If
        '    Next ilMkt
        ElseIf (tgMVef(ilLoop).sType = "R") Then
        'If (tgMVef(ilLoop).sType = "R") Then
            If UBound(tgMktMnf) > LBound(tgMktMnf) Then
                For ilMkt = LBound(tgMktMnf) To UBound(tgMktMnf) - 1 Step 1
                    If tgMktMnf(ilMkt).iCode = tgMVef(ilLoop).iMnfVehGp3Mkt Then
                        If Trim$(tgMktMnf(ilMkt).sRPU) = "Y" Then
                            rbcPost(4).Visible = True
                            rbcPost(4).Enabled = True
                            blVisible(4) = True
                        Else
                            'If tgSpf.sPostCalAff <> "D" Then
                            '    rbcPost(1).Enabled = True
                            '    rbcPost(2).Enabled = True
                            '    rbcPost(3).Enabled = True
                            'Else
                            '    ilPost1Visible = False
                            '    rbcPost(1).Visible = False
                            '    rbcPost(2).Visible = False
                            '    rbcPost(3).Visible = False
                            '    rbcPost(5).Move rbcPost(1).Left, rbcPost(1).Top
                            '    rbcPost(5).Visible = True
                            '    rbcPost(5).Enabled = True
                            'End If
                            If tgSpf.sPostCalAff <> "N" Then
                                rbcPost(1).Enabled = True
                                rbcPost(2).Enabled = True
                                rbcPost(3).Enabled = True
                                rbcPost(1).Visible = True
                                blVisible(1) = True
                                rbcPost(2).Visible = True
                                blVisible(2) = True
                                rbcPost(3).Visible = True
                                blVisible(3) = True
                            End If
                            If (Asc(tgSpf.sUsingFeatures8) And REPBYDT) = REPBYDT Then
                                rbcPost(5).Visible = True
                                rbcPost(5).Enabled = True
                                blVisible(5) = True
                                rbcPost(6).Visible = True
                                rbcPost(6).Enabled = True
                                blVisible(6) = True
                            End If
                        End If
                        Exit For
                    End If
                Next ilMkt
            Else
                'If tgSpf.sPostCalAff <> "D" Then
                If tgSpf.sPostCalAff <> "N" Then
                    rbcPost(1).Enabled = True
                    rbcPost(2).Enabled = True
                    rbcPost(3).Enabled = True
                    rbcPost(1).Visible = True
                    blVisible(1) = True
                    rbcPost(2).Visible = True
                    blVisible(2) = True
                    rbcPost(3).Visible = True
                    blVisible(3) = True
                End If
                If (Asc(tgSpf.sUsingFeatures8) And REPBYDT) = REPBYDT Then
                    rbcPost(5).Visible = True
                    rbcPost(5).Enabled = True
                    blVisible(5) = True
                    rbcPost(6).Visible = True
                    rbcPost(6).Enabled = True
                    blVisible(6) = True
                End If
            End If
        End If
    Next ilLoop
    If Not blVisible(4) Then
        If blVisible(5) Then
            If Not blVisible(3) Then
                rbcPost(5).Move rbcPost(1).Left, rbcPost(1).Top
                rbcPost(6).Move rbcPost(2).Left, rbcPost(2).Top
            Else
                rbcPost(6).Move rbcPost(5).Left, rbcPost(5).Top
                rbcPost(5).Move rbcPost(4).Left, rbcPost(4).Top
            End If
        End If
    Else
        If Not blVisible(3) Then
            rbcPost(4).Move rbcPost(1).Left, rbcPost(1).Top
            If blVisible(5) Then
                rbcPost(5).Move rbcPost(2).Left, rbcPost(2).Top
                rbcPost(6).Move rbcPost(3).Left, rbcPost(3).Top
            End If
        End If
    End If
    
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    'pbcPrinting.Move (PLogMkt.Width - pbcPrinting.Width) / 2, (PLogMkt.Height - pbcPrinting.Height) / 2
'    gCenterModalForm PLogMkt
    Screen.MousePointer = vbDefault
    Exit Sub

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
    'Close btrieve files
    igManUnload = YES
    Unload PLogMkt
    igManUnload = NO
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set PLogMkt = Nothing   'Remove data segment
End Sub
