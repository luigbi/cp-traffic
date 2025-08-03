VERSION 5.00
Begin VB.Form CRBPkg 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5565
   ClientLeft      =   2550
   ClientTop       =   2730
   ClientWidth     =   7650
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
   ScaleHeight     =   5565
   ScaleWidth      =   7650
   Begin VB.Frame FrcPackageType 
      Caption         =   "Package Type"
      Height          =   660
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   7395
      Begin VB.TextBox edcVirtualSpots 
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
         Left            =   5295
         MaxLength       =   5
         TabIndex        =   5
         Top             =   240
         Width           =   1065
      End
      Begin VB.OptionButton rbcPackageType 
         Caption         =   "Virtual"
         Height          =   225
         Index           =   1
         Left            =   1155
         TabIndex        =   3
         Top             =   300
         Width           =   870
      End
      Begin VB.OptionButton rbcPackageType 
         Caption         =   "Real"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   870
      End
      Begin VB.Label lacVirtualSpots 
         Caption         =   "Number of Virtual Spots per Week"
         Height          =   195
         Left            =   2250
         TabIndex        =   4
         Top             =   300
         Width           =   2880
      End
   End
   Begin VB.Frame FrcPackageBy 
      Caption         =   "Package by"
      Height          =   3900
      Left            =   120
      TabIndex        =   6
      Top             =   1050
      Width           =   7395
      Begin VB.PictureBox pbcLbcVehicle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1890
         Left            =   255
         ScaleHeight     =   1890
         ScaleWidth      =   450
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1845
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox edcDaypartName 
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
         Left            =   3855
         MaxLength       =   20
         TabIndex        =   17
         Top             =   1500
         Width           =   3270
      End
      Begin VB.TextBox edcVehicleName 
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
         Left            =   180
         MaxLength       =   40
         TabIndex        =   15
         Top             =   1500
         Width           =   3270
      End
      Begin VB.Frame frcCreateComment 
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   3570
         TabIndex        =   10
         Top             =   255
         Visible         =   0   'False
         Width           =   3675
         Begin VB.OptionButton rbcCreateComment 
            Caption         =   "Yes"
            Height          =   210
            Index           =   0
            Left            =   -15
            TabIndex        =   13
            Top             =   210
            Width           =   660
         End
         Begin VB.OptionButton rbcCreateComment 
            Caption         =   "No"
            Height          =   210
            Index           =   1
            Left            =   825
            TabIndex        =   12
            Top             =   210
            Value           =   -1  'True
            Width           =   585
         End
         Begin VB.Label lacCreateComment 
            Caption         =   "Place Vehicle Names into ""Other Comment"""
            Height          =   195
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   3750
         End
      End
      Begin VB.ListBox lbcVehicle 
         Height          =   1950
         ItemData        =   "CRBPkg.frx":0000
         Left            =   180
         List            =   "CRBPkg.frx":0002
         TabIndex        =   18
         Top             =   1815
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.ListBox lbcDaypart 
         Height          =   1950
         ItemData        =   "CRBPkg.frx":0004
         Left            =   3855
         List            =   "CRBPkg.frx":0006
         TabIndex        =   19
         Top             =   1830
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.OptionButton rbcPackageBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Create One Inclusive Package"
         Height          =   210
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   960
         Width           =   3105
      End
      Begin VB.OptionButton rbcPackageBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Create Package for each Vehicle"
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   660
         Width           =   3105
      End
      Begin VB.OptionButton rbcPackageBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Create Package for each Daypart"
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   345
         Width           =   3105
      End
      Begin VB.Label lacDaypartName 
         Caption         =   "Enter New Name OR Select from List:"
         Height          =   210
         Left            =   3855
         TabIndex        =   16
         Top             =   1290
         Width           =   3135
      End
      Begin VB.Label lacVehicleName 
         Caption         =   "Enter New Name OR Select from List:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   1290
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4065
      TabIndex        =   21
      Top             =   5145
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1515
      Width           =   120
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2475
      TabIndex        =   20
      Top             =   5145
      Width           =   1050
   End
   Begin VB.Label lacScreen 
      Caption         =   "Rate Bar Package"
      Height          =   210
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   2820
   End
End
Attribute VB_Name = "CRBPkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CRBPkg.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CCancel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract revision number increment screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imUpdateAllowed As Integer
Dim imINPBCPaint As Integer


Private Sub cmcCancel_Click()
    igTerminateReturn = 0   '0=Canceled
    sgPassValue = ""
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slStr As String
    Dim ilUpdateAllowed As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer

    If (rbcPackageType(0).Value = False) And (rbcPackageType(1).Value = False) Then
        Exit Sub
    End If
    If (rbcCreateComment(0).Value = False) And (rbcCreateComment(1).Value = False) Then
        Exit Sub
    End If
    If (rbcPackageBy(0).Value = False) And (rbcPackageBy(1).Value = False) And (rbcPackageBy(2).Value = False) Then
        Exit Sub
    End If
    If (rbcPackageBy(0).Value) Or (rbcPackageBy(2).Value) Then
        If (lbcVehicle.ListIndex < 0) And (Trim$(edcVehicleName.Text) = "") Then
            Exit Sub
        End If
        '1/19/10:  Check that the name is not a package name
        If lbcVehicle.ListIndex < 0 Then
            slStr = Trim$(edcVehicleName.Text)
        Else
            slStr = lbcVehicle.List(lbcVehicle.ListIndex)
        End If
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If StrComp(Trim$(tgMVef(ilVef).sName), slStr, vbTextCompare) = 0 Then
                If tgMVef(ilVef).sType = "G" Then
                    ilRet = MsgBox("Enter a different Package Vehicle Name as a Sport Vehicle Name is not allowed", vbOKOnly + vbExclamation, "Incomplete")
                    Exit Sub
                End If
            End If
        Next ilVef
    End If
    If rbcPackageType(1).Value Then
        If Trim$(edcVirtualSpots.Text) = "" Then
            Exit Sub
        Else
            If Val(Trim$(edcVirtualSpots.Text)) <= 0 Then
                Exit Sub
            End If
        End If
    End If
    If (rbcPackageBy(1).Value) Or (rbcPackageBy(2).Value) Then
        If (lbcDaypart.ListIndex < 0) And (Trim$(edcDaypartName.Text) = "") Then
            Exit Sub
        End If
        If lbcDaypart.ListIndex < 0 Then
            slStr = Trim$(edcDaypartName.Text)
            gFindMatch slStr, 0, lbcDaypart
            If gLastFound(lbcDaypart) < 0 Then
                igDPNameCallSource = CALLSOURCERATECARD
                sgDPName = slStr
                ilUpdateAllowed = imUpdateAllowed
                igVefCode = -1
                Daypart.Show vbModal
                imUpdateAllowed = ilUpdateAllowed
                gShowBranner imUpdateAllowed
                If igDPAltered Then
                    ilRet = gObtainRcfRifRdf()
                    mPkDPPop lbcDaypart
                End If
                If igDPNameCallSource = CALLDONE Then  'Done
                    igDPNameCallSource = CALLNONE
                    gFindMatch sgDPName, 0, lbcDaypart
                    sgDPName = ""
                    If gLastFound(lbcDaypart) >= 0 Then
                        lbcDaypart.ListIndex = gLastFound(lbcDaypart)
                    Else
                        lbcDaypart.ListIndex = -1
                        edcDaypartName.SetFocus
                        Exit Sub
                    End If
                End If
                If igDPNameCallSource = CALLCANCELLED Then  'Cancelled
                    igDPNameCallSource = CALLNONE
                    sgDPName = ""
                    edcDaypartName.SetFocus
                    Exit Sub
                End If
                If igDPNameCallSource = CALLTERMINATED Then
                    igDPNameCallSource = CALLNONE
                    sgDPName = ""
                    edcDaypartName.SetFocus
                    Exit Sub
                End If
            Else
                lbcDaypart.ListIndex = gLastFound(lbcDaypart)
            End If
        End If
    End If
    sgPassValue = ""
    If rbcPackageType(0).Value Then
        sgPassValue = "0"
    Else
        sgPassValue = "1"
    End If
    If (rbcCreateComment(0).Value) And (rbcPackageBy(0).Value) Then
        sgPassValue = sgPassValue & "|0"
    Else
        sgPassValue = sgPassValue & "|1"
    End If
    If (rbcPackageBy(0).Value) Then
        sgPassValue = sgPassValue & "|0"
    End If
    If (rbcPackageBy(1).Value) Then
        sgPassValue = sgPassValue & "|1"
    End If
    If (rbcPackageBy(2).Value) Then
        sgPassValue = sgPassValue & "|2"
    End If
    If (rbcPackageBy(0).Value) Or (rbcPackageBy(2).Value) Then
        If lbcVehicle.ListIndex < 0 Then
            sgPassValue = sgPassValue & "|" & Trim$(edcVehicleName.Text)
        Else
            sgPassValue = sgPassValue & "|" & lbcVehicle.List(lbcVehicle.ListIndex)
        End If
    Else
        sgPassValue = sgPassValue & "|"
    End If
    If (rbcPackageBy(1).Value) Or (rbcPackageBy(2).Value) Then
        sgPassValue = sgPassValue & "|" & lbcDaypart.List(lbcDaypart.ListIndex)
    Else
        sgPassValue = sgPassValue & "|"
    End If
    If rbcPackageType(0).Value Then
        sgPassValue = sgPassValue & "|" & "0"
    Else
        sgPassValue = sgPassValue & "|" & Trim$(edcVirtualSpots.Text)
    End If
    igTerminateReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
End Sub

Private Sub edcDaypartName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcVehicleName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcVirtualSpots_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcVirtualSpots_KeyPress(KeyAscii As Integer)
    Dim slStr As String

    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcVirtualSpots.Text
    slStr = Left$(slStr, edcVirtualSpots.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcVirtualSpots.SelStart - edcVirtualSpots.SelLength)
    If gCompNumberStr(slStr, "99999") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
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
'
'   mInit
'   Where:
'
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    CRBPkg.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterModalForm CRBPkg
    If tgSpf.sCPkAired = "Y" Then
        rbcPackageType(0).Enabled = True
    Else
        rbcPackageType(0).Enabled = False
    End If
    If tgSpf.sCPkOrdered = "Y" Then
        rbcPackageType(1).Enabled = True
        edcVirtualSpots.Text = "1"
    Else
        rbcPackageType(1).Enabled = False
    End If
    If (tgSpf.sCPkAired = "Y") And (tgSpf.sCPkOrdered <> "Y") Then
        rbcPackageType(0).Value = True
    End If
    If (tgSpf.sCPkAired <> "Y") And (tgSpf.sCPkOrdered = "Y") Then
        rbcPackageType(1).Value = True
    End If
    mPkDPPop lbcDaypart
    mPkVehPop lbcVehicle, False, True
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
    Unload CRBPkg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CRBPkg = Nothing   'Remove data segment
End Sub

Private Sub lbcVehicle_Click()
    pbcLbcVehicle_Paint
End Sub

Private Sub lbcVehicle_Scroll()
    pbcLbcVehicle_Paint
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcLbcVehicle_Paint()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilField                       slFields                                                *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLinesEnd As Integer
    Dim llWidth As Long
    Dim llFgColor As Long

    If imINPBCPaint Then
        Exit Sub
    End If
    imINPBCPaint = True
    pbcLbcVehicle.Move lbcVehicle.Left + 15, lbcVehicle.Top + 15, pbcLbcVehicle.Width - 30, lbcVehicle.Height - 30 '2115, 1560
    ilLinesEnd = lbcVehicle.TopIndex + lbcVehicle.Height \ fgListHtArial825
    If ilLinesEnd > lbcVehicle.ListCount Then
        ilLinesEnd = lbcVehicle.ListCount
    End If
    If lbcVehicle.ListCount <= lbcVehicle.Height \ fgListHtArial825 Then
        llWidth = lbcVehicle.Width - 30
    Else
        llWidth = lbcVehicle.Width - igScrollBarWidth - 30
    End If
    pbcLbcVehicle.Width = llWidth
    pbcLbcVehicle.Cls
    llFgColor = pbcLbcVehicle.ForeColor
    For ilLoop = lbcVehicle.TopIndex To ilLinesEnd - 1 Step 1
        pbcLbcVehicle.ForeColor = llFgColor
        If lbcVehicle.Selected(ilLoop) Then
            gPaintArea pbcLbcVehicle, CSng(0), CSng((ilLoop - lbcVehicle.TopIndex) * fgListHtArial825), CSng(pbcLbcVehicle.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
            pbcLbcVehicle.ForeColor = vbWhite
        Else
            slStr = tgPackVeh(ilLoop).sSortChar
            If Left$(slStr, 1) = "A" Then
                pbcLbcVehicle.ForeColor = vbBlue
            Else
                pbcLbcVehicle.ForeColor = vbBlack
            End If
        End If
        pbcLbcVehicle.CurrentX = 15
        pbcLbcVehicle.CurrentY = (ilLoop - lbcVehicle.TopIndex) * fgListHtArial825 + 15
        slStr = lbcVehicle.List(ilLoop)
        gAdjShowLen pbcLbcVehicle, slStr, pbcLbcVehicle.Width
        pbcLbcVehicle.Print slStr
        pbcLbcVehicle.ForeColor = llFgColor
    Next ilLoop
    imINPBCPaint = False
End Sub

Private Sub rbcPackageBy_Click(Index As Integer)
    imINPBCPaint = False
    If rbcPackageBy(Index).Value Then
        Select Case Index
            Case 0
                lacVehicleName.Left = 180
                edcVehicleName.Left = 180
                lacVehicleName.Visible = True
                edcVehicleName.Visible = True
                lacDaypartName.Visible = False
                edcDaypartName.Visible = False
                lbcVehicle.Left = 180
                lbcVehicle.Width = 5745
                lbcVehicle.Visible = True
                pbcLbcVehicle.Visible = True
                lbcDaypart.Visible = False
                frcCreateComment.Visible = True
            Case 1
                lacDaypartName.Left = 180
                edcDaypartName.Left = 180
                lacDaypartName.Visible = True
                edcDaypartName.Visible = True
                lacVehicleName.Visible = False
                edcVehicleName.Visible = False
                lbcDaypart.Left = 180
                lbcDaypart.Width = 5745
                lbcDaypart.Visible = True
                lbcVehicle.Visible = False
                pbcLbcVehicle.Visible = False
                frcCreateComment.Visible = False
            Case 2
                lacVehicleName.Left = 180
                edcVehicleName.Left = 180
                lacVehicleName.Visible = True
                edcVehicleName.Visible = True
                lacDaypartName.Left = 3855
                edcDaypartName.Left = 3855
                lacDaypartName.Visible = True
                edcDaypartName.Visible = True
                lbcVehicle.Left = 180
                lbcVehicle.Width = 3315
                lbcDaypart.Left = 3855
                lbcDaypart.Width = 3315
                lbcVehicle.Visible = True
                pbcLbcVehicle.Visible = True
                lbcDaypart.Visible = True
                frcCreateComment.Visible = False
        End Select
    End If
End Sub

Private Sub rbcPackageType_Click(Index As Integer)
    If rbcPackageType(Index).Value Then
        If Index = 0 Then
            lacVirtualSpots.Enabled = False
            edcVirtualSpots.Enabled = False
        Else
            lacVirtualSpots.Enabled = True
            edcVirtualSpots.Enabled = True
        End If
    End If
End Sub
