VERSION 5.00
Begin VB.Form RCSplit 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2325
   ClientLeft      =   2145
   ClientTop       =   2535
   ClientWidth     =   5595
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
   ScaleHeight     =   2325
   ScaleWidth      =   5595
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Top             =   1965
      Width           =   1050
   End
   Begin VB.PictureBox plcAvg 
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
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   5310
      TabIndex        =   6
      Top             =   1395
      Width           =   5370
      Begin VB.CheckBox ckcAvg 
         Caption         =   "Set Weekly Prices to Average Value within Flight Dates"
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
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Tag             =   "Increment the revsion number of the contract"
         Top             =   90
         Width           =   5085
      End
   End
   Begin VB.PictureBox plcWeeks 
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
      Height          =   975
      Left            =   1095
      ScaleHeight     =   915
      ScaleWidth      =   3390
      TabIndex        =   1
      Top             =   315
      Width           =   3450
      Begin VB.TextBox edcWeeks 
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
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   5
         Top             =   555
         Width           =   540
      End
      Begin VB.Label lacNoWeeks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   2580
         TabIndex        =   3
         Top             =   150
         Width           =   540
      End
      Begin VB.Label lacWeeks 
         Appearance      =   0  'Flat
         Caption         =   "New Number of Weeks"
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
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   4
         Top             =   570
         Width           =   2085
      End
      Begin VB.Label lacWeeks 
         Appearance      =   0  'Flat
         Caption         =   "Current Number of Weeks"
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
         Index           =   0
         Left            =   315
         TabIndex        =   2
         Top             =   180
         Width           =   2250
      End
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
      Left            =   120
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1935
      Width           =   120
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   30
      ScaleHeight     =   270
      ScaleWidth      =   2745
      TabIndex        =   0
      Top             =   0
      Width           =   2745
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1485
      TabIndex        =   8
      Top             =   1950
      Width           =   1050
   End
End
Attribute VB_Name = "RCSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rcsplit.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RCSplit.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract revision number increment screen code
Option Explicit
Option Compare Text
Dim smScreenCaption As String
Dim imUpdateAllowed As Integer
Dim imFirstActivate As Integer



Private Sub ckcAvg_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcCancel_Click()
    igRCReturn = 0
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slStr As String
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If igCurNoWks > 0 Then
        slStr = Trim$(edcWeeks.Text)
        igNewNoWks = Val(slStr)
    End If
    If ckcAvg.Value = vbChecked Then
        igAvgPrices = True
    Else
        igAvgPrices = False
    End If
    igRCReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
End Sub
Private Sub edcWeeks_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcWeeks_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcWeeks.Text
    slStr = Left$(slStr, edcWeeks.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcWeeks.SelStart - edcWeeks.SelLength)
    If gCompNumberStr(slStr, "54") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub Form_Activate()
    Dim slStr As String
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    If igCurNoWks > 0 Then
        slStr = Trim$(str$(igCurNoWks))
        lacNoWeeks.Caption = slStr
        edcWeeks.Text = slStr
    Else
        plcAvg.Move plcAvg.Left, plcWeeks.Top + (cmcDone.Top - plcWeeks.Top) \ 2 - plcAvg.Height \ 2
        plcWeeks.Visible = False
    End If
    ckcAvg.Value = vbUnchecked
    RCSplit.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
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
    If igCurNoWks > 0 Then
        smScreenCaption = "Weeks " & sgWkStartDate & "-" & sgWkEndDate
    Else
        smScreenCaption = "Price Set " & sgWkStartDate & "-" & sgWkEndDate
    End If
    RCSplit.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterModalForm RCSplit
    plcScreen_Paint
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
    Unload RCSplit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RCSplit = Nothing   'Remove data segment
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub
