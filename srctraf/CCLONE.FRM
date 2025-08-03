VERSION 5.00
Begin VB.Form CClone 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5085
   ClientLeft      =   345
   ClientTop       =   2955
   ClientWidth     =   9285
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
   ScaleHeight     =   5085
   ScaleWidth      =   9285
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
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
      Height          =   1770
      Left            =   2895
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2085
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "CCLONE.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   34
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   210
         Left            =   330
         TabIndex        =   35
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.Timer tmcChk 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9015
      Top             =   4035
   End
   Begin VB.Timer tmcDate 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8760
      Top             =   4620
   End
   Begin VB.CommandButton cmcClone 
      Appearance      =   0  'Flat
      Caption         =   "C&lone"
      Height          =   285
      Left            =   3375
      TabIndex        =   29
      Top             =   4680
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
      Left            =   105
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4740
      Width           =   120
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4635
      TabIndex        =   0
      Top             =   4680
      Width           =   1050
   End
   Begin VB.PictureBox plcClone 
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
      Height          =   4185
      Left            =   105
      ScaleHeight     =   4125
      ScaleWidth      =   8745
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   390
      Width           =   8805
      Begin VB.Frame frcTWeeks 
         Caption         =   "To Weeks"
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
         Height          =   3735
         Left            =   3780
         TabIndex        =   15
         Top             =   345
         Width           =   4725
         Begin VB.PictureBox pbcLbcTLines 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   2520
            Left            =   135
            ScaleHeight     =   2520
            ScaleWidth      =   3105
            TabIndex        =   39
            Top             =   810
            Width           =   3105
         End
         Begin VB.CheckBox ckcClonePrice 
            Caption         =   "Clone Price"
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
            Left            =   3315
            TabIndex        =   37
            Top             =   510
            Width           =   1320
         End
         Begin VB.TextBox edcWks 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   1
            Left            =   2610
            MaxLength       =   2
            TabIndex        =   25
            Text            =   "0"
            Top             =   495
            Width           =   375
         End
         Begin VB.TextBox edcWks 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   0
            Left            =   1080
            MaxLength       =   2
            TabIndex        =   23
            Text            =   "1"
            Top             =   495
            Width           =   375
         End
         Begin VB.ListBox lbcTDates 
            Appearance      =   0  'Flat
            Height          =   2550
            Left            =   3435
            MultiSelect     =   2  'Extended
            TabIndex        =   28
            Top             =   795
            Width           =   1110
         End
         Begin VB.TextBox edcTDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   0
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   17
            Top             =   240
            Width           =   930
         End
         Begin VB.CommandButton cmcTDate 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   2070
            Picture         =   "CCLONE.frx":2E1A
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            Width           =   195
         End
         Begin VB.TextBox edcTDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   1
            Left            =   3435
            MaxLength       =   10
            TabIndex        =   20
            Top             =   240
            Width           =   930
         End
         Begin VB.CommandButton cmcTDate 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   4365
            Picture         =   "CCLONE.frx":2F14
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Width           =   195
         End
         Begin VB.ListBox lbcTLines 
            Appearance      =   0  'Flat
            Height          =   2550
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   26
            Top             =   795
            Width           =   3135
         End
         Begin VB.CheckBox ckcTAll 
            Caption         =   "All"
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
            TabIndex        =   27
            Top             =   3450
            Width           =   1425
         End
         Begin VB.Label lacWks 
            Appearance      =   0  'Flat
            Caption         =   "Weeks Off"
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
            Index           =   1
            Left            =   1605
            TabIndex        =   24
            Top             =   510
            Width           =   990
         End
         Begin VB.Label lacWks 
            Appearance      =   0  'Flat
            Caption         =   "Weeks On"
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
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   510
            Width           =   990
         End
         Begin VB.Label lacTDate 
            Appearance      =   0  'Flat
            Caption         =   "Start Date"
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
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   255
            Width           =   990
         End
         Begin VB.Label lacTDate 
            Appearance      =   0  'Flat
            Caption         =   "End Date"
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
            Index           =   1
            Left            =   2400
            TabIndex        =   19
            Top             =   255
            Width           =   990
         End
      End
      Begin VB.Frame frcFWeeks 
         Caption         =   "From Weeks"
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
         Height          =   3720
         Left            =   210
         TabIndex        =   6
         Top             =   345
         Width           =   3435
         Begin VB.PictureBox pbcLbcFLines 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   2520
            Left            =   135
            ScaleHeight     =   2520
            ScaleWidth      =   3105
            TabIndex        =   38
            Top             =   810
            Width           =   3105
         End
         Begin VB.CheckBox ckcFAll 
            Caption         =   "All"
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
            TabIndex        =   14
            Top             =   3450
            Width           =   1425
         End
         Begin VB.ListBox lbcFLines 
            Appearance      =   0  'Flat
            Height          =   2550
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   13
            Top             =   795
            Width           =   3135
         End
         Begin VB.CommandButton cmcFDate 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   2070
            Picture         =   "CCLONE.frx":300E
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   525
            Width           =   195
         End
         Begin VB.TextBox edcFDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   1
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   11
            Top             =   525
            Width           =   930
         End
         Begin VB.CommandButton cmcFDate 
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   5.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   2070
            Picture         =   "CCLONE.frx":3108
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   195
         End
         Begin VB.TextBox edcFDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   210
            Index           =   0
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   8
            Top             =   240
            Width           =   930
         End
         Begin VB.Label lacFDate 
            Appearance      =   0  'Flat
            Caption         =   "End Date"
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
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   540
            Width           =   990
         End
         Begin VB.Label lacFDate 
            Appearance      =   0  'Flat
            Caption         =   "Start Date"
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
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   255
            Width           =   990
         End
      End
      Begin VB.PictureBox plcDir 
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2310
         ScaleHeight     =   225
         ScaleWidth      =   4425
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   75
         Width           =   4425
         Begin VB.OptionButton rbcDir 
            Caption         =   "Lines"
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
            Index           =   2
            Left            =   3375
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   15
            Width           =   870
         End
         Begin VB.OptionButton rbcDir 
            Caption         =   "Weeks Across"
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
            Index           =   0
            Left            =   45
            TabIndex        =   3
            Top             =   15
            Width           =   1620
         End
         Begin VB.OptionButton rbcDir 
            Caption         =   "Weeks Down"
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
            Index           =   1
            Left            =   1740
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   15
            Width           =   1545
         End
      End
   End
   Begin VB.Label plcScreen 
      Caption         =   "Clone"
      Height          =   225
      Left            =   75
      TabIndex        =   40
      Top             =   45
      Width           =   1710
   End
End
Attribute VB_Name = "CClone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CCLONE.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CClone.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract Clone screen code
Option Explicit
Option Compare Text
Dim smFStartDate As String
Dim smFEndDate As String
Dim smTStartDate As String
Dim smTEndDate As String
Dim lmEarliestDate As Long
Dim imWkDays(0 To 6) As Integer 'Days from flight
Dim lmWkSpots As Long        'Number of spots per week
Dim imWkCff As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassFocus As Integer
Dim imLbcArrowSetting As Integer
Dim imLbcMouseDown As Integer
Dim imDateIndex As Integer
Dim smTodaysDate As String
Dim imBypassAll As Integer
Dim imBypassLbc As Integer
Dim imFSelected() As Integer
Dim imTSelected() As Integer
Dim imClf() As Integer          'Clf value for line in lbcFLines
Dim imDGSave(0 To 7) As Integer   'Values saved for Gen, Index:1=Run every; 2=Daily(0)/Weekly(1)/Daily %(2)/Weekly%(3); 3=Preempt (0=Yes; 1=No) Not Used; 4 = Rank (priority)
'Dim imListField(1 To 4) As Integer
Dim imListField(0 To 4) As Integer
Dim imLBCtrls As Integer

'Dim imListFieldChar(1 To 3) As Integer
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Const LINECOMMINDEX = 21

Private Sub ckcFAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcFAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    If imBypassAll Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    tmcChk.Enabled = False
    ilValue = Value
    If lbcFLines.ListCount > 0 Then
        llRg = CLng(lbcFLines.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcFLines.hWnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    tmcChk.Enabled = True
    pbcLbcFLines_Paint
    Screen.MousePointer = vbDefault
End Sub
Private Sub ckcFAll_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub ckcTAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcTAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    If imBypassAll Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilValue = Value
    If lbcTLines.ListCount > 0 Then
        llRg = CLng(lbcTLines.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcTLines.hWnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    tmcChk.Enabled = True
    pbcLbcTLines_Paint
    Screen.MousePointer = vbDefault
End Sub
Private Sub ckcTAll_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    Select Case imDateIndex
        Case 0
            edcFDate(0).SelStart = 0
            edcFDate(0).SelLength = Len(edcFDate(0).Text)
            edcFDate(0).SetFocus
        Case 1
            edcFDate(1).SelStart = 0
            edcFDate(1).SelLength = Len(edcFDate(1).Text)
            edcFDate(1).SetFocus
        Case 2
            edcTDate(0).SelStart = 0
            edcTDate(0).SelLength = Len(edcTDate(0).Text)
            edcTDate(0).SetFocus
        Case 3
            edcTDate(1).SelStart = 0
            edcTDate(1).SelLength = Len(edcTDate(1).Text)
            edcTDate(1).SetFocus
    End Select
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    Select Case imDateIndex
        Case 0
            edcFDate(0).SelStart = 0
            edcFDate(0).SelLength = Len(edcFDate(0).Text)
            edcFDate(0).SetFocus
        Case 1
            edcFDate(1).SelStart = 0
            edcFDate(1).SelLength = Len(edcFDate(1).Text)
            edcFDate(1).SetFocus
        Case 2
            edcTDate(0).SelStart = 0
            edcTDate(0).SelLength = Len(edcTDate(0).Text)
            edcTDate(0).SetFocus
        Case 3
            edcTDate(1).SelStart = 0
            edcTDate(1).SelLength = Len(edcTDate(1).Text)
            edcTDate(1).SetFocus
    End Select
End Sub
Private Sub cmcClone_Click()
    Dim slDate As String
    Dim ilRet As Integer
    slDate = edcFDate(0).Text
    If Not gValidDate(slDate) Then
        Beep
        edcFDate(0).SetFocus
        Exit Sub
    End If
    If gWeekDayStr(slDate) <> 0 Then
        Beep
        ilRet = MsgBox("From Week Start Date must be a Monday", vbOKOnly + vbExclamation, "Error")
        edcFDate(0).SetFocus
        Exit Sub
    End If
    If rbcDir(1).Value Or rbcDir(2).Value Then
        slDate = edcFDate(1).Text
        If Not gValidDate(slDate) Then
            Beep
            edcFDate(1).SetFocus
            Exit Sub
        End If
        'If gWeekDayStr(slDate) <> 6 Then
        '    Beep
        '    ilRet = MsgBox("From Week End Date must be a Sunday", vbOkOnly + vbExclamation, "Error")
        '    edcFDate(0).SetFocus
        '    Exit Sub
        'End If
    End If
    If rbcDir(0).Value Or rbcDir(2).Value Then
        slDate = edcTDate(0).Text
        If Not gValidDate(slDate) Then
            Beep
            edcTDate(0).SetFocus
            Exit Sub
        End If
        If gWeekDayStr(slDate) <> 0 Then
            Beep
            ilRet = MsgBox("To Week Start Date must be a Monday", vbOKOnly + vbExclamation, "Error")
            edcTDate(0).SetFocus
            Exit Sub
        End If
        If rbcDir(0).Value Then
            slDate = edcTDate(1).Text
            If Not gValidDate(slDate) Then
                Beep
                edcTDate(1).SetFocus
                Exit Sub
            End If
        End If
    End If
    'If rbcDir(1).Value Or rbcDir(2).Value Then
    If rbcDir(1).Value Then
        If lbcFLines.SelCount <> 1 Then
            Beep
            ilRet = MsgBox("Only One From Line can be Specified", vbOKOnly + vbExclamation, "Error")
            lbcFLines.SetFocus
            Exit Sub
        End If
    End If
    'If rbcDir(0).Value Then
    If rbcDir(0).Value Or rbcDir(2).Value Then
        slDate = edcTDate(0).Text
        If lmEarliestDate > gDateValue(slDate) Then
            Beep
            ilRet = MsgBox("'To Start Date' must not be before " & Format$(lmEarliestDate, "m/d/yy"), vbOKOnly + vbExclamation, "Error")
            edcTDate(0).SetFocus
            Exit Sub
        End If
    End If
    If rbcDir(1).Value Then
        slDate = edcFDate(0).Text
        If lmEarliestDate > gDateValue(slDate) Then
            Beep
            ilRet = MsgBox("'From Start Date' must not be before " & Format$(lmEarliestDate, "m/d/yy"), vbOKOnly + vbExclamation, "Error")
            edcFDate(0).SetFocus
            Exit Sub
        End If
    End If
    If rbcDir(2).Value Then
        If gDateValue(edcFDate(1).Text) < gDateValue(edcFDate(0).Text) Then
            Beep
            ilRet = MsgBox("'From Start Date' must prior to 'From End Date'", vbOKOnly + vbExclamation, "Error")
            edcFDate(0).SetFocus
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    'If rbcDir(0).Value Or rbcDir(2).Value Then
    If rbcDir(0).Value Then
        mCloneAcross
    End If
    'If rbcDir(1).Value Or rbcDir(2).Value Then
    If rbcDir(1).Value Then
        mCloneDown
    End If
    If rbcDir(2).Value Then
        mCloneLine
    End If
    If igShiftDates <> 0 Then
        cmcDone.Caption = "&Done"
    End If
    Screen.MousePointer = vbDefault
    'igTerminateReturn = 1
    'mTerminate
End Sub
Private Sub cmcClone_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcDone_Click()
    igTerminateReturn = 0
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcFDate_Click(Index As Integer)
    plcCalendar.Visible = Not plcCalendar.Visible
    edcFDate(Index).SelStart = 0
    edcFDate(Index).SelLength = Len(edcFDate(Index).Text)
    edcFDate(Index).SetFocus
End Sub
Private Sub cmcFDate_GotFocus(Index As Integer)
    If imDateIndex <> Index Then
        plcCalendar.Visible = False
    End If
    imDateIndex = Index
    plcCalendar.Move plcClone.Left + frcFWeeks.Left + edcFDate(Index).Left, plcClone.Top + frcFWeeks.Top + edcFDate(Index).Top + edcFDate(Index).Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcTDate_Click(Index As Integer)
    plcCalendar.Visible = Not plcCalendar.Visible
    edcTDate(Index).SelStart = 0
    edcTDate(Index).SelLength = Len(edcTDate(Index).Text)
    edcTDate(Index).SetFocus
End Sub
Private Sub cmcTDate_GotFocus(Index As Integer)
    If imDateIndex <> Index + 2 Then
        plcCalendar.Visible = False
    End If
    imDateIndex = Index + 2
    plcCalendar.Move plcClone.Left + frcTWeeks.Left + edcTDate(Index).Left, plcClone.Top + frcTWeeks.Top + edcTDate(Index).Top + edcTDate(Index).Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcFDate_Change(Index As Integer)
    Dim slStr As String
    slStr = edcFDate(Index).Text
    If Not gValidDate(slStr) Then
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcFDate_GotFocus(Index As Integer)
    If Not imBypassFocus Then
        If imDateIndex <> Index Then
            plcCalendar.Visible = False
        End If
        imDateIndex = Index
        plcCalendar.Move plcClone.Left + frcFWeeks.Left + edcFDate(Index).Left, plcClone.Top + frcFWeeks.Top + edcFDate(Index).Top + edcFDate(Index).Height
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcFDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcFDate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcFDate(Index).SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcFDate_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcFDate(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcFDate(Index).Text = slDate
            End If
        End If
        edcFDate(Index).SelStart = 0
        edcFDate(Index).SelLength = Len(edcFDate(Index).Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcFDate(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcFDate(Index).Text = slDate
            End If
        End If
        edcFDate(Index).SelStart = 0
        edcFDate(Index).SelLength = Len(edcFDate(Index).Text)
    End If
End Sub
Private Sub edcTDate_Change(Index As Integer)
    Dim slStr As String
    tmcDate.Enabled = False
    slStr = edcTDate(Index).Text
    If Not gValidDate(slStr) Then
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    tmcDate.Enabled = True
End Sub
Private Sub edcTDate_GotFocus(Index As Integer)
    If Not imBypassFocus Then
        If imDateIndex <> Index + 2 Then
            plcCalendar.Visible = False
        End If
        imDateIndex = Index + 2
        plcCalendar.Move plcClone.Left + frcTWeeks.Left + edcTDate(Index).Left, plcClone.Top + frcTWeeks.Top + edcTDate(Index).Top + edcTDate(Index).Height
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcTDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcTDate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcTDate(Index).SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcTDate_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcTDate(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcTDate(Index).Text = slDate
            End If
        End If
        edcTDate(Index).SelStart = 0
        edcTDate(Index).SelLength = Len(edcTDate(Index).Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcTDate(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcTDate(Index).Text = slDate
            End If
        End If
        edcTDate(Index).SelStart = 0
        edcTDate(Index).SelLength = Len(edcTDate(Index).Text)
    End If
End Sub
Private Sub edcWks_Change(Index As Integer)
    tmcDate.Enabled = False
    tmcDate.Enabled = True
End Sub
Private Sub edcWks_GotFocus(Index As Integer)
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcWks_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcWks(Index).Text
    'If (Trim$(slStr) = "") And (KeyAscii = KEY0) Then
    '    Beep
    '    KeyAscii = 0
    '    Exit Sub
    'End If
    slStr = Left$(slStr, edcWks(Index).SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcWks(Index).SelStart - edcWks(Index).SelLength)
    If gCompNumberStr(slStr, "99") > 0 Then
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
        plcCalendar.Visible = False
        gFunctionKeyBranch KeyCode
    End If

End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase imFSelected
    Erase imTSelected
    Erase imClf
    
    Set CClone = Nothing   'Remove data segment

End Sub

Private Sub lbcFLines_Click()
    If imBypassLbc Then
        Exit Sub
    End If
    tmcChk.Enabled = False
    imBypassAll = True
    ckcFAll.Value = vbUnchecked
    imBypassAll = False
    If rbcDir(0).Value Then
        tmcChk.Enabled = True
    ElseIf rbcDir(1).Value Then
        tmcChk.Enabled = True
    ElseIf rbcDir(2).Value Then
        tmcChk.Enabled = True
    End If
    pbcLbcFLines_Paint
End Sub
Private Sub lbcFLines_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub lbcFLines_Scroll()
    pbcLbcFLines_Paint
End Sub

Private Sub lbcTDates_GotFocus()
    If tmcDate.Enabled Then
        tmcDate_Timer
    End If
    plcCalendar.Visible = False
End Sub
Private Sub lbcTLines_Click()
    If imBypassLbc Then
        Exit Sub
    End If
    tmcChk.Enabled = False
    imBypassAll = True
    ckcTAll.Value = vbUnchecked
    imBypassAll = False
    If rbcDir(1).Value Then
        tmcChk.Enabled = True
    End If
    pbcLbcTLines_Paint
End Sub
Private Sub lbcTLines_GotFocus()
    plcCalendar.Visible = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    Select Case imDateIndex
        Case 0
            slStr = edcFDate(0).Text
        Case 1
            slStr = edcFDate(1).Text
        Case 2
            slStr = edcTDate(0).Text
        Case 3
            slStr = edcTDate(1).Text
    End Select
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCloneAcross                    *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clone weeks across             *
'*                                                     *
'*******************************************************
Private Sub mCloneAcross()
    Dim ilFClf As Integer
    Dim llFDate As Long
    Dim ilDate As Integer
    Dim llTDate As Long
    Dim ilTCff As Integer
    Dim llTstDate As Long
    Dim llActPrice As Long
    Dim llPropPrice As Long
    Dim slPriceType As String
    Dim slStr As String
    Dim ilClf As Integer
    Dim ilClonePrice As Integer
    Dim slWDFlag As String
    ReDim ilDailySpots(0 To 6) As Integer

    smFStartDate = edcFDate(0).Text
    If Not gValidDate(smFStartDate) Then
        Beep
        Exit Sub
    End If
    ilClonePrice = False
    If Contract.plcScreen <> "Orders" Then
        If ckcClonePrice.Value = vbChecked Then
            ilClonePrice = True
        End If
    End If
    smFStartDate = gObtainPrevMonday(smFStartDate)
    llFDate = gDateValue(smFStartDate)
    For ilClf = 0 To lbcFLines.ListCount - 1 Step 1
        If lbcFLines.Selected(ilClf) Then
            ilFClf = imClf(ilClf)
            If tgClfCntr(ilFClf).iFirstCgf = -1 Then
                imWkCff = mGetFlightSpotsDays(ilFClf + 1, llFDate, imOvDays(), lmWkSpots, imWkDays(), slWDFlag, ilDailySpots())
                If imWkCff >= 0 Then
                    For ilDate = 0 To lbcTDates.ListCount - 1 Step 1
                        If lbcTDates.Selected(ilDate) Then
                            If igShiftDates = 0 Then
                                igShiftDates = 1
                            End If
                            slStr = lbcTDates.List(ilDate)
                            llTDate = gDateValue(slStr)
                            If slWDFlag = "D" Then
                                mAddWeek imLastCffUsed, ilFClf + 1, llTDate, "D", lmWkSpots, ilDailySpots()
                            Else
                                mAddWeek imLastCffUsed, ilFClf + 1, llTDate, "W", lmWkSpots, imWkDays()
                            End If
                            imWkCff = mGetFlightSpotsDays(ilFClf + 1, llFDate, imOvDays(), lmWkSpots, imWkDays(), slWDFlag, ilDailySpots())
                            If imWkCff >= 0 Then
                                slPriceType = tgCffCntr(imWkCff).CffRec.sPriceType
                                llActPrice = tgCffCntr(imWkCff).CffRec.lActPrice
                                llPropPrice = tgCffCntr(imWkCff).CffRec.lPropPrice
                                ilTCff = tgClfCntr(ilFClf).iFirstCff
                                Do While ilTCff <> -1
                                    gUnpackDateLong tgCffCntr(ilTCff).CffRec.iStartDate(0), tgCffCntr(ilTCff).CffRec.iStartDate(1), llTstDate
                                    If llTstDate = llTDate Then
                                        If Contract.plcScreen = "Orders" Then
                                            tgCffCntr(ilTCff).CffRec.sPriceType = slPriceType 'tgCffCntr(imWkCff).CffRec.sPriceType
                                            tgCffCntr(ilTCff).CffRec.lPropPrice = 0 'tgCffCntr(ilFCff).CffRec.lPropPrice
                                            tgCffCntr(ilTCff).CffRec.lActPrice = llActPrice 'tgCffCntr(imWkCff).CffRec.lActPrice
                                        Else
                                            'tgCffCntr(ilTCff).CffRec.sPriceType = "T"   'tgCffCntr(ilFCff).CffRec.sPriceType
                                            ''tgCffCntr(ilTCff).CffRec.lPropPrice = tgCffCntr(ilFCff).CffRec.lPropPrice
                                            ''tgCffCntr(ilTCff).CffRec.lActPrice = tgCffCntr(ilFCff).CffRec.lActPrice
                                            'slDate = Format$(llTDate, "m/d/yy")
                                            'tgCffCntr(ilTCff).CffRec.lPropPrice = mComputeRCWkPrice(igCloneRcfCode, tgClfCntr(ilFClf).ClfRec.iRdfCode, slDate, tgClfCntr(ilFClf).ClfRec.iVefCode, ilFClf + 1)
                                            'tgCffCntr(ilTCff).CffRec.lActPrice = 100 * tgCffCntr(ilTCff).CffRec.lPropPrice
                                            'Get Proposal Price
                                            mSetFlightPriceForWk ilFClf + 1, llTDate, "Rate Card", imOvDays(), tmRcf, tgClfCntr(ilFClf).ClfRec.iVefCode, imDGSave(), 0
                                            If (llActPrice <> llPropPrice * 100) Or (ilClonePrice) Then
                                                tgCffCntr(ilTCff).CffRec.sPriceType = slPriceType
                                                tgCffCntr(ilTCff).CffRec.lActPrice = llActPrice
                                            End If
                                        End If
                                        Exit Do
                                    End If
                                    ilTCff = tgCffCntr(ilTCff).iNextCff
                                Loop
                            End If
                        End If
                    Next ilDate
                End If
            End If
        End If
    Next ilClf
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCloneDown                      *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clone weeks down               *
'*                                                     *
'*******************************************************
Private Sub mCloneDown()
    Dim ilFClf As Integer
    Dim ilTClf As Integer
    Dim llFSDate As Long
    Dim llFEDate As Long
    Dim llTDate As Long
    Dim ilTCff As Integer
    Dim llTstDate As Long
    Dim ilDay As Integer
    Dim slLineType As String
    Dim ilClf As Integer
    Dim slWDFlag As String
    ReDim ilDailySpots(0 To 6) As Integer
    smFStartDate = edcFDate(0).Text
    If Not gValidDate(smFStartDate) Then
        Beep
        Exit Sub
    End If
    smFEndDate = edcFDate(1).Text
    If Not gValidDate(smFEndDate) Then
        Beep
        Exit Sub
    End If
    smFStartDate = gObtainPrevMonday(smFStartDate)
    llFSDate = gDateValue(smFStartDate)
    smFEndDate = gObtainNextSunday(smFEndDate)
    llFEDate = gDateValue(smFEndDate)
    For ilClf = 0 To lbcFLines.ListCount - 1 Step 1
        If lbcFLines.Selected(ilClf) Then
            ilFClf = imClf(ilClf)
            If tgClfCntr(ilFClf).iFirstCgf = -1 Then
                For llTDate = llFSDate To llFEDate Step 7
                    imWkCff = mGetFlightSpotsDays(ilFClf + 1, llTDate, imOvDays(), lmWkSpots, imWkDays(), slWDFlag, ilDailySpots())
                    If imWkCff >= 0 Then
                        For ilTClf = 0 To lbcTLines.ListCount - 1 Step 1
                            If lbcTLines.Selected(ilTClf) Then
                                slLineType = mGetLineType(ilTClf + 1)
                                If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                                    If igShiftDates = 0 Then
                                        igShiftDates = 1
                                    End If
                                    mGetFlightDays ilTClf + 1, imOvDays()
                                    For ilDay = 0 To 6 Step 1
                                        If imOvDays(ilDay) = 1 Then
                                            imWkDays(ilDay) = False
                                        Else
                                            imWkDays(ilDay) = True
                                        End If
                                    Next ilDay
                                    If slWDFlag = "D" Then
                                        mAddWeek imLastCffUsed, ilTClf + 1, llTDate, "D", lmWkSpots, ilDailySpots()
                                    Else
                                        mAddWeek imLastCffUsed, ilTClf + 1, llTDate, "W", lmWkSpots, imWkDays()
                                    End If
                                    ilTCff = tgClfCntr(ilTClf).iFirstCff
                                    Do While ilTCff <> -1
                                        gUnpackDateLong tgCffCntr(ilTCff).CffRec.iStartDate(0), tgCffCntr(ilTCff).CffRec.iStartDate(1), llTstDate
                                        If llTstDate = llTDate Then
                                            If Contract.plcScreen = "Orders" Then
                                                'tgCffCntr(ilTCff).CffRec.sPriceType = "T"
                                                'tgCffCntr(ilTCff).CffRec.lPropPrice = 0
                                                'tgCffCntr(ilTCff).CffRec.lActPrice = 0
                                                'Adjacent Price
                                                mSetFlightPriceForWk ilTClf + 1, llTDate, "", imOvDays(), tmRcf, tgClfCntr(ilTClf).ClfRec.iVefCode, imDGSave(), 0
                                            Else
                                                'tgCffCntr(ilTCff).CffRec.sPriceType = "T"   'tgCffCntr(ilFCff).CffRec.sPriceType
                                                ''tgCffCntr(ilTCff).CffRec.lPropPrice = tgCffCntr(ilFCff).CffRec.lPropPrice
                                                ''tgCffCntr(ilTCff).CffRec.lActPrice = tgCffCntr(ilFCff).CffRec.lActPrice
                                                'slDate = Format$(llTDate, "m/d/yy")
                                                'tgCffCntr(ilTCff).CffRec.lPropPrice = mComputeRCWkPrice(igCloneRcfCode, tgClfCntr(ilTClf).ClfRec.iRdfCode, slDate, tgClfCntr(ilTClf).ClfRec.iVefCode, ilTClf + 1)
                                                'tgCffCntr(ilTCff).CffRec.lActPrice = tgCffCntr(ilTCff).CffRec.lPropPrice
                                                mSetFlightPriceForWk ilTClf + 1, llTDate, "Rate Card", imOvDays(), tmRcf, tgClfCntr(ilTClf).ClfRec.iVefCode, imDGSave(), 0
                                            End If
                                            Exit Do
                                        End If
                                        ilTCff = tgCffCntr(ilTCff).iNextCff
                                    Loop
                                End If
                            End If
                        Next ilTClf
                    End If
                Next llTDate
            End If
        End If
    Next ilClf
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
    Dim slStr As String
    Dim ilClf As Integer
    Dim slRecCode As String
    Dim slLineType As String
    Dim ilTest As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slName As String
    Dim ilVef As Integer
    Dim ilLoop As Integer
    Dim ilSelHidden As Integer
    Dim ilSport As Integer
    
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imLBCDCtrls = 1
    imFirstActivate = True
    imBypassAll = False
    imBypassLbc = False
    imBypassFocus = False
    imChgMode = False
    imBSMode = False
    imLbcArrowSetting = False
    imLbcMouseDown = False
    imCalType = 0   'Standard
    mInitBox
    imListField(1) = 15
    imListField(2) = 6 * igAlignCharWidth '6
    imListField(3) = 20 * igAlignCharWidth '20
    imListField(4) = 40 * igAlignCharWidth  '40
    If sgPassValue = "" Then
        lmEarliestDate = 0
    Else
        lmEarliestDate = gDateValue(sgPassValue)
    End If
    smTodaysDate = Format$(gNow(), "mm/dd/yy")
    If igCloneLn > 0 And sgCloneDate <> "" Then
        edcFDate(0).Text = sgCloneDate
    Else
        slStr = gObtainNextMonday(smTodaysDate)
        gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
        pbcCalendar_Paint   'mBoxCalDate called within paint
        lacDate.Visible = False
    End If
    CClone.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterModalForm CClone
    ilSelHidden = False
    ReDim imClf(0 To 0) As Integer
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        slLineType = mGetLineType(ilClf + 1)
        ilSport = False
        If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
            ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
            If ilVef <> -1 Then
                If tgMVef(ilVef).sType = "G" Then
                    ilSport = True
                End If
            End If
        Else
            If mGameInPkg(ilClf + 1) Then
                ilSport = True
            End If
        End If
        If Not ilSport Then
            If slLineType <> "H" Then
                slStr = Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine))
            Else
                slStr = "  " & Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine))
            End If
            If tgClfCntr(ilClf).iCancel Then
                slStr = slStr & "| |" & "Canceled before Start"
            Else
                slRecCode = Trim$(str$(tgClfCntr(ilClf).ClfRec.iVefCode))
                If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                    For ilTest = 0 To UBound(tmVehicleCode) - 1 Step 1  'Contract!lbcVehicle.ListCount - 1 Step 1
                        slNameCode = tmVehicleCode(ilTest).sKey    'Contract!lbcVehicle.List(ilTest)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        On Error GoTo mInitErr
                        gCPErrorMsg ilRet, "mInit (gParseItem field 2)", CClone
                        On Error GoTo 0
                        If slRecCode = slCode Then
                            ilRet = gParseItem(slNameCode, 1, "\", slName)
                            On Error GoTo mInitErr
                            gCPErrorMsg ilRet, "mInit (gParseItem field 2)", CClone
                            On Error GoTo 0
                            ilRet = gParseItem(slName, 3, "|", slName)
                            On Error GoTo mInitErr
                            gCPErrorMsg ilRet, "mInit (gParseItem field 2)", CClone
                            On Error GoTo 0
                            slStr = slStr & "|" & slName
                            Exit For
                        End If
                    Next ilTest
                Else
                    If (tgClfCntr(ilClf).ClfRec.iVefCode = 0) Then
                        slStr = slStr & "|" & Trim$(smLnSave(1, ilClf + 1))
                    Else
                        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If tgClfCntr(ilClf).ClfRec.iVefCode = tgMVef(ilVef).iCode Then
                            ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
                            If ilVef <> -1 Then
                                slStr = slStr & "|" & Trim$(tgMVef(ilVef).sName)
                        '        Exit For
                            End If
                        'Next ilVef
                    End If
                End If
                slStr = slStr & "|" & mMakeDPName(ilClf + 1)  '
                'slStr = slStr & "|" & Trim$(Str$(tgClfCntr(ilClf).ClfRec.iLen))
                'If tgClfCntr(ilClf).iFirstCff <> -1 Then
                '    slStartDate = ""
                '    slEndDate = ""
                '    ilCff = tgClfCntr(ilClf).iFirstCff
                '    Do While ilCff <> -1
                '        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                '            If slStartDate = "" Then
                '                gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate
                '            End If
                '            gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate
                '        End If
                '        ilCff = tgCffCntr(ilCff).iNextCff
                '    Loop
                '    slStr = slStr & "|" & slStartDate & "-" & slEndDate
                '    slStr = slStr & "|" & slStartDate
                'End If
            End If
            If Not tgClfCntr(ilClf).iCancel Then
                'lbcFLines.AddItem gAlignStringByPixel(slStr, "|", imListField(), imListFieldChar())
                lbcFLines.AddItem slStr
                imClf(UBound(imClf)) = ilClf
                ReDim Preserve imClf(0 To UBound(imClf) + 1) As Integer
                If igCloneLn = tgClfCntr(ilClf).ClfRec.iLine Then
                    lbcFLines.Selected(lbcFLines.ListCount - 1) = True
                    If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                        ilSelHidden = True
                    End If
                Else
                    If ilSelHidden Then
                        If slLineType = "H" Then
                            lbcFLines.Selected(lbcFLines.ListCount - 1) = True
                        Else
                            ilSelHidden = False
                        End If
                    End If
                End If
            End If
            lbcTLines.AddItem slStr
        End If
    Next ilClf
    If lbcFLines.ListCount > 0 Then
        ReDim imFSelected(0 To lbcFLines.ListCount - 1) As Integer
    Else
        ReDim imFSelected(0 To 0) As Integer
    End If
    If lbcTLines.ListCount > 0 Then
        ReDim imTSelected(0 To lbcTLines.ListCount - 1) As Integer
    Else
        ReDim imTSelected(0 To 0) As Integer
    End If
    For ilLoop = 0 To lbcFLines.ListCount - 1 Step 1
        If lbcFLines.Selected(ilLoop) Then
            imFSelected(ilLoop) = True
        Else
            imFSelected(ilLoop) = False
        End If
    Next ilLoop
    For ilLoop = 0 To lbcTLines.ListCount - 1 Step 1
        If lbcTLines.Selected(ilLoop) Then
            imTSelected(ilLoop) = True
        Else
            imTSelected(ilLoop) = False
        End If
    Next ilLoop
    If Contract.plcScreen <> "Orders" Then
        ckcClonePrice.Value = vbChecked
    End If
    rbcDir(0).Value = True
    pbcLbcFLines_Paint
    pbcLbcTLines_Paint
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
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
'
'   mInitBox
'   Where:
'
    'Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
    'flTextHeight = pbcDate.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcClone.Move 105, 375
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
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
    Unload CClone
End Sub

Private Sub lbcTLines_Scroll()
    pbcLbcTLines_Paint
End Sub

Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                Select Case imDateIndex
                    Case 0
                        edcFDate(0).Text = Format$(llDate, "m/d/yy")
                        edcFDate(0).SelStart = 0
                        edcFDate(0).SelLength = Len(edcFDate(0).Text)
                        imBypassFocus = True
                        edcFDate(0).SetFocus
                    Case 1
                        edcFDate(1).Text = Format$(llDate, "m/d/yy")
                        edcFDate(1).SelStart = 0
                        edcFDate(1).SelLength = Len(edcFDate(1).Text)
                        imBypassFocus = True
                        edcFDate(1).SetFocus
                    Case 2
                        edcTDate(0).Text = Format$(llDate, "m/d/yy")
                        edcTDate(0).SelStart = 0
                        edcTDate(0).SelLength = Len(edcTDate(0).Text)
                        imBypassFocus = True
                        edcTDate(0).SetFocus
                    Case 3
                        edcTDate(1).Text = Format$(llDate, "m/d/yy")
                        edcTDate(1).SelStart = 0
                        edcTDate(1).SelLength = Len(edcTDate(1).Text)
                        imBypassFocus = True
                        edcTDate(1).SetFocus
                End Select
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    Select Case imDateIndex
        Case 0
            edcFDate(0).SetFocus
        Case 1
            edcFDate(1).SetFocus
        Case 2
            edcTDate(0).SetFocus
        Case 3
            edcTDate(1).SetFocus
    End Select
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcLbcFLines_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilFLinesEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 3) As String
    Dim llFgColor As Long
    Dim ilFieldIndex As Integer
    
    ilFLinesEnd = lbcFLines.TopIndex + lbcFLines.Height \ fgListHtArial825
    If ilFLinesEnd > lbcFLines.ListCount Then
        ilFLinesEnd = lbcFLines.ListCount
    End If
    If lbcFLines.ListCount <= lbcFLines.Height \ fgListHtArial825 Then
        llWidth = lbcFLines.Width - 30
    Else
        llWidth = lbcFLines.Width - igScrollBarWidth - 30
    End If
    pbcLbcFLines.Width = llWidth
    pbcLbcFLines.Cls
    llFgColor = pbcLbcFLines.ForeColor
    For ilLoop = lbcFLines.TopIndex To ilFLinesEnd - 1 Step 1
        pbcLbcFLines.ForeColor = llFgColor
        If lbcFLines.MultiSelect = 0 Then
            If lbcFLines.ListIndex = ilLoop Then
                gPaintArea pbcLbcFLines, CSng(0), CSng((ilLoop - lbcFLines.TopIndex) * fgListHtArial825), CSng(pbcLbcFLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcFLines.ForeColor = vbWhite
            End If
        Else
            If lbcFLines.Selected(ilLoop) Then
                gPaintArea pbcLbcFLines, CSng(0), CSng((ilLoop - lbcFLines.TopIndex) * fgListHtArial825), CSng(pbcLbcFLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcFLines.ForeColor = vbWhite
            End If
        End If
        slStr = lbcFLines.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcFLines.CurrentX = imListField(ilField)
            pbcLbcFLines.CurrentY = (ilLoop - lbcFLines.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcFLines, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcFLines.Print slStr
        Next ilField
        pbcLbcFLines.ForeColor = llFgColor
    Next ilLoop

End Sub

Private Sub pbcLbcTLines_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilTLinesEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 3) As String
    Dim llFgColor As Long
    Dim ilFieldIndex As Integer
    ilTLinesEnd = lbcTLines.TopIndex + lbcTLines.Height \ fgListHtArial825
    If ilTLinesEnd > lbcTLines.ListCount Then
        ilTLinesEnd = lbcTLines.ListCount
    End If
    If lbcTLines.ListCount <= lbcTLines.Height \ fgListHtArial825 Then
        llWidth = lbcTLines.Width - 30
    Else
        llWidth = lbcTLines.Width - igScrollBarWidth - 30
    End If
    pbcLbcTLines.Width = llWidth
    pbcLbcTLines.Cls
    llFgColor = pbcLbcTLines.ForeColor
    For ilLoop = lbcTLines.TopIndex To ilTLinesEnd - 1 Step 1
        pbcLbcTLines.ForeColor = llFgColor
        If lbcTLines.MultiSelect = 0 Then
            If lbcTLines.ListIndex = ilLoop Then
                gPaintArea pbcLbcTLines, CSng(0), CSng((ilLoop - lbcTLines.TopIndex) * fgListHtArial825), CSng(pbcLbcTLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcTLines.ForeColor = vbWhite
            End If
        Else
            If lbcTLines.Selected(ilLoop) Then
                gPaintArea pbcLbcTLines, CSng(0), CSng((ilLoop - lbcTLines.TopIndex) * fgListHtArial825), CSng(pbcLbcTLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcTLines.ForeColor = vbWhite
            End If
        End If
        slStr = lbcTLines.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcTLines.CurrentX = imListField(ilField)
            pbcLbcTLines.CurrentY = (ilLoop - lbcTLines.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcTLines, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcTLines.Print slStr
        Next ilField
        pbcLbcTLines.ForeColor = llFgColor
    Next ilLoop

End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcDir_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcDir(Index).Value
    'End of coded added
    Select Case Index
        Case 0  'Weeks Across
            edcFDate(1).Visible = False
            cmcFDate(1).Visible = False
            lacFDate(1).Visible = False
            lbcTLines.Visible = False
            pbcLbcTLines.Visible = False
            ckcTAll.Visible = False
            lacTDate(0).Visible = True
            lacTDate(1).Visible = True
            edcTDate(0).Visible = True
            edcTDate(1).Visible = True
            cmcTDate(0).Visible = True
            cmcTDate(1).Visible = True
            lacWks(0).Visible = True
            lacWks(1).Visible = True
            edcWks(0).Visible = True
            edcWks(1).Visible = True
            lbcTDates.Visible = True
            ckcFAll.Visible = True
            If Contract.plcScreen <> "Orders" Then
                ckcClonePrice.Visible = True
            Else
                ckcClonePrice.Visible = False
            End If
            tmcChk_Timer
        Case 1  'Weeks Down
            lacTDate(0).Visible = False
            lacTDate(1).Visible = False
            edcTDate(0).Visible = False
            cmcTDate(1).Visible = False
            cmcTDate(0).Visible = False
            edcTDate(1).Visible = False
            lacWks(0).Visible = False
            lacWks(1).Visible = False
            edcWks(0).Visible = False
            edcWks(1).Visible = False
            lbcTDates.Visible = False
            ckcClonePrice.Visible = False
            edcFDate(1).Visible = True
            lacFDate(1).Visible = True
            cmcFDate(1).Visible = True
            lbcTLines.Visible = True
            pbcLbcTLines.Visible = True
            ckcTAll.Visible = True
            If lbcFLines.SelCount > 1 Then
                If ckcFAll.Value = vbChecked Then
                    ckcFAll.Value = vbUnchecked
                Else
                    ckcFAll.Value = vbUnchecked
                    ckcFAll_Click
                End If
            End If
            ckcFAll.Visible = False
            tmcChk_Timer
        Case 2  'Lines
            edcFDate(1).Visible = True
            lacFDate(1).Visible = True
            cmcFDate(1).Visible = True
            lbcTLines.Visible = False
            pbcLbcTLines.Visible = False
            ckcTAll.Visible = False
            lacTDate(0).Visible = True
            lacTDate(1).Visible = False
            edcTDate(0).Visible = True
            edcTDate(1).Visible = False
            cmcTDate(0).Visible = True
            cmcTDate(1).Visible = False
            lacWks(0).Visible = False
            lacWks(1).Visible = False
            edcWks(0).Visible = False
            edcWks(1).Visible = False
            lbcTDates.Visible = False
            If lbcFLines.SelCount > 1 Then
                If ckcFAll.Value = vbChecked Then
                    ckcFAll.Value = vbUnchecked
                Else
                    ckcFAll_Click
                End If
            End If
            ckcFAll.Visible = False
            'If Contract.plcScreen <> "Orders" Then
            '    ckcClonePrice.Visible = True
            'Else
                ckcClonePrice.Visible = False
            'End If
            tmcChk_Timer
    End Select
End Sub
Private Sub rbcDir_GotFocus(Index As Integer)
    plcCalendar.Visible = False
End Sub
Private Sub tmcChk_Timer()
    Dim slLineType As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilFClf As Integer

    imBypassLbc = True
    tmcChk.Enabled = False
    If rbcDir(0).Value Then 'Weeks across
        For ilLoop = 0 To lbcFLines.ListCount - 1 Step 1
            If lbcFLines.Selected(ilLoop) Then
                slLineType = mGetLineType(ilLoop + 1)
                If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                    For ilIndex = ilLoop + 1 To lbcFLines.ListCount - 1 Step 1
                        slStr = Left$(lbcFLines.List(ilIndex), 2)
                        If slStr = "  " Then
                            lbcFLines.Selected(ilIndex) = True
                        Else
                            Exit For
                        End If
                    Next ilIndex
                End If
                ilFClf = imClf(ilLoop)
                If tgClfCntr(ilFClf).iFirstCgf <> -1 Then
                    lbcFLines.Selected(ilLoop) = False
                End If
            End If
        Next ilLoop
    ElseIf rbcDir(1).Value Then 'Weeks down
        If lbcFLines.SelCount > 1 Then
            For ilLoop = 0 To lbcFLines.ListCount - 1 Step 1
                If lbcFLines.Selected(ilLoop) And (Not imFSelected(ilLoop)) Then
                    ckcFAll.Value = vbUnchecked
                    ckcFAll_Click
                    lbcFLines.Selected(ilLoop) = True
                    Exit For
                End If
            Next ilLoop
        End If
        If lbcFLines.SelCount = 1 Then
            For ilLoop = 0 To lbcFLines.ListCount - 1 Step 1
                If lbcFLines.Selected(ilLoop) Then
                    slLineType = mGetLineType(ilLoop + 1)
                    If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                        lbcFLines.Selected(ilLoop) = False
                    End If
                    Exit For
                End If
            Next ilLoop
        End If
        For ilLoop = 0 To lbcFLines.ListCount - 1 Step 1
            If lbcFLines.Selected(ilLoop) Then
                ilFClf = imClf(ilLoop)
                If tgClfCntr(ilFClf).iFirstCgf <> -1 Then
                    lbcFLines.Selected(ilLoop) = False
                End If
            End If
        Next ilLoop
        For ilLoop = 0 To lbcTLines.ListCount - 1 Step 1
            If lbcTLines.Selected(ilLoop) Then
                slLineType = mGetLineType(ilLoop + 1)
                If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                    lbcTLines.Selected(ilLoop) = False
                End If
            End If
        Next ilLoop
    ElseIf rbcDir(2).Value Then
        For ilLoop = 0 To lbcFLines.ListCount - 1 Step 1
            If lbcFLines.Selected(ilLoop) Then
                slLineType = mGetLineType(ilLoop + 1)
                If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                    For ilIndex = ilLoop + 1 To lbcFLines.ListCount - 1 Step 1
                        slStr = Left$(lbcFLines.List(ilIndex), 2)
                        If slStr = "  " Then
                            lbcFLines.Selected(ilIndex) = True
                        Else
                            Exit For
                        End If
                    Next ilIndex
                End If
            End If
        Next ilLoop
    End If
    For ilLoop = 0 To lbcFLines.ListCount - 1 Step 1
        If lbcFLines.Selected(ilLoop) Then
            imFSelected(ilLoop) = True
        Else
            imFSelected(ilLoop) = False
        End If
    Next ilLoop
    pbcLbcFLines_Paint
    pbcLbcTLines_Paint
    imBypassLbc = False
End Sub
Private Sub tmcDate_Timer()
    Dim slStr As String
    Dim llTStart As Long
    Dim llTEnd As Long
    Dim llDate As Long
    Dim slDate As String
    Dim ilNoOn As Integer
    Dim ilNoOff As Integer
    Dim ilOn As Integer
    Dim ilOff As Integer
    Dim ilCycle As Integer  '0=On; 1=Off
    tmcDate.Enabled = False
    If rbcDir(1).Value Then
        Exit Sub
    End If
    smTStartDate = edcTDate(0).Text
    If Not gValidDate(smTStartDate) Then
        lbcTDates.Clear
        Exit Sub
    End If
    smTEndDate = edcTDate(1).Text
    If Not gValidDate(smTEndDate) Then
        lbcTDates.Clear
        Exit Sub
    End If
    slStr = edcWks(0).Text
    If slStr = "" Then
        lbcTDates.Clear
        Exit Sub
    End If
    ilOn = Val(slStr)
    If ilOn = 0 Then
        lbcTDates.Clear
        Exit Sub
    End If
    smTStartDate = gObtainPrevMonday(smTStartDate)
    llTStart = gDateValue(smTStartDate)
    smTEndDate = gObtainNextSunday(smTEndDate)
    llTEnd = gDateValue(smTEndDate)
    slStr = edcWks(0).Text
    slStr = edcWks(1).Text
    ilOff = Val(slStr)
    Screen.MousePointer = vbHourglass
    lbcTDates.Clear
    ilNoOn = 0
    ilNoOff = 0
    ilCycle = 0
    For llDate = llTStart To llTEnd Step 7
        slDate = Format$(llDate, "m/d/yy")
        lbcTDates.AddItem slDate
        If ilCycle = 0 Then
            lbcTDates.Selected(lbcTDates.ListCount - 1) = True
            ilNoOn = ilNoOn + 1
            If ilNoOn >= ilOn Then
                If ilOff > 0 Then
                    ilCycle = 1
                Else
                    ilCycle = 0
                End If
                ilNoOn = 0
            End If
        Else
            ilNoOff = ilNoOff + 1
            If ilNoOff >= ilOff Then
                ilCycle = 0
                ilNoOff = 0
            End If
        End If
    Next llDate
    Screen.MousePointer = vbDefault
End Sub

Private Sub mCloneLine()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLineNo                      ilPkLnNo                      ilDailySpots              *
'*                                                                                        *
'******************************************************************************************

    Dim ilFClf As Integer
    'Dim llTDate As Long
    'Dim ilTCff As Integer
    'Dim llTstDate As Long
    'Dim llActPrice As Long
    'Dim llPropPrice As Long
    'Dim slPriceType As String
    'Dim slStr As String
    Dim ilClf As Integer
    'Dim ilClonePrice As Integer
    'Dim slWDFlag As String
    Dim ilLoop As Integer
    Dim slLineType As String
    Dim ilLnRowNo As Integer
    Dim ilPrevIndex As Integer

    smFStartDate = edcFDate(0).Text
    If Not gValidDate(smFStartDate) Then
        Beep
        Exit Sub
    End If
    smFEndDate = edcFDate(1).Text
    If Not gValidDate(smFEndDate) Then
        Beep
        Exit Sub
    End If
    smTStartDate = edcTDate(0).Text
    If Not gValidDate(smTStartDate) Then
        Beep
        Exit Sub
    End If
    smFStartDate = gObtainPrevMonday(smFStartDate)
    smFEndDate = gObtainNextSunday(smFEndDate)
    smTStartDate = gObtainPrevMonday(smTStartDate)
    'ilClonePrice = True
    'If Contract.plcScreen <> "Orders" Then
    '    If ckcClonePrice.Value = vbChecked Then
    '        ilClonePrice = True
    '    End If
    'End If
    ilPrevIndex = -1
    ilLnRowNo = UBound(tgClfCntr) + 1
    For ilClf = 0 To lbcFLines.ListCount - 1 Step 1
        If lbcFLines.Selected(ilClf) Then
            igShiftDates = 2
            ilFClf = imClf(ilClf)
            mInitLine ilLnRowNo, ilFClf
            slLineType = mGetLineType(ilFClf + 1)
            If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                    If tgPackLn(ilLoop).iLineNo = tgClfCntr(ilFClf).ClfRec.iLine Then
                        tgPackLn(UBound(tgPackLn)) = tgPackLn(ilLoop)
                        ReDim Preserve tgPackLn(LBound(tgPackLn) To UBound(tgPackLn) + 1)
                        tgPackLn(UBound(tgPackLn) - 1).iLineNo = tgClfCntr(ilLnRowNo - 1).ClfRec.iLine
                        ilPrevIndex = UBound(tgPackLn) - 1
                        Exit For
                    End If
                Next ilLoop
            ElseIf (slLineType = "H") Then
                For ilLoop = ilClf - 1 To 0 Step -1
                    slLineType = mGetLineType(imClf(ilLoop) + 1)
                    If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                        If Not lbcFLines.Selected(ilLoop) Then
                            ilPrevIndex = -1
                        End If
                        Exit For
                    End If
                Next ilLoop
                If ilPrevIndex <> -1 Then
                    For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                        If tgPackLn(ilLoop).iLineNo = tgClfCntr(ilFClf).ClfRec.iLine Then
                            tgPackLn(UBound(tgPackLn)) = tgPackLn(ilLoop)
                            tgPackLn(ilPrevIndex).iIndex = UBound(tgPackLn)
                            ilPrevIndex = UBound(tgPackLn)
                            ReDim Preserve tgPackLn(LBound(tgPackLn) To UBound(tgPackLn) + 1)
                            tgPackLn(UBound(tgPackLn) - 1).iLineNo = tgClfCntr(ilLnRowNo - 1).ClfRec.iLine
                            Exit For
                        End If
                    Next ilLoop
                Else
                    tgClfCntr(UBound(tgClfCntr) - 1).ClfRec.sType = "S"
                    tgClfCntr(UBound(tgClfCntr) - 1).ClfRec.iPkLineNo = 0
                End If
            End If
            ilLnRowNo = UBound(tgClfCntr) + 1
        End If
    Next ilClf
End Sub

Private Sub mInitLine(ilLnRowNo As Integer, ilFClf As Integer)
    Dim ilLoop As Integer
    Dim ilLineNo As Integer
    Dim ilFCff As Integer
    Dim ilTCff As Integer
    Dim ilFCgf As Integer
    Dim ilTCgf As Integer
    Dim llFStartDate As Long
    Dim llFEndDate As Long
    Dim llTStartDate As Long
    Dim llSDate As Long
    Dim llEDate As Long

    llFStartDate = gDateValue(smFStartDate)
    llFEndDate = gDateValue(smFEndDate)
    llTStartDate = gDateValue(smTStartDate)

    If tgClfCntr(ilLnRowNo - 1).iStatus = -1 Then
        tgClfCntr(ilLnRowNo - 1).iStatus = 0
    End If
    ReDim Preserve tgClfCntr(0 To ilLnRowNo) As CLFLIST
    ReDim Preserve smLnSave(0 To 29, 0 To UBound(tgClfCntr) + 1) As String
    ReDim Preserve imLnSave(0 To 35, 0 To UBound(tgClfCntr) + 1) As Integer
    ReDim Preserve smLnShow(0 To LINECOMMINDEX, 0 To UBound(tgClfCntr) + 1) As String
    ReDim Preserve smLnComment(0 To UBound(tgClfCntr) + 1) As String
    ReDim Preserve smLnInitComment(0 To UBound(tgClfCntr) + 1) As String
    ReDim Preserve imLnCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
    ReDim Preserve imLnInitCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
    For ilLoop = LBound(imLnSave, FIELDBOUNDINDEX) To UBound(imLnSave, FIELDBOUNDINDEX) Step 1
        imLnSave(ilLoop, UBound(tgClfCntr) + 1) = -1
    Next ilLoop
    tgClfCntr(ilLnRowNo - 1) = tgClfCntr(ilFClf)
    ilLineNo = 0
    For ilLoop = LBound(tgClfCntr) To UBound(tgClfCntr) - 2 Step 1
        If (tgClfCntr(ilLoop).ClfRec.iLine > ilLineNo) And (tgClfCntr(ilLoop).iStatus >= 0) Then
            ilLineNo = tgClfCntr(ilLoop).ClfRec.iLine
        End If
    Next ilLoop
    For ilLoop = LBound(tgDelClf) To UBound(tgDelClf) - 2 Step 1
        If (tgDelClf(ilLoop).ClfRec.iLine > ilLineNo) And (tgDelClf(ilLoop).iStatus >= 0) Then
            ilLineNo = tgDelClf(ilLoop).ClfRec.iLine
        End If
    Next ilLoop
    tgClfCntr(ilLnRowNo - 1).ClfRec.iLine = ilLineNo + 1
    tgClfCntr(ilLnRowNo - 1).iStatus = 0
    tgClfCntr(ilLnRowNo - 1).lRecPos = 0
    tgClfCntr(ilLnRowNo - 1).iFirstCff = -1
    tgClfCntr(ilLnRowNo - 1).iFirstCgf = -1
    tgClfCntr(ilLnRowNo - 1).lFirstAvail = -1
    ilFCff = tgClfCntr(ilFClf).iFirstCff
    Do While ilFCff <> -1
        gUnpackDateLong tgCffCntr(ilFCff).CffRec.iStartDate(0), tgCffCntr(ilFCff).CffRec.iStartDate(1), llSDate
        gUnpackDateLong tgCffCntr(ilFCff).CffRec.iEndDate(0), tgCffCntr(ilFCff).CffRec.iEndDate(1), llEDate
        If (llEDate >= llFStartDate) And (llSDate <= llFEndDate) And (llSDate <= llEDate) Then
            If (tgCffCntr(ilFCff).iStatus = 0) Or (tgCffCntr(ilFCff).iStatus = 1) Then
                ilTCff = UBound(tgCffCntr)
                If tgClfCntr(ilLnRowNo - 1).iFirstCff = -1 Then
                    tgClfCntr(ilLnRowNo - 1).iFirstCff = ilTCff
                Else
                    tgCffCntr(ilTCff - 1).iNextCff = ilTCff
                End If
                tgCffCntr(ilTCff) = tgCffCntr(ilFCff)
                If tgClfCntr(ilFClf).iFirstCgf = -1 Then
                    If (llFStartDate > llSDate) And (llFStartDate < llEDate) Then
                        llSDate = llFStartDate
                    End If
                    If (llFEndDate < llEDate) Then
                        llEDate = llFEndDate
                    End If
                    llSDate = llSDate + llTStartDate - llFStartDate
                    llEDate = llEDate + llTStartDate - llFStartDate
                    gPackDateLong llSDate, tgCffCntr(ilTCff).CffRec.iStartDate(0), tgCffCntr(ilTCff).CffRec.iStartDate(1)
                    tgCffCntr(ilTCff).lStartDate = llSDate
                    gPackDateLong llEDate, tgCffCntr(ilTCff).CffRec.iEndDate(0), tgCffCntr(ilTCff).CffRec.iEndDate(1)
                    tgCffCntr(ilTCff).lEndDate = llEDate
                End If
                tgCffCntr(ilTCff).iNextCff = -1
                ReDim Preserve tgCffCntr(0 To UBound(tgCffCntr) + 1)
            End If
        End If
        ilFCff = tgCffCntr(ilFCff).iNextCff
    Loop
    ilFCgf = tgClfCntr(ilFClf).iFirstCgf
    Do While ilFCgf <> -1
        gUnpackDateLong tgCgfCntr(ilFCgf).CgfRec.iAirDate(0), tgCgfCntr(ilFCgf).CgfRec.iAirDate(1), llSDate
        If (llSDate >= llFStartDate) And (llSDate <= llFEndDate) Then
            If ((tgCgfCntr(ilFCgf).iStatus = 0) Or (tgCgfCntr(ilFCgf).iStatus = 1)) And (tgCgfCntr(ilFCgf).CgfRec.iNoSpots > 0) Then
                ilTCgf = UBound(tgCgfCntr)
                If tgClfCntr(ilLnRowNo - 1).iFirstCgf = -1 Then
                    tgClfCntr(ilLnRowNo - 1).iFirstCgf = ilTCgf
                Else
                    tgCgfCntr(ilTCgf - 1).iNextCgf = ilTCgf
                End If
                tgCgfCntr(ilTCgf) = tgCgfCntr(ilFCgf)
                tgCgfCntr(ilTCgf).iNextCgf = -1
                ReDim Preserve tgCgfCntr(0 To UBound(tgCgfCntr) + 1)
            End If
        End If
        ilFCgf = tgCgfCntr(ilFCgf).iNextCgf
    Loop
    tgClfCntr(ilLnRowNo).iStatus = -1
    tgClfCntr(ilLnRowNo).lRecPos = 0
    tgClfCntr(ilLnRowNo).iFirstCff = -1
    tgClfCntr(ilLnRowNo).iFirstCgf = -1
    tgClfCntr(ilLnRowNo).lFirstAvail = -1

End Sub
