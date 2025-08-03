VERSION 5.00
Begin VB.Form CFlight 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4275
   ClientLeft      =   2550
   ClientTop       =   2730
   ClientWidth     =   9330
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
   ScaleHeight     =   4275
   ScaleWidth      =   9330
   Begin VB.ListBox lbcPriceType 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      ItemData        =   "CFlight.frx":0000
      Left            =   3480
      List            =   "CFlight.frx":0002
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1035
      Visible         =   0   'False
      Width           =   1845
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
      Height          =   75
      Left            =   9735
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3435
      Width           =   45
   End
   Begin VB.PictureBox pbcDGWkly 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   4710
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   945
      Visible         =   0   'False
      Width           =   240
      Begin VB.CheckBox ckcDGDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   30
         TabIndex        =   20
         Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.ListBox lbcDWP 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "CFlight.frx":0004
      Left            =   5250
      List            =   "CFlight.frx":0006
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ListBox lbcDGRun 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "CFlight.frx":0008
      Left            =   2550
      List            =   "CFlight.frx":000A
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmcDGDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1635
      Picture         =   "CFlight.frx":000C
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   795
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDGEdit 
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
      Left            =   2895
      MaxLength       =   2
      TabIndex        =   8
      Top             =   870
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox edcDGDropDown 
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
      Left            =   585
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   7245
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "CFlight.frx":0106
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "CFlight.frx":0DC4
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.ListBox lbcDW 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "CFlight.frx":10CE
      Left            =   5550
      List            =   "CFlight.frx":10D0
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2655
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ListBox lbcPriceType 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   1
      ItemData        =   "CFlight.frx":10D2
      Left            =   5235
      List            =   "CFlight.frx":10D4
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1650
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox pbcAirDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5265
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2250
      Visible         =   0   'False
      Width           =   240
      Begin VB.CheckBox ckcAirDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   30
         TabIndex        =   33
         Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.TextBox edcAirSpots 
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
      Left            =   4260
      MaxLength       =   4
      TabIndex        =   31
      Top             =   1890
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox edcAirDaily 
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
      Left            =   3615
      MaxLength       =   2
      TabIndex        =   30
      Top             =   1845
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmcAirDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4665
      Picture         =   "CFlight.frx":10D6
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2610
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcAirDropDown 
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
      Left            =   3660
      MaxLength       =   20
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2610
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pbcAirWeek 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3600
      ScaleHeight     =   210
      ScaleWidth      =   435
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2175
      Visible         =   0   'False
      Width           =   435
   End
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
      Left            =   555
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "CFlight.frx":11D0
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   13
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
            TabIndex        =   14
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
         TabIndex        =   10
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
         TabIndex        =   12
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
         Left            =   315
         TabIndex        =   11
         Top             =   15
         Width           =   1305
      End
   End
   Begin VB.PictureBox pbcAirWk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   210
      Picture         =   "CFlight.frx":3FEA
      ScaleHeight     =   2175
      ScaleWidth      =   7095
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1350
      Visible         =   0   'False
      Width           =   7125
   End
   Begin VB.PictureBox pbcDGen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   255
      Picture         =   "CFlight.frx":14E18
      ScaleHeight     =   585
      ScaleWidth      =   8865
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   615
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.PictureBox pbcAir 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      Height          =   4245
      Left            =   15
      ScaleHeight     =   4215
      ScaleWidth      =   9285
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   9315
      Begin VB.Timer tmcStart 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   9060
         Top             =   3720
      End
      Begin VB.CommandButton cmcAirDone 
         Appearance      =   0  'Flat
         Caption         =   "&Done"
         Height          =   285
         Left            =   7950
         TabIndex        =   37
         Top             =   3300
         Width           =   840
      End
      Begin VB.VScrollBar vbcAirWk 
         Height          =   1980
         LargeChange     =   7
         Left            =   7455
         TabIndex        =   36
         Top             =   1395
         Width           =   255
      End
      Begin VB.PictureBox pbcAirSTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Height          =   105
         Left            =   45
         ScaleHeight     =   105
         ScaleWidth      =   75
         TabIndex        =   23
         Top             =   915
         Width           =   75
      End
      Begin VB.PictureBox pbcAirTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Height          =   15
         Left            =   585
         ScaleHeight     =   15
         ScaleWidth      =   30
         TabIndex        =   35
         Top             =   2775
         Width           =   30
      End
      Begin VB.CommandButton cmcAirCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Height          =   285
         Left            =   7965
         TabIndex        =   38
         Top             =   3705
         Width           =   840
      End
      Begin VB.PictureBox plcStartDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   705
         Left            =   165
         ScaleHeight     =   645
         ScaleWidth      =   8940
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   540
         Width           =   9000
      End
      Begin VB.PictureBox pbcDGSTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Height          =   105
         Left            =   0
         ScaleHeight     =   105
         ScaleWidth      =   90
         TabIndex        =   3
         Top             =   60
         Width           =   90
      End
      Begin VB.PictureBox pbcDGTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   60
         TabIndex        =   22
         Top             =   765
         Width           =   60
      End
      Begin VB.PictureBox plcAirScreen 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Height          =   195
         Left            =   30
         ScaleHeight     =   195
         ScaleWidth      =   2160
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   2160
      End
      Begin VB.CommandButton cmcGenNew 
         Appearance      =   0  'Flat
         Caption         =   "Update &Weeks"
         Height          =   285
         Left            =   7770
         TabIndex        =   39
         Top             =   1320
         Width           =   1380
      End
      Begin VB.CommandButton cmcGenAll 
         Appearance      =   0  'Flat
         Caption         =   "Update &Weeks"
         Height          =   285
         Left            =   7725
         TabIndex        =   40
         Top             =   1335
         Width           =   1515
      End
      Begin VB.PictureBox pbcAirDemo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   570
         Picture         =   "CFlight.frx":1A8DA
         ScaleHeight     =   570
         ScaleWidth      =   6510
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   3540
         Visible         =   0   'False
         Width           =   6510
      End
      Begin VB.CommandButton cmcAirClear 
         Appearance      =   0  'Flat
         Caption         =   "Clear"
         Height          =   285
         Left            =   7725
         TabIndex        =   41
         Top             =   1710
         Width           =   1515
      End
      Begin VB.ComboBox cbcAirLine 
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
         Height          =   315
         Left            =   2385
         TabIndex        =   2
         Top             =   90
         Width           =   6810
      End
      Begin VB.PictureBox pbcWkArrow 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Left            =   45
         Picture         =   "CFlight.frx":1EBD8
         ScaleHeight     =   180
         ScaleWidth      =   105
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1890
         Visible         =   0   'False
         Width           =   105
      End
   End
End
Attribute VB_Name = "CFlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CFlight.FRM on Wed 6/17/09 @ 12:56 PM
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

Const DPINDEX = 2
Const PRICEINDEX = 20

Const DGSTARTDATEINDEX = 1
Const DGNOWKSINDEX = 2
Const DGENDDATEINDEX = 3
Const DGRUNINDEX = 4
Const DGDWPINDEX = 5
Const DGSPOTWKINDEX = 6
Const DGDAILYINDEX = 7
Const DGPRICEINDEX = 14
'Const DGPREEMPTINDEX = 15
Const DGOVSTARTINDEX = 15   '16
Const DGOVENDINDEX = 16 '17

Const AIRWEEKINDEX = 1
Const WEEKNOINDEX = 2
Const WEEKSTARTDATEINDEX = 3
Const WEEKENDDATEINDEX = 4
Const DWINDEX = 5
Const SPOTPERWKINDEX = 6
Const DAILYSPOTSINDEX = 7
Const WKPRICEINDEX = 14


Private Sub cbcAirLine_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Dim slLineNo As String
    Dim ilLineNo As Integer
    Dim ilClf As Integer
    ReDim ilFlDays(0 To 6) As Integer
    Screen.MousePointer = vbHourglass
    mSetDate False
    tgClfCntr(imLnRowNo - 1).iResearch = False
    Contract.mPrgPop imLnRowNo, "L"
    mGetFlightDays imLnRowNo, ilFlDays()
    slStr = mSetDPShowName(imLnRowNo, DPINDEX, ilFlDays())
    Contract.mBuildWkInfo False, False, False
    imLnSave(8, imLnRowNo) = imDGSave(2)
    imSvLnChg = imLnChg
            'mGetFlightStartInfo True, False
    mGetLnPriceInfo imLnRowNo, slStr, imLnTotalType, PRICEINDEX         'Formats Price within Show
    Contract.mCntrTotal
    'mSetFlightEndInfo
    'Determine imLnRowNo
    slStr = cbcAirLine.List(cbcAirLine.ListIndex)
    ilRet = gParseItem(slStr, 1, " ", slLineNo)
    ilLineNo = Val(slLineNo)
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) Step 1
        If tgClfCntr(ilClf).ClfRec.iLine = ilLineNo Then
            imLnRowNo = ilClf + 1
            Exit For
        End If
    Next ilClf
    smLineType = mGetLineType(imLnRowNo)
    tgSpf.sAllowDailyBuys = smSvAllowDailyWeekly
    If mIsLnGame(imLnRowNo) Then
        tgSpf.sAllowDailyBuys = "N"
    End If
    mGetFlightStartInfo False, False
    Screen.MousePointer = vbDefault
End Sub

Private Sub cbcAirLine_Click()
    cbcAirLine_Change
End Sub

Private Sub cbcAirLine_GotFocus()
    mDGSetShow imDGBoxNo
    imDGBoxNo = -1
    mAirSetShow imAirBoxNo
    imAirBoxNo = -1
    pbcWkArrow.Visible = False
End Sub

Private Sub ckcAirDay_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAirDay.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    If Value Then
        If Not imSettingDay Then
            If (imStartDayIndex < 0) Or (imEndDayIndex < 0) Then
                If (imAirBoxNo - DAILYSPOTSINDEX > -imEndDayIndex) And (imAirBoxNo - DAILYSPOTSINDEX < -imStartDayIndex) Then
                    Beep
                    ckcAirDay.Value = vbUnchecked
                    Exit Sub
                End If
            Else
                If (imAirBoxNo - DAILYSPOTSINDEX < imStartDayIndex) Or (imAirBoxNo - DAILYSPOTSINDEX > imEndDayIndex) Then
                    Beep
                    ckcAirDay.Value = vbUnchecked
                    Exit Sub
                End If
            End If
            If Not imAllowedDays(imAirBoxNo - DAILYSPOTSINDEX + LBound(imAllowedDays)) Then
                Beep
                ckcAirDay.Value = vbUnchecked
                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub ckcAirDay_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcDGDay_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcDGDay.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    If Value Then
        If Not imSettingDay Then
            If (imStartDayIndex < 0) Or (imEndDayIndex < 0) Then
                If (imDGBoxNo - DGDAILYINDEX > -imEndDayIndex) And (imDGBoxNo - DGDAILYINDEX < -imStartDayIndex) Then
                    Beep
                    ckcDGDay.Value = vbUnchecked
                    Exit Sub
                End If
            Else
                If (imDGBoxNo - DGDAILYINDEX < imStartDayIndex) Or (imDGBoxNo - DGDAILYINDEX > imEndDayIndex) Then
                    Beep
                    ckcDGDay.Value = vbUnchecked
                    Exit Sub
                End If
            End If
            If Not imAllowedDays(imDGBoxNo - DGDAILYINDEX + LBound(imAllowedDays)) Then
                Beep
                ckcDGDay.Value = vbUnchecked
                Exit Sub
            End If
        End If
    End If
    mDGSetCommands

End Sub

Private Sub ckcDGDay_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcAirDropDown_Click()
    Select Case imAirBoxNo
        Case DWINDEX
            lbcDW.Visible = Not lbcDW.Visible
        Case WKPRICEINDEX
            lbcPriceType(1).Visible = Not lbcPriceType(1).Visible
    End Select
    edcAirDropDown.SelStart = 0
    edcAirDropDown.SelLength = Len(edcAirDropDown.Text)
    edcAirDropDown.SetFocus

End Sub

Private Sub cmcAirDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    'If plcPower.Visible Then
    '    If (imPBDateIndex = 0) Or (imPBDateIndex = 1) Then
    '        edcPBDate(imPBDateIndex).SelStart = 0
    '        edcPBDate(imPBDateIndex).SelLength = Len(edcPBDate(imPBDateIndex).Text)
    '        imBypassFocus = True
    '        edcPBDate(imPBDateIndex).SetFocus
    '    End If
    'Else
    '    If imNoBoxNo = NODATEINDEX Then
    '        edcShareDropDown.SelStart = 0
    '        edcShareDropDown.SelLength = Len(edcShareDropDown.Text)
    '        edcShareDropDown.SetFocus
    '    Else
            edcDGDropDown.SelStart = 0
            edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
            edcDGDropDown.SetFocus
    '    End If
    'End If
End Sub


Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    'If plcPower.Visible Then
    '    If (imPBDateIndex = 0) Or (imPBDateIndex = 1) Then
    '        edcPBDate(imPBDateIndex).SelStart = 0
    '        edcPBDate(imPBDateIndex).SelLength = Len(edcPBDate(imPBDateIndex).Text)
    '        imBypassFocus = True
    '        edcPBDate(imPBDateIndex).SetFocus
    '    End If
    'Else
    '    If imNoBoxNo = NODATEINDEX Then
    '        edcShareDropDown.SelStart = 0
    '        edcShareDropDown.SelLength = Len(edcShareDropDown.Text)
    '        edcShareDropDown.SetFocus
    '    Else
            edcDGDropDown.SelStart = 0
            edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
            edcDGDropDown.SetFocus
    '    End If
    'End If

End Sub

Private Sub cmcDGDropDown_Click()
    'If (imLnBoxNo <> OV1INDEX) And (imLnBoxNo <> OV2INDEX) Then
        '11/2/9: moved to CFlight
        Select Case imDGBoxNo
            Case DGSTARTDATEINDEX
                plcCalendar.Visible = Not plcCalendar.Visible
            Case DGNOWKSINDEX
                edcDGEdit.SelStart = 0
                edcDGEdit.SelLength = Len(edcDGEdit.Text)
                edcDGEdit.SetFocus
                Exit Sub
            Case DGENDDATEINDEX
                plcCalendar.Visible = Not plcCalendar.Visible
            Case DGRUNINDEX
                lbcDGRun.Visible = Not lbcDGRun.Visible
            Case DGDWPINDEX
                lbcDWP.Visible = Not lbcDWP.Visible
            Case DGSPOTWKINDEX
                edcDGEdit.SelStart = 0
                edcDGEdit.SelLength = Len(edcDGEdit.Text)
                edcDGEdit.SetFocus
                Exit Sub
            Case DGPRICEINDEX
                lbcPriceType(0).Visible = Not lbcPriceType(0).Visible
            Case DGOVSTARTINDEX
                plcTme.Visible = Not plcTme.Visible
            Case DGOVENDINDEX
                plcTme.Visible = Not plcTme.Visible
        End Select
    'Else
        'If imOVType <> 2 Then
        '    Select Case imOVBoxNo
        '        Case OVSTARTTIMEINDEX
        '            plcTme.Visible = Not plcTme.Visible
        '        Case OVENDTIMEINDEX
        '            plcTme.Visible = Not plcTme.Visible
        '        Case OVBBOPENINDEX
        '            lbcBBLen.Visible = Not lbcBBLen.Visible
        '        Case OVBBCLOSEINDEX
        '            lbcBBLen.Visible = Not lbcBBLen.Visible
        '    End Select
        'Else
        '    Select Case imOVBoxNo
        '        Case OV2STARTTIMEINDEX
        '            plcTme.Visible = Not plcTme.Visible
        '        Case OV2ENDTIMEINDEX
        '            plcTme.Visible = Not plcTme.Visible
        '        Case OV2SPLITINDEX
        '            lbcRegion.Visible = Not lbcRegion.Visible
        '    End Select
        'End If
    'End If
    edcDGDropDown.SelStart = 0
    edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
    edcDGDropDown.SetFocus

End Sub

Private Sub cmcDGDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcGenAll_Click()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilPaint As Integer
    If gDateValue(smDGSave(11)) < gDateValue(smDGSave(10)) Then
        Beep
        MsgBox "End Date must be after or equal to " & smDGSave(10), vbOKOnly + vbExclamation, "Warning"
        imDGBoxNo = DGENDDATEINDEX
        mDGEnableBox imDGBoxNo
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mDGSetShow imDGBoxNo
    imDGBoxNo = -1
    ilPaint = False
    If (imDGSave(2) = 0) Or (imDGSave(2) = 2) Then  'Change daily to weekly
        For ilLoop = 0 To 6 Step 1
            If tmDGCtrls(DGDAILYINDEX + ilLoop).sShow = " R" Then
                slStr = ""
                gSetShow pbcDGen, slStr, tmDGCtrls(DGDAILYINDEX + ilLoop)
                ilPaint = True
            End If
        Next ilLoop
    Else
        For ilLoop = 0 To 6 Step 1
            If tmDGCtrls(DGDAILYINDEX + ilLoop).sShow = " R" Then
                slStr = "4"
                gSetShow pbcDGen, slStr, tmDGCtrls(DGDAILYINDEX + ilLoop)
                ilPaint = True
            End If
        Next ilLoop
    End If
    If ilPaint Then
        pbcDGen_Paint
    End If
    mCreateDate False
    'Reset price
    If (smLineType = "O") Or (smLineType = "A") Or (smLineType = "E") Then
        mSetAirPkPrice
    End If
    mFlightTotals
    mGetDemoFlightTotals
    imAirRowNo = 1
    pbcAirWk.Cls
    pbcAirWk_Paint
    pbcAirSTab.SetFocus
    cmcGenNew.Visible = True
    cmcGenNew.Enabled = True
    cmcGenAll.Visible = False
    mDGSetCommands
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcGenAll_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcAirDaily_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcAirDaily_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    If (KeyAscii = KEYBACKSPACE) Then
        Exit Sub
    End If
    If Not imSettingDay Then
        'Allow user to enter a zero as replacement on illegal day
        If (KeyAscii <> KEY0) Or (Len(edcAirDaily.Text) = 0) Or (Len(edcAirDaily.Text) <> edcAirDaily.SelLength) Then
            If (imStartDayIndex < 0) Or (imEndDayIndex < 0) Then
                If (imAirBoxNo - DAILYSPOTSINDEX > -imEndDayIndex) And (imAirBoxNo - DAILYSPOTSINDEX < -imStartDayIndex) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            Else
                If (imAirBoxNo - DAILYSPOTSINDEX < imStartDayIndex) Or (imAirBoxNo - DAILYSPOTSINDEX > imEndDayIndex) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            If Not imAllowedDays(imAirBoxNo - DAILYSPOTSINDEX + LBound(imAllowedDays)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcAirDaily.Text
    slStr = Left$(slStr, edcAirDaily.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcAirDaily.SelStart - edcAirDaily.SelLength)
    If gCompNumberStr(slStr, "99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcAirDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imAirBoxNo
        Case DWINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcAirDropDown, lbcDW, imBSMode, imComboBoxIndex
            If tgSpf.sAllowDailyBuys <> "Y" Then
                slStr = edcAirDropDown.Text
                If InStr(1, slStr, "Daily", 1) > 0 Then
                    edcAirDropDown.Text = "Weekly"   'Force to weekly
                    gMatchLookAhead edcAirDropDown, lbcDW, imBSMode, imComboBoxIndex
                End If
            End If
            imLbcArrowSetting = False
        Case WKPRICEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcAirDropDown, lbcPriceType(1), imBSMode, slStr)
            imLbcArrowSetting = False
    End Select
End Sub
Private Sub edcAirDropDown_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcAirDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcAirDropDown_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    Dim slMaxValue As String

    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If ActiveControl.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imAirBoxNo
        Case DWINDEX
        Case WKPRICEINDEX
            slStr = edcAirDropDown.Text
            slStr = Left$(slStr, edcAirDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcAirDropDown.SelStart - edcAirDropDown.SelLength)
            If (Left$(slStr, 1) < "0") And (Left$(slStr, 1) > "9") Then
                If (KeyAscii = Asc("N")) Or (KeyAscii = Asc("n")) Then
                    lbcPriceType(1).ListIndex = 1
                    KeyAscii = 0
                    Exit Sub
                End If
                If (KeyAscii = Asc("M")) Or (KeyAscii = Asc("m")) Then
                    lbcPriceType(1).ListIndex = 2
                    KeyAscii = 0
                    Exit Sub
                End If
                If (KeyAscii = Asc("B")) Or (KeyAscii = Asc("b")) Then
                    lbcPriceType(1).ListIndex = 3
                    KeyAscii = 0
                    Exit Sub
                End If
                If (KeyAscii = Asc("S")) Or (KeyAscii = Asc("s")) Then
                    lbcPriceType(1).ListIndex = 4
                    KeyAscii = 0
                    Exit Sub
                End If
                If (KeyAscii = Asc("P")) Or (KeyAscii = Asc("p")) Then
                    lbcPriceType(1).ListIndex = 5
                    KeyAscii = 0
                    Exit Sub
                End If
                If (KeyAscii = Asc("R")) Or (KeyAscii = Asc("r")) Then
                    lbcPriceType(1).ListIndex = 6
                    KeyAscii = 0
                    Exit Sub
                End If
                If (KeyAscii = Asc("A")) Or (KeyAscii = Asc("a")) Then
                    lbcPriceType(1).ListIndex = 7
                    KeyAscii = 0
                    Exit Sub
                End If
                Exit Sub
            End If
            imChgMode = True
            lbcPriceType(1).ListIndex = 0
            If Trim$(edcAirDropDown.Text) > "" Then
                If (Left$(edcAirDropDown.Text, 1) < "0") Or (Left$(edcAirDropDown.Text, 1) > "9") Then
                    edcAirDropDown.Text = ""
                End If
            End If
            imChgMode = False
            ilPos = InStr(edcAirDropDown.SelText, ".")
            If ilPos = 0 Then
                ilPos = InStr(edcAirDropDown.Text, ".")    'Disallow multi-decimal points
                If ilPos > 0 Then
                    If KeyAscii = KEYDECPOINT Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcAirDropDown.Text
            slStr = Left$(slStr, edcAirDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcAirDropDown.SelStart - edcAirDropDown.SelLength)
            slMaxValue = "9999999.99"
            If Trim(tmFTCtrls(2).sShow) <> "" Then
                If gCompNumberStr(tmFTCtrls(2).sShow, "0") <> 0 Then
                    slMaxValue = gDivStr("21000000", tmFTCtrls(2).sShow)
                    If gCompNumberStr(slMaxValue, "9999999.99") > 0 Then
                        slMaxValue = "9999999.99"
                    End If
                End If
            End If
            If gCompNumberStr(slStr, slMaxValue) > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub
Private Sub edcAirDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imAirBoxNo
            Case DWINDEX
                If tgSpf.sAllowDailyBuys = "Y" Then
                    gProcessArrowKey Shift, KeyCode, lbcDW, imLbcArrowSetting
                End If
            Case WKPRICEINDEX
                gProcessArrowKey Shift, KeyCode, lbcPriceType(1), imLbcArrowSetting
        End Select
        edcAirDropDown.SelStart = 0
        edcAirDropDown.SelLength = Len(edcAirDropDown.Text)
    End If
End Sub

Private Sub edcAirSpots_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcAirSpots_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim slMaxValue As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcAirSpots.Text
    If (Trim$(slStr) = "") And (KeyAscii = KEY0) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = Left$(slStr, edcAirSpots.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcAirSpots.SelStart - edcAirSpots.SelLength)
    slMaxValue = "999"
    If Trim(tmFTCtrls(3).sShow) <> "" Then
        If gCompNumberStr(tmFTCtrls(3).sShow, "0") <> 0 Then
            slMaxValue = gDivStr("21000000", tmFTCtrls(3).sShow)
            If gCompNumberStr(slMaxValue, "999") > 0 Then
                slMaxValue = "999"
            End If
        End If
    End If
    If gCompNumberStr(slStr, slMaxValue) > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcDGEdit_Change()
    Dim llWkEndDate As Long
    Dim slStr As String
    Dim slEnd As String
    If imDGChgMode Then
        Exit Sub
    End If
    imDGChgMode = True
    Select Case imDGBoxNo
        Case DGNOWKSINDEX
            If (Trim$(edcDGEdit.Text) = "") Or (edcDGEdit.Text = "0") Then
                edcDGEdit.Text = "1"
            End If
            smDGSave(1) = edcDGEdit.Text
            If (smDGSave(11) = "") Or (Not gValidDate(smDGSave(11))) Then
                llWkEndDate = gDateValue(smDGSave(10)) + 7 * Val(edcDGEdit.Text) - 1
            Else
                slEnd = gAdjYear(smDGSave(11))
                llWkEndDate = gDateValue(smDGSave(10)) + 7 * Val(edcDGEdit.Text) - 7
                Do Until gWeekDayLong(llWkEndDate) = gWeekDayStr(slEnd)
                    llWkEndDate = llWkEndDate + 1
                Loop
            End If
            smDGSave(11) = Format$(llWkEndDate, "m/d/yy")
            slStr = gAddDayToDate(smDGSave(11))
            gSetShow pbcDGen, slStr, tmDGCtrls(DGENDDATEINDEX)
            pbcDGen.Cls
            pbcDGen_Paint
    End Select
    mDGSetCommands
    imDGChgMode = False
End Sub
Private Sub edcDGEdit_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDGEdit_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim slMaxValue As String

    Select Case imDGBoxNo
        Case DGNOWKSINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDGEdit.Text
            If (Trim$(slStr) = "") And (KeyAscii = KEY0) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = Left$(slStr, edcDGEdit.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDGEdit.SelStart - edcDGEdit.SelLength)
            If gCompNumberStr(slStr, "0") = 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            If gCompNumberStr(slStr, "99") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case DGSPOTWKINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDGEdit.Text
            slStr = Left$(slStr, edcDGEdit.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDGEdit.SelStart - edcDGEdit.SelLength)
            slMaxValue = "999"
            If Trim(smLnSave(18, imLnRowNo)) <> "" Then
                If gCompNumberStr(smLnSave(18, imLnRowNo), "0") <> 0 Then
                    slMaxValue = gDivStr("21000000", smLnSave(18, imLnRowNo))
                    If gCompNumberStr(slMaxValue, "999") > 0 Then
                        slMaxValue = "999"
                    End If
                End If
            End If
            If gCompNumberStr(slStr, slMaxValue) > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case DGDAILYINDEX To DGDAILYINDEX + 6
            If (KeyAscii = KEYBACKSPACE) Then
                Exit Sub
            End If
            If Not imSettingDay Then
                'Allow user to enter a zero as replacement on illegal day
                If (KeyAscii <> KEY0) Or (Len(edcDGEdit.Text) = 0) Or (Len(edcDGEdit.Text) <> edcDGEdit.SelLength) Then
                    If (imStartDayIndex < 0) Or (imEndDayIndex < 0) Then
                        If (imDGBoxNo - DGDAILYINDEX > -imEndDayIndex) And (imDGBoxNo - DGDAILYINDEX < -imStartDayIndex) Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Else
                        If (imDGBoxNo - DGDAILYINDEX < imStartDayIndex) Or (imDGBoxNo - DGDAILYINDEX > imEndDayIndex) Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                    If Not imAllowedDays(imDGBoxNo - DGDAILYINDEX + LBound(imAllowedDays)) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDGEdit.Text
            slStr = Left$(slStr, edcDGEdit.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDGEdit.SelStart - edcDGEdit.SelLength)
            If gCompNumberStr(slStr, "99") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
    'If (imDGBoxNo = DGNOWKSINDEX) Or (imDGBoxNo = DGSPOTWKINDEX) Then
    '    gNumOutLine KeyAscii, imcNumOutline
    'End If
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
    pbcDGSTab.Visible = False
    pbcDGTab.Visible = False
    pbcAirSTab.Visible = False
    pbcAirTab.Visible = False
    pbcDGSTab.Visible = True
    pbcDGTab.Visible = True
    pbcAirSTab.Visible = True
    pbcAirTab.Visible = True
    
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
    pbcWkArrow.Picture = IconTraf!imcArrow.Picture
    pbcWkArrow.Width = 90
    pbcWkArrow.height = 165
    
    lbcDGRun.AddItem "Week"
    lbcDGRun.AddItem "2nd"
    lbcDGRun.AddItem "3rd"
    lbcDGRun.AddItem "4th"
    lbcDWP.Clear
    lbcDWP.AddItem "Daily"
    lbcDWP.AddItem "Weekly"
    lbcDW.Clear
    lbcDW.AddItem "Daily"
    lbcDW.AddItem "Weekly"
    lbcPriceType(0).AddItem "[Dollars]"
    lbcPriceType(0).AddItem "Rate Card"
    lbcPriceType(0).AddItem "Avg Rate Card"
    lbcPriceType(0).AddItem "N/C"
    lbcPriceType(0).AddItem "MG"
    lbcPriceType(0).AddItem "Bonus"
    lbcPriceType(0).AddItem "Spinoff"
    lbcPriceType(0).AddItem "Package"
    lbcPriceType(0).AddItem "Recapturable"
    lbcPriceType(0).AddItem "ADU"
    lbcPriceType(1).AddItem "[Dollars]"
    lbcPriceType(1).AddItem "N/C"
    lbcPriceType(1).AddItem "MG"
    lbcPriceType(1).AddItem "Bonus"
    lbcPriceType(1).AddItem "Spinoff"
    lbcPriceType(1).AddItem "Package"
    lbcPriceType(1).AddItem "Recapturable"
    lbcPriceType(1).AddItem "ADU"
    gCenterModalForm CFlight
    tmcStart.Enabled = True
    
    'mGetFlightStartInfo True, False 'Require to call mGetDate as cbcAirLine calls mSetDate
    'mPopAirLine
    'gCenterModalForm CFlight
    'Screen.MousePointer = vbDefault
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
    Unload CFlight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CFlight = Nothing   'Remove data segment
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mDGEnableBox                    *
'*                                                     *
'*             Created:8/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mDGEnableBox(ilBoxNo As Integer)
'
'   mDGEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilBox As Integer
    Dim ilStartDayIndex As Integer
    Dim ilEndDayIndex As Integer
    'If (ilBoxNo < LBound(tmDGCtrls)) Or (ilBoxNo > UBound(tmDGCtrls)) Then
    If (ilBoxNo < imLBCTCtrls) Or (ilBoxNo > UBound(tmDGCtrls)) Then
        Exit Sub
    End If
    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        cmcAirCancel.SetFocus   'pbcClickFocus.SetFocus
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case DGSTARTDATEINDEX 'Start Date
            'gShowHelpMess tmClfHelp(), CLFSTARTDATE
            edcDGDropDown.Width = tmDGCtrls(DGSTARTDATEINDEX).fBoxW - cmcDGDropDown.Width
            edcDGDropDown.MaxLength = 10
            gMoveTableCtrl pbcDGen, edcDGDropDown, tmDGCtrls(DGSTARTDATEINDEX).fBoxX, tmDGCtrls(DGSTARTDATEINDEX).fBoxY
            edcDGDropDown.Text = smDGSave(10)
            cmcDGDropDown.Move edcDGDropDown.Left + edcDGDropDown.Width, edcDGDropDown.Top
            plcCalendar.Move edcDGDropDown.Left, edcDGDropDown.Top + edcDGDropDown.height
            gObtainMonthYear imCalType, smDGSave(10), imCalMonth, imCalYear
            edcDGDropDown.SelStart = 0
            edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            plcCalendar.Visible = True
            edcDGDropDown.SetFocus
        Case DGNOWKSINDEX 'Number of weeks
            'gShowHelpMess tmClfHelp(), CLFNOWKS
            edcDGEdit.Width = tmDGCtrls(DGNOWKSINDEX).fBoxW ' - cmcDGDropDown.Width / 2
            edcDGEdit.MaxLength = 2
            edcDGEdit.Text = smDGSave(1)
            gMoveTableCtrl pbcDGen, edcDGEdit, tmDGCtrls(DGNOWKSINDEX).fBoxX, tmDGCtrls(DGNOWKSINDEX).fBoxY
            edcDGEdit.Visible = True
            edcDGEdit.SetFocus
        Case DGENDDATEINDEX 'End Date
            'gShowHelpMess tmClfHelp(), CLFENDDATE
            edcDGDropDown.Width = tmDGCtrls(DGENDDATEINDEX).fBoxW - cmcDGDropDown.Width
            edcDGDropDown.MaxLength = 10
            gMoveTableCtrl pbcDGen, edcDGDropDown, tmDGCtrls(DGENDDATEINDEX).fBoxX, tmDGCtrls(DGENDDATEINDEX).fBoxY
            edcDGDropDown.Text = smDGSave(11)
            cmcDGDropDown.Move edcDGDropDown.Left + edcDGDropDown.Width, edcDGDropDown.Top
            plcCalendar.Move edcDGDropDown.Left, edcDGDropDown.Top + edcDGDropDown.height
            gObtainMonthYear imCalType, smDGSave(11), imCalMonth, imCalYear
            edcDGDropDown.SelStart = 0
            edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            plcCalendar.Visible = True
            edcDGDropDown.SetFocus
        Case DGRUNINDEX 'Run the weeks of...
            'gShowHelpMess tmClfHelp(), CLFRUNEVERY
            lbcDGRun.height = gListBoxHeight(lbcDGRun.ListCount, 4)
            edcDGDropDown.Width = tmDGCtrls(DGRUNINDEX).fBoxW - cmcDGDropDown.Width
            edcDGDropDown.MaxLength = 4
            gMoveTableCtrl pbcDGen, edcDGDropDown, tmDGCtrls(DGRUNINDEX).fBoxX, tmDGCtrls(DGRUNINDEX).fBoxY
            cmcDGDropDown.Move edcDGDropDown.Left + edcDGDropDown.Width, edcDGDropDown.Top
            imChgMode = True
            If imDGSave(1) >= 0 Then
                lbcDGRun.ListIndex = imDGSave(1)
            Else
                lbcDGRun.ListIndex = -1
            End If
            imComboBoxIndex = lbcDGRun.ListIndex
            If lbcDGRun.ListIndex >= 0 Then
                edcDGDropDown.Text = lbcDGRun.List(lbcDGRun.ListIndex)
            Else
                edcDGDropDown.Text = ""
            End If
            imChgMode = False
            lbcDGRun.Move edcDGDropDown.Left, edcDGDropDown.Top + edcDGDropDown.height
            edcDGDropDown.SelStart = 0
            edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            edcDGDropDown.SetFocus
        Case DGDWPINDEX
            lbcDWP.height = gListBoxHeight(lbcDWP.ListCount, 4)
            edcDGDropDown.Width = tmDGCtrls(DGDWPINDEX).fBoxW
            edcDGDropDown.MaxLength = 6
            gMoveTableCtrl pbcDGen, edcDGDropDown, tmDGCtrls(DGDWPINDEX).fBoxX, tmDGCtrls(DGDWPINDEX).fBoxY
            cmcDGDropDown.Move edcDGDropDown.Left + edcDGDropDown.Width, edcDGDropDown.Top
            imChgMode = True
            If imDGSave(2) >= 0 Then
                lbcDWP.ListIndex = imDGSave(2)
            Else
                lbcDWP.ListIndex = 1    'Weekly is default
            End If
            imComboBoxIndex = lbcDWP.ListIndex
            If lbcDWP.ListIndex >= 0 Then
                edcDGDropDown.Text = lbcDWP.List(lbcDWP.ListIndex)
            Else
                edcDGDropDown.Text = ""
            End If
            imChgMode = False
            lbcDWP.Move edcDGDropDown.Left, edcDGDropDown.Top + edcDGDropDown.height
            edcDGDropDown.SelStart = 0
            edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            edcDGDropDown.SetFocus
        Case DGSPOTWKINDEX 'Spots/week index
            If tgClfCntr(imLnRowNo - 1).iGame Then
                'gShowHelpMess tmClfHelp(), CLFSPOTSWKGAME
                Exit Sub
            End If
            tmDGCtrls(DGSPOTWKINDEX).iReq = True
            'gShowHelpMess tmClfHelp(), CLFSPOTSWK
            edcDGEdit.Width = tmDGCtrls(DGSPOTWKINDEX).fBoxW ' - cmcDGDropDown.Width
            edcDGEdit.MaxLength = 4
            edcDGEdit.Text = smDGSave(2)
            gMoveTableCtrl pbcDGen, edcDGEdit, tmDGCtrls(DGSPOTWKINDEX).fBoxX, tmDGCtrls(DGSPOTWKINDEX).fBoxY
            edcDGEdit.Visible = True  'Set visibility
            edcDGEdit.SetFocus
        Case DGDAILYINDEX To DGDAILYINDEX + 6 'Day index
            If Val(smDGSave(1)) > 1 Then
                imStartDayIndex = 0
                imEndDayIndex = 6
            Else
                ilStartDayIndex = gWeekDayStr(smDGSave(10))
                ilEndDayIndex = gWeekDayStr(smDGSave(11))
                If ilStartDayIndex <= ilEndDayIndex Then
                    imStartDayIndex = ilStartDayIndex
                    imEndDayIndex = ilEndDayIndex
                Else
                    imStartDayIndex = -ilStartDayIndex
                    imEndDayIndex = -ilEndDayIndex
                End If
            End If
            ilBox = ilBoxNo
            imSettingDay = True
            If ((imDGSave(2) = 0) Or (imDGSave(2) = 2)) And (Not tgClfCntr(imLnRowNo - 1).iGame) Then
                'gShowHelpMess tmClfHelp(), CLFDAILY
                edcDGEdit.Width = tmDGCtrls(DGDAILYINDEX).fBoxW
                gMoveTableCtrl pbcDGen, edcDGEdit, tmDGCtrls(ilBoxNo).fBoxX, tmDGCtrls(ilBoxNo).fBoxY
                edcDGEdit.MaxLength = 2
                edcDGEdit.Text = smDGSave(3 + ilBoxNo - DGDAILYINDEX)
                If edcDGEdit.Text = "" Then
                    edcDGEdit.Text = "1"
                    If (imStartDayIndex < 0) Or (imEndDayIndex < 0) Then
                        If (ilBoxNo - DGDAILYINDEX > -imEndDayIndex) And (ilBoxNo - DGDAILYINDEX < -imStartDayIndex) Then
                            edcDGEdit.Text = ""
                        End If
                    Else
                        If (ilBoxNo - DGDAILYINDEX < imStartDayIndex) Or (ilBoxNo - DGDAILYINDEX > imEndDayIndex) Then
                            edcDGEdit.Text = ""
                        End If
                    End If
                    If Not imAllowedDays(ilBoxNo - DGDAILYINDEX + LBound(imAllowedDays)) Then
                        edcDGEdit.Text = ""
                    End If
                End If
                edcDGEdit.Visible = True  'Set visibility
                edcDGEdit.SetFocus
            Else
                If tgClfCntr(imLnRowNo - 1).iGame Then
                    'gShowHelpMess tmClfHelp(), CLFGAMEDAYS
                Else
                    'gShowHelpMess tmClfHelp(), CLFWEEKLY
                End If
                gMoveTableCtrl pbcDGen, pbcDGWkly, tmDGCtrls(ilBoxNo).fBoxX, tmDGCtrls(ilBoxNo).fBoxY
                If smDGSave(3 + ilBoxNo - DGDAILYINDEX) = "Y" Then
                    ckcDGDay.Value = vbChecked
                ElseIf smDGSave(3 + ilBoxNo - DGDAILYINDEX) = "N" Then
                    ckcDGDay.Value = vbUnchecked
                Else
                    ckcDGDay.Value = vbChecked
                    If (imStartDayIndex < 0) Or (imEndDayIndex < 0) Then
                        If (ilBoxNo - DGDAILYINDEX > -imEndDayIndex) And (ilBoxNo - DGDAILYINDEX < -imStartDayIndex) Then
                            ckcDGDay.Value = vbUnchecked
                        End If
                    Else
                        If (ilBoxNo - DGDAILYINDEX < imStartDayIndex) Or (ilBoxNo - DGDAILYINDEX > imEndDayIndex) Then
                            ckcDGDay.Value = vbUnchecked
                        End If
                    End If
                    If Not imAllowedDays(ilBoxNo - DGDAILYINDEX + LBound(imAllowedDays)) Then
                        ckcDGDay.Value = vbUnchecked
                    End If
                End If
                pbcDGWkly.Visible = True  'Set visibility
                ckcDGDay.SetFocus
            End If
            imSettingDay = False
        Case DGPRICEINDEX
            lbcPriceType(0).height = gListBoxHeight(lbcPriceType(0).ListCount, 5)
            edcDGDropDown.Width = tmDGCtrls(DGPRICEINDEX).fBoxW
            edcDGDropDown.MaxLength = 12
            gMoveTableCtrl pbcDGen, edcDGDropDown, tmDGCtrls(DGPRICEINDEX).fBoxX, tmDGCtrls(DGPRICEINDEX).fBoxY
            cmcDGDropDown.Move edcDGDropDown.Left + edcDGDropDown.Width, edcDGDropDown.Top
            imChgMode = True
            gFindMatch smDGSave(12), 1, lbcPriceType(0)
            If gLastFound(lbcPriceType(0)) > 0 Then
                lbcPriceType(0).ListIndex = gLastFound(lbcPriceType(0))
                edcDGDropDown.Text = lbcPriceType(0).List(lbcPriceType(0).ListIndex)
            Else
                If smDGSave(12) = "" Then
                    lbcPriceType(0).ListIndex = 1
                    edcDGDropDown.Text = lbcPriceType(0).List(1)
                Else
                    lbcPriceType(0).ListIndex = 0
                    edcDGDropDown.Text = smDGSave(12)
                End If
            End If
            imComboBoxIndex = lbcPriceType(0).ListIndex
            imChgMode = False
            lbcPriceType(0).Move edcDGDropDown.Left, edcDGDropDown.Top + edcDGDropDown.height
            edcDGDropDown.SelStart = 0
            edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            edcDGDropDown.SetFocus
        Case DGOVSTARTINDEX
            edcDGDropDown.Width = tmDGCtrls(ilBoxNo).fBoxW - cmcDGDropDown.Width
            edcDGDropDown.MaxLength = 10
            gMoveTableCtrl pbcDGen, edcDGDropDown, tmDGCtrls(ilBoxNo).fBoxX, tmDGCtrls(ilBoxNo).fBoxY
            cmcDGDropDown.Move edcDGDropDown.Left + edcDGDropDown.Width, edcDGDropDown.Top
            plcTme.Move edcDGDropDown.Left, edcDGDropDown.Top + edcDGDropDown.height
            edcDGDropDown.Text = smDGSave(13)
            edcDGDropDown.SelStart = 0
            edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
            edcDGDropDown.Visible = True  'Set visibility
            cmcDGDropDown.Visible = True
            'plcTme.Visible = True
            edcDGDropDown.SetFocus
        Case DGOVENDINDEX
            edcDGDropDown.Width = tmDGCtrls(ilBoxNo).fBoxW - cmcDGDropDown.Width
            edcDGDropDown.MaxLength = 10
            gMoveTableCtrl pbcDGen, edcDGDropDown, tmDGCtrls(ilBoxNo).fBoxX, tmDGCtrls(ilBoxNo).fBoxY
            cmcDGDropDown.Move edcDGDropDown.Left + edcDGDropDown.Width, edcDGDropDown.Top
            plcTme.Move cmcDGDropDown.Left + cmcDGDropDown.Width - plcTme.Width, edcDGDropDown.Top + edcDGDropDown.height
            edcDGDropDown.Text = smDGSave(14)
            edcDGDropDown.SelStart = 0
            edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
            edcDGDropDown.Visible = True  'Set visibility
            cmcDGDropDown.Visible = True
            'plcTme.Visible = True
            edcDGDropDown.SetFocus
    End Select
    mDGSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDGSetCommands                  *
'*                                                     *
'*             Created:10/07/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mDGSetCommands()
'
'   mDGSetCommands
'   Where:
'
    Dim ilDefined As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        cmcGenAll.Enabled = False
        cmcGenNew.Enabled = False
        cmcAirClear.Caption = "C&lear"
        cmcAirClear.Enabled = False
        Exit Sub
    End If
    slDate = gObtainNextMonday(gIncOneDay(Format$(gNow(), "m/d/yy")))
    If lmAllowedChgWkStartDate >= gDateValue(slDate) Then
        slDate = gObtainNextMonday(Format$(lmAllowedChgWkStartDate + 1, "m/d/yy"))
    End If
    'If UBound(smAirSave, 2) > LBound(smAirSave, 2) Then
    If UBound(smAirSave, 2) > imLB1Or2 Then
        'slDate = smAirSave(LBound(smAirSave, 2), 1)
        cmcAirClear.Caption = "C&lear: " & slDate
        cmcAirClear.Enabled = True
    Else
        cmcAirClear.Caption = "C&lear"
        cmcAirClear.Enabled = False
    End If
    If Not cmcGenAll.Visible Then   'cmcGenNew is visible (only add week not defined)
        Exit Sub
    End If
    ilDefined = True
    'For ilLoop = LBound(imDGSave) To UBound(imDGSave) - 2 Step 1
    For ilLoop = 1 To 2 Step 1
        If imDGSave(ilLoop) < 0 Then
            If (ilLoop = DGDWPINDEX) And (Not (tgClfCntr(imLnRowNo - 1).iGame)) Then
                ilDefined = False
                Exit For
            End If
        End If
    Next ilLoop
    If smDGSave(1) = "" Then
        ilDefined = False
    End If
    'If weekly buy and no weekly spots
    If ((imDGSave(2) = 1) Or (imDGSave(2) = 3)) And (smDGSave(2) = "") Then
        ilDefined = False
    End If
    If (Not gValidDate(smDGSave(10))) Or (Not gValidDate(smDGSave(11))) Then
        ilDefined = False
    End If
'    If (smDGSave(3) = "") And (smDGSave(4) = "") And (smDGSave(5) = "") And (smDGSave(6) = "") And (smDGSave(7) = "") And (smDGSave(8) = "") And (smDGSave(9) = "") Then
'        ilDefined = False
'    End If
    If (ilDefined) And (imUpdateAllowed) And (Not igRepNetChgAllowed) Then
        cmcGenAll.Enabled = True
    Else
        cmcGenAll.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDGSetFocus                     *
'*                                                     *
'*             Created:8/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mDGSetFocus(ilBoxNo As Integer)
'
'   mDGSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    'If (ilBoxNo < LBound(tmDGCtrls)) Or (ilBoxNo > UBound(tmDGCtrls)) Then
    If (ilBoxNo < imLBCTCtrls) Or (ilBoxNo > UBound(tmDGCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case DGSTARTDATEINDEX 'Start Date
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            edcDGDropDown.SetFocus
        Case DGNOWKSINDEX 'Number of weeks
            'cmcDGDropDown.Visible = True
            edcDGEdit.Visible = True
            edcDGEdit.SetFocus
        Case DGENDDATEINDEX 'End Date
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            edcDGDropDown.SetFocus
        Case DGRUNINDEX 'Run the weeks of...
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            edcDGDropDown.SetFocus
        Case DGDWPINDEX
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            edcDGDropDown.SetFocus
        Case DGSPOTWKINDEX 'Spots/week index
            If tgClfCntr(imLnRowNo - 1).iGame Then
                Exit Sub
            End If
            edcDGEdit.Visible = True  'Set visibility
            'cmcDGDropDown.Visible = True
            edcDGEdit.SetFocus
        Case DGDAILYINDEX To DGDAILYINDEX + 6 'Day index
            If ((imDGSave(2) = 0) Or (imDGSave(2) = 2)) And (Not tgClfCntr(imLnRowNo - 1).iGame) Then
                edcDGEdit.Visible = True  'Set visibility
                edcDGEdit.SetFocus
            Else
                pbcDGWkly.Visible = True  'Set visibility
                ckcDGDay.SetFocus
            End If
        Case DGPRICEINDEX
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            edcDGDropDown.SetFocus
        'Case DGPREEMPTINDEX
        Case DGOVSTARTINDEX
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            edcDGDropDown.SetFocus
        Case DGOVENDINDEX
            edcDGDropDown.Visible = True
            cmcDGDropDown.Visible = True
            edcDGDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDGSetShow                      *
'*                                                     *
'*             Created:8/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mDGSetShow(ilBoxNo As Integer)
'
'   mDGSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim ilLoop As Integer
    'If (ilBoxNo < LBound(tmDGCtrls)) Or (ilBoxNo > UBound(tmDGCtrls)) Then
    If (ilBoxNo < imLBCTCtrls) Or (ilBoxNo > UBound(tmDGCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case DGSTARTDATEINDEX 'Start Dates
            plcCalendar.Visible = False
            edcDGDropDown.Visible = False
            cmcDGDropDown.Visible = False
            slStr = edcDGDropDown.Text
            If gValidDate(slStr) Then   'Check that date is valid (press cancel with invalid date showing)
                smDGSave(10) = edcDGDropDown.Text
                slStr = gAddDayToDate(smDGSave(10))
            Else
                Beep
                smDGSave(10) = ""
                slStr = ""
            End If
            gSetShow pbcDGen, slStr, tmDGCtrls(DGSTARTDATEINDEX)
        Case DGNOWKSINDEX 'Number of weeks
            edcDGEdit.Visible = False
            smDGSave(1) = edcDGEdit.Text
            slStr = edcDGEdit.Text
            gSetShow pbcDGen, slStr, tmDGCtrls(ilBoxNo)
        Case DGENDDATEINDEX 'End dates
            plcCalendar.Visible = False
            edcDGDropDown.Visible = False
            cmcDGDropDown.Visible = False
            slStr = edcDGDropDown.Text
            If gValidDate(slStr) Then   'Check that date is valid (press cancel with invalid date showing)
                smDGSave(11) = edcDGDropDown.Text
                slStr = gAddDayToDate(smDGSave(11))
            Else
                Beep
                smDGSave(11) = ""
                slStr = ""
            End If
            gSetShow pbcDGen, slStr, tmDGCtrls(DGENDDATEINDEX)
        Case DGRUNINDEX 'Run weeks of...
            lbcDGRun.Visible = False
            edcDGDropDown.Visible = False
            cmcDGDropDown.Visible = False
            imDGSave(1) = lbcDGRun.ListIndex
            Select Case lbcDGRun.ListIndex
                Case 0
                    slStr = "Week"
                Case 1
                    slStr = "2nd"
                Case 2
                    slStr = "3rd"
                Case 3
                    slStr = "4th"
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcDGen, slStr, tmDGCtrls(ilBoxNo)
        Case DGDWPINDEX 'Daily; Weekly; % by Daily; % by Weekly
            lbcDWP.Visible = False
            edcDGDropDown.Visible = False
            cmcDGDropDown.Visible = False
            If imDGSave(2) <> lbcDWP.ListIndex Then
                If (imDGSave(2) = 0) Or (imDGSave(2) = 2) Then  'Change daily to weekly
                    smDGSave(2) = "0"
                    For ilLoop = 0 To 6 Step 1
                        smDGSave(2) = gAddStr(smDGSave(2), smDGSave(3 + ilLoop))
                    Next ilLoop
                    For ilLoop = 0 To 6 Step 1
                        If smDGSave(3 + ilLoop) <> "" Then
                            smDGSave(3 + ilLoop) = "Y"
                            slStr = "4"
                        Else
                            smDGSave(3 + ilLoop) = ""
                            slStr = "  "
                        End If
                        If tmDGCtrls(DGDAILYINDEX + ilLoop).sShow <> " R" Then
                            gSetShow pbcDGen, slStr, tmDGCtrls(DGDAILYINDEX + ilLoop)
                        End If
                    Next ilLoop
                Else    'Change weekly to daily
                    For ilLoop = 0 To 6 Step 1
                        smDGSave(3 + ilLoop) = ""
                        slStr = "  "
                        If tmDGCtrls(DGDAILYINDEX + ilLoop).sShow <> " R" Then
                            gSetShow pbcDGen, slStr, tmDGCtrls(DGDAILYINDEX + ilLoop)
                        End If
                    Next ilLoop
                    smDGSave(2) = ""
                End If
                gSetShow pbcDGen, smDGSave(2), tmDGCtrls(DGSPOTWKINDEX)
            End If
            imDGSave(2) = lbcDWP.ListIndex
            tmDGCtrls(DGSPOTWKINDEX).iReq = False
            Select Case lbcDWP.ListIndex
                Case 0
                    slStr = "Daily"
                Case 1
                    slStr = "Weekly"
                    tmDGCtrls(DGSPOTWKINDEX).iReq = True
                Case 2
                    slStr = "Daily by %"
                Case 3
                    slStr = "Weekly by %"
                    tmDGCtrls(DGSPOTWKINDEX).iReq = True
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcDGen, slStr, tmDGCtrls(ilBoxNo)
        Case DGSPOTWKINDEX   'Spots/week
            edcDGEdit.Visible = False  'Set visibility
            smDGSave(2) = edcDGEdit.Text
            slStr = edcDGEdit.Text
            gSetShow pbcDGen, slStr, tmDGCtrls(ilBoxNo)
        Case DGDAILYINDEX To DGDAILYINDEX + 6 'Daily
            If ((imDGSave(2) = 0) Or (imDGSave(2) = 2)) And (Not tgClfCntr(imLnRowNo - 1).iGame) Then
                edcDGEdit.Visible = False  'Set visibility
                smDGSave(3 + ilBoxNo - DGDAILYINDEX) = edcDGEdit.Text
                slStr = edcDGEdit.Text
                gSetShow pbcDGen, slStr, tmDGCtrls(ilBoxNo)
            Else
                pbcDGWkly.Visible = False  'Set visibility
                If ckcDGDay.Value = vbChecked Then
                    smDGSave(3 + ilBoxNo - DGDAILYINDEX) = "Y"
                    slStr = "4"
                Else
                    smDGSave(3 + ilBoxNo - DGDAILYINDEX) = "N"
                    slStr = "  "
                End If
                gSetShow pbcDGen, slStr, tmDGCtrls(ilBoxNo)
            End If
        Case DGPRICEINDEX
            edcDGDropDown.Visible = False
            cmcDGDropDown.Visible = False
            lbcPriceType(0).Visible = False
            If (tgUrf(0).sChgPrices <> "V") Then
                If lbcPriceType(0).ListIndex <= 0 Then
                    smDGSave(12) = edcDGDropDown.Text
                Else
                    smDGSave(12) = lbcPriceType(0).List(lbcPriceType(0).ListIndex)
                End If
            End If
            If smDGSave(12) <> "" Then
                If (Left$(smDGSave(12), 1) < "0") Or (Left$(smDGSave(12), 1) > "9") Then
                    gSetShow pbcDGen, smDGSave(12), tmDGCtrls(DGPRICEINDEX)
                Else
                    gFormatStr smDGSave(12), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    gSetShow pbcDGen, slStr, tmDGCtrls(DGPRICEINDEX)
                End If
            Else
                gSetShow pbcDGen, "", tmDGCtrls(DGPRICEINDEX)
            End If
        Case DGOVSTARTINDEX
            plcTme.Visible = False
            cmcDGDropDown.Visible = False
            edcDGDropDown.Visible = False  'Set visibility
            slStr = edcDGDropDown.Text
            If gValidTime(slStr) Then   'Check that time is valid (press cancel with invalid date showing)
                smDGSave(13) = edcDGDropDown.Text
            Else
                Beep
                smDGSave(13) = ""
                slStr = ""
            End If
            gSetShow pbcDGen, slStr, tmDGCtrls(ilBoxNo)
        Case DGOVENDINDEX
            plcTme.Visible = False
            cmcDGDropDown.Visible = False
            edcDGDropDown.Visible = False  'Set visibility
            slStr = edcDGDropDown.Text
            If gValidTime(slStr) Then   'Check that time is valid (press cancel with invalid date showing)
                smDGSave(14) = edcDGDropDown.Text
            Else
                Beep
                smDGSave(14) = ""
                slStr = ""
            End If
            gSetShow pbcDGen, slStr, tmDGCtrls(ilBoxNo)
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mAirEnableBox                   *
'*                                                     *
'*             Created:8/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mAirEnableBox(ilBoxNo As Integer)
'
'   mAirEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilBox As Integer
    Dim ilRowNo As Integer
    If ilBoxNo < imLBCTCtrls Or ilBoxNo > UBound(tmAirCtrls) Then
        Exit Sub
    End If
    If (imAirRowNo < vbcAirWk.Value) Or (imAirRowNo > vbcAirWk.Value + vbcAirWk.LargeChange) Then
        mAirSetShow ilBoxNo
        pbcWkArrow.Visible = False
        Exit Sub
    End If
    pbcWkArrow.Move pbcWkArrow.Left, pbcAirWk.Top - pbcAir.Top + tmAirCtrls(1).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15) - 15
    pbcWkArrow.Visible = True
    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        cmcAirCancel.SetFocus   'pbcClickFocus.SetFocus
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case AIRWEEKINDEX 'Vehicle
            imAiredWk = imAirSave(1, imAirRowNo)
            pbcAirWeek.Width = tmAirCtrls(AIRWEEKINDEX).fBoxW
            'gShowHelpMess tmClfHelp(), CLFFLTAIRWK
            gMoveTableCtrl pbcAirWk, pbcAirWeek, tmAirCtrls(AIRWEEKINDEX).fBoxX, tmAirCtrls(AIRWEEKINDEX).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15)
            pbcAirWeek_Paint
            pbcAirWeek.Visible = True  'Set visibility
            pbcAirWeek.SetFocus
        Case WEEKNOINDEX 'Week number
            'gShowHelpMess tmClfHelp(), CLFFLTWKNO
        Case WEEKSTARTDATEINDEX 'Week Start date
            'gShowHelpMess tmClfHelp(), CLFFLTWKST
        Case WEEKENDDATEINDEX   'Week end date
            'gShowHelpMess tmClfHelp(), CLFFLTWKEND
        Case DWINDEX
            lbcDW.Clear
            If imDGSave(2) <= 1 Then
                lbcDW.AddItem "Daily"
                lbcDW.AddItem "Weekly"
            Else
                lbcDW.AddItem "Daily by %"
                lbcDW.AddItem "Weekly by %"
            End If
            lbcDW.height = gListBoxHeight(lbcDW.ListCount, 2)
            edcAirDropDown.Width = tmAirCtrls(DWINDEX).fBoxW
            edcAirDropDown.MaxLength = 6
            gMoveTableCtrl pbcAirWk, edcAirDropDown, tmAirCtrls(DWINDEX).fBoxX, tmAirCtrls(DWINDEX).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15)
            cmcAirDropDown.Move edcAirDropDown.Left + edcAirDropDown.Width, edcAirDropDown.Top
            imChgMode = True
            If (imAirSave(2, imAirRowNo) >= 0) And (imAirSave(2, imAirRowNo) < lbcDW.ListCount) Then
                lbcDW.ListIndex = imAirSave(2, imAirRowNo)
                imComboBoxIndex = lbcDW.ListIndex
                edcAirDropDown.Text = lbcDW.List(imAirSave(2, imAirRowNo))
            Else
                If (imDGSave(2) >= 0) And (imDGSave(2) < lbcDW.ListCount) Then
                    lbcDW.ListIndex = imDGSave(2)
                    imComboBoxIndex = lbcDW.ListIndex
                    edcAirDropDown.Text = lbcDW.List(imComboBoxIndex)
                Else
                    If tgSpf.sAllowDailyBuys = "Y" Then
                        lbcDW.ListIndex = 0
                    Else
                        lbcDW.ListIndex = 1
                    End If
                    imComboBoxIndex = lbcDW.ListIndex
                    edcAirDropDown.Text = lbcDW.List(0)
                End If
            End If
            imChgMode = False
            lbcDW.Move edcAirDropDown.Left, edcAirDropDown.Top + edcAirDropDown.height
            edcAirDropDown.SelStart = 0
            edcAirDropDown.SelLength = Len(edcAirDropDown.Text)
            edcAirDropDown.Visible = True
            cmcAirDropDown.Visible = True
            edcAirDropDown.SetFocus
        Case SPOTPERWKINDEX 'Spots/week index
            If tgClfCntr(imLnRowNo - 1).iGame Then
                'gShowHelpMess tmClfHelp(), CLFFLTSPTWKGAME
                Exit Sub
            End If
            'gShowHelpMess tmClfHelp(), CLFFLTSPTWK
            gMoveTableCtrl pbcAirWk, edcAirSpots, tmAirCtrls(SPOTPERWKINDEX).fBoxX, tmAirCtrls(SPOTPERWKINDEX).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15)
            If Len(Trim$(smAirSave(3, imAirRowNo))) > 0 Then
                edcAirSpots.Text = Trim$(smAirSave(3, imAirRowNo))
            Else
                If imAirSave(2, imAirRowNo) = imDGSave(2) Then
                    'Obtain from top specification
                    edcAirSpots.Text = smDGSave(2)
                Else
                    edcAirSpots.Text = Trim$(smAirSave(3, imAirRowNo))
                End If
            End If
            edcAirSpots.Visible = True  'Set visibility
            edcAirSpots.SetFocus
        Case DAILYSPOTSINDEX To DAILYSPOTSINDEX + 6 'Day index
            If imAirRowNo = 1 Then
                imStartDayIndex = gWeekDayStr(smAirSave(1, imAirRowNo))
            Else
                If (smAirSave(13, imAirRowNo) = "Y") And (smAirSave(13, imAirRowNo + 1) <> "Y") Then
                    imStartDayIndex = gWeekDayStr(smAirSave(1, imAirRowNo))
                Else
                    imStartDayIndex = 0
                End If
            End If
            If imAirRowNo = UBound(smAirSave, 2) - 1 Then
                imEndDayIndex = gWeekDayStr(smAirSave(2, imAirRowNo))
            Else
                If (smAirSave(13, imAirRowNo) = "Y") And (smAirSave(13, imAirRowNo + 1) = "Y") Then
                    imEndDayIndex = gWeekDayStr(smAirSave(2, imAirRowNo))
                Else
                    imEndDayIndex = 6
                End If
            End If
            ilBox = ilBoxNo
            If smAirSave(13, imAirRowNo) = "Y" Then
                If smAirSave(13, imAirRowNo + 1) = "Y" Then
                    ilRowNo = imAirRowNo + 1
                Else
                    ilRowNo = imAirRowNo - 1
                End If
            End If
            imSettingDay = True
            If (imAirSave(2, imAirRowNo) = 0) And (Not tgClfCntr(imLnRowNo - 1).iGame) Then
                'gShowHelpMess tmClfHelp(), CLFFLTDAILY
                gMoveTableCtrl pbcAirWk, edcAirDaily, tmAirCtrls(ilBoxNo).fBoxX, tmAirCtrls(ilBoxNo).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15)
                edcAirDaily.Text = Trim$(smAirSave(4 + ilBoxNo - DAILYSPOTSINDEX, imAirRowNo))
                edcAirDaily.Visible = True  'Set visibility
                edcAirDaily.SetFocus
            Else
                If tgClfCntr(imLnRowNo - 1).iGame Then
                    'gShowHelpMess tmClfHelp(), CLFFLTGAMEDAYS
                Else
                    'gShowHelpMess tmClfHelp(), CLFFLTWEEKLY
                End If
                gMoveTableCtrl pbcAirWk, pbcAirDays, tmAirCtrls(ilBoxNo).fBoxX, tmAirCtrls(ilBoxNo).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15)
                If smAirSave(4 + ilBoxNo - DAILYSPOTSINDEX, imAirRowNo) = "Y" Then
                    ckcAirDay.Value = vbChecked
                Else
                    ckcAirDay.Value = vbUnchecked
                End If
                pbcAirDays.Visible = True  'Set visibility
                ckcAirDay.SetFocus
            End If
            imSettingDay = False
        Case WKPRICEINDEX
            lbcPriceType(1).height = gListBoxHeight(lbcPriceType(1).ListCount, 5)
            edcAirDropDown.Width = tmAirCtrls(WKPRICEINDEX).fBoxW
            edcAirDropDown.MaxLength = 12
            gMoveTableCtrl pbcAirWk, edcAirDropDown, tmAirCtrls(WKPRICEINDEX).fBoxX, tmAirCtrls(WKPRICEINDEX).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15)
            cmcAirDropDown.Move edcAirDropDown.Left + edcAirDropDown.Width, edcAirDropDown.Top
            imChgMode = True
            gFindMatch smAirSave(11, imAirRowNo), 1, lbcPriceType(1)
            If gLastFound(lbcPriceType(1)) > 0 Then
                lbcPriceType(1).ListIndex = gLastFound(lbcPriceType(1))
                edcAirDropDown.Text = lbcPriceType(1).List(lbcPriceType(1).ListIndex)
            Else
                lbcPriceType(1).ListIndex = 0
                edcAirDropDown.Text = smAirSave(11, imAirRowNo)
            End If
            imChgMode = False
            If imAirRowNo - vbcAirWk.Value > 5 Then
                lbcPriceType(1).Move cmcAirDropDown.Left + cmcAirDropDown.Width - lbcPriceType(1).Width, edcAirDropDown.Top - lbcPriceType(1).height
            Else
                lbcPriceType(1).Move cmcAirDropDown.Left + cmcAirDropDown.Width - lbcPriceType(1).Width, edcAirDropDown.Top + edcAirDropDown.height
            End If
            edcAirDropDown.SelStart = 0
            edcAirDropDown.SelLength = Len(edcAirDropDown.Text)
            edcAirDropDown.Visible = True
            cmcAirDropDown.Visible = True
            edcAirDropDown.SetFocus
    End Select
    'mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAirSetFocus                    *
'*                                                     *
'*             Created:8/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mAirSetFocus(ilBoxNo As Integer)
'
'   mAirSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBCTCtrls Or ilBoxNo > UBound(tmAirCtrls) Then
        Exit Sub
    End If
    If (imAirRowNo < vbcAirWk.Value) Or (imAirRowNo > vbcAirWk.Value + vbcAirWk.LargeChange) Then
        mAirSetShow ilBoxNo
        pbcWkArrow.Visible = False
        Exit Sub
    End If
    pbcWkArrow.Move pbcWkArrow.Left, pbcAirWk.Top - pbcAir.Top + tmAirCtrls(1).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15) - 15
    pbcWkArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case AIRWEEKINDEX 'Vehicle
            pbcAirWeek.Visible = True  'Set visibility
            pbcAirWeek.SetFocus
        Case WEEKNOINDEX 'Week number
        Case WEEKSTARTDATEINDEX 'Week Start date
        Case WEEKENDDATEINDEX   'Week end date
        Case DWINDEX
            edcAirDropDown.Visible = True
            cmcAirDropDown.Visible = True
            edcAirDropDown.SetFocus
        Case SPOTPERWKINDEX 'Spots/week index
            edcAirSpots.Visible = True  'Set visibility
            edcAirSpots.SetFocus
        Case DAILYSPOTSINDEX To DAILYSPOTSINDEX + 6 'Day index
            If (imAirSave(2, imAirRowNo) = 0) And (Not tgClfCntr(imLnRowNo - 1).iGame) Then
                edcAirDaily.Visible = True  'Set visibility
                edcAirDaily.SetFocus
            Else
                pbcAirDays.Visible = True  'Set visibility
                ckcAirDay.SetFocus
            End If
        Case WKPRICEINDEX
            edcAirDropDown.Visible = True
            cmcAirDropDown.Visible = True
            edcAirDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAirSetShow                     *
'*                                                     *
'*             Created:7/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mAirSetShow(ilBoxNo As Integer)
'
'   mAirSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilDay As Integer
    Dim llDate As Long
    Dim llAllowedChgWkStartDate As Long
    Dim llEndPrevPeriod As Long
    Dim ilRet As Integer
    Dim ilChg As Integer
    Dim ilBilledCount As Integer

    pbcWkArrow.Visible = False
    'If (ilBoxNo < LBound(smAirShow, 1)) Or (ilBoxNo > UBound(smAirShow, 1)) Then
    If (ilBoxNo < imLB1Or2) Or (ilBoxNo > UBound(smAirShow, 1)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case AIRWEEKINDEX 'Yes or No
            pbcAirWeek.Visible = False  'Set visibility
            If imAirSave(1, imAirRowNo) = 0 Then
                slStr = "Yes"
            Else
                slStr = "No"
                '6/11/12: Check if allowed
                llDate = gDateValue(smAirSave(1, imAirRowNo))
                ilRet = mGetMGCount(imLnRowNo, hmCHF, hmClf, hmSdf, hmSmf, llDate, llDate + 6, ilBilledCount)
                If ilRet > 0 Then
                    imAirSave(1, imAirRowNo) = 0
                    slStr = "Yes"
                    gSetShow pbcAirWk, slStr, tmAirCtrls(ilBoxNo)
                    MsgBox "Cannot set week as Not Aired because MG/Outside spots already aired from the week", vbOKOnly + vbExclamation, "Incomplete"
                    Exit Sub
                End If
            End If
            gSetShow pbcAirWk, slStr, tmAirCtrls(ilBoxNo)
            smAirShow(AIRWEEKINDEX, imAirRowNo) = tmAirCtrls(ilBoxNo).sShow
            If smAirSave(13, imAirRowNo + 1) = "Y" Then
                imAirSave(1, imAirRowNo + 1) = imAirSave(1, imAirRowNo)
                If imAirSave(1, imAirRowNo + 1) = 0 Then
                    slStr = "Yes"
                Else
                    slStr = "No"
                End If
                gSetShow pbcAirWk, slStr, tmAirCtrls(ilBoxNo)
                smAirShow(AIRWEEKINDEX, imAirRowNo + 1) = tmAirCtrls(ilBoxNo).sShow
            End If
            If imAiredWk <> imAirSave(1, imAirRowNo) And (imAiredWk = 1) Then
                'Set daily/weekly and spot counts
                If imMatchPattern = YES Then
                    If imDGSave(2) = 0 Then
                        'Daily
                        imAirSave(2, imAirRowNo) = 0
                        For ilDay = 0 To 6 Step 1
                            smAirSave(4 + ilDay, imAirRowNo) = smDGSave(3 + ilDay)
                        Next ilDay
                    ElseIf imDGSave(2) = 1 Then
                        imAirSave(2, imAirRowNo) = 1
                        smAirSave(3, imAirRowNo) = smDGSave(2)
                        For ilDay = 0 To 6 Step 1
                            smAirSave(4 + ilDay, imAirRowNo) = smDGSave(3 + ilDay)
                        Next ilDay
                    End If
                Else
                    'Model from adjacent week
                    'For ilLoop = imAirRowNo - 1 To LBound(smAirSave, 2) Step -1
                    For ilLoop = imAirRowNo - 1 To imLB1Or2 Step -1
                        If imAirSave(1, ilLoop) = 0 Then
                            imAirSave(2, imAirRowNo) = imAirSave(2, ilLoop)
                            smAirSave(3, imAirRowNo) = smAirSave(3, ilLoop)
                            For ilDay = 0 To 6 Step 1
                                smAirSave(4 + ilDay, imAirRowNo) = smAirSave(4 + ilDay, ilLoop)
                            Next ilDay
                            Exit For
                        End If
                    Next ilLoop
                End If
            End If
        Case WEEKNOINDEX 'Week number
        Case WEEKSTARTDATEINDEX 'Week Start date
        Case WEEKENDDATEINDEX   'Week end date
        Case DWINDEX
            lbcDW.Visible = False
            edcAirDropDown.Visible = False
            cmcAirDropDown.Visible = False
            If imAirSave(2, imAirRowNo) <> lbcDW.ListIndex Then
                tgClfCntr(imLnRowNo - 1).iResearch = False
                If imAirSave(2, imAirRowNo) = 0 Then  'Change daily to weekly
                    smAirSave(3, imAirRowNo) = "0"
                    For ilLoop = 0 To 6 Step 1
                        smAirSave(3, imAirRowNo) = gAddStr(smAirSave(3, imAirRowNo), smAirSave(4 + ilLoop, imAirRowNo))
                    Next ilLoop
                    For ilLoop = 0 To 6 Step 1
                        If smAirSave(4 + ilLoop, imAirRowNo) <> "" Then
                            smAirSave(4 + ilLoop, imAirRowNo) = "Y"
                            slStr = "4"
                        Else
                            smAirSave(4 + ilLoop, imAirRowNo) = ""
                            slStr = "  "
                        End If
                        gSetShow pbcAirWk, slStr, tmAirCtrls(DAILYSPOTSINDEX + ilLoop)
                        smAirShow(DAILYSPOTSINDEX + ilLoop, imAirRowNo) = tmAirCtrls(DAILYSPOTSINDEX + ilLoop).sShow
                    Next ilLoop
                Else    'Change weekly to daily
                    smAirSave(3, imAirRowNo) = ""
                    gSetShow pbcAirWk, slStr, tmAirCtrls(SPOTPERWKINDEX)
                    smAirShow(SPOTPERWKINDEX, imAirRowNo) = tmAirCtrls(SPOTPERWKINDEX).sShow
                    For ilLoop = 0 To 6 Step 1
                        smAirSave(4 + ilLoop, imAirRowNo) = ""
                        slStr = "  "
                        gSetShow pbcAirWk, slStr, tmAirCtrls(DAILYSPOTSINDEX + ilLoop)
                        smAirShow(DAILYSPOTSINDEX + ilLoop, imAirRowNo) = tmAirCtrls(DAILYSPOTSINDEX + ilLoop).sShow
                    Next ilLoop
                End If
                gSetShow pbcAirWk, smAirSave(3, imAirRowNo), tmAirCtrls(SPOTPERWKINDEX)
                smAirShow(SPOTPERWKINDEX, imAirRowNo) = tmAirCtrls(SPOTPERWKINDEX).sShow
            End If
            imAirSave(2, imAirRowNo) = lbcDW.ListIndex
            If imDGSave(2) <= 1 Then
                Select Case imAirSave(2, imAirRowNo)
                    Case 0
                        slStr = "Daily"
                    Case 1
                        slStr = "Weekly"
                    Case Else
                        slStr = ""
                End Select
            Else
                Select Case imAirSave(2, imAirRowNo)
                    Case 0
                        slStr = "Daily by %"
                    Case 1
                        slStr = "Weekly by %"
                    Case Else
                        slStr = ""
                End Select
            End If
            gSetShow pbcAirWk, slStr, tmAirCtrls(ilBoxNo)
            smAirShow(ilBoxNo, imAirRowNo) = tmAirCtrls(ilBoxNo).sShow
        Case SPOTPERWKINDEX   'Spots/week
            edcAirSpots.Visible = False  'Set visibility
            slStr = edcAirSpots.Text
            If smAirSave(3, imAirRowNo) <> edcAirSpots.Text Then
                llDate = gDateValue(smAirSave(1, imAirRowNo))
                If (llDate >= lmMonNowDate) And (llDate <= lmNowDate) And (tgUrf(0).sChgCntr <> "I") Then
                    llAllowedChgWkStartDate = mLnDate(imLnRowNo, smLnLLD, 0)
                    If llAllowedChgWkStartDate < tgClfCntr(imLnRowNo - 1).lEndBilledDate Then
                        llAllowedChgWkStartDate = tgClfCntr(imLnRowNo - 1).lEndBilledDate
                    End If
                    gUnpackDateLong tgSpf.iRPRP(0), tgSpf.iRPRP(1), llEndPrevPeriod
                    If (llAllowedChgWkStartDate < llEndPrevPeriod) Then
                        llAllowedChgWkStartDate = llEndPrevPeriod + 1
                    End If
                    If (llAllowedChgWkStartDate >= llDate) And (tgClfCntr(imLnRowNo - 1).iStatus > 0) And ((tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "F") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "I") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "A") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "N") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "P") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "M")) Then
                        mGetCurrentWkNoSpots imLnRowNo, hmCHF, hmClf
                        If tgClfCntr(imLnRowNo - 1).iMonThruToDaysSpots > 0 Then
                            If Val(slStr) < tgClfCntr(imLnRowNo - 1).iMonThruToDaysSpots Then
                                MsgBox "Cannot define spots less than " & Trim$(str$(tgClfCntr(imLnRowNo - 1).iMonThruToDaysSpots)) & " in current week as spots already aired", vbOKOnly + vbExclamation, "Incomplete"
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                '3/14/13:  Allow spots to be removed if billing by Order
                If (Val(slStr) = 0) And (Asc(tgSpf.sUsingFeatures10) And REPLACEDELWKWITHFILLS) = REPLACEDELWKWITHFILLS Then
                    ilChg = mReplaceWithFills(hmCHF, hmClf, hmSdf, hmSmf, llDate, lmMonNowDate, imPastMsgAsked, "W")
                    If Not ilChg Then
                        Exit Sub
                    End If
                Else
                    '6/9/12: Disallow number less then MG count
                    If (Val(smAirSave(3, imAirRowNo)) > Val(edcAirSpots.Text)) Then
                        ilRet = mGetMGCount(imLnRowNo, hmCHF, hmClf, hmSdf, hmSmf, llDate, llDate + 6, ilBilledCount)
                        If (Val(edcAirSpots.Text) < ilRet) Then
                            MsgBox "Cannot define spots less than " & Trim$(str$(ilRet)) & " in week as MG/Outside spots already aired", vbOKOnly + vbExclamation, "Incomplete"
                            Exit Sub
                        End If
                    End If
                    
                    If (llDate < lmMonNowDate) And (tgUrf(0).sChgCntr <> "I") Then
                        If (tgClfCntr(imLnRowNo - 1).iStatus = 0) Or (Not tgClfCntr(imLnRowNo - 1).iLineSchd) Or (imLnRowNo = UBound(tgClfCntr) + 1) Or (Val(smAirSave(3, imAirRowNo)) <> Val(edcAirSpots.Text)) Then
                            If imPastMsgAsked = 0 Then
                                ilRet = MsgBox("You are Changing a Past Week, Ok to Proceed", vbYesNo + vbDefaultButton2 + vbQuestion, "Warning")
                                imPastMsgAsked = ilRet
                                If ilRet = vbNo Then
                                    Exit Sub
                                End If
                            Else
                                If imPastMsgAsked = vbNo Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            MsgBox "Changing a Past Week not Permitted", vbOKOnly + vbExclamation, "Warning"
                            Exit Sub
                        End If
                    End If
                End If
                gSetShow pbcAirWk, slStr, tmAirCtrls(SPOTPERWKINDEX)
                smAirShow(SPOTPERWKINDEX, imAirRowNo) = tmAirCtrls(SPOTPERWKINDEX).sShow
                tgClfCntr(imLnRowNo - 1).iResearch = False
                imComputeReq = True
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                smAirSave(3, imAirRowNo) = edcAirSpots.Text
            End If
        Case DAILYSPOTSINDEX To DAILYSPOTSINDEX + 6 'Agency
            If (imAirSave(2, imAirRowNo) = 0) And (Not tgClfCntr(imLnRowNo - 1).iGame) Then
                edcAirDaily.Visible = False  'Set visibility
                slStr = edcAirDaily.Text
                If smAirSave(4 + ilBoxNo - DAILYSPOTSINDEX, imAirRowNo) <> edcAirDaily.Text Then
                    llDate = gDateValue(smAirSave(1, imAirRowNo)) + ilBoxNo - DAILYSPOTSINDEX
                    If (llDate >= lmMonNowDate) And (llDate <= lmNowDate) And (tgUrf(0).sChgCntr <> "I") Then
                        llAllowedChgWkStartDate = mLnDate(imLnRowNo, smLnLLD, 0)
                        If llAllowedChgWkStartDate < tgClfCntr(imLnRowNo - 1).lEndBilledDate Then
                            llAllowedChgWkStartDate = tgClfCntr(imLnRowNo - 1).lEndBilledDate
                        End If
                        gUnpackDateLong tgSpf.iRPRP(0), tgSpf.iRPRP(1), llEndPrevPeriod
                        If (llAllowedChgWkStartDate < llEndPrevPeriod) Then
                            llAllowedChgWkStartDate = llEndPrevPeriod + 1
                        End If
                        If (llAllowedChgWkStartDate >= llDate) And (tgClfCntr(imLnRowNo - 1).iStatus > 0) And ((tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "F") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "I") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "A") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "N") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "M")) Then
                            MsgBox "Changing a Past Day not Permitted", vbOKOnly + vbExclamation, "Warning"
                            Exit Sub
                        End If
                    End If
                    
                    '3/14/13:  Allow spots to be removed if billing by Order
                    If (Val(slStr) = 0) And (Asc(tgSpf.sUsingFeatures10) And REPLACEDELWKWITHFILLS) = REPLACEDELWKWITHFILLS Then
                        ilChg = mReplaceWithFills(hmCHF, hmClf, hmSdf, hmSmf, llDate, lmMonNowDate, imPastMsgAsked, "D")
                        If Not ilChg Then
                            Exit Sub
                        End If
                    Else
                        '6/9/12: Disallow number less then MG count
                        If (Val(smAirSave(4 + ilBoxNo - DAILYSPOTSINDEX, imAirRowNo)) > Val(edcAirDaily.Text)) Then
                            ilRet = mGetMGCount(imLnRowNo, hmCHF, hmClf, hmSdf, hmSmf, llDate, llDate, ilBilledCount)
                            If (Val(edcAirDaily.Text) < ilRet) Then
                                MsgBox "Cannot define spots less than " & Trim$(str$(ilRet)) & " in day as MG/Outside spots already aired", vbOKOnly + vbExclamation, "Incomplete"
                                Exit Sub
                            End If
                        End If
                        
                        If (llDate < lmMonNowDate) And (tgUrf(0).sChgCntr <> "I") Then
                            If (tgClfCntr(imLnRowNo - 1).iStatus = 0) Or (Not tgClfCntr(imLnRowNo - 1).iLineSchd) Or (imLnRowNo = UBound(tgClfCntr) + 1) Or (Val(smAirSave(4 + ilBoxNo - DAILYSPOTSINDEX, imAirRowNo)) > Val(edcAirDaily.Text)) Then
                                If imPastMsgAsked = 0 Then
                                    ilRet = MsgBox("You are Changing a Past Week, Ok to Proceed", vbYesNo + vbDefaultButton2 + vbQuestion, "Warning")
                                    imPastMsgAsked = ilRet
                                    If ilRet = vbNo Then
                                        Exit Sub
                                    End If
                                Else
                                    If imPastMsgAsked = vbNo Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                MsgBox "Changing a Past Week not Permitted", vbOKOnly + vbExclamation, "Warning"
                                Exit Sub
                            End If
                        End If
                    End If
                    gSetShow pbcAirWk, slStr, tmAirCtrls(ilBoxNo)
                    smAirShow(ilBoxNo, imAirRowNo) = tmAirCtrls(ilBoxNo).sShow
                    tgClfCntr(imLnRowNo - 1).iResearch = False
                    imComputeReq = True
                    If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                        imLnChg = True
                    End If
                    smAirSave(4 + ilBoxNo - DAILYSPOTSINDEX, imAirRowNo) = edcAirDaily.Text
                End If
            Else
                pbcAirDays.Visible = False  'Set visibility
                If ckcAirDay.Value = vbChecked Then
                    slStr = "4"
                    If smAirSave(4 + ilBoxNo - DAILYSPOTSINDEX, imAirRowNo) <> "Y" Then
                        tgClfCntr(imLnRowNo - 1).iResearch = False
                        imComputeReq = True
                        If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                            imLnChg = True
                        End If
                        smAirSave(4 + ilBoxNo - DAILYSPOTSINDEX, imAirRowNo) = "Y"
                    End If
                Else
                    slStr = "  "
                    If smAirSave(4 + ilBoxNo - DAILYSPOTSINDEX, imAirRowNo) <> "N" Then
                        tgClfCntr(imLnRowNo - 1).iResearch = False
                        imComputeReq = True
                        If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                            imLnChg = True
                        End If
                        smAirSave(4 + ilBoxNo - DAILYSPOTSINDEX, imAirRowNo) = "N"
                    End If
                End If
                gSetShow pbcAirWk, slStr, tmAirCtrls(ilBoxNo)
                smAirShow(ilBoxNo, imAirRowNo) = tmAirCtrls(ilBoxNo).sShow
            End If
        Case WKPRICEINDEX
            lbcPriceType(1).Visible = False
            edcAirDropDown.Visible = False
            cmcAirDropDown.Visible = False
            If (tgUrf(0).sChgPrices = "V") Then
                edcAirDropDown.Text = smAirSave(11, imAirRowNo)
            End If
            slStr = edcAirDropDown.Text
            If slStr <> "" Then
                If (Left$(slStr, 1) < "0") Or (Left$(slStr, 1) > "9") Then
                    gSetShow pbcAirWk, slStr, tmAirCtrls(WKPRICEINDEX)
                    smAirShow(WKPRICEINDEX, imAirRowNo) = tmAirCtrls(WKPRICEINDEX).sShow
                Else
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    gSetShow pbcAirWk, slStr, tmAirCtrls(WKPRICEINDEX)
                    smAirShow(WKPRICEINDEX, imAirRowNo) = tmAirCtrls(WKPRICEINDEX).sShow
                End If
            Else
                gSetShow pbcAirWk, "", tmAirCtrls(WKPRICEINDEX)
                smAirShow(WKPRICEINDEX, imAirRowNo) = tmAirCtrls(WKPRICEINDEX).sShow
            End If
            If smAirSave(11, imAirRowNo) <> edcAirDropDown.Text Then
                tgClfCntr(imLnRowNo - 1).iResearch = False
                imComputeReq = True
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                smAirSave(11, imAirRowNo) = edcAirDropDown.Text
            End If
    End Select
    Screen.MousePointer = vbHourglass
    mSetAirPkPrice
    mFlightTotals
    mGetDemoFlightTotals
    Screen.MousePointer = vbDefault
End Sub

Private Sub pbcAirDemo_Paint()
    Dim ilBox As Integer

    For ilBox = imLBCTCtrls To UBound(tmDFTCtrls) Step 1
        pbcAirDemo.CurrentX = tmDFTCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcAirDemo.CurrentY = tmDFTCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
        pbcAirDemo.Print tmDFTCtrls(ilBox).sShow
    Next ilBox

End Sub

Private Sub pbcAirSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilDay As Integer
    Dim slStr As String
    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        If imAirBoxNo <> -1 Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If GetFocus() <> pbcAirSTab.HWnd Then
        Exit Sub
    End If
    If imGenReq Then
        If (cmcGenAll.Visible) And (cmcGenAll.Enabled) Then
            cmcGenAll.SetFocus
        ElseIf (cmcGenNew.Visible) And (cmcGenNew.Enabled) Then
            cmcGenNew.SetFocus
        Else
            cmcAirCancel.SetFocus
        End If
        Exit Sub
    End If
    mDGSetShow imDGBoxNo
    imDGBoxNo = -1
    imTabDirection = -1  'Set-right to left
    Select Case imAirBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            imAirRowNo = 0
            ilBox = 1
            vbcAirWk.Value = vbcAirWk.Min
            'Don't scroll to current week as it look bad
            Do
                If imAirRowNo = UBound(smAirSave, 2) - 1 Then
                    imAirBoxNo = -1
                    cmcAirDone.SetFocus
                    Exit Sub
                End If
                imAirRowNo = imAirRowNo + 1
                imSettingValue = True
                If imAirRowNo > vbcAirWk.Value + vbcAirWk.LargeChange Then
                    imSettingValue = True
                    vbcAirWk.Value = vbcAirWk.Value + 1
                End If
                imSettingValue = False
            Loop While (lmAllowedChgWkStartDate > gDateValue(smAirSave(2, imAirRowNo))) And (tgClfCntr(imLnRowNo - 1).iStatus > 0) And ((tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "F") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "I") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "M"))
        Case SPOTPERWKINDEX
            mAirSetShow imAirBoxNo
            If imAirRowNo > 1 Then
                If (smAirSave(13, imAirRowNo - 1) = "Y") And (smAirSave(13, imAirRowNo) = "Y") Then
                    If imAirRowNo <= 1 Then
                        imAirBoxNo = -1
                        cmcAirDone.SetFocus
                        Exit Sub
                    Else
                        ilDay = gWeekDayStr(smAirSave(2, imAirRowNo - 1))
                        ilBox = DAILYSPOTSINDEX + ilDay
                        imAirRowNo = imAirRowNo - 1
                        If imAirRowNo < vbcAirWk.Value Then
                            imSettingValue = True
                            vbcAirWk.Value = vbcAirWk.Value - 1
                            imSettingValue = False
                        End If
                        imAirBoxNo = ilBox
                        mAirEnableBox ilBox
                        Exit Sub
                    End If
                End If
            End If
            If (Trim$(smAirSave(3, imAirRowNo)) = "") Or (Val(smAirSave(3, imAirRowNo)) = 0) Then
                Beep
                ilBox = SPOTPERWKINDEX
            Else
                'Disallow Daily/Weekly from being changed (mixture not allowed 11/29/02)
                'ilBox = imAirBoxNo - 1
                ilBox = AIRWEEKINDEX
            End If
            imAirBoxNo = ilBox
            mAirEnableBox ilBox
            Exit Sub
        Case DAILYSPOTSINDEX
            mAirSetShow imAirBoxNo
            mSetAirDateFromSpots
            If tgClfCntr(imLnRowNo - 1).iGame Then
                ilBox = AIRWEEKINDEX
            Else
                If imAirSave(2, imAirRowNo) = 0 Then
                    'Disallow Daily/Weekly from being changed (mixture not allowed 11/29/02)
                    'ilBox = DWINDEX
                    ilBox = AIRWEEKINDEX
                Else
                    ilBox = SPOTPERWKINDEX
                End If
            End If
            imAirBoxNo = ilBox
            mAirEnableBox ilBox
            Exit Sub
        Case AIRWEEKINDEX 'Name (first control within header)
            mAirSetShow imAirBoxNo
            Do
                If imAirRowNo <= 1 Then
                    imAirBoxNo = -1
                    cmcAirDone.SetFocus
                    Exit Sub
                End If
                imAirRowNo = imAirRowNo - 1
                If imAirRowNo < vbcAirWk.Value Then
                    imSettingValue = True
                    vbcAirWk.Value = vbcAirWk.Value - 1
                    imSettingValue = False
                End If
            Loop While (lmAllowedChgWkStartDate > gDateValue(smAirSave(2, imAirRowNo))) And (tgClfCntr(imLnRowNo - 1).iStatus > 0) And ((tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "F") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "I") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "M"))
            If imAirSave(1, imAirRowNo) = 1 Then
                ilBox = AIRWEEKINDEX
            Else
                If ((smLineType = "O") Or (smLineType = "A")) And (tgSpf.sCAudPkg <> "Y") Then
                    ilBox = DAILYSPOTSINDEX + 6
                Else
                    If smLineType = "H" Then
                        slStr = mGetPkLineType(imLnRowNo)
                        If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                            ilBox = DAILYSPOTSINDEX + 6
                        Else
                            ilBox = WKPRICEINDEX
                        End If
                    Else
                        ilBox = WKPRICEINDEX
                    End If
                    If (ilBox = WKPRICEINDEX) And ((tgClfCntr(imLnRowNo - 1).iAcqCostSpotRate = 2) Or (igFreezeState = 1)) Then
                        ilBox = DAILYSPOTSINDEX + 6
                    End If
                End If
            End If
            imAirBoxNo = ilBox
            mAirEnableBox ilBox
            Exit Sub
        Case Else
            ilBox = imAirBoxNo - 1
    End Select
    mAirSetShow imAirBoxNo
    imAirBoxNo = ilBox
    mAirEnableBox ilBox
End Sub
Private Sub pbcAirTab_GotFocus()
    Dim ilBox As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilFound As Integer

    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If GetFocus() <> pbcAirTab.HWnd Then
        Exit Sub
    End If
    mDGSetShow imDGBoxNo
    imDGBoxNo = -1
    imTabDirection = 0  'Set-Left to right
    ilBox = imAirBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imTabDirection = -1  'Set-Right to left
                imAirRowNo = UBound(smAirSave, 2) - 1
                imSettingValue = True
                If imAirRowNo <= vbcAirWk.LargeChange + 1 Then
                    vbcAirWk.Value = 1
                Else
                    vbcAirWk.Value = imAirRowNo - vbcAirWk.LargeChange '- 1 'was 2
                End If
                imSettingValue = False
                If ((smLineType = "O") Or (smLineType = "A")) And (tgSpf.sCAudPkg <> "Y") Then
                    ilBox = DAILYSPOTSINDEX + 6
                Else
                    If smLineType = "H" Then
                        slStr = mGetPkLineType(imLnRowNo)
                        If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                            ilBox = DAILYSPOTSINDEX + 6
                        Else
                            ilBox = WKPRICEINDEX
                        End If
                    Else
                        ilBox = WKPRICEINDEX
                    End If
                End If
                'If last week in past then all week must be in past
                If (lmAllowedChgWkStartDate > gDateValue(smAirSave(2, imAirRowNo))) And (tgClfCntr(imLnRowNo - 1).iStatus > 0) And ((tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "F") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "I") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "M")) Then
                    imAirBoxNo = -1
                    cmcAirDone.SetFocus
                    Exit Sub
                End If
                If (ilBox = WKPRICEINDEX) And ((tgClfCntr(imLnRowNo - 1).iAcqCostSpotRate = 2) Or (igFreezeState = 1)) Then
                    imAirBoxNo = -1
                    cmcAirDone.SetFocus
                    Exit Sub
                End If
            Case AIRWEEKINDEX
                If imAirSave(1, imAirRowNo) = 1 Then
                    mAirSetShow imAirBoxNo
                    Do
                        If imAirRowNo = UBound(smAirSave, 2) - 1 Then
                            imAirBoxNo = -1
                            cmcAirDone.SetFocus
                            Exit Sub
                        End If
                        imAirRowNo = imAirRowNo + 1
                        If imAirRowNo > vbcAirWk.Value + vbcAirWk.LargeChange Then
                            imSettingValue = True
                            vbcAirWk.Value = vbcAirWk.Value + 1
                            imSettingValue = False
                        End If
                    Loop While (lmAllowedChgWkStartDate > gDateValue(smAirSave(2, imAirRowNo))) And (tgClfCntr(imLnRowNo - 1).iStatus > 0) And ((tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "F") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "I") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "M"))
                    ilBox = AIRWEEKINDEX
                    If imAirRowNo > 1 Then  'Test for split week
                        If (ilBox = AIRWEEKINDEX) And (smAirSave(13, imAirRowNo - 1) = "Y") And (smAirSave(13, imAirRowNo) = "Y") Then
                            ilBox = SPOTPERWKINDEX 'DWINDEX
                        End If
                    End If
                    imAirBoxNo = ilBox
                    mAirEnableBox ilBox
                    Exit Sub
                Else
                    If tgClfCntr(imLnRowNo - 1).iGame Then
                        ilBox = DAILYSPOTSINDEX
                    Else
                        'Disallow Daily/Weekly from being changed (mixture not allowed 11/29/02)
                        If imAirSave(2, imAirRowNo) < 0 Then
                            If (imDGSave(2) >= 0) And (imDGSave(2) < lbcDW.ListCount) Then
                                lbcDW.ListIndex = imDGSave(2)
                                imComboBoxIndex = lbcDW.ListIndex
                                edcAirDropDown.Text = lbcDW.List(imComboBoxIndex)
                            Else
                                If (tgSpf.sAllowDailyBuys = "Y") Then
                                    lbcDW.ListIndex = 0
                                    imComboBoxIndex = lbcDW.ListIndex
                                    edcAirDropDown.Text = lbcDW.List(0)
                                    imAirSave(2, imAirRowNo) = 0
                                Else
                                    lbcDW.ListIndex = 1
                                    imComboBoxIndex = lbcDW.ListIndex
                                    edcAirDropDown.Text = lbcDW.List(1)
                                    imAirSave(2, imAirRowNo) = 1
                                End If
                            End If
                        Else
                            lbcDW.ListIndex = imAirSave(2, imAirRowNo)
                            imComboBoxIndex = lbcDW.ListIndex
                            edcAirDropDown.Text = lbcDW.List(imAirSave(2, imAirRowNo))
                        End If
                        ilBox = DWINDEX
                        mAirSetShow ilBox
                        ilFound = False
                    End If
                End If
            Case DWINDEX
                If lbcDW.ListIndex = 0 Then
                    ilBox = DAILYSPOTSINDEX
                Else
                    ilBox = SPOTPERWKINDEX
                End If
            Case SPOTPERWKINDEX
                mAirSetShow imAirBoxNo
                If (Trim$(smAirSave(3, imAirRowNo)) = "") Or (Val(smAirSave(3, imAirRowNo)) = 0) Then
                    Beep
                    ilBox = SPOTPERWKINDEX
                Else
                    ilBox = imAirBoxNo + 1
                End If
                imAirBoxNo = ilBox
                mAirEnableBox ilBox
                Exit Sub
            Case WKPRICEINDEX, DAILYSPOTSINDEX + 6 'DAILYSPOTSINDEX + 6'Last control within header
                If imAirBoxNo = DAILYSPOTSINDEX + 6 Then
                    If ((smLineType = "O") Or (smLineType = "A")) And (tgSpf.sCAudPkg <> "Y") Then
                        ilRet = True
                    Else
                        If smLineType = "H" Then
                            slStr = mGetPkLineType(imLnRowNo)
                            If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                                ilRet = True
                            Else
                                ilRet = False
                            End If
                        Else
                            ilRet = False
                        End If
                    End If
                Else
                    ilRet = True
                End If
                If (tgClfCntr(imLnRowNo - 1).iAcqCostSpotRate = 2) Or (igFreezeState = 1) Then
                    ilRet = True
                End If
                If (imAirBoxNo = WKPRICEINDEX) Or ilRet Then
                    mAirSetShow imAirBoxNo
                    mSetAirDateFromSpots
                    Do
                        If imAirRowNo = UBound(smAirSave, 2) - 1 Then
                            imAirBoxNo = -1
                            cmcAirDone.SetFocus
                            Exit Sub
                        End If
                        imAirRowNo = imAirRowNo + 1
                        If imAirRowNo > vbcAirWk.Value + vbcAirWk.LargeChange Then
                            imSettingValue = True
                            vbcAirWk.Value = vbcAirWk.Value + 1
                            imSettingValue = False
                        End If
                    Loop While (lmAllowedChgWkStartDate > gDateValue(smAirSave(2, imAirRowNo))) And (tgClfCntr(imLnRowNo - 1).iStatus > 0) And ((tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "F") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "I") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "M"))
                    ilBox = AIRWEEKINDEX
                    If imAirRowNo > 1 Then
                        If (ilBox = AIRWEEKINDEX) And (smAirSave(13, imAirRowNo - 1) = "Y") And (smAirSave(13, imAirRowNo) = "Y") Then
                            ilFound = False
                            ilBox = DWINDEX
                        End If
                    End If
                    imAirBoxNo = ilBox
                    mAirEnableBox ilBox
                    Exit Sub
                Else
                    ilBox = imAirBoxNo + 1
                End If
            Case Else
                ilBox = imAirBoxNo + 1
        End Select
    Loop While Not ilFound
    mAirSetShow imAirBoxNo
    imAirBoxNo = ilBox
    mAirEnableBox ilBox
End Sub
Private Sub pbcAirWeek_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcAirWeek_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        If imAirSave(1, imAirRowNo) <> 0 Then
            If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                imLnChg = True
            End If
        End If
        imAirSave(1, imAirRowNo) = 0
        pbcAirWeek_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imAirSave(1, imAirRowNo) <> 1 Then
            If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                imLnChg = True
            End If
        End If
        imAirSave(1, imAirRowNo) = 1
        pbcAirWeek_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imAirSave(1, imAirRowNo) = 0 Then
            imAirSave(1, imAirRowNo) = 1
        Else
            imAirSave(1, imAirRowNo) = 0
        End If
        If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
            imLnChg = True
        End If
        pbcAirWeek_Paint
    End If
End Sub
Private Sub pbcAirWeek_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imAirSave(1, imAirRowNo) = 0 Then
        If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
            imLnChg = True
        End If
        imAirSave(1, imAirRowNo) = 1
    Else
        If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
            imLnChg = True
        End If
        imAirSave(1, imAirRowNo) = 0
    End If
    pbcAirWeek_Paint
End Sub
Private Sub pbcAirWeek_Paint()
    pbcAirWeek.Cls
    pbcAirWeek.CurrentX = fgBoxInsetX
    pbcAirWeek.CurrentY = 30 'fgBoxInsetY
    If imAirSave(1, imAirRowNo) = 0 Then
        pbcAirWeek.Print "Y"
    Else
        pbcAirWeek.Print "N"
    End If
End Sub
Private Sub pbcAirWk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilAirRowNo As Integer
    Dim slStr As String
    Dim llAllowedChgWkStartDate As Long
    Dim llEndPrevPeriod As Long
    Dim llAirDate As Long
    Dim ilSvWkDailyFlag As Integer
    Dim ilLoop As Integer

    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        cmcAirCancel.SetFocus
        Exit Sub
    End If
    ilCompRow = vbcAirWk.LargeChange + 1
    If UBound(smAirShow, 2) - 1 > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smAirShow, 2) - 1
    End If
    'llAllowedChgWkStartDate = lmAllowedChgWkStartDate
    ilSvWkDailyFlag = imLnSave(8, imLnRowNo)
    imLnSave(8, imLnRowNo) = 0
    llAllowedChgWkStartDate = mLnDate(imLnRowNo, smLnLLD, 2)
    imLnSave(8, imLnRowNo) = ilSvWkDailyFlag
    If llAllowedChgWkStartDate < tgClfCntr(imLnRowNo - 1).lEndBilledDate Then
        llAllowedChgWkStartDate = tgClfCntr(imLnRowNo - 1).lEndBilledDate
    End If
    gUnpackDateLong tgSpf.iRPRP(0), tgSpf.iRPRP(1), llEndPrevPeriod
    If (llAllowedChgWkStartDate < llEndPrevPeriod) Then
        llAllowedChgWkStartDate = llEndPrevPeriod + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCTCtrls To UBound(tmAirCtrls) Step 1
            If (X >= tmAirCtrls(ilBox).fBoxX) And (X <= (tmAirCtrls(ilBox).fBoxX + tmAirCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmAirCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmAirCtrls(ilBox).fBoxY + tmAirCtrls(ilBox).fBoxH)) Then
                    mAirSetShow imAirBoxNo
                    ilAirRowNo = ilRow + vbcAirWk.Value - 1
                    If (ilBox >= DAILYSPOTSINDEX) And (ilBox <= DAILYSPOTSINDEX + 6) Then
                        llAirDate = gDateValue(gObtainPrevMonday(smAirSave(1, ilAirRowNo))) + ilBox - DAILYSPOTSINDEX
                    Else
                        llAirDate = gDateValue(gObtainPrevMonday(smAirSave(1, ilAirRowNo)))
                    End If
                    If (llAllowedChgWkStartDate > llAirDate) And (tgClfCntr(imLnRowNo - 1).iStatus > 0) And ((tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "F") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "I") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "A") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "N") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "M") Or (tgClfCntr(imLnRowNo - 1).iLineSchd)) Then
                        Beep
                        MsgBox "Can't alter, Start Date in Past", vbOKOnly + vbExclamation, "Warning"
                        'If imDGBoxNo > 0 Then
                        '    mDGSetFocus imDGBoxNo
                        'ElseIf imAirBoxNo > 0 Then
                        '    mAirEnableBox imAirBoxNo
                        'Else
                            cmcAirDone.SetFocus
                        'End If
                        Exit Sub
                    End If
                    'If (tgClfCntr(imLnRowNo - 1).iGame) And ((ilBox = SPOTPERWKINDEX) Or (ilBox = DWINDEX) Or ((ilBox >= WEEKNOINDEX) And (ilBox <= WEEKENDDATEINDEX))) Then
                    If mIsLnGame(imLnRowNo) And (ilBox <> WKPRICEINDEX) Then
                        Beep
                        'If imDGBoxNo > 0 Then
                        '    mDGSetFocus imDGBoxNo
                        'ElseIf imAirBoxNo > 0 Then
                        '    mAirEnableBox imAirBoxNo
                        'Else
                            cmcAirDone.SetFocus
                        'End If
                        Exit Sub
                    End If
                    If (imAirSave(2, ilAirRowNo) = 0) And (ilBox = SPOTPERWKINDEX) Then
                        Beep
                        MsgBox "Can't defined Spots/Week as Line Defined as Daily", vbOKOnly + vbExclamation, "Warning"
                        'If imDGBoxNo > 0 Then
                        '    mDGSetFocus imDGBoxNo
                        'ElseIf imAirBoxNo > 0 Then
                        '    mAirEnableBox imAirBoxNo
                        'Else
                            cmcAirDone.SetFocus
                        'End If
                        Exit Sub
                    End If
                    If (ilBox >= WEEKNOINDEX) And (ilBox <= WEEKENDDATEINDEX) Then
                        Beep
                        MsgBox "Can't change week Dates", vbOKOnly + vbExclamation, "Warning"
                        'If imDGBoxNo > 0 Then
                        '    mDGSetFocus imDGBoxNo
                        'ElseIf imAirBoxNo > 0 Then
                        '    mAirEnableBox imAirBoxNo
                        'Else
                            cmcAirDone.SetFocus
                        'End If
                        Exit Sub
                    End If
                    If (ilBox > AIRWEEKINDEX) And (imAirSave(1, ilAirRowNo) = 1) Then
                        Beep
                        MsgBox "Must be made an Airing Week First", vbOKOnly + vbExclamation, "Warning"
                        'If imDGBoxNo > 0 Then
                        '    mDGSetFocus imDGBoxNo
                        'ElseIf imAirBoxNo > 0 Then
                        '    mAirEnableBox imAirBoxNo
                        'Else
                            cmcAirDone.SetFocus
                        'End If
                        Exit Sub
                    End If
                    'Disallow Daily/Weekly from being changed (mixture not allowed 11/29/02)
                    If (ilBox = DWINDEX) And (imAirSave(2, ilAirRowNo) >= 0) Then
                        Beep
                        MsgBox "Can't change Daily/Weekly indicator", vbOKOnly + vbExclamation, "Warning"
                        'If imDGBoxNo > 0 Then
                        '    mDGSetFocus imDGBoxNo
                        'ElseIf imAirBoxNo > 0 Then
                        '    mAirEnableBox imAirBoxNo
                        'Else
                            cmcAirDone.SetFocus
                        'End If
                        Exit Sub
                    End If
                    If ilAirRowNo > 1 Then
                        If (ilBox = AIRWEEKINDEX) And (smAirSave(13, ilAirRowNo - 1) = "Y") And (smAirSave(13, ilAirRowNo) = "Y") Then
                            Beep
                            If imDGBoxNo > 0 Then
                                mDGSetFocus imDGBoxNo
                            Else
                                mAirEnableBox imAirBoxNo
                            End If
                            Exit Sub
                        End If
                        If (ilBox = DWINDEX) And (smAirSave(13, ilAirRowNo - 1) = "Y") And (smAirSave(13, ilAirRowNo) = "Y") Then
                            Beep
                            If imDGBoxNo > 0 Then
                                mDGSetFocus imDGBoxNo
                            ElseIf imAirBoxNo > 0 Then
                                mAirEnableBox imAirBoxNo
                            Else
                                cmcAirDone.SetFocus
                            End If
                            Exit Sub
                        End If
                    End If
                    If (ilBox = WKPRICEINDEX) And ((smLineType = "O") Or (smLineType = "A")) And (tgSpf.sCAudPkg <> "Y") Then
                        Beep
                        MsgBox "Spot Prices on Package Lines are Calculated Automatically", vbOKOnly + vbExclamation, "Warning"
                        'If imDGBoxNo > 0 Then
                        '    mDGSetFocus imDGBoxNo
                        'ElseIf imAirBoxNo > 0 Then
                        '    mAirEnableBox imAirBoxNo
                        'Else
                            cmcAirDone.SetFocus
                        'End If
                        Exit Sub
                    End If
                    'If (ilBox = WKPRICEINDEX) And (tgUrf(0).sChgPrices = "V") Then
                    If (ilBox = WKPRICEINDEX) And (Not Contract.mAllowedToChgSpotPrice(imLnRowNo)) Then
                        Beep
                        If (tgClfCntr(imLnRowNo - 1).iInPast) And (tgClfCntr(imLnRowNo - 1).lEndBilledDate > 0) And (tgClfCntr(imLnRowNo - 1).iAcqCostSpotRate = 2) Then
                            MsgBox "Spot Price change disallowed on flight screen since Spot Price and Acquisition Cost must match and weeks have been billed", vbOKOnly + vbExclamation, "Warning"
                        Else
                            MsgBox "User option prohibits changing of Spot Price", vbOKOnly + vbExclamation, "Warning"
                        End If
                        'If imDGBoxNo > 0 Then
                        '    mDGSetFocus imDGBoxNo
                        'ElseIf imAirBoxNo > 0 Then
                        '    mAirEnableBox imAirBoxNo
                        'Else
                            cmcAirDone.SetFocus
                        'End If
                        Exit Sub
                    End If
                    If (ilBox = WKPRICEINDEX) And (smLineType = "H") Then
                        slStr = mGetPkLineType(imLnRowNo)
                        If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                            Beep
                            MsgBox "Spot Prices on Package Lines are Calculated Automatically", vbOKOnly + vbExclamation, "Warning"
                            'If imDGBoxNo > 0 Then
                            '    mDGSetFocus imDGBoxNo
                            'ElseIf imAirBoxNo > 0 Then
                            '    mAirEnableBox imAirBoxNo
                            'Else
                                cmcAirDone.SetFocus
                            'End If
                            Exit Sub
                        End If
                    End If
                    If (ilBox = WKPRICEINDEX) And ((tgClfCntr(imLnRowNo - 1).iAcqCostSpotRate = 2) Or (igFreezeState = 1)) Then
                        Beep
                        If tgClfCntr(imLnRowNo - 1).iAcqCostSpotRate = 2 Then
                            MsgBox "Spot Price change disallowed on flight screen since Spot Price and Acquisition Cost must match", vbOKOnly + vbExclamation, "Warning"
                        Else
                            MsgBox "Spot Price change disallowed on flight screen when Freeze Calculation selected", vbOKOnly + vbExclamation, "Warning"
                        End If
                        'If imDGBoxNo > 0 Then
                        '    mDGSetFocus imDGBoxNo
                        'ElseIf imAirBoxNo > 0 Then
                        '    mAirEnableBox imAirBoxNo
                        'Else
                            cmcAirDone.SetFocus
                        'End If
                        Exit Sub
                    End If
                    If (smLineType = "H") Then
                        For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                            If tgPackLn(ilLoop).iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine Then
                                'If tgPackLn(ilLoop).iAlter <> True Then
                                If tgPackLn(ilLoop).sAlter <> "Y" Then
                                    Beep
                                    MsgBox "Can't Alter Hidden Line", vbOKOnly + vbExclamation, "Warning"
                                    'If imDGBoxNo > 0 Then
                                    '    mDGSetFocus imDGBoxNo
                                    'ElseIf imAirBoxNo > 0 Then
                                    '    mAirEnableBox imAirBoxNo
                                    'Else
                                        cmcAirDone.SetFocus
                                    'End If
                                    Exit Sub
                                End If
                            End If
                        Next ilLoop
                    End If
                    mDGSetShow imDGBoxNo
                    imDGBoxNo = -1
                    imAirRowNo = ilAirRowNo
                    imAirBoxNo = ilBox
                    mAirEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    'If imDGBoxNo > 0 Then
    '    mDGSetFocus imDGBoxNo
    'ElseIf imAirBoxNo > 0 Then
    '    mAirEnableBox imAirBoxNo
    'Else
        cmcAirDone.SetFocus
    'End If
End Sub
Private Sub pbcAirWk_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slFont As String

    ilStartRow = vbcAirWk.Value  'Top location
    ilEndRow = vbcAirWk.Value + vbcAirWk.LargeChange
    If ilEndRow > UBound(smAirShow, 2) - 1 Then
        ilEndRow = UBound(smAirShow, 2) - 1  'don't include blank row
    End If
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBCTCtrls To UBound(tmAirCtrls) Step 1
            If (ilBox >= DWINDEX) And (ilBox <= WKPRICEINDEX) And (imAirSave(2, ilRow) = 1) Then
                gPaintArea pbcAirWk, tmAirCtrls(ilBox).fBoxX, tmAirCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmAirCtrls(ilBox).fBoxW - 15, tmAirCtrls(ilBox).fBoxH - 15, WHITE
            ElseIf (ilBox = WKPRICEINDEX) And ((smLineType = "O") Or (smLineType = "A")) Then
                gPaintArea pbcAirWk, tmAirCtrls(ilBox).fBoxX, tmAirCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmAirCtrls(ilBox).fBoxW - 15, tmAirCtrls(ilBox).fBoxH - 15, WHITE
            End If
            If (Asc(smAirShow(AIRWEEKINDEX, ilRow)) = Asc("Y")) Or ((ilBox < DWINDEX) And (Asc(smAirShow(AIRWEEKINDEX, ilRow)) = Asc("N"))) Then
                pbcAirWk.CurrentX = tmAirCtrls(ilBox).fBoxX + fgBoxInsetX
                If (Not tgClfCntr(imLnRowNo - 1).iGame) Or (ilBox <> DWINDEX) Then
                    If (ilBox >= DAILYSPOTSINDEX) And (ilBox <= DAILYSPOTSINDEX + 6) And (imAirSave(2, ilRow) = 1) Then 'Weekly
                        pbcAirWk.CurrentY = tmAirCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) + 15 '+ fgBoxInsetY
                        slFont = pbcAirWk.FontName
                        pbcAirWk.FontName = "Monotype Sorts"
                        pbcAirWk.FontBold = False
                        pbcAirWk.Print smAirShow(ilBox, ilRow)
                        pbcAirWk.FontName = slFont
                        pbcAirWk.FontBold = True
                    Else
                        If (ilBox >= DAILYSPOTSINDEX) And (ilBox <= DAILYSPOTSINDEX + 6) And (tgClfCntr(imLnRowNo - 1).iGame) Then
                            pbcAirWk.CurrentY = tmAirCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) + 15 '+ fgBoxInsetY
                            slFont = pbcAirWk.FontName
                            pbcAirWk.FontName = "Monotype Sorts"
                            pbcAirWk.FontBold = False
                            pbcAirWk.Print smAirShow(ilBox, ilRow)
                            pbcAirWk.FontName = slFont
                            pbcAirWk.FontBold = True
                        Else
                            If (ilBox > imLBCTCtrls) Or (smAirSave(13, ilRow) <> "Y") Or (ilRow = ilStartRow) Then
                                pbcAirWk.CurrentY = tmAirCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                                pbcAirWk.Print smAirShow(ilBox, ilRow)
                            Else
                                If ilRow < UBound(smAirSave, 2) - 1 Then
                                    If smAirSave(13, ilRow + 1) = "Y" Then
                                        pbcAirWk.CurrentY = tmAirCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                                        pbcAirWk.Print smAirShow(ilBox, ilRow)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next ilBox
    Next ilRow
    For ilBox = imLBCTCtrls To UBound(tmFTCtrls) Step 1
        gPaintArea pbcAirWk, tmFTCtrls(ilBox).fBoxX, tmFTCtrls(ilBox).fBoxY, tmFTCtrls(ilBox).fBoxW - 15, tmFTCtrls(ilBox).fBoxH - 15, WHITE
        pbcAirWk.CurrentX = tmFTCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcAirWk.CurrentY = tmFTCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
        pbcAirWk.Print tmFTCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcCalendar_Click()
    'gShowHelpMess tmClfHelp(), CLFCALENDAR
End Sub
Private Sub pbcCalendar_GotFocus()
    'If plcPower.Visible Then
    '    tmcDelay.Enabled = False
    '    imDelaySrce = -1
    'End If
End Sub
Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    mAirSetShow imAirBoxNo
    imAirBoxNo = -1
    pbcWkArrow.Visible = False
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                'If imNoBoxNo = NODATEINDEX Then
                '    edcShareDropDown.Text = Format$(llDate, "m/d/yy")
                '    edcShareDropDown.SelStart = 0
                '    edcShareDropDown.SelLength = Len(edcShareDropDown.Text)
                '    edcShareDropDown.SetFocus
                '    Exit Sub
                'ElseIf (imDGBoxNo = DGSTARTDATEINDEX) Or (imDGBoxNo = DGENDDATEINDEX) Then
                    edcDGDropDown.Text = Format$(llDate, "m/d/yy")
                    edcDGDropDown.SelStart = 0
                    edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
                    edcDGDropDown.SetFocus
                    Exit Sub
                'End If
                'If plcPower.Visible Then
                '    tmcDelay.Enabled = False
                '    imDelaySrce = -1
                '    If (imPBDateIndex = 0) Or (imPBDateIndex = 1) Then
                '        edcPBDate(imPBDateIndex).Text = Format$(llDate, "m/d/yy")
                '        edcPBDate(imPBDateIndex).SelStart = 0
                '        edcPBDate(imPBDateIndex).SelLength = Len(edcPBDate(imPBDateIndex).Text)
                '        edcPBDate(imPBDateIndex).SetFocus
                '        'DoEvents
                '        'mPBSetCommands
                '    End If
                '    Exit Sub
                'End If
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    'If imNoBoxNo = NODATEINDEX Then
    '    edcShareDropDown.SetFocus
    'ElseIf (imDGBoxNo = DGSTARTDATEINDEX) Or (imDGBoxNo = DGENDDATEINDEX) Then
        edcDGDropDown.SetFocus
    'Else
    '    If plcPower.Visible Then
    '        If (imPBDateIndex = 0) Or (imPBDateIndex = 1) Then
    '            edcPBDate(imPBDateIndex).SetFocus
    '        End If
    '    End If
    'End If
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    If imSettingCD Then
        imSettingCD = False
        Exit Sub
    End If
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub

Private Sub pbcClickFocus_GotFocus()
    mDGSetShow imDGBoxNo
    imDGBoxNo = -1
    mAirSetShow imAirBoxNo
    imAirBoxNo = -1
    pbcWkArrow.Visible = False
End Sub

Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    'If (imLnBoxNo <> OV1INDEX) And (imLnBoxNo <> OV2INDEX) Then
                        Select Case imDGBoxNo
                            Case DGOVSTARTINDEX
                                imBypassFocus = True    'Don't change select text
                                edcDGDropDown.SetFocus
                                'SendKeys slKey
                                gSendKeys edcDGDropDown, slKey
                            Case DGOVENDINDEX
                                imBypassFocus = True    'Don't change select text
                                edcDGDropDown.SetFocus
                                'SendKeys slKey
                                gSendKeys edcDGDropDown, slKey
                        End Select
                    'Else
                    '    If imOVType <> 2 Then
                    '        Select Case imOVBoxNo
                    '            Case OVSTARTTIMEINDEX
                    '                imBypassFocus = True    'Don't change select text
                    '                edcDGDropDown.SetFocus
                    '                'SendKeys slKey
                    '                gSendKeys edcDGDropDown, slKey
                    '            Case OVENDTIMEINDEX
                    '                imBypassFocus = True    'Don't change select text
                    '                edcDGDropDown.SetFocus
                    '                'SendKeys slKey
                    '                gSendKeys edcDGDropDown, slKey
                    '        End Select
                    '    Else
                    '        Select Case imOVBoxNo
                    '            Case OV2STARTTIMEINDEX
                    '                imBypassFocus = True    'Don't change select text
                    '                edcDGDropDown.SetFocus
                    '                'SendKeys slKey
                    '                gSendKeys edcDGDropDown, slKey
                    '            Case OV2ENDTIMEINDEX
                    '                imBypassFocus = True    'Don't change select text
                    '                edcDGDropDown.SetFocus
                    '                'SendKeys slKey
                    '                gSendKeys edcDGDropDown, slKey
                    '        End Select
                    '    End If
                    'End If
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetAirDateFromSpots            *
'*                                                     *
'*             Created:7/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Change start/End date of week  *
'*                      based on days spots assigned   *
'*                                                     *
'*******************************************************
Private Sub mSetAirDateFromSpots()
    Dim ilDay As Integer
    Dim ilStartDay As Integer
    Dim ilEndDay As Integer
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim slDate As String
    'Don't shrink dates
    Exit Sub
    ilStartDay = -1
    ilEndDay = -1
    If (smAirSave(3, imAirRowNo) = "") And (Not tgClfCntr(imLnRowNo - 1).iGame) Then
        For ilLoop = 0 To 6 Step 1
            If Val(smAirSave(4 + ilLoop, imAirRowNo)) > 0 Then
                If ilStartDay = -1 Then
                    ilStartDay = ilLoop
                End If
                ilEndDay = ilLoop
            End If
        Next ilLoop
    Else
        For ilLoop = 0 To 6 Step 1
            If smAirSave(4 + ilLoop, imAirRowNo) = "Y" Then
                If ilStartDay = -1 Then
                    ilStartDay = ilLoop
                End If
                ilEndDay = ilLoop
            End If
        Next ilLoop
    End If
    ilDay = gWeekDayStr(smAirSave(1, imAirRowNo))
    If ilStartDay <> ilDay Then
        If ilStartDay <> -1 Then
            llDate = gDateValue(smAirSave(1, imAirRowNo))
            llDate = llDate - ilDay + ilStartDay
            slDate = Format$(llDate, "m/d/yy")
            smAirSave(1, imAirRowNo) = slDate
            slDate = gAddDayToDate(slDate)
            gSetShow pbcAirWk, slDate, tmAirCtrls(WEEKSTARTDATEINDEX)
            smAirShow(WEEKSTARTDATEINDEX, imAirRowNo) = tmAirCtrls(WEEKSTARTDATEINDEX).sShow
            gPaintArea pbcAirWk, tmAirCtrls(WEEKSTARTDATEINDEX).fBoxX, tmAirCtrls(WEEKSTARTDATEINDEX).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15), tmAirCtrls(WEEKSTARTDATEINDEX).fBoxW - 15, tmAirCtrls(WEEKSTARTDATEINDEX).fBoxH - 15, WHITE
            pbcAirWk.CurrentX = tmAirCtrls(WEEKSTARTDATEINDEX).fBoxX + fgBoxInsetX
            pbcAirWk.CurrentY = tmAirCtrls(WEEKSTARTDATEINDEX).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            pbcAirWk.Print smAirShow(WEEKSTARTDATEINDEX, imAirRowNo)
        End If
    End If
    ilDay = gWeekDayStr(smAirSave(2, imAirRowNo))
    If ilEndDay <> ilDay Then
        If ilEndDay <> -1 Then
            llDate = gDateValue(smAirSave(2, imAirRowNo))
            llDate = llDate - ilDay + ilEndDay
            slDate = Format$(llDate, "m/d/yy")
            smAirSave(2, imAirRowNo) = slDate
            slDate = gAddDayToDate(slDate)
            gSetShow pbcAirWk, slDate, tmAirCtrls(WEEKENDDATEINDEX)
            smAirShow(WEEKENDDATEINDEX, imAirRowNo) = tmAirCtrls(WEEKENDDATEINDEX).sShow
            gPaintArea pbcAirWk, tmAirCtrls(WEEKENDDATEINDEX).fBoxX, tmAirCtrls(WEEKENDDATEINDEX).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15), tmAirCtrls(WEEKENDDATEINDEX).fBoxW - 15, tmAirCtrls(WEEKENDDATEINDEX).fBoxH - 15, WHITE
            pbcAirWk.CurrentX = tmAirCtrls(WEEKENDDATEINDEX).fBoxX + fgBoxInsetX
            pbcAirWk.CurrentY = tmAirCtrls(WEEKENDDATEINDEX).fBoxY + (imAirRowNo - vbcAirWk.Value) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            pbcAirWk.Print smAirShow(WEEKENDDATEINDEX, imAirRowNo)
        End If
    End If
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
    'If imNoBoxNo = NODATEINDEX Then
    '    slStr = edcShareDropDown.Text
    'ElseIf (imDGBoxNo = DGSTARTDATEINDEX) Or (imDGBoxNo = DGENDDATEINDEX) Then
        slStr = edcDGDropDown.Text
    'Else
    '    If plcPower.Visible Then
    '        If (imPBDateIndex = 0) Or (imPBDateIndex = 1) Then
    '            slStr = edcPBDate(imPBDateIndex).Text
    '        End If
    '    End If
    'End If
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

Private Sub cmcGenNew_Click()
    'This button will cause only new weeks to be added- old week can be replaced if airing
    'This sub is used after a line has flight weeks defined.
    Dim ilRet As Integer
    Dim slStr As String
    'Test that past billed past weeks are not removed

    If (tgClfCntr(imLnRowNo - 1).iInPast) Then  'Line started in the past (base on Least of LLD or ToDays Date)
        If (tgClfCntr(imLnRowNo - 1).iStatus > 0) Then  'Old line
            If (lmAllowedChgWkStartDate > gDateValue(smDGSave(10))) Then
                mDGSetShow imDGBoxNo
                Beep
                'ilRet = MsgBox("Dates prior or equal to " & Format$(lmAllowedChgWkStartDate, "m/d/yy") & " will not be altered", vbOkCancel + vbExclamation, "Warning")
                ilRet = MsgBox("Dates prior to " & Format$(lmAllowedChgWkStartDate, "m/d/yy") & " will not be altered", vbOKCancel + vbExclamation, "Warning")
                If ilRet = vbCancel Then
                    imDGBoxNo = DGENDDATEINDEX
                    mDGEnableBox imDGBoxNo
                    Exit Sub
                End If
                '2/2/14: Disallow dates in past to be created
                smDGSave(10) = gObtainNextMonday(Format$(lmAllowedChgWkStartDate, "m/d/yy"))
                slStr = smDGSave(10)
                gSetShow pbcDGen, slStr, tmDGCtrls(DGSTARTDATEINDEX)
                pbcDGen.Refresh
            End If
        End If
    End If
    If gDateValue(smDGSave(11)) < gDateValue(smDGSave(10)) Then
        mDGSetShow imDGBoxNo
        Beep
        MsgBox "End Date must be after or equal to " & smDGSave(10), vbOKOnly + vbExclamation, "Warning"
        imDGBoxNo = DGENDDATEINDEX
        mDGEnableBox imDGBoxNo
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mDGSetShow imDGBoxNo
    imDGBoxNo = -1
    mCreateNewDate
    pbcAirSTab.SetFocus
    mDGSetCommands
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmcGenNew_GotFocus()
    'gShowHelpMess tmClfHelp(), CLFBTGENNEW
    mAirSetShow imAirBoxNo  'Line air week
    imAirBoxNo = -1
    imAirRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcAirCancel_Click()
    If imSvLnRowNo <> imLnRowNo Then
        imLnRowNo = imSvLnRowNo
        smLineType = mGetLineType(imLnRowNo)
        mGetDate True
        imSvLnRowNo = -1
    End If
    imAirCancel = True
    'mLnSetShow imLnBoxNo    'True grid off
    mFlightReturn
    imLnChg = imSvLnChg
    imComputeReq = imSvComputeReq
    'mSetCommands
    '10/20/15: Changed from Comment to Price becaus ethe expansion of package lines moved from Comment to Price
    'Alternate solution would be imLnBoxNo = wkspotsindex + 12 and edcLnDropDown.Text = "P"
    imLnBoxNo = -PRICEINDEX '-LINECOMMINDEX
    'mLnEnableBox imLnBoxNo
    Screen.MousePointer = vbDefault
    imInFlightScreen = False
    igTerminateReturn = 0   'Cancel selected
    mTerminate
End Sub
Private Sub cmcAirCancel_GotFocus()
    mDGSetShow imDGBoxNo
    imDGBoxNo = -1
    mAirSetShow imAirBoxNo
    imAirBoxNo = -1
    pbcWkArrow.Visible = False
    'gShowHelpMess tmClfHelp(), CLFBTFLTCANCEL
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcAirClear_Click()
    Dim ilLoop As Integer
    Dim llDelDate As Long
    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    tgClfCntr(imLnRowNo - 1).iResearch = False
    llDelDate = gDateValue(gObtainNextMonday(gIncOneDay(Format$(gNow(), "m/d/yy"))))
    If lmAllowedChgWkStartDate >= llDelDate Then
        llDelDate = gDateValue(gObtainNextMonday(Format$(lmAllowedChgWkStartDate + 1, "m/d/yy")))
    End If
    'For ilLoop = LBound(smAirSave, 2) To UBound(smAirSave, 2) - 1 Step 1
    For ilLoop = imLB1Or2 To UBound(smAirSave, 2) - 1 Step 1
        If gDateValue(smAirSave(1, ilLoop)) >= llDelDate Then
            'If ilLoop > LBound(smAirSave, 2) Then
            If ilLoop > imLB1Or2 Then
                ReDim Preserve smAirShow(0 To 14, 0 To ilLoop - 1) As String
                ReDim Preserve smAirSave(0 To 13, 0 To ilLoop - 1) As String
                ReDim Preserve imAirSave(0 To 2, 0 To ilLoop - 1) As Integer
                ReDim Preserve smAirShow(0 To 14, 0 To ilLoop) As String
                ReDim Preserve smAirSave(0 To 13, 0 To ilLoop) As String
                ReDim Preserve imAirSave(0 To 2, 0 To ilLoop) As Integer
                Exit For
            Else
                ReDim smAirShow(0 To 14, 0 To ilLoop) As String
                ReDim smAirSave(0 To 13, 0 To ilLoop) As String
                ReDim imAirSave(0 To 2, 0 To ilLoop) As Integer
                Exit For
            End If
        End If
    Next ilLoop
    pbcAirWk.Cls
    mFlightTotals
    mGetDemoFlightTotals
    imSettingValue = True
    vbcAirWk.Min = imLB1Or2 'LBound(smAirShow, 2)
    If UBound(smAirShow, 2) - 2 <= vbcAirWk.LargeChange Then
        vbcAirWk.Max = imLB1Or2 'LBound(smAirShow, 2)
    Else
        vbcAirWk.Max = UBound(smAirShow, 2) - 1 - vbcAirWk.LargeChange
    End If
    If vbcAirWk.Value = vbcAirWk.Min Then
        pbcAirWk_Paint
    Else
        vbcAirWk.Value = vbcAirWk.Min
    End If
    imSettingValue = False
    imAirBoxNo = -1
    imDGBoxNo = -1
    pbcDGSTab.SetFocus
    'If UBound(smAirSave, 2) <= LBound(smAirSave, 2) Then
    If UBound(smAirSave, 2) <= imLB1Or2 Then
        cmcGenNew.Visible = False
        cmcGenNew.Enabled = False
        cmcGenAll.Visible = True
        mDGSetCommands
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmcAirClear_GotFocus()
    mDGSetShow imDGBoxNo
    imDGBoxNo = -1
    mAirSetShow imAirBoxNo
    imAirBoxNo = -1
    pbcWkArrow.Visible = False
End Sub
Private Sub cmcAirDone_Click()
    Dim ilRet As Integer
    Dim slStr As String
    'If UBound(smAirSave, 2) <= LBound(smAirSave, 2) Then
    If UBound(smAirSave, 2) <= imLB1Or2 Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("No Weeks Defined", vbOKOnly + vbExclamation, "Error")
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If imSvLnRowNo <> imLnRowNo Then
        mSetDate False
        tgClfCntr(imLnRowNo - 1).iResearch = False
        imLnSave(8, imLnRowNo) = imDGSave(2)
        Contract.mBuildWkInfo False, False, False
        mGetLnPriceInfo imLnRowNo, slStr, imLnTotalType, PRICEINDEX         'Formats Price within Show
        Contract.mCntrTotal
        imLnRowNo = imSvLnRowNo
        smLineType = mGetLineType(imLnRowNo)
        'imDGSave(2) = imLnSave(8, imLnRowNo)
        mGetDate True
        imSvLnRowNo = -1
        tgClfCntr(imLnRowNo - 1).iResearch = False
    End If
    imAirCancel = False
    'mLnSetShow imLnBoxNo    'True grid off
    'Screen.MousePointer = vbDefault
    mFlightReturn
    '10/20/15: Changed from Comment to Price becaus ethe expansion of package lines moved from Comment to Price
    'Alternate solution would be imLnBoxNo = wkspotsindex + 12 and edcLnDropDown.Text = "P"
    imLnBoxNo = -PRICEINDEX '-LINECOMMINDEX
    Screen.MousePointer = vbDefault
    imInFlightScreen = False
    igTerminateReturn = 1   'Done selected
    mTerminate
End Sub
Private Sub cmcAirDone_GotFocus()
    mDGSetShow imDGBoxNo
    imDGBoxNo = -1
    mAirSetShow imAirBoxNo
    imAirBoxNo = -1
    pbcWkArrow.Visible = False
    'gShowHelpMess tmClfHelp(), CLFBTFLTDONE
    gCtrlGotFocus ActiveControl
End Sub

Private Sub plcAirScreen_Paint()
    plcAirScreen.CurrentX = 0
    plcAirScreen.CurrentY = 0
    plcAirScreen.Print "Flights"
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mGetFlightStartInfo True, False 'Require to call mGetDate as cbcAirLine calls mSetDate
    mPopAirLine
    Screen.MousePointer = vbDefault
End Sub

Private Sub vbcAirWk_Change()
    If imSettingValue Then
        pbcAirWk.Cls
        pbcAirWk_Paint
        imSettingValue = False
    Else
        mAirSetShow imAirBoxNo
        pbcAirWk.Cls
        pbcAirWk_Paint
        mAirEnableBox imAirBoxNo
    End If
End Sub
Private Sub vbcAirWk_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcDGRun_Click()
    If imLbcMouseDown Then
        imLbcArrowSetting = False
    Else
        imLbcArrowSetting = True
    End If
    gProcessLbcClick lbcDGRun, edcDGDropDown, imChgMode, imLbcArrowSetting
    mDGSetCommands
    imLbcMouseDown = False
End Sub
Private Sub lbcDGRun_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcDW_Click()
    Dim slStr As String

    If imLbcMouseDown Then
        imLbcArrowSetting = False
    Else
        imLbcArrowSetting = True
    End If
    gProcessLbcClick lbcDW, edcAirDropDown, imChgMode, imLbcArrowSetting
    If tgSpf.sAllowDailyBuys <> "Y" Then
        slStr = edcAirDropDown.Text
        If InStr(1, slStr, "Daily", 1) > 0 Then
            lbcDW.ListIndex = lbcDW.ListIndex + 1   'Force to weekly
            gProcessLbcClick lbcDW, edcAirDropDown, imChgMode, imLbcArrowSetting
        End If
    End If
    imLbcMouseDown = False
End Sub
Private Sub lbcDW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcDWP_Click()
    Dim slStr As String

    If imLbcMouseDown Then
        imLbcArrowSetting = False
    Else
        imLbcArrowSetting = True
    End If
    gProcessLbcClick lbcDWP, edcDGDropDown, imChgMode, imLbcArrowSetting
    If tgSpf.sAllowDailyBuys <> "Y" Then
        slStr = edcDGDropDown.Text
        If InStr(1, slStr, "Daily", 1) > 0 Then
            lbcDWP.ListIndex = lbcDWP.ListIndex + 1   'Force to weekly
            gProcessLbcClick lbcDWP, edcDGDropDown, imChgMode, imLbcArrowSetting
        End If
    End If
    mDGSetCommands
    imLbcMouseDown = False
End Sub
Private Sub lbcDWP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateNewDate                  *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create only new flights        *
'*                                                     *
'*******************************************************
Private Sub mCreateNewDate()
    Dim ilUpperLimit As Integer
    Dim ilLoop As Integer
    Dim llOrigSDate As Long
    Dim llOrigEDate As Long
    Dim llNewSDate As Long
    Dim llNewEDate As Long
    Dim llDate As Long
    Dim llSuDate As Long
    Dim llEarliestDate As Long
    Dim llLatestDate As Long
    Dim ilNewIndex As Integer
    Dim ilOrigIndex As Integer
    Dim ilWkNo As Integer
    Dim slStr As String
    Dim ilDay As Integer
    Dim ilIndex As Integer
    Dim ilSpots As Integer
    Dim ilMGSpots As Integer
    Dim ilNewWeekStatus As Integer  '-1=Week does not exist; 0=Air Week; 1+=NonAir Week
    'Retain original flights so they can be remapped to weeks created
    ilUpperLimit = UBound(smAirShow, 2)
    ReDim slOrigAirShow(0 To 14, 0 To ilUpperLimit) As String
    ReDim slOrigAirSave(0 To 13, 0 To ilUpperLimit) As String
    ReDim ilOrigAirSave(0 To 2, 0 To ilUpperLimit) As Integer
    Dim ilBilledCount As Integer
    
    For ilLoop = LBound(smAirSave, 2) To UBound(smAirSave, 2) Step 1
        For ilIndex = LBound(smAirSave, 1) To UBound(smAirSave, 1) Step 1
            slOrigAirSave(ilIndex, ilLoop) = smAirSave(ilIndex, ilLoop)
        Next ilIndex
    Next ilLoop
    For ilLoop = LBound(imAirSave, 2) To UBound(imAirSave, 2) Step 1
        For ilIndex = LBound(imAirSave, 1) To UBound(imAirSave, 1) Step 1
            ilOrigAirSave(ilIndex, ilLoop) = imAirSave(ilIndex, ilLoop)
        Next ilIndex
    Next ilLoop
    For ilLoop = LBound(smAirShow, 2) To UBound(smAirShow, 2) Step 1
        For ilIndex = LBound(smAirShow, 1) To UBound(smAirShow, 1) Step 1
            slOrigAirShow(ilIndex, ilLoop) = smAirShow(ilIndex, ilLoop)
        Next ilIndex
    Next ilLoop
    mCreateDate False 'Create all dates
    ilUpperLimit = UBound(smAirShow, 2)
    ReDim slNewAirShow(0 To 14, 0 To ilUpperLimit) As String
    ReDim slNewAirSave(0 To 13, 0 To ilUpperLimit) As String
    ReDim ilNewAirSave(0 To 2, 0 To ilUpperLimit) As Integer
    For ilLoop = LBound(smAirSave, 2) To UBound(smAirSave, 2) Step 1
        For ilIndex = LBound(smAirSave, 1) To UBound(smAirSave, 1) Step 1
            slNewAirSave(ilIndex, ilLoop) = smAirSave(ilIndex, ilLoop)
        Next ilIndex
    Next ilLoop
    For ilLoop = LBound(imAirSave, 2) To UBound(imAirSave, 2) Step 1
        For ilIndex = LBound(imAirSave, 1) To UBound(imAirSave, 1) Step 1
            ilNewAirSave(ilIndex, ilLoop) = imAirSave(ilIndex, ilLoop)
        Next ilIndex
    Next ilLoop
    For ilLoop = LBound(smAirShow, 2) To UBound(smAirShow, 2) Step 1
        For ilIndex = LBound(smAirShow, 1) To UBound(smAirShow, 1) Step 1
            slNewAirShow(ilIndex, ilLoop) = smAirShow(ilIndex, ilLoop)
        Next ilIndex
    Next ilLoop


    'Merge flights (Index 1=Start Date; 2=End Date)
    If UBound(slOrigAirSave, 2) > LBound(slOrigAirSave, 2) Then
        llOrigSDate = gDateValue(slOrigAirSave(1, 1))
        llOrigEDate = gDateValue(slOrigAirSave(2, UBound(slOrigAirSave, 2) - 1))
    Else
        llOrigSDate = 0
        llOrigEDate = 0
    End If
    llNewSDate = gDateValue(slNewAirSave(1, 1))
    llNewEDate = gDateValue(slNewAirSave(2, UBound(smAirSave, 2) - 1))
    If (llOrigSDate < llNewSDate) And (llOrigSDate > 0) Then
        llEarliestDate = llOrigSDate
    Else
        llEarliestDate = llNewSDate
    End If
    If (llOrigEDate > llNewEDate) And (llOrigEDate > 0) Then
        llLatestDate = llOrigEDate
    Else
        llLatestDate = llNewEDate
    End If
    ReDim smAirShow(0 To 14, 0 To 1) As String
    ReDim smAirSave(0 To 13, 0 To 1) As String
    ReDim imAirSave(0 To 2, 0 To 1) As Integer
    ilUpperLimit = UBound(smAirShow, 2)
    llDate = llEarliestDate
    llSuDate = llDate
    Do While gWeekDayLong(llSuDate) <> 6
        llSuDate = llSuDate + 1
    Loop
    ilWkNo = 1
    Do
        'Determine if new or original should be used
        ilOrigIndex = -1
        For ilLoop = LBound(slOrigAirSave, 2) To UBound(slOrigAirSave, 2) - 1 Step 1
            If ((gDateValue(slOrigAirSave(1, ilLoop)) >= llDate) And (gDateValue(slOrigAirSave(1, ilLoop)) <= llSuDate)) Or ((gDateValue(slOrigAirSave(2, ilLoop)) >= llDate) And (gDateValue(slOrigAirSave(2, ilLoop)) <= llSuDate)) Then
                If ilOrigAirSave(1, ilLoop) = 0 Then    'Air week
                    ilOrigIndex = ilLoop
                End If
                Exit For
            End If
        Next ilLoop
        ilNewIndex = -1
        ilNewWeekStatus = -1
        For ilLoop = LBound(slNewAirSave, 2) To UBound(slNewAirSave, 2) - 1 Step 1
            If ((gDateValue(slNewAirSave(1, ilLoop)) >= llDate) And (gDateValue(slNewAirSave(1, ilLoop)) <= llSuDate)) Or ((gDateValue(slNewAirSave(2, ilLoop)) >= llDate) And (gDateValue(slNewAirSave(2, ilLoop)) <= llSuDate)) Then
                If ilNewAirSave(1, ilLoop) = 0 Then    'Air week
                    ilNewIndex = ilLoop
                    ilNewWeekStatus = 0
                Else
                    ilNewWeekStatus = ilLoop
                End If
                Exit For
            End If
        Next ilLoop
        If ilOrigIndex > -1 Then
            If ilOrigAirSave(2, ilOrigIndex) = 0 Then 'Daily
                ilSpots = 0
                For ilDay = 4 To 10 Step 1
                    ilSpots = ilSpots + Val(slOrigAirSave(ilDay, ilOrigIndex))
                Next ilDay
            Else
                ilSpots = Val(slOrigAirSave(3, ilOrigIndex))
            End If
        Else
            ilSpots = -1
        End If
        If ((ilOrigIndex > -1) And (ilNewIndex > -1)) Or ((ilOrigIndex > -1) And (ilNewIndex = -1) And (ilNewWeekStatus > -1) And (ilSpots = 0)) Then   'Merge
            'Replace original with new if original was an air week or Zero spots and not in the past
            'Later we might want to update specified fields only
            If ilNewIndex = -1 Then
                ilNewIndex = ilNewWeekStatus
            End If
            '6/11/12
            ilMGSpots = mGetMGCount(imLnRowNo, hmCHF, hmClf, hmSdf, hmSmf, gDateValue(slOrigAirSave(1, ilOrigIndex)), gDateValue(slOrigAirSave(1, ilOrigIndex)) + 6, ilBilledCount)
            If lmAllowedChgWkStartDate <= gDateValue(slOrigAirSave(1, ilOrigIndex)) And (ilMGSpots <= 0) Then
                For ilIndex = LBound(smAirSave, 1) To UBound(smAirSave, 1) Step 1
                    smAirSave(ilIndex, ilUpperLimit) = slNewAirSave(ilIndex, ilNewIndex)
                Next ilIndex
                For ilIndex = LBound(imAirSave, 1) To UBound(imAirSave, 1) Step 1
                    imAirSave(ilIndex, ilUpperLimit) = ilNewAirSave(ilIndex, ilNewIndex)
                Next ilIndex
                For ilIndex = LBound(smAirShow, 1) To UBound(smAirShow, 1) Step 1
                    smAirShow(ilIndex, ilUpperLimit) = slNewAirShow(ilIndex, ilNewIndex)
                Next ilIndex
            Else
                For ilIndex = LBound(smAirSave, 1) To UBound(smAirSave, 1) Step 1
                    smAirSave(ilIndex, ilUpperLimit) = slOrigAirSave(ilIndex, ilOrigIndex)
                Next ilIndex
                For ilIndex = LBound(imAirSave, 1) To UBound(imAirSave, 1) Step 1
                    imAirSave(ilIndex, ilUpperLimit) = ilOrigAirSave(ilIndex, ilOrigIndex)
                Next ilIndex
                For ilIndex = LBound(smAirShow, 1) To UBound(smAirShow, 1) Step 1
                    smAirShow(ilIndex, ilUpperLimit) = slOrigAirShow(ilIndex, ilOrigIndex)
                Next ilIndex
            End If
        ElseIf (ilOrigIndex > -1) Then
            For ilIndex = LBound(smAirSave, 1) To UBound(smAirSave, 1) Step 1
                smAirSave(ilIndex, ilUpperLimit) = slOrigAirSave(ilIndex, ilOrigIndex)
            Next ilIndex
            For ilIndex = LBound(imAirSave, 1) To UBound(imAirSave, 1) Step 1
                imAirSave(ilIndex, ilUpperLimit) = ilOrigAirSave(ilIndex, ilOrigIndex)
            Next ilIndex
            For ilIndex = LBound(smAirShow, 1) To UBound(smAirShow, 1) Step 1
                smAirShow(ilIndex, ilUpperLimit) = slOrigAirShow(ilIndex, ilOrigIndex)
            Next ilIndex
        ElseIf (ilNewIndex > -1) Then
            For ilIndex = LBound(smAirSave, 1) To UBound(smAirSave, 1) Step 1
                smAirSave(ilIndex, ilUpperLimit) = slNewAirSave(ilIndex, ilNewIndex)
            Next ilIndex
            For ilIndex = LBound(imAirSave, 1) To UBound(imAirSave, 1) Step 1
                imAirSave(ilIndex, ilUpperLimit) = ilNewAirSave(ilIndex, ilNewIndex)
            Next ilIndex
            For ilIndex = LBound(smAirShow, 1) To UBound(smAirShow, 1) Step 1
                smAirShow(ilIndex, ilUpperLimit) = slNewAirShow(ilIndex, ilNewIndex)
            Next ilIndex
        Else
            'Create non air week
            imAirSave(1, ilUpperLimit) = 1
            slStr = "No"
            gSetShow pbcAirWk, slStr, tmAirCtrls(AIRWEEKINDEX)
            smAirShow(AIRWEEKINDEX, ilUpperLimit) = tmAirCtrls(AIRWEEKINDEX).sShow
            slStr = Format$(llDate, "m/d/yy")
            smAirSave(1, ilUpperLimit) = slStr
            slStr = Format$(llSuDate, "m/d/yy")
            smAirSave(2, ilUpperLimit) = slStr
            slStr = ""
            smAirSave(3, ilUpperLimit) = slStr
            gSetShow pbcAirWk, slStr, tmAirCtrls(SPOTPERWKINDEX)
            smAirShow(SPOTPERWKINDEX, ilUpperLimit) = tmAirCtrls(SPOTPERWKINDEX).sShow
            For ilDay = 0 To 6 Step 1   'Initialize days
                slStr = ""
                smAirSave(4 + ilDay, ilUpperLimit) = slStr
                gSetShow pbcAirWk, slStr, tmAirCtrls(DAILYSPOTSINDEX + ilDay)
                smAirShow(DAILYSPOTSINDEX + ilDay, ilUpperLimit) = tmAirCtrls(DAILYSPOTSINDEX + ilDay).sShow
            Next ilDay
            smAirSave(13, ilUpperLimit) = "N"
            If (imDGSave(2) = 0) Or (imDGSave(2) = 2) Then    'Daily Or % Daily
                imAirSave(2, ilUpperLimit) = 0
                If imDGSave(2) > 1 Then
                    slStr = "Daily by %"
                Else
                    slStr = "Daily"
                End If
            Else
                imAirSave(2, ilUpperLimit) = 1
                If imDGSave(2) > 1 Then
                    slStr = "Weekly by %"
                Else
                    slStr = "Weekly"
                End If
            End If
            gSetShow pbcAirWk, slStr, tmAirCtrls(DWINDEX)
            smAirShow(DWINDEX, ilUpperLimit) = tmAirCtrls(DWINDEX).sShow
            mAdjDates ilUpperLimit
        End If
        slStr = Trim$(str$(ilWkNo))
        gSetShow pbcAirWk, slStr, tmAirCtrls(WEEKNOINDEX)
        smAirShow(WEEKNOINDEX, ilUpperLimit) = tmAirCtrls(WEEKNOINDEX).sShow
        ilUpperLimit = ilUpperLimit + 1
        ReDim Preserve smAirShow(0 To 14, 0 To ilUpperLimit) As String
        ReDim Preserve smAirSave(0 To 13, 0 To ilUpperLimit) As String
        ReDim Preserve imAirSave(0 To 2, 0 To ilUpperLimit) As Integer
        'Next Monday
        ilWkNo = ilWkNo + 1
        llDate = llDate + 1
        Do While gWeekDayLong(llDate) <> 0
            llDate = llDate + 1
        Loop
        llSuDate = llDate + 6
    Loop While llDate <= llLatestDate
    smLineType = mGetLineType(imLnRowNo)
    If (smLineType = "O") Or (smLineType = "A") Or (smLineType = "E") Then
        mSetAirPkPrice
    End If
    mFlightTotals
    mGetDemoFlightTotals
    vbcAirWk.Min = imLB1Or2 'LBound(smAirShow, 2)
    If UBound(smAirShow, 2) - 2 <= vbcAirWk.LargeChange Then
        vbcAirWk.Max = imLB1Or2 'LBound(smAirShow, 2)
    Else
        vbcAirWk.Max = UBound(smAirShow, 2) - 1 - vbcAirWk.LargeChange
    End If
    vbcAirWk.Value = vbcAirWk.Min
    If ilUpperLimit > 1 Then
        smLnSave(14, imLnRowNo) = smAirSave(1, 1)
        smLnSave(15, imLnRowNo) = smAirSave(2, ilUpperLimit - 1)
    Else
        smLnSave(14, imLnRowNo) = ""
        smLnSave(15, imLnRowNo) = ""
    End If
    imAirRowNo = 1
    pbcAirWk.Cls
    pbcAirWk_Paint
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFlightReturn                   *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process return from flight     *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mFlightReturn()
    Dim slStr As String
    Dim ilSvFreezeState As Integer
    ReDim ilFlDays(0 To 6) As Integer

    tgSpf.sAllowDailyBuys = smSvAllowDailyWeekly
    Contract!cmcDone.Enabled = True
    Contract!cmcCancel.Enabled = True
    'cmcPower.Enabled = True
    Contract!cmcRates.Enabled = True
    mAirSetShow imAirBoxNo
    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        pbcDGen.Visible = False
        pbcAir.Visible = False
        pbcAirWk.Visible = False
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If imInFlightReturn Then
        Exit Sub
    End If
    imInFlightReturn = True
    'Check that current line is still the one with data- focus sent
    'to another control without pressing done or cancel; treat as if
    'done pressed
    If (imSvLnRowNo <> imLnRowNo) And (imLnRowNo > 0) And (imSvLnRowNo > 0) Then
        mSetDate False
        imLnSave(8, imLnRowNo) = imDGSave(2)
        Contract.mBuildWkInfo False, False, False
        mGetLnPriceInfo imLnRowNo, slStr, imLnTotalType, PRICEINDEX         'Formats Price within Show
        imLnRowNo = imSvLnRowNo
        'imDGSave(2) = imLnSave(8, imLnRowNo)
        mGetDate True
    End If
    imSvLnRowNo = -1
    'If imAirCancel = False Then
        mSetDate False
    'End If
    imLnSave(8, imLnRowNo) = imDGSave(2)
    mGetDate True    'reset as this control is called twice (once for done or cancel, next for advancing control)
                'Also required to set pattern flag
    Contract.mPrgPop imLnRowNo, "L"
    mGetFlightDays imLnRowNo, ilFlDays()
    slStr = mSetDPShowName(imLnRowNo, DPINDEX, ilFlDays())
    Contract.mBuildWkInfo True, False, False
    If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
        Contract.mSetHdDate False, False
    Else
        Contract.mSetHdDate True, False
    End If
    If Contract!plcDemoBar.Visible Then
        If imDBTotals < 5 Then
            Contract!pbcDemoBar(0).Cls
            Contract!pbcDemoBar(1).Cls
            tgClfCntr(imLnRowNo - 1).iResearch = False
            'mGetResearchTotals False
            mGetResearchTotals imLnRowNo, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imDBTotals, False, imCppCpm, imCntrVefCode, imVpfIndex, imTerminate, imChgMode, imInResearch, imLnTotalType
            Contract.pbcDemoBar_Paint 1
            Contract.pbcDemoBar_Paint 0
        Else
            Contract.pbcDemoBar(1).Cls
            Contract.pbcDemoBar_Paint 1
        End If
    End If
    'Determine if HSCur can be reset
    pbcDGen.Visible = False
    pbcAir.Visible = False
    pbcAirWk.Visible = False
    If (lmSvHSStartDate + 7 * (imSvHSCurWkNo - 1) >= lmHSStartDate) And (lmSvHSStartDate + 7 * (imSvHSCurWkNo - 1) <= lmHSStartDate + 7 * (imHSEndWkNo - 1)) Then
        imHSCurWkNo = ((lmSvHSStartDate + 7 * (imSvHSCurWkNo - 1)) - lmHSStartDate) \ 7 + 1
    End If
    mGetFlightStartInfo True, True
    ilSvFreezeState = igFreezeState
    igFreezeState = 0
    mLnTotals imLnRowNo, imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), "N", imNoWkDays()
    igFreezeState = ilSvFreezeState
    tgClfCntr(imLnRowNo - 1).iSPChg = False
    igLnSPChg = False
    mGetLnPriceInfo imLnRowNo, slStr, imLnTotalType, PRICEINDEX         'Formats Price within Show
    Contract.mCntrTotal
    imAirCancel = False
    imAirBoxNo = -1
    'If plcDemoBar.Visible Then
    '    pbcDemoBar(1).Cls
    '    mBuildWkInfo True
    'End If
    imInFlightReturn = False
    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopAirLine                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate with Lines not CBS    *
'*                      used in the flight screen      *
'*                                                     *
'*******************************************************
Private Sub mPopAirLine()
    Dim ilClf As Integer
    'Dim ilVef As Integer
    Dim ilVeh As Integer
    Dim slLineType As String
    Dim ilLineIndex As Integer
    Dim slNameCode As String
    'Dim slCode As String
    Dim slVefName As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilMax As Integer
    Dim ilAddLine As Integer
    cbcAirLine.Clear
    ilLineIndex = -1
    ilMax = UBound(tgClfCntr) - 1
    If imLnRowNo = UBound(tgClfCntr) + 1 Then
        'ilMax = UBound(tgClfCntr)
        If UBound(smLnSave, LINEBOUNDINDEX) <= UBound(tgClfCntr) Then
            ilMax = UBound(tgClfCntr) - 1
        Else
            ilMax = UBound(tgClfCntr)
        End If
    End If
    For ilClf = LBound(tgClfCntr) To ilMax Step 1
        ilAddLine = True
        If tgClfCntr(ilClf).iCancel Then
            If imLnRowNo <> ilClf + 1 Then
                ilAddLine = False
            End If
        End If
        If mIsLnGame(ilClf + 1) Then
            ilAddLine = False
        End If
        slLineType = mGetLineType(ilClf + 1)
        If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
            If mGameInPkg(ilClf + 1) Then
                ilAddLine = False
            End If
        End If
        If ilAddLine Then
            slLineType = mGetLineType(ilClf + 1)
            If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                gFindMatch smLnSave(1, ilClf + 1), 0, Contract.lbcLnVehicle(igTabMapIndex)
                If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                    slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
                    ilRet = gParseItem(slNameCode, 1, "\", slVefName)
                    ilRet = gParseItem(slVefName, 3, "|", slVefName)
                    'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    '    If imLnSave(1, ilClf + 1) = tgMRdf(ilRdf).iCode Then
                            slStr = mMakeDPName(ilClf + 1)  'mMakePrgName(tgMRdf(ilRdf).iCode)
                            If tgClfCntr(ilClf).iCancel Then
                                cbcAirLine.AddItem Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine)) & " " & Trim$(slVefName) & " " & slStr & " " & smLnSave(16, ilClf + 1) & "s " & "Canceled"
                            Else
                                cbcAirLine.AddItem Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine)) & " " & Trim$(slVefName) & " " & slStr & " " & smLnSave(16, ilClf + 1) & "s " & smLnSave(14, ilClf + 1) & "-" & smLnSave(15, ilClf + 1)
                            End If
                            If imLnRowNo = ilClf + 1 Then
                                ilLineIndex = cbcAirLine.ListCount - 1
                            End If
                    '        Exit For
                    '    End If
                    'Next ilRdf
                End If
            Else
                For ilVeh = LBound(tgPackVeh) To UBound(tgPackVeh) - 1 Step 1
                    If StrComp(Trim$(smLnSave(1, ilClf + 1)), Trim$(tgPackVeh(ilVeh).sName), 1) = 0 Then
                        slVefName = tgPackVeh(ilVeh).sName
                        'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                        '    If StrComp(smLnSave(20, ilClf + 1), Trim$(tgMRdf(ilRdf).sName), 1) = 0 Then
                                slStr = mMakeDPName(ilClf + 1)  'mMakePrgName(tgMRdf(ilRdf).iCode)
                                If tgClfCntr(ilClf).iCancel Then
                                    cbcAirLine.AddItem Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine)) & " " & Trim$(slVefName) & " " & slStr & " " & smLnSave(16, ilClf + 1) & "s " & "Canceled"
                                Else
                                    cbcAirLine.AddItem Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine)) & " " & Trim$(slVefName) & " " & slStr & " " & smLnSave(16, ilClf + 1) & "s " & smLnSave(14, ilClf + 1) & "-" & smLnSave(15, ilClf + 1)
                                End If
                                If imLnRowNo = ilClf + 1 Then
                                    ilLineIndex = cbcAirLine.ListCount - 1
                                End If
                        '        Exit For
                        '    End If
                        'Next ilRdf
                    End If
                Next ilVeh
            End If
        End If
    Next ilClf
    If ilLineIndex >= 0 Then
        cbcAirLine.ListIndex = ilLineIndex
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetFlightStartInfo             *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Flight Start Information   *
'*                      used in the flight screen      *
'*                                                     *
'*******************************************************
Public Sub mGetFlightStartInfo(ilMakeShowOnly As Integer, ilRetainZeroWks As Integer)
'
'   mGetFlightStartInfo ilMakeShowOnly, ilRetainZeroWks
'   Where:
'       ilMakeShowOnly(I)- True if only want to set arrays (only only to avoid double painting)
'       ilRetainZeroWks(I)- True if retain weeks without spots; False means remove weeks (starting flights)
'
    Dim slStr As String
    Dim slDate As String
    Dim ilNoWks As Integer
    Dim llEndDate As Long
    Dim llStartDate As Long
    Dim ilDay As Integer
    Dim ilPattern As Integer
    Dim ilPatternOK As Integer
    Dim ilWkNo As Integer
    Dim ilStartDGNew As Integer
    Dim ilAllBlank As Integer
    Dim ilLoop As Integer

    lmAllowedChgWkStartDate = mLnDate(imLnRowNo, smLnLLD, 0)
    If lmAllowedChgWkStartDate < tgClfCntr(imLnRowNo - 1).lEndBilledDate Then
        lmAllowedChgWkStartDate = tgClfCntr(imLnRowNo - 1).lEndBilledDate
    End If
    imCalType = Contract.lbcBillCycle.ListIndex
    'If (imCalType = 1) Or (imCalType = 3) Then
    If (imCalType = 1) Then
        If Left(Contract.lbcBillCycle.List(Contract.lbcBillCycle.ListIndex), 1) = "C" Then
            imCalType = 1
        Else
            imCalType = 0
        End If
    Else
        imCalType = 0
    End If
    mGetDate ilRetainZeroWks
    If smLnSave(14, imLnRowNo) <> "" Then   'Start Date defined
        'cmcGenNew.Enabled = True
        cmcGenAll.Visible = False
        cmcGenNew.Visible = True
        cmcGenNew.Enabled = True
        imGenReq = False
        smDGSave(10) = smLnSave(14, imLnRowNo)
        smDGSave(11) = smLnSave(15, imLnRowNo)
        smOrigStartDate = smDGSave(10)
        smOrigEndDate = smDGSave(11)
        If gDateValue(smDGSave(10)) < lmAllowedChgWkStartDate Then  'First unbilled date or Vehicle Last Log Date
            smDGSave(10) = gObtainNextMonday(Format$(lmAllowedChgWkStartDate, "m/d/yy"))
            If gDateValue(smDGSave(11)) < gDateValue(smDGSave(10)) Then
                smDGSave(11) = gObtainNextSunday(smDGSave(10))
            End If
        ElseIf gDateValue(smDGSave(10)) < gDateValue(smLnLLD) Then  'Todays date or Vehicle last log date
            smDGSave(10) = gObtainNextMonday(Format$(smLnLLD, "m/d/yy"))
            If gDateValue(smDGSave(11)) < gDateValue(smDGSave(10)) Then
                smDGSave(11) = gObtainNextSunday(smDGSave(10))
            End If
        End If
        llEndDate = gDateValue(smDGSave(11))
        llStartDate = gDateValue(smDGSave(10))
        ilNoWks = (llEndDate - llStartDate + 7) \ 7
        smDGSave(1) = Trim$(str$(ilNoWks))
    Else
        'cmcGenNew.Enabled = False
        cmcGenNew.Visible = False
        cmcGenAll.Visible = True
        cmcGenAll.Enabled = False   'Enable when all specifications defined
        imGenReq = True
        ilStartDGNew = False
        If imLnRowNo = 1 Then
            ilStartDGNew = True
        Else
            If tgClfCntr(imLnRowNo - 2).iCancel Then    'previous line is canceled before start
                ilStartDGNew = True
            Else
                'Temporary remove using Now when previous line is in the past
                'The three lines of code are OK- remove comment if what the
                'feature back in
            End If
        End If
        If ilStartDGNew Then
            slDate = gObtainMondayFromToday()
            smDGSave(10) = Format$(slDate, "m/d/yy")
            smDGSave(11) = gObtainNextSunday(slDate)
            smDGSave(1) = "1"
            smOrigStartDate = ""
            smOrigEndDate = ""
        Else
            smDGSave(10) = gObtainPrevMonday(smLnSave(14, imLnRowNo - 1))
            smDGSave(11) = gObtainNextSunday(smLnSave(15, imLnRowNo - 1))
            llEndDate = gDateValue(smDGSave(11))
            llStartDate = gDateValue(smDGSave(10))
            ilNoWks = (llEndDate - llStartDate + 7) \ 7
            smDGSave(1) = Trim$(str$(ilNoWks))
            smOrigStartDate = ""
            smOrigEndDate = ""
        End If
    End If
    slStr = gAddDayToDate(smDGSave(10))
    gSetShow pbcDGen, slStr, tmDGCtrls(DGSTARTDATEINDEX)
    slStr = gAddDayToDate(smDGSave(11))
    gSetShow pbcDGen, slStr, tmDGCtrls(DGENDDATEINDEX)
    slStr = smDGSave(10)
    'If (lbcBillCycle.ListIndex <> 1) And (lbcBillCycle.ListIndex <> 3) Then
    '    slStr = gObtainEndStd(slStr)
    'End If
    If (Contract.lbcBillCycle.ListIndex >= 0) And (Contract.lbcBillCycle.ListIndex <= 2) Then
        slStr = Left(Contract.lbcBillCycle.List(Contract.lbcBillCycle.ListIndex), 1)
        If Left(Contract.lbcBillCycle.List(Contract.lbcBillCycle.ListIndex), 1) = "C" Then
            slStr = gObtainEndCal(slStr)
        ElseIf Left(Contract.lbcBillCycle.List(Contract.lbcBillCycle.ListIndex), 1) = "W" Then
            slStr = gObtainNextSunday(slStr)
        Else
            slStr = gObtainEndStd(slStr)
        End If
    End If
    lacCalName.Caption = gMonthYearFormat(slStr)
    slStr = smDGSave(1)
    gSetShow pbcDGen, slStr, tmDGCtrls(DGNOWKSINDEX)
    slStr = mGetLineType(imLnRowNo)
    If (slStr <> "O") And (slStr <> "A") And (slStr <> "E") Then
        mDPAllowedDays imLnSave(1, imLnRowNo), imAllowedDays()
    Else
        For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
            If StrComp(smLnSave(20, imLnRowNo), Trim$(tgMRdf(ilLoop).sName), 1) = 0 Then
                mDPAllowedDays tgMRdf(ilLoop).iCode, imAllowedDays()
                Exit For
            End If
        Next ilLoop
    End If
    imSettingDay = True
    If (imMatchPattern = YES) And (smLnSave(14, imLnRowNo) <> "") Then
        smDGSave(2) = smAirSave(3, 1)
        slStr = smAirSave(3, 1)
        gSetShow pbcDGen, slStr, tmDGCtrls(DGSPOTWKINDEX)
        If (smAirSave(3, 1) <> "") Or (tgClfCntr(imLnRowNo - 1).iGame) Then
            For ilDay = 0 To 6 Step 1
                If (smAirSave(4 + ilDay, 1) = "Y") Or ((smAirSave(13, 1) = "Y") And (smAirSave(4 + ilDay, 2) = "Y")) Then
                    slStr = "4"
                    gSetShow pbcDGen, slStr, tmDGCtrls(DGDAILYINDEX + ilDay)
                    smDGSave(3 + ilDay) = "Y"
                Else
                    slStr = "  "
                    gSetShow pbcDGen, slStr, tmDGCtrls(DGDAILYINDEX + ilDay)
                    smDGSave(3 + ilDay) = "N"
                End If
            Next ilDay
        Else
            For ilDay = 0 To 6 Step 1
                slStr = smAirSave(4 + ilDay, 1)
                If (slStr = "") And (smAirSave(13, 1) = "Y") Then
                    slStr = smAirSave(4 + ilDay, 2)
                End If
                gSetShow pbcDGen, slStr, tmDGCtrls(DGDAILYINDEX + ilDay)
                smDGSave(3 + ilDay) = smAirSave(4 + ilDay, 1)
            Next ilDay
        End If
    Else
        smDGSave(2) = ""
        slStr = ""
        gSetShow pbcDGen, slStr, tmDGCtrls(DGSPOTWKINDEX)
        For ilDay = 0 To 6 Step 1
            slStr = ""
            gSetShow pbcDGen, slStr, tmDGCtrls(DGDAILYINDEX + ilDay)
            smDGSave(3 + ilDay) = ""
        Next ilDay
        If smLnSave(14, imLnRowNo) = "" Then
            If (imDGSave(2) = 0) Or (imDGSave(2) = 2) Then
                For ilDay = 0 To 6 Step 1
                    slStr = " R"    'Show as *
                    gSetShow pbcDGen, slStr, tmDGCtrls(DGDAILYINDEX + ilDay)
                Next ilDay
            Else
                For ilDay = 0 To 6 Step 1
                    If imOvDays(ilDay) = 1 Then
                        slStr = "  "
                        gSetShow pbcDGen, slStr, tmDGCtrls(DGDAILYINDEX + ilDay)
                        smDGSave(3 + ilDay) = "N"
                    Else
                        slStr = "4"
                        gSetShow pbcDGen, slStr, tmDGCtrls(DGDAILYINDEX + ilDay)
                        smDGSave(3 + ilDay) = "Y"
                   End If
                Next ilDay
            End If
        End If
    End If
    If tgClfCntr(imLnRowNo - 1).iGame Then
        ilAllBlank = True
        For ilDay = 0 To 6 Step 1
            If smDGSave(3 + ilDay) = "Y" Then
                ilAllBlank = False
                Exit For
            End If
        Next ilDay
        If ilAllBlank Then
            For ilDay = 0 To 6 Step 1
                If imAllowedDays(ilDay + LBound(imAllowedDays)) Then
                    smDGSave(3 + ilDay) = "Y"
                End If
            Next ilDay
        End If
    End If
    If imDGSave(2) = 0 Then
        slStr = "Daily"
    ElseIf imDGSave(2) = 1 Then
        slStr = "Weekly"
    ElseIf imDGSave(2) = 2 Then
        slStr = "Daily by %"
    ElseIf imDGSave(2) = 3 Then
        slStr = "Weekly by %"
    Else
        slStr = "Weekly"
        imDGSave(2) = 1
    End If
    gSetShow pbcDGen, slStr, tmDGCtrls(DGDWPINDEX)
    imSettingDay = False
    smOrigPattern = ""
    slStr = ""
    gSetShow pbcDGen, slStr, tmDGCtrls(DGRUNINDEX)
    For ilPattern = 0 To 3 Step 1
        ilPatternOK = True
        ilWkNo = 1
        'For ilLoop = LBound(smAirSave, 2) To UBound(smAirSave, 2) - 1 Step 1
        For ilLoop = imLB1Or2 To UBound(smAirSave, 2) - 1 Step 1
            Select Case ilPattern
                Case 0  'All week on
                    If imAirSave(1, ilLoop) <> 0 Then 'not an air week
                        ilPatternOK = False
                        Exit For
                    End If
                Case 1 To 3  'Every other week
                    If (ilWkNo Mod (ilPattern + 1)) = 1 Then    'On week
                        If imAirSave(1, ilLoop) <> 0 Then
                            ilPatternOK = False
                            Exit For
                        End If
                    Else    'Off week
                        If imAirSave(1, ilLoop) = 0 Then    'Week is on- should be off
                            ilPatternOK = False
                            Exit For
                        End If
                    End If
            End Select
            If (smAirSave(13, ilLoop) = "N") Or ((smAirSave(13, ilLoop) = "Y") And (smAirSave(13, ilLoop + 1) = "Y")) Then
                ilWkNo = ilWkNo + 1
            End If
        Next ilLoop
        imDGSave(1) = 0
        If ilPatternOK Then
            smOrigPattern = Trim$(str$(ilPattern))
            imDGSave(1) = ilPattern
            Select Case ilPattern
                Case 0
                    slStr = "Week"
                Case 1
                    slStr = "2nd"
                Case 2
                    slStr = "3rd"
                Case 3
                    slStr = "4th"
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcDGen, slStr, tmDGCtrls(DGRUNINDEX)
            Exit For
        End If
    Next ilPattern
    If (imDGSave(1) >= 0) And (imDGSave(1) <= 3) Then
        lbcDGRun.ListIndex = imDGSave(1)
    End If
    smOrigNoSpots = smDGSave(1)
    If smOrigNoSpots = "" Then
        For ilDay = 0 To 6 Step 1
            smOrigDaily(ilDay + 1) = smDGSave(3 + ilDay)
        Next ilDay
    Else
        For ilDay = 0 To 6 Step 1
            If smDGSave(3 + ilDay) = "Y" Then
                smOrigDaily(ilDay + 1) = "1"
            Else
                smOrigDaily(ilDay + 1) = "0"
            End If
        Next ilDay
    End If
    If (smLineType <> "O") And (smLineType <> "A") Then
        If imDollarMatch = YES Then
            smDGSave(12) = smAirSave(11, 1)
        Else
            smDGSave(12) = ""
        End If
        If smDGSave(12) <> "" Then
            If (Left$(smDGSave(12), 1) < "0") Or (Left$(smDGSave(12), 1) > "9") Then
                gSetShow pbcDGen, smDGSave(12), tmDGCtrls(DGPRICEINDEX)
            Else
                gFormatStr smDGSave(12), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                gSetShow pbcDGen, slStr, tmDGCtrls(DGPRICEINDEX)
            End If
        Else
            'gSetShow pbcDGen, "", tmDGCtrls(DGPRICEINDEX)
            If lbcPriceType(0).ListIndex = 1 Then
                lbcPriceType_Click 0
            Else
                lbcPriceType(0).ListIndex = 1
            End If
            smDGSave(12) = lbcPriceType(0).List(1)
            gSetShow pbcDGen, smDGSave(12), tmDGCtrls(DGPRICEINDEX)
        End If
    Else
        smDGSave(12) = ""
        gSetShow pbcDGen, smDGSave(12), tmDGCtrls(DGPRICEINDEX)
        smDGSave(12) = "0"
    End If
    If smDGSave(13) <> "" Then
        If gValidTime(smDGSave(13)) Then
            slStr = smDGSave(13)
        Else
            slStr = ""
        End If
    Else
        slStr = ""
    End If
    gSetShow pbcDGen, slStr, tmDGCtrls(DGOVSTARTINDEX)
    If smDGSave(14) <> "" Then
        If gValidTime(smDGSave(14)) Then
            slStr = smDGSave(14)
        Else
            slStr = ""
        End If
    Else
        slStr = ""
    End If
    gSetShow pbcDGen, slStr, tmDGCtrls(DGOVENDINDEX)
    mGetDemoFlightTotals
    If ilMakeShowOnly Then
        Exit Sub
    End If
    mMoveFlight
    mDGSetCommands
    pbcDGen.Cls
    pbcDGen_Paint
    pbcDGen.Visible = True
    pbcAir.Visible = True  'Set visibility
    pbcAirWk.Visible = True
    pbcAirWk.Cls
    pbcAirWk_Paint
    imAirRowNo = 1
    'plcAirScreen.Caption = "Flights"    '"Flights for Line # " & Trim$(Str$(tgClfCntr(imLnRowNo - 1).ClfRec.iLine))
    plcAirScreen.Cls
    plcAirScreen.CurrentX = 0
    plcAirScreen.CurrentY = 0
    plcAirScreen.Print "Flights"
    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        imAirBoxNo = -1
        imDGBoxNo = -1
        cmcAirCancel.SetFocus
    Else
        If tgClfCntr(imLnRowNo - 1).iFirstCff = -1 Then
            imAirBoxNo = -1
            imDGBoxNo = -1
            If pbcDGSTab.Visible Then
                pbcDGSTab.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Else
            imAirBoxNo = -1
            imDGBoxNo = -1
            cmcAirDone.SetFocus
        End If
    End If
End Sub
Private Sub lbcPriceType_Click(Index As Integer)
    If imLbcMouseDown Then
        imLbcArrowSetting = False
    Else
        imLbcArrowSetting = True
    End If
    If Index = 0 Then
        'If imLnBoxNo = PRICEINDEX Then
        '    gProcessLbcClick lbcPriceType(Index), edcLNDropDown, imChgMode, imLbcArrowSetting
        'Else
            gProcessLbcClick lbcPriceType(Index), edcDGDropDown, imChgMode, imLbcArrowSetting
        'End If
    Else
        gProcessLbcClick lbcPriceType(Index), edcAirDropDown, imChgMode, imLbcArrowSetting
    End If
    imLbcMouseDown = False
End Sub
Private Sub lbcPriceType_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveFlight                     *
'*                                                     *
'*             Created:8/09/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move flight screen             *
'*                                                     *
'*******************************************************
Private Sub mMoveFlight()
    pbcDGSTab.Left = -pbcDGSTab.Width - 60
    pbcDGTab.Left = -pbcDGTab.Width - 60
    pbcAirSTab.Left = -pbcAirSTab.Width - 60
    pbcAirTab.Left = -pbcAirTab.Width - 60
    pbcClickFocus.Left = -pbcClickFocus.Left - 60
    If (sgCntrScreen = "Orders") Or (smLineType = "O") Or (smLineType = "A") Or (smLineType = "E") Then
        ''pbcAir.Move 120, 1380, pbcAir.Width, 3675    '1350 '450, 1395
        'pbcAir.Move Contract.Width / 2 - pbcAir.Width / 2, 1380, pbcAir.Width, 3675    '1350 '450, 1395
        pbcAir.Move 30, 30, pbcAir.Width, 3675    '1350 '450, 1395
        cbcAirLine.Move pbcAir.Width - cbcAirLine.Width - 90, 90
        plcStartDate.Move 165, cbcAirLine.Top + cbcAirLine.height + 30, plcStartDate.Width, pbcDGen.height + fgPanelAdj
        cbcAirLine.Move plcStartDate.Left + plcStartDate.Width - cbcAirLine.Width
        pbcDGen.Move pbcAir.Left + plcStartDate.Left + fgBevelX + 15, pbcAir.Top + plcStartDate.Top + fgBevelY + 15
        pbcAirWk.Move pbcAir.Left + 165, pbcDGen.Top + pbcDGen.height + 2 * fgBevelX + 30
        vbcAirWk.Move pbcAirWk.Left + pbcAirWk.Width - pbcAir.Left - 15, pbcAirWk.Top - pbcAir.Top - 15, vbcAirWk.Width, pbcAirWk.height
        pbcWkArrow.Move 30 'pbcAirWk.Left - pbcWkArrow.Width - 15    'set arrow    'Vehicle
        cmcGenNew.Move vbcAirWk.Left + vbcAirWk.Width + (pbcAir.Width - vbcAirWk.Left - vbcAirWk.Width) \ 2 - cmcGenNew.Width \ 2, pbcAirWk.Top - pbcAir.Top
        cmcGenAll.Move vbcAirWk.Left + vbcAirWk.Width + (pbcAir.Width - vbcAirWk.Left - vbcAirWk.Width) \ 2 - cmcGenAll.Width \ 2, pbcAirWk.Top - pbcAir.Top
        cmcAirClear.Move vbcAirWk.Left + vbcAirWk.Width + (pbcAir.Width - vbcAirWk.Left - vbcAirWk.Width) \ 2 - cmcAirClear.Width \ 2, cmcGenNew.Top + cmcGenNew.height + 60
        cmcAirCancel.Move vbcAirWk.Left + vbcAirWk.Width + (pbcAir.Width - vbcAirWk.Left - vbcAirWk.Width) \ 2 - cmcAirCancel.Width \ 2, pbcAir.height - cmcAirCancel.height - 120 'pbcAirWk.top + pbcAirWk.Height - pbcAir.top - cmcAirCancel.Height
        cmcAirDone.Move vbcAirWk.Left + vbcAirWk.Width + (pbcAir.Width - vbcAirWk.Left - vbcAirWk.Width) \ 2 - cmcAirDone.Width \ 2, cmcAirCancel.Top - cmcAirDone.height - 60
        pbcAirDemo.Visible = False
    Else
        ''pbcAir.Move 120, 915, pbcAir.Width, 3675 + pbcAirDemo.Height - 105 '1350 '450, 1395
        'pbcAir.Move Contract.Width / 2 - pbcAir.Width / 2, 915, pbcAir.Width, 3675 + pbcAirDemo.Height - 105 '1350 '450, 1395
        pbcAir.Move 30, 30, pbcAir.Width, 3675 + pbcAirDemo.height - 105 '1350 '450, 1395
        cbcAirLine.Move pbcAir.Width - cbcAirLine.Width - 90, 90 '1710, 90
        plcStartDate.Move 165, cbcAirLine.Top + cbcAirLine.height, plcStartDate.Width, pbcDGen.height + fgPanelAdj
        cbcAirLine.Move plcStartDate.Left + plcStartDate.Width - cbcAirLine.Width, 30
        pbcDGen.Move pbcAir.Left + plcStartDate.Left + fgBevelX + 15, pbcAir.Top + plcStartDate.Top + fgBevelY + 15
        pbcAirWk.Move pbcAir.Left + 165, pbcDGen.Top + pbcDGen.height + 2 * fgBevelX - 15 '+ 30
        vbcAirWk.Move pbcAirWk.Left + pbcAirWk.Width - pbcAir.Left - 15, pbcAirWk.Top - pbcAir.Top - 15, vbcAirWk.Width, pbcAirWk.height
        pbcWkArrow.Move 30 'pbcAirWk.Left - pbcWkArrow.Width - 15    'set arrow    'Vehicle
        pbcAirDemo.Move 165 + (pbcAirWk.Width + vbcAirWk.Width - pbcAirDemo.Width) / 2, pbcAir.height - pbcAirDemo.height - 60
        cmcGenNew.Move vbcAirWk.Left + vbcAirWk.Width + (pbcAir.Width - vbcAirWk.Left - vbcAirWk.Width) \ 2 - cmcGenNew.Width \ 2, pbcAirWk.Top - pbcAir.Top
        cmcGenAll.Move vbcAirWk.Left + vbcAirWk.Width + (pbcAir.Width - vbcAirWk.Left - vbcAirWk.Width) \ 2 - cmcGenAll.Width \ 2, pbcAirWk.Top - pbcAir.Top
        cmcAirClear.Move vbcAirWk.Left + vbcAirWk.Width + (pbcAir.Width - vbcAirWk.Left - vbcAirWk.Width) \ 2 - cmcAirClear.Width \ 2, cmcGenNew.Top + cmcGenNew.height + 60
        cmcAirCancel.Move vbcAirWk.Left + vbcAirWk.Width + (pbcAir.Width - vbcAirWk.Left - vbcAirWk.Width) \ 2 - cmcAirCancel.Width \ 2, pbcAir.height - cmcAirCancel.height - 120 'pbcAirWk.top + pbcAirWk.Height - pbcAir.top - cmcAirCancel.Height
        cmcAirDone.Move vbcAirWk.Left + vbcAirWk.Width + (pbcAir.Width - vbcAirWk.Left - vbcAirWk.Width) \ 2 - cmcAirDone.Width \ 2, cmcAirCancel.Top - cmcAirDone.height - 60
        pbcAirDemo.Visible = True
    End If
    CFlight.height = pbcAir.height + 180
End Sub
Private Sub pbcDGen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    
    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        cmcAirCancel.SetFocus
        Exit Sub
    End If
    'For ilBox = LBound(tmDGCtrls) To UBound(tmDGCtrls) Step 1
    For ilBox = imLBCTCtrls To UBound(tmDGCtrls) Step 1
        If (X >= tmDGCtrls(ilBox).fBoxX) And (X <= (tmDGCtrls(ilBox).fBoxX + tmDGCtrls(ilBox).fBoxW)) Then
            If (Y >= tmDGCtrls(ilBox).fBoxY) And (Y <= (tmDGCtrls(ilBox).fBoxY + tmDGCtrls(ilBox).fBoxH)) Then
                mDGSetShow imDGBoxNo
                'If ((tgClfCntr(imLnRowNo - 1).iGame) Or (imDGSave(2) = 0) Or (imDGSave(2) = 2)) And (ilBox = DGSPOTWKINDEX) Then
                If ((imDGSave(2) = 0) Or (imDGSave(2) = 2)) And (ilBox = DGSPOTWKINDEX) Then
                    Beep
                    MsgBox "Can't defined Spots/Week as Line Defined as Daily", vbOKOnly + vbExclamation, "Warning"
                    If imDGBoxNo >= 0 Then
                        mDGSetFocus imDGBoxNo
                    Else
                        mAirSetFocus imAirBoxNo
                    End If
                    Exit Sub
                End If
                If mIsLnGame(imLnRowNo) Then
                    Beep
                    MsgBox "Event Line, click on weeks in Line area to define Spots per Event", vbOKOnly + vbExclamation, "Warning"
                    If imDGBoxNo >= 0 Then
                        mDGSetFocus imDGBoxNo
                    Else
                        mAirSetFocus imAirBoxNo
                    End If
                    Exit Sub
                End If
                If (Not tgClfCntr(imLnRowNo - 1).iOverride) And ((ilBox = DGOVSTARTINDEX) Or (ilBox = DGOVENDINDEX)) Then
                    Beep
                    MsgBox "Not allowed to Override Daypart times", vbOKOnly + vbExclamation, "Warning"
                    If imDGBoxNo >= 0 Then
                        mDGSetFocus imDGBoxNo
                    Else
                        mAirSetFocus imAirBoxNo
                    End If
                    Exit Sub
                End If
                If (ilBox >= DGDAILYINDEX) And (imDGSave(2) < 0) And (Not (tgClfCntr(imLnRowNo - 1).iGame)) Then
                    Beep
                    ilBox = DGDWPINDEX
                End If
                If (ilBox = DGPRICEINDEX) And ((smLineType = "O") Or (smLineType = "A")) And (tgSpf.sCAudPkg <> "Y") Then
                    Beep
                    MsgBox "Spot Prices on Package Lines are Calculated Automatically", vbOKOnly + vbExclamation, "Warning"
                    If imDGBoxNo > 0 Then
                        mDGSetFocus imDGBoxNo
                    Else
                        mAirEnableBox imAirBoxNo
                    End If
                    Exit Sub
                End If
                'If (ilBox = DGPRICEINDEX) And (tgUrf(0).sChgPrices = "V") Then
                If (ilBox = DGPRICEINDEX) And (Not Contract.mAllowedToChgSpotPrice(imLnRowNo)) Then
                    Beep
                    'MsgBox "Not Allowed to Change Spot Price", vbOkOnly + vbExclamation, "Warning"
                    If imDGBoxNo > 0 Then
                        mDGSetFocus imDGBoxNo
                    Else
                        mAirEnableBox imAirBoxNo
                    End If
                    Exit Sub
                End If
                If (ilBox = DGPRICEINDEX) And (smLineType = "H") Then
                    slStr = mGetPkLineType(imLnRowNo)
                    If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                        Beep
                        MsgBox "Spot Prices on Hidden Lines are Calculated Automatically", vbOKOnly + vbExclamation, "Warning"
                        If imDGBoxNo > 0 Then
                            mDGSetFocus imDGBoxNo
                        Else
                            mAirEnableBox imAirBoxNo
                        End If
                        Exit Sub
                    End If
                End If
                If (smLineType = "H") Then
                    For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                        If tgPackLn(ilLoop).iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine Then
                            'If tgPackLn(ilLoop).iAlter <> True Then
                            If tgPackLn(ilLoop).sAlter <> "Y" Then
                                Beep
                                MsgBox "Can't Alter Hidden Line", vbOKOnly + vbExclamation, "Warning"
                                If imDGBoxNo > 0 Then
                                    mDGSetFocus imDGBoxNo
                                Else
                                    mAirEnableBox imAirBoxNo
                                End If
                                Exit Sub
                            End If
                        End If
                    Next ilLoop
                End If
                mAirSetShow imAirBoxNo
                imAirBoxNo = -1
                imDGBoxNo = ilBox
                mDGEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    If imDGBoxNo >= 0 Then
        mDGSetFocus imDGBoxNo
    Else
        mAirSetFocus imAirBoxNo
    End If
End Sub
Private Sub pbcDGen_Paint()
    Dim ilBox As Integer
    Dim slFont As String
    Dim llColor As Long
    llColor = 0
    'For ilBox = LBound(tmDGCtrls) To UBound(tmDGCtrls) Step 1
    For ilBox = imLBCTCtrls To UBound(tmDGCtrls) Step 1
        gPaintArea pbcDGen, tmDGCtrls(ilBox).fBoxX, tmDGCtrls(ilBox).fBoxY, tmDGCtrls(ilBox).fBoxW - 15, tmDGCtrls(ilBox).fBoxH - 15, WHITE
        pbcDGen.CurrentX = tmDGCtrls(ilBox).fBoxX + fgBoxInsetX
        If (Not tgClfCntr(imLnRowNo - 1).iGame) Or (ilBox <> DGDWPINDEX) Then
            If (ilBox >= DGDAILYINDEX) And (ilBox <= DGDAILYINDEX + 6) And (Trim$(tmDGCtrls(ilBox).sShow) = "R") Then
                pbcDGen.CurrentY = tmDGCtrls(ilBox).fBoxY + 15 '+ fgBoxInsetY
                slFont = pbcDGen.FontName
                llColor = pbcDGen.ForeColor
                pbcDGen.FontBold = False
                pbcDGen.FontName = "Monotype Sorts"
                pbcDGen.ForeColor = BLUE
                pbcDGen.Print tmDGCtrls(ilBox).sShow
                pbcDGen.FontName = slFont
                pbcDGen.ForeColor = llColor
                pbcDGen.FontBold = True
            Else
                If (ilBox >= DGDAILYINDEX) And (ilBox <= DGDAILYINDEX + 6) And ((imDGSave(2) = 1) Or (imDGSave(2) = 3)) Then
                    pbcDGen.CurrentY = tmDGCtrls(ilBox).fBoxY + 15 '+ fgBoxInsetY
                    slFont = pbcDGen.FontName
                    pbcDGen.FontName = "Monotype Sorts"
                    pbcDGen.FontBold = False
                    pbcDGen.Print tmDGCtrls(ilBox).sShow
                    pbcDGen.FontName = slFont
                    pbcDGen.FontBold = True
                Else
                    If (ilBox >= DGDAILYINDEX) And (ilBox <= DGDAILYINDEX + 6) And (tgClfCntr(imLnRowNo - 1).iGame) Then
                        pbcDGen.CurrentY = tmDGCtrls(ilBox).fBoxY + 15 '+ fgBoxInsetY
                        slFont = pbcDGen.FontName
                        pbcDGen.FontName = "Monotype Sorts"
                        pbcDGen.FontBold = False
                        pbcDGen.Print tmDGCtrls(ilBox).sShow
                        pbcDGen.FontName = slFont
                        pbcDGen.FontBold = True
                    Else
                        pbcDGen.CurrentY = tmDGCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
                        pbcDGen.Print tmDGCtrls(ilBox).sShow
                    End If
                End If
            End If
        End If
    Next ilBox
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetDate                        *
'*                                                     *
'*             Created:8/03/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move values from record to     *
'*                      working area                   *
'*                                                     *
'*******************************************************
Public Sub mGetDate(ilRetainZeroWks As Integer)
    Dim ilCff As Integer
    Dim ilPCff As Integer
    Dim ilPrev As Integer
    Dim ilUpperLimit As Integer
    Dim ilWkNo As Integer
    'Dim slStartDate As String
    Dim llStartDate As Long
    Dim llSplitStart As Long
    Dim llSplitEnd As Long
    'Dim slEndDate As String
    Dim llEndDate As Long
    Dim llDate As Long
    Dim llPrevEndDate As Long
    Dim slStr As String
    Dim ilDay As Integer
    Dim ilMonth As Integer
    Dim ilLoop As Integer
    Dim llDateAdj As Long
    Dim ilNoSpots As Integer
    Dim ilNoAirDays As Integer
    Dim ilNoDays As Integer
    Dim ilStartDay As Integer
    Dim ilEndDay As Integer
    Dim ilNoCff As Integer
    Dim ilDyWk As Integer
    Dim llWkPrice As Long
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilRetainWkWithoutSpots As Integer
    Dim ilRetainWk As Integer
    Dim ilPrevCff As Integer
    Dim slLastPriceType As String
    Dim slLastPrice As String
    Dim slLineType As String
    Dim llSpots As Long
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim ilDays(0 To 6) As Integer
    ReDim smAirShow(0 To 14, 0 To 1) As String
    ReDim smAirSave(0 To 13, 0 To 1) As String
    ReDim imAirSave(0 To 2, 0 To 1) As Integer
    
    If bmInCancel Then
        Exit Sub
    End If
    If imLnRowNo <= 0 Then
        Exit Sub
    End If
    slLineType = mGetLineType(imLnRowNo)
    imDGSave(2) = imLnSave(8, imLnRowNo)    'check each week setting ilDyWk
    ilDyWk = -1
    imDGSave(7) = imLnSave(31, imLnRowNo)
    imDGSave(3) = imLnSave(6, imLnRowNo)
    If imLnSave(4, imLnRowNo) = 1 Then  'Time override
        smDGSave(13) = smLnSave(2, imLnRowNo)
        smDGSave(14) = smLnSave(3, imLnRowNo)
        llOvStartTime = gTimeToLong(smDGSave(13), False)
        llOvEndTime = gTimeToLong(smDGSave(14), True)
    Else
        smDGSave(13) = ""
        smDGSave(14) = ""
        llOvStartTime = 0
        llOvEndTime = 0
    End If
    If Not ilRetainZeroWks Then
        ilRetainWkWithoutSpots = True
        ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
        Do While ilCff <> -1
            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                    For ilLoop = 0 To 6 Step 1
                        If tgCffCntr(ilCff).CffRec.iDay(ilLoop) > 0 Then
                            ilRetainWkWithoutSpots = False
                            Exit Do
                        End If
                    Next ilLoop
                Else
                    If (tgCffCntr(ilCff).CffRec.iSpotsWk > 0) Or (tgCffCntr(ilCff).CffRec.iXSpotsWk > 0) Then
                        ilRetainWkWithoutSpots = False
                        Exit Do
                    End If
                End If

            End If
            ilCff = tgCffCntr(ilCff).iNextCff
        Loop
        If Not ilRetainWkWithoutSpots Then
            'ilPrevCff = -1
            ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
            ilPrevCff = ilCff
            Do While ilCff <> -1
                ilRetainWk = False
                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                    If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                        For ilLoop = 0 To 6 Step 1
                            If tgCffCntr(ilCff).CffRec.iDay(ilLoop) > 0 Then
                                ilRetainWk = True
                                Exit For
                            End If
                        Next ilLoop
                    Else
                        If (tgCffCntr(ilCff).CffRec.iSpotsWk > 0) Or (tgCffCntr(ilCff).CffRec.iXSpotsWk > 0) Then
                            ilRetainWk = True
                        End If
                    End If
                End If
                If ilRetainWk Then
                    ilPrevCff = ilCff
                    ilCff = tgCffCntr(ilCff).iNextCff
                Else
                    If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = -1) Then
                        'If ilPrevCff = tgClfCntr(imLnRowNo - 1).iFirstCff Then
                        If ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff Then
                            tgClfCntr(imLnRowNo - 1).iFirstCff = tgCffCntr(ilCff).iNextCff
                            ilPrevCff = tgCffCntr(ilCff).iNextCff
                        Else
                            tgCffCntr(ilPrevCff).iNextCff = tgCffCntr(ilCff).iNextCff
                        End If
                        ilPrev = ilCff
                        tgCffCntr(ilCff).iStatus = -1   'New to not used
                        ilCff = tgCffCntr(ilCff).iNextCff
                        tgCffCntr(ilPrev).iNextCff = -1
                    Else
                        ilPrevCff = ilCff
                        tgCffCntr(ilCff).iStatus = 2    'Old to delete
                        ilCff = tgCffCntr(ilCff).iNextCff
                    End If

                End If
            Loop
        End If
    End If
    imMatchPattern = YES
    imDollarMatch = YES
    imAllWeeks = True
    ilUpperLimit = UBound(smAirShow, 2)
    ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
    ilWkNo = 1
    slLastPriceType = ""
    slLastPrice = ""
    Do While ilCff <> -1
        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
            'gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate    'Week Start date
            'gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate    'Week Start date
            llStartDate = tgCffCntr(ilCff).lStartDate   'gDateValue(slStartDate)
            llEndDate = tgCffCntr(ilCff).lEndDate   'gDateValue(slEndDate)
            If ilWkNo > 1 Then
                Do While gWeekDayLong(llPrevEndDate) <> 6   '6=sunday
                    llPrevEndDate = llPrevEndDate + 1
                Loop
                'Generate Non-Air Weeks
                For llDate = llPrevEndDate + 7 To llStartDate Step 7    'Show non-air weeks
                    imAirSave(1, ilUpperLimit) = 1
                    slStr = "No"
                    gSetShow pbcAirWk, slStr, tmAirCtrls(AIRWEEKINDEX)
                    smAirShow(AIRWEEKINDEX, ilUpperLimit) = tmAirCtrls(AIRWEEKINDEX).sShow
                    slStr = Trim$(str$(ilWkNo))
                    gSetShow pbcAirWk, slStr, tmAirCtrls(WEEKNOINDEX)
                    smAirShow(WEEKNOINDEX, ilUpperLimit) = tmAirCtrls(WEEKNOINDEX).sShow
                    slStr = Format$(llDate - 6, "m/d/yy")
                    smAirSave(1, ilUpperLimit) = slStr
                    slStr = Format$(llDate, "m/d/yy")
                    smAirSave(2, ilUpperLimit) = slStr
                    slStr = ""
                    smAirSave(3, ilUpperLimit) = slStr
                    gSetShow pbcAirWk, slStr, tmAirCtrls(SPOTPERWKINDEX)
                    smAirShow(SPOTPERWKINDEX, ilUpperLimit) = tmAirCtrls(SPOTPERWKINDEX).sShow
                    For ilDay = 0 To 6 Step 1
                        slStr = ""
                        smAirSave(4 + ilDay, ilUpperLimit) = slStr
                        gSetShow pbcAirWk, slStr, tmAirCtrls(DAILYSPOTSINDEX + ilDay)
                        smAirShow(DAILYSPOTSINDEX + ilDay, ilUpperLimit) = tmAirCtrls(DAILYSPOTSINDEX + ilDay).sShow
                    Next ilDay
                    smAirSave(13, ilUpperLimit) = "N"
                    ilWkNo = ilWkNo + 1
                    ilUpperLimit = ilUpperLimit + 1
                    ReDim Preserve smAirShow(0 To 14, 0 To ilUpperLimit) As String
                    ReDim Preserve smAirSave(0 To 13, 0 To ilUpperLimit) As String
                    ReDim Preserve imAirSave(0 To 2, 0 To ilUpperLimit) As Integer
                Next llDate
            End If
            For llDate = llStartDate To llEndDate Step 7
                imAirSave(1, ilUpperLimit) = 0
                slStr = "Yes"
                gSetShow pbcAirWk, slStr, tmAirCtrls(AIRWEEKINDEX)
                smAirShow(AIRWEEKINDEX, ilUpperLimit) = tmAirCtrls(AIRWEEKINDEX).sShow
                slStr = Trim$(str$(ilWkNo))
                gSetShow pbcAirWk, slStr, tmAirCtrls(WEEKNOINDEX)
                smAirShow(WEEKNOINDEX, ilUpperLimit) = tmAirCtrls(WEEKNOINDEX).sShow
                slStr = Format$(llDate, "m/d/yy")
                smAirSave(1, ilUpperLimit) = slStr
                ilStartDay = gWeekDayStr(smAirSave(1, ilUpperLimit))
                If llDate + 6 - ilStartDay <= llEndDate Then
                    slStr = Format$(llDate + 6 - ilStartDay, "m/d/yy")
                Else
                    slStr = Format$(llEndDate, "m/d/yy")
                End If
                smAirSave(2, ilUpperLimit) = slStr
                ilEndDay = gWeekDayStr(smAirSave(2, ilUpperLimit))
                'Adjust start and end days to bracket defined days
                For ilDay = ilStartDay To ilEndDay Step 1
                    If tgCffCntr(ilCff).CffRec.iDay(ilDay) <> 0 Then
                        Exit For
                    End If
                    ilStartDay = ilStartDay + 1
                Next ilDay
                For ilDay = ilEndDay To ilStartDay Step -1
                    If tgCffCntr(ilCff).CffRec.iDay(ilDay) <> 0 Then
                        Exit For
                    End If
                    ilEndDay = ilEndDay - 1
                Next ilDay
                For ilDay = 0 To 6 Step 1   'Initialize days
                    slStr = ""
                    smAirSave(4 + ilDay, ilUpperLimit) = slStr
                    gSetShow pbcAirWk, slStr, tmAirCtrls(DAILYSPOTSINDEX + ilDay)
                    smAirShow(DAILYSPOTSINDEX + ilDay, ilUpperLimit) = tmAirCtrls(DAILYSPOTSINDEX + ilDay).sShow
                Next ilDay
                If (tgCffCntr(ilCff).CffRec.iSpotsWk > 0) Or (tgCffCntr(ilCff).CffRec.iXSpotsWk > 0) Or (tgCffCntr(ilCff).CffRec.sDyWk = "W") Or (tgClfCntr(imLnRowNo - 1).iGame) Then
                    If tgClfCntr(imLnRowNo - 1).iGame Then
                        slStr = ""
                    Else
                        slStr = Trim$(str$(tgCffCntr(ilCff).CffRec.iSpotsWk))
                        imAirSave(2, ilUpperLimit) = 1
                    End If
                    smAirSave(3, ilUpperLimit) = slStr
                    gSetShow pbcAirWk, slStr, tmAirCtrls(SPOTPERWKINDEX)
                    smAirShow(SPOTPERWKINDEX, ilUpperLimit) = tmAirCtrls(SPOTPERWKINDEX).sShow
                    For ilDay = ilStartDay To ilEndDay Step 1
                        If tgCffCntr(ilCff).CffRec.iDay(ilDay) <> 0 Then
                            smAirSave(4 + ilDay, ilUpperLimit) = "Y"
                            slStr = "4"
                            gSetShow pbcAirWk, slStr, tmAirCtrls(DAILYSPOTSINDEX + ilDay)
                            smAirShow(DAILYSPOTSINDEX + ilDay, ilUpperLimit) = tmAirCtrls(DAILYSPOTSINDEX + ilDay).sShow
                        Else
                            smAirSave(4 + ilDay, ilUpperLimit) = "N"
                            slStr = "  "
                            gSetShow pbcAirWk, slStr, tmAirCtrls(DAILYSPOTSINDEX + ilDay)
                            smAirShow(DAILYSPOTSINDEX + ilDay, ilUpperLimit) = tmAirCtrls(DAILYSPOTSINDEX + ilDay).sShow
                        End If
                    Next ilDay
                Else
                    slStr = ""
                    smAirSave(3, ilUpperLimit) = slStr
                    imAirSave(2, ilUpperLimit) = 0
                    gSetShow pbcAirWk, slStr, tmAirCtrls(SPOTPERWKINDEX)
                    smAirShow(SPOTPERWKINDEX, ilUpperLimit) = tmAirCtrls(SPOTPERWKINDEX).sShow
                    For ilDay = ilStartDay To ilEndDay Step 1
                        slStr = Trim$(str$(tgCffCntr(ilCff).CffRec.iDay(ilDay)))
                        smAirSave(4 + ilDay, ilUpperLimit) = slStr
                        gSetShow pbcAirWk, slStr, tmAirCtrls(DAILYSPOTSINDEX + ilDay)
                        smAirShow(DAILYSPOTSINDEX + ilDay, ilUpperLimit) = tmAirCtrls(DAILYSPOTSINDEX + ilDay).sShow
                    Next ilDay
                End If
                smAirSave(13, ilUpperLimit) = "N"
                'Test if same pattern is in all week- used later to set pattern on flight screen
                If (ilUpperLimit > 1) And (imMatchPattern = YES) Then
                    If smAirSave(3, 1) <> smAirSave(3, ilUpperLimit) Then
                        imMatchPattern = NO
                    Else
                        For ilDay = 0 To 6 Step 1
                            If smAirSave(4 + ilDay, 1) <> smAirSave(4 + ilDay, ilUpperLimit) Then
                                imMatchPattern = NO
                                Exit For
                            End If
                        Next ilDay
                    End If
                End If
                If ilDyWk = -1 Then
                    If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                        ilDyWk = 1
                    Else
                        ilDyWk = 2
                    End If
                Else
                    If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                        If ilDyWk <> 1 Then
                            ilDyWk = 3
                        End If
                    Else
                        If ilDyWk <> 2 Then
                            ilDyWk = 3
                        End If
                    End If
                End If
                'Get Price
                'Get proposal price here incase of * as price as smAirSave(12,..) will be set
                smAirSave(12, ilUpperLimit) = gLongToStrDec(tgCffCntr(ilCff).CffRec.lPropPrice, 0)  '2)
                Select Case tgCffCntr(ilCff).CffRec.sPriceType
                    Case "T"    'True
                        'gPDNToStr tgCffCntr(ilCff).CffRec.sActPrice, 2, smLnSave(17, ilLoop + 1)
                        'smAirSave(11, ilUpperLimit) = gLongToStrDec(tgCffCntr(ilCff).CffRec.lActPrice, 2)
                        If slLineType <> "E" Then
                            smAirSave(11, ilUpperLimit) = gLongToStrDec(tgCffCntr(ilCff).CffRec.lActPrice, 2)
                        Else
                            smAirSave(11, ilUpperLimit) = gLongToStrDec(tgCffCntr(ilCff).CffRec.lActPrice, 2)
                        End If
                        slLastPriceType = tgCffCntr(ilCff).CffRec.sPriceType
                        slLastPrice = smAirSave(11, ilUpperLimit)
                    Case "N"    'No Charge
                        smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(1)
                        slLastPriceType = tgCffCntr(ilCff).CffRec.sPriceType
                        slLastPrice = smAirSave(11, ilUpperLimit)
                    Case "M"    'MG Line
                        smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(2)
                        slLastPriceType = tgCffCntr(ilCff).CffRec.sPriceType
                        slLastPrice = smAirSave(11, ilUpperLimit)
                    Case "B"    'Bonus
                        smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(3)
                        slLastPriceType = tgCffCntr(ilCff).CffRec.sPriceType
                        slLastPrice = smAirSave(11, ilUpperLimit)
                    Case "S"    'Spinoff
                        smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(4)
                        slLastPriceType = tgCffCntr(ilCff).CffRec.sPriceType
                        slLastPrice = smAirSave(11, ilUpperLimit)
                    Case "P"    'Package
                        smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(5)
                        slLastPriceType = tgCffCntr(ilCff).CffRec.sPriceType
                        slLastPrice = smAirSave(11, ilUpperLimit)
                    Case "R"    'Recapturable
                        smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(6)
                        slLastPriceType = tgCffCntr(ilCff).CffRec.sPriceType
                        slLastPrice = smAirSave(11, ilUpperLimit)
                    Case "A"    'ADU
                        smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(7)
                        slLastPriceType = tgCffCntr(ilCff).CffRec.sPriceType
                        slLastPrice = smAirSave(11, ilUpperLimit)
                    Case "*"    'Compute price from rate card- set in duplicate flights (CDupl) and mAddWeek
                        'tgCffCntr(ilCff).CffRec.sPriceType = "T"   'Set at end of flight area
                        If (sgCntrScreen = "Orders") Then
                            'Find line with a price and use it
                            If slLastPriceType <> "" Then
                                smAirSave(11, ilUpperLimit) = slLastPrice
                            Else
                                'Price not found in past weeks, look into future weeks
                                smAirSave(11, ilUpperLimit) = ""
                                ilPCff = tgCffCntr(ilCff).iNextCff
                                Do While ilPCff <> -1
                                    Select Case tgCffCntr(ilPCff).CffRec.sPriceType
                                        Case "T"    'True
                                            'gPDNToStr tgCffCntr(ilCff).CffRec.sActPrice, 2, smLnSave(17, ilLoop + 1)
                                            If slLineType <> "E" Then
                                                smAirSave(11, ilUpperLimit) = gLongToStrDec(tgCffCntr(ilPCff).CffRec.lActPrice, 2)
                                            Else
                                                smAirSave(11, ilUpperLimit) = gLongToStrDec(tgCffCntr(ilPCff).CffRec.lActPrice, 2)
                                            End If
                                            slLastPriceType = tgCffCntr(ilPCff).CffRec.sPriceType
                                            slLastPrice = smAirSave(11, ilUpperLimit)
                                            Exit Do
                                        Case "N"    'No Charge
                                            smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(1)
                                            slLastPriceType = tgCffCntr(ilPCff).CffRec.sPriceType
                                            slLastPrice = smAirSave(11, ilUpperLimit)
                                            Exit Do
                                        Case "M"    'MG Line
                                            smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(2)
                                            slLastPriceType = tgCffCntr(ilPCff).CffRec.sPriceType
                                            slLastPrice = smAirSave(11, ilUpperLimit)
                                            Exit Do
                                        Case "B"    'Bonus
                                            smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(3)
                                            slLastPriceType = tgCffCntr(ilPCff).CffRec.sPriceType
                                            slLastPrice = smAirSave(11, ilUpperLimit)
                                            Exit Do
                                        Case "S"    'Spinoff
                                            smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(4)
                                            slLastPriceType = tgCffCntr(ilPCff).CffRec.sPriceType
                                            slLastPrice = smAirSave(11, ilUpperLimit)
                                            Exit Do
                                        Case "P"    'Package
                                            smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(5)
                                            slLastPriceType = tgCffCntr(ilPCff).CffRec.sPriceType
                                            slLastPrice = smAirSave(11, ilUpperLimit)
                                            Exit Do
                                        Case "R"    'Recapturable
                                            smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(6)
                                            slLastPriceType = tgCffCntr(ilPCff).CffRec.sPriceType
                                            slLastPrice = smAirSave(11, ilUpperLimit)
                                            Exit Do
                                        Case "A"    'ADU
                                            smAirSave(11, ilUpperLimit) = lbcPriceType(1).List(7)
                                            slLastPriceType = tgCffCntr(ilPCff).CffRec.sPriceType
                                            slLastPrice = smAirSave(11, ilUpperLimit)
                                            Exit Do
                                    End Select
                                    ilPCff = tgCffCntr(ilPCff).iNextCff
                                Loop
                            End If
                            If slLastPriceType = "" Then
                                smAirSave(11, ilUpperLimit) = ".00"
                            End If
                            smAirSave(12, ilUpperLimit) = ""
                        Else
                            mCVehOptIndex imLnRowNo, ilVefCode, ilVpfIndex, imTerminate
                            llWkPrice = mComputeRCWkPrice(tmRcf.iCode, imLnSave(1, imLnRowNo), smAirSave(1, ilUpperLimit), ilVefCode, imLnRowNo)
                            If llWkPrice > 0 Then
                                If imAirSave(2, ilUpperLimit) = 0 Then
                                    llSpots = 0
                                    For ilDay = 4 To 10 Step 1
                                        llSpots = llSpots + Val(smAirSave(ilDay, ilUpperLimit))
                                        If Val(smAirSave(ilDay, ilUpperLimit)) > 0 Then
                                            ilDays(ilDay - 4) = True
                                        Else
                                            ilDays(ilDay - 4) = False
                                        End If
                                    Next ilDay
                                Else
                                    llSpots = Val(smAirSave(3, ilUpperLimit))
                                    For ilDay = 4 To 10 Step 1
                                        If smAirSave(ilDay, ilUpperLimit) = "Y" Then
                                            ilDays(ilDay - 4) = True
                                        Else
                                            ilDays(ilDay - 4) = False
                                        End If
                                    Next ilDay
                                End If
                                If smDGSave(13) <> "" Then
                                    llOvStartTime = gTimeToLong(smDGSave(13), False)
                                    llOvEndTime = gTimeToLong(smDGSave(14), True)
                                Else
                                    llOvStartTime = 0
                                    llOvEndTime = 0
                                End If
                                slStr = mRCPriceAdjust(tmRcf, llWkPrice, True, imLnRowNo, imDGSave(), llOvStartTime, llOvEndTime, Val(smLnSave(16, imLnRowNo)), llSpots, ilDays(), imLnSave(1, imLnRowNo))
                                smAirSave(11, ilUpperLimit) = gRoundStr(slStr, "1.00", 0)    'gLongToStrDec(llWkPrice, 0) '& ".00"
                            Else
                                smAirSave(11, ilUpperLimit) = ""
                            End If
                            smAirSave(12, ilUpperLimit) = gRoundStr(smAirSave(11, ilUpperLimit), "1.00", 0) 'smAirSave(11, ilUpperLimit)
                        End If
                    Case Else
                        smAirSave(11, ilUpperLimit) = ""
                End Select
                If smAirSave(11, ilUpperLimit) <> "" Then
                    If (Left$(smAirSave(11, ilUpperLimit), 1) < "0") Or (Left$(smAirSave(11, ilUpperLimit), 1) > "9") Then
                        gSetShow pbcAirWk, smAirSave(11, ilUpperLimit), tmAirCtrls(WKPRICEINDEX)
                        smAirShow(WKPRICEINDEX, ilUpperLimit) = tmAirCtrls(WKPRICEINDEX).sShow
                    Else
                        gFormatStr smAirSave(11, ilUpperLimit), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                        gSetShow pbcAirWk, slStr, tmAirCtrls(WKPRICEINDEX)
                        smAirShow(WKPRICEINDEX, ilUpperLimit) = tmAirCtrls(WKPRICEINDEX).sShow
                    End If
                Else
                    gSetShow pbcAirWk, "", tmAirCtrls(WKPRICEINDEX)
                    smAirShow(WKPRICEINDEX, ilUpperLimit) = tmAirCtrls(WKPRICEINDEX).sShow
                End If
                If (ilUpperLimit > 1) And (imDollarMatch = YES) Then
                    If smAirSave(11, 1) <> smAirSave(11, ilUpperLimit) Then
                        imDollarMatch = False
                    End If
                End If

                ilWkNo = ilWkNo + 1
                ilUpperLimit = ilUpperLimit + 1
                ReDim Preserve smAirShow(0 To 14, 0 To ilUpperLimit) As String
                ReDim Preserve smAirSave(0 To 13, 0 To ilUpperLimit) As String
                ReDim Preserve imAirSave(0 To 2, 0 To ilUpperLimit) As Integer
            Next llDate
            llPrevEndDate = llEndDate
        End If
        If tgCffCntr(ilCff).CffRec.sPriceType = "*" Then
            tgCffCntr(ilCff).CffRec.sPriceType = "T"
        End If
        ilCff = tgCffCntr(ilCff).iNextCff
    Loop
    ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
    ilNoCff = 0
    Do While ilCff <> -1
        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
            ilNoCff = ilNoCff + 1
            If ilNoCff = 1 Then
                Select Case tgCffCntr(ilCff).CffRec.sPriceType
                    Case "T"
                        imLnSave(7, imLnRowNo) = 0
                        smLnSave(17, imLnRowNo) = gLongToStrDec(tgCffCntr(ilCff).CffRec.lActPrice, 2)
                    Case "N"
                        imLnSave(7, imLnRowNo) = 1
                        smLnSave(17, imLnRowNo) = "N/C"
                    Case "M"
                        imLnSave(7, imLnRowNo) = 2
                        smLnSave(17, imLnRowNo) = "MG"
                    Case "B"
                        imLnSave(7, imLnRowNo) = 3
                        smLnSave(17, imLnRowNo) = "Bonus"
                    Case "S"
                        imLnSave(7, imLnRowNo) = 4
                        smLnSave(17, imLnRowNo) = "Spinoff"
                    Case "P"
                        imLnSave(7, imLnRowNo) = 5
                        smLnSave(17, imLnRowNo) = "Package"
                    Case "R"
                        imLnSave(7, imLnRowNo) = 6
                        smLnSave(17, imLnRowNo) = "Recapturable"
                    Case "A"
                        imLnSave(7, imLnRowNo) = 7
                        smLnSave(17, imLnRowNo) = "ADU"
                End Select
            End If
        End If
        ilCff = tgCffCntr(ilCff).iNextCff
    Loop
    If ilNoCff <= 1 Then
        imLnSave(9, imLnRowNo) = True
    Else
        imLnSave(9, imLnRowNo) = False
    End If
    If ilDyWk = 1 Then
        imDGSave(2) = 0
    ElseIf ilDyWk = 2 Then
        imDGSave(2) = 1
    End If
    'Determine if spots on first week must be split with last week
    'Adjust dates
    'For ilLoop = LBound(smAirSave, 2) To UBound(smAirSave, 2) - 1 Step 1
    For ilLoop = imLB1Or2 To UBound(smAirSave, 2) - 1 Step 1
        mAdjDates ilLoop
        If imAirSave(1, ilLoop) = 0 Then    'Air Week
            'If smAirSave(3, ilLoop) = "" Then
            If imAirSave(2, ilLoop) = 0 Then
                'imAirSave(2, ilLoop) = 0
                If imDGSave(2) > 1 Then
                    slStr = "Daily by %"
                Else
                    slStr = "Daily"
                End If
            Else
                'imAirSave(2, ilLoop) = 1
                If imDGSave(2) > 1 Then
                    slStr = "Weekly by %"
                Else
                    slStr = "Weekly"
                End If
            End If
        Else
            'Note: First and last week must be airing weeks- therefore set Mon-Sun
            'Non-Air Week: Initialize values
            imAirSave(2, ilLoop) = imDGSave(2)
            If (imDGSave(2) = 0) Or (imDGSave(2) = 2) Then
                smAirSave(3, ilLoop) = ""
                For ilDay = 0 To 6 Step 1
                    smAirSave(4 + ilDay, ilLoop) = smDGSave(3 + ilDay)
                Next ilDay
            ElseIf (imDGSave(2) = 1) Or (imDGSave(2) = 3) Then
                smAirSave(3, ilLoop) = smDGSave(2)
                For ilDay = 0 To 6 Step 1
                    If smDGSave(3 + ilDay) = "Y" Then
                        smAirSave(4 + ilDay, ilLoop) = "Y"
                    Else
                        smAirSave(4 + ilDay, ilLoop) = "N"
                    End If
                Next ilDay
            Else
                smAirSave(3, ilLoop) = ""
            End If
            slStr = ""
        End If
        gSetShow pbcAirWk, slStr, tmAirCtrls(DWINDEX)
        smAirShow(DWINDEX, ilLoop) = tmAirCtrls(DWINDEX).sShow
    Next ilLoop
    mFlightTotals
    vbcAirWk.Min = imLB1Or2 'LBound(smAirShow, 2)
    If UBound(smAirShow, 2) - 2 <= vbcAirWk.LargeChange Then
        vbcAirWk.Max = imLB1Or2 'LBound(smAirShow, 2)
    Else
        vbcAirWk.Max = UBound(smAirShow, 2) - 1 - vbcAirWk.LargeChange
    End If
    vbcAirWk.Value = vbcAirWk.Min
    If ilUpperLimit > 1 Then
        smLnSave(14, imLnRowNo) = smAirSave(1, 1)
        smLnSave(15, imLnRowNo) = smAirSave(2, ilUpperLimit - 1)
    Else
        smLnSave(14, imLnRowNo) = ""
        smLnSave(15, imLnRowNo) = ""
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetDate                        *
'*                                                     *
'*             Created:8/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move values from work area to  *
'*                      records                        *
'*                                                     *
'*******************************************************
Private Sub mSetDate(ilRemoveZeroSpotWk As Integer)
    Dim ilWkIndex As Integer
    Dim ilCff As Integer
    Dim ilPrevCff As Integer
    Dim ilPrev As Integer
    Dim ilNoWkCombine As Integer
    Dim ilIndex As Integer
    Dim ilMatch As Integer
    Dim ilUpperBound As Integer
    'Dim llStartDate As Long
    Dim slStartDate As String
    'Dim llEndDate As Long
    Dim slEndDate As String
    Dim ilDay As Integer
    Dim ilCffIndex As Integer
    Dim ilLoop As Integer
    'Dim slTestDate As String
    Dim ilMoveIndex As Integer
    Dim ilWkEnd As Integer
    Dim llSpots As Long
    'Dim ilResetCff As Integer
    Dim ilLastCffUsed As Integer
    Dim ilClf As Integer
    Dim ilTestCff As Integer
    Dim ilFound As Integer
    Dim slLineType As String
    Dim llDate As Long
    Dim llPkStartDate As Long
    Dim llPkEndDate As Long
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    Dim ilAirWk As Integer
    Dim ilStdPkg As Integer
    Dim slWDFlag As String
    ReDim ilDailySpots(0 To 6) As Integer
    ReDim ilDays(0 To 6) As Integer
    Dim slStr As String
    Dim llWkPrice As Long
    slLineType = mGetLineType(imLnRowNo)
    If (smLnSave(2, imLnRowNo) <> smDGSave(13)) Or (smLnSave(3, imLnRowNo) <> smDGSave(14)) Then
        If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
            imLnChg = True
        End If
    End If
    If (Trim$(smDGSave(13)) = "") Or (Trim$(smDGSave(14)) = "") Then
        imLnSave(4, imLnRowNo) = 0
        smLnSave(2, imLnRowNo) = ""
        smLnSave(3, imLnRowNo) = ""
    ElseIf (Not gValidTime(smDGSave(13))) Or (Not gValidTime(smDGSave(14))) Then
        imLnSave(4, imLnRowNo) = 0
        smLnSave(2, imLnRowNo) = ""
        smLnSave(3, imLnRowNo) = ""
    Else
        imLnSave(4, imLnRowNo) = 1
        smLnSave(2, imLnRowNo) = smDGSave(13)
        smLnSave(3, imLnRowNo) = smDGSave(14)
    End If
    ReDim tmFlWkInfo(0 To 0) As FLWKINFO
    ilStdPkg = False
    If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
        For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
            If tgPackLn(ilLoop).iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine Then
                If (tgPackLn(ilLoop).iStdPkg = True) Or (slLineType = "E") Then
                    ilStdPkg = True
                End If
                Exit For
            End If
        Next ilLoop
    End If
    If ilStdPkg Then
        llPkStartDate = 0
        llPkEndDate = 0
        ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
        Do While ilCff <> -1
            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                llFlStartDate = tgCffCntr(ilCff).lStartDate
                llFlEndDate = tgCffCntr(ilCff).lEndDate
                If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                    llSpots = 0
                    For ilDay = 0 To 6 Step 1
                        llSpots = llSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                        ilDailySpots(ilDay) = tgCffCntr(ilCff).CffRec.iDay(ilDay)
                    Next ilDay
                    slWDFlag = "D"
                Else
                    slWDFlag = "W"
                    llSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                End If
                If llSpots > 0 Then
                    If llPkStartDate = 0 Then
                        llPkStartDate = llFlStartDate
                        llPkEndDate = llFlEndDate
                    Else
                        If llFlStartDate < llPkStartDate Then
                            llPkStartDate = llFlStartDate
                        End If
                        If llFlEndDate > llPkEndDate Then
                            llPkEndDate = llFlEndDate
                        End If
                    End If
                End If
            End If
            ilCff = tgCffCntr(ilCff).iNextCff
        Loop
        llPkStartDate = gDateValue(gObtainPrevMonday(Format$(llPkStartDate, "m/d/yy")))
        For llDate = llPkStartDate To llPkEndDate Step 7
            ilAirWk = False
            ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
            Do While ilCff <> -1
                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                    llFlStartDate = tgCffCntr(ilCff).lStartDate
                    llFlEndDate = tgCffCntr(ilCff).lEndDate
                    If (llFlEndDate >= llDate) And (llFlStartDate <= llDate + 6) Then
                        ilAirWk = True
                        If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                            llSpots = 0
                            For ilDay = 0 To 6 Step 1
                                If (llDate + ilDay >= llFlStartDate) And (llDate + ilDay <= llFlEndDate) Then
                                    llSpots = llSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                End If
                            Next ilDay
                        Else
                            llSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                        End If
                        Exit Do
                    End If
                End If
                ilCff = tgCffCntr(ilCff).iNextCff
            Loop
            If ilAirWk Then
                tmFlWkInfo(UBound(tmFlWkInfo)).lDate = llDate
                tmFlWkInfo(UBound(tmFlWkInfo)).iSpots = llSpots
                tmFlWkInfo(UBound(tmFlWkInfo)).iStatus = 1
                ReDim Preserve tmFlWkInfo(0 To UBound(tmFlWkInfo) + 1) As FLWKINFO
            End If
        Next llDate
    End If
    ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
    ilUpperBound = UBound(tgCffCntr)
    ilLastCffUsed = LBound(tgCffCntr)
    'Combine daily split weeks
    ilWkIndex = imLB1Or2    'LBound(smAirSave, 2)
    ilWkEnd = UBound(smAirSave, 2)
    Do
        If (smAirSave(13, ilWkIndex) = "Y") And (imAirSave(2, ilWkIndex) = 0) Then
            smAirSave(13, ilWkIndex) = "N"
            ilIndex = gWeekDayStr(smAirSave(1, ilWkIndex + 1))
            For ilDay = ilIndex To 6 Step 1
                smAirSave(4 + ilDay, ilWkIndex) = smAirSave(4 + ilDay, ilWkIndex + 1)
            Next ilDay
            For ilMoveIndex = ilWkIndex + 2 To ilWkEnd Step 1
                For ilLoop = LBound(smAirSave, 1) To UBound(smAirSave, 1) Step 1
                    smAirSave(ilLoop, ilMoveIndex - 1) = smAirSave(ilLoop, ilMoveIndex)
                Next ilLoop
                For ilLoop = LBound(imAirSave, 1) To UBound(imAirSave, 1) Step 1
                    imAirSave(ilLoop, ilMoveIndex - 1) = imAirSave(ilLoop, ilMoveIndex)
                Next ilLoop
                For ilLoop = LBound(smAirShow, 1) To UBound(smAirShow, 1) Step 1
                    smAirShow(ilLoop, ilMoveIndex - 1) = smAirShow(ilLoop, ilMoveIndex)
                Next ilLoop
            Next ilMoveIndex
            ilWkEnd = ilWkEnd - 1
            ReDim Preserve smAirShow(0 To 14, 0 To ilWkEnd) As String
            ReDim Preserve smAirSave(0 To 13, 0 To ilWkEnd) As String
            ReDim Preserve imAirSave(0 To 2, 0 To ilWkEnd) As Integer
        End If
        ilWkIndex = ilWkIndex + 1
    Loop Until ilWkIndex >= ilWkEnd
    If ilRemoveZeroSpotWk Then
        'For ilWkIndex = LBound(smAirSave, 2) To UBound(smAirSave, 2) - 1 Step 1
        For ilWkIndex = imLB1Or2 To UBound(smAirSave, 2) - 1 Step 1
            If (smAirSave(13, ilWkIndex) = "N") And (imAirSave(1, ilWkIndex) = 0) Then  'Test non-split weeks only
                If (imAirSave(2, ilWkIndex) = 0) And (Not tgClfCntr(imLnRowNo - 1).iGame) Then
                    llSpots = 0
                    For ilDay = 0 To 6 Step 1
                        If smAirSave(4 + ilDay, ilWkIndex) <> "" Then
                            llSpots = llSpots + Val(smAirSave(4 + ilDay, ilWkIndex))
                        End If
                    Next ilDay
                Else
                    If tgClfCntr(imLnRowNo - 1).iGame Then
                        llSpots = 1
                    Else
                        llSpots = Val(smAirSave(3, ilWkIndex))
                    End If
                End If
                If llSpots = 0 Then
                    imAirSave(1, ilWkIndex) = 1
                End If
            End If
        Next ilWkIndex
    End If
    ilPrevCff = -1
    ilWkIndex = imLB1Or2    'LBound(smAirSave, 2)
    Do While ilWkIndex < UBound(smAirSave, 2)
        ilNoWkCombine = 1
        If imAirSave(1, ilWkIndex) = 0 Then  'Combine weeks( Air weeks only)
            If (smAirSave(13, ilWkIndex) = "N") And (Not mIsLnGame(imLnRowNo)) Then  'Not a split week
                ilMatch = True
                For ilLoop = ilWkIndex + 1 To UBound(smAirSave, 2) - 1 Step 1
                    If imAirSave(1, ilLoop) <> 0 Then
                        Exit For
                    End If
                    If smAirSave(13, ilLoop) = "Y" Then
                        Exit For
                    End If
                    'For ilIndex = LBound(smAirSave, 1) + 2 To UBound(smAirSave, 1) - 1 Step 1
                    For ilIndex = imLB1Or2 + 2 To UBound(smAirSave, 1) - 1 Step 1
                        If StrComp(smAirSave(ilIndex, ilWkIndex), smAirSave(ilIndex, ilLoop), 1) <> 0 Then
                            ilMatch = False
                            Exit For
                        End If
                    Next ilIndex
                    If Not ilMatch Then
                        Exit For
                    End If
                    ilNoWkCombine = ilNoWkCombine + 1
                Next ilLoop
            End If
            slStartDate = smAirSave(1, ilWkIndex)
            If smAirSave(13, ilWkIndex + ilNoWkCombine - 1) = "Y" Then
                slEndDate = smAirSave(2, ilWkIndex + ilNoWkCombine)
            Else
                slEndDate = smAirSave(2, ilWkIndex + ilNoWkCombine - 1)
            End If
            If ilCff = -1 Then  'Use next available
                'ilCffIndex = ilUpperBound
                '6/7/16: Moved code here
                'GoSub lObtainNextCff2
                ilCffIndex = -1
                For ilLoop = ilLastCffUsed To UBound(tgCffCntr) - 1 Step 1
                    If tgCffCntr(ilLoop).iStatus = -1 Then
                        'Test Chain
                        ilFound = False
                        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) Step 1
                            ilTestCff = tgClfCntr(ilClf).iFirstCff
                            If ilTestCff = ilLoop Then
                                ilFound = True
                                Exit For
                            End If
                            Do While ilTestCff <> -1
                                ilTestCff = tgCffCntr(ilTestCff).iNextCff
                                If ilTestCff = ilLoop Then
                                    ilFound = True
                                    Exit For
                                End If
                            Loop
                        Next ilClf
                        If Not ilFound Then
                            ilCffIndex = ilLoop
                            ilLastCffUsed = ilLoop
                            Exit For
                        End If
                    End If
                Next ilLoop
                If ilCffIndex = -1 Then
                    ilCffIndex = UBound(tgCffCntr)
                    ilLastCffUsed = ilCffIndex
                    ReDim Preserve tgCffCntr(0 To ilCffIndex + 100) As CFFLIST
                    For ilLoop = ilCffIndex + 1 To UBound(tgCffCntr) Step 1
                        tgCffCntr(ilLoop).iStatus = -1 'Not Used
                        tgCffCntr(ilLoop).lRecPos = 0
                        tgCffCntr(ilLoop).iNextCff = -1
                    Next ilLoop
                End If
                tgCffCntr(ilCffIndex).iNextCff = -1
                tgCffCntr(ilCffIndex).iStatus = 0
                tgCffCntr(ilCffIndex).lRecPos = 0
                If tgClfCntr(imLnRowNo - 1).iFirstCff = -1 Then
                    tgClfCntr(imLnRowNo - 1).iFirstCff = ilCffIndex 'ilUpperBound
                Else
                    tgCffCntr(ilPrevCff).iNextCff = ilCffIndex  'ilUpperBound
                End If
                ilPrevCff = ilCffIndex  'ilUpperBound
            Else
                ilCffIndex = ilCff
                ilPrevCff = ilCff
                ilCff = tgCffCntr(ilCff).iNextCff
            End If
            If tgCffCntr(ilCffIndex).iStatus = -1 Then
                tgCffCntr(ilCffIndex).iStatus = 0   'New to not used
            ElseIf tgCffCntr(ilCffIndex).iStatus = 2 Then
                tgCffCntr(ilCffIndex).iStatus = 1    'Old to use
            End If
            If tgCffCntr(ilCffIndex).iStatus = 0 Then
                tgCffCntr(ilCffIndex).CffRec.lChfCode = tgChfCntr.lCode
            End If
            tgCffCntr(ilCffIndex).CffRec.iClfLine = tgClfCntr(imLnRowNo - 1).ClfRec.iLine
            tgCffCntr(ilCffIndex).CffRec.iCntRevNo = tgClfCntr(imLnRowNo - 1).ClfRec.iCntRevNo
            tgCffCntr(ilCffIndex).CffRec.iPropVer = tgClfCntr(imLnRowNo - 1).ClfRec.iPropVer
            gPackDate slStartDate, tgCffCntr(ilCffIndex).CffRec.iStartDate(0), tgCffCntr(ilCffIndex).CffRec.iStartDate(1)
            gPackDate slEndDate, tgCffCntr(ilCffIndex).CffRec.iEndDate(0), tgCffCntr(ilCffIndex).CffRec.iEndDate(1)
            tgCffCntr(ilCffIndex).lStartDate = gDateValue(slStartDate)
            tgCffCntr(ilCffIndex).lEndDate = gDateValue(slEndDate)
            If (imAirSave(2, ilWkIndex) = 0) And (Not tgClfCntr(imLnRowNo - 1).iGame) Then
                tgCffCntr(ilCffIndex).CffRec.iSpotsWk = 0
                tgCffCntr(ilCffIndex).CffRec.sDyWk = "D"
                llSpots = 0
                For ilDay = 0 To 6 Step 1
                    If smAirSave(4 + ilDay, ilWkIndex) = "" Then
                        tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = 0
                    Else
                        tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = Val(smAirSave(4 + ilDay, ilWkIndex))
                        llSpots = llSpots + tgCffCntr(ilCffIndex).CffRec.iDay(ilDay)
                    End If
                Next ilDay
            Else
                If tgClfCntr(imLnRowNo - 1).iGame Then
                    tgCffCntr(ilCffIndex).CffRec.iSpotsWk = 0
                Else
                    tgCffCntr(ilCffIndex).CffRec.sDyWk = "W"
                    tgCffCntr(ilCffIndex).CffRec.iSpotsWk = Val(smAirSave(3, ilWkIndex))
                    llSpots = tgCffCntr(ilCffIndex).CffRec.iSpotsWk
                End If
                For ilDay = 0 To 6 Step 1
                    If smAirSave(4 + ilDay, ilWkIndex) = "Y" Then
                        tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = 1
                    Else
                        tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = 0
                    End If
                Next ilDay
            End If
            'Test if split week
            tgCffCntr(ilCffIndex).CffRec.sDelete = "N"
            If smAirSave(13, ilWkIndex) = "Y" Then
                ilWkIndex = ilWkIndex + 1
                tgCffCntr(ilCffIndex).CffRec.iXSpotsWk = Val(smAirSave(3, ilWkIndex))
                For ilDay = 0 To 6 Step 1
                    If smAirSave(4 + ilDay, ilWkIndex) = "Y" Then
                        tgCffCntr(ilCffIndex).CffRec.sXDay(ilDay) = "1"
                    Else
                        tgCffCntr(ilCffIndex).CffRec.sXDay(ilDay) = "0"
                    End If
                Next ilDay
            Else
                tgCffCntr(ilCffIndex).CffRec.iXSpotsWk = 0
                For ilDay = 0 To 6 Step 1
                    tgCffCntr(ilCffIndex).CffRec.sXDay(ilDay) = "0"
                Next ilDay
            End If
            'Determine if week existed- if so add week
            If ilStdPkg Then
                For llDate = gDateValue(slStartDate) To gDateValue(slEndDate) Step 7
                    ilFound = False
                    For ilLoop = LBound(tmFlWkInfo) To UBound(tmFlWkInfo) Step 1
                        If (llDate >= tmFlWkInfo(ilLoop).lDate) And (llDate <= tmFlWkInfo(ilLoop).lDate + 6) Then
                            If llSpots = tmFlWkInfo(ilLoop).iSpots Then
                                ilFound = True
                            End If
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        ilFound = mAddStdPkgWk(hmPvf, imLnRowNo, llDate, slWDFlag, llSpots, ilDailySpots(), ilLastCffUsed, tmRcf, imDGSave(), imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), hmCHF, hmClf, hmSdf, hmSmf, hmGhf, hmGsf)
                    End If
                Next llDate
            End If
            tgCffCntr(ilCffIndex).CffRec.sPriceType = ""
            gFindMatch smAirSave(11, ilWkIndex), 1, lbcPriceType(1)
            If gLastFound(lbcPriceType(1)) > 0 Then
                tgCffCntr(ilCffIndex).CffRec.sPriceType = Left$(lbcPriceType(1).List(gLastFound(lbcPriceType(1))), 1)
                tgCffCntr(ilCffIndex).CffRec.lActPrice = 0  'Later- might want to store average package price
            Else
                If smAirSave(11, ilWkIndex) <> "" Then
                    tgCffCntr(ilCffIndex).CffRec.sPriceType = "T"
                    If smLineType <> "E" Then
                        If tgCffCntr(ilCffIndex).CffRec.lActPrice <> gStrDecToLong(smAirSave(11, ilWkIndex), 2) Then
                            tgClfCntr(imLnRowNo - 1).sCurrentPrice = ""
                        End If
                        tgCffCntr(ilCffIndex).CffRec.lActPrice = gStrDecToLong(smAirSave(11, ilWkIndex), 2)
                    Else
                        If llSpots > 0 Then
                            tgCffCntr(ilCffIndex).CffRec.lActPrice = gStrDecToLong(smAirSave(11, ilWkIndex), 2)
                        Else
                            tgCffCntr(ilCffIndex).CffRec.lActPrice = 0
                        End If
                    End If
                Else
                    tgCffCntr(ilCffIndex).CffRec.sPriceType = ""    'Not defined
                    tgCffCntr(ilCffIndex).CffRec.lActPrice = 0
                End If
            End If
            '3/18/17: Determine actual price if not specified
            If Trim$(smAirSave(12, ilWkIndex)) = "" Then
                llWkPrice = mComputeRCWkPrice(tmRcf.iCode, imLnSave(1, imLnRowNo), slStartDate, tgClfCntr(imLnRowNo - 1).ClfRec.iVefCode, imLnRowNo)
                If llWkPrice > 0 Then
                    For ilDay = 0 To 6 Step 1
                        If tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) > 0 Then
                            ilDays(ilDay) = True
                        Else
                            ilDays(ilDay) = False
                        End If
                    Next ilDay
                    slStr = mRCPriceAdjust(tmRcf, llWkPrice, False, imLnRowNo, imDGSave(), 0, 0, Val(smLnSave(16, imLnRowNo)), llSpots, ilDays(), imLnSave(1, imLnRowNo))
                    smAirSave(12, ilWkIndex) = gRoundStr(slStr, "1.00", 0)    'gLongToStrDec(llWkPrice, 0) '& ".00"
                End If
            End If
            tgCffCntr(ilCffIndex).CffRec.lPropPrice = gStrDecToLong(smAirSave(12, ilWkIndex), 0)    '2)
        End If
        ilWkIndex = ilWkIndex + ilNoWkCombine
    Loop
    'If all flights removed- then its CBS
    If (ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff) And (ilCff <> -1) Then
        tgClfCntr(imLnRowNo - 1).iCancel = True
        gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate
        slEndDate = gDecOneDay(slStartDate)
        gPackDate slEndDate, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
        gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), tgCffCntr(ilCff).lStartDate    'Week Start date
        gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), tgCffCntr(ilCff).lEndDate    'Week Start date
        smLnSave(14, imLnRowNo) = slStartDate
        smLnSave(15, imLnRowNo) = slEndDate
        If tgCffCntr(ilCff).iStatus = -1 Then
            tgCffCntr(ilCff).iStatus = 0    'Reset to new
        ElseIf tgCffCntr(ilCff).iStatus = 2 Then
            tgCffCntr(ilCff).iStatus = 1    'Reset old
        End If
        ilPrevCff = ilCff
        ilCff = tgCffCntr(ilCff).iNextCff
        Do While ilCff <> -1
            If tgCffCntr(ilCff).iStatus = 0 Then
                tgCffCntr(ilPrevCff).iNextCff = tgCffCntr(ilCff).iNextCff
                tgCffCntr(ilCff).iStatus = -1   'New to not used
                ilPrev = ilCff
                ilCff = tgCffCntr(ilCff).iNextCff
                tgCffCntr(ilPrev).iNextCff = -1
            ElseIf tgCffCntr(ilCff).iStatus = 1 Then
                ilPrevCff = ilCff
                tgCffCntr(ilCff).iStatus = 2    'Old to delete
                ilCff = tgCffCntr(ilCff).iNextCff
            Else
                ilCff = tgCffCntr(ilCff).iNextCff
            End If
        Loop
    Else
        'Note ilPrevCff is set to prev cff
        Do While ilCff <> -1
            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = -1) Then
                tgCffCntr(ilPrevCff).iNextCff = tgCffCntr(ilCff).iNextCff
                tgCffCntr(ilCff).iStatus = -1   'New to not used
                ilPrev = ilCff
                ilCff = tgCffCntr(ilCff).iNextCff
                tgCffCntr(ilPrev).iNextCff = -1
            Else
                ilPrevCff = ilCff
                tgCffCntr(ilCff).iStatus = 2    'Old to delete
                ilCff = tgCffCntr(ilCff).iNextCff
            End If
        Loop
    End If
    If ilStdPkg Then
        For ilLoop = LBound(tmFlWkInfo) To UBound(tmFlWkInfo) - 1 Step 1
            llDate = tmFlWkInfo(ilLoop).lDate
            ilFound = False
            ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
            Do While ilCff <> -1
                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                    llFlStartDate = tgCffCntr(ilCff).lStartDate
                    llFlEndDate = tgCffCntr(ilCff).lEndDate
                    If (llFlEndDate >= llDate) And (llFlStartDate <= llDate + 6) Then
                        ilFound = True
                        Exit Do
                    End If
                End If
                ilCff = tgCffCntr(ilCff).iNextCff
            Loop
            If Not ilFound Then   'Delete week
                ilFound = mAddStdPkgWk(hmPvf, imLnRowNo, llDate, "W", 0, ilDailySpots(), ilLastCffUsed, tmRcf, imDGSave(), imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), hmCHF, hmClf, hmSdf, hmSmf, hmGhf, hmGsf)
            End If
        Next ilLoop
    End If
    If (Not imInSave) Or (slLineType <> "H") Then
        mLnTotals imLnRowNo, imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), "N", imNoWkDays()
    End If
    Exit Sub

End Sub
Private Sub pbcDGSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        If imDGBoxNo <> -1 Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If GetFocus() <> pbcDGSTab.HWnd Then
        Exit Sub
    End If
    mAirSetShow imAirBoxNo
    imAirBoxNo = -1
    imTabDirection = -1  'Set-Right to Left
    Select Case imDGBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            If tgClfCntr(imLnRowNo - 1).iInPast Then    'Only allow GenNew if weeks are in the past
                ilBox = DGNOWKSINDEX
            Else
                ilBox = DGSTARTDATEINDEX
            End If
        Case DGNOWKSINDEX
                ilBox = imDGBoxNo - 1
        Case DGSPOTWKINDEX
            mDGSetShow imDGBoxNo
            If (Trim$(smDGSave(2)) = "") Or (Val(smDGSave(2)) = 0) Then
                Beep
                ilBox = imDGBoxNo
            Else
                ilBox = imDGBoxNo - 1
            End If
            imDGBoxNo = ilBox
            mDGEnableBox ilBox
            Exit Sub
        Case DGDAILYINDEX
            If tgClfCntr(imLnRowNo - 1).iGame Then
                ilBox = DGRUNINDEX
            Else
                If (imDGSave(2) = 1) Or (imDGSave(2) = 3) Then
                    ilBox = DGSPOTWKINDEX
                Else
                    ilBox = DGDWPINDEX
                End If
            End If
        Case DGSTARTDATEINDEX
            slStr = edcDGDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDGDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDGDropDown.SetFocus
                Exit Sub
            End If
            mDGSetShow imDGBoxNo
            If (cmcGenAll.Visible) And (cmcGenAll.Enabled) Then
                imDGBoxNo = -1
                cmcGenAll.SetFocus
                Exit Sub
            Else
                If (cmcGenNew.Visible) And (cmcGenNew.Enabled) Then
                    imDGBoxNo = -1
                    cmcGenNew.SetFocus
                    Exit Sub
                Else
                    'ilBox = DGDAILYINDEX + 6
                    If (smLineType = "O") Or (smLineType = "A") Then
                        ilBox = DGDAILYINDEX + 6
                    Else
                        If Contract.mAllowedToChgSpotPrice(imLnRowNo) Then
                            ilBox = DGPRICEINDEX
                        Else
                            ilBox = DGDAILYINDEX + 6
                        End If
                    End If
                    imDGBoxNo = ilBox
                    mDGEnableBox ilBox
                    Exit Sub
                End If
            End If
        Case DGENDDATEINDEX
            slStr = edcDGDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDGDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDGDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imDGBoxNo - 1
        Case 0
        Case Else
            ilBox = imDGBoxNo - 1
    End Select
    mDGSetShow imDGBoxNo
    imDGBoxNo = ilBox
    mDGEnableBox ilBox
End Sub
Private Sub pbcDGTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If GetFocus() <> pbcDGTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = 0  'Reset-Left to right
    Select Case imDGBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            mAirSetShow imAirBoxNo
            imAirBoxNo = -1
            'ilBox = DGPRICEINDEX    'DGDAILYINDEX + 6
            If ((smLineType = "O") Or (smLineType = "A")) And (tgSpf.sCAudPkg <> "Y") Then
                ilBox = DGDAILYINDEX + 6  'DGPREEMPTINDEX
            Else
                If smLineType = "H" Then
                    slStr = mGetPkLineType(imLnRowNo)
                    If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                        ilBox = DGDAILYINDEX + 6  'DGPREEMPTINDEX
                    Else
                        ilBox = DGPRICEINDEX
                    End If
                Else
                    ilBox = DGPRICEINDEX
                End If
            End If
        Case DGSTARTDATEINDEX
            slStr = edcDGDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDGDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDGDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imDGBoxNo + 1
        Case DGENDDATEINDEX
            slStr = edcDGDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDGDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDGDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imDGBoxNo + 1
        Case DGRUNINDEX
            If tgClfCntr(imLnRowNo - 1).iGame Then
                ilBox = DGDAILYINDEX
            Else
                ilBox = DGDWPINDEX
            End If
        Case DGDWPINDEX
            If (lbcDWP.ListIndex = 1) Or (lbcDWP.ListIndex = 3) Then    'Weekly
                ilBox = DGSPOTWKINDEX
            Else    'Daily
                ilBox = DGDAILYINDEX
            End If
        Case DGSPOTWKINDEX
            mDGSetShow imDGBoxNo
            If (Trim$(smDGSave(2)) = "") Or (Val(smDGSave(2)) = 0) Then
                Beep
                ilBox = imDGBoxNo
            Else
                ilBox = imDGBoxNo + 1
            End If
            imDGBoxNo = ilBox
            mDGEnableBox ilBox
            Exit Sub
        Case DGDAILYINDEX + 6
            If ((smLineType = "O") Or (smLineType = "A")) And (tgSpf.sCAudPkg <> "Y") Then
                ilBox = DGOVSTARTINDEX  'DGPREEMPTINDEX
            ElseIf Not Contract.mAllowedToChgSpotPrice(imLnRowNo) Then
                ilBox = DGOVSTARTINDEX  'DGPREEMPTINDEX
            Else
                If smLineType = "H" Then
                    slStr = mGetPkLineType(imLnRowNo)
                    If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                        If (Not tgClfCntr(imLnRowNo - 1).iOverride) Then
                            mDGSetShow imDGBoxNo
                
                            If (cmcGenAll.Visible) And (cmcGenAll.Enabled) Then
                                imDGBoxNo = -1
                                cmcGenAll.SetFocus
                                Exit Sub
                            Else
                                If (cmcGenNew.Visible) And (cmcGenNew.Enabled) Then
                                    imDGBoxNo = -1
                                    cmcGenNew.SetFocus
                                    Exit Sub
                                Else
                                    If tgClfCntr(imLnRowNo - 1).iInPast Then    'Only allow GenNew if weeks are in the past
                                        ilBox = DGNOWKSINDEX
                                    Else
                                        ilBox = DGSTARTDATEINDEX
                                    End If
                                    imDGBoxNo = ilBox
                                    mDGEnableBox ilBox
                                    Exit Sub
                                End If
                            End If
                        Else
                            ilBox = DGOVSTARTINDEX  'DGPREEMPTINDEX
                        End If
                    Else
                        ilBox = DGPRICEINDEX
                    End If
                Else
                    ilBox = DGPRICEINDEX
                End If
            End If
        Case DGPRICEINDEX
            If (Not tgClfCntr(imLnRowNo - 1).iOverride) Then
                mDGSetShow imDGBoxNo
    
                If (cmcGenAll.Visible) And (cmcGenAll.Enabled) Then
                    imDGBoxNo = -1
                    cmcGenAll.SetFocus
                    Exit Sub
                Else
                    If (cmcGenNew.Visible) And (cmcGenNew.Enabled) Then
                        imDGBoxNo = -1
                        cmcGenNew.SetFocus
                        Exit Sub
                    Else
                        If tgClfCntr(imLnRowNo - 1).iInPast Then    'Only allow GenNew if weeks are in the past
                            ilBox = DGNOWKSINDEX
                        Else
                            ilBox = DGSTARTDATEINDEX
                        End If
                        imDGBoxNo = ilBox
                        mDGEnableBox ilBox
                        Exit Sub
                    End If
                End If
            Else
                ilBox = DGOVSTARTINDEX  'DGPREEMPTINDEX
            End If
        Case DGOVENDINDEX   'DGDAILYINDEX + 6'Last control within header
            mDGSetShow imDGBoxNo

            If (cmcGenAll.Visible) And (cmcGenAll.Enabled) Then
                imDGBoxNo = -1
                cmcGenAll.SetFocus
                Exit Sub
            Else
                If (cmcGenNew.Visible) And (cmcGenNew.Enabled) Then
                    imDGBoxNo = -1
                    cmcGenNew.SetFocus
                    Exit Sub
                Else
                    If tgClfCntr(imLnRowNo - 1).iInPast Then    'Only allow GenNew if weeks are in the past
                        ilBox = DGNOWKSINDEX
                    Else
                        ilBox = DGSTARTDATEINDEX
                    End If
                    imDGBoxNo = ilBox
                    mDGEnableBox ilBox
                    Exit Sub
                End If
            End If
        Case Else
            ilBox = imDGBoxNo + 1
    End Select
    mDGSetShow imDGBoxNo
    imDGBoxNo = ilBox
    mDGEnableBox ilBox
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetDemoFlightTotals            *
'*                                                     *
'*             Created:9/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute the Research Values for*
'*                      the flight                     *
'*                                                     *
'*******************************************************
Sub mGetDemoFlightTotals()
    Dim ilLoop As Integer
    Dim ilDBTotals As Integer
    pbcAirDemo.Cls
    For ilLoop = imLBCTCtrls To UBound(tmDFTCtrls) - 1 Step 1
        tmDFTCtrls(ilLoop).sShow = ""
    Next ilLoop
    ilDBTotals = imDBTotals
    imDBTotals = 0  'Line totals
    mGetResearchTotals imLnRowNo, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imDBTotals, True, imCppCpm, imCntrVefCode, imVpfIndex, imTerminate, imChgMode, imInResearch, imLnTotalType
    imDBTotals = ilDBTotals
    pbcAirDemo_Paint
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFlightTotals                   *
'*                                                     *
'*             Created:8/03/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute Flight Totlas          *
'*                                                     *
'*******************************************************
Sub mFlightTotals()
    Dim ilLoop As Integer
    Dim ilAirWks As Integer
    Dim ilWks As Integer
    'Dim ilNoSpots As Integer
    Dim llNoSpots As Long
    'Dim llDollars As Long
    Dim slDollars As String
    Dim ilDay As Integer
    Dim slStr As String
    ilAirWks = 0
    ilWks = 0
    llNoSpots = 0
    'llDollars = 0
    slDollars = "0.00"
    'For ilLoop = LBound(smAirSave, 2) To UBound(smAirSave, 2) - 1 Step 1
    For ilLoop = imLB1Or2 To UBound(smAirSave, 2) - 1 Step 1
        If imAirSave(1, ilLoop) = 0 Then    'Aired week
            If smAirSave(13, ilLoop) = "Y" Then
                If ilLoop <> UBound(smAirSave, 2) - 1 Then
                    If smAirSave(13, ilLoop + 1) = "Y" Then
                        ilAirWks = ilAirWks + 1
                        ilWks = ilWks + 1
                    End If
                End If
            Else
                ilAirWks = ilAirWks + 1
                ilWks = ilWks + 1
            End If
            If imAirSave(2, ilLoop) = 0 Then    'Daily
                For ilDay = 4 To 10 Step 1
                    llNoSpots = llNoSpots + Val(smAirSave(ilDay, ilLoop))
                    If smLineType <> "E" Then
                        'llDollars = llDollars + Val(smAirSave(ilDay, ilLoop)) * gStrDecToLong(smAirSave(11, ilLoop), 2)
                        slDollars = gAddStr(slDollars, gMulStr(smAirSave(ilDay, ilLoop), smAirSave(11, ilLoop)))
                    Else
                        'llDollars = llDollars + gStrDecToLong(smAirSave(11, ilLoop), 2)
                        slDollars = gAddStr(slDollars, smAirSave(11, ilLoop))
                    End If
                Next ilDay
            Else
                llNoSpots = llNoSpots + Val(smAirSave(3, ilLoop))
                If smLineType <> "E" Then
                    'llDollars = llDollars + Val(smAirSave(3, ilLoop)) * gStrDecToLong(smAirSave(11, ilLoop), 2)
                    slDollars = gAddStr(slDollars, gMulStr(smAirSave(3, ilLoop), smAirSave(11, ilLoop)))
                Else
                    'llDollars = llDollars + gStrDecToLong(smAirSave(11, ilLoop), 2)
                    slDollars = gAddStr(slDollars, smAirSave(11, ilLoop))
                End If
            End If
        Else
            If smAirSave(13, ilLoop) = "Y" Then
                If ilLoop <> UBound(smAirSave, 2) - 1 Then
                    If smAirSave(13, ilLoop + 1) = "Y" Then
                        ilWks = ilWks + 1
                    End If
                End If
            Else
                ilWks = ilWks + 1
            End If
        End If
    Next ilLoop
    If (smLineType = "O") Or (smLineType = "A") Or (smLineType = "E") Then
        'llDollars = gStrDecToLong(smLnSave(18, imLnRowNo), 2)
        slDollars = smLnSave(18, imLnRowNo)
    End If
    slStr = Trim$(str$(ilAirWks)) & "/" & Trim$(str$(ilWks))
    gSetShow pbcAirWk, slStr, tmFTCtrls(1)
    slStr = Trim$(str$(llNoSpots))
    gSetShow pbcAirWk, slStr, tmFTCtrls(2)
    'slStr = gLongToStrDec(llDollars, 2)
    slStr = gRoundStr(slDollars, ".01", 2)
    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
    gSetShow pbcAirWk, slStr, tmFTCtrls(3)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetAirPkPrice                  *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create price on flight screen  *
'*                      for Package lines              *
'*                                                     *
'*******************************************************
Sub mSetAirPkPrice()
    Dim ilLoop As Integer
    Dim ilNoSpots As Integer
    Dim ilTNoSpots As Integer
    Dim ilDay As Integer
    Dim llActPrice As Long
    Dim slStr As String
    Dim llDate As Long
    Dim llPrevBilledDollars As Long
    Dim ilPrevBilledSpots As Integer
    If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
        Exit Sub
    End If
    ilNoSpots = 0
        'Avg by week
        'For ilLoop = LBound(smAirSave, 2) To UBound(smAirSave, 2) - 1 Step 1
        For ilLoop = imLB1Or2 To UBound(smAirSave, 2) - 1 Step 1
            If imAirSave(1, ilLoop) = 0 Then    'Aired week
                If (smLineType = "O") Or (smLineType = "A") Then
                    If imAirSave(2, ilLoop) = 0 Then    'Daily
                        ilNoSpots = 0
                        For ilDay = 4 To 10 Step 1
                            ilNoSpots = ilNoSpots + Val(smAirSave(ilDay, ilLoop))
                        Next ilDay
                    Else
                        ilNoSpots = Val(smAirSave(3, ilLoop))
                    End If
                Else
                    ilNoSpots = 1
                End If
                If ilNoSpots > 0 Then
                    If smLineType <> "E" Then
                        'llActPrice = gStrDecToLong(smAirSave(11, ilLoop), 2) / ilNoSpots '+ .5
                        llActPrice = gStrDecToLong(smAirSave(11, ilLoop), 2)
                    Else
                        llActPrice = gStrDecToLong(smAirSave(11, ilLoop), 2)
                    End If
                Else
                    llActPrice = 0
                End If
                smAirSave(11, ilLoop) = gLongToStrDec(llActPrice, 2)
                gFormatStr smAirSave(11, ilLoop), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                gSetShow pbcAirWk, slStr, tmAirCtrls(WKPRICEINDEX)
                smAirShow(WKPRICEINDEX, ilLoop) = tmAirCtrls(WKPRICEINDEX).sShow
            End If
        Next ilLoop
    pbcAirWk_Paint
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateDate                     *
'*                                                     *
'*             Created:8/09/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create flights (first time line*
'*                      flights defined)               *
'*                                                     *
'*******************************************************
Public Sub mCreateDate(ilPaint As Integer)
'
'   mCreateDate ilPaint
'   Where:
'       ilPaint (I)- True=Paint week after creating
'                    False=Don't paint
'
    Dim ilUpperLimit As Integer
    Dim ilWkNo As Integer
    Dim llStartDate As Long
    Dim llSplitStart As Long
    Dim llSplitEnd As Long
    Dim llEndDate As Long
    Dim llDate As Long
    Dim slStr As String
    Dim ilDay As Integer
    Dim ilMonth As Integer
    Dim llDateAdj As Long
    Dim ilNoSpots As Integer
    Dim ilNoAirDays As Integer
    Dim ilNoDays As Integer
    Dim ilLoop As Integer
    Dim ilStartDay As Integer
    Dim ilEndDay As Integer
    Dim ilSplitSpotFirstLast As Integer
    Dim ilFirstWkNoSpots As Integer
    Dim ilLastWkNoSpots As Integer
    ReDim smAirShow(0 To 14, 0 To 1) As String
    ReDim smAirSave(0 To 13, 0 To 1) As String
    ReDim imAirSave(0 To 2, 0 To 1) As Integer
    ilUpperLimit = UBound(smAirShow, 2)
    ilWkNo = 1
    llStartDate = gDateValue(smDGSave(10))
    llEndDate = gDateValue(smDGSave(11))
    'Set days if user pressed GenAll prior to selecting days
    If imGenReq Then
        If (imDGSave(2) = 0) Or (imDGSave(2) = 2) Then    'Daily Or % Daily
            For ilLoop = 0 To 6 Step 1
                If imAllowedDays(ilLoop + LBound(imAllowedDays)) Then
                    If smDGSave(3 + ilLoop) = "" Then
                        smDGSave(3 + ilLoop) = "0"
                    End If
                End If
            Next ilLoop
        Else    'Weekly or % weekly
            For ilLoop = 0 To 6 Step 1
                If imAllowedDays(ilLoop + LBound(imAllowedDays)) Then
                    If smDGSave(3 + ilLoop) = "" Then
                        smDGSave(3 + ilLoop) = "Y"
                    End If
                End If
            Next ilLoop
        End If
    End If
    mCVehOptIndex imLnRowNo, imCntrVefCode, imVpfIndex, imTerminate  'Get imCntrVefCode
    imGenReq = False
    ilSplitSpotFirstLast = False
    If (imDGSave(2) = 1) Or (imDGSave(2) = 3) Then 'Weekly or % weekly
        ilStartDay = gWeekDayLong(llStartDate)
        ilEndDay = gWeekDayLong(llEndDate)
        If ilEndDay < ilStartDay Then
            ilSplitSpotFirstLast = True
            For ilDay = ilStartDay To 6 Step 1
                If smDGSave(3 + ilDay) = "Y" Then
                    ilNoAirDays = ilNoAirDays + 1
                    ilNoDays = ilNoDays + 1
                End If
            Next ilDay
            For ilDay = 0 To ilEndDay Step 1
                If smDGSave(3 + ilDay) = "Y" Then
                    ilNoDays = ilNoDays + 1
                End If
            Next ilDay
            If ilNoDays <> 0 Then
                ilFirstWkNoSpots = (Val(smDGSave(2)) * ilNoAirDays) / ilNoDays
            Else
                ilFirstWkNoSpots = Val(smDGSave(2))
            End If
            ilLastWkNoSpots = Val(smDGSave(2)) - ilFirstWkNoSpots
        End If
    End If
    llDate = llStartDate
    Do While llDate <= llEndDate
        Select Case imDGSave(1)
            Case 0  'Every Week
                imAirSave(1, ilUpperLimit) = 0
                slStr = "Yes"
                gSetShow pbcAirWk, slStr, tmAirCtrls(AIRWEEKINDEX)
                smAirShow(AIRWEEKINDEX, ilUpperLimit) = tmAirCtrls(AIRWEEKINDEX).sShow
            Case 1 To 3
                If (ilWkNo Mod (imDGSave(1) + 1)) = 1 Then
                    imAirSave(1, ilUpperLimit) = 0
                    slStr = "Yes"
                Else
                    imAirSave(1, ilUpperLimit) = 1
                    slStr = "No"
                End If
        End Select
        gSetShow pbcAirWk, slStr, tmAirCtrls(AIRWEEKINDEX)
        smAirShow(AIRWEEKINDEX, ilUpperLimit) = tmAirCtrls(AIRWEEKINDEX).sShow
        slStr = Trim$(str$(ilWkNo))
        gSetShow pbcAirWk, slStr, tmAirCtrls(WEEKNOINDEX)
        smAirShow(WEEKNOINDEX, ilUpperLimit) = tmAirCtrls(WEEKNOINDEX).sShow
        slStr = Format$(llDate, "m/d/yy")
        smAirSave(1, ilUpperLimit) = slStr
        ilStartDay = gWeekDayLong(llDate)
        If llDate + 6 - ilStartDay <= llEndDate Then
            slStr = Format$(llDate + 6 - ilStartDay, "m/d/yy")
            ilEndDay = gWeekDayLong(llDate + 6 - ilStartDay)
        Else
            slStr = Format$(llEndDate, "m/d/yy")
            ilEndDay = gWeekDayLong(llEndDate)
        End If
        smAirSave(2, ilUpperLimit) = slStr
        If imAirSave(1, ilUpperLimit) = 0 Then  '0=Air week
            slStr = smDGSave(2)
        Else
            slStr = ""
        End If
        'If (ilUpperLimit = LBound(smAirShow, 2)) And (ilSplitSpotFirstLast) Then
        If (ilUpperLimit = imLB1Or2) And (ilSplitSpotFirstLast) Then
            slStr = Trim$(str$(ilFirstWkNoSpots))
        End If
        If (gDateValue(smAirSave(2, ilUpperLimit)) = llEndDate) And (ilSplitSpotFirstLast) Then
            slStr = Trim$(str$(ilLastWkNoSpots))
        End If
        smAirSave(3, ilUpperLimit) = slStr
        gSetShow pbcAirWk, slStr, tmAirCtrls(SPOTPERWKINDEX)
        smAirShow(SPOTPERWKINDEX, ilUpperLimit) = tmAirCtrls(SPOTPERWKINDEX).sShow
        For ilDay = 0 To 6 Step 1   'Initialize days
            slStr = ""
            smAirSave(4 + ilDay, ilUpperLimit) = slStr
            gSetShow pbcAirWk, slStr, tmAirCtrls(DAILYSPOTSINDEX + ilDay)
            smAirShow(DAILYSPOTSINDEX + ilDay, ilUpperLimit) = tmAirCtrls(DAILYSPOTSINDEX + ilDay).sShow
        Next ilDay
        If (((imDGSave(2) = 0) Or (imDGSave(2) = 2)) And (Not tgClfCntr(imLnRowNo - 1).iGame)) Or (imAirSave(1, ilUpperLimit) = 1) Then
            For ilDay = ilStartDay To ilEndDay Step 1
                If imAirSave(1, ilUpperLimit) = 0 Then
                    slStr = smDGSave(3 + ilDay)
                Else
                    slStr = ""
                End If
                smAirSave(4 + ilDay, ilUpperLimit) = slStr
                gSetShow pbcAirWk, slStr, tmAirCtrls(DAILYSPOTSINDEX + ilDay)
                smAirShow(DAILYSPOTSINDEX + ilDay, ilUpperLimit) = tmAirCtrls(DAILYSPOTSINDEX + ilDay).sShow
            Next ilDay
        Else
            For ilDay = ilStartDay To ilEndDay Step 1
                If smDGSave(3 + ilDay) = "Y" Then
                    smAirSave(4 + ilDay, ilUpperLimit) = "Y"
                    slStr = "4"
                Else
                    smAirSave(4 + ilDay, ilUpperLimit) = "N"
                    slStr = " "
                End If
                gSetShow pbcAirWk, slStr, tmAirCtrls(DAILYSPOTSINDEX + ilDay)
                smAirShow(DAILYSPOTSINDEX + ilDay, ilUpperLimit) = tmAirCtrls(DAILYSPOTSINDEX + ilDay).sShow
            Next ilDay
        End If
        smAirSave(13, ilUpperLimit) = "N"

        ilWkNo = ilWkNo + 1
        ilUpperLimit = ilUpperLimit + 1
        ReDim Preserve smAirShow(0 To 14, 0 To ilUpperLimit) As String
        ReDim Preserve smAirSave(0 To 13, 0 To ilUpperLimit) As String
        ReDim Preserve imAirSave(0 To 2, 0 To ilUpperLimit) As Integer
        Do
            llDate = llDate + 1
        Loop Until gWeekDayLong(llDate) = 0  '0=Monday
    Loop
    'Set Rate Card Prices and Compute average price
    'If smDGSave(12) = lbcPriceType(0).List(2) Then  'Average- Compute value
'    mSetFlightPrice smDGSave(12), True
    'For ilLoop = LBound(smAirSave, 2) To UBound(smAirSave, 2) - 1 Step 1
    For ilLoop = imLB1Or2 To UBound(smAirSave, 2) - 1 Step 1
        mAdjDates ilLoop
'        If smAirSave(3, ilLoop) = "" Then
        If (imDGSave(2) = 0) Or (imDGSave(2) = 2) Then    'Daily Or % Daily
            imAirSave(2, ilLoop) = 0
            If imDGSave(2) > 1 Then
                slStr = "Daily by %"
            Else
                slStr = "Daily"
            End If
        Else
            imAirSave(2, ilLoop) = 1
            If imDGSave(2) > 1 Then
                slStr = "Weekly by %"
            Else
                slStr = "Weekly"
            End If
        End If
        gSetShow pbcAirWk, slStr, tmAirCtrls(DWINDEX)
        smAirShow(DWINDEX, ilLoop) = tmAirCtrls(DWINDEX).sShow
    Next ilLoop
    mSetFlightPrice smDGSave(12), True
    vbcAirWk.Min = imLB1Or2 'LBound(smAirShow, 2)
    If UBound(smAirShow, 2) - 2 <= vbcAirWk.LargeChange Then
        vbcAirWk.Max = imLB1Or2 'LBound(smAirShow, 2)
    Else
        vbcAirWk.Max = UBound(smAirShow, 2) - 1 - vbcAirWk.LargeChange
    End If
    vbcAirWk.Value = vbcAirWk.Min
    smLnSave(14, imLnRowNo) = smDGSave(10)
    smLnSave(15, imLnRowNo) = smDGSave(11)
    If ilPaint Then
        mFlightTotals
        mGetDemoFlightTotals
        imAirRowNo = 1
        pbcAirWk.Cls
        pbcAirWk_Paint
    End If
    tgClfCntr(imLnRowNo - 1).lFirstAvail = -1
    tgClfCntr(imLnRowNo - 1).iResearch = False
    tgClfCntr(imLnRowNo - 1).iCancel = False
    tgClfCntr(imLnRowNo - 1).ClfRec.sHideCBS = "N"
    imLnChg = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mAdjDates                       *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Adjust dates so dates are only *
'*                      days that have spots           *
'*                                                     *
'*******************************************************
Public Sub mAdjDates(ilFlightNo As Integer)
    Dim ilDay As Integer
    Dim slDate As String
    Dim ilAdjDates As Integer
    'This code works but ABC doesn't want this feature
    'If line dates are mon thur sun and only buying
    'wed and thur, this code will change the start
    'and end dates to be wed and thur respectively.
    '(multi weeks long are also changed)
    'Might want a table question for this feature
    'to be added later- set ilAdjDates
    ilAdjDates = False
    'Advance start date if required
    slDate = smAirSave(1, ilFlightNo)
    If ilAdjDates Then
        If (smAirSave(13, ilFlightNo) = "N") Or ((smAirSave(13, ilFlightNo) = "Y") And (smAirSave(13, ilFlightNo + 1) = "Y")) Then
            Do While gWeekDayStr(slDate) <> 0  '0=Monday
                slDate = gDecOneDay(slDate)
            Loop
            ilDay = 0
            Do While ((smAirSave(4 + ilDay, ilFlightNo) = "N") Or (smAirSave(4 + ilDay, ilFlightNo) = "") Or (smAirSave(4 + ilDay, ilFlightNo) = "0"))
                slDate = gIncOneDay(slDate)
                ilDay = ilDay + 1
                If ilDay > 6 Then
                    slDate = smAirSave(1, ilFlightNo)
                    Exit Do
                End If
            Loop
        End If
    End If
    smAirSave(1, ilFlightNo) = slDate
    slDate = gAddDayToDate(slDate)
    gSetShow pbcAirWk, slDate, tmAirCtrls(WEEKSTARTDATEINDEX)
    smAirShow(WEEKSTARTDATEINDEX, ilFlightNo) = tmAirCtrls(WEEKSTARTDATEINDEX).sShow
    'Decrement end date if required
    slDate = smAirSave(2, ilFlightNo)
    If ilAdjDates Then
        If (smAirSave(13, ilFlightNo) = "N") Or ((smAirSave(13, ilFlightNo) = "Y") And (smAirSave(13, ilFlightNo + 1) <> "Y")) Then
            Do While gWeekDayStr(slDate) <> 6  '6=Sunday
                slDate = gIncOneDay(slDate)
            Loop
            ilDay = 6
            Do While ((smAirSave(4 + ilDay, ilFlightNo) = "N") Or (smAirSave(4 + ilDay, ilFlightNo) = "") Or (smAirSave(4 + ilDay, ilFlightNo) = "0"))
                slDate = gDecOneDay(slDate)
                ilDay = ilDay - 1
                If ilDay < 0 Then
                    slDate = smAirSave(2, ilFlightNo)
                    Exit Do
                End If
            Loop
        End If
    End If
    smAirSave(2, ilFlightNo) = slDate
    slDate = gAddDayToDate(slDate)
    gSetShow pbcAirWk, slDate, tmAirCtrls(WEEKENDDATEINDEX)
    smAirShow(WEEKENDDATEINDEX, ilFlightNo) = tmAirCtrls(WEEKENDDATEINDEX).sShow
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFlightPrice                 *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set flight prices              *
'*                                                     *
'*******************************************************
Public Sub mSetFlightPrice(slInPrice As String, ilFromFlightArea As Integer)
    'Dim llTotalPrice As Long
    Dim slTotalPrice As String
    Dim slWkPrice As String
    Dim ilTotalSpots As Integer
    Dim ilLoop As Integer
    Dim llSpots As Long
    Dim llWkPrice As Long
    Dim ilDay As Integer
    Dim llAvgPrice As Long
    Dim slAvgPrice As String
    Dim slRound As String
    Dim slStr As String
    Dim ilCff As Integer
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    Dim llWkEndDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim llDate As Long
    Dim slDate As String
    Dim llPrice As Long
    Dim slPriceType As String
    'Dim ilUpper As Integer
    Dim ilFirstWk As Integer
    Dim ilSetPropPrice As Integer
    Dim ilLastCffUsed As Integer
    Dim ilCffIndex As Integer
    Dim ilClf As Integer
    Dim ilTestCff As Integer
    Dim ilFound As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim ilDays(0 To 6) As Integer
    Dim llEndbilledDate As Long

    'llTotalPrice = 0
    slTotalPrice = ".00"
    ilTotalSpots = 0
    mCVehOptIndex imLnRowNo, imCntrVefCode, imVpfIndex, imTerminate
    llEndbilledDate = tgClfCntr(imLnRowNo - 1).lEndBilledDate
    If (tgClfCntr(imLnRowNo - 1).iChgBilledPrice) And (Not ilFromFlightArea) Then
        llEndbilledDate = 0
        '1/3/07: Disallow period closed price to be changed
        gUnpackDateLong tgSpf.iRPRP(0), tgSpf.iRPRP(1), llEndbilledDate
    End If
    ilLastCffUsed = LBound(tgCffCntr)
    If slInPrice = lbcPriceType(0).List(2) Then  'Use average rate card price
        If ilFromFlightArea Then
            'For ilLoop = LBound(smAirSave, 2) To UBound(smAirSave, 2) - 1 Step 1
            For ilLoop = imLB1Or2 To UBound(smAirSave, 2) - 1 Step 1
                llDate = gDateValue(smAirSave(1, ilLoop))
                If (llEndbilledDate = 0) Or ((llEndbilledDate <> 0) And (llEndbilledDate < llDate)) Then
                    If imAirSave(1, ilLoop) = 0 Then    'Air week
                        'Determine number of spots
                        llSpots = 0
                        If imAirSave(2, ilLoop) = 0 Then
                            For ilDay = 4 To 10 Step 1
                                llSpots = llSpots + Val(smAirSave(ilDay, ilLoop))
                                If Val(smAirSave(ilDay, ilLoop)) > 0 Then
                                    ilDays(ilDay - 4) = True
                                Else
                                    ilDays(ilDay - 4) = False
                                End If
                            Next ilDay
                        Else
                            llSpots = Val(smAirSave(3, ilLoop))
                            For ilDay = 4 To 10 Step 1
                                If smAirSave(ilDay, ilLoop) = "Y" Then
                                    ilDays(ilDay - 4) = True
                                Else
                                    ilDays(ilDay - 4) = False
                                End If
                            Next ilDay
                        End If
                        llWkPrice = mComputeRCWkPrice(tmRcf.iCode, imLnSave(1, imLnRowNo), smAirSave(1, ilLoop), imCntrVefCode, imLnRowNo)
                        If llWkPrice > 0 Then
                            If smDGSave(13) <> "" Then
                                llOvStartTime = gTimeToLong(smDGSave(13), False)
                                llOvEndTime = gTimeToLong(smDGSave(14), False)
                            Else
                                llOvStartTime = 0
                                llOvEndTime = 0
                            End If
                            slStr = mRCPriceAdjust(tmRcf, llWkPrice, True, imLnRowNo, imDGSave(), llOvStartTime, llOvEndTime, Val(smLnSave(16, imLnRowNo)), llSpots, ilDays(), imLnSave(1, imLnRowNo))
                            smAirSave(12, ilLoop) = gRoundStr(slStr, "1.00", 0)    'gLongToStrDec(llWkPrice, 0) '& ".00"
                            llWkPrice = gStrDecToLong(smAirSave(12, ilLoop), 0)
                        Else
                            smAirSave(12, ilLoop) = ""
                        End If
                        'llTotalPrice = llTotalPrice + llSpots * llWkPrice
                        If smLineType <> "E" Then
                            slWkPrice = gMulStr(gLongToStrDec(llSpots, 0), gLongToStrDec(llWkPrice, 2))
                            slTotalPrice = gAddStr(slTotalPrice, slWkPrice)
                            ilTotalSpots = ilTotalSpots + llSpots
                        Else
                            slWkPrice = gLongToStrDec(llWkPrice, 2)
                            slTotalPrice = gAddStr(slTotalPrice, slWkPrice)
                            ilTotalSpots = ilTotalSpots + 1
                        End If
                    End If
                End If
            Next ilLoop
        Else
            ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
            Do While ilCff <> -1
                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                    'gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llFlStartDate    'Week Start date
                    'gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llFlEndDate    'Week Start date
                    llFlStartDate = tgCffCntr(ilCff).lStartDate
                    llFlEndDate = tgCffCntr(ilCff).lEndDate
                    llDate = llFlStartDate
                    llWkEndDate = llDate + 6
                    If llWkEndDate > llFlEndDate Then
                        llWkEndDate = llFlEndDate
                    End If
                    Do
                        If (llEndbilledDate = 0) Or ((llEndbilledDate <> 0) And (llEndbilledDate < llDate)) Then
                            slDate = Format$(llDate, "m/d/yy")
                            If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                llSpots = 0
                                For ilLoop = gWeekDayLong(llDate) To gWeekDayLong(llWkEndDate) Step 1
                                    llSpots = llSpots + tgCffCntr(ilCff).CffRec.iDay(ilLoop)
                                Next ilLoop
                            Else
                                llSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                            End If
                            For ilLoop = 0 To 6 Step 1
                                If tgCffCntr(ilCff).CffRec.iDay(ilLoop) > 0 Then
                                    ilDays(ilLoop) = True
                                Else
                                    ilDays(ilLoop) = False
                                End If
                            Next ilLoop
                            llWkPrice = mComputeRCWkPrice(tmRcf.iCode, imLnSave(1, imLnRowNo), slDate, imCntrVefCode, imLnRowNo)
                            If llWkPrice > 0 Then
                                slStr = mRCPriceAdjust(tmRcf, llWkPrice, False, imLnRowNo, imDGSave(), 0, 0, Val(smLnSave(16, imLnRowNo)), llSpots, ilDays(), imLnSave(1, imLnRowNo))
                                slStr = gRoundStr(slStr, "1.00", 0)    'gLongToStrDec(llWkPrice, 0) '& ".00"
                                llWkPrice = gStrDecToLong(slStr, 0)
                            End If
                            'llTotalPrice = llTotalPrice + llSpots * llWkPrice
                            If smLineType <> "E" Then
                                slWkPrice = gMulStr(gLongToStrDec(llSpots, 0), gLongToStrDec(llWkPrice, 2))
                                slTotalPrice = gAddStr(slTotalPrice, slWkPrice)
                                ilTotalSpots = ilTotalSpots + llSpots
                            Else
                                slWkPrice = gLongToStrDec(llWkPrice, 2)
                                slTotalPrice = gAddStr(slTotalPrice, slWkPrice)
                                ilTotalSpots = ilTotalSpots + 1
                            End If
                        End If
                        llDate = llDate + 1
                        Do While gWeekDayLong(llDate) <> 0
                            llDate = llDate + 1
                        Loop
                        llWkEndDate = llDate + 6
                        If llWkEndDate > llFlEndDate Then
                            llWkEndDate = llFlEndDate
                        End If
                    Loop While llDate <= llFlEndDate
                End If
                ilCff = tgCffCntr(ilCff).iNextCff
            Loop
        End If
        If ilTotalSpots > 0 Then
            llAvgPrice = gStrDecToInt(gDivStr(slTotalPrice, gIntToStrDec(ilTotalSpots, 0)), 0)   'llTotalPrice / ilTotalSpots
        Else
            llAvgPrice = 0
        End If
        slAvgPrice = Trim$(str$(llAvgPrice)) & ".00"
        If tmRcf.iCode <> 0 Then
            slRound = gLongToStrDec(tmRcf.lRound, 2)
            slAvgPrice = gRoundStr(slAvgPrice, slRound, 2)
        End If
    End If
    If ilFromFlightArea Then
        'For ilLoop = LBound(smAirSave, 2) To UBound(smAirSave, 2) - 1 Step 1
        For ilLoop = imLB1Or2 To UBound(smAirSave, 2) - 1 Step 1
            llDate = gDateValue(smAirSave(1, ilLoop))
            If (llEndbilledDate = 0) Or ((llEndbilledDate <> 0) And (llEndbilledDate < llDate)) Then
                If imAirSave(1, ilLoop) = 0 Then    'Air week
                    If slInPrice = lbcPriceType(0).List(1) Then  'Use rate card price
                        llWkPrice = mComputeRCWkPrice(tmRcf.iCode, imLnSave(1, imLnRowNo), smAirSave(1, ilLoop), imCntrVefCode, imLnRowNo)
                        If llWkPrice > 0 Then
                            llSpots = 0
                            If imAirSave(2, ilLoop) = 0 Then
                                For ilDay = 4 To 10 Step 1
                                    llSpots = llSpots + Val(smAirSave(ilDay, ilLoop))
                                    If Val(smAirSave(ilDay, ilLoop)) > 0 Then
                                        ilDays(ilDay - 4) = True
                                    Else
                                        ilDays(ilDay - 4) = False
                                    End If
                                Next ilDay
                            Else
                                llSpots = Val(smAirSave(3, ilLoop))
                                For ilDay = 4 To 10 Step 1
                                    If smAirSave(ilDay, ilLoop) = "Y" Then
                                        ilDays(ilDay - 4) = True
                                    Else
                                        ilDays(ilDay - 4) = False
                                    End If
                                Next ilDay
                            End If
                            If smDGSave(13) <> "" Then
                                llOvStartTime = gTimeToLong(smDGSave(13), False)
                                llOvEndTime = gTimeToLong(smDGSave(14), False)
                            Else
                                llOvStartTime = 0
                                llOvEndTime = 0
                            End If
                            slStr = mRCPriceAdjust(tmRcf, llWkPrice, True, imLnRowNo, imDGSave(), llOvStartTime, llOvEndTime, Val(smLnSave(16, imLnRowNo)), llSpots, ilDays(), imLnSave(1, imLnRowNo))
                        Else
                            slStr = "0"
                        End If
                        smAirSave(11, ilLoop) = slStr
                        smAirSave(12, ilLoop) = slStr
                    ElseIf slInPrice = lbcPriceType(0).List(2) Then  'Use average rate card price
                        smAirSave(11, ilLoop) = slAvgPrice
                    Else
                        smAirSave(11, ilLoop) = slInPrice
                    End If
                Else
                    smAirSave(11, ilLoop) = ""
                End If
                If smAirSave(11, ilLoop) <> "" Then
                    If (Left$(smAirSave(11, ilLoop), 1) < "0") Or (Left$(smAirSave(11, ilLoop), 1) > "9") Then
                        gSetShow pbcAirWk, smAirSave(11, ilLoop), tmAirCtrls(WKPRICEINDEX)
                        smAirShow(WKPRICEINDEX, ilLoop) = tmAirCtrls(WKPRICEINDEX).sShow
                    Else
                        gFormatStr smAirSave(11, ilLoop), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                        gSetShow pbcAirWk, slStr, tmAirCtrls(WKPRICEINDEX)
                        smAirShow(WKPRICEINDEX, ilLoop) = tmAirCtrls(WKPRICEINDEX).sShow
                    End If
                Else
                    gSetShow pbcAirWk, "", tmAirCtrls(WKPRICEINDEX)
                    smAirShow(WKPRICEINDEX, ilLoop) = tmAirCtrls(WKPRICEINDEX).sShow
                End If
            End If
        Next ilLoop
    Else
       '10843 remove dead code
'        If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType > 1) Then
'            If lmHSStartDate < 0 Then
'                Contract.mSetHdDate True, False
'            End If
'            llSDate = lmHSStartDate + 7 * (imHSCurWkNo - 1)
'            slDate = Format(llSDate, "m/d/yy")
'            gObtainMonthYear 0, slDate, ilCff, ilClf
'            slDate = Trim$(str$(3 * ((ilCff - 1) \ 3) + 1)) & "/15/" & Trim$(str$(ilClf))
'            llFlStartDate = gDateValue(gObtainStartStd(slDate))
'            llEDate = llSDate + 7 * 12 - 1
'            ilCff = (ilCff - 1) \ 3 + 1
'            llSpots = ilCff
'            Do While llSpots = ilCff
'                llEDate = llEDate + 7
'                slDate = Format(llEDate, "m/d/yy")
'                gObtainMonthYear 0, slDate, ilCff, ilClf
'                ilCff = (ilCff - 1) \ 3 + 1
'            Loop
'            llEDate = llEDate - 7
'            llSpots = 0
'        Else
            llSDate = 0
            llEDate = 999999999
       ' End If
        ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
        Do While ilCff <> -1
            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                'gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llFlStartDate    'Week Start date
                'gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llFlEndDate    'Week Start date
                llFlStartDate = tgCffCntr(ilCff).lStartDate
                llFlEndDate = tgCffCntr(ilCff).lEndDate
                llDate = llFlStartDate
                llWkEndDate = llDate + 6
                If llWkEndDate > llFlEndDate Then
                    llWkEndDate = llFlEndDate
                End If
                ilFirstWk = True
                Do
                    If (llDate >= llSDate) And (llDate <= llEDate) Then
                        If (llEndbilledDate = 0) Or ((llEndbilledDate <> 0) And (llEndbilledDate < llDate)) Then
                            slDate = Format$(llDate, "m/d/yy")
                            llWkPrice = mComputeRCWkPrice(tmRcf.iCode, imLnSave(1, imLnRowNo), slDate, imCntrVefCode, imLnRowNo)
                            If llWkPrice > 0 Then
                                For ilDay = 0 To 6 Step 1
                                    ilDays(ilDay) = False
                                Next ilDay
                                If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                    llSpots = 0
                                    For ilDay = 0 To 6 Step 1
                                        If (llDate + ilDay >= llFlStartDate) And (llDate + ilDay <= llFlEndDate) Then
                                            llSpots = llSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                            If tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0 Then
                                                ilDays(ilDay) = True
                                            End If
                                        End If
                                    Next ilDay
                                Else
                                    llSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                                    For ilDay = 0 To 6 Step 1
                                        If (llDate + ilDay >= llFlStartDate) And (llDate + ilDay <= llFlEndDate) Then
                                            If (tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilDay) = "1") Then
                                                ilDays(ilDay) = True
                                            End If
                                        End If
                                    Next ilDay
                                End If
                                slStr = mRCPriceAdjust(tmRcf, llWkPrice, False, imLnRowNo, imDGSave(), 0, 0, Val(smLnSave(16, imLnRowNo)), llSpots, ilDays(), imLnSave(1, imLnRowNo))
                                llWkPrice = gStrDecToLong(slStr, 0)
                            End If
                            If slInPrice = lbcPriceType(0).List(1) Then  'Use rate card price
                                slPriceType = "T"
                                llPrice = 100 * llWkPrice
                            ElseIf slInPrice = lbcPriceType(0).List(2) Then  'Use average rate card price
                                slPriceType = "T"
                                llPrice = gStrDecToLong(slAvgPrice, 2)  'llAvgPrice
                            Else
                                gFindMatch slInPrice, 1, lbcPriceType(0)
                                If gLastFound(lbcPriceType(0)) > 0 Then
                                    lbcPriceType(0).ListIndex = gLastFound(lbcPriceType(0))
                                    llPrice = 0
                                    Select Case lbcPriceType(0).ListIndex
                                        Case 3  'N/C
                                            slPriceType = "N"
                                        Case 4  'MG
                                            slPriceType = "M"
                                        Case 5  'Bonus
                                            slPriceType = "B"
                                        Case 6  'Spinoff
                                            slPriceType = "S"
                                        Case 7  'Package
                                            slPriceType = "P"
                                        Case 8  'Recaptureable
                                            slPriceType = "R"
                                        Case 9  'ADU
                                            slPriceType = "A"
                                    End Select
                                Else
                                    slPriceType = "T"
                                    llPrice = gStrDecToLong(slInPrice, 2)
                                End If
                            End If
                            If ilFirstWk Then
                                ilSetPropPrice = False
                                If tgCffCntr(ilCff).CffRec.sPriceType = "*" Then
                                    tgCffCntr(ilCff).CffRec.lPropPrice = llWkPrice
                                    ilSetPropPrice = True
                                End If
                                tgCffCntr(ilCff).CffRec.sPriceType = slPriceType
                                If smLineType <> "E" Then
                                    tgCffCntr(ilCff).CffRec.lActPrice = llPrice
                                Else
                                    tgCffCntr(ilCff).CffRec.lActPrice = llPrice
                                End If
                                ilFirstWk = False
                            Else
                                If (tgCffCntr(ilCff).CffRec.lPropPrice <> llWkPrice) Or (tgCffCntr(ilCff).CffRec.lActPrice <> llPrice) Then
                                    'Split flight
                                     ilCffIndex = -1
                                    For ilLoop = ilLastCffUsed To UBound(tgCffCntr) - 1 Step 1
                                        If tgCffCntr(ilLoop).iStatus = -1 Then
                                            'Test Chain
                                            ilFound = False
                                            For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) Step 1
                                                ilTestCff = tgClfCntr(ilClf).iFirstCff
                                                If ilTestCff = ilLoop Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                                Do While ilTestCff <> -1
                                                    ilTestCff = tgCffCntr(ilTestCff).iNextCff
                                                    If ilTestCff = ilLoop Then
                                                        ilFound = True
                                                        Exit For
                                                    End If
                                                Loop
                                            Next ilClf
                                            If Not ilFound Then
                                                ilCffIndex = ilLoop
                                                ilLastCffUsed = ilLoop
                                                Exit For
                                            End If
                                        End If
                                    Next ilLoop
                                    If ilCffIndex = -1 Then
                                        ilCffIndex = UBound(tgCffCntr)
                                        ilLastCffUsed = ilCffIndex
                                        ReDim Preserve tgCffCntr(0 To ilCffIndex + 100) As CFFLIST
                                        For ilLoop = ilCffIndex + 1 To UBound(tgCffCntr) Step 1
                                            tgCffCntr(ilLoop).iStatus = -1 'Not Used
                                            tgCffCntr(ilLoop).lRecPos = 0
                                            tgCffCntr(ilLoop).iNextCff = -1
                                        Next ilLoop
                                    End If
                                    tgCffCntr(ilCffIndex) = tgCffCntr(ilCff)
                                    tgCffCntr(ilCffIndex).iStatus = 0
                                    tgCffCntr(ilCff).iNextCff = ilCffIndex
                                    'Set End Date
                                    gPackDateLong llDate - 1, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
                                    tgCffCntr(ilCff).lEndDate = llDate - 1
                                    ilCff = ilCffIndex
                                    gPackDateLong llDate, tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1)
                                    tgCffCntr(ilCff).lStartDate = llDate
                                    If ilSetPropPrice Then
                                        tgCffCntr(ilCff).CffRec.lPropPrice = llWkPrice
                                    End If
                                    tgCffCntr(ilCff).CffRec.sPriceType = slPriceType
                                    If smLineType <> "E" Then
                                        tgCffCntr(ilCff).CffRec.lActPrice = llPrice
                                    Else
                                        tgCffCntr(ilCff).CffRec.lActPrice = llPrice
                                    End If
                                End If
                            End If
                        End If
                    ElseIf llDate > llEDate Then
                        Exit Do
                    End If
                    llDate = llDate + 1
                    Do While gWeekDayLong(llDate) <> 0
                        llDate = llDate + 1
                    Loop
                    llWkEndDate = llDate + 6
                    If llWkEndDate > llFlEndDate Then
                        llWkEndDate = llFlEndDate
                    End If
                Loop While llDate <= llFlEndDate
            End If
            ilCff = tgCffCntr(ilCff).iNextCff
        Loop
       '10843 remove dead code
'        If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType > 1) Then
'            tgClfCntr(imLnRowNo - 1).sCurrentPrice = ""
'        Else
            tgClfCntr(imLnRowNo - 1).sCurrentPrice = slInPrice
'        End If
        'If tgUrf(0).sUseComputeCMC <> "I" Then
            mInitLnTotals imLnRowNo, imLnTotalType, PRICEINDEX
        'End If
    End If
    Exit Sub

End Sub
Private Sub edcDGDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim llWkEndDate As Long
    Dim ilNoWks As Integer
    If imDGChgMode Then
        Exit Sub
    End If
    imDGChgMode = True
    'If (imLnBoxNo <> OV1INDEX) And (imLnBoxNo <> OV2INDEX) Then
        'If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 6) And (imLnSave(8, imLnRowNo) = 0) And (imOVType = 3) And (igFreezeState = 1) And (pbcOVWD(0).Visible) Then
        'If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 6) And (imLnSave(8, imLnRowNo) = 0) And (imOVType = 3) And (pbcOVWD(0).Visible) Then
        '    Exit Sub
        'End If
        Select Case imDGBoxNo
            Case DGSTARTDATEINDEX
                slStr = edcDGDropDown.Text
                If Not gValidDate(slStr) Then
                    imDGChgMode = False
                    mDGSetCommands
                    Exit Sub
                End If
                smDGSave(10) = slStr
                llWkEndDate = gDateValue(smDGSave(10)) + 7 * Val(smDGSave(1)) - 7
                If gValidDate(smDGSave(11)) Then
                    Do Until gWeekDayLong(llWkEndDate) = gWeekDayStr(smDGSave(11))
                        llWkEndDate = llWkEndDate + 1
                    Loop
                End If
                smDGSave(11) = Format$(llWkEndDate, "m/d/yy")
                slDate = gAddDayToDate(smDGSave(11))
                gSetShow pbcDGen, slDate, tmDGCtrls(DGENDDATEINDEX)
                If (gDateValue(smDGSave(10)) >= lmCalStartDate) And (gDateValue(smDGSave(10)) <= lmCalEndDate) Then
                    pbcDGen_Paint
                    mBoxCalDate
                Else
                    pbcDGen.Cls
                    pbcDGen_Paint
                    gObtainMonthYear imCalType, smDGSave(10), imCalMonth, imCalYear
                    pbcCalendar_Paint   'mBoxCalDate called within paint
                End If
            Case DGENDDATEINDEX
                slStr = edcDGDropDown.Text
                If gValidDate(slStr) Then
                    If gDateValue(slStr) < gDateValue(smDGSave(10)) Then
                        imDGChgMode = False
                        mDGSetCommands
                        Exit Sub
                    End If
                Else
                    imDGChgMode = False
                    mDGSetCommands
                    Exit Sub
                End If
                smDGSave(11) = slStr
                ilNoWks = (gDateValue(smDGSave(11)) - gDateValue(smDGSave(10)) + 7) \ 7
                smDGSave(1) = Trim$(str$(ilNoWks))
                slStr = smDGSave(1)
                gSetShow pbcDGen, slStr, tmDGCtrls(DGNOWKSINDEX)
                If (gDateValue(smDGSave(11)) >= lmCalStartDate) And (gDateValue(smDGSave(11)) <= lmCalEndDate) Then
                    pbcDGen_Paint
                    mBoxCalDate
                Else
                    pbcDGen.Cls
                    pbcDGen_Paint
                    gObtainMonthYear imCalType, smDGSave(11), imCalMonth, imCalYear
                    pbcCalendar_Paint   'mBoxCalDate called within paint
                End If
            Case DGRUNINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDGDropDown, lbcDGRun, imBSMode, imComboBoxIndex
                imLbcArrowSetting = False
            Case DGDWPINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDGDropDown, lbcDWP, imBSMode, imComboBoxIndex
                If tgSpf.sAllowDailyBuys <> "Y" Then
                   slStr = edcDGDropDown.Text
                    If InStr(1, slStr, "Daily", 1) > 0 Then
                        edcDGDropDown.Text = "Weekly"   'Force to weekly
                        gMatchLookAhead edcDGDropDown, lbcDWP, imBSMode, imComboBoxIndex
                    End If
                End If
                imLbcArrowSetting = False
            Case DGPRICEINDEX
                imLbcArrowSetting = True
                ilRet = gOptionalLookAhead(edcDGDropDown, lbcPriceType(0), imBSMode, slStr)
                imLbcArrowSetting = False
                If (lbcPriceType(0).ListIndex = 0) Then
                    If Trim$(edcDGDropDown.Text) > "" Then
                        If (Left$(edcDGDropDown.Text, 1) < "0") Or (Left$(edcDGDropDown.Text, 1) > "9") Then
                            If Left$(edcDGDropDown.Text, 1) <> "." Then
                                edcDGDropDown.Text = ""
                            End If
                       End If
                    End If
                End If
            Case DGOVSTARTINDEX
            Case DGOVENDINDEX
        End Select
        mDGSetCommands
    'Else
    '    If imOVType <> 2 Then
    '        Select Case imOVBoxNo
    '            Case OVWDINDEX
    '            Case OVSTARTTIMEINDEX
    '            Case OVENDTIMEINDEX
    '            Case OVRANKINDEX
    '            Case OVBBOPENINDEX
    '                imLbcArrowSetting = True
    '                gMatchLookAhead edcDGDropDown, lbcBBLen, imBSMode, imComboBoxIndex
    '                imLbcArrowSetting = False
    '            Case OVBBCLOSEINDEX
    '                imLbcArrowSetting = True
    '                gMatchLookAhead edcDGDropDown, lbcBBLen, imBSMode, imComboBoxIndex
    '                imLbcArrowSetting = False
    '            Case ACT1LINEUPINDEX
    '        End Select
    '    Else
    '        Select Case imOVBoxNo
    '            Case OV2STARTTIMEINDEX
    '            Case OV2ENDTIMEINDEX
    '            Case OV2SPLITINDEX
    '                imLbcArrowSetting = True
    '                ilRet = gOptionalLookAhead(edcDGDropDown, lbcRegion, imBSMode, slStr)
    '                If ilRet = 1 Then
    '                    lbcRegion.ListIndex = 0
    '                End If
    '            Case OV2ALLOCATIONINDEX
    '            Case OV2ACQUISITIONINDEX
    '        End Select
    '    End If
    'End If
    imDGChgMode = False
End Sub

Private Sub edcDGDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub

Private Sub edcDGDropDown_GotFocus()
    'If (imLnBoxNo <> OV1INDEX) And (imLnBoxNo <> OV2INDEX) Then
        Select Case imDGBoxNo
            Case DGRUNINDEX
                If lbcDGRun.ListCount = 1 Then
                    lbcDGRun.ListIndex = 0
                    If imTabDirection = -1 Then  'Right To Left
                        pbcDGSTab.SetFocus
                    Else
                        pbcDGTab.SetFocus
                    End If
                    Exit Sub
                End If
        End Select
    'End If
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDGDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDGDropDown_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slStr As String
    Dim slMaxValue As String

    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If ActiveControl.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'If (imLnBoxNo <> OV1INDEX) And (imLnBoxNo <> OV2INDEX) Then
        'If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 6) And (imLnSave(8, imLnRowNo) = 0) And (imOVType = 3) And (igFreezeState = 1) And (pbcOVWD(0).Visible) Then
        'If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 6) And (imLnSave(8, imLnRowNo) = 0) And (imOVType = 3) And (pbcOVWD(0).Visible) Then
        '    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        '        Beep
        '        KeyAscii = 0
        '        Exit Sub
        '    End If
        '    slStr = edcDGDropDown.Text
        '    slStr = Left$(slStr, edcDGDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDGDropDown.SelStart - edcDGDropDown.SelLength)
        '    slMaxValue = "999"
        '    If Trim(smLnSave(18, imLnRowNo)) <> "" Then
        '        If gCompNumberStr(smLnSave(18, imLnRowNo), "0") <> 0 Then
        '            slMaxValue = gDivStr("21000000", smLnSave(18, imLnRowNo))
        '            If gCompNumberStr(slMaxValue, "999") > 0 Then
        '                slMaxValue = "999"
        '            End If
        '        End If
        '    End If
        '    If gCompNumberStr(slStr, slMaxValue) > 0 Then
        '        Beep
        '        KeyAscii = 0
        '        Exit Sub
        '    End If
        '    Exit Sub
        'End If
        Select Case imDGBoxNo
            Case DGSTARTDATEINDEX
                'Filter characters (allow only BackSpace, numbers 0 thru 9
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            Case DGENDDATEINDEX
                'Filter characters (allow only BackSpace, numbers 0 thru 9
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            Case DGRUNINDEX
            Case DGDWPINDEX
            Case DGPRICEINDEX
                slStr = edcDGDropDown.Text
                slStr = Left$(slStr, edcDGDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDGDropDown.SelStart - edcDGDropDown.SelLength)
                If (Left$(slStr, 1) < "0") And (Left$(slStr, 1) > "9") Then
                    If (KeyAscii = Asc("R")) Or (KeyAscii = Asc("r")) Then
                        lbcPriceType(0).ListIndex = 1
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If (KeyAscii = Asc("A")) Or (KeyAscii = Asc("a")) Then
                        lbcPriceType(0).ListIndex = 2
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If (KeyAscii = Asc("N")) Or (KeyAscii = Asc("n")) Then
                        lbcPriceType(0).ListIndex = 3
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If (KeyAscii = Asc("M")) Or (KeyAscii = Asc("m")) Then
                        lbcPriceType(0).ListIndex = 4
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If (KeyAscii = Asc("B")) Or (KeyAscii = Asc("b")) Then
                        lbcPriceType(0).ListIndex = 5
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If (KeyAscii = Asc("S")) Or (KeyAscii = Asc("s")) Then
                        lbcPriceType(0).ListIndex = 6
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If (KeyAscii = Asc("P")) Or (KeyAscii = Asc("p")) Then
                        lbcPriceType(0).ListIndex = 7
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If (KeyAscii = Asc("R")) Or (KeyAscii = Asc("r")) Then
                        lbcPriceType(0).ListIndex = 8
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If (KeyAscii = Asc("D")) Or (KeyAscii = Asc("d")) Then
                        lbcPriceType(0).ListIndex = 9
                        KeyAscii = 0
                        Exit Sub
                    End If
                    Exit Sub
                End If
                imChgMode = True
                lbcPriceType(0).ListIndex = 0
                If Trim$(edcDGDropDown.Text) > "" Then
                    If (Left$(edcDGDropDown.Text, 1) < "0") Or (Left$(edcDGDropDown.Text, 1) > "9") Then
                        edcDGDropDown.Text = ""
                    End If
                End If
                imChgMode = False
                ilPos = InStr(edcDGDropDown.SelText, ".")
                If ilPos = 0 Then
                    ilPos = InStr(edcDGDropDown.Text, ".")    'Disallow multi-decimal points
                    If ilPos > 0 Then
                        If KeyAscii = KEYDECPOINT Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                End If
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
                slStr = edcDGDropDown.Text
                slStr = Left$(slStr, edcDGDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDGDropDown.SelStart - edcDGDropDown.SelLength)
                slMaxValue = "9999999.99"
                If Trim(smLnSave(23, imLnRowNo)) <> "" Then
                    If gCompNumberStr(smLnSave(23, imLnRowNo), "0") <> 0 Then
                        slMaxValue = gDivStr("21000000", smLnSave(23, imLnRowNo))
                        If gCompNumberStr(slMaxValue, "9999999.99") > 0 Then
                            slMaxValue = "9999999.99"
                        End If
                    End If
                End If
                If gCompNumberStr(slStr, slMaxValue) > 0 Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            Case DGOVSTARTINDEX
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
                gTimeOutLine KeyAscii, imcTmeOutline
            Case DGOVENDINDEX
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
                gTimeOutLine KeyAscii, imcTmeOutline
        End Select
    'Else
    '    If imOVType <> 2 Then
    '        Select Case imOVBoxNo
    '            Case OVWDINDEX
    '                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
    '                    Beep
    '                    KeyAscii = 0
    '                    Exit Sub
    '                End If
    '                slStr = edcDGDropDown.Text
    '                slStr = Left$(slStr, edcDGDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDGDropDown.SelStart - edcDGDropDown.SelLength)
    '                If gCompNumberStr(slStr, "999") > 0 Then
    '                    Beep
    '                    KeyAscii = 0
    '                    Exit Sub
    '                End If
    '            Case OVSTARTTIMEINDEX
    '                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
    '                    ilFound = False
    '                    For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
    '                        If KeyAscii = igLegalTime(ilLoop) Then
    '                            ilFound = True
    '                            Exit For
    '                        End If
    '                    Next ilLoop
    '                    If Not ilFound Then
    '                        Beep
    '                        KeyAscii = 0
    '                        Exit Sub
    '                    End If
    '                End If
    '                gTimeOutLine KeyAscii, imcTmeOutline
    '            Case OVENDTIMEINDEX
    '                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
    '                    ilFound = False
    '                    For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
    '                        If KeyAscii = igLegalTime(ilLoop) Then
    '                            ilFound = True
    '                            Exit For
    '                        End If
    '                    Next ilLoop
    '                    If Not ilFound Then
    '                        Beep
    '                        KeyAscii = 0
    '                        Exit Sub
    '                    End If
    '                End If
    '                gTimeOutLine KeyAscii, imcTmeOutline
    '            Case OVRANKINDEX
    '                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
    '                    Beep
    '                    KeyAscii = 0
    '                    Exit Sub
    '                End If
    '                slStr = edcDGDropDown.Text
    '                slStr = Left$(slStr, edcDGDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDGDropDown.SelStart - edcDGDropDown.SelLength)
    '                If gCompNumberStr(slStr, "2000") > 0 Then
    '                    Beep
    '                    KeyAscii = 0
    '                    Exit Sub
    '                End If
    '            Case OVBBOPENINDEX
    '            Case OVBBCLOSEINDEX
    '            Case ACT1LINEUPINDEX
    '        End Select
    '    Else
    '        Select Case imOVBoxNo
    '            Case OV2STARTTIMEINDEX
    '                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
    '                    ilFound = False
    '                    For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
    '                        If KeyAscii = igLegalTime(ilLoop) Then
    '                            ilFound = True
    '                            Exit For
    '                        End If
    '                    Next ilLoop
    '                    If Not ilFound Then
    '                        Beep
    '                        KeyAscii = 0
    '                        Exit Sub
    '                    End If
    '                End If
    '                gTimeOutLine KeyAscii, imcTmeOutline
    '            Case OV2ENDTIMEINDEX
    '                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
    '                    ilFound = False
    '                    For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
    '                        If KeyAscii = igLegalTime(ilLoop) Then
    '                            ilFound = True
    '                            Exit For
    '                        End If
    '                    Next ilLoop
    '                    If Not ilFound Then
    '                        Beep
    '                        KeyAscii = 0
    '                        Exit Sub
    '                    End If
    '                End If
    '                gTimeOutLine KeyAscii, imcTmeOutline
    '            Case OV2SPLITINDEX
    '            Case OV2ALLOCATIONINDEX
    '                ilPos = InStr(edcDGDropDown.SelText, ".")
    '                If ilPos = 0 Then
    '                    ilPos = InStr(edcDGDropDown.Text, ".")    'Disallow multi-decimal points
    '                    If ilPos > 0 Then
    '                        If KeyAscii = KEYDECPOINT Then
    '                            Beep
    '                            KeyAscii = 0
    '                            Exit Sub
    '                        End If
    '                    End If
    '                End If
    '                'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    '                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
    '                    Beep
    '                    KeyAscii = 0
    '                    Exit Sub
    '                End If
    '                slStr = edcDGDropDown.Text
    '                slStr = Left$(slStr, edcDGDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDGDropDown.SelStart - edcDGDropDown.SelLength)
    '                If gCompNumberStr(slStr, "100.00") > 0 Then
    '                    Beep
    '                    KeyAscii = 0
    '                    Exit Sub
    '                End If
    '            Case OV2ACQUISITIONINDEX
    '                ilPos = InStr(edcDGDropDown.SelText, ".")
    '                If ilPos = 0 Then
    '                    ilPos = InStr(edcDGDropDown.Text, ".")    'Disallow multi-decimal points
    '                    If ilPos > 0 Then
    '                        If KeyAscii = KEYDECPOINT Then
    '                            Beep
    '                            KeyAscii = 0
    '                            Exit Sub
    '                        End If
    '                    End If
    '                End If
    '                'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    '                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
    '                    Beep
    '                    KeyAscii = 0
    '                    Exit Sub
    '                End If
    '                slStr = edcDGDropDown.Text
    '                slStr = Left$(slStr, edcDGDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDGDropDown.SelStart - edcDGDropDown.SelLength)
    '                slMaxValue = "9999999.99"
    '                If Trim(smLnSave(23, imLnRowNo)) <> "" Then
    '                    If gCompNumberStr(smLnSave(23, imLnRowNo), "0") <> 0 Then
    '                        slMaxValue = gDivStr("21000000", smLnSave(23, imLnRowNo))
    '                        If gCompNumberStr(slMaxValue, "9999999.99") > 0 Then
    '                            slMaxValue = "9999999.99"
    '                        End If
    '                    End If
    '                End If
    '                If gCompNumberStr(slStr, slMaxValue) > 0 Then
    '                    Beep
    '                    KeyAscii = 0
    '                    Exit Sub
    '                End If
    '        End Select
    '    End If
    'End If
End Sub
Private Sub edcDGDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        'If (imLnBoxNo <> OV1INDEX) And (imLnBoxNo <> OV2INDEX) Then
            'If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 6) And (imLnSave(8, imLnRowNo) = 0) And (imOVType = 3) And (igFreezeState = 1) And (pbcOVWD(0).Visible) Then
            'If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 6) And (imLnSave(8, imLnRowNo) = 0) And (imOVType = 3) And (pbcOVWD(0).Visible) Then
            '    edcDGDropDown.SelStart = 0
            '    edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
            '    Exit Sub
            'End If
            Select Case imDGBoxNo
                Case DGSTARTDATEINDEX
                    If (Shift And vbAltMask) > 0 Then
                        plcCalendar.Visible = Not plcCalendar.Visible
                    Else
                        If gValidDate(smDGSave(10)) Then
                            If KeyCode = KEYUP Then 'Up arrow
                                smDGSave(10) = Format$(gDateValue(smDGSave(10)) - 7, "m/d/yy")
                            Else
                                smDGSave(10) = Format$(gDateValue(smDGSave(10)) + 7, "m/d/yy")
                            End If
                            gObtainMonthYear imCalType, smDGSave(10), imCalMonth, imCalYear
                            edcDGDropDown.Text = smDGSave(10)
                        End If
                    End If
                Case DGENDDATEINDEX
                    If (Shift And vbAltMask) > 0 Then
                        plcCalendar.Visible = Not plcCalendar.Visible
                    Else
                        If gValidDate(smDGSave(11)) Then
                            If KeyCode = KEYUP Then 'Up arrow
                                smDGSave(11) = Format$(gDateValue(smDGSave(11)) - 7, "m/d/yy")
                            Else
                                smDGSave(11) = Format$(gDateValue(smDGSave(11)) + 7, "m/d/yy")
                            End If
                            gObtainMonthYear imCalType, smDGSave(11), imCalMonth, imCalYear
                            edcDGDropDown.Text = smDGSave(11)
                        End If
                    End If
                Case DGRUNINDEX
                    gProcessArrowKey Shift, KeyCode, lbcDGRun, imLbcArrowSetting
                Case DGDWPINDEX
                    If tgSpf.sAllowDailyBuys = "Y" Then
                        gProcessArrowKey Shift, KeyCode, lbcDWP, imLbcArrowSetting
                    End If
                Case DGPRICEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcPriceType(0), imLbcArrowSetting
                Case DGOVSTARTINDEX
                    If (Shift And vbAltMask) > 0 Then
                        plcTme.Visible = Not plcTme.Visible
                    End If
                Case DGOVENDINDEX
                    If (Shift And vbAltMask) > 0 Then
                        plcTme.Visible = Not plcTme.Visible
                    End If
            End Select
        'Else
        '    If imOVType <> 2 Then
        '        Select Case imOVBoxNo
        '            Case OVSTARTTIMEINDEX
        '                If (Shift And vbAltMask) > 0 Then
        '                    plcTme.Visible = Not plcTme.Visible
        '                End If
        '            Case OVENDTIMEINDEX
        '                If (Shift And vbAltMask) > 0 Then
        '                    plcTme.Visible = Not plcTme.Visible
        '                End If
        '            Case OVBBOPENINDEX
        '                gProcessArrowKey Shift, KeyCode, lbcBBLen, imLbcArrowSetting
        '            Case OVBBCLOSEINDEX
        '                gProcessArrowKey Shift, KeyCode, lbcBBLen, imLbcArrowSetting
        '        End Select
        '    Else
        '    Select Case imOVBoxNo
        '        Case OV2STARTTIMEINDEX
        '            If (Shift And vbAltMask) > 0 Then
        '                plcTme.Visible = Not plcTme.Visible
        '            End If
        '        Case OV2ENDTIMEINDEX
        '            If (Shift And vbAltMask) > 0 Then
        '                plcTme.Visible = Not plcTme.Visible
        '            End If
        '        Case OV2SPLITINDEX
        '            gProcessArrowKey Shift, KeyCode, lbcRegion, imLbcArrowSetting
        '        Case OV2ALLOCATIONINDEX
        '        Case OV2ACQUISITIONINDEX
        '    End Select
        '    End If
        'End If
        edcDGDropDown.SelStart = 0
        edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        'If (imLnBoxNo <> OV1INDEX) And (imLnBoxNo <> OV2INDEX) Then
            'If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 6) And (imLnSave(8, imLnRowNo) = 0) And (imOVType = 3) And (igFreezeState = 1) And (pbcOVWD(0).Visible) Then
            'If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 6) And (imLnSave(8, imLnRowNo) = 0) And (imOVType = 3) And (pbcOVWD(0).Visible) Then
            '    edcDGDropDown.SelStart = 0
            '    edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
            '    Exit Sub
            'End If
            Select Case imDGBoxNo
                Case DGSTARTDATEINDEX
                    If (Shift And vbAltMask) > 0 Then
                    Else
                        If gValidDate(smDGSave(10)) Then
                            If KeyCode = KEYLEFT Then 'Left arrow
                                smDGSave(10) = Format$(gDateValue(smDGSave(10)) - 1, "m/d/yy")
                            Else
                                smDGSave(10) = Format$(gDateValue(smDGSave(10)) + 1, "m/d/yy")
                            End If
                            gObtainMonthYear imCalType, smDGSave(10), imCalMonth, imCalYear
                            edcDGDropDown.Text = smDGSave(10)
                        End If
                    End If
                Case DGENDDATEINDEX
                    If (Shift And vbAltMask) > 0 Then
                    Else
                        If gValidDate(smDGSave(11)) Then
                            If KeyCode = KEYLEFT Then 'Left arrow
                                smDGSave(11) = Format$(gDateValue(smDGSave(11)) - 1, "m/d/yy")
                            Else
                                smDGSave(11) = Format$(gDateValue(smDGSave(11)) + 1, "m/d/yy")
                            End If
                            gObtainMonthYear imCalType, smDGSave(11), imCalMonth, imCalYear
                            edcDGDropDown.Text = smDGSave(11)
                        End If
                    End If
                Case DGRUNINDEX
            End Select
        'Else
        '    Select Case imOVBoxNo
        '    End Select
        'End If
        edcDGDropDown.SelStart = 0
        edcDGDropDown.SelLength = Len(edcDGDropDown.Text)
    End If
End Sub

Private Sub edcDGDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        'If (imLnBoxNo <> OV1INDEX) And (imLnBoxNo <> OV2INDEX) Then
        'Else
        '    If imOVType <> 2 Then
        '    Else
        '        Select Case imOVBoxNo
        '            Case OV2SPLITINDEX
        '                If imTabDirection = -1 Then  'Right To Left
        '                    pbcLnSTab.SetFocus
        '                Else
        '                    pbcLnTab.SetFocus
        '                End If
        '            Exit Sub
        '        End Select
        '    End If
        'End If
        imDoubleClickName = False
    End If
End Sub
