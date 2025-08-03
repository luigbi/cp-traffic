VERSION 5.00
Begin VB.Form RCTerms 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5910
   ClientLeft      =   240
   ClientTop       =   1935
   ClientWidth     =   9285
   ClipControls    =   0   'False
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5910
   ScaleWidth      =   9285
   Begin VB.PictureBox pbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1305
      Left            =   510
      Picture         =   "Rcterms.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   7050
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5745
      Visible         =   0   'False
      Width           =   7080
   End
   Begin VB.PictureBox pbcTDDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2820
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3705
      Visible         =   0   'False
      Width           =   210
      Begin VB.CheckBox ckcTDDay 
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
         Left            =   15
         TabIndex        =   28
         Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.PictureBox pbcSTab 
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
      Height          =   90
      Left            =   90
      ScaleHeight     =   90
      ScaleWidth      =   75
      TabIndex        =   14
      Top             =   345
      Width           =   75
   End
   Begin VB.PictureBox pbcTab 
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
      Height          =   90
      Left            =   135
      ScaleHeight     =   90
      ScaleWidth      =   75
      TabIndex        =   36
      Top             =   420
      Width           =   75
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   5730
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3825
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "Rcterms.frx":1D516
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   33
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
            Picture         =   "Rcterms.frx":1E1D4
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
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
      Left            =   5340
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   1995
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Rcterms.frx":1E4DE
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   9
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
            TabIndex        =   10
            Top             =   405
            Visible         =   0   'False
            Width           =   300
         End
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
         TabIndex        =   41
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmcSpecDropDown 
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
      Left            =   4230
      Picture         =   "Rcterms.frx":212F8
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcBySpotType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   645
      ScaleHeight     =   210
      ScaleWidth      =   1005
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2205
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox edcSpecDropDown 
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
      Left            =   3105
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmcDropDown 
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
      Left            =   2850
      Picture         =   "Rcterms.frx":213F2
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropDown 
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
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2205
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.ListBox lbcBaseLen 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   7605
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4680
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.PictureBox pbcSpotType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   345
      Picture         =   "Rcterms.frx":214EC
      ScaleHeight     =   1395
      ScaleWidth      =   3510
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   615
      Width           =   3510
   End
   Begin VB.PictureBox pbcGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   345
      Picture         =   "Rcterms.frx":31926
      ScaleHeight     =   2760
      ScaleWidth      =   1545
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2730
      Width           =   1545
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
      Height          =   90
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   105
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5715
      Width           =   105
   End
   Begin VB.ListBox lbcCombo 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   40
      Top             =   105
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox pbcArrow 
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
      Height          =   180
      Left            =   165
      Picture         =   "Rcterms.frx":364E0
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3510
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcSpecTab 
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
      Height          =   60
      Left            =   1815
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   13
      Top             =   420
      Width           =   60
   End
   Begin VB.PictureBox pbcSpecSTab 
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
      Height          =   105
      Left            =   1785
      ScaleHeight     =   105
      ScaleWidth      =   90
      TabIndex        =   1
      Top             =   15
      Width           =   90
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   60
      ScaleHeight     =   240
      ScaleWidth      =   1470
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1470
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      HelpContextID   =   2
      Left            =   5025
      TabIndex        =   38
      Top             =   5625
      Width           =   945
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      HelpContextID   =   1
      Left            =   3315
      TabIndex        =   37
      Top             =   5625
      Width           =   945
   End
   Begin VB.PictureBox pbcSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1965
      Picture         =   "Rcterms.frx":367EA
      ScaleHeight     =   375
      ScaleWidth      =   7185
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   7185
   End
   Begin VB.PictureBox plcSpec 
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
      Height          =   480
      Left            =   1905
      ScaleHeight     =   420
      ScaleWidth      =   7260
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   15
      Width           =   7320
   End
   Begin VB.PictureBox pbcLength 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   4425
      Picture         =   "Rcterms.frx":3993C
      ScaleHeight     =   2385
      ScaleWidth      =   1830
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   615
      Width           =   1830
      Begin VB.Label lacLength 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Index"
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
         Height          =   195
         Left            =   780
         TabIndex        =   25
         Top             =   210
         Width           =   1035
      End
   End
   Begin VB.PictureBox pbcSpotFreq 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   7020
      Picture         =   "Rcterms.frx":3E6D2
      ScaleHeight     =   1215
      ScaleWidth      =   2145
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   615
      Width           =   2145
      Begin VB.Label lacSpotFreq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Index"
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
         Height          =   195
         Left            =   1110
         TabIndex        =   19
         Top             =   225
         Width           =   1005
      End
   End
   Begin VB.PictureBox pbcWeekFreq 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   6990
      Picture         =   "Rcterms.frx":414F4
      ScaleHeight     =   1215
      ScaleWidth      =   2145
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   2145
      Begin VB.Label lacWeekFreq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Index"
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
         Height          =   195
         Left            =   1110
         TabIndex        =   21
         Top             =   210
         Width           =   1005
      End
   End
   Begin VB.PictureBox pbcDayRate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   2115
      Picture         =   "Rcterms.frx":44316
      ScaleHeight     =   2385
      ScaleWidth      =   2640
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2640
      Begin VB.Label lacDayRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Index"
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
         Height          =   195
         Left            =   1590
         TabIndex        =   29
         Top             =   210
         Width           =   1035
      End
   End
   Begin VB.PictureBox pbcHourRate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   4980
      Picture         =   "Rcterms.frx":4B0F8
      ScaleHeight     =   2385
      ScaleWidth      =   4170
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   4170
      Begin VB.Label lacHourRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Index"
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
         Height          =   195
         Left            =   3135
         TabIndex        =   31
         Top             =   225
         Width           =   1005
      End
   End
   Begin VB.PictureBox plcTerms 
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
      Height          =   5010
      Left            =   60
      ScaleHeight     =   4950
      ScaleWidth      =   9075
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   570
      Width           =   9135
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   225
      Picture         =   "Rcterms.frx":55F72
      Top             =   285
      Width           =   480
   End
End
Attribute VB_Name = "RCTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rcterms.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RCTerms.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Rate Card Terms input screen code
'
'   This program uses tmRcf for storing the values tgRcfI is the image retained in
'   the rate card program.  Using tmRcf eliminates the save area.
'
Option Explicit
Option Compare Text
Dim tmSpecCtrls(0 To 7)  As FIELDAREA
Dim imLBSpecCtrls As Integer
Dim imSpecBoxNo As Integer   'Current Rate Card Term Box
Dim tmSTCtrls(0 To 12) As FIELDAREA  'Spot Type (By, Value: 3 times)
Dim imLBSTCtrls As Integer
Dim tmSFCtrls(0 To 12) As FIELDAREA 'Spot Freq (From, To, Value: 4 Times)
Dim imLBSFCtrls As Integer
Dim tmWFCtrls(0 To 12) As FIELDAREA 'Week Freq (From, To, Value: 4 Times)
Dim imLBWFCtrls As Integer
Dim tmLenCtrls(0 To 20)  As FIELDAREA
Dim imLBLenCtrls As Integer
Dim tmGDCtrls(0 To 12) As FIELDAREA  'Grid price
Dim imLBGDCtrls As Integer
Dim tmDRCtrls(0 To 80) As FIELDAREA 'Day Rate
Dim imLBDRCtrls As Integer
Dim tmHRCtrls(0 To 100) As FIELDAREA
Dim imLBHRCtrls As Integer
Dim imType As Integer       'Which one has focus (SPOTTYPE; SPOTFREQ; WEEKFREQ; LENGTH; GRID; DAYRATE; HOURRATE)
Dim imBoxNo As Integer
'Dim imLenBoxNo As Integer
'Dim imLenRowNo As Integer
'Dim imMaxNoLen As Integer
'Dim imSpotBoxNo As Integer
'Dim imSpotRowNo As Integer
'Dim imMaxNoSpot As Integer
'Dim imWeekBoxNo As Integer
'Dim imWeekRowNo As Integer
'Dim imMaxNoWeek As Integer
'Dim imSingleBoxNo As Integer
'Dim imFixedBoxNo As Integer
'Dim imNonBoxNo As Integer
'Dim imDayBoxNo As Integer
'Dim imDayRowNo As Integer
'Dim imHourBoxNo As Integer
'Dim imHourRowNo As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imEditType As Integer   '0=Ascii; 1=Numbers with decimal; 2=Numbers without decimal; 3= Date; 4=Time
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim tmRcf As RCF    'Working image
Dim imComboBoxIndex As Integer
Dim imLbcArrowSetting As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imVpfIndex As Integer       '-1=All vehicles; >0 vehicle selected
Dim imBypassFocus As Integer
Dim imVehIndex As Integer   'Save vehicle index
Dim imUpdateAllowed As Integer
Dim smSign5 As String * 5
Dim smZero5 As String * 5
Dim smSign3 As String * 3
Dim smZero3 As String * 3
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
'Mouse down
Const NAMEINDEX = 1     'Rate Card Name control/field
Const VEHICLEINDEX = 2
Const YEARINDEX = 3   'Start date control/field
Const STARTDATEINDEX = 4
Const NOGRIDSINDEX = 5  'Number of Grids control/index
Const ROUNDINDEX = 6    'Round to nearest control/index
Const BASELENINDEX = 7  'Base Length control/field
Const STBYINDEX = 1
Const STVALUEINDEX = 2
Const SPOTFROMINDEX = 1 'Spot frequency "From" control/field
Const SPOTTOINDEX = 2   'Spot frequency "TO" control/field
Const SPOTVALUEINDEX = 3 'Spot frequency value control/field
Const WEEKFROMINDEX = 1 'Week frequency "From" control/field
Const WEEKTOINDEX = 2   'Week frequency "TO" control/field
Const WEEKVALUEINDEX = 3 'Week frequency value control/field
Const LENINDEX = 1 'Spot length value control/field
Const LENVALUEINDEX = 2 'Spot length value control/field
Const GRIDVALUEINDEX = 1 'Grid value control/field
Const DRDAYINDEX = 1 'Day value control/field
Const DRVALUEINDEX = 8
Const HRSTARTINDEX = 1
Const HRENDINDEX = 2
Const HRDAYINDEX = 3
Const HRVALUEINDEX = 10
Const SPOTTYPE = 1
Const Length = 2
Const SPOTFREQ = 3
Const WEEKFREQ = 4
Const GRID = 5
Const DAYRATE = 6
Const HOURRATE = 7
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcSpecDropDown.SelStart = 0
    edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    edcSpecDropDown.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcSpecDropDown.SelStart = 0
    edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    edcSpecDropDown.SetFocus
End Sub
Private Sub cmcCancel_Click()
    igRCReturn = 0
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
    mSetAll
End Sub
Private Sub cmcDone_Click()
    Dim tlRcf1 As MOVEREC
    Dim tlRcf2 As MOVEREC
    Dim ilLoop As Integer
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If mTestFields() = NO Then
        mSpecEnableBox imSpecBoxNo
        Exit Sub
    End If
    If igRCMode = 0 Then    'New-set level
        tmRcf.iTodayGrid = 1
    End If
    'If (StrComp(tmRcf.sBBValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sBBValue, smZero5, 0) <> 0) Then
    'Else
    If tmRcf.lBBValue = 0 Then
        tmRcf.sUseBB = "N"
    End If
    'If (StrComp(tmRcf.sFPValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sFPValue, smZero5, 0) <> 0) Then
    'Else
    If tmRcf.lFPValue = 0 Then
        tmRcf.sUseFP = "N"
    End If
    If tmRcf.lNPValue = 0 Then
        tmRcf.sUseNP = "N"
    End If
    If tmRcf.lPrefDTValue = 0 Then
        tmRcf.sUsePrefDT = "N"
    End If
    If tmRcf.l1stPosValue = 0 Then
        tmRcf.sUse1stPos = "N"
    End If
    If tmRcf.lSoloAvailValue = 0 Then
        tmRcf.sUseSoloAvail = "N"
    End If
    tmRcf.sUseLen = "N"
    For ilLoop = LBound(tmRcf.iLen) To UBound(tmRcf.iLen) Step 1
        'If (StrComp(tmRcf.sValue(ilLoop), smSign5, 0) <> 0) And (StrComp(tmRcf.sValue(ilLoop), smZero5, 0) <> 0) Then
        If tmRcf.lValue(ilLoop) <> 0 Then
            If lacLength.Caption = "Index" Then
                tmRcf.sUseLen = "I"
            Else
                tmRcf.sUseLen = "D"
            End If
            Exit For
        End If
    Next ilLoop

    tmRcf.sUseSpot = "N"
    For ilLoop = LBound(tmRcf.iSpotMin) To UBound(tmRcf.iSpotMin) Step 1
        'If (StrComp(tmRcf.sSpotVal(ilLoop), smSign5, 0) <> 0) And (StrComp(tmRcf.sSpotVal(ilLoop), smZero5, 0) <> 0) Then
        If tmRcf.lSpotVal(ilLoop) <> 0 Then
            If lacSpotFreq.Caption = "Index" Then
                tmRcf.sUseSpot = "I"
            Else
                tmRcf.sUseSpot = "D"
            End If
            Exit For
        End If
    Next ilLoop

    tmRcf.sUseWeek = "N"
    If pbcWeekFreq.Visible Then
        For ilLoop = LBound(tmRcf.iWkMin) To UBound(tmRcf.iWkMin) Step 1
            'If (StrComp(tmRcf.sWkVal(ilLoop), smSign5, 0) <> 0) And (StrComp(tmRcf.sWkVal(ilLoop), smZero5, 0) <> 0) Then
            If tmRcf.lWkVal(ilLoop) <> 0 Then
                If lacWeekFreq.Caption = "Index" Then
                    tmRcf.sUseWeek = "I"
                Else
                    tmRcf.sUseWeek = "D"
                End If
                Exit For
            End If
        Next ilLoop
    End If

    tmRcf.sUseDay = "N"
    For ilLoop = LBound(tmRcf.lDyRate) To UBound(tmRcf.lDyRate) Step 1
        'If (StrComp(tmRcf.sDyRate(ilLoop), smSign5, 0) <> 0) And (StrComp(tmRcf.sDyRate(ilLoop), smZero5, 0) <> 0) Then
        If tmRcf.lDyRate(ilLoop) <> 0 Then
            If lacDayRate.Caption = "Index" Then
                tmRcf.sUseDay = "I"
            Else
                tmRcf.sUseDay = "D"
            End If
            Exit For
        End If
    Next ilLoop


    tmRcf.sUseHour = "N"
    If pbcHourRate.Visible Then
        For ilLoop = LBound(tmRcf.lHrRate) To UBound(tmRcf.lHrRate) Step 1
            'If (StrComp(tmRcf.sHrRate(ilLoop), smSign5, 0) <> 0) And (StrComp(tmRcf.sHrRate(ilLoop), smZero5, 0) <> 0) Then
            If tmRcf.lHrRate(ilLoop) <> 0 Then
                If lacHourRate.Caption = "Index" Then
                    tmRcf.sUseHour = "I"
                Else
                    tmRcf.sUseHour = "D"
                End If
                Exit For
            End If
        Next ilLoop
    End If
    LSet tlRcf1 = tgRcfI
    LSet tlRcf2 = tmRcf
    If StrComp(tlRcf1.sChar, tlRcf2.sChar, 0) <> 0 Then
        igRcfChg = True
        tgRcfI = tmRcf
    End If
    igRCReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
    mSetAll
End Sub
Private Sub cmcDropDown_Click()
    Dim ilBox As Integer
    Dim ilIndex As Integer
    If imType = HOURRATE Then
        ilIndex = (imBoxNo - 1) \ HRVALUEINDEX + 1
        ilBox = (imBoxNo - 1) Mod HRVALUEINDEX + 1
        Select Case ilBox
            Case 1
                plcTme.Visible = Not plcTme.Visible
            Case 2
                plcTme.Visible = Not plcTme.Visible
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
        edcDropDown.SetFocus
    End If
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcSpecDropDown_Click()
    Select Case imSpecBoxNo
        Case VEHICLEINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
        Case STARTDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case BASELENINDEX
            lbcBaseLen.Visible = Not lbcBaseLen.Visible
    End Select
    edcSpecDropDown.SelStart = 0
    edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    edcSpecDropDown.SetFocus
End Sub
Private Sub cmcSpecDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    If imEditType = 0 Then  'Any characters
        Exit Sub
    End If
    If imEditType = 1 Then  'one decimal place only (money or percent)
        ilPos = InStr(edcDropDown.SelText, ".")
        If ilPos = 0 Then
            ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
                If ilPos > 0 Then
                    If KeyAscii = KEYDECPOINT Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
        End If
        'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) And (KeyAscii <> KEYPOS) And (KeyAscii <> KEYNEG) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        Exit Sub
        slStr = edcDropDown.Text
        slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
        If gCompNumberStr(slStr, "9999999.99") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    If imEditType = 2 Then  'Number without decimal point
        'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYPOS) And (KeyAscii <> KEYNEG) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        Exit Sub
        slStr = edcDropDown.Text
        slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
        'If imSpecBoxNo = NOGRIDSINDEX Then
        '    If gCompNumberStr(slStr, "12") > 0 Then
        '        Beep
        '        KeyAscii = 0
        '        Exit Sub
        '    End If
        'Else
            If gCompNumberStr(slStr, "999") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        'End If
    End If
    If imEditType = 3 Then  'Date
        'Filter characters (allow only BackSpace, numbers 0 thru 9, slash  (Date)
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        Exit Sub
    End If
    If imEditType = 4 Then  'Time
        'Filter characters (allow only BackSpace, numbers 0 thru 9, slash  (Date)
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
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilBox As Integer
    Dim ilIndex As Integer
    If imType = HOURRATE Then
        ilIndex = (imBoxNo - 1) \ HRVALUEINDEX + 1
        ilBox = (imBoxNo - 1) Mod HRVALUEINDEX + 1
        If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
            Select Case ilBox
                Case 1
                    If (Shift And vbAltMask) > 0 Then
                        plcTme.Visible = Not plcTme.Visible
                    End If
                    edcDropDown.SelStart = 0
                    edcDropDown.SelLength = Len(edcDropDown.Text)
                Case 2
                    If (Shift And vbAltMask) > 0 Then
                        plcTme.Visible = Not plcTme.Visible
                    End If
                    edcDropDown.SelStart = 0
                    edcDropDown.SelLength = Len(edcDropDown.Text)
            End Select
        End If
    End If
End Sub
Private Sub edcSpecDropDown_Change()
    Dim slStr As String
    Select Case imSpecBoxNo
        Case VEHICLEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcSpecDropDown, lbcVehicle, imBSMode, imComboBoxIndex
            If lbcVehicle.ListIndex <> imVehIndex Then
                'mModelFromNewestRC  'Only call if changed vehicle
                imVehIndex = lbcVehicle.ListIndex
            End If
        Case STARTDATEINDEX
            slStr = edcSpecDropDown.Text
            If Not gValidDate(slStr) Then
                Exit Sub
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case BASELENINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcSpecDropDown, lbcBaseLen, imBSMode, imComboBoxIndex
    End Select
    imLbcArrowSetting = False
End Sub
Private Sub edcSpecDropDown_GotFocus()
    Select Case imSpecBoxNo
        Case VEHICLEINDEX
            If lbcVehicle.ListCount = 1 Then
                lbcVehicle.ListIndex = 0
                'If imTabDirection = -1 Then 'Right to left
                '    pbcSpecSTab.SetFocus
                'Else
                '    pbcSpecTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case STARTDATEINDEX
        Case BASELENINDEX
            If lbcBaseLen.ListCount = 1 Then
                lbcBaseLen.ListIndex = 0
                'If imTabDirection = -1 Then 'Right to left
                '    pbcSpecSTab.SetFocus
                'Else
                '    pbcSpecTab.SetFocus
                'End If
                'Exit Sub
            End If
    End Select
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcSpecDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcSpecDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcSpecDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case imSpecBoxNo
        Case NAMEINDEX 'Name
            If (KeyAscii = KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case VEHICLEINDEX
        Case STARTDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case BASELENINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub
Private Sub edcSpecDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imSpecBoxNo
            Case VEHICLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
            Case STARTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcSpecDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpecDropDown.Text = slDate
                    End If
                End If
            Case BASELENINDEX
                gProcessArrowKey Shift, KeyCode, lbcBaseLen, imLbcArrowSetting
        End Select
        edcSpecDropDown.SelStart = 0
        edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imSpecBoxNo
            Case VEHICLEINDEX
            Case STARTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcSpecDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpecDropDown.Text = slDate
                    End If
                End If
                edcSpecDropDown.SelStart = 0
                edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            Case BASELENINDEX
        End Select
    End If
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Or ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB) And (tgUrf(0).iMnfHubCode > 0) Then
        pbcSpec.Enabled = False
        pbcSpotType.Enabled = False
        pbcSpotFreq.Enabled = False
        pbcWeekFreq.Enabled = False
        pbcLength.Enabled = False
        pbcGrid.Enabled = False
        pbcDayRate.Enabled = False
        pbcHourRate.Enabled = False
        pbcSpecSTab.Enabled = False
        pbcSpecTab.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    RCTerms.Refresh
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
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
        If imSpecBoxNo > 0 Then
            mSpecEnableBox imSpecBoxNo
        ElseIf imBoxNo > 0 Then
            mEnableBox
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not igManUnload Then
        mSpecSetShow imSpecBoxNo
    End If
    Set RCTerms = Nothing   'Remove data segment
End Sub

Private Sub imcKey_Click()
    pbcKey.Visible = Not pbcKey.Visible
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pbcKey.Visible = True
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pbcKey.Visible = False
End Sub

Private Sub lacDayRate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    If tmRcf.sUseDay = "I" Then
        tmRcf.sUseDay = "D"
        lacDayRate.Caption = "Dollar"
    Else
        tmRcf.sUseDay = "I"
        lacDayRate.Caption = "Index"
    End If
    For ilLoop = LBound(tmRcf.lDyRate) To UBound(tmRcf.lDyRate) Step 1
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmRcf.sDyRate(ilLoop)
        tmRcf.lDyRate(ilLoop) = 0
        'tmDRCtrls(8 * ilLoop).sShow = ""
        tmDRCtrls(8 * (ilLoop + 1)).sShow = ""
    Next ilLoop
    pbcDayRate.Cls
    pbcDayRate_Paint
End Sub
Private Sub lacHourRate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    If tmRcf.sUseHour = "I" Then
        tmRcf.sUseHour = "D"
        lacHourRate.Caption = "Dollar"
    Else
        tmRcf.sUseHour = "I"
        lacHourRate.Caption = "Index"
    End If
    For ilLoop = LBound(tmRcf.lHrRate) To UBound(tmRcf.lHrRate) Step 1
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmRcf.sHrRate(ilLoop)
        tmRcf.lHrRate(ilLoop) = 0
        tmHRCtrls(10 * (ilLoop + 1)).sShow = ""
    Next ilLoop
    pbcHourRate.Cls
    pbcHourRate_Paint
End Sub
Private Sub lacLength_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    If tmRcf.sUseLen = "I" Then
        tmRcf.sUseLen = "D"
        lacLength.Caption = "Dollar"
    Else
        tmRcf.sUseLen = "I"
        lacLength.Caption = "Index"
    End If
    For ilLoop = LBound(tmRcf.iLen) To UBound(tmRcf.iLen) Step 1
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmRcf.sValue(ilLoop)
        tmRcf.lValue(ilLoop) = 0
        tmLenCtrls(2 * (ilLoop + 1)).sShow = ""
    Next ilLoop
    pbcLength.Cls
    pbcLength_Paint
End Sub
Private Sub lacSpotFreq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer

    If tmRcf.sUseSpot = "I" Then
        tmRcf.sUseSpot = "D"
        lacSpotFreq.Caption = "Dollar"
    Else
        tmRcf.sUseSpot = "I"
        lacSpotFreq.Caption = "Index"
    End If
    For ilLoop = LBound(tmRcf.iSpotMin) To UBound(tmRcf.iSpotMin) Step 1
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmRcf.sSpotVal(ilLoop)
        tmRcf.lSpotVal(ilLoop) = 0
        tmSFCtrls(3 * (ilLoop + 1)).sShow = ""
    Next ilLoop
    pbcSpotFreq.Cls
    pbcSpotFreq_Paint
End Sub
Private Sub lacWeekFreq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    If tmRcf.sUseWeek = "I" Then
        tmRcf.sUseWeek = "D"
        lacWeekFreq.Caption = "Dollar"
    Else
        tmRcf.sUseWeek = "I"
        lacWeekFreq.Caption = "Index"
    End If
    For ilLoop = LBound(tmRcf.iWkMin) To UBound(tmRcf.iWkMin) Step 1
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmRcf.sWkVal(ilLoop)
        tmRcf.lWkVal(ilLoop) = 0
        tmWFCtrls(3 * (ilLoop + 1)).sShow = ""
    Next ilLoop
    pbcWeekFreq.Cls
    pbcWeekFreq_Paint
End Sub
Private Sub lbcBaseLen_Click()
    gProcessLbcClick lbcBaseLen, edcSpecDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcBaseLen_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcVehicle_Click()
    gProcessLbcClick lbcVehicle, edcSpecDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
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
    slStr = edcSpecDropDown.Text
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
'*      Procedure Name:mClearRec                       *
'*                                                     *
'*             Created:7/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up a New rate card         *
'*                                                     *
'*******************************************************
Private Sub mClearRec()
    Dim ilLoop As Integer
    Dim ilDay As Integer
    tmRcf.iCode = 0     'set only in mClearRec
    tmRcf.sName = ""
    tmRcf.iVefCode = -32000
    tmRcf.iStartDate(0) = 0
    tmRcf.iStartDate(1) = 0
    tmRcf.iEndDate(0) = 0
    tmRcf.iEndDate(1) = 0
    tmRcf.iGridsUsed = 0
    tmRcf.iTodayGrid = 0
    tmRcf.iBaseLen = 0
    'slStr = ""
    'gStrToPDN slStr, 2, 3, tmRcf.sRound
    tmRcf.lRound = 0
    tmRcf.sUseBB = "N"
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tmRcf.sBBValue
    tmRcf.lBBValue = 0
    tmRcf.sUseFP = "N"
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tmRcf.sFPValue
    tmRcf.lFPValue = 0
    tmRcf.sUseNP = "N"
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tmRcf.sNPValue
    tmRcf.lNPValue = 0

    tmRcf.sUsePrefDT = "N"
    tmRcf.lPrefDTValue = 0
    tmRcf.sUse1stPos = "N"
    tmRcf.l1stPosValue = 0
    tmRcf.sUseSoloAvail = "N"
    tmRcf.lSoloAvailValue = 0

    tmRcf.sUseLen = "N"
    For ilLoop = LBound(tmRcf.iLen) To UBound(tmRcf.iLen) Step 1
        tmRcf.iLen(ilLoop) = 0  'Lengths set in mLenPop
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmRcf.sValue(ilLoop)
        tmRcf.lValue(ilLoop) = 0
    Next ilLoop
    tmRcf.sUseSpot = "N"
    For ilLoop = LBound(tmRcf.iSpotMin) To UBound(tmRcf.iSpotMin) Step 1
        tmRcf.iSpotMin(ilLoop) = 0
        tmRcf.iSpotMax(ilLoop) = 0
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmRcf.sSpotVal(ilLoop)
        tmRcf.lSpotVal(ilLoop) = 0
    Next ilLoop
    tmRcf.sUseWeek = "N"
    For ilLoop = LBound(tmRcf.iWkMin) To UBound(tmRcf.iWkMin) Step 1
        tmRcf.iWkMin(ilLoop) = 0
        tmRcf.iWkMax(ilLoop) = 0
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmRcf.sWkVal(ilLoop)
        tmRcf.lWkVal(ilLoop) = 0
    Next ilLoop
    tmRcf.sUseDay = "N"
    For ilLoop = LBound(tmRcf.lDyRate) To UBound(tmRcf.lDyRate) Step 1
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmRcf.sDyRate(ilLoop)
        tmRcf.lDyRate(ilLoop) = 0
        For ilDay = 1 To 7 Step 1
            tmRcf.sDay(ilLoop, ilDay - 1) = "N"
        Next ilDay
    Next ilLoop
    tmRcf.sUseHour = "N"
    For ilLoop = LBound(tmRcf.lHrRate) To UBound(tmRcf.lHrRate) Step 1
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmRcf.sHrRate(ilLoop)
        tmRcf.lHrRate(ilLoop) = 0
        For ilDay = 1 To 7 Step 1
            tmRcf.sHrDay(ilLoop, ilDay - 1) = "N"
        Next ilDay
        tmRcf.iHrStartTime(0, ilLoop) = 0
        tmRcf.iHrStartTime(1, ilLoop) = 0
        tmRcf.iHrEndTime(0, ilLoop) = 0
        tmRcf.iHrEndTime(1, ilLoop) = 0
    Next ilLoop
    For ilLoop = LBound(tmRcf.iFltNo) To UBound(tmRcf.iFltNo) Step 1
        tmRcf.iFltNo(ilLoop) = 1
    Next ilLoop
    For ilLoop = LBound(tmRcf.lGridIndex) To UBound(tmRcf.lGridIndex) Step 1
        'slStr = ""
        'gStrToPDN slStr, 4, 3, tmRcf.sGridIndex(ilLoop)
        tmRcf.lGridIndex(ilLoop) = 0
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:7/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox()
'
'   mSpotEnableBox imBoxNo
'   Where:
'       imBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilBox As Integer
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    If imType <= 0 Then
        Exit Sub
    End If
    Select Case imType
        Case SPOTTYPE
            If (imBoxNo < imLBSTCtrls) Or (imBoxNo > UBound(tmSTCtrls)) Then
                Exit Sub
            End If
            pbcArrow.Move pbcSpotType.Left - pbcArrow.Width - 15, pbcSpotType.Top + tmSTCtrls(imBoxNo).fBoxY - 15
            pbcArrow.Visible = True
            edcDropDown.Width = tmSTCtrls(imBoxNo).fBoxW
            pbcBySpotType.Width = tmSTCtrls(imBoxNo).fBoxW
            ilIndex = (imBoxNo - 1) \ STVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod STVALUEINDEX + 1
            Select Case ilBox
                Case 1    'Spot Type By Dollar or Index
                    gMoveTableCtrl pbcSpotType, pbcBySpotType, tmSTCtrls(imBoxNo).fBoxX, tmSTCtrls(imBoxNo).fBoxY
                    pbcBySpotType.Visible = True
                    pbcBySpotType.SetFocus
                Case 2  'From Value
                    gMoveTableCtrl pbcSpotType, edcDropDown, tmSTCtrls(imBoxNo).fBoxX, tmSTCtrls(imBoxNo).fBoxY
                    Select Case ilIndex
                        Case 1
                            If tmRcf.sUseBB <> "I" Then
                                edcDropDown.MaxLength = 10
                                imEditType = 1
                            Else
                                edcDropDown.MaxLength = 6
                                imEditType = 1
                            End If
                            'gPDNToStr tmRcf.sBBValue, 2, slStr
                            slStr = gLongToStrDec(tmRcf.lBBValue, 2)
                        Case 2
                            If tmRcf.sUseFP <> "I" Then
                                edcDropDown.MaxLength = 10
                                imEditType = 1
                            Else
                                edcDropDown.MaxLength = 6
                                imEditType = 1
                            End If
                            'gPDNToStr tmRcf.sFPValue, 2, slStr
                            slStr = gLongToStrDec(tmRcf.lFPValue, 2)
                        Case 3
                            If tmRcf.sUseNP <> "I" Then
                                edcDropDown.MaxLength = 10
                                imEditType = 1
                            Else
                                edcDropDown.MaxLength = 6
                                imEditType = 1
                            End If
                            'gPDNToStr tmRcf.sNPValue, 2, slStr
                            slStr = gLongToStrDec(tmRcf.lNPValue, 2)
                        Case 4
                            If tmRcf.sUsePrefDT <> "I" Then
                                edcDropDown.MaxLength = 10
                                imEditType = 1
                            Else
                                edcDropDown.MaxLength = 6
                                imEditType = 1
                            End If
                            'gPDNToStr tmRcf.sBBValue, 2, slStr
                            slStr = gLongToStrDec(tmRcf.lPrefDTValue, 2)
                        Case 5
                            If tmRcf.sUse1stPos <> "I" Then
                                edcDropDown.MaxLength = 10
                                imEditType = 1
                            Else
                                edcDropDown.MaxLength = 6
                                imEditType = 1
                            End If
                            'gPDNToStr tmRcf.sFPValue, 2, slStr
                            slStr = gLongToStrDec(tmRcf.l1stPosValue, 2)
                        Case 6
                            If tmRcf.sUseSoloAvail <> "I" Then
                                edcDropDown.MaxLength = 10
                                imEditType = 1
                            Else
                                edcDropDown.MaxLength = 6
                                imEditType = 1
                            End If
                            'gPDNToStr tmRcf.sNPValue, 2, slStr
                            slStr = gLongToStrDec(tmRcf.lSoloAvailValue, 2)
                    End Select
                    edcDropDown.Text = slStr
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
            End Select
        Case SPOTFREQ
            If (imBoxNo < imLBSFCtrls) Or (imBoxNo > UBound(tmSFCtrls)) Then
                Exit Sub
            End If
            pbcArrow.Move pbcSpotFreq.Left - pbcArrow.Width - 15, pbcSpotFreq.Top + tmSFCtrls(imBoxNo).fBoxY - 15
            pbcArrow.Visible = True
            edcDropDown.Width = tmSFCtrls(imBoxNo).fBoxW
            ilIndex = (imBoxNo - 1) \ SPOTVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod SPOTVALUEINDEX + 1
            Select Case ilBox
                Case 1    'From Value
                    gMoveTableCtrl pbcSpotFreq, edcDropDown, tmSFCtrls(imBoxNo).fBoxX, tmSFCtrls(imBoxNo).fBoxY
                    edcDropDown.MaxLength = 3
                    imEditType = 2
                    If tmRcf.iSpotMin(ilIndex - 1) <> 0 Then
                        slStr = Trim$(str$(tmRcf.iSpotMin(ilIndex - 1)))
                    Else
                        If igRCMode = 0 Then    'New
                            If ilIndex > 1 Then
                                'If tmRcf.iSpotMax(ilIndex - 1) > 0 Then
                                '    slStr = Trim$(str$(tmRcf.iSpotMax(ilIndex - 1) + 1))
                                If tmRcf.iSpotMax(ilIndex - 2) > 0 Then
                                    slStr = Trim$(str$(tmRcf.iSpotMax(ilIndex - 2) + 1))
                                Else
                                    slStr = ""
                                End If
                            Else
                                slStr = ""
                            End If
                        Else
                            slStr = ""
                        End If
                    End If
                    edcDropDown.Text = slStr
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
                Case 2   'From Value
                    gMoveTableCtrl pbcSpotFreq, edcDropDown, tmSFCtrls(imBoxNo).fBoxX, tmSFCtrls(imBoxNo).fBoxY
                    edcDropDown.MaxLength = 3
                    imEditType = 2
                    If tmRcf.iSpotMin(ilIndex - 1) <> 0 Then
                        If tmRcf.iSpotMax(ilIndex - 1) <> 0 Then
                            slStr = Trim$(str$(tmRcf.iSpotMax(ilIndex - 1)))
                        Else
                            slStr = "+"
                        End If
                    Else
                        slStr = ""
                    End If
                    edcDropDown.Text = slStr
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
                Case 3  'From Value
                    gMoveTableCtrl pbcSpotFreq, edcDropDown, tmSFCtrls(imBoxNo).fBoxX, tmSFCtrls(imBoxNo).fBoxY
                    If lacSpotFreq.Caption <> "Index" Then
                        edcDropDown.MaxLength = 10
                        imEditType = 1
                    Else
                        edcDropDown.MaxLength = 6
                        imEditType = 1
                    End If
                    If tmRcf.iSpotMin(ilIndex - 1) <> 0 Then
                        'gPDNToStr tmRcf.sSpotVal(ilIndex), 2, slStr
                        slStr = gLongToStrDec(tmRcf.lSpotVal(ilIndex - 1), 2)
                        edcDropDown.Text = slStr
                    Else
                        edcDropDown.Text = ""
                    End If
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
            End Select
        Case WEEKFREQ
            If (imBoxNo < imLBWFCtrls) Or (imBoxNo > UBound(tmWFCtrls)) Then
                Exit Sub
            End If
            pbcArrow.Move pbcWeekFreq.Left - pbcArrow.Width - 15, pbcWeekFreq.Top + tmWFCtrls(imBoxNo).fBoxY - 15
            pbcArrow.Visible = True
            edcDropDown.Width = tmWFCtrls(imBoxNo).fBoxW
            ilIndex = (imBoxNo - 1) \ WEEKVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod WEEKVALUEINDEX + 1
            Select Case ilBox 'Branch on box type (control)
                Case 1 'From Value
                    gMoveTableCtrl pbcWeekFreq, edcDropDown, tmWFCtrls(imBoxNo).fBoxX, tmWFCtrls(imBoxNo).fBoxY
                    edcDropDown.MaxLength = 3
                    imEditType = 2
                    If tmRcf.iWkMin(ilIndex - 1) <> 0 Then
                        slStr = Trim$(str$(tmRcf.iWkMin(ilIndex - 1)))
                    Else
                        If igRCMode = 0 Then    'New
                            If ilIndex > 1 Then
                                'If tmRcf.iWkMax(ilIndex - 1) > 0 Then
                                '    slStr = Trim$(str$(tmRcf.iWkMax(ilIndex - 1) + 1))
                                If tmRcf.iWkMax(ilIndex - 2) > 0 Then
                                    slStr = Trim$(str$(tmRcf.iWkMax(ilIndex - 2) + 1))
                                Else
                                    slStr = ""
                                End If
                            Else
                                slStr = ""
                            End If
                        Else
                            slStr = ""
                        End If
                    End If
                    edcDropDown.Text = slStr
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
                Case 2  'From Value
                    gMoveTableCtrl pbcWeekFreq, edcDropDown, tmWFCtrls(imBoxNo).fBoxX, tmWFCtrls(imBoxNo).fBoxY
                    If tmRcf.iWkMin(ilIndex - 1) <> 0 Then
                        If tmRcf.iWkMax(ilIndex - 1) <> 0 Then
                            slStr = Trim$(str$(tmRcf.iWkMax(ilIndex - 1)))
                        Else
                            slStr = "+"
                        End If
                    Else
                        slStr = ""
                    End If
                    edcDropDown.Text = slStr
                    edcDropDown.MaxLength = 3
                    imEditType = 2
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
                Case 3  'Value
                    gMoveTableCtrl pbcWeekFreq, edcDropDown, tmWFCtrls(imBoxNo).fBoxX, tmWFCtrls(imBoxNo).fBoxY
                    If lacWeekFreq.Caption <> "Index" Then
                        edcDropDown.MaxLength = 10
                        imEditType = 1
                    Else
                        edcDropDown.MaxLength = 6
                        imEditType = 1
                    End If
                    If tmRcf.iWkMin(ilIndex - 1) <> 0 Then
                        'gPDNToStr tmRcf.sWkVal(ilIndex), 2, slStr
                        slStr = gLongToStrDec(tmRcf.lWkVal(ilIndex - 1), 2)
                        edcDropDown.Text = slStr
                    Else
                        edcDropDown.Text = ""
                    End If
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
            End Select
        Case Length
            If (imBoxNo < imLBLenCtrls) Or (imBoxNo > UBound(tmLenCtrls)) Then
                Exit Sub
            End If
            If (tmLenCtrls(1).sShow = "") And (igRCMode = 0) Then
                ilBoxNo = 1
                For ilRowNo = LBound(tmRcf.lValue) To UBound(tmRcf.lValue)
                    If (tmRcf.iLen(ilRowNo) <> 0) Then
                        slStr = Trim$(str$(tmRcf.iLen(ilRowNo)))
                        gSetShow pbcLength, slStr, tmLenCtrls(ilBoxNo)
                        ilBoxNo = ilBoxNo + 1
                    Else
                        slStr = ""
                        gSetShow pbcLength, slStr, tmLenCtrls(ilBoxNo)
                        ilBoxNo = ilBoxNo + 1
                        slStr = ""
                    End If
                    ilBoxNo = ilBoxNo + 1
                Next ilRowNo
                pbcLength.Cls
                pbcLength_Paint
            End If
            pbcArrow.Move pbcLength.Left - pbcArrow.Width - 15, pbcLength.Top + tmLenCtrls(imBoxNo).fBoxY - 15
            pbcArrow.Visible = True
            edcDropDown.Width = tmLenCtrls(imBoxNo).fBoxW
            ilIndex = (imBoxNo - 1) \ LENVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod LENVALUEINDEX + 1
            If lacLength.Caption <> "Index" Then
                edcDropDown.MaxLength = 10
                imEditType = 1
            Else
                edcDropDown.MaxLength = 6
                imEditType = 1
            End If
            Select Case ilBox
                Case 2 'Value
                    gMoveTableCtrl pbcLength, edcDropDown, tmLenCtrls(imBoxNo).fBoxX, tmLenCtrls(imBoxNo).fBoxY
                    'gPDNToStr tmRcf.sValue(ilIndex), 2, slStr
                    slStr = gLongToStrDec(tmRcf.lValue(ilIndex - 1), 2)
                    edcDropDown.Text = slStr
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
            End Select
        Case GRID
            If (imBoxNo < imLBGDCtrls) Or (imBoxNo > UBound(tmGDCtrls)) Then
                Exit Sub
            End If
            pbcArrow.Move pbcGrid.Left - pbcArrow.Width - 15, pbcGrid.Top + tmGDCtrls(imBoxNo).fBoxY - 15
            pbcArrow.Visible = True
            edcDropDown.MaxLength = 6
            imEditType = 1
            ilIndex = imBoxNo
            Select Case imBoxNo
                Case GRIDVALUEINDEX To GRIDVALUEINDEX + 11 'Value
                    gMoveTableCtrl pbcGrid, edcDropDown, tmGDCtrls(imBoxNo).fBoxX, tmGDCtrls(imBoxNo).fBoxY
                    'gPDNToStr tmRcf.sGridIndex(ilIndex), 4, slStr
                    slStr = gLongToStrDec(tmRcf.lGridIndex(ilIndex - 1), 4)
                    edcDropDown.Text = slStr
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
            End Select
        Case DAYRATE
            If (imBoxNo < imLBDRCtrls) Or (imBoxNo > UBound(tmDRCtrls)) Then
                Exit Sub
            End If
            pbcArrow.Move pbcDayRate.Left - pbcArrow.Width - 15, pbcDayRate.Top + tmDRCtrls(imBoxNo).fBoxY - 15
            pbcArrow.Visible = True
            edcDropDown.Width = tmDRCtrls(imBoxNo).fBoxW
            ilIndex = (imBoxNo - 1) \ DRVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod DRVALUEINDEX + 1
            Select Case ilBox
                Case 1, 2, 3, 4, 5, 6, 7  'Mo-Su
                    gMoveTableCtrl pbcDayRate, pbcTDDays, tmDRCtrls(imBoxNo).fBoxX, tmDRCtrls(imBoxNo).fBoxY
                    If tmRcf.sDay(ilIndex - 1, ilBox - 1) = "Y" Then
                        ckcTDDay.Value = vbChecked
                    Else
                        ckcTDDay.Value = vbUnchecked
                    End If
                    pbcTDDays.Visible = True
                    ckcTDDay.SetFocus
                Case 8  'Value
                    gMoveTableCtrl pbcDayRate, edcDropDown, tmDRCtrls(imBoxNo).fBoxX, tmDRCtrls(imBoxNo).fBoxY
                    If lacDayRate.Caption <> "Index" Then
                        edcDropDown.MaxLength = 10
                        imEditType = 1
                    Else
                        edcDropDown.MaxLength = 6
                        imEditType = 1
                    End If
                    'gPDNToStr tmRcf.sDyRate(ilIndex), 2, slStr
                    slStr = gLongToStrDec(tmRcf.lDyRate(ilIndex - 1), 2)
                    edcDropDown.Text = slStr
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
            End Select
        Case HOURRATE
            If (imBoxNo < imLBHRCtrls) Or (imBoxNo > UBound(tmHRCtrls)) Then
                Exit Sub
            End If
            pbcArrow.Move pbcHourRate.Left - pbcArrow.Width - 15, pbcHourRate.Top + tmHRCtrls(imBoxNo).fBoxY - 15
            pbcArrow.Visible = True
            edcDropDown.Width = tmHRCtrls(imBoxNo).fBoxW
            ilIndex = (imBoxNo - 1) \ HRVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod HRVALUEINDEX + 1
            Select Case ilBox
                Case 1
                    edcDropDown.MaxLength = 10
                    imEditType = 4
                    gMoveTableCtrl pbcHourRate, edcDropDown, tmHRCtrls(imBoxNo).fBoxX, tmHRCtrls(imBoxNo).fBoxY
                    cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                    If ilIndex <= 2 Then
                        plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                    Else
                        plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
                    End If
                    'If (tmRcf.iHrStartTime(0, ilIndex) <> 0) Or (tmRcf.iHrStartTime(1, ilIndex) <> 0) Then
                        gUnpackTime tmRcf.iHrStartTime(0, ilIndex - 1), tmRcf.iHrStartTime(1, ilIndex - 1), "A", "1", slStr
                    'Else
                    '    slStr = ""
                    'End If
                    edcDropDown.Text = slStr
                    plcTme.Visible = False
                    edcDropDown.Visible = True  'Set visibility
                    cmcDropDown.Visible = True
                    edcDropDown.SetFocus
                Case 2
                    edcDropDown.MaxLength = 10
                    imEditType = 4
                    gMoveTableCtrl pbcHourRate, edcDropDown, tmHRCtrls(imBoxNo).fBoxX - cmcDropDown.Width, tmHRCtrls(imBoxNo).fBoxY '+ (imTDRowNo - 1) * (fgBoxGridH + 15)
                    cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                    If ilIndex <= 2 Then
                        plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                    Else
                        plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
                    End If
                    'If (tmRcf.iHrEndTime(0, ilIndex) <> 0) Or (tmRcf.iHrEndTime(1, ilIndex) <> 0) Then
                        gUnpackTime tmRcf.iHrEndTime(0, ilIndex - 1), tmRcf.iHrEndTime(1, ilIndex - 1), "A", "1", slStr
                    'Else
                    '    slStr = ""
                    'End If
                    edcDropDown.Text = slStr
                    plcTme.Visible = False
                    edcDropDown.Visible = True  'Set visibility
                    cmcDropDown.Visible = True
                    edcDropDown.SetFocus
                Case 3, 4, 5, 6, 7, 8, 9  'Mo-Su
                    gMoveTableCtrl pbcHourRate, pbcTDDays, tmHRCtrls(imBoxNo).fBoxX, tmHRCtrls(imBoxNo).fBoxY
                    'If tmRcf.sHrDay(ilIndex, ilBox - 2) = "Y" Then
                    If tmRcf.sHrDay(ilIndex - 1, ilBox - 3) = "Y" Then
                        ckcTDDay.Value = vbChecked
                    Else
                        ckcTDDay.Value = vbUnchecked
                    End If
                    pbcTDDays.Visible = True
                    ckcTDDay.SetFocus
                Case 10  'Value
                    gMoveTableCtrl pbcHourRate, edcDropDown, tmHRCtrls(imBoxNo).fBoxX, tmHRCtrls(imBoxNo).fBoxY
                    If lacHourRate.Caption <> "Index" Then
                        edcDropDown.MaxLength = 10
                        imEditType = 1
                    Else
                        edcDropDown.MaxLength = 6
                        imEditType = 1
                    End If
                    'gPDNToStr tmRcf.sHrRate(ilIndex), 2, slStr
                    slStr = gLongToStrDec(tmRcf.lHrRate(ilIndex - 1), 2)
                    edcDropDown.Text = slStr
                    edcDropDown.Visible = True  'Set visibility
                    edcDropDown.SetFocus
            End Select
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:7/10/93       By:D. LeVine      *
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
    Dim hlVsf As Integer
    Dim ilDestroy As Integer
    Dim slStr As String
    Dim slZero6 As String * 6
    Screen.MousePointer = vbHourglass
    imLBSpecCtrls = 1
    imLBSTCtrls = 1
    imLBSFCtrls = 1
    imLBWFCtrls = 1
    imLBLenCtrls = 1
    imLBGDCtrls = 1
    imLBDRCtrls = 1
    imLBHRCtrls = 1
    imFirstActivate = True
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    ilDestroy = False
    imSpecBoxNo = -1 'Initialize current Box to N/A
    imType = -1
    imBoxNo = -1
    slStr = "0"
    gStrToPDN slStr, 2, 6, slZero6
    slStr = "0"
    gStrToPDN slStr, 2, 5, smSign5
    smZero5 = slZero6   'Remove sign bits
    'gStrToPDN slStr, 2, 5, smZero5
    slStr = "0"
    gStrToPDN slStr, 4, 3, smSign3
    smZero3 = slZero6   'Remove sign bits
    'gStrToPDN slStr, 4, 3, smZero3
    'imLenBoxNo = -1
    'imLenRowNo = 0
    'imSpotBoxNo = -1
    'imSpotRowNo = 0
    'imWeekBoxNo = -1
    'imWeekRowNo = 0
    'imSingleBoxNo = -1
    'imFixedBoxNo = -1
    'imNonBoxNo = -1
    'imDayBoxNo = -1
    'imDayRowNo = 0
    'imHourBoxNo = -1
    'imHourRowNo = 0
    'imVehIndex = -2
    imTerminate = False
    imChgMode = False
    imBSMode = False
    imBypassFocus = False
    imcKey.Picture = IconTraf!imcKey.Picture
    imTabDirection = 0  'left to right
    imLbcArrowSetting = False
    imCalType = 0   'Standard
    If igRCMode = 0 Then 'New
        If igRcfModel > 0 Then
            mModelRcf igRcfModel
        Else
            mClearRec
        End If
        If RateCard!cbcSelect.ListIndex <> 0 Then   'Not [New]
            tmRcf.sName = RateCard!cbcSelect.Text
        End If
    Else
        tmRcf = tgRcfI
    End If
    lbcVehicle.Clear 'Force population
    mVehPop lbcVehicle
    If imTerminate Then
        Exit Sub
    End If
    'mLenPop
    'If imTerminate Then
    '    Exit Sub
    'End If
    Screen.MousePointer = vbHourglass
    If igRCMode = 1 Then 'Change
    Else
        'imMaxNoLen = 0
        'For ilLoop = LBound(tmRcf.iLen) To UBound(tmRcf.iLen) Step 1
        '    If tmRcf.iLen(ilLoop) = 0 Then
        '        Exit For
        '    End If
        ''    imMaxNoLen = imMaxNoLen + 1
        'Next ilLoop
        'imMaxNoSpot = UBound(tmRcf.iSpotMin) - 1
        'imMaxNoWeek = UBound(tmRcf.iWkMin) - 1
    End If
    RCTerms.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterModalForm RCTerms
    lacLength.Enabled = False
    lacSpotFreq.Enabled = False
    lacDayRate.Enabled = False
    lacHourRate.Enabled = False
    lacWeekFreq.Enabled = False
    pbcBySpotType.Enabled = False
    mInitBox
    mInitShow True
    Screen.MousePointer = vbDefault
    Exit Sub

    On Error GoTo 0
    If ilDestroy Then
        ilRet = btrClose(hlVsf)
        btrDestroy hlVsf
    End If
    igRCReturn = 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:7/10/93       By:D. LeVine      *
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
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
    flTextHeight = pbcSpec.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcSpec.Move 1905, 30, pbcSpec.Width + fgPanelAdj, pbcSpec.Height + fgPanelAdj
    pbcSpec.Move plcSpec.Left + fgBevelX, plcSpec.Top + fgBevelY
    plcTerms.Move 90, 555, 9135, 5010
    pbcSpotType.Move 345, 615
    pbcLength.Move 4425, 615
    pbcSpotFreq.Move 7020, 615
    pbcWeekFreq.Move 7020, 1860
    pbcGrid.Move 345, 2730
    pbcDayRate.Move 2115, 3120
    pbcHourRate.Move 4980, 3120

    pbcKey.Move plcTerms.Left, plcTerms.Top

    'Name
    gSetCtrl tmSpecCtrls(NAMEINDEX), 30, 30, 930, fgBoxStH
    'Vehicle
    gSetCtrl tmSpecCtrls(VEHICLEINDEX), 975, tmSpecCtrls(NAMEINDEX).fBoxY, 1500, fgBoxStH
    'Year
    gSetCtrl tmSpecCtrls(YEARINDEX), 2490, tmSpecCtrls(NAMEINDEX).fBoxY, 585, fgBoxStH
    'Start Date
    gSetCtrl tmSpecCtrls(STARTDATEINDEX), 3090, tmSpecCtrls(NAMEINDEX).fBoxY, 930, fgBoxStH
    'Grid Levels
    gSetCtrl tmSpecCtrls(NOGRIDSINDEX), 4035, tmSpecCtrls(NAMEINDEX).fBoxY, 930, fgBoxStH
    'Round to Nearest
    gSetCtrl tmSpecCtrls(ROUNDINDEX), 4980, tmSpecCtrls(NAMEINDEX).fBoxY, 1230, fgBoxStH
    tmSpecCtrls(ROUNDINDEX).iReq = False
    'Base Length
    gSetCtrl tmSpecCtrls(BASELENINDEX), 6225, tmSpecCtrls(NAMEINDEX).fBoxY, 930, fgBoxStH
'    gMoveFormCtrl pbcTerms, cbcLen, tmSpecCtrls(BASELENINDEX).fBoxX, tmSpecCtrls(BASELENINDEX).fBoxY
    tmSpecCtrls(BASELENINDEX).iReq = False
    'Spot Type
    For ilLoop = 1 To 6 Step 1
        gSetCtrl tmSTCtrls(STBYINDEX + 2 * (ilLoop - 1)), 1470, 225 + (ilLoop - 1) * (fgBoxGridH + 15), 1005, fgBoxGridH
        tmSTCtrls(STBYINDEX + 2 * (ilLoop - 1)).iReq = False
        gSetCtrl tmSTCtrls(STVALUEINDEX + 2 * (ilLoop - 1)), 2490, 225 + (ilLoop - 1) * (fgBoxGridH + 15), 1005, fgBoxGridH
        tmSTCtrls(STVALUEINDEX + 2 * (ilLoop - 1)).iReq = False
    Next ilLoop
    'Spot freq and value
    For ilLoop = 1 To 4 Step 1
        gSetCtrl tmSFCtrls(SPOTFROMINDEX + 3 * (ilLoop - 1)), 30, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 525, fgBoxGridH
        tmSFCtrls(SPOTFROMINDEX + 3 * (ilLoop - 1)).iReq = False
        gSetCtrl tmSFCtrls(SPOTTOINDEX + 3 * (ilLoop - 1)), 570, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 525, fgBoxGridH
        tmSFCtrls(SPOTTOINDEX + 3 * (ilLoop - 1)).iReq = False
        gSetCtrl tmSFCtrls(SPOTVALUEINDEX + 3 * (ilLoop - 1)), 1110, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 1020, fgBoxGridH
        tmSFCtrls(SPOTVALUEINDEX + 3 * (ilLoop - 1)).iReq = False
    Next ilLoop
    'Week freq and value
    For ilLoop = 1 To 4 Step 1
        gSetCtrl tmWFCtrls(WEEKFROMINDEX + 3 * (ilLoop - 1)), 30, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 525, fgBoxGridH
        tmWFCtrls(WEEKFROMINDEX + 3 * (ilLoop - 1)).iReq = False
        gSetCtrl tmWFCtrls(WEEKTOINDEX + 3 * (ilLoop - 1)), 570, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 525, fgBoxGridH
        tmWFCtrls(WEEKTOINDEX + 3 * (ilLoop - 1)).iReq = False
        gSetCtrl tmWFCtrls(WEEKVALUEINDEX + 3 * (ilLoop - 1)), 1110, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 1020, fgBoxGridH
        tmWFCtrls(WEEKVALUEINDEX + 3 * (ilLoop - 1)).iReq = False
    Next ilLoop
    'Lengths
    For ilLoop = 1 To 10 Step 1
        gSetCtrl tmLenCtrls(LENINDEX + 2 * (ilLoop - 1)), 30, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 750, fgBoxGridH
        tmLenCtrls(LENINDEX + 2 * (ilLoop - 1)).iReq = False
        gSetCtrl tmLenCtrls(LENVALUEINDEX + 2 * (ilLoop - 1)), 795, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 1005, fgBoxGridH
        tmLenCtrls(LENVALUEINDEX + 2 * (ilLoop - 1)).iReq = False
    Next ilLoop
    'Grid
    For ilLoop = 1 To 12 Step 1
        gSetCtrl tmGDCtrls(GRIDVALUEINDEX + ilLoop - 1), 510, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 1005, fgBoxGridH
        tmGDCtrls(GRIDVALUEINDEX + ilLoop - 1).iReq = False
    Next ilLoop
    'Day Rate
    For ilLoop = 1 To 10 Step 1
        gSetCtrl tmDRCtrls(DRDAYINDEX + 8 * (ilLoop - 1)), 30, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmDRCtrls(DRDAYINDEX + 8 * (ilLoop - 1)).iReq = False
        gSetCtrl tmDRCtrls(DRDAYINDEX + 1 + 8 * (ilLoop - 1)), 255, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmDRCtrls(DRDAYINDEX + 1 + 8 * (ilLoop - 1)).iReq = False
        gSetCtrl tmDRCtrls(DRDAYINDEX + 2 + 8 * (ilLoop - 1)), 480, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmDRCtrls(DRDAYINDEX + 2 + 8 * (ilLoop - 1)).iReq = False
        gSetCtrl tmDRCtrls(DRDAYINDEX + 3 + 8 * (ilLoop - 1)), 705, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmDRCtrls(DRDAYINDEX + 3 + 8 * (ilLoop - 1)).iReq = False
        gSetCtrl tmDRCtrls(DRDAYINDEX + 4 + 8 * (ilLoop - 1)), 930, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmDRCtrls(DRDAYINDEX + 4 + 8 * (ilLoop - 1)).iReq = False
        gSetCtrl tmDRCtrls(DRDAYINDEX + 5 + 8 * (ilLoop - 1)), 1155, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmDRCtrls(DRDAYINDEX + 5 + 8 * (ilLoop - 1)).iReq = False
        gSetCtrl tmDRCtrls(DRDAYINDEX + 6 + 8 * (ilLoop - 1)), 1380, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmDRCtrls(DRDAYINDEX + 6 + 8 * (ilLoop - 1)).iReq = False
        gSetCtrl tmDRCtrls(DRVALUEINDEX + 8 * (ilLoop - 1)), 1605, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 1005, fgBoxGridH
        tmDRCtrls(DRVALUEINDEX + 8 * (ilLoop - 1)).iReq = False
    Next ilLoop
    'Hour Rate
    For ilLoop = 1 To 10 Step 1
        gSetCtrl tmHRCtrls(HRSTARTINDEX + 10 * (ilLoop - 1)), 30, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 750, fgBoxGridH
        tmHRCtrls(HRSTARTINDEX + 10 * (ilLoop - 1)).iReq = False
        gSetCtrl tmHRCtrls(HRENDINDEX + 10 * (ilLoop - 1)), 795, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 750, fgBoxGridH
        tmHRCtrls(HRENDINDEX + 10 * (ilLoop - 1)).iReq = False
        gSetCtrl tmHRCtrls(HRDAYINDEX + 10 * (ilLoop - 1)), 1530 + 30, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmHRCtrls(HRDAYINDEX + 10 * (ilLoop - 1)).iReq = False
        gSetCtrl tmHRCtrls(HRDAYINDEX + 1 + 10 * (ilLoop - 1)), 1530 + 255, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmHRCtrls(HRDAYINDEX + 1 + 10 * (ilLoop - 1)).iReq = False
        gSetCtrl tmHRCtrls(HRDAYINDEX + 2 + 10 * (ilLoop - 1)), 1530 + 480, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmHRCtrls(HRDAYINDEX + 2 + 10 * (ilLoop - 1)).iReq = False
        gSetCtrl tmHRCtrls(HRDAYINDEX + 3 + 10 * (ilLoop - 1)), 1530 + 705, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmHRCtrls(HRDAYINDEX + 3 + 10 * (ilLoop - 1)).iReq = False
        gSetCtrl tmHRCtrls(HRDAYINDEX + 4 + 10 * (ilLoop - 1)), 1530 + 930, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmHRCtrls(HRDAYINDEX + 4 + 10 * (ilLoop - 1)).iReq = False
        gSetCtrl tmHRCtrls(HRDAYINDEX + 5 + 10 * (ilLoop - 1)), 1530 + 1155, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmHRCtrls(HRDAYINDEX + 5 + 10 * (ilLoop - 1)).iReq = False
        gSetCtrl tmHRCtrls(HRDAYINDEX + 6 + 10 * (ilLoop - 1)), 1530 + 1380, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 210, fgBoxGridH
        tmHRCtrls(HRDAYINDEX + 6 + 10 * (ilLoop - 1)).iReq = False
        gSetCtrl tmHRCtrls(HRVALUEINDEX + 10 * (ilLoop - 1)), 1530 + 1605, 420 + (ilLoop - 1) * (fgBoxGridH + 15), 1005, fgBoxGridH
        tmHRCtrls(HRVALUEINDEX + 10 * (ilLoop - 1)).iReq = False
    Next ilLoop
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitShow                       *
'*                                                     *
'*             Created:7/09/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show values         *
'*                                                     *
'*******************************************************
Private Sub mInitShow(ilInitNameFac As Integer)

    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    Dim slStr As String
    Dim slRecCode As String
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilDay As Integer
    If ilInitNameFac Then
        slStr = tmRcf.sName
        gSetShow pbcSpec, slStr, tmSpecCtrls(NAMEINDEX)
        lbcVehicle.ListIndex = -1
    End If
    If tmRcf.iVefCode <> 0 Then
        'If tmRcf.iVefCode > 0 Then
            slRecCode = Trim$(str$(tmRcf.iVefCode))
            For ilLoop = 0 To UBound(tgUserVehicle) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mInitShowErr
                gCPErrorMsg ilRet, "mInitShow (gParseItem field 2)", RCTerms
                On Error GoTo 0
                If slRecCode = slCode Then
                    lbcVehicle.ListIndex = ilLoop
                    Exit For
                End If
            Next ilLoop
        'ElseIf tmRcf.iVefCode < 0 Then
        '    slRecCode = Trim$(Str$(-tmRcf.iVefCode))
        '    For ilLoop = 0 To lbcCombo.ListCount - 1 Step 1
        '        slNameCode = lbcCombo.List(ilLoop)
        '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        '        On Error GoTo mInitShowErr
        '        gCPErrorMsg ilRet, "mInitShow (gParseItem field 2)", RCTerms
        '        On Error GoTo 0
        '        If slRecCode = slCode Then
        '            lbcVehicle.ListIndex = ilLoop + Traffic!lbcUserVehicle.ListCount
        '            Exit For
        '        End If
        '    Next ilLoop
        'End If
    Else
        'If Not ilInitNameFac Then
        If ilInitNameFac Then
            lbcVehicle.ListIndex = 0    'All vehicles
        End If
    End If
    If lbcVehicle.ListIndex >= 0 Then
        slStr = lbcVehicle.List(lbcVehicle.ListIndex)
    Else
        slStr = ""
    End If
    gSetShow pbcSpec, slStr, tmSpecCtrls(VEHICLEINDEX)
    If tmRcf.iYear > 0 Then
        slStr = Trim$(str$(tmRcf.iYear))
    Else
        slStr = ""
    End If
    gSetShow pbcSpec, slStr, tmSpecCtrls(YEARINDEX)
    gUnpackDate tmRcf.iStartDate(0), tmRcf.iStartDate(1), slStr
    slStr = gFormatDate(slStr)
    gSetShow pbcSpec, slStr, tmSpecCtrls(STARTDATEINDEX)
    If igRCMode = 1 Then    'Change
        mLenPop
        If imTerminate Then
            Exit Sub
        End If
        slStr = Trim$(str$(tmRcf.iGridsUsed))
        gSetShow pbcSpec, slStr, tmSpecCtrls(NOGRIDSINDEX)
        slStr = Trim$(str$(tmRcf.iBaseLen))
        gSetShow pbcSpec, slStr, tmSpecCtrls(BASELENINDEX)
        'gPDNToStr tmRcf.sRound, 2, slStr
        slStr = gLongToStrDec(tmRcf.lRound, 2)
        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        gSetShow pbcSpec, slStr, tmSpecCtrls(ROUNDINDEX)
    End If
    If tmRcf.sUseLen = "N" Then
        tmRcf.sUseLen = "I"
    End If
    If tmRcf.sUseLen = "D" Then
        lacLength.Caption = "Dollar"
    Else
        lacLength.Caption = "Index"
    End If
    If tmRcf.sUseSpot = "N" Then
        tmRcf.sUseSpot = "I"
    End If
    If tmRcf.sUseSpot = "D" Then
        lacSpotFreq.Caption = "Dollar"
    Else
        lacSpotFreq.Caption = "Index"
    End If
    If tmRcf.sUseWeek = "N" Then
        tmRcf.sUseWeek = "I"
    End If
    If tmRcf.sUseWeek = "D" Then
        lacWeekFreq.Caption = "Dollar"
    Else
        lacWeekFreq.Caption = "Index"
    End If
    If tmRcf.sUseDay = "N" Then
        tmRcf.sUseDay = "I"
    End If
    If tmRcf.sUseDay = "D" Then
        lacDayRate.Caption = "Dollar"
    Else
        lacDayRate.Caption = "Index"
    End If
    If tmRcf.sUseHour = "N" Then
        tmRcf.sUseHour = "I"
    End If
    If tmRcf.sUseHour = "D" Then
        lacHourRate.Caption = "Dollar"
    Else
        lacHourRate.Caption = "Index"
    End If
    ilBoxNo = 1
    'Check if any length defined- if not create
    ilFound = False
    For ilRowNo = LBound(tmRcf.lValue) To UBound(tmRcf.lValue)
        If (tmRcf.iLen(ilRowNo) <> 0) Then
            ilFound = True
            Exit For
        End If
    Next ilRowNo
    If Not ilFound Then
        For ilLoop = 0 To lbcBaseLen.ListCount - 1 Step 1
            slStr = lbcBaseLen.List(ilLoop)
            'tmRcf.iLen(ilLoop + 1) = Val(slStr)
            tmRcf.iLen(ilLoop) = Val(slStr)
        Next ilLoop
    End If
    For ilRowNo = LBound(tmRcf.lValue) To UBound(tmRcf.lValue)
        If (tmRcf.iLen(ilRowNo) <> 0) Then
            slStr = Trim$(str$(tmRcf.iLen(ilRowNo)))
            gSetShow pbcLength, slStr, tmLenCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
            'If (StrComp(tmRcf.sValue(ilRowNo), smSign5, 0) <> 0) And (StrComp(tmRcf.sValue(ilRowNo), smZero5, 0) <> 0) Then
            If tmRcf.lValue(ilRowNo) <> 0 Then
                'gPDNToStr tmRcf.sValue(ilRowNo), 2, slStr
                slStr = gLongToStrDec(tmRcf.lValue(ilRowNo), 2)
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            Else
                slStr = ""
            End If
        Else
            slStr = ""
            gSetShow pbcLength, slStr, tmLenCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
            slStr = ""
        End If
        gSetShow pbcLength, slStr, tmLenCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
    Next ilRowNo
    ilBoxNo = 1
    For ilRowNo = LBound(tmRcf.iSpotMin) To UBound(tmRcf.iSpotMin) Step 1
        If (tmRcf.iSpotMin(ilRowNo) <> 0) Then
            slStr = Trim$(str$(tmRcf.iSpotMin(ilRowNo)))
            gSetShow pbcSpotFreq, slStr, tmSFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
            If tmRcf.iSpotMax(ilRowNo) <> 0 Then
                slStr = Trim$(str$(tmRcf.iSpotMax(ilRowNo)))
            Else
                slStr = "+"
            End If
            gSetShow pbcSpotFreq, slStr, tmSFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
            'gPDNToStr tmRcf.sSpotVal(ilRowNo), 2, slStr
            slStr = gLongToStrDec(tmRcf.lSpotVal(ilRowNo), 2)
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            gSetShow pbcSpotFreq, slStr, tmSFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
        Else
            slStr = ""
            gSetShow pbcSpotFreq, slStr, tmSFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
            slStr = ""
            gSetShow pbcSpotFreq, slStr, tmSFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
            slStr = ""
            gSetShow pbcSpotFreq, slStr, tmSFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
        End If
    Next ilRowNo
    ilBoxNo = 1
    For ilRowNo = LBound(tmRcf.iWkMin) To UBound(tmRcf.iWkMin) Step 1
        If (tmRcf.iWkMin(ilRowNo) <> 0) Then
            slStr = Trim$(str$(tmRcf.iWkMin(ilRowNo)))
            gSetShow pbcWeekFreq, slStr, tmWFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
            If tmRcf.iWkMax(ilRowNo) <> 0 Then
                slStr = Trim$(str$(tmRcf.iWkMax(ilRowNo)))
            Else
                slStr = "+"
            End If
            gSetShow pbcWeekFreq, slStr, tmWFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
            'gPDNToStr tmRcf.sWkVal(ilRowNo), 2, slStr
            slStr = gLongToStrDec(tmRcf.lWkVal(ilRowNo), 2)
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            gSetShow pbcWeekFreq, slStr, tmWFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
        Else
            slStr = ""
            gSetShow pbcWeekFreq, slStr, tmWFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
            slStr = ""
            gSetShow pbcWeekFreq, slStr, tmWFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
            slStr = ""
            gSetShow pbcWeekFreq, slStr, tmWFCtrls(ilBoxNo)
            ilBoxNo = ilBoxNo + 1
        End If
    Next ilRowNo
    ilBoxNo = 1
    If tmRcf.sUseBB <> "N" Then
        If tmRcf.sUseBB = "D" Then
            slStr = "Dollar"
        Else
            slStr = "Index"
        End If
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        'If (StrComp(tmRcf.sBBValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sBBValue, smZero5, 0) <> 0) Then
        If tmRcf.lBBValue <> 0 Then
            'gPDNToStr tmRcf.sBBValue, 2, slStr
            slStr = gLongToStrDec(tmRcf.lBBValue, 2)
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        Else
            slStr = ""
        End If
    Else
        tmRcf.sUseBB = "I"
        slStr = ""
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        slStr = ""
    End If
    gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
    ilBoxNo = ilBoxNo + 1
    If tmRcf.sUseFP <> "N" Then
        If tmRcf.sUseFP = "D" Then
            slStr = "Dollar"
        Else
            slStr = "Index"
        End If
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        'If (StrComp(tmRcf.sFPValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sFPValue, smZero5, 0) <> 0) Then
        If tmRcf.lFPValue <> 0 Then
            'gPDNToStr tmRcf.sFPValue, 2, slStr
            slStr = gLongToStrDec(tmRcf.lFPValue, 2)
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        Else
            slStr = ""
        End If
    Else
        tmRcf.sUseFP = "I"
        slStr = ""
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        slStr = ""
    End If
    gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
    ilBoxNo = ilBoxNo + 1
    If tmRcf.sUseNP <> "N" Then
        If tmRcf.sUseNP = "D" Then
            slStr = "Dollar"
        Else
            slStr = "Index"
        End If
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        'If (StrComp(tmRcf.sNPValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sNPValue, smZero5, 0) <> 0) Then
        If tmRcf.lNPValue <> 0 Then
            'gPDNToStr tmRcf.sNPValue, 2, slStr
            slStr = gLongToStrDec(tmRcf.lNPValue, 2)
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        Else
            slStr = ""
        End If
    Else
        tmRcf.sUseNP = "I"
        slStr = ""
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        slStr = ""
    End If
    gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)

    ilBoxNo = ilBoxNo + 1
    If (tmRcf.sUsePrefDT <> "N") And (Trim$(tmRcf.sUsePrefDT) <> "") Then
        If tmRcf.sUsePrefDT = "D" Then
            slStr = "Dollar"
        Else
            slStr = "Index"
        End If
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        'If (StrComp(tmRcf.sFPValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sFPValue, smZero5, 0) <> 0) Then
        If tmRcf.lPrefDTValue <> 0 Then
            'gPDNToStr tmRcf.sFPValue, 2, slStr
            slStr = gLongToStrDec(tmRcf.lPrefDTValue, 2)
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        Else
            slStr = ""
        End If
    Else
        tmRcf.sUsePrefDT = "I"
        slStr = ""
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        slStr = ""
    End If
    gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
    ilBoxNo = ilBoxNo + 1
    If (tmRcf.sUse1stPos <> "N") And (Trim$(tmRcf.sUse1stPos) <> "") Then
        If tmRcf.sUse1stPos = "D" Then
            slStr = "Dollar"
        Else
            slStr = "Index"
        End If
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        'If (StrComp(tmRcf.sFPValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sFPValue, smZero5, 0) <> 0) Then
        If tmRcf.l1stPosValue <> 0 Then
            'gPDNToStr tmRcf.sFPValue, 2, slStr
            slStr = gLongToStrDec(tmRcf.l1stPosValue, 2)
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        Else
            slStr = ""
        End If
    Else
        tmRcf.sUse1stPos = "I"
        slStr = ""
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        slStr = ""
    End If
    gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
    ilBoxNo = ilBoxNo + 1
    If (tmRcf.sUseSoloAvail <> "N") And (Trim$(tmRcf.sUseSoloAvail) <> "") Then
        If tmRcf.sUseSoloAvail = "D" Then
            slStr = "Dollar"
        Else
            slStr = "Index"
        End If
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        'If (StrComp(tmRcf.sFPValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sFPValue, smZero5, 0) <> 0) Then
        If tmRcf.lSoloAvailValue <> 0 Then
            'gPDNToStr tmRcf.sFPValue, 2, slStr
            slStr = gLongToStrDec(tmRcf.lSoloAvailValue, 2)
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        Else
            slStr = ""
        End If
    Else
        tmRcf.sUseSoloAvail = "I"
        slStr = ""
        gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        slStr = ""
    End If
    gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)


    ilBoxNo = 1
    For ilRowNo = LBound(tmRcf.lGridIndex) To UBound(tmRcf.lGridIndex) Step 1
        'If (StrComp(tmRcf.sGridIndex(ilRowNo), smSign3, 0) <> 0) And (StrComp(tmRcf.sGridIndex(ilRowNo), smZero3, 0) <> 0) Then
        If tmRcf.lGridIndex(ilRowNo) <> 0 Then
            'gPDNToStr tmRcf.sGridIndex(ilRowNo), 4, slStr
            slStr = gLongToStrDec(tmRcf.lGridIndex(ilRowNo), 4)
        Else
            slStr = ""
        End If
        gSetShow pbcGrid, slStr, tmGDCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
    Next ilRowNo
    ilBoxNo = 1
    For ilRowNo = LBound(tmRcf.lDyRate) To UBound(tmRcf.lDyRate) Step 1
        If tmRcf.sUseDay <> "N" Then
            For ilDay = 1 To 7 Step 1
                If tmRcf.sDay(ilRowNo, ilDay - 1) = "Y" Then
                    slStr = "Y"
                Else
                    slStr = " "
                End If
                gSetShow pbcDayRate, slStr, tmDRCtrls(ilBoxNo)
                ilBoxNo = ilBoxNo + 1
            Next ilDay
            'If (StrComp(tmRcf.sDyRate(ilRowNo), smSign5, 0) <> 0) And (StrComp(tmRcf.sDyRate(ilRowNo), smZero5, 0) <> 0) Then
            If tmRcf.lDyRate(ilRowNo) <> 0 Then
                'gPDNToStr tmRcf.sDyRate(ilRowNo), 2, slStr
                slStr = gLongToStrDec(tmRcf.lDyRate(ilRowNo), 2)
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            Else
                slStr = ""
            End If
        Else
            For ilDay = 1 To 7 Step 1
                slStr = " "
                gSetShow pbcDayRate, slStr, tmDRCtrls(ilBoxNo)
                ilBoxNo = ilBoxNo + 1
            Next ilDay
            slStr = ""
        End If
        gSetShow pbcDayRate, slStr, tmDRCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
    Next ilRowNo
    ilBoxNo = 1
    For ilRowNo = LBound(tmRcf.lHrRate) To UBound(tmRcf.lHrRate) Step 1
        If (tmRcf.iHrStartTime(0, ilRowNo) <> 0) Or (tmRcf.iHrStartTime(1, ilRowNo) <> 0) Then
            gUnpackTime tmRcf.iHrStartTime(0, ilRowNo), tmRcf.iHrStartTime(1, ilRowNo), "A", "1", slStr
        Else
            slStr = ""
        End If
        gSetShow pbcHourRate, slStr, tmHRCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        If (tmRcf.iHrEndTime(0, ilRowNo) <> 0) Or (tmRcf.iHrEndTime(1, ilRowNo) <> 0) Then
            gUnpackTime tmRcf.iHrEndTime(0, ilRowNo), tmRcf.iHrEndTime(1, ilRowNo), "A", "1", slStr
        Else
            slStr = ""
        End If
        gSetShow pbcHourRate, slStr, tmHRCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
        If tmRcf.sUseHour <> "N" Then
            For ilDay = 1 To 7 Step 1
                If tmRcf.sHrDay(ilRowNo, ilDay - 1) = "Y" Then
                    slStr = "Y"
                Else
                    slStr = " "
                End If
                gSetShow pbcHourRate, slStr, tmHRCtrls(ilBoxNo)
                ilBoxNo = ilBoxNo + 1
            Next ilDay
            'If (StrComp(tmRcf.sHrRate(ilRowNo), smSign5, 0) <> 0) And (StrComp(tmRcf.sHrRate(ilRowNo), smZero5, 0) <> 0) Then
            If tmRcf.lHrRate(ilRowNo) <> 0 Then
                'gPDNToStr tmRcf.sHrRate(ilRowNo), 2, slStr
                slStr = gLongToStrDec(tmRcf.lHrRate(ilRowNo), 2)
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            Else
                slStr = ""
            End If
        Else
            For ilDay = 1 To 7 Step 1
                slStr = " "
                gSetShow pbcHourRate, slStr, tmHRCtrls(ilBoxNo)
                ilBoxNo = ilBoxNo + 1
            Next ilDay
            slStr = ""
        End If
        gSetShow pbcHourRate, slStr, tmHRCtrls(ilBoxNo)
        ilBoxNo = ilBoxNo + 1
    Next ilRowNo
    Exit Sub
mInitShowErr:
    On Error GoTo 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLenPop                         *
'*                                                     *
'*             Created:7/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate length box from       *
'*                      lengths specified for Vehicle *
'*                                                     *
'*******************************************************
Private Sub mLenPop()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim hlVsf As Integer
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim ilDestroy As Integer

    lbcBaseLen.Clear
    If lbcVehicle.ListIndex <= 0 Then
        imVpfIndex = -1
    'ElseIf lbcVehicle.ListIndex < Traffic!lbcUserVehicle.ListCount Then
    Else
        slNameCode = tgUserVehicle(lbcVehicle.ListIndex - 1).sKey 'Traffic!lbcUserVehicle.List(lbcVehicle.ListIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mLenPopErr
        gCPErrorMsg ilRet, "mLenPop (gParseItem field 2)", RCTerms
        On Error GoTo 0
        ilCode = CInt(slCode)
        imVpfIndex = gVpfFind(RCTerms, ilCode)
    'Else
    '    slNameCode = lbcCombo.List(lbcVehicle.ListIndex - Traffic!lbcUserVehicle.ListCount)
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '    On Error GoTo mLenPopErr
    '    gCPErrorMsg ilRet, "mLenPop (gParseItem field 2)", RCTerms
    '    On Error GoTo 0
    '    hlVsf = CBtrvTable()
    '    ilDestroy = True
    '    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    '    On Error GoTo mLenPopErr
    '    gBtrvErrorMsg ilRet, "mLenPop (btrOpen)" & "Vsf.Btr", RCTerms
    '    On Error GoTo 0
    '    ilRecLen = Len(tlVsf)  'btrRecordLength(hlVpf)  'Get and save record length
    '    tlSrchKey.iCode = CInt(slCode)
    '    ilRet = btrGetEqual(hlVsf, tlVsf, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    '    On Error GoTo mLenPopErr
    '    gBtrvErrorMsg ilRet, "mLenPop (btrGetEqual)", RCTerms
    '    On Error GoTo 0
    '    ilRet = btrClose(hlVsf)
    '    btrDestroy hlVsf
    '    ilDestroy = False
    '    imVpfIndex = gVpfFind(RCTerms, tlVsf.iFSCode(0))
    End If
    If imVpfIndex < 0 Then
        For ilLoop = LBound(tgSpf.iSLen) To UBound(tgSpf.iSLen) Step 1
            If tgSpf.iSLen(ilLoop) <> 0 Then
                If igRCMode = 0 Then 'New
                    ilFound = False
                    For ilIndex = LBound(tmRcf.iLen) To UBound(tmRcf.iLen) Step 1
                        If tgSpf.iSLen(ilLoop) = tmRcf.iLen(ilIndex) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilIndex
                    If Not ilFound Then
                        For ilIndex = LBound(tmRcf.iLen) To UBound(tmRcf.iLen) Step 1
                            If tmRcf.iLen(ilIndex) = 0 Then
                                tmRcf.iLen(ilIndex) = tgSpf.iSLen(ilLoop)
                                Exit For
                            End If
                        Next ilIndex
                    End If
                End If
                lbcBaseLen.AddItem Trim$(str$(tgSpf.iSLen(ilLoop)))
            End If
        Next ilLoop
    Else
        For ilLoop = LBound(tgVpf(imVpfIndex).iSLen) To UBound(tgVpf(imVpfIndex).iSLen) Step 1
            If tgVpf(imVpfIndex).iSLen(ilLoop) <> 0 Then
                If igRCMode = 0 Then 'New
                    ilFound = False
                    For ilIndex = LBound(tmRcf.iLen) To UBound(tmRcf.iLen) Step 1
                        If tgVpf(imVpfIndex).iSLen(ilLoop) = tmRcf.iLen(ilIndex) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilIndex
                    If Not ilFound Then
                        For ilIndex = LBound(tmRcf.iLen) To UBound(tmRcf.iLen) Step 1
                            If tmRcf.iLen(ilIndex) = 0 Then
                                tmRcf.iLen(ilIndex) = tgVpf(imVpfIndex).iSLen(ilLoop)
                                Exit For
                            End If
                        Next ilIndex
                    End If
                End If
                lbcBaseLen.AddItem Trim$(str$(tgVpf(imVpfIndex).iSLen(ilLoop)))
            End If
        Next ilLoop
    End If
    Exit Sub
mLenPopErr:
    On Error GoTo 0
    If ilDestroy Then
        ilRet = btrClose(hlVsf)
        btrDestroy hlVsf
    End If
    igRCReturn = 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mModelFromOldestRC              *
'*                                                     *
'*             Created:7/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Model new rate card from       *
'*                      specified rate card            *
'*                                                     *
'*******************************************************
Private Sub mModelRcf(ilRcfCode As Integer)
    Dim ilRecLen As Integer     'RcF record length
    Dim hlRcf As Integer        'User Option file handle
    Dim tlRcf As RCF
    Dim tlSrchKey As INTKEY0
    Dim ilRet As Integer
    Dim ilLoop As Integer
    hlRcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlRcf, "", sgDBPath & "Rcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mModelRcfErr
    gBtrvErrorMsg ilRet, "mModelRcf (btrOpen):" & "Rcf.Btr", RCTerms
    On Error GoTo 0
    ilRecLen = Len(tlRcf)  'btrRecordLength(hlRcf)  'Get and save record length
    tlSrchKey.iCode = ilRcfCode
    ilRet = btrGetEqual(hlRcf, tmRcf, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mModelRcfErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", RCTerms
    On Error GoTo 0
    tmRcf.sName = ""
    tmRcf.iVefCode = -32000
    tmRcf.iYear = 0
    tmRcf.iStartDate(0) = 0
    tmRcf.iStartDate(1) = 0
    tmRcf.iEndDate(0) = 0
    tmRcf.iEndDate(1) = 0
    For ilLoop = LBound(tmRcf.iLen) To UBound(tmRcf.iLen) Step 1
        If tmRcf.iLen(ilLoop) = 0 Then
            Exit For
        End If
    Next ilLoop
    ilRet = btrClose(hlRcf)
    btrDestroy hlRcf
    Exit Sub
mModelRcfErr:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    ilRet = btrClose(hlRcf)
    btrDestroy hlRcf
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetAll                        *
'*                                                     *
'*             Created:7/15/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear all controls             *
'*                                                     *
'*******************************************************
Private Sub mSetAll()
    mSetShow
    imBoxNo = -1
    imType = -1
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:7/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSpotEnableBox imBoxNo
'   Where:
'       imBoxNo (I)- Number of the Control to be enabled
'
    Dim ilBox As Integer
    If imType <= 0 Then
        Exit Sub
    End If
    Select Case imType
        Case SPOTTYPE
            ilBox = (imBoxNo - 1) Mod 3 + 1
            Select Case ilBox

            End Select
        Case SPOTFREQ
            If (imBoxNo < imLBSFCtrls) Or (imBoxNo > UBound(tmSFCtrls)) Then
                Exit Sub
            End If
            ilBox = (imBoxNo - 1) Mod 3 + 1
            Select Case ilBox
                Case 1    'From Value
                    edcDropDown.SetFocus
                Case 2   'From Value
                    edcDropDown.SetFocus
                Case 3  'From Value
                    edcDropDown.SetFocus
            End Select
        Case WEEKFREQ
            If (imBoxNo < imLBWFCtrls) Or (imBoxNo > UBound(tmWFCtrls)) Then
                Exit Sub
            End If
            ilBox = (imBoxNo - 1) Mod 3 + 1
            Select Case ilBox 'Branch on box type (control)
                Case 1 'From Value
                    edcDropDown.SetFocus
                Case 2  'From Value
                    edcDropDown.SetFocus
                Case 3  'Value
                    edcDropDown.SetFocus
            End Select
        Case Length
            If (imBoxNo < imLBLenCtrls) Or (imBoxNo > UBound(tmLenCtrls)) Then
                Exit Sub
            End If
            ilBox = (imBoxNo - 1) Mod 2 + 1
            Select Case ilBox
                Case 2 'Value
                    edcDropDown.SetFocus
            End Select
        Case GRID
            If (imBoxNo < imLBGDCtrls) Or (imBoxNo > UBound(tmGDCtrls)) Then
                Exit Sub
            End If
            Select Case imBoxNo
                Case GRIDVALUEINDEX To GRIDVALUEINDEX + 11 'Value
                    edcDropDown.SetFocus
            End Select
        Case DAYRATE
            If (imBoxNo < imLBDRCtrls) Or (imBoxNo > UBound(tmDRCtrls)) Then
                Exit Sub
            End If
            ilBox = (imBoxNo - 1) Mod 8 + 1
            Select Case ilBox
                Case 1, 2, 3, 4, 5, 6, 7  'Mo-Su
                    ckcTDDay.SetFocus
                Case 8  'Value
                    edcDropDown.SetFocus
            End Select
        Case HOURRATE
            If (imBoxNo < imLBHRCtrls) Or (imBoxNo > UBound(tmHRCtrls)) Then
                Exit Sub
            End If
            ilBox = (imBoxNo - 1) Mod 10 + 1
            Select Case ilBox
                Case 1
                    edcDropDown.SetFocus
                Case 2
                    edcDropDown.SetFocus
                Case 3, 4, 5, 6, 7, 8, 9  'Mo-Su
                    ckcTDDay.SetFocus
                Case 10  'Value
                    edcDropDown.SetFocus
            End Select
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:7/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set show and save values       *
'*                                                     *
'*******************************************************
Private Sub mSetShow()
'
'   mSetShow
'   Where:
'       imBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilBox As Integer
    Dim slvalue As String
    pbcArrow.Visible = False
    If imType <= 0 Then
        Exit Sub
    End If
    Select Case imType
        Case SPOTTYPE
            If (imBoxNo < imLBSTCtrls) Or (imBoxNo > UBound(tmSTCtrls)) Then
                Exit Sub
            End If
            edcDropDown.Visible = False
            pbcBySpotType.Visible = False
            ilIndex = (imBoxNo - 1) \ STVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod STVALUEINDEX + 1
            Select Case ilBox
                Case 1
                    Select Case ilIndex
                        Case 1
                            If tmRcf.sUseBB <> "I" Then
                                slStr = "Dollar"
                            Else
                                slStr = "Index"
                            End If
                        Case 2
                            If tmRcf.sUseFP <> "I" Then
                                slStr = "Dollar"
                            Else
                                slStr = "Index"
                            End If
                        Case 3
                            If tmRcf.sUseNP <> "I" Then
                                slStr = "Dollar"
                            Else
                                slStr = "Index"
                            End If
                        Case 4
                            If tmRcf.sUsePrefDT <> "I" Then
                                slStr = "Dollar"
                            Else
                                slStr = "Index"
                            End If
                        Case 5
                            If tmRcf.sUse1stPos <> "I" Then
                                slStr = "Dollar"
                            Else
                                slStr = "Index"
                            End If
                        Case 6
                            If tmRcf.sUseSoloAvail <> "I" Then
                                slStr = "Dollar"
                            Else
                                slStr = "Index"
                            End If
                    End Select
                    gSetShow pbcSpotType, slStr, tmSTCtrls(imBoxNo)
                Case 2
                    slStr = Trim$(edcDropDown.Text)
                    slvalue = slStr
                    Select Case ilIndex
                        Case 1
                            'gStrToPDN slStr, 2, 5, tmRcf.sBBValue
                            tmRcf.lBBValue = gStrDecToLong(slStr, 2)
                            'If (StrComp(tmRcf.sBBValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sBBValue, smZero5, 0) <> 0) Then
                            If tmRcf.lBBValue <> 0 Then
                                gFormatStr slvalue, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                            Else
                                slStr = ""
                            End If
                        Case 2
                            'gStrToPDN slStr, 2, 5, tmRcf.sFPValue
                            tmRcf.lFPValue = gStrDecToLong(slStr, 2)
                            'If (StrComp(tmRcf.sFPValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sFPValue, smZero5, 0) <> 0) Then
                            If tmRcf.lFPValue <> 0 Then
                                gFormatStr slvalue, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                            Else
                                slStr = ""
                            End If
                        Case 3
                            'gStrToPDN slStr, 2, 5, tmRcf.sNPValue
                            tmRcf.lNPValue = gStrDecToLong(slStr, 2)
                            'If (StrComp(tmRcf.sNPValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sNPValue, smZero5, 0) <> 0) Then
                            If tmRcf.lNPValue <> 0 Then
                                gFormatStr slvalue, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                            Else
                                slStr = ""
                            End If
                        Case 4
                            'gStrToPDN slStr, 2, 5, tmRcf.sBBValue
                            tmRcf.lPrefDTValue = gStrDecToLong(slStr, 2)
                            'If (StrComp(tmRcf.sBBValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sBBValue, smZero5, 0) <> 0) Then
                            If tmRcf.lPrefDTValue <> 0 Then
                                gFormatStr slvalue, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                            Else
                                slStr = ""
                            End If
                        Case 5
                            'gStrToPDN slStr, 2, 5, tmRcf.sFPValue
                            tmRcf.l1stPosValue = gStrDecToLong(slStr, 2)
                            'If (StrComp(tmRcf.sFPValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sFPValue, smZero5, 0) <> 0) Then
                            If tmRcf.l1stPosValue <> 0 Then
                                gFormatStr slvalue, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                            Else
                                slStr = ""
                            End If
                        Case 6
                            'gStrToPDN slStr, 2, 5, tmRcf.sNPValue
                            tmRcf.lSoloAvailValue = gStrDecToLong(slStr, 2)
                            'If (StrComp(tmRcf.sNPValue, smSign5, 0) <> 0) And (StrComp(tmRcf.sNPValue, smZero5, 0) <> 0) Then
                            If tmRcf.lSoloAvailValue <> 0 Then
                                gFormatStr slvalue, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                            Else
                                slStr = ""
                            End If
                    End Select
                    gSetShow pbcSpotFreq, slStr, tmSTCtrls(imBoxNo)
            End Select
        Case SPOTFREQ
            If (imBoxNo < imLBSFCtrls) Or (imBoxNo > UBound(tmSFCtrls)) Then
                Exit Sub
            End If
            edcDropDown.Visible = False
            ilIndex = (imBoxNo - 1) \ SPOTVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod SPOTVALUEINDEX + 1
            Select Case ilBox
                Case 1    'From Value
                    slStr = Trim$(edcDropDown.Text)
                    If slStr <> "" Then
                        tmRcf.iSpotMin(ilIndex - 1) = Val(slStr)
                    Else
                        tmRcf.iSpotMin(ilIndex - 1) = 0
                    End If
                    gSetShow pbcSpotFreq, slStr, tmSFCtrls(imBoxNo)
                Case 2   'From Value
                    slStr = Trim$(edcDropDown.Text)
                    If slStr <> "" Then
                        If slStr <> "+" Then
                            tmRcf.iSpotMax(ilIndex - 1) = Val(slStr)
                            If tmRcf.iSpotMax(ilIndex - 1) = 0 Then
                                slStr = "+"
                            End If
                        Else
                            tmRcf.iSpotMax(ilIndex - 1) = 0
                        End If
                    Else
                        tmRcf.iSpotMax(ilIndex - 1) = 0
                    End If
                    gSetShow pbcSpotFreq, slStr, tmSFCtrls(imBoxNo)
                    If tmRcf.iSpotMax(ilIndex - 1) = 0 Then 'Remove all other rows
                        For ilLoop = ilIndex + 1 To UBound(tmRcf.iSpotMin) + 1 Step 1
                            tmRcf.iSpotMin(ilLoop - 1) = 0
                            slStr = ""
                            gSetShow pbcSpotFreq, slStr, tmSFCtrls(3 * (ilLoop - 1) + 1)
                            tmRcf.iSpotMax(ilLoop - 1) = 0
                            slStr = ""
                            gSetShow pbcSpotFreq, slStr, tmSFCtrls(3 * (ilLoop - 1) + 2)
                            'slStr = ""
                            'gStrToPDN slStr, 2, 5, tmRcf.sSpotVal(ilLoop)
                            tmRcf.lSpotVal(ilLoop - 1) = 0
                            slStr = ""
                            gSetShow pbcSpotFreq, slStr, tmSFCtrls(3 * (ilLoop - 1) + 3)
                        Next ilLoop
                    End If
                Case 3  'From Value
                    slStr = Trim$(edcDropDown.Text)
                    slvalue = slStr
                    'gStrToPDN slStr, 2, 5, tmRcf.sSpotVal(ilIndex)
                    tmRcf.lSpotVal(ilIndex - 1) = gStrDecToLong(slStr, 2)
                    'If (StrComp(tmRcf.sSpotVal(ilIndex), smSign5, 0) <> 0) And (StrComp(tmRcf.sSpotVal(ilIndex), smZero5, 0) <> 0) Then
                    If tmRcf.lSpotVal(ilIndex - 1) <> 0 Then
                        gFormatStr slvalue, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    Else
                        slStr = ""
                    End If
                    gSetShow pbcSpotFreq, slStr, tmSFCtrls(imBoxNo)
            End Select
        Case WEEKFREQ
            If (imBoxNo < imLBWFCtrls) Or (imBoxNo > UBound(tmWFCtrls)) Then
                Exit Sub
            End If
            edcDropDown.Visible = False
            ilIndex = (imBoxNo - 1) \ WEEKVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod WEEKVALUEINDEX + 1
            Select Case ilBox
                Case 1    'From Value
                    slStr = Trim$(edcDropDown.Text)
                    If slStr <> "" Then
                        tmRcf.iWkMin(ilIndex - 1) = Val(slStr)
                    Else
                        tmRcf.iWkMin(ilIndex - 1) = 0
                    End If
                    gSetShow pbcWeekFreq, slStr, tmWFCtrls(imBoxNo)
                Case 2   'From Value
                    slStr = Trim$(edcDropDown.Text)
                    If slStr <> "" Then
                        If slStr <> "+" Then
                            tmRcf.iWkMax(ilIndex - 1) = Val(slStr)
                            If tmRcf.iWkMax(ilIndex - 1) = 0 Then
                                slStr = "+"
                            End If
                        Else
                            tmRcf.iWkMax(ilIndex - 1) = 0
                        End If
                    Else
                        tmRcf.iWkMax(ilIndex - 1) = 0
                    End If
                    gSetShow pbcWeekFreq, slStr, tmWFCtrls(imBoxNo)
                    If tmRcf.iWkMax(ilIndex - 1) = 0 Then 'Remove all other rows
                        For ilLoop = ilIndex + 1 To UBound(tmRcf.iWkMin) + 1 Step 1
                            tmRcf.iWkMin(ilLoop - 1) = 0
                            slStr = ""
                            gSetShow pbcWeekFreq, slStr, tmWFCtrls(3 * (ilLoop - 1) + 1)
                            tmRcf.iWkMax(ilLoop - 1) = 0
                            slStr = ""
                            gSetShow pbcWeekFreq, slStr, tmWFCtrls(3 * (ilLoop - 1) + 2)
                            'slStr = ""
                            'gStrToPDN slStr, 2, 5, tmRcf.sWkVal(ilLoop)
                            tmRcf.lWkVal(ilLoop - 1) = 0
                            slStr = ""
                            gSetShow pbcWeekFreq, slStr, tmWFCtrls(3 * (ilLoop - 1) + 3)
                        Next ilLoop
                    End If
                Case 3  'From Value
                    slStr = Trim$(edcDropDown.Text)
                    slvalue = slStr
                    'gStrToPDN slStr, 2, 5, tmRcf.sWkVal(ilIndex)
                    tmRcf.lWkVal(ilIndex) = gStrDecToLong(slStr, 2)
                    'If (StrComp(tmRcf.sWkVal(ilIndex), smSign5, 0) <> 0) And (StrComp(tmRcf.sWkVal(ilIndex), smZero5, 0) <> 0) Then
                    If tmRcf.lWkVal(ilIndex) <> 0 Then
                        gFormatStr slvalue, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    End If
                    gSetShow pbcWeekFreq, slStr, tmWFCtrls(imBoxNo)
            End Select
        Case Length
            If (imBoxNo < imLBLenCtrls) Or (imBoxNo > UBound(tmLenCtrls)) Then
                Exit Sub
            End If
            edcDropDown.Visible = False
            ilIndex = (imBoxNo - 1) \ LENVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod LENVALUEINDEX + 1
            slStr = Trim$(edcDropDown.Text)
            Select Case ilBox
                Case 2 'Value
                    'gStrToPDN slStr, 2, 5, tmRcf.sValue(ilIndex)
                    tmRcf.lValue(ilIndex - 1) = gStrDecToLong(slStr, 2)
                    'If (StrComp(tmRcf.sValue(ilIndex), smSign5, 0) <> 0) And (StrComp(tmRcf.sValue(ilIndex), smZero5, 0) <> 0) Then
                    If tmRcf.lValue(ilIndex - 1) <> 0 Then
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    Else
                        slStr = ""
                    End If
                    gSetShow pbcLength, slStr, tmLenCtrls(imBoxNo)
            End Select
        Case GRID
            If (imBoxNo < imLBGDCtrls) Or (imBoxNo > UBound(tmGDCtrls)) Then
                Exit Sub
            End If
            edcDropDown.Visible = False
            ilIndex = imBoxNo
            ilBox = imBoxNo
            slStr = Trim$(edcDropDown.Text)
            Select Case ilBox
                Case GRIDVALUEINDEX To GRIDVALUEINDEX + 11 'Value
                    'gStrToPDN slStr, 4, 3, tmRcf.sGridIndex(ilIndex)
                    tmRcf.lGridIndex(ilIndex - 1) = gStrDecToLong(slStr, 4)
                    'If (StrComp(tmRcf.sGridIndex(ilIndex), smSign3, 0) <> 0) And (StrComp(tmRcf.sGridIndex(ilIndex), smZero3, 0) <> 0) Then
                    If tmRcf.lGridIndex(ilIndex - 1) <> 0 Then
                        gSetShow pbcGrid, slStr, tmGDCtrls(ilIndex)
                    Else
                        slStr = ""
                        gSetShow pbcGrid, slStr, tmGDCtrls(ilIndex)
                    End If
            End Select
        Case DAYRATE
            If (imBoxNo < imLBDRCtrls) Or (imBoxNo > UBound(tmDRCtrls)) Then
                Exit Sub
            End If
            edcDropDown.Visible = False
            pbcTDDays.Visible = False
            ilIndex = (imBoxNo - 1) \ DRVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod DRVALUEINDEX + 1
            Select Case ilBox
                Case 1, 2, 3, 4, 5, 6, 7  'Mo-Su
                    If ckcTDDay.Value = vbChecked Then
                        tmRcf.sDay(ilIndex - 1, ilBox - 1) = "Y"
                        slStr = "Y"
                    Else
                        tmRcf.sDay(ilIndex - 1, ilBox - 1) = "N"
                        slStr = " "
                    End If
                    gSetShow pbcDayRate, slStr, tmDRCtrls(imBoxNo)
                Case 8  'Value
                    slStr = Trim$(edcDropDown.Text)
                    slvalue = slStr
                    'gStrToPDN slStr, 2, 5, tmRcf.sDyRate(ilIndex)
                    tmRcf.lDyRate(ilIndex - 1) = gStrDecToLong(slStr, 2)
                    'If (StrComp(tmRcf.sDyRate(ilIndex), smSign5, 0) <> 0) And (StrComp(tmRcf.sDyRate(ilIndex), smZero5, 0) <> 0) Then
                    If tmRcf.lDyRate(ilIndex - 1) <> 0 Then
                        gFormatStr slvalue, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    Else
                        slStr = ""
                    End If
                    gSetShow pbcDayRate, slStr, tmDRCtrls(imBoxNo)
            End Select
        Case HOURRATE
            If (imBoxNo < imLBHRCtrls) Or (imBoxNo > UBound(tmHRCtrls)) Then
                Exit Sub
            End If
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            plcTme.Visible = False
            pbcTDDays.Visible = False
            ilIndex = (imBoxNo - 1) \ HRVALUEINDEX + 1
            ilBox = (imBoxNo - 1) Mod HRVALUEINDEX + 1
            Select Case ilBox
                Case 1
                    slStr = Trim$(edcDropDown.Text)
                    gPackTime slStr, tmRcf.iHrStartTime(0, ilIndex - 1), tmRcf.iHrStartTime(1, ilIndex - 1)
                    gSetShow pbcHourRate, slStr, tmHRCtrls(imBoxNo)
                Case 2
                    slStr = Trim$(edcDropDown.Text)
                    gPackTime slStr, tmRcf.iHrEndTime(0, ilIndex - 1), tmRcf.iHrEndTime(1, ilIndex - 1)
                    gSetShow pbcHourRate, slStr, tmHRCtrls(imBoxNo)
                Case 3, 4, 5, 6, 7, 8, 9  'Mo-Su
                    If ckcTDDay.Value = vbChecked Then
                        'tmRcf.sHrDay(ilIndex, ilBox - 2) = "Y"
                        tmRcf.sHrDay(ilIndex - 1, ilBox - 3) = "Y"
                        slStr = "Y"
                    Else
                        'tmRcf.sHrDay(ilIndex, ilBox - 2) = "N"
                        tmRcf.sHrDay(ilIndex - 1, ilBox - 3) = "N"
                        slStr = " "
                    End If
                    gSetShow pbcHourRate, slStr, tmHRCtrls(imBoxNo)
                Case 10  'Value
                    slStr = Trim$(edcDropDown.Text)
                    slvalue = slStr
                    'gStrToPDN slStr, 2, 5, tmRcf.sHrRate(ilIndex)
                    tmRcf.lHrRate(ilIndex - 1) = gStrDecToLong(slStr, 2)
                    'If (StrComp(tmRcf.sHrRate(ilIndex), smSign5, 0) <> 0) And (StrComp(tmRcf.sHrRate(ilIndex), smZero5, 0) <> 0) Then
                    If tmRcf.lHrRate(ilIndex - 1) <> 0 Then
                        gFormatStr slvalue, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    Else
                        slStr = ""
                    End If
                    gSetShow pbcHourRate, slStr, tmHRCtrls(imBoxNo)
            End Select
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecEnableBox                  *
'*                                                     *
'*             Created:7/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSpecEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim slDate As String
    Dim slStdDate As String

    If ilBoxNo < imLBSpecCtrls Or ilBoxNo > UBound(tmSpecCtrls) Then
        Exit Sub
    End If

    edcSpecDropDown.Width = tmSpecCtrls(ilBoxNo).fBoxW
    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(NAMEINDEX).fBoxX, tmSpecCtrls(NAMEINDEX).fBoxY
            edcSpecDropDown.MaxLength = 4
            imEditType = 0
            edcSpecDropDown.Text = Trim$(tmRcf.sName)
            edcSpecDropDown.Visible = True  'Set visibility
            edcSpecDropDown.SetFocus
        Case VEHICLEINDEX
            mVehPop lbcVehicle
            If imTerminate Then
                Exit Sub
            End If
            lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 15)
            edcSpecDropDown.Width = tmSpecCtrls(VEHICLEINDEX).fBoxW - cmcSpecDropDown.Width
            If tgSpf.iVehLen <= 40 Then
                edcSpecDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcSpecDropDown.MaxLength = 20
            End If
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(VEHICLEINDEX).fBoxX, tmSpecCtrls(VEHICLEINDEX).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            imChgMode = True
            If lbcVehicle.ListIndex >= 0 Then
                imComboBoxIndex = lbcVehicle.ListIndex
                edcSpecDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
            Else
                gFindMatch sgUserDefVehicleName, 0, lbcVehicle
                If gLastFound(lbcVehicle) >= 0 Then
                    lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                    imComboBoxIndex = lbcVehicle.ListIndex
                    edcSpecDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                Else
                    lbcVehicle.ListIndex = 0
                    imComboBoxIndex = lbcVehicle.ListIndex
                    edcSpecDropDown.Text = lbcVehicle.List(0)
                End If
            End If
            imChgMode = False
            lbcVehicle.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.Height
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case YEARINDEX 'Year
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(YEARINDEX).fBoxX, tmSpecCtrls(YEARINDEX).fBoxY
            edcSpecDropDown.MaxLength = 4
            imEditType = 2
            If tmRcf.iYear = 0 Then
                'If Start Date asked before year- then get year from start date
                slDate = Format$(gNow(), "m/d/yy")   'Get year
                slDate = gObtainEndStd(slDate)
                gObtainMonthYear 0, slDate, ilMonth, ilYear
                slStr = Trim$(str$(ilYear)) '""
            Else
                slStr = Trim$(str$(tmRcf.iYear))
            End If
            edcSpecDropDown.Text = slStr
            edcSpecDropDown.Visible = True  'Set visibility
            edcSpecDropDown.SetFocus
        Case STARTDATEINDEX 'Start date
            edcSpecDropDown.Width = tmSpecCtrls(STARTDATEINDEX).fBoxW - cmcSpecDropDown.Width
            edcSpecDropDown.MaxLength = 10
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(STARTDATEINDEX).fBoxX, tmSpecCtrls(STARTDATEINDEX).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            plcCalendar.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.Height
            If (tmRcf.iStartDate(0) = 0) And (tmRcf.iStartDate(1) = 0) Then
                slStr = Format$(gDateValue(gNow()) + 40, "m/d/yy")
                slStr = gObtainStartStd(slStr)
                slStdDate = "1/15/" & Trim$(str$(tmRcf.iYear))
                slStdDate = gObtainStartStd(slStdDate)
                If gDateValue(slStdDate) > gDateValue(slStr) Then
                    slStr = slStdDate
                End If
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
                slStr = slStr
            Else
                gUnpackDate tmRcf.iStartDate(0), tmRcf.iStartDate(1), slStr    'Last log date
            End If
            edcSpecDropDown.Text = slStr
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case NOGRIDSINDEX 'No Grid Levels
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(NOGRIDSINDEX).fBoxX, tmSpecCtrls(NOGRIDSINDEX).fBoxY
            edcSpecDropDown.MaxLength = 2
            imEditType = 2
            slStr = Trim$(str$(tmRcf.iGridsUsed))
            edcSpecDropDown.Text = slStr
            edcSpecDropDown.Visible = True  'Set visibility
            edcSpecDropDown.SetFocus
        Case ROUNDINDEX 'Round to nearset
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(ROUNDINDEX).fBoxX, tmSpecCtrls(ROUNDINDEX).fBoxY
            edcSpecDropDown.MaxLength = 6
            imEditType = 1
            'gPDNToStr tmRcf.sRound, 2, slStr
            slStr = gLongToStrDec(tmRcf.lRound, 2)
            edcSpecDropDown.Text = slStr
            edcSpecDropDown.Visible = True  'Set visibility
            edcSpecDropDown.SetFocus
        Case BASELENINDEX 'Base Length
            lbcBaseLen.Height = gListBoxHeight(lbcBaseLen.ListCount, 12)
            edcSpecDropDown.Width = tmSpecCtrls(BASELENINDEX).fBoxW - cmcSpecDropDown.Width
            edcSpecDropDown.MaxLength = 3
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(BASELENINDEX).fBoxX, tmSpecCtrls(BASELENINDEX).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            imChgMode = True
            If lbcBaseLen.ListIndex >= 0 Then
                imComboBoxIndex = lbcBaseLen.ListIndex
                edcSpecDropDown.Text = lbcBaseLen.List(lbcBaseLen.ListIndex)
            Else
                slStr = Trim$(str$(tmRcf.iBaseLen))
                gFindMatch slStr, 0, lbcBaseLen
                If gLastFound(lbcBaseLen) >= 0 Then
                    lbcBaseLen.ListIndex = gLastFound(lbcBaseLen)
                    imComboBoxIndex = lbcBaseLen.ListIndex
                    edcSpecDropDown.Text = lbcBaseLen.List(lbcBaseLen.ListIndex)
                Else
                    lbcBaseLen.ListIndex = -1
                    imComboBoxIndex = lbcBaseLen.ListIndex
                    edcSpecDropDown.Text = ""
                End If
            End If
            imChgMode = False
            lbcBaseLen.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.Height
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:7/15/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus                      *
'*                                                     *
'*******************************************************
Private Sub mSpecSetFocus()

    If imSpecBoxNo >= imLBSpecCtrls And imSpecBoxNo <= UBound(tmSpecCtrls) Then
        Select Case imSpecBoxNo 'Branch on box type (control)
            Case NAMEINDEX 'Name
                edcSpecDropDown.SetFocus
            Case VEHICLEINDEX
                edcSpecDropDown.SetFocus
            Case YEARINDEX
                edcSpecDropDown.SetFocus
            Case STARTDATEINDEX 'Start date
                edcSpecDropDown.SetFocus
            Case NOGRIDSINDEX 'No Grid Levels
                edcSpecDropDown.SetFocus
            Case BASELENINDEX 'Base Length
                edcSpecDropDown.SetFocus
            Case ROUNDINDEX 'Round to nearset
                edcSpecDropDown.SetFocus
        End Select
        Exit Sub
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecSetShow                    *
'*                                                     *
'*             Created:7/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSpecSetShow(ilBoxNo As Integer)
'
'   mSpecSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slLen As String
    Dim ilVefCode As Integer
    If (ilBoxNo < imLBSpecCtrls) Or (ilBoxNo > UBound(tmSpecCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcSpecDropDown.Visible = False  'Set visibility
            slStr = edcSpecDropDown.Text
            tmRcf.sName = slStr
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case VEHICLEINDEX
            lbcVehicle.Visible = False
            edcSpecDropDown.Visible = False
            cmcSpecDropDown.Visible = False
            If lbcVehicle.ListIndex <= 0 Then
                ilVefCode = 0
            Else
                slNameCode = tgUserVehicle(lbcVehicle.ListIndex - 1).sKey 'Traffic!lbcUserVehicle.List(lbcVehicle.ListIndex)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mSpecSetShowErr
                gCPErrorMsg ilRet, "mmSpecSetShow (gParseItem field 2)", RCTerms
                On Error GoTo 0
                ilVefCode = CInt(slCode)
            End If
            If tmRcf.iVefCode <> ilVefCode Then
                tmRcf.iVefCode = ilVefCode
                mLenPop
                If imTerminate Then
                    Exit Sub
                End If
                If tmRcf.iBaseLen = 0 Then
                    'Set len: 60=Radio; 30=TV
                    If imVpfIndex >= 0 Then
                        If (tgVpf(imVpfIndex).sGMedium = "R") Or (tgVpf(imVpfIndex).sGMedium = "N") Then
                            slLen = "60"
                        Else
                            slLen = "30"
                        End If
                    Else
                        slLen = "30"
                    End If
                Else
                    slLen = Trim$(str$(tmRcf.iBaseLen))
                End If
                gFindMatch slLen, 0, lbcBaseLen
                imChgMode = True
                If gLastFound(lbcBaseLen) >= 0 Then
                    lbcBaseLen.ListIndex = gLastFound(lbcBaseLen)
                Else
                    lbcBaseLen.ListIndex = 0
                End If
                imChgMode = False
            End If
            slStr = edcSpecDropDown.Text
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case YEARINDEX
            edcSpecDropDown.Visible = False  'Set visibility
            slStr = edcSpecDropDown.Text
            If Trim$(slStr) <> "" Then
                tmRcf.iYear = Val(slStr)
                If (tmRcf.iYear >= 0) And (tmRcf.iYear <= 69) Then
                    tmRcf.iYear = 2000 + tmRcf.iYear
                ElseIf (tmRcf.iYear >= 70) And (tmRcf.iYear <= 99) Then
                    tmRcf.iYear = 1900 + tmRcf.iYear
                End If
            Else
                tmRcf.iYear = 0
            End If
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case STARTDATEINDEX 'Start Date
            plcCalendar.Visible = False
            cmcSpecDropDown.Visible = False
            edcSpecDropDown.Visible = False  'Set visibility
            slStr = edcSpecDropDown.Text
            If Not gValidDate(slStr) Then
                slStr = ""
            End If
            gPackDate slStr, tmRcf.iStartDate(0), tmRcf.iStartDate(1)
            slStr = gFormatDate(slStr)
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case NOGRIDSINDEX 'Number of Grids
            edcSpecDropDown.Visible = False  'Set visibility
            slStr = edcSpecDropDown.Text
            If Trim$(slStr) <> "" Then
                tmRcf.iGridsUsed = Val(slStr)
            Else
                tmRcf.iGridsUsed = 0
            End If
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case ROUNDINDEX 'Round to Nearest
            edcSpecDropDown.Visible = False  'Set visibility
            slStr = edcSpecDropDown.Text
            'gStrToPDN slStr, 2, 3, tmRcf.sRound
            tmRcf.lRound = gStrDecToLong(slStr, 2)
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case BASELENINDEX
            lbcBaseLen.Visible = False
            edcSpecDropDown.Visible = False
            cmcSpecDropDown.Visible = False
            slStr = edcSpecDropDown.Text
            If Trim$(slStr) <> "" Then
                tmRcf.iBaseLen = Val(slStr)
            '    ilRePaintLen = False
            '    For ilLoop = LBound(tmRcf.iLen) To UBound(tmRcf.iLen) Step 1
            '        If tmRcf.iLen(ilLoop) = tmRcf.iBaseLen Then
            '            If tmLenCtrls(ilLoop).sShow = "Base" Then
            '                Exit For
            '            End If
            '            slStr = "0"
            '            gStrToPDN slStr, 2, 5, tmRcf.sValue(ilLoop)
            '            slStr = "Base"
            '            gSetShow pbcLenInp, slStr, tmLenCtrls(ilLoop)
            '            ilRePaintLen = True
            '        Else
            '            If tmLenCtrls(ilLoop).sShow = "Base" Then
            '                tmLenCtrls(ilLoop).sShow = ""
            '                ilRePaintLen = True
            '            End If
            '        End If
            '    Next ilLoop
            '    If ilRePaintLen Then
            '        pbcLenInp.Cls
            '        pbcLenInp_Paint
            '    End If
            Else
                tmRcf.iBaseLen = 0
            End If
            slStr = Trim$(str$(tmRcf.iBaseLen))
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
    End Select
    Exit Sub
mSpecSetShowErr:
    On Error GoTo 0
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
    Unload RCTerms
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields() As Integer
'
'   iRet = mTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim slStr As String
    Dim slName As String
    Dim slVehName As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slRecCode As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilMonth As Integer
    Dim ilYear As Integer

    gUnpackDate tmRcf.iStartDate(0), tmRcf.iStartDate(1), slStr    'Last log date
    If Trim$(slStr) = "" Then
        Beep
        ilRes = MsgBox("Start date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imSpecBoxNo = STARTDATEINDEX
        mTestFields = NO
        Exit Function
    End If
    gObtainMonthYear 0, slStr, ilMonth, ilYear  'Standard month
    If tmRcf.iYear <> ilYear Then
        Beep
        ilRes = MsgBox("Start date year must match Rate Card Year", vbOKOnly + vbExclamation, "Incomplete")
        imSpecBoxNo = STARTDATEINDEX
        mTestFields = NO
        Exit Function
    End If
    If igRCMode = 1 Then
        mTestFields = YES
        Exit Function
    End If
    If Trim$(tmRcf.sName) = "" Then
        Beep
        ilRes = MsgBox("Rate Card Number must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imSpecBoxNo = NAMEINDEX
        mTestFields = NO
        Exit Function
    End If
    For ilLoop = LBound(tmRateCard) To UBound(tmRateCard) - 1 Step 1
        slNameCode = tmRateCard(ilLoop).sKey    'lbcRateCard.List(ilSelectIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slName)
        ilRet = gParseItem(slName, 1, "/", slName)
        slName = Left$(slName, Len(slName) - 4)
        If StrComp(Trim$(tmRcf.sName), Trim$(slName), 1) = 0 Then
            Beep
            MsgBox "Rate Card # already defined, enter a different #", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
            imSpecBoxNo = NAMEINDEX
            mTestFields = NO
            Exit Function
        End If
    Next ilLoop
    If tmRcf.iVefCode = -32000 Then
        Beep
        ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imSpecBoxNo = VEHICLEINDEX
        mTestFields = NO
        Exit Function
    End If
    slName = Trim$(tmRcf.sName)
    slVehName = ""
    If tmRcf.iVefCode = 0 Then
        slVehName = "[All Vehicles]"
    'ElseIf tmRcf.iVefCode > 0 Then
    Else
        slRecCode = Trim$(str$(tmRcf.iVefCode))
        For ilLoop = 0 To UBound(tgUserVehicle) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
            slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mTestFieldsErr
            gCPErrorMsg ilRet, "mTestFields (gParseItem field 2)", RCTerms
            On Error GoTo 0
            If slRecCode = slCode Then
                ilRet = gParseItem(slNameCode, 1, "\", slVehName)
                ilRet = gParseItem(slVehName, 3, "|", slVehName)
                On Error GoTo mTestFieldsErr
                gCPErrorMsg ilRet, "mTestFields (gParseItem field 1)", RCTerms
                On Error GoTo 0
                Exit For
            End If
        Next ilLoop
    'ElseIf tmRcf.iVefCode < 0 Then
    '    slRecCode = Trim$(Str$(-tmRcf.iVefCode))
    '    For ilLoop = 0 To lbcCombo.ListCount - 1 Step 1
    '        slNameCode = lbcCombo.List(ilLoop)
    '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '        On Error GoTo mTestFieldsErr
    '        gCPErrorMsg ilRet, "mTestFields (gParseItem field 2)", RCTerms
    '        On Error GoTo 0
    '        If slRecCode = slCode Then
    '            ilRet = gParseItem(slNameCode, 1, "\", slVehName)
    '            On Error GoTo mTestFieldsErr
    '            gCPErrorMsg ilRet, "mTestFields (gParseItem field 1)", RCTerms
    '            On Error GoTo 0
    '            Exit For
    '        End If
    '    Next ilLoop
    End If
    slName = slName & "\" & Trim$(slVehName)
    gFindMatch slName, 0, RateCard!cbcSelect    'Determine if name exist
    If gLastFound(RateCard!cbcSelect) <> -1 Then   'Name found
        If gLastFound(RateCard!cbcSelect) <> imSelectedIndex Then
            slStr = tmRcf.sName
            If slStr = RateCard!cbcSelect.List(gLastFound(RateCard!cbcSelect)) Then
                Beep
                MsgBox "Rate Card already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                imSpecBoxNo = NAMEINDEX
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    gUnpackDate tmRcf.iStartDate(0), tmRcf.iStartDate(1), slStr    'Last log date
    If Trim$(slStr) = "" Then
        Beep
        ilRes = MsgBox("Start date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imSpecBoxNo = STARTDATEINDEX
        mTestFields = NO
        Exit Function
    End If
    gUnpackDateForSort tmRcf.iStartDate(0), tmRcf.iStartDate(1), slName
    slVehName = Trim$(slVehName)
    'For ilLoop = 0 To UBound(tmRateCard) - 1 Step 1 'RateCard!lbcRateCard.ListCount - 1 Step 1
    '    'ilPos = InStr(RateCard!lbcRateCard.List(ilLoop), slVehName)
    '    ilPos = InStr(tmRateCard(ilLoop).sKey, slVehName)
    '    If ilPos > 0 Then
    '        slNameCode = tmRateCard(ilLoop).sKey   'RateCard!lbcRateCard.List(ilLoop)
    '        ilRet = gParseItem(slNameCode, 1, "\", slCode)
    '        On Error GoTo mTestFieldsErr
    '        gCPErrorMsg ilRet, "mTestFields (gParseItem field 1)", RCTerms
    '        On Error GoTo 0
    '        slCode = gSubStr("99999", slCode)
    '        If StrComp(slName, slCode, 1) <= 0 Then
    '            Beep
    '            ilRes = MsgBox("Invalid Start date specified", vbOkOnly + vbExclamation, "Incomplete")
    '            imSpecBoxNo = STARTDATEINDEX
    '            mTestFields = No
    '            Exit Function
    '        End If
    '    End If
    'Next ilLoop
    mTestFields = YES
    Exit Function
mTestFieldsErr:
    On Error GoTo 0
    mTestFields = NO
    Beep
    MsgBox "Invalid Rate Card specified, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
    imSpecBoxNo = NAMEINDEX
    mTestFields = NO
    Exit Function
End Function
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
Private Sub mVehPop(lbcCtrl As control)
    Dim ilRet As Integer
    'ilRet = gPopUserVehicleBox(RCTerms, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcCtrl, Traffic!lbcUserVehicle)
    'ilRet = gPopUserVehicleBox(RCTerms, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcCtrl, tgUserVehicle(), sgUserVehicleTag)
    ilRet = gPopUserVehicleBox(RCTerms, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH, lbcCtrl, tgUserVehicle(), sgUserVehicleTag)
    'ilRet = gPopUserVehComboBox(RCTerms, lbcCtrl, Traffic!lbcUserVehicle, lbcCombo)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        'gCPErrorMsg ilRet, "mVehPop (gPopUserVehComboBox: Vehicle/Combo)", RCTerms
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", RCTerms
        On Error GoTo 0
        lbcCtrl.AddItem "[All Vehicle]", 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub pbcBySpotType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcBySpotType_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("I")) Or (KeyAscii = Asc("i")) Then
        If imBoxNo = 1 Then
            tmRcf.sUseBB = "I"
        ElseIf imBoxNo = 3 Then
            tmRcf.sUseFP = "I"
        ElseIf imBoxNo = 5 Then
            tmRcf.sUseNP = "I"
        ElseIf imBoxNo = 7 Then
            tmRcf.sUsePrefDT = "I"
        ElseIf imBoxNo = 9 Then
            tmRcf.sUse1stPos = "I"
        ElseIf imBoxNo = 11 Then
            tmRcf.sUseSoloAvail = "I"
        End If
        pbcBySpotType_Paint
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If imBoxNo = 1 Then
            tmRcf.sUseBB = "D"
        ElseIf imBoxNo = 3 Then
            tmRcf.sUseFP = "D"
        ElseIf imBoxNo = 5 Then
            tmRcf.sUseNP = "D"
        ElseIf imBoxNo = 7 Then
            tmRcf.sUsePrefDT = "D"
        ElseIf imBoxNo = 9 Then
            tmRcf.sUse1stPos = "D"
        ElseIf imBoxNo = 11 Then
            tmRcf.sUseSoloAvail = "D"
        End If
        pbcBySpotType_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imBoxNo = 1 Then
            If tmRcf.sUseBB = "I" Then
                tmRcf.sUseBB = "D"
                pbcBySpotType_Paint
            ElseIf tmRcf.sUseBB = "D" Then
                tmRcf.sUseBB = "I"
                pbcBySpotType_Paint
            End If
        ElseIf imBoxNo = 3 Then
            If tmRcf.sUseFP = "I" Then
                tmRcf.sUseFP = "D"
                pbcBySpotType_Paint
            ElseIf tmRcf.sUseFP = "D" Then
                tmRcf.sUseFP = "I"
                pbcBySpotType_Paint
            End If
        ElseIf imBoxNo = 5 Then
            If tmRcf.sUseNP = "I" Then
                tmRcf.sUseNP = "D"
                pbcBySpotType_Paint
            ElseIf tmRcf.sUseNP = "D" Then
                tmRcf.sUseNP = "I"
                pbcBySpotType_Paint
            End If
        ElseIf imBoxNo = 7 Then
            If tmRcf.sUsePrefDT = "I" Then
                tmRcf.sUsePrefDT = "D"
                pbcBySpotType_Paint
            ElseIf tmRcf.sUsePrefDT = "D" Then
                tmRcf.sUsePrefDT = "I"
                pbcBySpotType_Paint
            End If
        ElseIf imBoxNo = 9 Then
            If tmRcf.sUse1stPos = "I" Then
                tmRcf.sUse1stPos = "D"
                pbcBySpotType_Paint
            ElseIf tmRcf.sUse1stPos = "D" Then
                tmRcf.sUse1stPos = "I"
                pbcBySpotType_Paint
            End If
        ElseIf imBoxNo = 11 Then
            If tmRcf.sUseSoloAvail = "I" Then
                tmRcf.sUseSoloAvail = "D"
                pbcBySpotType_Paint
            ElseIf tmRcf.sUseSoloAvail = "D" Then
                tmRcf.sUseSoloAvail = "I"
                pbcBySpotType_Paint
            End If
       End If
    End If
End Sub
Private Sub pbcBySpotType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = 1 Then
        If tmRcf.sUseBB = "I" Then
            tmRcf.sUseBB = "D"
        Else
            tmRcf.sUseBB = "I"
        End If
    ElseIf imBoxNo = 3 Then
        If tmRcf.sUseFP = "I" Then
            tmRcf.sUseFP = "D"
        Else
            tmRcf.sUseFP = "I"
        End If
    ElseIf imBoxNo = 5 Then
        If tmRcf.sUseNP = "I" Then
            tmRcf.sUseNP = "D"
        Else
            tmRcf.sUseNP = "I"
        End If
    ElseIf imBoxNo = 7 Then
        If tmRcf.sUsePrefDT = "I" Then
            tmRcf.sUsePrefDT = "D"
        Else
            tmRcf.sUsePrefDT = "I"
        End If
    ElseIf imBoxNo = 9 Then
        If tmRcf.sUse1stPos = "I" Then
            tmRcf.sUse1stPos = "D"
        Else
            tmRcf.sUse1stPos = "I"
        End If
    ElseIf imBoxNo = 11 Then
        If tmRcf.sUseSoloAvail = "I" Then
            tmRcf.sUseSoloAvail = "D"
        Else
            tmRcf.sUseSoloAvail = "I"
        End If
    End If
    pbcBySpotType_Paint
End Sub
Private Sub pbcBySpotType_Paint()
    Dim slStr As String
    pbcBySpotType.Cls
    pbcBySpotType.CurrentX = fgBoxInsetX
    pbcBySpotType.CurrentY = 0 'fgBoxInsetY
    If imBoxNo = 1 Then
        slStr = tmRcf.sUseBB
    ElseIf imBoxNo = 3 Then
        slStr = tmRcf.sUseFP
    ElseIf imBoxNo = 5 Then
        slStr = tmRcf.sUseNP
    ElseIf imBoxNo = 7 Then
        slStr = tmRcf.sUsePrefDT
    ElseIf imBoxNo = 9 Then
        slStr = tmRcf.sUse1stPos
    ElseIf imBoxNo = 11 Then
        slStr = tmRcf.sUseSoloAvail
    End If
    If slStr = "I" Then
        pbcBySpotType.Print "Index"
    Else
        pbcBySpotType.Print "Dollar"
    End If

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
                edcSpecDropDown.Text = Format$(llDate, "m/d/yy")
                edcSpecDropDown.SelStart = 0
                edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
                imBypassFocus = True
                edcSpecDropDown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcSpecDropDown.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetAll
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcDayRate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBDRCtrls To UBound(tmDRCtrls) Step 1
        If (X >= tmDRCtrls(ilBox).fBoxX) And (X <= tmDRCtrls(ilBox).fBoxX + tmDRCtrls(ilBox).fBoxW) Then
            If (Y >= tmDRCtrls(ilBox).fBoxY) And (Y <= tmDRCtrls(ilBox).fBoxY + tmDRCtrls(ilBox).fBoxH) Then
                mSetAll
                imType = DAYRATE
                imBoxNo = ilBox
                mEnableBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus
End Sub
Private Sub pbcDayRate_Paint()
    Dim ilBox As Integer
    For ilBox = imLBDRCtrls To UBound(tmDRCtrls) Step 1
        gPaintArea pbcDayRate, tmDRCtrls(ilBox).fBoxX, tmDRCtrls(ilBox).fBoxY, tmDRCtrls(ilBox).fBoxW - 15, tmDRCtrls(ilBox).fBoxH - 15, WHITE
        pbcDayRate.CurrentX = tmDRCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcDayRate.CurrentY = tmDRCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
        pbcDayRate.Print tmDRCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBGDCtrls To UBound(tmGDCtrls) Step 1
        If (X >= tmGDCtrls(ilBox).fBoxX) And (X <= tmGDCtrls(ilBox).fBoxX + tmGDCtrls(ilBox).fBoxW) Then
            If (Y >= tmGDCtrls(ilBox).fBoxY) And (Y <= tmGDCtrls(ilBox).fBoxY + tmGDCtrls(ilBox).fBoxH) Then
                mSetAll
                imType = GRID
                imBoxNo = ilBox
                mEnableBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus
End Sub
Private Sub pbcGrid_Paint()
    Dim ilBox As Integer
    For ilBox = imLBGDCtrls To UBound(tmGDCtrls) Step 1
        gPaintArea pbcGrid, tmGDCtrls(ilBox).fBoxX, tmGDCtrls(ilBox).fBoxY, tmGDCtrls(ilBox).fBoxW - 15, tmGDCtrls(ilBox).fBoxH - 15, WHITE
        pbcGrid.CurrentX = tmGDCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcGrid.CurrentY = tmGDCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
        pbcGrid.Print tmGDCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcHourRate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBHRCtrls To UBound(tmHRCtrls) Step 1
        If (X >= tmHRCtrls(ilBox).fBoxX) And (X <= tmHRCtrls(ilBox).fBoxX + tmHRCtrls(ilBox).fBoxW) Then
            If (Y >= tmHRCtrls(ilBox).fBoxY) And (Y <= tmHRCtrls(ilBox).fBoxY + tmHRCtrls(ilBox).fBoxH) Then
                mSetAll
                imType = HOURRATE
                imBoxNo = ilBox
                mEnableBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus
End Sub
Private Sub pbcHourRate_Paint()
    Dim ilBox As Integer
    For ilBox = imLBHRCtrls To UBound(tmHRCtrls) Step 1
        gPaintArea pbcHourRate, tmHRCtrls(ilBox).fBoxX, tmHRCtrls(ilBox).fBoxY, tmHRCtrls(ilBox).fBoxW - 15, tmHRCtrls(ilBox).fBoxH - 15, WHITE
        pbcHourRate.CurrentX = tmHRCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcHourRate.CurrentY = tmHRCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
        pbcHourRate.Print tmHRCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcLength_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBLenCtrls To UBound(tmLenCtrls) Step 1
        If (X >= tmLenCtrls(ilBox).fBoxX) And (X <= tmLenCtrls(ilBox).fBoxX + tmLenCtrls(ilBox).fBoxW) Then
            If (Y >= tmLenCtrls(ilBox).fBoxY) And (Y <= tmLenCtrls(ilBox).fBoxY + tmLenCtrls(ilBox).fBoxH) Then
                mSetAll
                imType = Length
                imBoxNo = ilBox
                mEnableBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus
End Sub
Private Sub pbcLength_Paint()
    Dim ilBox As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    llColor = pbcLength.ForeColor
    slFontName = pbcLength.FontName
    flFontSize = pbcLength.FontSize
    pbcLength.ForeColor = BLUE
    pbcLength.FontBold = False
    pbcLength.FontSize = 7
    pbcLength.FontName = "Arial"
    pbcLength.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    For ilBox = imLBLenCtrls To UBound(tmLenCtrls) Step 2
        gPaintArea pbcLength, tmLenCtrls(ilBox).fBoxX, tmLenCtrls(ilBox).fBoxY, tmLenCtrls(ilBox).fBoxW - 15, tmLenCtrls(ilBox).fBoxH - 15, WHITE
        pbcLength.CurrentX = gCenterShowStr(pbcLength, tmLenCtrls(ilBox).sShow, tmLenCtrls(ilBox))   'tmLenCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcLength.CurrentY = tmLenCtrls(ilBox).fBoxY '- 30'+ fgBoxInsetY
        pbcLength.Print tmLenCtrls(ilBox).sShow
    Next ilBox
    pbcLength.FontSize = flFontSize
    pbcLength.FontName = slFontName
    pbcLength.FontSize = flFontSize
    pbcLength.ForeColor = llColor
    pbcLength.FontBold = True
    For ilBox = imLBLenCtrls + 1 To UBound(tmLenCtrls) Step 2
        gPaintArea pbcLength, tmLenCtrls(ilBox).fBoxX, tmLenCtrls(ilBox).fBoxY, tmLenCtrls(ilBox).fBoxW - 15, tmLenCtrls(ilBox).fBoxH - 15, WHITE
        pbcLength.CurrentX = tmLenCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcLength.CurrentY = tmLenCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
        pbcLength.Print tmLenCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        If (X >= tmSpecCtrls(ilBox).fBoxX) And (X <= tmSpecCtrls(ilBox).fBoxX + tmSpecCtrls(ilBox).fBoxW) Then
            If (Y >= tmSpecCtrls(ilBox).fBoxY) And (Y <= tmSpecCtrls(ilBox).fBoxY + tmSpecCtrls(ilBox).fBoxH) Then
                If igRCMode = 1 Then 'Change
                    If (ilBox <> BASELENINDEX) And (ilBox <> ROUNDINDEX) And (ilBox <> NOGRIDSINDEX) And (ilBox <> STARTDATEINDEX) Then
                        Beep
                        mSpecSetFocus
                        Exit Sub
                    End If
                End If
                mSetShow
                mSpecSetShow imSpecBoxNo
                imSpecBoxNo = ilBox
                mSpecEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSpecSetFocus
End Sub
Private Sub pbcSpec_Paint()
    Dim ilBox As Integer
    For ilBox = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        pbcSpec.CurrentX = tmSpecCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcSpec.CurrentY = tmSpecCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcSpec.Print tmSpecCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcSpecSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcSpecSTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = -1  'Right to left
    Select Case imSpecBoxNo
        Case -1
            imTabDirection = 0  'left to right
            If igRCMode = 0 Then 'New
                ilBox = NAMEINDEX
            Else
                ilBox = NOGRIDSINDEX
            End If
        Case 1 'Name (first control within header)
            mSpecSetShow imSpecBoxNo
            imSpecBoxNo = -1
            cmcDone.SetFocus
            Exit Sub
        Case STARTDATEINDEX
            slStr = edcSpecDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcSpecDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcSpecDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imSpecBoxNo - 1
        Case Else
            If (imSpecBoxNo = NOGRIDSINDEX) And (igRCMode = 1) Then
                mSpecSetShow imSpecBoxNo
                imSpecBoxNo = -1
                cmcDone.SetFocus
                Exit Sub
            End If
            ilBox = imSpecBoxNo - 1
    End Select
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = ilBox
    mSpecEnableBox ilBox
End Sub
Private Sub pbcSpecTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcSpecTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = 0  'left to right
    Select Case imSpecBoxNo
        Case -1
            imTabDirection = -1  'Right to left
            ilBox = BASELENINDEX
        Case STARTDATEINDEX
            slStr = edcSpecDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcSpecDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcSpecDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imSpecBoxNo + 1
        Case BASELENINDEX
            mSpecSetShow imSpecBoxNo
            imSpecBoxNo = -1
            pbcSTab.SetFocus
            Exit Sub
        Case Else
            ilBox = imSpecBoxNo + 1
    End Select
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = ilBox
    mSpecEnableBox ilBox
End Sub
Private Sub pbcSpotFreq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBSFCtrls To UBound(tmSFCtrls) Step 1
        If (X >= tmSFCtrls(ilBox).fBoxX) And (X <= tmSFCtrls(ilBox).fBoxX + tmSFCtrls(ilBox).fBoxW) Then
            If (Y >= tmSFCtrls(ilBox).fBoxY) And (Y <= tmSFCtrls(ilBox).fBoxY + tmSFCtrls(ilBox).fBoxH) Then
                mSetAll
                imType = SPOTFREQ
                imBoxNo = ilBox
                mEnableBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus
End Sub
Private Sub pbcSpotFreq_Paint()
    Dim ilBox As Integer
    For ilBox = imLBSFCtrls To UBound(tmSFCtrls) Step 1
        gPaintArea pbcSpotFreq, tmSFCtrls(ilBox).fBoxX, tmSFCtrls(ilBox).fBoxY, tmSFCtrls(ilBox).fBoxW - 15, tmSFCtrls(ilBox).fBoxH - 15, WHITE
        pbcSpotFreq.CurrentX = tmSFCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcSpotFreq.CurrentY = tmSFCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
        pbcSpotFreq.Print tmSFCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcSpotType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                       slStr                                                   *
'******************************************************************************************

    Dim ilBox As Integer

    For ilBox = imLBSTCtrls To UBound(tmSTCtrls) Step 1
        If (X >= tmSTCtrls(ilBox).fBoxX) And (X <= tmSTCtrls(ilBox).fBoxX + tmSTCtrls(ilBox).fBoxW) Then
            If (Y >= tmSTCtrls(ilBox).fBoxY) And (Y <= tmSTCtrls(ilBox).fBoxY + tmSTCtrls(ilBox).fBoxH) Then
                If (ilBox - 1) Mod STVALUEINDEX + 1 = 1 Then
                    Beep
                    Exit Sub
                End If
                If (ilBox - 1) Mod STVALUEINDEX + 1 = 2 Then
                    mSetSpotTypeIndex ilBox - 1
                End If
                mSetAll
                imType = SPOTTYPE
                imBoxNo = ilBox
                mEnableBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus
End Sub
Private Sub pbcSpotType_Paint()
    Dim ilBox As Integer
    For ilBox = imLBSTCtrls To UBound(tmSTCtrls) Step 1
        gPaintArea pbcBySpotType, tmSTCtrls(ilBox).fBoxX, tmSTCtrls(ilBox).fBoxY, tmSTCtrls(ilBox).fBoxW - 15, tmSTCtrls(ilBox).fBoxH - 15, WHITE
        pbcSpotType.CurrentX = tmSTCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcSpotType.CurrentY = tmSTCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
        pbcSpotType.Print tmSTCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcSTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************

    Dim ilType As Integer
    Dim ilBoxNo As Integer
    Dim ilFound As Integer

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    ilType = imType
    ilBoxNo = imBoxNo
    ilFound = False
    Do
        Select Case ilType
            Case -1
                ilBoxNo = -1
                ilType = SPOTTYPE
            Case SPOTTYPE
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = STBYINDEX
                        mSetSpotTypeIndex ilBoxNo
                        ilBoxNo = ilBoxNo + 1
                        ilFound = True
                    Case STBYINDEX
                        mSetShow
                        imType = -1
                        imBoxNo = -1
                        cmcDone.SetFocus
                        Exit Sub
                    Case Else
                        ilBoxNo = ilBoxNo - 1
                        ilFound = True
                        If (ilBoxNo - 1) Mod STVALUEINDEX + 1 = 1 Then
                            mSetSpotTypeIndex ilBoxNo
                            ilFound = False
                        End If
                End Select
            Case Length
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = LENVALUEINDEX
                        ilFound = True
                    Case LENVALUEINDEX
                        mSetShow
                        imType = SPOTTYPE
                        imBoxNo = -1
                        pbcTab.SetFocus
                        Exit Sub
                    Case Else
                        ilBoxNo = ilBoxNo - 2
                        ilFound = True
                End Select
            Case SPOTFREQ
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = SPOTFROMINDEX
                        ilFound = True
                    Case SPOTFROMINDEX
                        mSetShow
                        imType = Length
                        imBoxNo = -1
                        pbcTab.SetFocus
                        Exit Sub
                    Case Else
                        ilBoxNo = ilBoxNo - 1
                        ilFound = True
                End Select
            Case WEEKFREQ
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = WEEKFROMINDEX
                        ilFound = True
                    Case WEEKFROMINDEX
                        mSetShow
                        imType = SPOTFREQ
                        imBoxNo = -1
                        pbcTab.SetFocus
                        Exit Sub
                    Case Else
                        ilBoxNo = ilBoxNo - 1
                        ilFound = True
                End Select
            Case GRID
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = GRIDVALUEINDEX
                        ilFound = True
                    Case GRIDVALUEINDEX
                        mSetShow
                        imType = SPOTFREQ  'WEEKFREQ
                        imBoxNo = -1
                        pbcTab.SetFocus
                        Exit Sub
                    Case Else
                        ilBoxNo = ilBoxNo - 1
                        ilFound = True
                End Select
            Case DAYRATE
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = DRDAYINDEX
                        ilFound = True
                    Case DRDAYINDEX
                        mSetShow
                        imType = GRID
                        imBoxNo = -1
                        pbcTab.SetFocus
                        Exit Sub
                    Case Else
                        ilBoxNo = ilBoxNo - 1
                        ilFound = True
                End Select
            Case HOURRATE
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = HRSTARTINDEX
                        ilFound = True
                    Case HRSTARTINDEX
                        mSetShow
                        imType = DAYRATE
                        imBoxNo = -1
                        pbcTab.SetFocus
                        Exit Sub
                    Case Else
                        ilBoxNo = ilBoxNo - 1
                        ilFound = True
                End Select
        End Select
    Loop Until ilFound
    mSetShow
    imType = ilType
    imBoxNo = ilBoxNo
    mEnableBox
End Sub
Private Sub pbcTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************

    Dim ilType As Integer
    Dim ilBoxNo As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    mSetShow
    ilType = imType
    ilBoxNo = imBoxNo
    ilFound = False
    Do
        Select Case ilType
            Case -1
                ilBoxNo = -1
                If pbcHourRate.Visible Then
                    ilType = HOURRATE
                Else
                    ilType = DAYRATE
                End If
            Case SPOTTYPE
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = STBYINDEX
                        For ilLoop = UBound(tmSTCtrls) To STVALUEINDEX Step -1
                            If Trim$(tmSTCtrls(ilLoop).sShow) <> "" Then
                                ilBoxNo = ilLoop
                                Exit For
                            End If
                        Next ilLoop
                        If (ilBoxNo - 1) Mod STVALUEINDEX + 1 = 1 Then
                            mSetSpotTypeIndex ilBoxNo
                            ilBoxNo = ilBoxNo + 1
                        End If
                        ilFound = True
                    Case UBound(tmSTCtrls)
                        imType = Length 'SPOTFREQ
                        imBoxNo = -1
                        pbcSTab.SetFocus
                        Exit Sub
                    Case Else
                        ilFound = True
                        ilBoxNo = ilBoxNo + 1
                        If (ilBoxNo - 1) Mod STVALUEINDEX + 1 = 1 Then
                            mSetSpotTypeIndex ilBoxNo
                            ilFound = False
                        End If
                End Select
            Case Length
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = LENVALUEINDEX
                        For ilLoop = UBound(tmLenCtrls) To LENVALUEINDEX Step -2
                            If Trim$(tmLenCtrls(ilLoop - 1).sShow) <> "" Then  'Check for length
                                ilBoxNo = ilLoop
                                Exit For
                            End If
                        Next ilLoop
                        ilFound = True
                    Case UBound(tmLenCtrls), UBound(tmLenCtrls) + 1
                        imType = SPOTFREQ 'GRID
                        imBoxNo = -1
                        pbcSTab.SetFocus
                        Exit Sub
                    Case Else
                        'If next length is blank- bypass
                        If Trim$(tmLenCtrls(ilBoxNo + 1).sShow) <> "" Then
                            ilFound = True
                        End If
                        ilBoxNo = ilBoxNo + 2
                End Select
            Case SPOTFREQ
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = SPOTFROMINDEX
                        For ilLoop = UBound(tmSFCtrls) - SPOTVALUEINDEX + 1 To SPOTFROMINDEX Step -SPOTVALUEINDEX
                            If Trim$(tmSFCtrls(ilLoop).sShow) <> "" Then
                                ilBoxNo = ilLoop + SPOTVALUEINDEX - 1
                                Exit For
                            End If
                        Next ilLoop
                        ilFound = True
                    Case UBound(tmSFCtrls)
                        imType = GRID   'WEEKFREQ
                        imBoxNo = -1
                        pbcSTab.SetFocus
                        Exit Sub
                    Case Else
                        If (ilBoxNo - 1) Mod SPOTVALUEINDEX = 0 Then
                            For ilLoop = ilBoxNo To UBound(tmSFCtrls) Step 1
                                If Trim$(tmSFCtrls(ilLoop).sShow) <> "" Then
                                    ilFound = True
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                imType = GRID   'WEEKFREQ
                                imBoxNo = -1
                                pbcSTab.SetFocus
                                Exit Sub
                            End If
                        Else
                            ilFound = True
                        End If
                        ilBoxNo = ilBoxNo + 1
                End Select
            Case WEEKFREQ
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = WEEKFROMINDEX
                        For ilLoop = UBound(tmWFCtrls) - WEEKVALUEINDEX + 1 To WEEKFROMINDEX Step -WEEKVALUEINDEX
                            If Trim$(tmWFCtrls(ilLoop).sShow) <> "" Then
                                ilBoxNo = ilLoop + WEEKVALUEINDEX - 1
                                Exit For
                            End If
                        Next ilLoop
                        ilFound = True
                    Case UBound(tmWFCtrls)
                        imType = GRID   'LENGTH
                        imBoxNo = -1
                        pbcSTab.SetFocus
                        Exit Sub
                    Case Else
                        If (ilBoxNo - 1) Mod WEEKVALUEINDEX = 0 Then
                            For ilLoop = ilBoxNo To UBound(tmWFCtrls) Step 1
                                If Trim$(tmWFCtrls(ilLoop).sShow) <> "" Then
                                    ilFound = True
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                imType = GRID
                                imBoxNo = -1
                                pbcSTab.SetFocus
                                Exit Sub
                            End If
                        Else
                            ilFound = True
                        End If
                        ilBoxNo = ilBoxNo + 1
                End Select
            Case GRID
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = GRIDVALUEINDEX
                        For ilLoop = UBound(tmGDCtrls) To GRIDVALUEINDEX Step -1
                            If Trim$(tmGDCtrls(ilLoop).sShow) <> "" Then
                                ilBoxNo = ilLoop
                                Exit For
                            End If
                        Next ilLoop
                        ilFound = True
                    Case UBound(tmGDCtrls)
                        imType = DAYRATE
                        imBoxNo = -1
                        pbcSTab.SetFocus
                        Exit Sub
                    Case Else
                        ilFound = True
                        ilBoxNo = ilBoxNo + 1
                End Select
            Case DAYRATE
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = DRDAYINDEX
                        For ilLoop = UBound(tmDRCtrls) - DRVALUEINDEX + 1 To DRDAYINDEX Step -DRVALUEINDEX
                            If Trim$(tmDRCtrls(ilLoop).sShow) <> "" Then
                                ilBoxNo = ilLoop + DRVALUEINDEX - 1
                                Exit For
                            End If
                        Next ilLoop
                        ilFound = True
                    Case UBound(tmDRCtrls)
                        If pbcHourRate.Visible Then
                            imType = HOURRATE
                            imBoxNo = -1
                            pbcSTab.SetFocus
                            Exit Sub
                        Else
                            imType = -1
                            imBoxNo = -1
                            cmcDone.SetFocus
                            Exit Sub
                        End If
                    Case Else
                        If (ilBoxNo - 1) Mod DRVALUEINDEX + 1 = 8 Then
                            For ilLoop = ilBoxNo To UBound(tmDRCtrls) Step 1
                                If Trim$(tmDRCtrls(ilLoop).sShow) <> "" Then
                                    ilFound = True
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                If pbcHourRate.Visible Then
                                    imType = HOURRATE
                                    imBoxNo = -1
                                    pbcSTab.SetFocus
                                    Exit Sub
                                Else
                                    imType = -1
                                    imBoxNo = -1
                                    cmcDone.SetFocus
                                    Exit Sub
                                End If
                            End If
                        Else
                            ilFound = True
                        End If
                        ilBoxNo = ilBoxNo + 1
                End Select
            Case HOURRATE
                Select Case ilBoxNo
                    Case -1
                        ilBoxNo = HRSTARTINDEX
                        For ilLoop = UBound(tmHRCtrls) - HRVALUEINDEX + 1 To HRSTARTINDEX Step -HRVALUEINDEX
                            If Trim$(tmHRCtrls(ilLoop).sShow) <> "" Then
                                ilBoxNo = ilLoop + HRVALUEINDEX - 1
                                Exit For
                            End If
                        Next ilLoop
                        ilFound = True
                    Case UBound(tmHRCtrls)
                        imType = -1
                        imBoxNo = -1
                        cmcDone.SetFocus
                        Exit Sub
                    Case Else
                        If (ilBoxNo - 1) Mod HRVALUEINDEX = 0 Then
                            For ilLoop = ilBoxNo To UBound(tmHRCtrls) Step 1
                                If Trim$(tmHRCtrls(ilLoop).sShow) <> "" Then
                                    ilFound = True
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                imType = -1
                                imBoxNo = -1
                                cmcDone.SetFocus
                                Exit Sub
                            End If
                        Else
                            ilFound = True
                        End If
                        ilBoxNo = ilBoxNo + 1
                End Select
        End Select
    Loop Until ilFound
    imType = ilType
    imBoxNo = ilBoxNo
    mEnableBox
End Sub
Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    Dim ilBox As Integer
    Dim ilIndex As Integer
    If imType = HOURRATE Then
        ilIndex = (imBoxNo - 1) \ HRVALUEINDEX + 1
        ilBox = (imBoxNo - 1) Mod HRVALUEINDEX + 1
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
    End If
End Sub
Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    Dim ilBox As Integer
    Dim ilIndex As Integer
    If imType = HOURRATE Then
        ilIndex = (imBoxNo - 1) \ HRVALUEINDEX + 1
        ilBox = (imBoxNo - 1) Mod HRVALUEINDEX + 1
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
                        Select Case ilBox
                            Case 1
                                imBypassFocus = True    'Don't change select text
                                edcDropDown.SetFocus
                                'SendKeys slKey
                                gSendKeys edcDropDown, slKey
                            Case 2
                                imBypassFocus = True    'Don't change select text
                                edcDropDown.SetFocus
                                'SendKeys slKey
                                gSendKeys edcDropDown, slKey
                        End Select
                        Exit Sub
                    End If
                    flX = flX + fgPadDeltaX
                Next ilColNo
            End If
            flY = flY + fgPadDeltaY
        Next ilRowNo
    End If
End Sub
Private Sub pbcWeekFreq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBWFCtrls To UBound(tmWFCtrls) Step 1
        If (X >= tmWFCtrls(ilBox).fBoxX) And (X <= tmWFCtrls(ilBox).fBoxX + tmWFCtrls(ilBox).fBoxW) Then
            If (Y >= tmWFCtrls(ilBox).fBoxY) And (Y <= tmWFCtrls(ilBox).fBoxY + tmWFCtrls(ilBox).fBoxH) Then
                mSetAll
                imType = WEEKFREQ
                imBoxNo = ilBox
                mEnableBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus
End Sub
Private Sub pbcWeekFreq_Paint()
    Dim ilBox As Integer
    For ilBox = imLBWFCtrls To UBound(tmWFCtrls) Step 1
        gPaintArea pbcWeekFreq, tmWFCtrls(ilBox).fBoxX, tmWFCtrls(ilBox).fBoxY, tmWFCtrls(ilBox).fBoxW - 15, tmWFCtrls(ilBox).fBoxH - 15, WHITE
        pbcWeekFreq.CurrentX = tmWFCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcWeekFreq.CurrentY = tmWFCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
        pbcWeekFreq.Print tmWFCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcTerms_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Rate Card Terms"
End Sub

Private Sub mSetSpotTypeIndex(ilBoxNo As Integer)
    Dim slStr As String
    Dim ilIndex As Integer

    ilIndex = (ilBoxNo - 1) \ STVALUEINDEX + 1
    If ilIndex = 1 Then
        If (tmRcf.sUseBB <> "I") And (tmRcf.sUseBB <> "D") Or (Trim$(tmSTCtrls(ilBoxNo).sShow) = "") Then
            tmRcf.sUseBB = "I"
            slStr = "Index"
            gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
            pbcSpotType.Cls
            pbcSpotType_Paint
        End If
    ElseIf ilIndex = 2 Then
        If (tmRcf.sUseFP <> "I") And (tmRcf.sUseFP <> "D") Or (Trim$(tmSTCtrls(ilBoxNo).sShow) = "") Then
            tmRcf.sUseFP = "I"
            slStr = "Index"
            gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
            pbcSpotType.Cls
            pbcSpotType_Paint
        End If
    ElseIf ilIndex = 3 Then
        If (tmRcf.sUseNP <> "I") And (tmRcf.sUseNP <> "D") Or (Trim$(tmSTCtrls(ilBoxNo).sShow) = "") Then
            tmRcf.sUseNP = "I"
            slStr = "Index"
            gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
            pbcSpotType.Cls
            pbcSpotType_Paint
        End If
    ElseIf ilIndex = 4 Then
        If (tmRcf.sUsePrefDT <> "I") And (tmRcf.sUsePrefDT <> "D") Or (Trim$(tmSTCtrls(ilBoxNo).sShow) = "") Then
            tmRcf.sUsePrefDT = "I"
            slStr = "Index"
            gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
            pbcSpotType.Cls
            pbcSpotType_Paint
        End If
    ElseIf ilIndex = 5 Then
        If (tmRcf.sUse1stPos <> "I") And (tmRcf.sUse1stPos <> "D") Or (Trim$(tmSTCtrls(ilBoxNo).sShow) = "") Then
            tmRcf.sUse1stPos = "I"
            slStr = "Index"
            gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
            pbcSpotType.Cls
            pbcSpotType_Paint
        End If
    ElseIf ilIndex = 6 Then
        If (tmRcf.sUseSoloAvail <> "I") And (tmRcf.sUseSoloAvail <> "D") Or (Trim$(tmSTCtrls(ilBoxNo).sShow) = "") Then
            tmRcf.sUseSoloAvail = "I"
            slStr = "Index"
            gSetShow pbcSpotType, slStr, tmSTCtrls(ilBoxNo)
            pbcSpotType.Cls
            pbcSpotType_Paint
        End If
    End If
End Sub
