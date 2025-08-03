VERSION 5.00
Begin VB.Form FeedPlge 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5100
   ClientLeft      =   855
   ClientTop       =   1470
   ClientWidth     =   7185
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5100
   ScaleWidth      =   7185
   Begin VB.PictureBox plcLen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   675
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2895
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcLen 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "FeedPlge.frx":0000
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcLenOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcLenInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "FeedPlge.frx":0CBE
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.TextBox edcQAdj 
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
      Left            =   5550
      MaxLength       =   20
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   870
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmcQAdj 
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
      Left            =   6555
      Picture         =   "FeedPlge.frx":0FC8
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   870
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcStartNew 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   6975
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   180
      Width           =   105
   End
   Begin VB.PictureBox pbcSpecTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   75
      ScaleHeight     =   45
      ScaleWidth      =   90
      TabIndex        =   13
      Top             =   1005
      Width           =   90
   End
   Begin VB.PictureBox pbcSpecSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   105
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   570
      Width           =   60
   End
   Begin VB.TextBox edcSpecDropdown 
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
      Left            =   825
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmcSpecDropdown 
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
      Left            =   1830
      Picture         =   "FeedPlge.frx":10C2
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   5550
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2925
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "FeedPlge.frx":11BC
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "FeedPlge.frx":1E7A
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
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
      End
   End
   Begin VB.PictureBox pbcSpec 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   405
      Picture         =   "FeedPlge.frx":2184
      ScaleHeight     =   390
      ScaleWidth      =   2100
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   660
      Width           =   2100
   End
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
      Left            =   5490
      TabIndex        =   24
      Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
      Top             =   2205
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.ComboBox cbcSelect 
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
      Left            =   2955
      TabIndex        =   0
      Top             =   120
      Width           =   3795
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6315
      Top             =   3510
   End
   Begin VB.PictureBox plcNum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   4200
      ScaleHeight     =   1140
      ScaleWidth      =   1095
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2310
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcNum 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1050
         Left            =   45
         Picture         =   "FeedPlge.frx":4CD6
         ScaleHeight     =   1050
         ScaleWidth      =   1020
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcNumOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            Top             =   15
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcNumInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   300
            Picture         =   "FeedPlge.frx":5848
            Top             =   255
            Visible         =   0   'False
            Width           =   480
         End
      End
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
      Left            =   135
      Picture         =   "FeedPlge.frx":5B52
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1815
      Visible         =   0   'False
      Width           =   105
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
      Left            =   1665
      Picture         =   "FeedPlge.frx":5E5C
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2235
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
      Left            =   660
      MaxLength       =   20
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   1020
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
      Left            =   1920
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2310
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         Picture         =   "FeedPlge.frx":5F56
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   11
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
            TabIndex        =   12
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
         TabIndex        =   8
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
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
      Height          =   285
      Left            =   6270
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4455
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
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
      Height          =   285
      Left            =   6315
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
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
      Height          =   285
      Left            =   6225
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3045
      Visible         =   0   'False
      Width           =   525
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
      Height          =   60
      Left            =   -15
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4560
      Width           =   75
   End
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   4140
      TabIndex        =   34
      Top             =   4560
      Width           =   945
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
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   30
      Top             =   4200
      Width           =   60
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
      Height          =   120
      Left            =   -15
      ScaleHeight     =   120
      ScaleWidth      =   105
      TabIndex        =   19
      Top             =   1080
      Width           =   105
   End
   Begin VB.VScrollBar vbcPledge 
      Height          =   2910
      LargeChange     =   11
      Left            =   6465
      Max             =   1
      Min             =   1
      TabIndex        =   31
      Top             =   1260
      Value           =   1
      Width           =   240
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3045
      TabIndex        =   33
      Top             =   4560
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1950
      TabIndex        =   32
      Top             =   4560
      Width           =   1050
   End
   Begin VB.PictureBox pbcPledge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   420
      Picture         =   "FeedPlge.frx":8D70
      ScaleHeight     =   2895
      ScaleWidth      =   6030
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1230
      Width           =   6030
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   30
         TabIndex        =   39
         Top             =   570
         Visible         =   0   'False
         Width           =   5985
      End
   End
   Begin VB.PictureBox plcPledge 
      BackColor       =   &H00FFFFFF&
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
      Height          =   3030
      Left            =   360
      ScaleHeight     =   2970
      ScaleWidth      =   6345
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1170
      Width           =   6405
   End
   Begin VB.PictureBox plcSpec 
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   360
      ScaleHeight     =   405
      ScaleWidth      =   2145
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   615
      Width           =   2205
   End
   Begin VB.Label plcScreen 
      Caption         =   "Pledge"
      Height          =   225
      Left            =   60
      TabIndex        =   40
      Top             =   0
      Width           =   825
   End
   Begin VB.Label lacQAdj 
      Appearance      =   0  'Flat
      Caption         =   "Pledge Adjustment (± Time)"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3210
      TabIndex        =   14
      Top             =   855
      Width           =   2370
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   5640
      Picture         =   "FeedPlge.frx":4242A
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   285
      Top             =   4545
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "FeedPlge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of FeedPlge.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: FeedPlge.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Commission input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim tmSpecCtrls(0 To 2)  As FIELDAREA
Dim imLBSpecCtrls As Integer
Dim tmCtrls(0 To 18)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current event name Box
Dim imRowNo As Integer  'Current event row
Dim smSpecSave(1 To 2) As String
Dim smSave() As String  'Feed: 1=Mo; 2= Tu; 3=We; 4=Th; 5=Fr; 6=Sa; 7=Su; 8=Start Time; 9=End Time
                        'Pledge: 10=Mo; 11= Tu; 12=We; 13=Th; 14=Fr;  15=Sa; 16=Su; 17=Start Time; 18=End Time
Dim smShow() As String



Dim tmFpf As FPF        'FPF record image
Dim tmFpfSrchKey As INTKEY0    'FPF key record image
Dim tmFpfSrchKey1 As FPFKEY1    'FPF key record image
Dim tmFpfSrchKey2 As FPFKEY2    'FPF key record image
Dim hmFpf As Integer    'Sale Commission file handle
Dim imFpfRecLen As Integer        'FPF record length
Dim tmFdf As FDF        'FPF record image
Dim tmFdfSrchKey As INTKEY0    'FPF key record image
Dim tmFdfSrchKey1 As FDFKEY1    'FPF key record image
Dim hmFdf As Integer    'Sale Commission file handle
Dim imFdfRecLen As Integer        'FPF record length
Dim imFdfChg As Integer     'Indicates if field changed
Dim imFpfChg As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imFpfSelectIndex As Integer  'Index of selected record (0 if new)
Dim imComboBoxIndex As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imSettingValue As Integer
Dim smNowDate As String
Dim imInNew As Integer
Dim imFirstTimeSelect As Integer
Dim imSpecBoxNo As Integer
Dim imAdjRequired As Integer    'Flag that indicates adjustment required

'Calendar variables
Dim tmCDCtrls(1 To 7) As FIELDAREA  'Field area image
Dim imCalYear As Integer        'Month of displayed calendar
Dim imCalMonth As Integer       'Year of displayed calendar
Dim lmCalStartDate As Long      'Start date of displayed calendar
Dim lmCalEndDate As Long        'End date of displayed calendar
Dim imCalType As Integer        'Calendar type
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer

'Const VEHINDEX = 1
'Const GOALPVEHINDEX = 4
'Const REMNANTPVEHINDEX = 5

Const SPECSTARTDATEINDEX = 1
Const SPECENDDATEINDEX = 2

Const FDMOINDEX = 1
Const FDTUINDEX = 2
Const FDWEINDEX = 3
Const FDTHINDEX = 4
Const FDFRINDEX = 5
Const FDSAINDEX = 6
Const FDSUINDEX = 7
Const FDSTIMEINDEX = 8
Const FDETIMEINDEX = 9
Const PDMOINDEX = 10
Const PDTUINDEX = 11
Const PDWEINDEX = 12
Const PDTHINDEX = 13
Const PDFRINDEX = 14
Const PDSAINDEX = 15
Const PDSUINDEX = 16
Const PDSTIMEINDEX = 17
Const PDETIMEINDEX = 18

'*******************************************************
'*                                                     *
'*      Procedure Name:mStartNew                       *
'*                                                     *
'*             Created:7/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up a New rate card and     *
'*                      initiate RCTerms               *
'*                                                     *
'*******************************************************
Private Function mStartNew() As Integer
    Dim ilRet As Integer
    imInNew = True
    If (cbcSelect.ListCount > 1) Then
        PdModel.Show vbModal
        If (igPdReturn = 0) Or (igPdCodeFpf = 0) Then    'Cancelled
            mStartNew = True
            imInNew = False
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    Else
        igPdCodeFpf = 0
        mStartNew = True
        imInNew = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    Screen.MousePointer = vbHourglass    '
    lacQAdj.Visible = True
    edcQAdj.Visible = True
    cmcQAdj.Visible = True
    ilRet = mReadRec(True)
    pbcPledge.Cls
    mMoveRecToCtrl True
    mInitSpecShow
    mInitShow
    mSetMinMax
    mStartNew = True
    Screen.MousePointer = vbDefault
    mSetCommands
    imInNew = False
    Exit Function

    On Error GoTo 0
    mStartNew = False
    Exit Function
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecTestFields                     *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mSpecTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mSpecTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim ilRes As Integer    'Result of MsgBox

    If (ilCtrlNo = SPECSTARTDATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSpecSave(1), "", "Start Date must be specified", tmSpecCtrls(SPECSTARTDATEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imSpecBoxNo = SPECSTARTDATEINDEX
            End If
            mSpecTestFields = NO
            Exit Function
        End If
        If Not gValidDate(smSpecSave(1)) Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                ilRes = MsgBox("Start Date must be specified correctly", vbOkOnly + vbExclamation, "Incomplete")
                imSpecBoxNo = SPECSTARTDATEINDEX
            End If
            mSpecTestFields = NO
            Exit Function
        End If
    End If
    If smSpecSave(2) <> "" Then
        If (ilCtrlNo = SPECENDDATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedStr(smSpecSave(2), "", "End Date must be specified", tmSpecCtrls(SPECSTARTDATEINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imSpecBoxNo = SPECSTARTDATEINDEX
                End If
                mSpecTestFields = NO
                Exit Function
            End If
            If Not gValidDate(smSpecSave(2)) Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    ilRes = MsgBox("End Date must be specified correctly", vbOkOnly + vbExclamation, "Incomplete")
                    imSpecBoxNo = SPECENDDATEINDEX
                End If
                mSpecTestFields = NO
                Exit Function
            End If
        End If
    End If
    If ilCtrlNo = TESTALLCTRLS Then
        If gDateValue(smNowDate) < gDateValue(smSpecSave(1)) Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                ilRes = MsgBox("Start Date must be in the future of todays date", vbOkOnly + vbExclamation, "Incomplete")
                imSpecBoxNo = SPECSTARTDATEINDEX
            End If
            mSpecTestFields = NO
            Exit Function
        End If
    End If
    mSpecTestFields = YES
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecSetShow                    *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSpecSetShow(ilBoxNo As Integer)
'
'   mSpecSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If ilBoxNo < imLBSpecCtrls Or ilBoxNo > UBound(tmSpecCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case SPECSTARTDATEINDEX
            plcCalendar.Visible = False
            cmcSpecDropdown.Visible = False
            edcSpecDropdown.Visible = False  'Set visibility
            slStr = edcSpecDropdown.Text
            If gValidDate(slStr) Then
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
                If gDateValue(smSpecSave(1)) <> gDateValue(edcSpecDropdown.Text) Then
                    imFpfChg = True
                    tmSpecCtrls(ilBoxNo).iChg = True
                End If
                smSpecSave(1) = edcSpecDropdown.Text
            Else
                Beep
                edcSpecDropdown.Text = smSpecSave(1)
            End If
        Case SPECENDDATEINDEX
            plcCalendar.Visible = False
            cmcSpecDropdown.Visible = False
            edcSpecDropdown.Visible = False  'Set visibility
            slStr = edcSpecDropdown.Text
            If Len(slStr) > 0 Then
                If gValidDate(slStr) Then
                    gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
                    If gDateValue(smSpecSave(1)) <> gDateValue(edcSpecDropdown.Text) Then
                        imFpfChg = True
                        tmSpecCtrls(ilBoxNo).iChg = True
                    End If
                    smSpecSave(2) = edcSpecDropdown.Text
                Else
                    Beep
                    edcSpecDropdown.Text = smSpecSave(2)
                End If
            Else
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
                smSpecSave(2) = ""
            End If
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecSetFocus                   *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSpecSetFocus(ilBoxNo As Integer)
'
'   mSpecSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBSpecCtrls Or ilBoxNo > UBound(tmSpecCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case SPECSTARTDATEINDEX
            If edcSpecDropdown.Enabled Then
                edcSpecDropdown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case SPECENDDATEINDEX
            If edcSpecDropdown.Enabled Then
                edcSpecDropdown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecEnableBox                  *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSpecEnableBox(ilBoxNo As Integer)
'
'   mSpecEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBSpecCtrls Or ilBoxNo > UBound(tmSpecCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case SPECSTARTDATEINDEX
            edcSpecDropdown.Width = tmSpecCtrls(ilBoxNo).fBoxW
            edcSpecDropdown.MaxLength = 10
            gMoveFormCtrl pbcSpec, edcSpecDropdown, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            cmcSpecDropdown.Move edcSpecDropdown.Left + edcSpecDropdown.Width, edcSpecDropdown.Top
            plcCalendar.Move edcSpecDropdown.Left, edcSpecDropdown.Top + edcSpecDropdown.Height
            edcSpecDropdown.Text = Trim$(smSpecSave(1))
            edcSpecDropdown.SelStart = 0
            edcSpecDropdown.SelLength = Len(edcSpecDropdown.Text)
            edcSpecDropdown.Visible = True  'Set visibility
            cmcSpecDropdown.Visible = True
            plcCalendar.Visible = True
            edcSpecDropdown.SetFocus
        Case SPECENDDATEINDEX
            edcSpecDropdown.Width = tmSpecCtrls(ilBoxNo).fBoxW
            edcSpecDropdown.MaxLength = 10
            gMoveFormCtrl pbcSpec, edcSpecDropdown, tmSpecCtrls(ilBoxNo).fBoxX - cmcSpecDropdown, tmSpecCtrls(ilBoxNo).fBoxY
            cmcSpecDropdown.Move edcSpecDropdown.Left + edcSpecDropdown.Width, edcSpecDropdown.Top
            plcCalendar.Move edcSpecDropdown.Left, edcSpecDropdown.Top + edcSpecDropdown.Height
            edcSpecDropdown.Text = Trim$(smSpecSave(2))
            edcSpecDropdown.SelStart = 0
            edcSpecDropdown.SelLength = Len(edcSpecDropdown.Text)
            edcSpecDropdown.Visible = True  'Set visibility
            cmcSpecDropdown.Visible = True
            plcCalendar.Visible = True
            edcSpecDropdown.SetFocus
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitSpecShow                   *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mInitSpecShow()
'
'   mInitSpecShow
'   Where:
'
    Dim slStr As String
    Dim ilBoxNo As Integer
    For ilBoxNo = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        Select Case ilBoxNo 'Branch on box type (control)
            Case SPECSTARTDATEINDEX
                slStr = smSpecSave(1)
                slStr = gFormatDate(slStr)
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case SPECENDDATEINDEX
                slStr = smSpecSave(2)
                slStr = gFormatDate(slStr)
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        End Select
    Next ilBoxNo
End Sub
Private Sub cbcSelect_Change()
    Dim ilRet As Integer
    If imChgMode = False Then
        imChgMode = True
        Screen.MousePointer = vbHourglass  'Wait
        If cbcSelect.Text <> "" Then
            gManLookAhead cbcSelect, imBSMode, imComboBoxIndex
        End If
        imFirstTimeSelect = True
        imFpfSelectIndex = cbcSelect.ListIndex
        mClearCtrlFields
        If imFpfSelectIndex <= 0 Then
            lacQAdj.Visible = False
            edcQAdj.Visible = False
            cmcQAdj.Visible = False
        Else
            lacQAdj.Visible = True
            edcQAdj.Visible = True
            cmcQAdj.Visible = True
        End If
        ilRet = mReadRec(False)
        pbcPledge.Cls
        mMoveRecToCtrl False
        mInitSpecShow
        mInitShow
        mSetMinMax
        pbcSpec_Paint
        mSetCommands
        imChgMode = False
        Screen.MousePointer = vbDefault    'Default
    End If
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcSelect_GotFocus()

    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    plcLen.Visible = False
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
    If imFirstFocus Then
        imFirstFocus = False
    End If
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields 'Make sure all fields cleared
        If pbcSpecSTab.Enabled Then
            pbcSpecSTab.SetFocus
        Else
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    imComboBoxIndex = imFpfSelectIndex
    gCtrlGotFocus cbcSelect
    Exit Sub
End Sub
Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcSpecDropdown.SelStart = 0
    edcSpecDropdown.SelLength = Len(edcSpecDropdown.Text)
    edcSpecDropdown.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcSpecDropdown.SelStart = 0
    edcSpecDropdown.SelLength = Len(edcSpecDropdown.Text)
    edcSpecDropdown.SetFocus
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If mSaveRecChg(True) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    mMakeAdj
    plcLen.Visible = False
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
'        Case VEHINDEX
'            lbcVehicle.Visible = Not lbcVehicle.Visible
'        Case SDATEINDEX
'            plcCalendar.Visible = Not plcCalendar.Visible
'        Case EDATEINDEX
'            plcCalendar.Visible = Not plcCalendar.Visible
'        Case GOALPVEHINDEX
'            plcNum.Visible = Not plcNum.Visible
'        Case REMNANTPVEHINDEX
'            plcNum.Visible = Not plcNum.Visible
        Case FDSTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case FDETIMEINDEX
            plcTme.Visible = Not plcTme.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcQAdj_Click()
    plcLen.Visible = Not plcLen.Visible
    edcQAdj.SelStart = 0
    edcQAdj.SelLength = Len(edcDropDown.Text)
    edcQAdj.SetFocus
End Sub

Private Sub cmcQAdj_GotFocus()
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
    plcLen.Move edcQAdj.Left, edcQAdj.Top + edcQAdj.Height
End Sub

Private Sub cmcSave_Click()
    Dim ilCode As Integer
    Dim ilLoop As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    ReDim tgPlgeDel(1 To 1) As FDFREC
    ilCode = tmFpf.iCode
    cbcSelect.Clear
    mPopulate
    pbcPledge.Cls
    pbcPledge_Paint
    imSpecBoxNo = -1
    imBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
    imFdfChg = False
    imFpfChg = False
    mSetCommands
    For ilLoop = 0 To cbcSelect.ListCount - 1 Step 1
        If ilCode = cbcSelect.ItemData(ilLoop) Then
            cbcSelect.ListIndex = ilLoop
            Exit For
        End If
    Next ilLoop

End Sub

Private Sub cmcSave_GotFocus()
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    mMakeAdj
    plcLen.Visible = False
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
End Sub

Private Sub cmcSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcSpecDropDown_Click()
    Select Case imSpecBoxNo
        Case SPECSTARTDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case SPECENDDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
    End Select
    edcSpecDropdown.SelStart = 0
    edcSpecDropdown.SelLength = Len(edcSpecDropdown.Text)
    edcSpecDropdown.SetFocus
End Sub

Private Sub cmcSpecDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_Change()
    Select Case imBoxNo
'        Case VEHINDEX
'            imLbcArrowSetting = True
'            gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
'        Case SDATEINDEX
'            slStr = edcDropDown.Text
'            If Not gValidDate(slStr) Then
'                lacDate.Visible = False
'                Exit Sub
'            End If
'            lacDate.Visible = True
'            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
'            pbcCalendar_Paint   'mBoxCalDate called within paint
'        Case EDATEINDEX
'            slStr = edcDropDown.Text
'            If Not gValidDate(slStr) Then
'                lacDate.Visible = False
'                Exit Sub
'            End If
'            lacDate.Visible = True
'            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
'            pbcCalendar_Paint   'mBoxCalDate called within paint
'        Case GOALPVEHINDEX
'        Case REMNANTPVEHINDEX
    End Select
    imLbcArrowSetting = False
End Sub
Private Sub edcDropDown_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
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
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilKeyAscii As Integer
    ilKeyAscii = KeyAscii
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imBoxNo
'        Case VEHINDEX
'        Case SDATEINDEX
'            'Filter characters (allow only BackSpace, numbers 0 thru 9
'            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
'                Beep
'                KeyAscii = 0
'                Exit Sub
'            End If
'        Case EDATEINDEX
'            'Disallow TFN for alternate
'            If (Len(edcDropDown.Text) = edcDropDown.SelLength) Then
'                If (KeyAscii = Asc("T")) Or (KeyAscii = Asc("t")) Then
'                    edcDropDown.Text = "TFN"
'                    edcDropDown.SelStart = 0
'                    edcDropDown.SelLength = 3
'                    KeyAscii = 0
'                    Exit Sub
'                End If
'            End If
'            'Filter characters (allow only BackSpace, numbers 0 thru 9
'            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
'                Beep
'                KeyAscii = 0
'                Exit Sub
'            End If
'        Case GOALPVEHINDEX
'            If Not mDropDownKeyPress(ilKeyAscii, False) Then
'                KeyAscii = 0
'                Exit Sub
'            End If
'        Case REMNANTPVEHINDEX
'            If Not mDropDownKeyPress(ilKeyAscii, False) Then
'                KeyAscii = 0
'                Exit Sub
'            End If
        Case FDSTIMEINDEX, FDETIMEINDEX
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
        Case PDSTIMEINDEX, PDETIMEINDEX
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
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KeyUp) Or (KeyCode = KeyDown) Then
        Select Case imBoxNo
'            Case VEHINDEX
'                If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
'                    gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
'                End If
'            Case SDATEINDEX
'                If (Shift And vbAltMask) > 0 Then
'                    plcCalendar.Visible = Not plcCalendar.Visible
'                Else
'                    slDate = edcDropDown.Text
'                    If gValidDate(slDate) Then
'                        If KeyCode = KEYUP Then 'Up arrow
'                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
'                        Else
'                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
'                        End If
'                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
'                        edcDropDown.Text = slDate
'                    End If
'                End If
'            Case EDATEINDEX
'                If (Shift And vbAltMask) > 0 Then
'                    plcCalendar.Visible = Not plcCalendar.Visible
'                Else
'                    slDate = edcDropDown.Text
'                    If gValidDate(slDate) Then
'                        If KeyCode = KEYUP Then 'Up arrow
'                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
'                        Else
'                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
'                        End If
'                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
'                        edcDropDown.Text = slDate
'                    End If
'                End If
'            Case GOALPVEHINDEX
'            Case REMNANTPVEHINDEX
            Case FDSTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case FDETIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case PDSTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case PDETIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNo
'            Case VEHINDEX
'            Case SDATEINDEX
'                If (Shift And vbAltMask) > 0 Then
'                Else
'                    slDate = edcDropDown.Text
'                    If gValidDate(slDate) Then
'                        If KeyCode = KEYLEFT Then 'Up arrow
'                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
'                        Else
'                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
'                        End If
'                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
'                        edcDropDown.Text = slDate
'                    End If
'                End If
'                edcDropDown.SelStart = 0
'                edcDropDown.SelLength = Len(edcDropDown.Text)
'            Case EDATEINDEX
'                If (Shift And vbAltMask) > 0 Then
'                Else
'                    slDate = edcDropDown.Text
'                    If gValidDate(slDate) Then
'                        If KeyCode = KEYLEFT Then 'Up arrow
'                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
'                        Else
'                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
'                        End If
'                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
'                        edcDropDown.Text = slDate
'                    End If
'                End If
'                edcDropDown.SelStart = 0
'                edcDropDown.SelLength = Len(edcDropDown.Text)
'            Case GOALPVEHINDEX
'            Case REMNANTPVEHINDEX
        End Select
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcQAdj_Change()
    Dim slStr As String
    Dim slNeg As String

    slStr = edcQAdj.Text
    If InStr(1, slStr, "-", vbTextCompare) = 1 Then
        slNeg = "-"
        slStr = Mid$(slStr, 2)
    Else
        slNeg = ""
    End If
    If (gValidLength(slStr)) And (Trim$(slStr) <> "") Then
        imAdjRequired = True
    Else
        imAdjRequired = False
    End If
End Sub

Private Sub edcQAdj_GotFocus()
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    plcLen.Move edcQAdj.Left, edcQAdj.Top + edcQAdj.Height
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcQAdj_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcQAdj_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcSpecDropdown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        ilFound = False
        If (KeyAscii <> KEYNEG) Or (edcQAdj.Text <> "") Then
            For ilLoop = LBound(igLegalLength) To UBound(igLegalLength) Step 1
                If KeyAscii = igLegalLength(ilLoop) Then
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
    End If
    gLengthOutLine KeyAscii, imcLenOutline
End Sub

Private Sub edcQAdj_KeyUp(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KeyDown) Then
        imDirProcess = KeyCode 'mDirection 0
        pbcSpecTab.SetFocus
        Exit Sub
    End If
    If (KeyCode = KeyUp) Or (KeyCode = KeyDown) Then
        If (Shift And vbAltMask) > 0 Then
            plcLen.Visible = Not plcLen.Visible
        End If
        edcSpecDropdown.SelStart = 0
        edcSpecDropdown.SelLength = Len(edcSpecDropdown.Text)
    End If
End Sub

Private Sub edcSpecDropDown_Change()
    Dim slStr As String
    Select Case imSpecBoxNo
        Case SPECSTARTDATEINDEX
            slStr = edcSpecDropdown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case SPECENDDATEINDEX
            slStr = edcSpecDropdown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
    End Select
End Sub

Private Sub edcSpecDropDown_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
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
    Dim ilKeyAscii As Integer
    ilKeyAscii = KeyAscii
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imSpecBoxNo
        Case SPECSTARTDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case SPECENDDATEINDEX
            'Disallow TFN for alternate
            If (Len(edcSpecDropdown.Text) = edcSpecDropdown.SelLength) Then
                If (KeyAscii = Asc("T")) Or (KeyAscii = Asc("t")) Then
                    edcSpecDropdown.Text = "TFN"
                    edcSpecDropdown.SelStart = 0
                    edcSpecDropdown.SelLength = 3
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub

Private Sub edcSpecDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KeyUp) Or (KeyCode = KeyDown) Then
        Select Case imBoxNo
            Case SPECSTARTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcSpecDropdown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KeyUp Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpecDropdown.Text = slDate
                    End If
                End If
            Case SPECENDDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcSpecDropdown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KeyUp Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpecDropdown.Text = slDate
                    End If
                End If
        End Select
        edcSpecDropdown.SelStart = 0
        edcSpecDropdown.SelLength = Len(edcSpecDropdown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNo
            Case SPECSTARTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcSpecDropdown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpecDropdown.Text = slDate
                    End If
                End If
                edcSpecDropdown.SelStart = 0
                edcSpecDropdown.SelLength = Len(edcSpecDropdown.Text)
            Case SPECENDDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcSpecDropdown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpecDropdown.Text = slDate
                    End If
                End If
                edcSpecDropdown.SelStart = 0
                edcSpecDropdown.SelLength = Len(edcSpecDropdown.Text)
        End Select
    End If
End Sub

Private Sub Form_Activate()
    If imInNew Then
        Exit Sub
    End If
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True  'To get Alt J and Alt L keys
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(FEEDJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcPledge.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcPledge.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    Me.ZOrder 0 'Send to front
    FeedPlge.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcNum.Visible = False
        If (cbcSelect.Enabled) And (imBoxNo > 0) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            cbcSelect.Enabled = True
        End If
    End If

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    'btrStopAppl
    'End
    igJobShowing(FEEDJOB) = False
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub imcHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub imcTrash_Click()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilUpperBound As Integer
    Dim ilRowNo As Integer
    Dim ilFpf As Integer
    If (imRowNo < vbcPledge.Value) Or (imRowNo > vbcPledge.Value + vbcPledge.LargeChange) Then
        Exit Sub
    End If
    ilRowNo = imRowNo
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
    ilUpperBound = UBound(smSave, 2)
    ilFpf = ilRowNo
    If ilFpf = ilUpperBound Then
        mInitNew ilFpf
    Else
        If ilFpf > 0 Then
            If tgPlgeRec(ilFpf).iStatus = 1 Then
                tgPlgeDel(UBound(tgPlgeDel)).tFdf = tgPlgeRec(ilFpf).tFdf
                tgPlgeDel(UBound(tgPlgeDel)).iStatus = tgPlgeRec(ilFpf).iStatus
                tgPlgeDel(UBound(tgPlgeDel)).lRecPos = tgPlgeRec(ilFpf).lRecPos
                ReDim Preserve tgPlgeDel(1 To UBound(tgPlgeDel) + 1) As FDFREC
            End If
            ilFpf = ilRowNo
            'Remove record from tgRjf1Rec- Leave tgPjf2Rec
            For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
                tgPlgeRec(ilLoop) = tgPlgeRec(ilLoop + 1)
            Next ilLoop
            ReDim Preserve tgPlgeRec(1 To UBound(tgPlgeRec) - 1) As FDFREC
        End If
        For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
            For ilIndex = 1 To UBound(smSave, 1) Step 1
                smSave(ilIndex, ilLoop) = smSave(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(smShow, 1) Step 1
                smShow(ilIndex, ilLoop) = smShow(ilIndex, ilLoop + 1)
            Next ilIndex
        Next ilLoop
        ilUpperBound = UBound(smSave, 2)
        ReDim Preserve smShow(1 To 18, 1 To ilUpperBound - 1) As String 'Values shown in program area
        ReDim Preserve smSave(1 To 18, 1 To ilUpperBound - 1) As String    'Values saved (program name) in program area
        imFdfChg = True
    End If
    mSetCommands
    lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    imSettingValue = True
    vbcPledge.Min = LBound(smShow, 2)
    imSettingValue = True
    If UBound(smShow, 2) - 1 <= vbcPledge.LargeChange + 1 Then ' + 1 Then
        vbcPledge.Max = LBound(smShow, 2)
    Else
        vbcPledge.Max = UBound(smShow, 2) - vbcPledge.LargeChange
    End If
    imSettingValue = True
    vbcPledge.Value = vbcPledge.Min
    imSettingValue = True
    pbcPledge.Cls
    pbcPledge_Paint
End Sub
Private Sub imcTrash_DragDrop(Source As Control, X As Single, Y As Single)
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        lacFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
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

    slStr = edcSpecDropdown.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(Str$(Day(llDate)))
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
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear each control on the      *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'
    Dim ilLoop As Integer

    imFdfChg = False
    imFpfChg = False
    pbcSpec.Cls
    pbcPledge.Cls
    ReDim tgPlgeRec(1 To 1) As FDFREC
    tgPlgeRec(1).iStatus = -1
    tgPlgeRec(1).lRecPos = 0
    tgPlgeRec(1).iDateChg = False
    ReDim tgPlgeDel(1 To 1) As FDFREC
    tgPlgeDel(1).iStatus = -1
    tgPlgeDel(1).lRecPos = 0
    For ilLoop = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        tmSpecCtrls(ilLoop).sShow = ""
    Next ilLoop
    smSpecSave(1) = ""
    smSpecSave(2) = ""
    ReDim smShow(1 To 18, 1 To 1) As String 'Values shown in program area
    ReDim smSave(1 To 18, 1 To 1) As String 'Values saved (program name) in program area
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, 1) = ""
    Next ilLoop
    vbcPledge.Min = LBound(smShow, 2)
    imSettingValue = True
    vbcPledge.Max = LBound(smShow, 2)
    imSettingValue = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If
    If (imRowNo < vbcPledge.Value) Or (imRowNo >= vbcPledge.Value + vbcPledge.LargeChange + 1) Then
        mSetShow ilBoxNo
        Exit Sub
    End If
    lacFrame.Move 0, tmCtrls(FDMOINDEX).fBoxY + (imRowNo - vbcPledge.Value) * (fgBoxGridH + 15) - 30
    lacFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcPledge.Top + tmCtrls(FDMOINDEX).fBoxY + (imRowNo - vbcPledge.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True

    Select Case ilBoxNo 'Branch on box type (control)
        Case FDMOINDEX To FDSUINDEX
            gMoveTableCtrl pbcPledge, ckcAirDay, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcPledge.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(1 + ilBoxNo - FDMOINDEX, imRowNo))
            If slStr = "" Or slStr = "Y" Then
                ckcAirDay.Value = vbChecked
            Else
                ckcAirDay.Value = vbUnchecked
            End If
            ckcAirDay.Visible = True
            ckcAirDay.SetFocus
        Case FDSTIMEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 11
            gMoveTableCtrl pbcPledge, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcPledge.Value) * (fgBoxGridH + 15)
            edcDropDown.Left = edcDropDown.Left - cmcDropDown.Width
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            slStr = Trim$(smSave(8, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            If imRowNo > UBound(smSave, 2) - 1 Then
                plcTme.Visible = True
            End If
            edcDropDown.SetFocus
        Case FDETIMEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 11
            gMoveTableCtrl pbcPledge, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcPledge.Value) * (fgBoxGridH + 15)
            edcDropDown.Left = edcDropDown.Left - cmcDropDown.Width
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            slStr = Trim$(smSave(9, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            If imRowNo > UBound(smSave, 2) - 1 Then
                plcTme.Visible = True
            End If
            edcDropDown.SetFocus
        Case PDMOINDEX To PDSUINDEX
            gMoveTableCtrl pbcPledge, ckcAirDay, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcPledge.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(10 + ilBoxNo - PDMOINDEX, imRowNo))
            If slStr = "" Or slStr = "Y" Then
                ckcAirDay.Value = vbChecked
            Else
                ckcAirDay.Value = vbUnchecked
            End If
            ckcAirDay.Visible = True
            ckcAirDay.SetFocus
        Case PDSTIMEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 11
            gMoveTableCtrl pbcPledge, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcPledge.Value) * (fgBoxGridH + 15)
            edcDropDown.Left = edcDropDown.Left - cmcDropDown.Width
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            slStr = Trim$(smSave(17, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            If imRowNo > UBound(smSave, 2) - 1 Then
                plcTme.Visible = True
            End If
            edcDropDown.SetFocus
        Case PDETIMEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 11
            gMoveTableCtrl pbcPledge, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcPledge.Value) * (fgBoxGridH + 15)
            edcDropDown.Left = edcDropDown.Left - cmcDropDown.Width
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            slStr = Trim$(smSave(18, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            If imRowNo > UBound(smSave, 2) - 1 Then
                plcTme.Visible = True
            End If
            edcDropDown.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
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
    Dim ilLoop As Integer
    Dim slDate As String
    imTerminate = False
    imFirstActivate = True
    imInNew = False
    imFirstTimeSelect = True
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165

    Screen.MousePointer = vbHourglass
    imLBSpecCtrls = 1
    imLBCtrls = 1
    igJobShowing(FEEDJOB) = True
    imFirstActivate = True
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    FeedPlge.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    mInitBox
    smNowDate = Format$(Now, "m/d/yy")
    gCenterForm FeedPlge
    'FeedPlge.Show
    Screen.MousePointer = vbHourglass
    ReDim tgPlgeRec(1 To 1) As FDFREC
    tgPlgeRec(1).iStatus = -1
    tgPlgeRec(1).lRecPos = 0
    ReDim tgPlgeDel(1 To 1) As FDFREC
    tgPlgeDel(1).iStatus = -1
    tgPlgeDel(1).lRecPos = 0
    ReDim smShow(1 To 18, 1 To 1) As String 'Values shown in program area
    ReDim smSave(1 To 18, 1 To 1) As String 'Values saved (program name) in program area
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, 1) = ""
    Next ilLoop
'    mInitDDE
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    imFirstFocus = True
    imDoubleClickName = False
    imLbcMouseDown = False
    imCalType = 0               'Standard type
    imBoxNo = -1                'Initialize current Box to N/A
    imRowNo = -1
    imSpecBoxNo = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imChgMode = False
    imBSMode = False
    imBypassFocus = False
    imBypassSetting = False
    imFpfSelectIndex = -1
    imFdfChg = False
    imFpfChg = False
    imAdjRequired = False
    hmFpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmFpf, "", sgDBPath & "FPF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: FPF.Btr)", FeedPlge
    On Error GoTo 0
    imFpfRecLen = Len(tmFpf)
    hmFdf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmFdf, "", sgDBPath & "FDF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: FDF.Btr)", FeedPlge
    On Error GoTo 0
    imFdfRecLen = Len(tmFdf)
    cbcSelect.Clear 'Force list box to be populated
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    slDate = Format$(gNow(), "m/d/yy")
    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
    If cbcSelect.ListCount > 0 Then
        cbcSelect.ListIndex = 0
    End If
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
'*             Created:5/17/93       By:D. LeVine      *
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
    flTextHeight = pbcPledge.TextHeight("1") - 35
    'Position panel and picture areas with panel
    'plcSelect.Move 3555, 120
    cbcSelect.Move 2955, 120

    plcSpec.Move 360, 615, pbcSpec.Width + fgPanelAdj, pbcSpec.Height + fgPanelAdj
    pbcSpec.Move plcSpec.Left + fgBevelX, plcSpec.Top + fgBevelY

    plcPledge.Move 360, 1170, pbcPledge.Width + fgPanelAdj + vbcPledge.Width, pbcPledge.Height + fgPanelAdj
    pbcPledge.Move plcPledge.Left + fgBevelX, plcPledge.Top + fgBevelY
    vbcPledge.Move pbcPledge.Left + pbcPledge.Width - 15, pbcPledge.Top
    'Start Date
    gSetCtrl tmSpecCtrls(SPECSTARTDATEINDEX), 30, 30, 1020, fgBoxStH
    'End Date
    gSetCtrl tmSpecCtrls(SPECENDDATEINDEX), 1065, tmSpecCtrls(SPECSTARTDATEINDEX).fBoxY, 1020, fgBoxStH

    'Monday Spots
    gSetCtrl tmCtrls(FDMOINDEX), 30, 555, 225, fgBoxGridH
    'Tuesday Spots
    gSetCtrl tmCtrls(FDTUINDEX), 270, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Wednesday Spots
    gSetCtrl tmCtrls(FDWEINDEX), 510, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Thursday Spots
    gSetCtrl tmCtrls(FDTHINDEX), 750, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Friday Spots
    gSetCtrl tmCtrls(FDFRINDEX), 990, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Saturday Spots
    gSetCtrl tmCtrls(FDSAINDEX), 1230, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Sunday Spots
    gSetCtrl tmCtrls(FDSUINDEX), 1470, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Start Time
    gSetCtrl tmCtrls(FDSTIMEINDEX), 1710, tmCtrls(FDMOINDEX).fBoxY, 645, fgBoxGridH
    'End Time
    gSetCtrl tmCtrls(FDETIMEINDEX), 2370, tmCtrls(FDMOINDEX).fBoxY, 645, fgBoxGridH

    'Monday Spots
    gSetCtrl tmCtrls(PDMOINDEX), 3030, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Tuesday Spots
    gSetCtrl tmCtrls(PDTUINDEX), 3270, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Wednesday Spots
    gSetCtrl tmCtrls(PDWEINDEX), 3510, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Thursday Spots
    gSetCtrl tmCtrls(PDTHINDEX), 3750, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Friday Spots
    gSetCtrl tmCtrls(PDFRINDEX), 3990, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Saturday Spots
    gSetCtrl tmCtrls(PDSAINDEX), 4230, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Sunday Spots
    gSetCtrl tmCtrls(PDSUINDEX), 4470, tmCtrls(FDMOINDEX).fBoxY, 225, fgBoxGridH
    'Start Time
    gSetCtrl tmCtrls(PDSTIMEINDEX), 4710, tmCtrls(FDMOINDEX).fBoxY, 645, fgBoxGridH
    'End Time
    gSetCtrl tmCtrls(PDETIMEINDEX), 5370, tmCtrls(FDMOINDEX).fBoxY, 645, fgBoxGridH

    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNew                        *
'*                                                     *
'*             Created:9/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize values              *
'*                                                     *
'*******************************************************
Private Sub mInitNew(ilRowNo As Integer)
    Dim ilLoop As Integer

    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, ilRowNo) = ""
    Next ilLoop
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, ilRowNo) = ""
    Next ilLoop
    tgPlgeRec(ilRowNo).iStatus = 0
    tgPlgeRec(ilRowNo).lRecPos = 0
    tgPlgeRec(ilRowNo).iDateChg = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitShow                       *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mInitShow()
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    For ilRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        For ilBoxNo = imLBCtrls To UBound(tmCtrls) Step 1
            Select Case ilBoxNo
                Case FDMOINDEX To FDSUINDEX
                    slStr = smSave(1 + ilBoxNo - FDMOINDEX, ilRowNo)
                    gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case FDSTIMEINDEX
                    slStr = Trim$(smSave(8, ilRowNo))
                    gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case FDETIMEINDEX
                    slStr = Trim$(smSave(9, ilRowNo))
                    gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case PDMOINDEX To PDSUINDEX
                    slStr = smSave(10 + ilBoxNo - PDMOINDEX, ilRowNo)
                    gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case PDSTIMEINDEX
                    slStr = Trim$(smSave(17, ilRowNo))
                    gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case PDETIMEINDEX
                    slStr = Trim$(smSave(18, ilRowNo))
                    gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
            End Select
        Next ilBoxNo
    Next ilRowNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
'
'   mMoveCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim ilRowNo As Integer

    tmFpf.iFnfCode = igPledgeFnfCode
    tmFpf.iVefCode = igPledgeVefCode
    gPackDate smSpecSave(1), tmFpf.iEffStartDate(0), tmFpf.iEffStartDate(1)
    If (Trim$(smSpecSave(2)) = "") Or (StrComp(smSpecSave(2), "TFN", vbTextCompare) = 0) Then
        gPackDate "12/31/2068", tmFpf.iEffEndDate(0), tmFpf.iEffEndDate(1)
    Else
        gPackDate smSpecSave(2), tmFpf.iEffEndDate(0), tmFpf.iEffEndDate(1)
    End If
    tmFpf.iUrfCode = tgUrf(0).iCode

    For ilRowNo = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
        'Feed Days
        For ilLoop = 0 To 6 Step 1
            tgPlgeRec(ilRowNo).tFdf.sFeedDays(ilLoop) = smSave(ilLoop + 1, ilRowNo)
        Next ilLoop
        'Feed Times
        gPackTime smSave(8, ilRowNo), tgPlgeRec(ilRowNo).tFdf.iFeedStartTime(0), tgPlgeRec(ilRowNo).tFdf.iFeedStartTime(1)
        gPackTime smSave(9, ilRowNo), tgPlgeRec(ilRowNo).tFdf.iFeedEndTime(0), tgPlgeRec(ilRowNo).tFdf.iFeedEndTime(1)
        'Pledge Days
        For ilLoop = 0 To 6 Step 1
            tgPlgeRec(ilRowNo).tFdf.sPledgeDays(ilLoop) = smSave(ilLoop + 10, ilRowNo)
        Next ilLoop
        'Pledge Times
        gPackTime smSave(17, ilRowNo), tgPlgeRec(ilRowNo).tFdf.iPledgeStartTime(0), tgPlgeRec(ilRowNo).tFdf.iPledgeStartTime(1)
        gPackTime smSave(18, ilRowNo), tgPlgeRec(ilRowNo).tFdf.iPledgeEndTime(0), tgPlgeRec(ilRowNo).tFdf.iPledgeEndTime(1)
    Next ilRowNo
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl(ilFromModel As Integer)
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim ilRowNo As Integer

    If Not ilFromModel Then
        If (imFpfSelectIndex <= 0) Then
            Exit Sub
        End If
        gUnpackDate tmFpf.iEffStartDate(0), tmFpf.iEffStartDate(1), smSpecSave(1)
        gUnpackDate tmFpf.iEffEndDate(0), tmFpf.iEffEndDate(1), smSpecSave(2)
        If Trim$(smSpecSave(2)) <> "" Then
            If gDateValue(smSpecSave(2)) = gDateValue("12/31/2068") Then
                smSpecSave(2) = "TFN"
            End If
        End If
    Else
        smSpecSave(1) = ""
        smSpecSave(2) = ""
    End If
    For ilLoop = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        tmSpecCtrls(ilLoop).iChg = False
    Next ilLoop

    ilUpper = UBound(tgPlgeRec)
    ReDim smShow(1 To 18, 1 To ilUpper) As String 'Values shown in program area
    ReDim smSave(1 To 18, 1 To ilUpper) As String 'Values saved (program name) in program area
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, ilUpper) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, ilUpper) = ""
    Next ilLoop
    'Init value in the case that no records are associated with the salesperson
    If ilUpper = LBound(tgPlgeRec) Then
        ilRowNo = imRowNo
        imRowNo = 1
        mInitNew imRowNo
        imRowNo = ilRowNo
    End If
    For ilRowNo = LBound(tgPlgeRec) To UBound(tgPlgeRec) - 1 Step 1
        'Feed Days
        For ilLoop = 0 To 6 Step 1
            smSave(ilLoop + 1, ilRowNo) = tgPlgeRec(ilRowNo).tFdf.sFeedDays(ilLoop)
        Next ilLoop
        'Feed Times
        gUnpackTime tgPlgeRec(ilRowNo).tFdf.iFeedStartTime(0), tgPlgeRec(ilRowNo).tFdf.iFeedStartTime(1), "A", "1", smSave(8, ilRowNo)
        gUnpackTime tgPlgeRec(ilRowNo).tFdf.iFeedEndTime(0), tgPlgeRec(ilRowNo).tFdf.iFeedEndTime(1), "A", "1", smSave(9, ilRowNo)
        'Pledge Days
        For ilLoop = 0 To 6 Step 1
            smSave(ilLoop + 10, ilRowNo) = tgPlgeRec(ilRowNo).tFdf.sPledgeDays(ilLoop)
        Next ilLoop
        'Pledge Times
        gUnpackTime tgPlgeRec(ilRowNo).tFdf.iPledgeStartTime(0), tgPlgeRec(ilRowNo).tFdf.iPledgeStartTime(1), "A", "1", smSave(17, ilRowNo)
        gUnpackTime tgPlgeRec(ilRowNo).tFdf.iPledgeEndTime(0), tgPlgeRec(ilRowNo).tFdf.iPledgeEndTime(1), "A", "1", smSave(18, ilRowNo)
    Next ilRowNo
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Salesperson list box  *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer
    Dim slStartDate As String
    Dim slEndDate As String

    'Populate with each unique effective date


    tmFpfSrchKey2.iFnfCode = igPledgeFnfCode
    tmFpfSrchKey2.iVefCode = igPledgeVefCode
    tmFpfSrchKey2.iEffStartDate(0) = 0
    tmFpfSrchKey2.iEffStartDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmFpf, tmFpf, imFpfRecLen, tmFpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmFpf.iFnfCode = igPledgeFnfCode) And (tmFpf.iVefCode = igPledgeVefCode)
        gUnpackDate tmFpf.iEffStartDate(0), tmFpf.iEffStartDate(1), slStartDate
        gUnpackDate tmFpf.iEffEndDate(0), tmFpf.iEffEndDate(1), slEndDate
        If Trim$(slEndDate) <> "" Then
            If gDateValue(slEndDate) = gDateValue("12/31/2068") Then
                slEndDate = "TFN"
            End If
        End If
        'set to zero to have date in descending order
        cbcSelect.AddItem slStartDate & "-" & slEndDate, 0
        cbcSelect.ItemData(cbcSelect.NewIndex) = Trim$(Str$(tmFpf.iCode))
        ilRet = btrGetNext(hmFpf, tmFpf, imFpfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    cbcSelect.AddItem "[New]", 0
    cbcSelect.ItemData(cbcSelect.NewIndex) = 0
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilFromModel As Integer) As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilFpfCode As Integer
    Dim ilUpper As Integer
    Dim llNoRec As Long
    Dim ilOffset As Integer
    Dim slStr As String
    Dim slStartTime As String
    Dim llStartTime As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    ReDim tgPlgeRec(1 To 1) As FDFREC
    tgPlgeRec(1).iStatus = -1
    tgPlgeRec(1).lRecPos = 0
    tgPlgeRec(1).iDateChg = False
    ReDim tgPlgeDel(1 To 1) As FDFREC
    tgPlgeDel(1).iStatus = -1
    tgPlgeDel(1).lRecPos = 0
    ilUpper = 1
    If (imFpfSelectIndex < 0) And (Not ilFromModel) Then
        mReadRec = False
        Exit Function
    End If
    btrExtClear hmFdf   'Clear any previous extend operation
    ilExtLen = Len(tgPlgeRec(1).tFdf)  'Extract operation record size
    If Not ilFromModel Then
        ilFpfCode = cbcSelect.ItemData(imFpfSelectIndex)
    Else
        ilFpfCode = igPdCodeFpf
    End If
    'Build program images from newest
    tmFpfSrchKey.iCode = ilFpfCode
    ilRet = btrGetEqual(hmFpf, tmFpf, imFpfRecLen, tmFpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mReadRec = False
        Exit Function
    End If
    tmFdfSrchKey1.iFpfCode = ilFpfCode
    ilRet = btrGetEqual(hmFdf, tmFdf, imFdfRecLen, tmFdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
        ilRet = BTRV_ERR_END_OF_FILE
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mReadRec = False
            Exit Function
        End If
    End If
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmFdf, llNoRec, -1, "UC", "Fdf", "") '"EG") 'Set extract limits (all records)
        ilOffset = gFieldOffset("Fdf", "FDFFPFCODE") 'GetOffSetForInt(tmFdf, tmFdf.iSlfCode)
        tlIntTypeBuff.iType = ilFpfCode
        ilRet = btrExtAddLogicConst(hmFdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Fdf.Btr", FeedPlge
        On Error GoTo 0
        ilRet = btrExtAddField(hmFdf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mReadRec (btrExtAddField):" & "Fdf.Btr", FeedPlge
        On Error GoTo 0
        ilRet = btrExtGetNext(hmFdf, tmFdf, ilExtLen, tgPlgeRec(ilUpper).lRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrExtGetNextExt):" & "Fdf.Btr", FeedPlge
            On Error GoTo 0
            ilExtLen = Len(tmFdf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmFdf, tmFdf, ilExtLen, tgPlgeRec(ilUpper).lRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slStr = ""

                gUnpackTimeLong tmFdf.iFeedStartTime(0), tmFdf.iFeedStartTime(0), False, llStartTime
                slStartTime = Trim$(Str$(llStartTime))
                Do While Len(slStartTime) < 6
                    slStartTime = "0" & slStartTime
                Loop
                tgPlgeRec(ilUpper).sKey = slStr & slStartTime
                tgPlgeRec(ilUpper).tFdf = tmFdf
                If ilFromModel Then
                    tgPlgeRec(ilUpper).tFdf.iCode = 0
                    tgPlgeRec(ilUpper).iStatus = 0
                    tgPlgeRec(ilUpper).lRecPos = 0
                Else
                    tgPlgeRec(ilUpper).iStatus = 1
                End If
                tgPlgeRec(ilUpper).iDateChg = False
                ilUpper = ilUpper + 1
                ReDim Preserve tgPlgeRec(1 To ilUpper) As FDFREC
                tgPlgeRec(ilUpper).iStatus = -1
                tgPlgeRec(ilUpper).lRecPos = 0
                ilRet = btrExtGetNext(hmFdf, tmFdf, ilExtLen, tgPlgeRec(ilUpper).lRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmFdf, tgPlgeRec(ilUpper).tFdf, ilExtLen, tgPlgeRec(ilUpper).lRecPos)
                Loop
            Loop
        End If
    End If
    If ilUpper > 1 Then
        ArraySortTyp fnAV(tgPlgeRec(), 1), UBound(tgPlgeRec) - 1, 0, LenB(tgPlgeRec(1)), 0, LenB(tgPlgeRec(1).sKey), 0
    End If
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:6/29/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilLoop As Integer   'For loop control
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilRowNo As Integer
    Dim slMsg As String
    Dim ilFpfCode As Integer
    Dim ilTFpfCode As Integer
    Dim ilFdf As Integer
    Dim tlFdf As FDF
    Dim tlFdf1 As MOVEREC
    Dim tlFdf2 As MOVEREC
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    If mSpecTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        mSaveRec = False
        Exit Function
    End If
    For ilRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        If mTestSaveFields(ilRowNo) = NO Then
            mSaveRec = False
            imRowNo = ilRowNo
            Exit Function
        End If
    Next ilRowNo
    If mOverlappingTimes() Then
        mSaveRec = False
        Exit Function
    End If
    ilTFpfCode = mOverlappingDates()
    If ilTFpfCode = -1 Then
        mSaveRec = False
        Exit Function
    End If
    ilRet = btrBeginTrans(hmFdf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 1", vbOkOnly + vbExclamation, "Pledge")
        Exit Function
    End If
    If ilTFpfCode > 0 Then
        tmFpfSrchKey.iCode = ilTFpfCode
        ilRet = btrGetEqual(hmFpf, tmFpf, imFpfRecLen, tmFpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        slDate = gDecOneDay(smSpecSave(1))
        gPackDate slDate, tmFpf.iEffEndDate(0), tmFpf.iEffEndDate(1)
        ilRet = btrUpdate(hmFpf, tmFpf, imFpfRecLen)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            ilCRet = btrAbortTrans(hmFdf)
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at FPF-0", vbOkOnly + vbExclamation, "Pledge")
            Exit Function
        End If
    End If
    If imFpfSelectIndex > 0 Then
        tmFpfSrchKey.iCode = cbcSelect.ItemData(imFpfSelectIndex)
        ilRet = btrGetEqual(hmFpf, tmFpf, imFpfRecLen, tmFpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    End If
    mMoveCtrlToRec
    Screen.MousePointer = vbHourglass  'Wait
    If imTerminate Then
        Screen.MousePointer = vbDefault    'Default
        mSaveRec = False
        Exit Function
    End If
    ilFpfCode = cbcSelect.ItemData(imFpfSelectIndex)
    If imFpfSelectIndex = 0 Then
        tmFpf.iCode = 0
        ilRet = btrInsert(hmFpf, tmFpf, imFpfRecLen, INDEXKEY0)
    Else
        ilRet = btrUpdate(hmFpf, tmFpf, imFpfRecLen)
    End If
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        ilCRet = btrAbortTrans(hmFdf)
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at FPF-1", vbOkOnly + vbExclamation, "Pledge")
        Exit Function
    End If
    ilLoop = 0
    For ilFdf = LBound(tgPlgeRec) To UBound(tgPlgeRec) - 1 Step 1
        Do  'Loop until record updated or added
            tgPlgeRec(ilFdf).tFdf.iFpfCode = tmFpf.iCode
            If (tgPlgeRec(ilFdf).iStatus = 0) Then  'New selected
                'User
'                gPackDate smNowDate, tgPlgeRec(ilFdf).tFdf.iDateEntrd(0), tgPlgeRec(ilFdf).tFdf.iDateEntrd(1)
                tgPlgeRec(ilFdf).tFdf.iCode = 0
'                tgPlgeRec(ilFdf).tFdf.iSlfCode = ilSlfCode
'                tgPlgeRec(ilFdf).tFdf.iUrfCode = tgUrf(0).iCode
                ilRet = btrInsert(hmFdf, tgPlgeRec(ilFdf).tFdf, imFdfRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmFpf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 2", vbOkOnly + vbExclamation, "Pledge")
                    Exit Function
                End If
                slMsg = "mSaveRec (btrInsert: Sales Commission)"
                ilRet = btrGetPosition(hmFpf, tgPlgeRec(ilFdf).lRecPos)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmFpf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 3", vbOkOnly + vbExclamation, "Pledge")
                    Exit Function
                End If
                tgPlgeRec(ilFdf).iStatus = 1
            ElseIf (tgPlgeRec(ilFdf).iStatus = 1) Then  'Old record-Update
                slMsg = "mSaveRec (btrGetDirect: Sales Commission)"
                tmFdfSrchKey.iCode = tgPlgeRec(ilFdf).tFdf.iCode
                ilRet = btrGetEqual(hmFdf, tlFdf, imFdfRecLen, tmFdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmFdf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 4", vbOkOnly + vbExclamation, "Pledge")
                    Exit Function
                End If
                tlFdf1 = tlFdf
                tlFdf2 = tgPlgeRec(ilFdf).tFdf
                If StrComp(tlFdf1.sChar, tlFdf2.sChar, 0) <> 0 Then
'                    tgPlgeRec(ilFdf).tFdf.iUrfCode = tgUrf(0).iCode
                    ilRet = btrUpdate(hmFdf, tgPlgeRec(ilFdf).tFdf, imFdfRecLen)
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrUpdate: Sales Commission)"
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            ilCRet = btrAbortTrans(hmFdf)
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 5", vbOkOnly + vbExclamation, "Pledge")
            Exit Function
        End If
    Next ilFdf
    For ilFdf = LBound(tgPlgeDel) To UBound(tgPlgeDel) - 1 Step 1
        If tgPlgeDel(ilFdf).iStatus = 1 Then
            Do
                slMsg = "mSaveRec (btrGetEqual: Sales Commission)"
                tmFdfSrchKey.iCode = tgPlgeDel(ilFdf).tFdf.iCode
                ilRet = btrGetEqual(hmFdf, tmFdf, imFdfRecLen, tmFdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmFdf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 6", vbOkOnly + vbExclamation, "Pledge")
                    Exit Function
                End If
                ilRet = btrDelete(hmFpf)
                slMsg = "mSaveRec (btrDelete: Sales Commission)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                ilCRet = btrAbortTrans(hmFpf)
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 7", vbOkOnly + vbExclamation, "Pledge")
                Exit Function
            End If
        End If
    Next ilFdf
    ilRet = btrEndTrans(hmFdf)
    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function

    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if record altered and*
'*                      requires updating              *
'*                                                     *
'*******************************************************
Private Function mSaveRecChg(ilAsk As Integer) As Integer
'
'   iAsk = True
'   iRet = mSaveRecChg(iAsk)
'   Where:
'       iAsk (I)- True = Ask if changed records should be updated;
'                 False= Update record if required without asking user
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRes As Integer
    Dim slMess As String
    Dim ilLoop As Integer
    Dim ilNew As Integer
    If (imFdfChg Or imFpfChg) And (UBound(tgPlgeRec) > LBound(tgPlgeRec)) Or (UBound(tgPlgeDel) > LBound(tgPlgeDel)) Then
        If ilAsk Then
            ilNew = True
            For ilLoop = LBound(tgPlgeRec) To UBound(tgPlgeRec) - 1 Step 1
                If tgPlgeRec(ilLoop).iStatus <> 0 Then
                    ilNew = False
                    Exit For
                End If
            Next ilLoop
            For ilLoop = LBound(tgPlgeDel) To UBound(tgPlgeDel) - 1 Step 1
                If tgPlgeDel(ilLoop).iStatus <> 0 Then
                    ilNew = False
                    Exit For
                End If
            Next ilLoop
            If Not ilNew Then
                slMess = "Save Changes"
            Else
                slMess = "Add Changes"
            End If
            ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
            If ilRes = vbCancel Then
                mSaveRecChg = False
                Exit Function
            End If
            If ilRes = vbYes Then
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
        Else
            ilRes = mSaveRec()
            mSaveRecChg = ilRes
            Exit Function
        End If
    End If
    mSaveRecChg = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:5/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mSetChg(ilBoxNo As Integer)
'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    Dim ilAltered As Integer
    If (imBypassSetting) Or (Not imUpdateAllowed) Then
        Exit Sub
    End If
    ilAltered = imFdfChg
    If imFpfChg Then
        ilAltered = True
    End If
    If (Not ilAltered) And (UBound(tgPlgeDel) > LBound(tgPlgeDel)) Then
        ilAltered = True
    End If
    If ilAltered Then
        pbcPledge.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        cbcSelect.Enabled = False
    Else
        If (imFpfSelectIndex < 0) Then
            pbcPledge.Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            cbcSelect.Enabled = True
        Else
            pbcPledge.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            cbcSelect.Enabled = True
        End If
    End If
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields() = YES) And (ilAltered) And (UBound(tgPlgeRec) > 1) And (imUpdateAllowed) Then
        cmcSave.Enabled = True
    Else
        cmcSave.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case FDMOINDEX To FDSUINDEX
            ckcAirDay.Visible = True
            ckcAirDay.SetFocus
        Case FDSTIMEINDEX
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case FDETIMEINDEX
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PDMOINDEX To PDSUINDEX
            ckcAirDay.Visible = True
            ckcAirDay.SetFocus
        Case PDSTIMEINDEX
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PDETIMEINDEX
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetMinMax                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set scroll bar min/max         *
'*                                                     *
'*******************************************************
Private Sub mSetMinMax()
    imSettingValue = True
    vbcPledge.Min = LBound(smShow, 2)
    imSettingValue = True
    If UBound(smShow, 2) - 1 <= vbcPledge.LargeChange + 1 Then ' + 1 Then
        vbcPledge.Max = LBound(smShow, 2)
    Else
        vbcPledge.Max = UBound(smShow, 2) - vbcPledge.LargeChange
    End If
    imSettingValue = True
    If vbcPledge.Value = vbcPledge.Min Then
        vbcPledge_Change
    Else
        vbcPledge.Value = vbcPledge.Min
    End If
    imSettingValue = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    pbcArrow.Visible = False
    lacFrame.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case FDMOINDEX To FDSUINDEX
            ckcAirDay.Visible = False
            If ckcAirDay.Value = vbChecked Then
                If smSave(1 + ilBoxNo - FDMOINDEX, imRowNo) <> "Y" Then
                    imFdfChg = True
                    smSave(1 + ilBoxNo - FDMOINDEX, imRowNo) = "Y"
                End If
            Else
                If smSave(1 + ilBoxNo - FDMOINDEX, imRowNo) <> "N" Then
                    imFdfChg = True
                    smSave(1 + ilBoxNo - FDMOINDEX, imRowNo) = "N"
                End If
            End If
            slStr = smSave(1 + ilBoxNo - FDMOINDEX, imRowNo)
            gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
        Case FDSTIMEINDEX
            plcTme.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidTime(slStr) Then
                gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
                If Trim$(smSave(8, imRowNo)) <> slStr Then
                    imFdfChg = True
                    smSave(8, imRowNo) = slStr
                End If
            End If
        Case FDETIMEINDEX
            plcTme.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidTime(slStr) Then
                gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
                If Trim$(smSave(9, imRowNo)) <> slStr Then
                    imFdfChg = True
                    smSave(9, imRowNo) = slStr
                End If
            End If
        Case PDMOINDEX To PDSUINDEX
            ckcAirDay.Visible = False
            If ckcAirDay.Value = vbChecked Then
                If smSave(10 + ilBoxNo - PDMOINDEX, imRowNo) <> "Y" Then
                    imFdfChg = True
                    smSave(10 + ilBoxNo - PDMOINDEX, imRowNo) = "Y"
                End If
            Else
                If smSave(10 + ilBoxNo - PDMOINDEX, imRowNo) <> "N" Then
                    imFdfChg = True
                    smSave(10 + ilBoxNo - PDMOINDEX, imRowNo) = "N"
                End If
            End If
            slStr = smSave(10 + ilBoxNo - PDMOINDEX, imRowNo)
            gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
        Case PDSTIMEINDEX
            plcTme.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidTime(slStr) Then
                gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
                If Trim$(smSave(17, imRowNo)) <> slStr Then
                    imFdfChg = True
                    smSave(17, imRowNo) = slStr
                End If
            End If
        Case PDETIMEINDEX
            plcTme.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidTime(slStr) Then
                gSetShow pbcPledge, slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
                If Trim$(smSave(18, imRowNo)) <> slStr Then
                    imFdfChg = True
                    smSave(18, imRowNo) = slStr
                End If
            End If
    End Select
    mSetCommands
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
    Dim ilRet As Integer
    Erase tgPlgeRec
    Erase tgPlgeDel

    btrExtClear hmFpf   'Clear any previous extend operation
    ilRet = btrClose(hmFpf)
    btrExtClear hmFdf   'Clear any previous extend operation
    ilRet = btrClose(hmFdf)
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload FeedPlge
    Set FeedPlge = Nothing
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
    Dim ilRowNo As Integer
    For ilRowNo = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
        If Trim$(smSave(8, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSave(9, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSave(17, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSave(18, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
    Next ilRowNo
    mTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSaveFields                 *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields(ilRowNo As Integer) As Integer
'
'   iRet = mTestSaveFields(ilRowNo)
'   Where:
'       ilRowNo (I)- row number to be checked
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilLoop As Integer
    Dim ilDayFound As Integer

    If (Not gValidTime(smSave(8, ilRowNo))) Or (Trim$(smSave(8, ilRowNo)) = "") Then
        Beep
        ilRes = MsgBox("Feed Start Time must be specified correctly", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = FDSTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If (Not gValidTime(smSave(9, ilRowNo))) Or (Trim$(smSave(9, ilRowNo)) = "") Then
        Beep
        ilRes = MsgBox("Feed End Time must be specified correctly", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = FDSTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    'Check times
    If gTimeToLong(smSave(8, ilRowNo), False) > gTimeToLong(smSave(9, ilRowNo), True) Then
        Beep
        ilRes = MsgBox("Feed Start Time must be prior to End Time", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = FDSTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If (Not gValidTime(smSave(17, ilRowNo))) Or (Trim$(smSave(17, ilRowNo)) = "") Then
        Beep
        ilRes = MsgBox("Pledge Start Time must be specified correctly", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = FDSTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If (Not gValidTime(smSave(18, ilRowNo))) Or (Trim$(smSave(18, ilRowNo)) = "") Then
        Beep
        ilRes = MsgBox("Pledge End Time must be specified correctly", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = FDSTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    'Check times
    If gTimeToLong(smSave(17, ilRowNo), False) > gTimeToLong(smSave(18, ilRowNo), True) Then
        Beep
        ilRes = MsgBox("Pledge Start Time must be prior to End Time", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = FDSTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    ilDayFound = False
    For ilLoop = 0 To 6 Step 1
        If smSave(1 + ilLoop, ilRowNo) <> "N" Then
            ilDayFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilDayFound Then
        Beep
        ilRes = MsgBox("Feed Week Day must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = FDMOINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    ilDayFound = False
    For ilLoop = 0 To 6 Step 1
        If smSave(10 + ilLoop, ilRowNo) <> "N" Then
            ilDayFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilDayFound Then
        Beep
        ilRes = MsgBox("Pledge Week Day must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = FDMOINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    mTestSaveFields = YES
End Function
Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0

    llDate = lmCalStartDate

    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(Str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcSpecDropdown.Text = Format$(llDate, "m/d/yy")
                edcSpecDropdown.SelStart = 0
                edcSpecDropdown.SelLength = Len(edcSpecDropdown.Text)
                imBypassFocus = True
                edcSpecDropdown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate

    edcSpecDropdown.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(Str$(imCalMonth)) & "/15/" & Trim$(Str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    plcCalendar.Visible = False
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcClickFocus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcLen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcLenInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 5 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcLenInv.Move flX, flY
                    imcLenInv.Visible = True
                    imcLenOutline.Move flX - 15, flY - 15
                    imcLenOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub pbcLen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcLenInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 5 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcLenInv.Move flX, flY
                    imcLenOutline.Move flX - 15, flY - 15
                    imcLenOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "H"
                                Case 2
                                    slKey = "M"
                                Case 3
                                    slKey = "S"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                    End Select
                    imBypassFocus = True    'Don't change select text
                    edcQAdj.SetFocus
                    'SendKeys slKey
                    gSendKeys edcQAdj, slKey
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub pbcPledge_GotFocus()
    plcLen.Visible = False
    mMakeAdj
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
End Sub

Private Sub pbcPledge_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcPledge_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If Button = 2 Then
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    ilCompRow = vbcPledge.LargeChange + 1
    If UBound(tgPlgeRec) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tgPlgeRec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcPledge.Value - 1
                    If ilRowNo > UBound(smSave, 2) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    mSpecSetShow imSpecBoxNo
                    imSpecBoxNo = -1
                    mSetShow imBoxNo
                    imRowNo = ilRow + vbcPledge.Value - 1
                    If (imRowNo = UBound(smSave, 2)) And (Trim$(smSave(1, imRowNo)) = "") Then
                        mInitNew imRowNo
                    End If
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    If imSpecBoxNo > 0 Then
        mSpecSetFocus imSpecBoxNo
    ElseIf imBoxNo > 0 Then
        mSetFocus imBoxNo
    End If
End Sub
Private Sub pbcPledge_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long

    ilStartRow = vbcPledge.Value '+ 1  'Top location
    ilEndRow = vbcPledge.Value + vbcPledge.LargeChange ' + 1
    If ilEndRow > UBound(smSave, 2) Then
        If Trim$(smShow(1, UBound(smShow, 2))) <> "" Then
            ilEndRow = UBound(smSave, 2) 'include blank row as it might have data
        Else
            ilEndRow = UBound(smSave, 2) - 1
        End If
    End If
    llColor = pbcPledge.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        If ilRow = UBound(smSave, 2) Then
            pbcPledge.ForeColor = DARKPURPLE
        Else
            pbcPledge.ForeColor = llColor
        End If
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            pbcPledge.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcPledge.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            slStr = Trim$(smShow(ilBox, ilRow))
            pbcPledge.Print slStr
        Next ilBox
    Next ilRow
    pbcPledge.ForeColor = llColor
End Sub
Private Sub pbcNum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    imcNumInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 4 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            For ilColNo = 1 To 3 Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcNumInv.Move flX, flY
                    imcNumInv.Visible = True
                    imcNumOutline.Move flX - 15, flY - 15
                    imcNumOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcNum_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    imcNumInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 4 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            For ilColNo = 1 To 3 Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcNumInv.Move flX, flY
                    imcNumOutline.Move flX - 15, flY - 15
                    imcNumOutline.Visible = True
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
                                Case 3
                                    slKey = "."
                            End Select
                    End Select
                    imBypassFocus = True    'Don't change select text
                    edcDropDown.SetFocus
                    'SendKeys slKey
                    gSendKeys edcDropDown, slKey
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub pbcSpec_GotFocus()
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    plcLen.Visible = False
    mMakeAdj
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
End Sub

Private Sub pbcSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim slStr As String
    For ilBox = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        If (X >= tmSpecCtrls(ilBox).fBoxX) And (X <= tmSpecCtrls(ilBox).fBoxX + tmSpecCtrls(ilBox).fBoxW) Then
            If (Y >= tmSpecCtrls(ilBox).fBoxY) And (Y <= tmSpecCtrls(ilBox).fBoxY + tmSpecCtrls(ilBox).fBoxH) Then
                If (imSpecBoxNo = SPECSTARTDATEINDEX) And (edcSpecDropdown.Text <> "") Then
                    slStr = edcSpecDropdown.Text
                    If Not gValidDate(slStr) Then
                        Beep
                        edcSpecDropdown.SetFocus
                        Exit Sub
                    End If
                End If
                If (imSpecBoxNo = SPECENDDATEINDEX) And (edcSpecDropdown.Text <> "") Then
                    slStr = edcSpecDropdown.Text
                    If Not gValidDate(slStr) Then
                        Beep
                        edcSpecDropdown.SetFocus
                        Exit Sub
                    End If
                End If
                mSetShow imBoxNo
                imBoxNo = -1
                imRowNo = -1
                pbcArrow.Visible = False
                lacFrame.Visible = False
                mSpecSetShow imSpecBoxNo
                imSpecBoxNo = ilBox
                mSpecEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    If imSpecBoxNo > 0 Then
        mSpecSetFocus imSpecBoxNo
    ElseIf imBoxNo > 0 Then
        mSetFocus imBoxNo
    End If
End Sub

Private Sub pbcSpec_Paint()
    Dim ilBox As Integer
    pbcSpec.Cls
    For ilBox = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        pbcSpec.CurrentX = tmSpecCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcSpec.CurrentY = tmSpecCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcSpec.Print tmSpecCtrls(ilBox).sShow
    Next ilBox
End Sub

Private Sub pbcSpecSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSpecSTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = -1  'Set-Right to left
    plcLen.Visible = False
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    If (imSpecBoxNo >= imLBSpecCtrls) And (imSpecBoxNo <= UBound(tmSpecCtrls)) Then
        If (Not cbcSelect.Enabled) Then
            If mSpecTestFields(imSpecBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mSpecEnableBox imSpecBoxNo
                Exit Sub
            End If
        End If
    End If
    Select Case imSpecBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            ilBox = SPECSTARTDATEINDEX
        Case SPECSTARTDATEINDEX 'Type (first control within header)
            mSpecSetShow imSpecBoxNo
            imSpecBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            cmcCancel.SetFocus
        Case Else
            ilBox = imSpecBoxNo - 1
    End Select
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = ilBox
    mSpecEnableBox ilBox

End Sub

Private Sub pbcSpecTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSpecTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = 0  'Set-Left to right
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    mSpecSetShow imSpecBoxNo
    If (imSpecBoxNo >= imLBSpecCtrls) And (imSpecBoxNo <= UBound(tmSpecCtrls)) Then
        If mSpecTestFields(imSpecBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            imDirProcess = -1
            mSpecEnableBox imSpecBoxNo
            Exit Sub
        End If
    End If
    Select Case imSpecBoxNo
        Case -1 'Shift tab from button
            imTabDirection = -1  'Set-Right to left
            ilBox = SPECENDDATEINDEX
        Case SPECENDDATEINDEX    'last control
            imSpecBoxNo = -1
            If mSpecTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
                Beep
                mSpecEnableBox imSpecBoxNo
                Exit Sub
            End If
            If pbcSTab.Enabled Then
                pbcSTab.SetFocus
            Else
                pbcSpecSTab.SetFocus
            End If
            Exit Sub
        Case Else
            ilBox = imSpecBoxNo + 1
    End Select
    imSpecBoxNo = ilBox
    mSpecEnableBox ilBox
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    plcLen.Visible = False
    mMakeAdj
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    imTabDirection = -1 'Set- Right to left
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            If (UBound(smSave, 2) = 1) Then
                imTabDirection = 0  'Set-Left to right
                imRowNo = 1
                mInitNew imRowNo
            Else
                If UBound(smSave, 2) <= vbcPledge.LargeChange Then 'was <=
                    vbcPledge.Max = LBound(smSave, 2)
                Else
                    vbcPledge.Max = UBound(smSave, 2) - vbcPledge.LargeChange '- 1
                End If
                imRowNo = 1
                If imRowNo >= UBound(smSave, 2) Then
                    mInitNew imRowNo
                End If
                imSettingValue = True
                vbcPledge.Value = vbcPledge.Min
                imSettingValue = False
            End If
            ilBox = FDMOINDEX
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case FDMOINDEX, 0
            mSetShow imBoxNo
            If (imBoxNo < 1) And (imRowNo < 1) Then 'Modelled from Proposal
                Exit Sub
            End If
            ilBox = PDETIMEINDEX
            If imRowNo <= 1 Then
                imBoxNo = -1
                imRowNo = -1
                cmcDone.SetFocus
                Exit Sub
            End If
            imRowNo = imRowNo - 1
            If imRowNo < vbcPledge.Value Then
                imSettingValue = True
                vbcPledge.Value = vbcPledge.Value - 1
                imSettingValue = False
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
'        Case SDATEINDEX
'            slStr = edcDropDown.Text
'            If slStr <> "" Then
'                If Not gValidDate(slStr) Then
'                    Beep
'                    edcDropDown.SetFocus
'                    Exit Sub
'                End If
'            Else
'                Beep
'                edcDropDown.SetFocus
'                Exit Sub
'            End If
'            ilBox = imBoxNo - 1
'        Case EDATEINDEX
'            slStr = edcDropDown.Text
'            If (slStr <> "") And (StrComp(slStr, "TFN", 1) <> 0) Then
'                If Not gValidDate(slStr) Then
'                    Beep
'                    edcDropDown.SetFocus
'                    Exit Sub
'                End If
'            End If
'            ilBox = imBoxNo - 1
        Case FDSTIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo - 1
        Case FDETIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo - 1
        Case PDSTIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo - 1
        Case PDETIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo - 1
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcStartNew_GotFocus()
    Dim ilRet As Integer
    If imInNew Then
        Exit Sub
    End If
    plcLen.Visible = False
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    If (imFpfSelectIndex = 0) And (imFirstTimeSelect) Then
        imFirstTimeSelect = False
        ilRet = mStartNew()
        If Not ilRet Then
            imTerminate = True
            mTerminate
            Exit Sub
        End If
    End If
    mSetCommands
    pbcSpecSTab.SetFocus
End Sub

Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim slStr As String

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            imRowNo = UBound(smSave, 2) - 1
            imSettingValue = True
            If imRowNo <= vbcPledge.LargeChange + 1 Then
                vbcPledge.Value = 1
            Else
                vbcPledge.Value = imRowNo - vbcPledge.LargeChange - 1
            End If
            imSettingValue = False
            ilBox = PDETIMEINDEX
        Case PDETIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            mSetShow imBoxNo
            If mTestSaveFields(imRowNo) = NO Then
                mEnableBox imBoxNo
                Exit Sub
            End If
            If imRowNo >= UBound(smSave, 2) Then
                imFdfChg = True
                ReDim Preserve smShow(1 To 18, 1 To imRowNo + 1) As String 'Values shown in program area
                ReDim Preserve smSave(1 To 18, 1 To imRowNo + 1) As String 'Values saved (program name) in program area
                For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                    smShow(ilLoop, imRowNo + 1) = ""
                Next ilLoop
                For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                    smSave(ilLoop, imRowNo + 1) = ""
                Next ilLoop
                ReDim Preserve tgPlgeRec(1 To UBound(tgPlgeRec) + 1) As FDFREC
                tgPlgeRec(UBound(tgPlgeRec)).iStatus = 0
                tgPlgeRec(UBound(tgPlgeRec)).lRecPos = 0
            End If
            If imRowNo >= UBound(smSave, 2) - 1 Then
                imRowNo = imRowNo + 1
                mInitNew imRowNo
                If UBound(smSave, 2) <= vbcPledge.LargeChange Then 'was <=
                    vbcPledge.Max = LBound(smSave, 2) '- 1
                Else
                    vbcPledge.Max = UBound(smSave, 2) - vbcPledge.LargeChange '- 1
                End If
            Else
                imRowNo = imRowNo + 1
            End If
            If imRowNo > vbcPledge.Value + vbcPledge.LargeChange Then
                imSettingValue = True
                vbcPledge.Value = vbcPledge.Value + 1
                imSettingValue = False
            End If
            If imRowNo >= UBound(smSave, 2) Then
                imBoxNo = 0
                mSetCommands
                'lacFrame.Move 0, tmCtrls(PROPNOINDEX).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15) - 30
                'lacFrame.Visible = True
                pbcArrow.Move pbcArrow.Left, plcPledge.Top + tmCtrls(FDMOINDEX).fBoxY + (imRowNo - vbcPledge.Value) * (fgBoxGridH + 15) + 45
                pbcArrow.Visible = True
                pbcArrow.SetFocus
                Exit Sub
            Else
                ilBox = FDMOINDEX
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case 0
            ilBox = FDMOINDEX
        Case FDSTIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo + 1
        Case FDETIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo + 1
        Case PDSTIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo + 1
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
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
                    Select Case imBoxNo
                        Case FDSTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                         Case FDETIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                        Case PDSTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                         Case PDETIMEINDEX
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
End Sub

Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub tmcDrag_Timer()
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            ilCompRow = vbcPledge.LargeChange + 1
            If UBound(smSave, 2) > ilCompRow Then
                ilMaxRow = ilCompRow
            Else
                ilMaxRow = UBound(smSave, 2)
            End If
            For ilRow = 1 To ilMaxRow Step 1
                If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(FDMOINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(FDMOINDEX).fBoxY + tmCtrls(FDMOINDEX).fBoxH)) Then
                    mSetShow imBoxNo
                    imBoxNo = -1
                    imRowNo = -1
                    imRowNo = ilRow + vbcPledge.Value - 1
                    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                    lacFrame.Move 0, tmCtrls(FDMOINDEX).fBoxY + (imRowNo - vbcPledge.Value) * (fgBoxGridH + 15) - 30
                    'If gInvertArea call then remove visible setting
                    lacFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcPledge.Top + tmCtrls(FDMOINDEX).fBoxY + (imRowNo - vbcPledge.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    lacFrame.Drag vbBeginDrag
                    lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                    Exit Sub
                End If
            Next ilRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub vbcPledge_Change()
    If imSettingValue Then
        pbcPledge.Cls
        pbcPledge_Paint
        imSettingValue = False
    Else
        mSetShow imBoxNo
        imBoxNo = -1
        imRowNo = -1
        pbcPledge.Cls
        pbcPledge_Paint
    End If
End Sub
Private Sub vbcPledge_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Function mOverlappingTimes() As Integer
    Dim ilRowNo As Integer
    Dim ilLoop As Integer
    Dim llSTime As Long
    Dim llETime As Long
    Dim llTTime As Long
    Dim ilDay As Integer
    Dim ilRet As Integer

    For ilRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        llSTime = gTimeToLong(smSave(8, ilRowNo), False)
        llETime = gTimeToLong(smSave(9, ilRowNo), True)
        For ilLoop = ilRowNo + 1 To UBound(smSave, 2) - 1 Step 1
            llTTime = gTimeToLong(smSave(8, ilLoop), False)
            If (llTTime >= llSTime) And (llTTime < llETime) Then
                For ilDay = 0 To 6 Step 1
                    If smSave(1 + ilDay, ilRowNo) = smSave(1 + ilDay, ilLoop) And smSave(1 + ilDay, ilRowNo) = "Y" Then
                        ilRet = MsgBox("Feed Times " & smSave(8, ilRowNo) & "-" & smSave(9, ilRowNo) & " in Conflict with " & smSave(17, ilLoop) & "-" & smSave(18, ilLoop) & ", Save terminated.", vbOkOnly + vbExclamation, "Conflict")
                        mOverlappingTimes = True
                        imRowNo = ilLoop
                        Exit Function
                    End If
                Next ilDay
            End If
            llTTime = gTimeToLong(smSave(9, ilLoop), True)
            If (llTTime > llSTime) And (llTTime <= llETime) Then
                For ilDay = 0 To 6 Step 1
                    If smSave(1 + ilDay, ilRowNo) = smSave(1 + ilDay, ilLoop) And smSave(1 + ilDay, ilRowNo) = "Y" Then
                        ilRet = MsgBox("Feed Times " & smSave(8, ilRowNo) & "-" & smSave(9, ilRowNo) & " in Conflict with " & smSave(17, ilLoop) & "-" & smSave(18, ilLoop) & ", Save terminated.", vbOkOnly + vbExclamation, "Conflict")
                        mOverlappingTimes = True
                        imRowNo = ilLoop
                        Exit Function
                    End If
                Next ilDay
            End If
        Next ilLoop
    Next ilRowNo
    mOverlappingTimes = False
End Function

Private Sub mMakeAdj()
    Dim ilRowNo As Integer
    Dim slLen As String
    Dim slStr As String
    Dim slNeg As String
    Dim slXMid As String

    If Not imAdjRequired Then
       Exit Sub
    End If
    slLen = edcQAdj.Text
    If InStr(1, slLen, "-", vbTextCompare) = 1 Then
        slNeg = "-"
        slLen = Mid$(slLen, 2)
    Else
        slNeg = ""
    End If
    If Not gValidLength(slLen) Then
        Exit Sub
    End If
    slLen = slNeg & slLen
    For ilRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        gAddTimeLength smSave(17, ilRowNo), slLen, "A", "1", smSave(17, ilRowNo), slXMid
        gAddTimeLength smSave(18, ilRowNo), slLen, "A", "1", smSave(18, ilRowNo), slXMid
        slStr = Trim$(smSave(17, ilRowNo))
        gSetShow pbcPledge, slStr, tmCtrls(PDSTIMEINDEX)
        smShow(PDSTIMEINDEX, ilRowNo) = tmCtrls(PDSTIMEINDEX).sShow
        slStr = Trim$(smSave(18, ilRowNo))
        gSetShow pbcPledge, slStr, tmCtrls(PDETIMEINDEX)
        smShow(PDETIMEINDEX, ilRowNo) = tmCtrls(PDETIMEINDEX).sShow
    Next ilRowNo
    imFdfChg = True
    edcQAdj.Text = ""
    pbcPledge.Cls
    pbcPledge_Paint
    mSetCommands
End Sub

Private Function mOverlappingDates() As Integer
    Dim ilLoop As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llTStartDate As Long
    Dim llTEndDate As Long
    Dim ilPos As Integer
    Dim slTStartDate As String
    Dim slTEndDate As String
    Dim ilIndex As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim slDates As String

    llStartDate = gDateValue(smSpecSave(1))
    If (Trim$(smSpecSave(2)) = "") Or (StrComp(smSpecSave(2), "TFN", vbTextCompare) = 0) Then
        llEndDate = gDateValue("12/31/2068")
    Else
        llEndDate = gDateValue(smSpecSave(2))
    End If
    ilIndex = -1
    For ilLoop = 1 To cbcSelect.ListCount - 1 Step 1
        If (imFpfSelectIndex <> ilLoop) Then
            slStr = cbcSelect.List(ilLoop)
            ilPos = InStr(1, slStr, "-", vbTextCompare)
            slTStartDate = Left$(slStr, ilPos - 1)
            slTEndDate = Mid$(slStr, ilPos + 1)
            If slTEndDate = "TFN" Then
                slTEndDate = "12/31/2068"
            End If
            llTStartDate = gDateValue(slTStartDate)
            llTEndDate = gDateValue(slTEndDate)
            If (llTEndDate >= llStartDate) And (llEndDate >= llTStartDate) Then
                If (llTStartDate < llStartDate) And (llTEndDate <= llEndDate) Then
                    ilIndex = ilLoop
                Else
                    ilRet = MsgBox("Dates specified conflict with " & slStr & " save terminated", vbOkOnly + vbExclamation, "Conflict")
                    mOverlappingDates = -1
                    Exit Function
                End If
            End If
        End If
    Next ilLoop
    If ilIndex <> -1 Then
        slStr = cbcSelect.List(ilIndex)
        ilPos = InStr(1, slStr, "-", vbTextCompare)
        slTStartDate = Left$(slStr, ilPos - 1)
        slTEndDate = Format$(llStartDate - 1, "m/d/yy")
        slDates = slTStartDate & "-" & slTEndDate
        ilRet = MsgBox("Dates specified in conflict with " & slStr & ", the Pledge dates will be changed to " & slDates, vbOkCancel + vbExclamation, "Conflict")
        If ilRet = vbCancel Then
            mOverlappingDates = -1
            Exit Function
        Else
            mOverlappingDates = cbcSelect.ItemData(ilIndex)
            Exit Function
        End If

    End If
    mOverlappingDates = 0
    Exit Function
End Function
