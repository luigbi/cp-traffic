VERSION 5.00
Begin VB.Form PostRep 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5970
   ClientLeft      =   630
   ClientTop       =   1680
   ClientWidth     =   9435
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5970
   ScaleWidth      =   9435
   Begin VB.TextBox edcWkly 
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
      Left            =   5460
      MaxLength       =   3
      TabIndex        =   15
      Top             =   3900
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox pbcWeekly 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   3960
      Picture         =   "PostRep.frx":0000
      ScaleHeight     =   1440
      ScaleWidth      =   4200
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3315
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.PictureBox plcWeekly 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1785
      Left            =   3795
      ScaleHeight     =   1785
      ScaleWidth      =   4515
      TabIndex        =   12
      Top             =   3105
      Visible         =   0   'False
      Width           =   4515
      Begin VB.PictureBox pbcWklyTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   30
         ScaleHeight     =   105
         ScaleWidth      =   75
         TabIndex        =   16
         Top             =   840
         Width           =   75
      End
      Begin VB.PictureBox pbcWklySTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   45
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   13
         Top             =   225
         Width           =   60
      End
   End
   Begin VB.PictureBox plcInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
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
      Height          =   885
      Left            =   210
      ScaleHeight     =   855
      ScaleWidth      =   9135
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3930
      Visible         =   0   'False
      Width           =   9165
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Line Comment:"
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
         Index           =   2
         Left            =   60
         TabIndex        =   39
         Top             =   585
         Width           =   9030
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Invoice #: xxxxxx  Vehicle Name:"
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
         Left            =   60
         TabIndex        =   31
         Top             =   45
         Width           =   9030
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Contract #: xxxxxx  Check #  Transaction: Date xx/xx/xx  Type xx  Action  xx"
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
         Index           =   1
         Left            =   60
         TabIndex        =   30
         Top             =   315
         Width           =   9030
      End
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
      Left            =   7845
      Picture         =   "PostRep.frx":13B42
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcGross 
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
      Left            =   4995
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1050
      Visible         =   0   'False
      Width           =   1275
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
      Left            =   6720
      MaxLength       =   20
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox edcNoSpots 
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
      HelpContextID   =   8
      Left            =   6645
      MaxLength       =   4
      TabIndex        =   10
      Top             =   1155
      Visible         =   0   'False
      Width           =   960
   End
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
      Height          =   1410
      Left            =   405
      Picture         =   "PostRep.frx":13C3C
      ScaleHeight     =   1380
      ScaleWidth      =   5115
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2505
      Visible         =   0   'False
      Width           =   5145
   End
   Begin VB.PictureBox pbcPost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4290
      Index           =   3
      Left            =   8640
      Picture         =   "PostRep.frx":2AC7E
      ScaleHeight     =   4290
      ScaleWidth      =   8835
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3225
      Visible         =   0   'False
      Width           =   8835
   End
   Begin VB.PictureBox pbcPost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4290
      Index           =   2
      Left            =   8880
      Picture         =   "PostRep.frx":A6868
      ScaleHeight     =   4290
      ScaleWidth      =   8835
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2820
      Visible         =   0   'False
      Width           =   8835
   End
   Begin VB.PictureBox pbcPost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4290
      Index           =   1
      Left            =   8955
      Picture         =   "PostRep.frx":122452
      ScaleHeight     =   4290
      ScaleWidth      =   8835
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   8835
   End
   Begin VB.PictureBox pbcPost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4290
      Index           =   0
      Left            =   9060
      Picture         =   "PostRep.frx":19E03C
      ScaleHeight     =   4290
      ScaleWidth      =   8835
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   8835
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8355
      Top             =   5370
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "PostRep.frx":219C26
      Left            =   2400
      List            =   "PostRep.frx":219C2D
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
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
      Left            =   30
      Picture         =   "PostRep.frx":219C3D
      ScaleHeight     =   180
      ScaleWidth      =   90
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   90
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
      Left            =   8775
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4230
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4140
      TabIndex        =   24
      Top             =   5400
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
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4560
      Width           =   75
   End
   Begin VB.CommandButton cmcImport 
      Appearance      =   0  'Flat
      Caption         =   "&Import"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5460
      TabIndex        =   25
      Top             =   5400
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
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   17
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
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   105
      TabIndex        =   4
      Top             =   525
      Width           =   105
   End
   Begin VB.PictureBox plcSelect 
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
      Height          =   405
      Left            =   2895
      ScaleHeight     =   345
      ScaleWidth      =   6345
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   75
      Width           =   6405
      Begin VB.ComboBox cbcInvDate 
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
         Left            =   30
         TabIndex        =   2
         Top             =   15
         Width           =   2565
      End
      Begin VB.ComboBox cbcNames 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2895
         TabIndex        =   3
         Top             =   15
         Width           =   3435
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2820
      TabIndex        =   23
      Top             =   5400
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1440
      TabIndex        =   22
      Top             =   5400
      Width           =   1050
   End
   Begin VB.PictureBox pbcPostRep 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   180
      Picture         =   "PostRep.frx":219F47
      ScaleHeight     =   4290
      ScaleWidth      =   8835
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   615
      Width           =   8835
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
         Left            =   0
         TabIndex        =   21
         Top             =   390
         Visible         =   0   'False
         Width           =   8820
      End
   End
   Begin VB.CommandButton cmcClear 
      Appearance      =   0  'Flat
      Caption         =   "C&lear"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      TabIndex        =   36
      Top             =   5400
      Width           =   945
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Received"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   135
      TabIndex        =   37
      Top             =   5025
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.PictureBox pbcPost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4290
      Index           =   4
      Left            =   9135
      Picture         =   "PostRep.frx":295B31
      ScaleHeight     =   4290
      ScaleWidth      =   8835
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2025
      Visible         =   0   'False
      Width           =   8835
   End
   Begin VB.VScrollBar vbcPostRep 
      Height          =   4275
      LargeChange     =   19
      Left            =   9015
      TabIndex        =   18
      Top             =   630
      Width           =   270
   End
   Begin VB.PictureBox plcPostRep 
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
      Height          =   4440
      Left            =   135
      ScaleHeight     =   4380
      ScaleWidth      =   9150
      TabIndex        =   5
      Top             =   570
      Width           =   9210
   End
   Begin VB.Label lacCalDates 
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
      Height          =   180
      Left            =   120
      TabIndex        =   38
      Top             =   5055
      Visible         =   0   'False
      Width           =   7620
   End
   Begin VB.Label lacScreen 
      Caption         =   "Remote Invoice Posting"
      Height          =   225
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2130
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   150
      Picture         =   "PostRep.frx":31171B
      Top             =   5250
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8850
      Picture         =   "PostRep.frx":311FE5
      Top             =   5190
      Width           =   480
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   195
      Picture         =   "PostRep.frx":3122EF
      Top             =   330
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   5310
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacTotals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7905
      TabIndex        =   20
      Top             =   4275
      Width           =   210
   End
End
Attribute VB_Name = "PostRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PostRep.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmRwfSrchKey                                                                          *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mWklySetFocus                                                                         *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PostRep.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract revision number increment screen code
Option Explicit
Option Compare Text

Dim hmMsg As Integer   'From file handle
Dim hmFrom As Integer

Dim tmMktCode() As SORTCODE
Dim imMktCode As Integer
Dim imMktVefCode() As Integer

Dim tmInvVehicle() As SORTCODE
Dim smInvVehicleTag As String

'Library calendar
Dim hmLcf As Integer        'Library calendar file handle
Dim tmLcf As LCF
Dim imLcfRecLen As Integer

Dim tmContract() As SORTCODE
Dim tmChfAdvtExt() As CHFADVTEXT

Dim tmVehicle() As SORTCODE
Dim imUpdateAllowed As Integer
Dim imFirstActivate As Integer

Dim tmAdvertiser() As SORTCODE
Dim smAdvertiserTag As String


'Billing Items
Dim tmCtrls(0 To 14) As FIELDAREA
Dim imLBCtrls As Integer
Dim tmWklyCtrls(0 To 5) As FIELDAREA
Dim imLBWklyCtrls As Integer

Dim hmCHF As Integer
Dim tmChf As CHF
Dim tmChfSrchKey As LONGKEY0
Dim tmChfSrchKey1 As CHFKEY1
Dim imCHFRecLen As Integer

Dim hmClf As Integer
Dim tmClf As CLF
Dim imClfRecLen As Integer

Dim hmCff As Integer            'Contract line flight file handle
Dim imCffRecLen As Integer        'CFF record length
Dim tmCff As CFF

Dim tmCxf As CXF            'CXF record image
Dim tmCxfSrchKey As LONGKEY0  'CXF key record image
Dim hmCxf As Integer        'CXF Handle
Dim imCxfRecLen As Integer      'CXF record length

Dim hmSbf As Integer        'Special billing
Dim tmSbf As SBF    'SBF record image of billing items
Dim tmSbfSrchKey1 As LONGKEY0            'SBF record image
Dim tmSbfSrchKey2 As SBFKEY2    'SBF key record image
Dim imSbfRecLen As Integer        'SBF record length
Dim lmSbfDel() As Long      'SbfCode of records to be deleted

Dim hmVsf As Integer            'Virtual Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim tmVsfSrchKey As LONGKEY0            'VSF record image
Dim imVsfRecLen As Integer        'VSF record length

Dim hmRwf As Integer            'Rep by Week file handle
Dim tmRwf As RWF                'RWF record image
Dim tmRwfSrchKey1 As RWFKEY1            'RWF record image
Dim imRwfRecLen As Integer        'RWF record length

Dim hmNrf As Integer            'Contract line flight file handle
Dim imNrfRecLen As Integer        'CFF record length
Dim tmNrf As NRF
Dim tmNetNames() As SORTCODE

Dim hmRdf As Integer            'Contract line flight file handle
Dim imRdfRecLen As Integer        'CFF record length
Dim tmRdf As RDF
Dim tmRdfSrchKey As INTKEY0            'RWF record image

Dim hmRvf As Integer            'Receivable file handle
Dim tmRvf As RVF                'RVF record image
Dim tmRvfSrchKey4 As RVFKEY4            'RVF record image (Advertiser code)
Dim imRvfRecLen As Integer        'RVF record length

Dim imMarketIndex As Integer
Dim imInvDateIndex As Integer

Dim imButton As Integer 'Value 1= Left button; 2=Right button; 4=Middle button
Dim imButtonRow As Integer
Dim imIgnoreRightMove As Integer

Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBoxNo As Integer
Dim imRowNo As Integer
Dim smSave() As String      'Values saved (1=C or T; 2=PctTrade; 3=Ordered No Spots; 4=Ordered Gross; 5=Aired No Spots;
                            '6=Aired Gross; 7= Bonus No Spots; 8=Billed, 9=Calendar Carry; 10=Missed carry; 11 Received Date;
                            '12=Posted Date; 13=MGs; 14=Missed; 15=Initial Aired Spots;16=Next billing no spots-for calendar posting); 17=Barter Paid;
                            '18-22=Aired by Week; 23-27=Carried by Week; 28-32=Ordered Spots by Week, 33=first 100 characters of line comment
                            
Dim imSave() As Integer     'Values saved (1=AdfCode; 2=VefCode; 3=lbcVehicle Index; 4=Valid vehicle; 5=Checked; 6=Spot Length if Barter used or Post by Time; 7=Line #)
Dim lmSave() As Long        'Values saved(1=ChfCode; 2=SbfCode; 3=Spot Price; 4=Acquisition Cost, 5=Station Invoice #)
Dim smShow() As String * 40     'Show values (Index number 1-10)
Dim smInfo() As String * 12     'Import Info for Right Mouse (1=Source[I=Import;F=Sbf;C=Contract;T=Total;S=Insert]; 2=Export Date; 3=Import Date;
                            '4=Sbf Date; 5=Combine ID; 6=Ref Inv #; 7=Tax1; 8=Tax2; 9=Ordered Spots, 10=Ordered Gross; 11=Comm Pct, 12=Print Rep Invoice Date, 13 = Gen AR Date)
Dim smWklyDates(0 To 5) As String   'Index zero ignored
Dim imWklyBoxNo As Integer
Dim imWklyRowNo As Integer
Dim imChg As Integer
Dim smStartStd As String    'Starting date for standard billing
Dim smEndStd As String      'Ending date for standard billing
Dim smStartCal As String    'Starting date for standard billing
Dim smEndCal As String      'Ending date for standard billing
Dim lmStartStd As Long    'Starting date for standard billing
Dim lmEndStd As Long      'Ending date for standard billing
Dim lmStartCal As Long    'Starting date for standard billing
Dim lmEndCal As Long      'Ending date for standard billing
Dim smNowDate As String
Dim imWkNo As Integer   'Number of weeks when posting by week

Dim smFieldValues(0 To 19) As String    'Index zero temporarily used
Dim imComboBoxIndex As Integer
Dim imSettingValue As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imTaxDefined As Integer
Dim imBypassFocus As Integer
Dim imSetAll As Integer

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Const LBONE = 1

'Const CONTRACTINDEX = 1
'Const CASHTRADEINDEX = 2
'Const ADVTINDEX = 3
'Const VEHICLEINDEX = 4
'Const ONOSPOTSINDEX = 5
'Const OGROSSINDEX = 6
'Const ANOSPOTSINDEX = 7
'Const AGROSSINDEX = 8
'Const ABONUSINDEX = 9
'Const DNOSPOTSINDEX = 10
'Const DGROSSINDEX = 11
Dim imMaxIndex As Integer
Dim imCHECKINDEX As Integer
Dim imCONTRACTINDEX As Integer
Dim imLINEINDEX  As Integer
'Dim imCASHTRADEINDEX As Integer
'Dim imADVTINDEX As Integer
Dim imCashTradeIndex As Integer
Dim imAdvtIndex As Integer
Dim imVEHICLEINDEX As Integer
Dim imADVTVEHDAYPARTINDEX As Integer
Dim imLENGTHINDEX As Integer
Dim imPRIORNOSPOTSINDEX As Integer
Dim imONOSPOTSINDEX As Integer
Dim imOPRICEINDEX As Integer
Dim imOGROSSINDEX As Integer
Dim imINVOICENOINDEX As Integer
Dim imAPREVNOSPOTSINDEX As Integer
Dim imANOSPOTSINDEX As Integer
Dim imANEXTNOSPOTSINDEX As Integer
Dim imMISSEDINDEX As Integer
Dim imMGINDEX As Integer
Dim imBONUSINDEX As Integer
Dim imDGROSSINDEX As Integer
Dim imRECEIVEDINDEX As Integer
Dim imPOSTEDINDEX As Integer
Dim imANOSPOTSPRIORINDEX As Integer
Dim imANOSPOTSCURRINDEX As Integer
Dim imDIFFSPOTSINDEX As Integer
Dim imDIFFGROSSINDEX As Integer
Dim imBONUSPREVINDEX As Integer
Dim imBONUSCURRINDEX As Integer
Dim imNNOSPOTSINDEX As Integer
Dim imNNOBONUSINDEX As Integer

Dim imTOTALINDEX As Integer

Const WKLYDATESINDEX = 1
Const WKLYORDERNOINDEX = 2
Const WKLYAIRNOINDEX = 3
Const WKLYCARRIEDNOINDEX = 4
Const WKLYTOTALINDEX = 5




'*******************************************************
'*                                                     *
'*      Procedure Name:mAdvtPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Advertiser list box   *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAdvtPop()
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = cbcNames.ListIndex
    If ilIndex >= 0 Then
        slName = cbcNames.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(Copy, cbcAdvt, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(PostRep, cbcNames, tmAdvertiser(), smAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", PostRep
        On Error GoTo 0
'        cbcAdvt.AddItem "[New]", 0  'Force as first item on list
        If ilIndex >= 0 Then
            gFindMatch slName, 0, cbcNames
            If gLastFound(cbcNames) >= 0 Then
                cbcNames.ListIndex = gLastFound(cbcNames)
            Else
                cbcNames.ListIndex = -1
            End If
        Else
            cbcNames.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mShowInfo                       *
'*                                                     *
'*             Created:5/13/94       By:D. Hannifan    *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show Sdf information           *
'*                                                     *
'*******************************************************
Sub mShowInfo()
    Dim ilVef As Integer
    Dim slVehName As String
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilRet As Integer

    If (imButtonRow >= LBONE) And (imButtonRow <= UBound(smSave, 2)) Then
        plcInfo.Move pbcPostRep.Left, tmCtrls(imCONTRACTINDEX).fBoxY + (imButtonRow + 4) * (fgBoxGridH + 15)
        If (Trim$(smInfo(1, imButtonRow)) = "T") Or (igPostType = 3) Then
            plcInfo.Visible = False
        ElseIf (igPostType = 1) Or (igPostType = 3) Then
            If (imMarketIndex >= 0) Then
                slNameCode = tmVehicle(imMarketIndex).sKey    'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                slVehName = ""
                'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                '    If tgMVef(ilVef).iCode = ilVefCode Then
                    ilVef = gBinarySearchVef(ilVefCode)
                    If ilVef <> -1 Then
                        slStr = gFormatPhoneNo(tgMVef(ilVef).sPhone)
                        slVehName = "Vehicle: " & Trim$(tgMVef(ilVef).sName) & " Contact: " & Trim$(tgMVef(ilVef).sContact) & " " & slStr
                '        Exit For
                    End If
                'Next ilVef
                If slVehName <> "" Then
                    lacInfo(0).Caption = "Received Date: " & Trim$(smSave(11, imButtonRow)) & " Posted Date: " & Trim$(smSave(12, imButtonRow)) & " Gen A/R Date: " & Trim$(smInfo(13, imButtonRow)) & " Inv Date: " & Trim$(smInfo(12, imButtonRow))
                    lacInfo(1).Caption = slVehName
                    If (Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER Then
                        lacInfo(1).Caption = lacInfo(1).Caption & " Len:" & imSave(6, imButtonRow) & "s"
                    End If
                    '6/7/15: replaced acquisition from site override with Barter in system options
                    'If ((Asc(tgSpf.sOverrideOptions) And SPACQUISITION) = SPACQUISITION) Or ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
                    If ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
                        slStr = Trim$(str$(lmSave(4, imButtonRow)))
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                        lacInfo(1).Caption = lacInfo(1).Caption & " Acq $'s: " & slStr
                    End If
                    If Trim$(smSave(33, imButtonRow)) <> "" Then
                        lacInfo(2).Caption = "Line Comment: " & Trim$(smSave(33, imButtonRow))
                    Else
                        lacInfo(2).Caption = ""
                    End If
                    plcInfo.Visible = True
                End If
            End If
        Else
            slVehName = ""
            'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '    If tgMVef(ilVef).iCode = imSave(2, imButtonRow) Then
                ilVef = gBinarySearchVef(imSave(2, imButtonRow))
                If ilVef <> -1 Then
                    slStr = gFormatPhoneNo(tgMVef(ilVef).sPhone)
                    If igPostType = 4 Then
                        slVehName = Trim$(tgMVef(ilVef).sName)
                    Else
                        slVehName = Trim$(tgMVef(ilVef).sName) & " Contact: " & Trim$(tgMVef(ilVef).sContact) & " " & slStr
                    End If
            '        Exit For
                End If
            'Next ilVef
            If igPostType = 4 Then
                lacInfo(0).Caption = "Export Date: " & Trim$(smInfo(2, imButtonRow)) & " Import Date: " & Trim$(smInfo(3, imButtonRow)) & " Gen A/R Date: " & Trim$(smInfo(13, imButtonRow)) & " Inv Date: " & Trim$(smInfo(12, imButtonRow)) & " Ref Invoice #: " & Trim$(smInfo(6, imButtonRow))
                lacInfo(1).Caption = "Vehicle " & slVehName & " Remote Order Totals: Spots " & Trim$(smInfo(9, imButtonRow)) & " Gross " & Trim$(smInfo(10, imButtonRow)) & " Agency Comm % " & Trim$(smInfo(11, imButtonRow))
            Else
                lacInfo(0).Caption = "Received Date: " & Trim$(smSave(11, imButtonRow)) & " Posted Date: " & Trim$(smSave(12, imButtonRow)) & " Gen A/R Date: " & Trim$(smInfo(13, imButtonRow)) & " Inv Date: " & Trim$(smInfo(12, imButtonRow))
                lacInfo(1).Caption = "Vehicle " & slVehName
                If (Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER Then
                    lacInfo(1).Caption = lacInfo(1).Caption & " Len:" & imSave(6, imButtonRow) & "s"
                End If
                '6/7/15: replaced acquisition from site override with Barter in system options
                'If ((Asc(tgSpf.sOverrideOptions) And SPACQUISITION) = SPACQUISITION) Or ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
                If ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
                    slStr = Trim$(str$(lmSave(4, imButtonRow)))
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    lacInfo(1).Caption = lacInfo(1).Caption & " Acq $'s: " & slStr
                End If
            End If
            If Trim$(smSave(33, imButtonRow)) <> "" Then
                lacInfo(2).Caption = "Line Comment: " & Trim$(smSave(33, imButtonRow))
            Else
                lacInfo(2).Caption = ""
            End If
            plcInfo.Visible = True
        End If
    Else
        plcInfo.Visible = False
    End If
End Sub


'Help messages
Private Sub cbcInvDate_Change()
    Dim slStr As String     'Text entered

    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True    'Set change mode to avoid infinite loop
        slStr = cbcInvDate.Text
        If slStr <> "" Then
            gManLookAhead cbcInvDate, imBSMode, imComboBoxIndex
            If cbcInvDate.ListIndex >= 0 Then
                tmcClick.Enabled = False
                tmcClick.Interval = 2000    '2 seconds
                tmcClick.Enabled = True
            End If
        End If
        imInvDateIndex = cbcInvDate.ListIndex
        imChgMode = False
    End If
    cmcImport.Enabled = False
    Exit Sub

    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    Exit Sub
End Sub

Private Sub cbcInvDate_Click()
    imComboBoxIndex = cbcInvDate.ListIndex
    cbcInvDate_Change
End Sub

Private Sub cbcInvDate_GotFocus()
    Dim slSvText As String   'Save so list box can be reset

    If imTerminate Then
        Exit Sub
    End If
    tmcClick.Enabled = False
    mWklySetShow imWklyBoxNo
    imWklyBoxNo = -1
    imWklyRowNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
'    gSetIndexFromText cbcInvDate
    slSvText = cbcInvDate.Text
    If cbcInvDate.ListCount <= 1 Then
        If cbcInvDate.ListCount = 1 Then
            cbcInvDate.ListIndex = 0
        End If
        mClearCtrlFields 'Make sure all fields cleared
        mSetCommands
'        pbcHdSTab.SetFocus
        Exit Sub
    End If
'    gShowHelpMess tmChfHelp(), CHFCNTRSELECT
    gCtrlGotFocus ActiveControl
    If (slSvText = "") Then
        cbcInvDate.ListIndex = 0
        cbcInvDate_Change
    Else
        gFindMatch slSvText, 1, cbcInvDate
        If gLastFound(cbcInvDate) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcInvDate)) Or (ilSvIndex <> cbcInvDate.ListIndex) Then
            If (slSvText <> cbcInvDate.List(gLastFound(cbcInvDate))) Then
                cbcInvDate.ListIndex = gLastFound(cbcInvDate)
            End If
        Else
            cbcInvDate.ListIndex = 0
            mClearCtrlFields
            cbcInvDate_Change
        End If
    End If
    mSetCommands
End Sub

Sub cbcInvDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Sub cbcInvDate_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcInvDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cbcNames_Change()
    Dim slStr As String

    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True
        tmcClick.Enabled = False
        slStr = Trim$(cbcNames.Text)
        If slStr <> "" Then
            gManLookAhead cbcNames, imBSMode, imComboBoxIndex
            If cbcNames.ListIndex >= 0 Then
                tmcClick.Interval = 2000    '2 seconds
                tmcClick.Enabled = True
            End If
        End If
        imMarketIndex = cbcNames.ListIndex
        imChgMode = False
    End If
    cmcImport.Enabled = False
    Exit Sub
End Sub

Private Sub cbcNames_Click()
    imComboBoxIndex = cbcNames.ListIndex
    cbcNames_Change
End Sub

Private Sub cbcNames_GotFocus()
    If imTerminate Then
        Exit Sub
    End If
    mWklySetShow imWklyBoxNo
    imWklyBoxNo = -1
    imWklyRowNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus cbcNames
End Sub

Private Sub cbcNames_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcNames_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcNames.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub ckcAll_Click()
    Dim ilRow As Integer
    If Not imSetAll Then
        Exit Sub
    End If
    For ilRow = LBONE To UBound(smSave, 2) - 1 Step 1
        If (Trim$(smSave(8, ilRow)) <> "Y") And (Trim$(smInfo(1, ilRow)) <> "T") And (Trim$(smShow(imCONTRACTINDEX, ilRow)) <> "") Then
            If ckcAll.Value = vbChecked Then
                imSave(5, ilRow) = True
            Else
                imSave(5, ilRow) = False
            End If
        End If
    Next ilRow
    pbcPostRep.Cls
    pbcPostRep_Paint
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub cmcCancel_GotFocus()
    mWklySetShow imWklyBoxNo
    imWklyBoxNo = -1
    imWklyRowNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcClear_Click()
    Dim ilDelete As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slMsg As String
    Dim ilCRet As Integer
    Dim tlSbf As SBF

    For ilLoop = LBONE To UBound(smSave, 2) - 1 Step 1
        If (Trim$(smInfo(1, ilLoop)) = "I") Or (Trim$(smInfo(1, ilLoop)) = "F") Or (Trim$(smInfo(1, ilLoop)) = "S") Then
            If Trim$(smSave(8, ilLoop)) = "Y" Then
                ilRet = MsgBox("Previously Bill, Clear not allowed", vbOKOnly + vbExclamation, "Clear")
                Exit Sub
            End If
        End If
    Next ilLoop
    ilDelete = False
    ilRet = MsgBox("This will remove all Imports and any Posting of Aired Spots, Continue", vbYesNo + vbQuestion, "Delete")
    If ilRet = vbYes Then
        ilDelete = True
    End If
    If ilDelete Then
        Screen.MousePointer = vbHourglass
        mWklySetShow imWklyBoxNo
        mSetShow imBoxNo
        imChg = True
        lacFrame.Visible = False
        pbcArrow.Visible = False

        For ilLoop = LBONE To UBound(smSave, 2) - 1 Step 1
            If (Trim$(smInfo(1, ilLoop)) = "I") Or (Trim$(smInfo(1, ilLoop)) = "F") Or (Trim$(smInfo(1, ilLoop)) = "S") Then
                If lmSave(2, ilLoop) > 0 Then
                    slMsg = "Clear (btrGetEqual: Posting)"
                    tmSbfSrchKey1.lCode = lmSave(2, ilLoop)
                    ilRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    On Error GoTo mClearErr
                    gBtrvErrorMsg ilRet, slMsg, PostRep
                    On Error GoTo 0
                    Do
                        ilRet = btrDelete(hmSbf)
                        If ilRet = BTRV_ERR_CONFLICT Then
                            tmSbfSrchKey1.lCode = lmSave(2, ilLoop)
                            ilCRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        End If
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    slMsg = "Clear (btrDelete: Posting)"
                    On Error GoTo mClearErr
                    gBtrvErrorMsg ilRet, slMsg, PostRep
                    On Error GoTo 0
                End If
            End If
        Next ilLoop
        'Repopulate as if no import done.
        tmcClick_Timer
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
mClearErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    Exit Sub
End Sub

Private Sub cmcDone_Click()
    If (igPostType <> 5) And (igPostType <> 6) Then
        If Not imUpdateAllowed Then
            cmcCancel_Click
            Exit Sub
        End If
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub cmcDone_GotFocus()
    mWklySetShow imWklyBoxNo
    imWklyBoxNo = -1
    imWklyRowNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case imVEHICLEINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcImport_Click()
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim ilRet As Integer
    Dim slMsgFile As String

    ilRet = MsgBox("Current Airing values will be replaced, Continue?", vbYesNo + vbQuestion, "Warning")
    If ilRet = vbNo Then
        Exit Sub
    End If
    mBuildDate
    'gObtainYearMonthDayStr smStartStd, True, slFYear, slFMonth, slFDay
    gObtainYearMonthDayStr smEndStd, True, slFYear, slFMonth, slFDay
    If imInvDateIndex >= 0 Then
        slFMonth = Left$(cbcInvDate.List(imInvDateIndex), 3)
    Else
        slFMonth = "???"
    End If
    igBrowserType = 7  'Mask
    ''sgBrowseMaskFile = "F" & Right$(slFYear, 2) & slFMonth & slFDay & "?.I??"
    'sgBrowseMaskFile = "?" & right$(slFYear, 2) & slFMonth & slFDay & "?.I??"
    sgBrowseMaskFile = slFMonth & right$(slFYear, 2) & "In?.??"
    sgBrowserTitle = "Import for " & cbcNames.List(imMarketIndex)
    Browser.Show vbModal
    sgBrowserTitle = ""
    If igBrowserReturn = 1 Then
        Screen.MousePointer = vbHourglass
        slMsgFile = sgBrowserFile
        'If InStr(slMsgFile, ":") = 0 Then
        If (InStr(slMsgFile, ":") = 0) And (Left$(slMsgFile, 2) <> "\\") Then
            slMsgFile = sgImportPath & slMsgFile
        End If
        ilRet = mOpenMsgFile(slMsgFile)
        If Not ilRet Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Print #hmMsg, "Import " & sgBrowserFile & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        'Remove Aired count if not previously defined and not imported
        mRemoveAirCount True
        pbcPostRep.Cls
        ilRet = mReadImportFile(sgBrowserFile)
        If ilRet Then
            Print #hmMsg, "Import Finish Successfully"
            Close #hmMsg
        Else
            Print #hmMsg, "** Import Failed or Terminated " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            MsgBox "See " & slMsgFile & " for errors related to Rejected Records"
        End If
        'pbcPostRep.Cls
        'Compute totals
        mRecomputeTotals
        vbcPostRep.Min = LBONE  'LBound(smSave, 2)
        If UBound(smSave, 2) <= vbcPostRep.LargeChange Then
            vbcPostRep.Max = LBONE  'LBound(smSave, 2)
        Else
            vbcPostRep.Max = UBound(smSave, 2) - vbcPostRep.LargeChange
        End If
        If vbcPostRep.Value = vbcPostRep.Min Then
            pbcPostRep_Paint
        Else
            vbcPostRep.Value = vbcPostRep.Min
        End If
        cmcClear.Enabled = False
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmcImport_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub cmcImport_GotFocus()
    mWklySetShow imWklyBoxNo
    imWklyBoxNo = -1
    imWklyRowNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUpdate_Click()
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
    mEnableBox imBoxNo
    If igPostType = 4 Then
        cmcClear.Enabled = True
    End If
    mSetCommands
End Sub
Private Sub cmcUpdate_GotFocus()
    mWklySetShow imWklyBoxNo
    imWklyBoxNo = -1
    imWklyRowNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcGross_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcGross_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(ActiveControl.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(ActiveControl.Text, ".")    'Disallow multi-decimal points
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
    slStr = edcGross.Text
    slStr = Left$(slStr, edcGross.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcGross.SelStart - edcGross.SelLength)
    If gCompNumberStr(slStr, "9999999.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_Change()
    Select Case imBoxNo
        Case imVEHICLEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
        Case imOPRICEINDEX
    End Select
    imLbcArrowSetting = False
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case imVEHICLEINDEX
            If lbcVehicle.ListCount = 1 Then
                lbcVehicle.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case imOPRICEINDEX
    End Select
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    Dim slComp As String
    Select Case imBoxNo
        Case imVEHICLEINDEX
            If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
                If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
                    imBSMode = True 'Force deletion of character prior to selected text
                End If
            End If
        Case imOPRICEINDEX
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
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            slComp = "9999999.99"
            If gCompNumberStr(slStr, slComp) > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case imINVOICENOINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, "999999999") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case Else
            If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
                If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
                    imBSMode = True 'Force deletion of character prior to selected text
                End If
            End If
    End Select

End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case imVEHICLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
        End Select
        imDoubleClickName = False
    End If
End Sub
Private Sub edcNoSpots_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcNoSpots_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcNoSpots.Text
    slStr = Left$(slStr, edcNoSpots.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcNoSpots.SelStart - edcNoSpots.SelLength)
    If gCompNumberStr(slStr, "9999") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcWkly_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcWkly_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcWkly.Text
    slStr = Left$(slStr, edcWkly.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcWkly.SelStart - edcWkly.SelLength)
    If gCompNumberStr(slStr, "9999") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
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
    Me.KeyPreview = True
    If (igWinStatus(INVOICESJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        If tgUrf(0).iSlfCode > 0 Then
            imUpdateAllowed = False
        Else
            imUpdateAllowed = True
        End If
    End If
    gShowBranner imUpdateAllowed
    Me.ZOrder 0 'Send to front
    Me.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If (cbcNames.Enabled) And (imBoxNo > 0) Then
            plcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            plcSelect.Enabled = True
        End If
    Else
        If KeyCode = KEYINSERT Then    'Insert Row
            imcInsert_Click
        ElseIf KeyCode = KEYDELETE Then
            imcTrash_Click
        End If

    End If
End Sub

Private Sub Form_Load()
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    If Not igManUnload Then
        mWklySetShow imWklyBoxNo
        imWklyBoxNo = -1
        imWklyRowNo = -1
        mSetShow imBoxNo
        imBoxNo = -1
        imRowNo = -1
        pbcArrow.Visible = False
        lacFrame.Visible = False
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            If imBoxNo <> -1 Then
                mEnableBox imBoxNo
            End If
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    Me.KeyPreview = False

    Erase tmMktCode
    Erase imMktVefCode

    Erase tmContract
    Erase tmChfAdvtExt
    Erase tmVehicle

    Erase tmAdvertiser
    smAdvertiserTag = ""

    Erase tmInvVehicle
    smInvVehicleTag = ""
    Erase lmSbfDel

    Erase tmNetNames

    ilRet = btrClose(hmCxf)
    btrDestroy hmCxf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmRwf)
    btrDestroy hmRwf
    ilRet = btrClose(hmNrf)
    btrDestroy hmNrf
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf

    Erase smSave
    Erase imSave
    Erase lmSave
    Erase smShow
    igJobShowing(POSTLOGSJOB) = False
    Set PostRep = Nothing   'Remove data segment
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub imcInsert_Click()
    Dim ilNewRow As Integer
    Dim ilIndex As Integer
    Dim ilCol As Integer

    mWklySetShow imWklyBoxNo
    imWklyBoxNo = -1
    imWklyRowNo = -1
    mSetShow imBoxNo
    pbcArrow.Visible = False
    lacFrame.Visible = False
    If (imRowNo >= LBONE) And (imRowNo <= UBound(smSave, 2)) Then
        If Trim$(smInfo(1, imRowNo)) <> "T" Then
            ilNewRow = imRowNo + 1
            'Move rows down and duplicate current row
            For ilIndex = UBound(smSave, 2) To ilNewRow Step -1
                For ilCol = LBONE To UBound(smSave, 1) Step 1
                    smSave(ilCol, ilIndex) = smSave(ilCol, ilIndex - 1)
                Next ilCol
                For ilCol = LBONE To UBound(imSave, 1) Step 1
                    imSave(ilCol, ilIndex) = imSave(ilCol, ilIndex - 1)
                Next ilCol
                For ilCol = LBONE To UBound(lmSave, 1) Step 1
                    lmSave(ilCol, ilIndex) = lmSave(ilCol, ilIndex - 1)
                Next ilCol
                For ilCol = LBONE To UBound(smShow, 1) Step 1
                    smShow(ilCol, ilIndex) = smShow(ilCol, ilIndex - 1)
                Next ilCol
                For ilCol = LBONE To UBound(smInfo, 1) Step 1
                    smInfo(ilCol, ilIndex) = smInfo(ilCol, ilIndex - 1)
                Next ilCol
            Next ilIndex
            ReDim Preserve smSave(0 To 33, 0 To UBound(smSave, 2) + 1) As String
            ReDim Preserve imSave(0 To 10, 0 To UBound(imSave, 2) + 1) As Integer
            ReDim Preserve lmSave(0 To 5, 0 To UBound(lmSave, 2) + 1) As Long
            ReDim Preserve smShow(0 To 10, 0 To UBound(smShow, 2) + 1) As String * 40
            ReDim Preserve smInfo(0 To 13, 0 To UBound(smInfo, 2) + 1) As String * 12
            smSave(3, ilNewRow) = ""    'Ordered # spots
            smSave(4, ilNewRow) = ""    'Ordered Gross
            smSave(5, ilNewRow) = "0"   'Aired # spots
            smSave(6, ilNewRow) = "0.00"    'Aired gross
            smSave(7, ilNewRow) = ""    'Bonus # Spots
            smSave(8, ilNewRow) = "N"   'Billed
            smSave(9, ilNewRow) = ""    'Calendar Carry
            smSave(10, ilNewRow) = ""   'Missed carry
            smSave(11, ilNewRow) = ""   'Received date
            smSave(12, ilNewRow) = ""   'Posted Date
            smSave(13, ilNewRow) = ""   'MGs
            smSave(14, ilNewRow) = ""   'Missed
            smSave(15, ilNewRow) = ""   'Initial Aired Spots
            smSave(16, ilNewRow) = "0"  'Next billing number of spots (calendar billing)
            smSave(17, ilNewRow) = "N"   'Week 1 date
            smSave(18, ilNewRow) = ""   'Aired weekly Spots
            smSave(19, ilNewRow) = ""   'Aired weekly Spots
            smSave(20, ilNewRow) = ""   'Aired weekly Spots
            smSave(21, ilNewRow) = ""   'Aired weekly Spots
            smSave(22, ilNewRow) = ""   'Aired weekly Spots
            smSave(23, ilNewRow) = ""   'Carried weekly Spots
            smSave(24, ilNewRow) = ""   'Carried weekly Spots
            smSave(25, ilNewRow) = ""   'Carried weekly Spots
            smSave(26, ilNewRow) = ""   'Carried weekly Spots
            smSave(27, ilNewRow) = ""   'Carried weekly Spots
            smSave(28, ilNewRow) = ""   'Order weekly Spots
            smSave(29, ilNewRow) = ""   'Order weekly Spots
            smSave(30, ilNewRow) = ""   'Order weekly Spots
            smSave(31, ilNewRow) = ""   'Order weekly Spots
            smSave(32, ilNewRow) = ""   'Order weekly Spots
            smSave(33, ilNewRow) = ""   'Line Comment
            gSetShow pbcPostRep, smSave(3, ilNewRow), tmCtrls(imONOSPOTSINDEX)
            smShow(imONOSPOTSINDEX, ilNewRow) = tmCtrls(imONOSPOTSINDEX).sShow
            gSetShow pbcPostRep, smSave(4, ilNewRow), tmCtrls(imOGROSSINDEX)
            smShow(imOGROSSINDEX, ilNewRow) = tmCtrls(imOGROSSINDEX).sShow
            gSetShow pbcPostRep, smSave(5, ilNewRow), tmCtrls(imANOSPOTSINDEX)
            smShow(imANOSPOTSINDEX, ilNewRow) = tmCtrls(imANOSPOTSINDEX).sShow
            If imANEXTNOSPOTSINDEX > 0 Then
                gSetShow pbcPostRep, smSave(16, ilNewRow), tmCtrls(imANEXTNOSPOTSINDEX)
                smShow(imANEXTNOSPOTSINDEX, ilNewRow) = tmCtrls(imANEXTNOSPOTSINDEX).sShow
            End If
            'gSetShow pbcPostRep, smSave(6, ilNewRow), tmCtrls(imAGROSSINDEX)
            'smShow(imAGROSSINDEX, ilNewRow) = tmCtrls(imAGROSSINDEX).sShow
            'gSetShow pbcPostRep, smSave(7, ilNewRow), tmCtrls(imABONUSINDEX)
            'smShow(imABONUSINDEX, ilNewRow) = tmCtrls(imABONUSINDEX).sShow
            'Retain advertiser and vehicle if posting by advertiser; retain advertiser if posting by advertiser
            'imSave(1, ilNewRow) and imSave(2, ilNewRow)
            If igPostType = 2 Then
                imSave(2, ilNewRow) = -1
                imSave(3, ilNewRow) = -1
            End If
            imSave(4, ilNewRow) = True  'Valid vehicle
            imSave(5, ilNewRow) = False 'Checked

            imSave(6, ilNewRow) = 0     'Spot Length

            'Retain chfCode (lmSave(1, ilNewRow) and price lmSave(3, ilNewRow)
            lmSave(2, ilNewRow) = 0
            lmSave(5, ilNewRow) = 0
            smInfo(1, ilNewRow) = "S"   'Insert
            smInfo(2, ilNewRow) = ""    'Export date
            smInfo(3, ilNewRow) = ""    'Import Date
            smInfo(4, ilNewRow) = ""    'sbf date
            smInfo(5, ilNewRow) = ""   'Combine ID
            smInfo(6, ilNewRow) = ""   'Ref Inv #
            smInfo(7, ilNewRow) = "0.00"    'Tax 1
            smInfo(8, ilNewRow) = "0.00"    'Tax 2
            smInfo(9, ilNewRow) = ""    'Ordered Spots
            smInfo(10, ilNewRow) = ""   'Ordered Gross
            smInfo(11, ilNewRow) = ""   'Comm Pct
            smInfo(12, ilNewRow) = ""   'Print Rep Invoice date
            smInfo(13, ilNewRow) = ""    'Gen AR date
            pbcPostRep.Cls
            vbcPostRep.Min = LBONE  'LBound(smSave, 2)
            If UBound(smSave, 2) <= vbcPostRep.LargeChange Then
                vbcPostRep.Max = LBONE  'LBound(smSave, 2)
            Else
                vbcPostRep.Max = UBound(smSave, 2) - vbcPostRep.LargeChange
            End If
            If vbcPostRep.Value = vbcPostRep.Min Then
                pbcPostRep_Paint
            Else
                vbcPostRep.Value = vbcPostRep.Min
            End If
            If igPostType = 2 Then
                imBoxNo = imVEHICLEINDEX
            Else
                imBoxNo = imOPRICEINDEX
            End If
            imRowNo = ilNewRow
            If imRowNo <= vbcPostRep.LargeChange + 1 Then
                vbcPostRep.Value = vbcPostRep.Min
            Else
                vbcPostRep.Value = imRowNo - vbcPostRep.LargeChange
            End If
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    cmcDone.SetFocus
End Sub

Private Sub imcInsert_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imChg = True
    mSetCommands
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

Sub imcTrash_Click()
    Dim ilRet As Integer
    Dim ilDelete As Integer
    Dim ilIndex As Integer
    Dim ilCol As Integer

    If (imRowNo < LBONE) Or (imRowNo > UBound(smSave, 2)) Then
        Exit Sub
    End If
    If InStr(1, smShow(imTOTALINDEX, imRowNo), "Total:", 1) > 0 Then
        Beep
        Exit Sub
    End If
    ilDelete = False
    If (Trim$(smInfo(1, imRowNo)) = "S") Then
        ilDelete = True
    ElseIf (Val(smSave(3, imRowNo)) = 0) And (gStrDecToLong(smSave(4, imRowNo), 2) = 0) Then
        ilRet = MsgBox("Ok to Delete Row", vbYesNo + vbQuestion, "Trash")
        If ilRet = vbYes Then
            ilDelete = True
        End If
    End If
    If ilDelete Then
        mWklySetShow imWklyBoxNo
        imWklyBoxNo = -1
        imWklyRowNo = -1
        mSetShow imBoxNo
        imChg = True
        lacFrame.Visible = False
        pbcArrow.Visible = False
        If lmSave(2, imRowNo) > 0 Then
            lmSbfDel(UBound(lmSbfDel)) = lmSave(2, imRowNo)
            ReDim Preserve lmSbfDel(0 To UBound(lmSbfDel) + 1) As Long
        End If
        Screen.MousePointer = vbHourglass
        'Move rows up
        For ilIndex = imRowNo To UBound(smSave, 2) - 1 Step 1
            For ilCol = LBONE To UBound(smSave, 1) Step 1
                smSave(ilCol, ilIndex) = smSave(ilCol, ilIndex + 1)
            Next ilCol
            For ilCol = LBONE To UBound(imSave, 1) Step 1
                imSave(ilCol, ilIndex) = imSave(ilCol, ilIndex + 1)
            Next ilCol
            For ilCol = LBONE To UBound(lmSave, 1) Step 1
                lmSave(ilCol, ilIndex) = lmSave(ilCol, ilIndex + 1)
            Next ilCol
            For ilCol = LBONE To UBound(smShow, 1) Step 1
                smShow(ilCol, ilIndex) = smShow(ilCol, ilIndex + 1)
            Next ilCol
            For ilCol = LBONE To UBound(smInfo, 1) Step 1
                smInfo(ilCol, ilIndex) = smInfo(ilCol, ilIndex + 1)
            Next ilCol
        Next ilIndex
        ReDim Preserve smSave(0 To 33, 0 To UBound(smSave, 2) - 1) As String
        ReDim Preserve imSave(0 To 10, 0 To UBound(imSave, 2) - 1) As Integer
        ReDim Preserve lmSave(0 To 5, 0 To UBound(lmSave, 2) - 1) As Long
        ReDim Preserve smShow(0 To 10, 0 To UBound(smShow, 2) - 1) As String * 40
        ReDim Preserve smInfo(0 To 13, 0 To UBound(smInfo, 2) - 1) As String * 12
        pbcPostRep.Cls
        mRecomputeTotals
        vbcPostRep.Min = LBONE  'LBound(smSave, 2)
        If UBound(smSave, 2) <= vbcPostRep.LargeChange Then
            vbcPostRep.Max = LBONE  'LBound(smSave, 2)
        Else
            vbcPostRep.Max = UBound(smSave, 2) - vbcPostRep.LargeChange
        End If
        If vbcPostRep.Value = vbcPostRep.Min Then
            pbcPostRep_Paint
        Else
            vbcPostRep.Value = vbcPostRep.Min
        End If
        mSetCommands
        If InStr(1, smShow(imTOTALINDEX, imRowNo), "Total:", 1) > 0 Then
            If imRowNo + 1 >= UBound(smSave, 2) Then
                imRowNo = -1
                If (igPostType = 5) Or (igPostType = 6) Then
                    cmcDone.SetFocus
                Else
                    cmcCancel.SetFocus
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            imRowNo = imRowNo + 1
        End If
        If imRowNo <= vbcPostRep.LargeChange + 1 Then
            vbcPostRep.Value = vbcPostRep.Min
        Else
            vbcPostRep.Value = imRowNo - vbcPostRep.LargeChange
        End If
        imBoxNo = 0
        lacFrame.Move 0, tmCtrls(imCONTRACTINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15) - 30
        lacFrame.Visible = True
        pbcArrow.Move pbcArrow.Left, plcPostRep.Top + tmCtrls(imCONTRACTINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15) + 45
        pbcArrow.Visible = True
        pbcArrow.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub lacScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub lacScreen_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub

Private Sub lacTotals_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub lbcVehicle_Click()
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mAddGrandTotalLine              *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute grans totals           *
'*                                                     *
'*******************************************************
Sub mAddGrandTotalLine()
    Dim ilRowNo As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilCol As Integer

    'Add Total line
    ilRowNo = UBound(smSave, 2)
    slStr = ""
    smSave(1, ilRowNo) = ""
    smSave(2, ilRowNo) = ""
    smSave(3, ilRowNo) = ""
    smSave(4, ilRowNo) = ""
    smSave(5, ilRowNo) = ""
    smSave(6, ilRowNo) = ""
    smSave(7, ilRowNo) = ""
    smSave(8, ilRowNo) = ""
    smSave(9, ilRowNo) = ""
    smSave(10, ilRowNo) = ""
    smSave(16, ilRowNo) = ""
    smSave(17, ilRowNo) = ""
    smSave(33, ilRowNo) = ""
    gSetShow pbcPostRep, slStr, tmCtrls(imONOSPOTSINDEX)
    smShow(imONOSPOTSINDEX, ilRowNo) = tmCtrls(imONOSPOTSINDEX).sShow
    gSetShow pbcPostRep, slStr, tmCtrls(imOGROSSINDEX)
    smShow(imOGROSSINDEX, ilRowNo) = tmCtrls(imOGROSSINDEX).sShow
    If imANOSPOTSINDEX > 0 Then
        gSetShow pbcPostRep, slStr, tmCtrls(imANOSPOTSINDEX)
        smShow(imANOSPOTSINDEX, ilRowNo) = tmCtrls(imANOSPOTSINDEX).sShow
    End If
    If imANEXTNOSPOTSINDEX > 0 Then
        gSetShow pbcPostRep, slStr, tmCtrls(imANEXTNOSPOTSINDEX)
        smShow(imANEXTNOSPOTSINDEX, ilRowNo) = tmCtrls(imANEXTNOSPOTSINDEX).sShow
    End If
    'gSetShow pbcPostRep, slStr, tmCtrls(imAGROSSINDEX)
    'smShow(imAGROSSINDEX, ilRowNo) = tmCtrls(imAGROSSINDEX).sShow
    imSave(2, ilRowNo) = 0
    imSave(3, ilRowNo) = -1
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, ilRowNo) = ""
    Next ilLoop
    slStr = "Grand Total:"
    gSetShow pbcPostRep, slStr, tmCtrls(imTOTALINDEX)
    smShow(imTOTALINDEX, ilRowNo) = tmCtrls(imTOTALINDEX).sShow
    For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
        smInfo(ilCol, ilRowNo) = ""
    Next ilCol
    smInfo(1, ilRowNo) = "T"
    ReDim Preserve smSave(0 To 33, 0 To ilRowNo + 1) As String
    ReDim Preserve imSave(0 To 10, 0 To ilRowNo + 1) As Integer
    ReDim Preserve lmSave(0 To 5, 0 To ilRowNo + 1) As Long
    ReDim Preserve smShow(0 To 10, 0 To ilRowNo + 1) As String * 40
    ReDim Preserve smInfo(0 To 13, 0 To ilRowNo + 1) As String * 12
    mGrandTotal
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mBuildDate                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Converted selected date        *
'*                                                     *
'*******************************************************
Sub mBuildDate()
    Dim slName As String
    Dim slMonth As String
    Dim slYear As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim llDate As Long

    smStartCal = ""
    smEndCal = ""
    lmStartCal = 0
    lmEndCal = 0
    smStartStd = ""
    smEndStd = ""
    lmStartStd = 0
    lmEndStd = 0
    lacCalDates.Caption = ""
    If imInvDateIndex < 0 Then
        Exit Sub
    End If
    'Build Dates
    slName = cbcInvDate.List(imInvDateIndex)
    ilRet = gParseItem(slName, 1, ",", slMonth)
    ilRet = gParseItem(slName, 2, ",", slYear)
    Select Case UCase$(slMonth)
        Case "JAN"
            slDate = "1/15/" & slYear
        Case "FEB"
            slDate = "2/15/" & slYear
        Case "MAR"
            slDate = "3/15/" & slYear
        Case "APR"
            slDate = "4/15/" & slYear
        Case "MAY"
            slDate = "5/15/" & slYear
        Case "June"
            slDate = "6/15/" & slYear
        Case "July"
            slDate = "7/15/" & slYear
        Case "AUG"
            slDate = "8/15/" & slYear
        Case "SEPT"
            slDate = "9/15/" & slYear
        Case "OCT"
            slDate = "10/15/" & slYear
        Case "NOV"
            slDate = "11/15/" & slYear
        Case "DEC"
            slDate = "12/15/" & slYear
    End Select
    smStartCal = gObtainStartCal(slDate)
    smEndCal = gObtainEndCal(smStartCal)
    smStartStd = gObtainStartStd(slDate)
    smEndStd = gObtainEndStd(smStartStd)
    lmStartStd = gDateValue(smStartStd)
    lmEndStd = gDateValue(smEndStd)
    lmStartCal = gDateValue(smStartCal)
    lmEndCal = gDateValue(smEndCal)
    mShowCalDates
    If tgSpf.sPostCalAff = "W" Then
        imWkNo = 1
        For llDate = gDateValue(smStartStd) To gDateValue(smEndStd) Step 7
            smWklyDates(imWkNo) = Format$(llDate, "m/d/yy") & "-" & Format$(llDate + 6, "m/d/yy")
            imWkNo = imWkNo + 1
        Next llDate
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
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
'       ilOnlyAddr (I)- Clear only fields after address
'
    Dim ilLoop As Integer
    Dim ilCol As Integer

    lbcVehicle.ListIndex = -1
    edcGross.Text = ""
    edcNoSpots.Text = ""
    For ilLoop = imLBCtrls To imMaxIndex Step 1
        tmCtrls(ilLoop).sShow = ""
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    ReDim smSave(0 To 33, 0 To 1) As String
    ReDim imSave(0 To 10, 0 To 1) As Integer
    ReDim lmSave(0 To 5, 0 To 1) As Long
    ReDim smShow(0 To 10, 0 To 1) As String * 40
    ReDim smInfo(0 To 13, 0 To 1) As String * 12
    ReDim lmSbfDel(0 To 0) As Long
    For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilCol, 1) = ""
    Next ilCol
    For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
        smInfo(ilCol, 1) = ""
    Next ilCol
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcMouseDown = False
    imChgMode = False
    imBSMode = False
    imChg = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDatePop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Date list box         *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mDatePop()
    Dim llLastStd As Long
    Dim llLastCal As Long
    Dim llEarliestDate As Long
    Dim slDate As String
    Dim llDate As Long
    'Dim ilMonth As Integer
    'Dim ilYear As Integer
    Dim slName As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slStr As String

    llEarliestDate = -1
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slDate
    llLastStd = gDateValue(slDate)
    gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), slDate
    llLastCal = gDateValue(slDate)
    'ilRet = gPopUserVehicleBox(PostRep, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcVehicle, tmInvVehicle(), smInvVehicleTag)
    'For ilLoop = 0 To UBound(tmInvVehicle) - 1 Step 1 'Traffic!lbcVehicle.ListCount - 1 To 0 Step -1
    '    slNameCode = tmInvVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
    '    ilRet = gParseItem(slNameCode, 1, "\", slName)    'Get application name
    '    ilRet = gParseItem(slName, 3, "|", slName)    'Get application name
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
    '    ilVefCode = Val(slCode)
    '    tmLcfSrchKey.sType = "O"  ' On Air code
    '    tmLcfSrchKey.sStatus = "C"  ' Current
    '    tmLcfSrchKey.iVefCode = ilVefCode
    '    slDate = Format$("1/1/95", "m/d/yy")
    '    gPackDate slDate, tmLcfSrchKey.iLogDate(0), tmLcfSrchKey.iLogDate(1)
    '    tmLcfSrchKey.iSeqNo = 0
    '    ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    '    If (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = ilVefCode) And (tmLcf.sStatus = "C") And (tmLcf.sType = "O") Then
    '        gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llDate
    '        If llEarliestDate = -1 Then
    '            llEarliestDate = llDate
    '        ElseIf llDate < llEarliestDate Then
    '            llEarliestDate = llDate
    '        End If
    '    End If
    'Next ilLoop
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slDate
    'If (slDate <> "") And (llEarliestDate > 0) Then
    If (slDate <> "") Then
        llDate = gDateValue(slDate)
        'Could back up earliestDate to start of SBF or at least to first unbilled date
        'For now just back it up 6 months
        imSbfRecLen = Len(tmSbf)
        tmSbfSrchKey2.sTranType = "T"
        gPackDate "1/1/2000", tmSbfSrchKey2.iDate(0), tmSbfSrchKey2.iDate(1)
        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) And (tmSbf.sTranType = "T") Then
            gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llEarliestDate
        Else
            llEarliestDate = llDate - 90
        End If
        If gDateValue(Format$(gNow(), "m/d/yy")) > llDate Then
            llDate = gDateValue(Format$(gNow(), "m/d/yy"))
        End If
        slDate = Format$(llDate, "m/d/yy")
        slDate = gObtainStartStd(slDate)
        llDate = gDateValue(slDate) - 1
        Do While llDate >= llEarliestDate
            slDate = Format$(llDate, "m/d/yy")
            slName = gMonthYearFormat(slDate)
            cbcInvDate.AddItem slName
            slDate = gObtainStartStd(slDate)
            llDate = gDateValue(slDate) - 1
        Loop
    End If
    ''Check previous month
    'slDate = Format$(gDateValue(Format$(gNow(), "m/d/yy")) - 20, "m/d/yy")
    'slName = gMonthYearFormat(slDate)
    'ilFound = False
    'For ilLoop = 0 To cbcInvDate.ListCount - 1 Step 1
    '    slStr = cbcInvDate.List(ilLoop)
    '    If StrComp(slName, slStr, 1) = 0 Then
    '        ilFound = True
    '        Exit For
    '    End If
    'Next ilLoop
    'If Not ilFound Then
    '    cbcInvDate.AddItem slName, 0
    'End If
    'slDate = Format$(gNow(), "m/d/yy")
    'slName = gMonthYearFormat(slDate)
    'ilFound = False
    'For ilLoop = 0 To cbcInvDate.ListCount - 1 Step 1
    '    slStr = cbcInvDate.List(ilLoop)
    '    If StrComp(slName, slStr, 1) = 0 Then
    '        ilFound = True
    '        Exit For
    '    End If
    'Next ilLoop
    'If Not ilFound Then
    '    cbcInvDate.AddItem slName, 0
    'End If
    'Set to next month
    slDate = Format$(llLastStd + 15, "m/d/yy")
    slName = gMonthYearFormat(slDate)
    For ilLoop = 0 To cbcInvDate.ListCount - 1 Step 1
        slStr = cbcInvDate.List(ilLoop)
        If StrComp(slName, slStr, 1) = 0 Then
            cbcInvDate.ListIndex = ilLoop
            Exit For
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxIndex) Then
        Exit Sub
    End If

    If (imRowNo < vbcPostRep.Value) Or (imRowNo >= vbcPostRep.Value + vbcPostRep.LargeChange + 1) Then
        mSetShow ilBoxNo
        pbcArrow.Visible = False
        lacFrame.Visible = False
        Exit Sub
    End If
    lacFrame.Move 0, tmCtrls(imCONTRACTINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15) - 30
    lacFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcPostRep.Top + tmCtrls(imCONTRACTINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    plcInfo.Visible = False
    Select Case ilBoxNo 'Branch on box type (control)
        Case imVEHICLEINDEX 'Vehicle
            'mVehPop
'            gShowHelpMess tmSbfHelp(), SBFVEHICLE
            lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 8)
            edcDropDown.Width = tmCtrls(imVEHICLEINDEX).fBoxW
            If tgSpf.iVehLen <= 40 Then
                edcDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveTableCtrl pbcPostRep, edcDropDown, tmCtrls(imVEHICLEINDEX).fBoxX, tmCtrls(imVEHICLEINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            imChgMode = True
            If imSave(3, imRowNo) >= 0 Then
                lbcVehicle.ListIndex = imSave(3, imRowNo)
                imComboBoxIndex = lbcVehicle.ListIndex
                edcDropDown.Text = lbcVehicle.List(imSave(3, imRowNo))
            Else
                lbcVehicle.ListIndex = 0
                imComboBoxIndex = lbcVehicle.ListIndex
                edcDropDown.Text = lbcVehicle.List(0)
            End If
            imChgMode = False
            If imRowNo - vbcPostRep.Value <= vbcPostRep.LargeChange \ 2 Then
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case imOPRICEINDEX
            edcDropDown.Width = tmCtrls(imOPRICEINDEX).fBoxW
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcPostRep, edcDropDown, tmCtrls(imOPRICEINDEX).fBoxX, tmCtrls(imOPRICEINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15)
            edcDropDown.Text = gLongToStrDec(lmSave(3, imRowNo), 2)
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case imINVOICENOINDEX
            edcDropDown.Width = tmCtrls(imINVOICENOINDEX).fBoxW
            edcDropDown.MaxLength = 9   '10
            gMoveTableCtrl pbcPostRep, edcDropDown, tmCtrls(imINVOICENOINDEX).fBoxX, tmCtrls(imINVOICENOINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15)
            If lmSave(5, imRowNo) > 0 Then
                edcDropDown.Text = lmSave(5, imRowNo)
            Else
                edcDropDown.Text = ""
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case imANOSPOTSINDEX
'            gShowHelpMess tmSbfHelp(), SBFITEMDESC
            If tgSpf.sPostCalAff <> "W" Then
                edcNoSpots.Width = tmCtrls(imANOSPOTSINDEX).fBoxW
                gMoveTableCtrl pbcPostRep, edcNoSpots, tmCtrls(imANOSPOTSINDEX).fBoxX, tmCtrls(imANOSPOTSINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15)
                edcNoSpots.Text = smSave(5, imRowNo)
                edcNoSpots.Visible = True  'Set visibility
                edcNoSpots.SetFocus
            Else
                plcWeekly.Height = pbcWeekly.Height + 345
                plcWeekly.Width = pbcWeekly.Width + 300
                plcWeekly.Move pbcPostRep.Left + tmCtrls(imANOSPOTSINDEX).fBoxX, pbcPostRep.Top + tmCtrls(imANOSPOTSINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15)
                plcWeekly.Move vbcPostRep.Left - plcWeekly.Width, pbcPostRep.Top + tmCtrls(imANOSPOTSINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15)
                If plcWeekly.Top + plcWeekly.Height > cmcUpdate.Top + cmcUpdate.Height Then
                    plcWeekly.Move plcWeekly.Left, pbcPostRep.Top + tmLnCtrls(imANOSPOTSINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15) - plcWeekly.Height + (fgBoxGridH)
                End If
                pbcWeekly.Move plcWeekly.Left + 150, plcWeekly.Top + 225
                plcWeekly.Visible = True
                pbcWeekly.Visible = True
                imWklyRowNo = 1
                imWklyBoxNo = WKLYAIRNOINDEX
                mWklyEnableBox imWklyBoxNo
            End If
        Case imANEXTNOSPOTSINDEX
            edcNoSpots.Width = tmCtrls(imANEXTNOSPOTSINDEX).fBoxW
            gMoveTableCtrl pbcPostRep, edcNoSpots, tmCtrls(imANEXTNOSPOTSINDEX).fBoxX, tmCtrls(imANEXTNOSPOTSINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15)
            edcNoSpots.Text = smSave(16, imRowNo)
            edcNoSpots.Visible = True  'Set visibility
            edcNoSpots.SetFocus
        'Case imAGROSSINDEX
'            gShowHelpMess tmSbfHelp(), SBFITEMCOST
        '    edcGross.Width = tmCtrls(imAGROSSINDEX).fBoxW
        '    gMoveTableCtrl pbcPostRep, edcGross, tmCtrls(imAGROSSINDEX).fBoxX, tmCtrls(imAGROSSINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15)
        '    edcGross.Text = smSave(6, imRowNo)
        '    edcGross.Visible = True  'Set visibility
        '    edcGross.SetFocus
        'Case imABONUSINDEX
        '    edcNoSpots.Width = tmCtrls(imABONUSINDEX).fBoxW
        '    gMoveTableCtrl pbcPostRep, edcNoSpots, tmCtrls(imABONUSINDEX).fBoxX, tmCtrls(imABONUSINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15)
        '    edcNoSpots.Text = smSave(7, imRowNo)
        '    edcNoSpots.Visible = True  'Set visibility
        '    edcNoSpots.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetCntr                        *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Contracts for market       *
'*                                                     *
'*******************************************************
Sub mGetCntr()
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStatus As String
    Dim slCntrType As String
    Dim ilHOType As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVeh As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim llChfCode As Long
    Dim ilVefCode As Integer
    Dim ilFound As Integer
    Dim ilChf As Integer
    Dim slKey As String
    Dim slStr As String
    Dim ilAdf As Integer
    Dim ilAgf As Integer

    'Moved to tmcClick
    'mBuildDate
    'ReDim imMktVefCode(0 To 0) As Integer
    'If imMarketIndex >= 0 Then
    '    slNameCode = tmMktCode(imMarketIndex).sKey    'lbcMster.List(ilLoop)
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '    imMktCode = Val(slCode)
    '    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
    '        If tgMVef(ilLoop).iMnfVehGp3Mkt = imMktCode Then
    '            imMktVefCode(UBound(imMktVefCode)) = tgMVef(ilLoop).iCode
    '            ReDim Preserve imMktVefCode(0 To UBound(imMktVefCode) + 1) As Integer
    '        End If
    '    Next ilLoop
    'Else
    '    imMktCode = -1
    'End If
    If ((lmStartStd > 0) Or (lmStartCal > 0)) And (imMktCode > 0) Then
        If (lmStartStd > 0) And (lmStartCal > 0) Then
            If lmStartStd < lmStartCal Then
                slStartDate = Format$(lmStartStd, "m/d/yy")
            Else
                slStartDate = Format$(lmStartCal, "m/d/yy")
            End If
        ElseIf lmStartStd > 0 Then
            slStartDate = Format$(lmStartStd, "m/d/yy")
        Else
            slStartDate = Format$(lmStartCal, "m/d/yy")
        End If
        If (lmEndStd > 0) And (lmEndCal > 0) Then
            If lmEndStd > lmEndCal Then
                slEndDate = Format$(lmEndStd, "m/d/yy")
            Else
                slEndDate = Format$(lmEndCal, "m/d/yy")
            End If
        ElseIf lmEndStd > 0 Then
            slEndDate = Format$(lmEndStd, "m/d/yy")
        Else
            slEndDate = Format$(lmEndCal, "m/d/yy")
        End If
        slStatus = "HO"
        slCntrType = ""
        ilHOType = 1
        sgCntrForDateStamp = ""
        ilRet = gObtainCntrForDate(PostRep, slStartDate, slEndDate, slStatus, slCntrType, ilHOType, tmChfAdvtExt())
    Else
        'ReDim tmChfAdvtExt(1 To 1) As CHFADVTEXT
        ReDim tmChfAdvtExt(0 To 0) As CHFADVTEXT
    End If
    ReDim tmContract(0 To 0) As SORTCODE
    For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
        For ilVeh = 0 To UBound(imMktVefCode) - 1 Step 1
            ilVefCode = imMktVefCode(ilVeh)
            ilFound = False
            If tmChfAdvtExt(ilChf).lVefCode > 0 Then
                If tmChfAdvtExt(ilChf).lVefCode = ilVefCode Then
                    ilFound = True
                End If
            ElseIf tmChfAdvtExt(ilChf).lVefCode < 0 Then
                tmVsfSrchKey.lCode = -tmChfAdvtExt(ilChf).lVefCode
                If tmVsf.lCode <> tmVsfSrchKey.lCode Then
                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                Do While ilRet = BTRV_ERR_NONE
                    For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                        If tmVsf.iFSCode(ilLoop) > 0 Then
                            If tmVsf.iFSCode(ilLoop) = ilVefCode Then
                                ilFound = True
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If ilFound Then
                        Exit Do
                    End If
                    If tmVsf.lLkVsfCode <= 0 Then
                        Exit Do
                    End If
                    tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
            If igPostType = 2 Then
                If imMktCode <> tmChfAdvtExt(ilChf).iAdfCode Then
                    ilFound = False
                End If
            End If
            'Test if invoices generated external
            'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
            '    If tgCommAdf(ilAdf).iCode = tmChfAdvtExt(ilChf).iAdfCode Then
                ilAdf = gBinarySearchAdf(tmChfAdvtExt(ilChf).iAdfCode)
                If ilAdf <> -1 Then
                    If tgCommAdf(ilAdf).sRepInvGen = "E" Then
                        ilFound = False
                    End If
            '        Exit For
                End If
            'Next ilAdf
            '6/9/10: Don't include contract that need to be posted by Date/Time: EDI Contracts
            If (tgSpf.sAEDII = "Y") And ((Asc(tgSpf.sUsingFeatures8) And REPBYDT) = REPBYDT) Then
                If (igPostType = 1) Or (igPostType = 2) Or (igPostType = 3) Then
                    ilAdf = gBinarySearchAdf(tmChfAdvtExt(ilChf).iAdfCode)
                    If ilAdf <> -1 Then
                        If (tgCommAdf(ilAdf).sBillAgyDir = "D") Then
                            If tgCommAdf(ilAdf).iArfInvCode > 0 Then
                                ilFound = False
                            End If
                        Else
                            ilAgf = gBinarySearchAgf(tmChfAdvtExt(ilChf).iAgfCode)
                            If ilAgf <> -1 Then
                                If tgCommAgf(ilAgf).iArfInvCode > 0 Then
                                    ilFound = False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If ilFound Then
                slStr = "Missing"
                'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                '    If tgCommAdf(ilAdf).iCode = tmChfAdvtExt(ilChf).iAdfCode Then
                    ilAdf = gBinarySearchAdf(tmChfAdvtExt(ilChf).iAdfCode)
                    If ilAdf <> -1 Then
                        If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
                            slStr = Trim$(tgCommAdf(ilAdf).sName) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID)
                        Else
                            slStr = Trim$(tgCommAdf(ilAdf).sName)
                        End If
                '        Exit For
                    End If
                'Next ilAdf
                Do While Len(slStr) < 30
                    slStr = slStr & " "
                Loop
                slKey = slStr
                'Contract number
                slStr = Trim$(str$(tmChfAdvtExt(ilChf).lCntrNo))
                Do While Len(slStr) < 8
                    slStr = "0" & slStr
                Loop
                slKey = slKey & slStr
                tmContract(UBound(tmContract)).sKey = slKey & "\" & Trim$(str$(tmChfAdvtExt(ilChf).lCode))
                ReDim Preserve tmContract(0 To UBound(tmContract) + 1) As SORTCODE
                Exit For
            End If
        Next ilVeh
    Next ilChf
    If UBound(tmContract) > 0 Then
        'ArraySortTyp tgSort(), tgSort(0), ilUpper, 0, Len(tgSort(0)), 0, -9, 0
        ArraySortTyp fnAV(tmContract(), 0), UBound(tmContract), 0, LenB(tmContract(0)), 0, LenB(tmContract(0).sKey), 0
    End If
    'Build images
    For ilChf = 0 To UBound(tmContract) - 1 Step 1
        slNameCode = tmContract(ilChf).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        llChfCode = Val(slCode)
        mObtainCntrInfo llChfCode
    Next ilChf
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
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
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilNoSpots As Integer
    Dim ilWk As Integer
    Dim ilAnyWkPosted As Integer

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxIndex) Then
        Exit Sub
    End If
    If (imRowNo < LBONE) Or (imRowNo > UBound(smSave, 2)) Then
        Exit Sub
    End If

    lacFrame.Visible = False
    pbcArrow.Visible = False
    Select Case ilBoxNo 'Branch on box type (control)
        Case imVEHICLEINDEX 'Vehicle
            lbcVehicle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcPostRep, slStr, tmCtrls(ilBoxNo)
            smShow(imVEHICLEINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow
            If imSave(3, imRowNo) <> lbcVehicle.ListIndex Then
                imSave(3, imRowNo) = lbcVehicle.ListIndex
                slNameCode = tmVehicle(imSave(3, imRowNo)).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                imSave(2, imRowNo) = Val(slCode)
                imChg = True
            End If
        Case imOPRICEINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
            gSetShow pbcPostRep, slStr, tmCtrls(ilBoxNo)
            smShow(imOPRICEINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow
            If lmSave(3, imRowNo) <> gStrDecToLong(slStr, 2) Then
                lmSave(3, imRowNo) = gStrDecToLong(slStr, 2)
                imChg = True
            End If
        Case imINVOICENOINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If lmSave(5, imRowNo) <> Val(slStr) Then
                lmSave(5, imRowNo) = slStr
                imChg = True
            End If
        Case imANOSPOTSINDEX
            plcWeekly.Visible = False
            pbcWeekly.Visible = False
            edcNoSpots.Visible = False
            If tgSpf.sPostCalAff = "W" Then
                ilNoSpots = 0
                ilAnyWkPosted = False
                For ilWk = 0 To 4 Step 1
                    If Trim$(smSave(18 + ilWk, imRowNo)) <> "" Then
                        ilAnyWkPosted = True
                        ilNoSpots = ilNoSpots + Val(smSave(18 + ilWk, imRowNo))
                    End If
                    If Trim$(smSave(23 + ilWk, imRowNo)) <> "" Then
                        ilAnyWkPosted = True
                        ilNoSpots = ilNoSpots + Val(smSave(23 + ilWk, imRowNo))
                    End If
                Next ilWk
                If ilAnyWkPosted Then
                    edcNoSpots.Text = ilNoSpots
                Else
                    edcNoSpots.Text = ""
                End If
            End If
            slStr = edcNoSpots.Text
            gSetShow pbcPostRep, slStr, tmCtrls(ilBoxNo)
            smShow(imANOSPOTSINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(5, imRowNo) <> edcNoSpots.Text Then
                imChg = True
                smSave(5, imRowNo) = edcNoSpots.Text
                mRecomputeTotals
                pbcPostRep.Cls
                pbcPostRep_Paint
            End If
        Case imANEXTNOSPOTSINDEX
            edcNoSpots.Visible = False
            slStr = edcNoSpots.Text
            gSetShow pbcPostRep, slStr, tmCtrls(ilBoxNo)
            smShow(imANEXTNOSPOTSINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(16, imRowNo) <> edcNoSpots.Text Then
                imChg = True
                smSave(16, imRowNo) = edcNoSpots.Text
                mRecomputeTotals
                pbcPostRep.Cls
                pbcPostRep_Paint
            End If
        'Case imAGROSSINDEX
        '    edcGross.Visible = False
        '    slStr = edcGross.Text
        '    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        '    gSetShow pbcPostRep, slStr, tmCtrls(ilBoxNo)
        '    smShow(imAGROSSINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow
        '    If smSave(6, imRowNo) <> edcGross.Text Then
        '        imChg = True
        '        smSave(6, imRowNo) = edcGross.Text
        '        slStr = smSave(6, imRowNo)
        '        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        '        gSetShow pbcPostRep, slStr, tmCtrls(imAGROSSINDEX)
        '        smShow(imAGROSSINDEX, imRowNo) = tmCtrls(imAGROSSINDEX).sShow
        '        mRecomputeTotals
        '        pbcPostRep.Cls
        '        pbcPostRep_Paint
        '    End If
        'Case imABONUSINDEX
        '    edcNoSpots.Visible = False
        '    slStr = edcNoSpots.Text
        '    gSetShow pbcPostRep, slStr, tmCtrls(ilBoxNo)
        '    smShow(imABONUSINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow
        '    If smSave(7, imRowNo) <> edcNoSpots.Text Then
        '        imChg = True
        '        smSave(7, imRowNo) = edcNoSpots.Text
        '        mRecomputeTotals
        '        pbcPostRep.Cls
        '        pbcPostRep_Paint
        '    End If
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
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

    For ilRowNo = LBONE To UBound(smSave, 2) - 1 Step 1
        If mTestSaveFields(ilRowNo) = NO Then
            imRowNo = ilRowNo
            imSettingValue = True
            If imRowNo <= vbcPostRep.LargeChange + 1 Then
                vbcPostRep.Value = vbcPostRep.Min
            Else
                vbcPostRep.Value = imRowNo - vbcPostRep.LargeChange
            End If
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
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields(ilRowNo As Integer) As Integer
'
'   iRet = mTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    If (imSave(3, ilRowNo) < 0) And (Trim$(smInfo(1, ilRowNo)) = "S") Then
        ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = imVEHICLEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    mTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGrandTotal                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute Grand Total            *
'*                                                     *
'*******************************************************
Sub mGrandTotal()
    Dim ilLoop As Integer
    Dim ilRowNo As Integer
    Dim slOGTotalNoPerWk As String
    Dim slOGTotalRate As String
    Dim slAGTotalNoPerWk As String
    Dim slAGTotalRate As String
    Dim slAMissed As String
    Dim slAMG As String
    Dim slABonus As String
    Dim slCalCarry As String
    Dim slPrior As String
    Dim slCalPrev As String
    Dim ilCol As Integer
    Dim ilCalCarryBonus As Integer
    Dim ilCurrBonus As Integer
    Dim ilPriorBonus As Integer

    slOGTotalNoPerWk = "0"
    slOGTotalRate = "0"
    slAGTotalNoPerWk = "0"
    slAGTotalRate = "0"
    slAMissed = "0"
    slAMG = "0"
    slABonus = "0"
    slCalCarry = "0"
    slCalPrev = "0"
    slPrior = "0"
    ilCalCarryBonus = 0
    ilCurrBonus = 0
    ilPriorBonus = 0
    For ilLoop = LBONE To UBound(smSave, 2) - 2 Step 1
        If ((Trim$(smInfo(1, ilLoop)) = "T") And ((igPostType = 4) Or (igPostType = 2) Or (igPostType = 5) Or (igPostType = 6))) Or (igPostType = 1) Or (igPostType = 3) Then
            slOGTotalNoPerWk = gAddStr(slOGTotalNoPerWk, smSave(3, ilLoop))
            If InStr(RTrim$(smSave(4, ilLoop)), ".") > 0 Then
                slOGTotalRate = gAddStr(slOGTotalRate, smSave(4, ilLoop))
            End If
            slAGTotalNoPerWk = gAddStr(slAGTotalNoPerWk, smSave(5, ilLoop))
            If InStr(RTrim$(smSave(6, ilLoop)), ".") > 0 Then
                slAGTotalRate = gAddStr(slAGTotalRate, smSave(6, ilLoop))
            End If
            slAMissed = gAddStr(slAMissed, smSave(14, ilLoop))
            slAMG = gAddStr(slAMG, smSave(13, ilLoop))
            slABonus = gAddStr(slABonus, smSave(7, ilLoop))
            slCalCarry = gAddStr(slCalCarry, smSave(16, ilLoop))
            slCalPrev = gAddStr(slCalPrev, smSave(9, ilLoop))
            slPrior = gAddStr(slPrior, smSave(10, ilLoop))
            ilCalCarryBonus = ilCalCarryBonus + imSave(9, ilLoop)
            ilCurrBonus = ilCurrBonus + imSave(8, ilLoop)
            ilPriorBonus = ilPriorBonus + imSave(10, ilLoop)
        End If
    Next ilLoop
    ilRowNo = UBound(smSave, 2) - 1
    'Set save values so that difference will be set
    smSave(3, ilRowNo) = slOGTotalNoPerWk
    smSave(4, ilRowNo) = slOGTotalRate
    smSave(5, ilRowNo) = slAGTotalNoPerWk
    smSave(6, ilRowNo) = slAGTotalRate
    smSave(7, ilRowNo) = slABonus
    smSave(13, ilRowNo) = slAMG
    smSave(14, ilRowNo) = slAMissed
    smSave(16, ilRowNo) = slCalCarry
    smSave(9, ilRowNo) = slCalPrev
    smSave(10, ilRowNo) = slPrior
    imSave(8, ilRowNo) = ilCurrBonus
    imSave(9, ilRowNo) = ilCalCarryBonus
    imSave(10, ilRowNo) = ilPriorBonus
    For ilCol = LBONE To UBound(smShow, 1) Step 1
        If ilCol <> imTOTALINDEX Then
            smShow(ilCol, ilRowNo) = ""
        Else
            smShow(ilCol, ilRowNo) = "Grand Total:"
        End If
    Next ilCol
    gSetShow pbcPostRep, slOGTotalNoPerWk, tmCtrls(imONOSPOTSINDEX)
    smShow(imONOSPOTSINDEX, ilRowNo) = tmCtrls(imONOSPOTSINDEX).sShow
    gSetShow pbcPostRep, slOGTotalRate, tmCtrls(imOGROSSINDEX)
    smShow(imOGROSSINDEX, ilRowNo) = tmCtrls(imOGROSSINDEX).sShow
    If imANOSPOTSINDEX > 0 Then
        gSetShow pbcPostRep, slAGTotalNoPerWk, tmCtrls(imANOSPOTSINDEX)
        smShow(imANOSPOTSINDEX, ilRowNo) = tmCtrls(imANOSPOTSINDEX).sShow
    End If
    If imANEXTNOSPOTSINDEX > 0 Then
        gSetShow pbcPostRep, slCalCarry, tmCtrls(imANEXTNOSPOTSINDEX)
        smShow(imANEXTNOSPOTSINDEX, ilRowNo) = tmCtrls(imANEXTNOSPOTSINDEX).sShow
    End If
    If imAPREVNOSPOTSINDEX > 0 Then
        gSetShow pbcPostRep, slCalPrev, tmCtrls(imAPREVNOSPOTSINDEX)
        smShow(imAPREVNOSPOTSINDEX, ilRowNo) = tmCtrls(imAPREVNOSPOTSINDEX).sShow
    End If
    If imPRIORNOSPOTSINDEX > 0 Then
        gSetShow pbcPostRep, slPrior, tmCtrls(imPRIORNOSPOTSINDEX)
        smShow(imPRIORNOSPOTSINDEX, ilRowNo) = tmCtrls(imPRIORNOSPOTSINDEX).sShow
    End If
    'gSetShow pbcPostRep, slAGTotalRate, tmCtrls(imAGROSSINDEX)
    'smShow(imAGROSSINDEX, ilRowNo) = tmCtrls(imAGROSSINDEX).sShow
    'gSetShow pbcPostRep, slABonus, tmCtrls(imABONUSINDEX)
    'smShow(imABONUSINDEX, ilRowNo) = tmCtrls(imABONUSINDEX).sShow
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
    Dim ilRet As Integer    'Return Status
    Dim ilCol As Integer
    'Dim tlSbf As SBF    'Only used to get size of SBF

    imLBWklyCtrls = 1
    imLBCtrls = 1
    imLINEINDEX = -100
    imADVTVEHDAYPARTINDEX = -100
    imLENGTHINDEX = -100
    imANOSPOTSPRIORINDEX = -100
    imANOSPOTSCURRINDEX = -100
    imDIFFSPOTSINDEX = -100
    imDIFFGROSSINDEX = -100
    imBONUSPREVINDEX = -100
    imBONUSCURRINDEX = -100
    imNNOSPOTSINDEX = -100
    imNNOBONUSINDEX = -100
    If (igPostType = 5) Or (igPostType = 6) Then  'Rep Spot Times
        pbcPostRep.Picture = pbcPost(4).Picture
        lacScreen.Caption = "Rep Spot Times"
        imCONTRACTINDEX = 1
        imLINEINDEX = 2
        imCashTradeIndex = -100
        imAdvtIndex = -100
        imVEHICLEINDEX = -100
        imADVTVEHDAYPARTINDEX = 3
        imLENGTHINDEX = 4
        imPRIORNOSPOTSINDEX = -100
        imONOSPOTSINDEX = 5
        imOPRICEINDEX = -100
        imOGROSSINDEX = 6
        imANOSPOTSPRIORINDEX = 7
        imANOSPOTSCURRINDEX = 8
        imANOSPOTSINDEX = -100
        imMISSEDINDEX = -100
        imMGINDEX = -100
        imBONUSINDEX = -100
        imDGROSSINDEX = -100
        imDIFFSPOTSINDEX = 9
        imDIFFGROSSINDEX = 10
        imBONUSPREVINDEX = 11
        imBONUSCURRINDEX = 12
        imNNOSPOTSINDEX = 13
        imNNOBONUSINDEX = 14
        imMaxIndex = 14
        imCHECKINDEX = -100
        'imVEHICLEINDEX = -100
        imINVOICENOINDEX = -100
        imAPREVNOSPOTSINDEX = -100
        imANEXTNOSPOTSINDEX = -100
        imRECEIVEDINDEX = -100
        imPOSTEDINDEX = -100
    ElseIf igPostType = 4 Then  'Cluster
        pbcPostRep.Picture = pbcPost(0).Picture
        lacScreen.Caption = "Post by Cluster"
        imCONTRACTINDEX = 1
        imCashTradeIndex = 2
        imAdvtIndex = 3
        imVEHICLEINDEX = 4
        imPRIORNOSPOTSINDEX = 5
        imONOSPOTSINDEX = 6
        imOPRICEINDEX = 7
        imOGROSSINDEX = 8
        imANOSPOTSINDEX = 9
        imMISSEDINDEX = 10
        imMGINDEX = 11
        imBONUSINDEX = 12
        imDGROSSINDEX = 13
        imMaxIndex = 13
        imCHECKINDEX = -100
        imINVOICENOINDEX = -100
        imAPREVNOSPOTSINDEX = -100
        imANEXTNOSPOTSINDEX = -100
        imRECEIVEDINDEX = -100
        imPOSTEDINDEX = -100
    ElseIf igPostType = 3 Then  'Received
        pbcPostRep.Picture = pbcPost(3).Picture
        lacScreen.Caption = "Post Received"
        ckcAll.Visible = True
        imCHECKINDEX = 1
        imCONTRACTINDEX = 2
        imCashTradeIndex = 3
        imAdvtIndex = 4
        imPRIORNOSPOTSINDEX = 5
        imONOSPOTSINDEX = 6
        imOPRICEINDEX = 7
        imOGROSSINDEX = 8
        imRECEIVEDINDEX = 9
        imPOSTEDINDEX = 10
        imMaxIndex = 10
        imVEHICLEINDEX = -100
        imINVOICENOINDEX = -100
        imAPREVNOSPOTSINDEX = -100
        imANOSPOTSINDEX = -100
        imANEXTNOSPOTSINDEX = -100
        imMISSEDINDEX = -100
        imMGINDEX = -100
        imBONUSINDEX = -100
        imDGROSSINDEX = -100
    Else    'Rep- non cluster
        If tgSpf.sPostCalAff = "C" Then
            pbcPostRep.Picture = pbcPost(2).Picture
            imCONTRACTINDEX = 1
            imCashTradeIndex = 2
            'imVEHICLEINDEX = 3
            imPRIORNOSPOTSINDEX = 4
            imONOSPOTSINDEX = 5
            imOPRICEINDEX = 6
            imOGROSSINDEX = 7
            imAPREVNOSPOTSINDEX = 8
            imANOSPOTSINDEX = 9
            imANEXTNOSPOTSINDEX = 10
            imMISSEDINDEX = 11
            imMGINDEX = 12
            imBONUSINDEX = 13
            imDGROSSINDEX = 14
            imMaxIndex = 14
            imCHECKINDEX = -100
            imINVOICENOINDEX = -100
            'imAdvtIndex = -100
            imRECEIVEDINDEX = -100
            imPOSTEDINDEX = -100
            lacCalDates.Visible = True
        ElseIf igPostType = 1 And tgSpf.sPostCalAff = "W" Then
            pbcPostRep.Picture = pbcPost(1).Picture
            imCONTRACTINDEX = 1
            imCashTradeIndex = 2
            'imAdvtIndex = 3
            imPRIORNOSPOTSINDEX = 4
            imONOSPOTSINDEX = 5
            imOPRICEINDEX = 6
            imOGROSSINDEX = 7
            imINVOICENOINDEX = 8
            imANOSPOTSINDEX = 9
            imMISSEDINDEX = 10
            imMGINDEX = 11
            imBONUSINDEX = 12
            imDGROSSINDEX = 13
            imMaxIndex = 13
            imCHECKINDEX = -100
            'imVEHICLEINDEX = -100
            imAPREVNOSPOTSINDEX = -100
            imANEXTNOSPOTSINDEX = -100
            imRECEIVEDINDEX = -100
            imPOSTEDINDEX = -100
        Else
            pbcPostRep.Picture = pbcPost(1).Picture
            imCONTRACTINDEX = 1
            imCashTradeIndex = 2
            'imAdvtIndex = 3
            imPRIORNOSPOTSINDEX = 4
            imONOSPOTSINDEX = 5
            imOPRICEINDEX = 6
            imOGROSSINDEX = 7
            imANOSPOTSINDEX = 8
            imMISSEDINDEX = 9
            imMGINDEX = 10
            imBONUSINDEX = 11
            imDGROSSINDEX = 12
            imMaxIndex = 12
            imCHECKINDEX = -100
            imINVOICENOINDEX = -100
            'imVEHICLEINDEX = -100
            imAPREVNOSPOTSINDEX = -100
            imANEXTNOSPOTSINDEX = -100
            imRECEIVEDINDEX = -100
            imPOSTEDINDEX = -100
        End If
        If igPostType = 1 Then
            lacScreen.Caption = "Post by Vehicle"
            imAdvtIndex = 3
            imVEHICLEINDEX = -100
        Else
            lacScreen.Caption = "Post by Advertiser"
            If tgSpf.sPostCalAff <> "W" Then
                imcInsert.Visible = True
            End If
            imVEHICLEINDEX = 3
            imAdvtIndex = -100
        End If
    End If
    If (igPostType = 5) Or (igPostType = 6) Then
        cmcDone.Left = cmcUpdate.Left
        cmcCancel.Visible = False
        cmcUpdate.Visible = False
        cmcImport.Visible = False
        cmcClear.Visible = False
        imcTrash.Visible = False
    ElseIf igPostType <> 4 Then
        cmcDone.Left = cmcCancel.Left
        cmcCancel.Left = cmcUpdate.Left
        cmcUpdate.Left = cmcImport.Left
        cmcImport.Visible = False
        cmcClear.Visible = False
    End If
    ReDim smSave(0 To 33, 0 To 1) As String
    ReDim imSave(0 To 10, 0 To 1) As Integer
    ReDim lmSave(0 To 5, 0 To 1) As Long
    ReDim smShow(0 To 10, 0 To 1) As String * 40
    ReDim smInfo(0 To 13, 0 To 1) As String * 12
    For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilCol, 1) = ""
    Next ilCol
    For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
        smInfo(ilCol, 1) = ""
    Next ilCol
    ReDim lmSbfDel(0 To 0) As Long
    Screen.MousePointer = vbHourglass
    smNowDate = Format$(gNow(), "m/d/yy")
    igJobShowing(POSTLOGSJOB) = True
    imFirstActivate = True
    imcKey.Picture = IconTraf!imcKey.Picture
    imcInsert.Picture = IconTraf!imcInsert.Picture
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    vbcPostRep.Min = LBONE  'LBound(smShow, 2)
    vbcPostRep.Max = LBONE  'LBound(smShow, 2)
    vbcPostRep.Value = vbcPostRep.Min
    'gPDNToStr tgSpf.sBTax(0), 2, slStr1
    'gPDNToStr tgSpf.sBTax(1), 2, slStr2
    'If (Val(slStr1) = 0) And (Val(slStr2) = 0) Then
    '12/17/06-Change to tax by agency or vehicle
    'If (tgSpf.iBTax(0) = 0) Or (tgSpf.iBTax(1) = 0) Then
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
        imTaxDefined = True
    Else
        imTaxDefined = False
    End If
    imSetAll = True
    imMarketIndex = -1
    imInvDateIndex = -1
    sgCntrForDateStamp = ""
    imIgnoreRightMove = False
    imTerminate = False
    imFirstActivate = True
    imBypassFocus = False
    imBoxNo = -1 'Initialize current Box to N/A
    imRowNo = -1
    imChg = False
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imSettingValue = False
    imChgMode = False
    imBSMode = False
    hmNrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmNrf, "", sgDBPath & "Nrf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Nrf.Btr)", PostRep
    On Error GoTo 0
    imNrfRecLen = Len(tmNrf) 'btrRecordLength(hmChf)    'Get Chf size
    If igPostType = 5 Then
        If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
            mNetNamePop
        Else
            mVehPop 1
        End If
    ElseIf igPostType = 6 Then
        mVehPop 1
    ElseIf igPostType = 4 Then
        mMarketPop
    ElseIf igPostType = 3 Then
        mVehPop 1
    ElseIf igPostType = 2 Then
        igPopExternalAdvt = False
        mAdvtPop
        igPopExternalAdvt = True
        'Reuired by insert
        mVehPop 2
    ElseIf igPostType = 1 Then
        mVehPop 1
    End If
    If imTerminate Then
        Exit Sub
    End If
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", PostRep
    On Error GoTo 0
    imCHFRecLen = Len(tmChf) 'btrRecordLength(hmChf)    'Get Chf size
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", PostRep
    On Error GoTo 0
    imClfRecLen = Len(tmClf) 'btrRecordLength(hmClf)    'Get Clf size
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", PostRep
    On Error GoTo 0
    imCffRecLen = Len(tmCff) 'btrRecordLength(hmCff)    'Get Cff size
    hmCxf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cxf.Btr)", PostRep
    On Error GoTo 0
    imCxfRecLen = Len(tmCxf) 'btrRecordLength(hmCff)    'Get Cff size
    hmSbf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sbf.Btr)", PostRep
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf) 'btrRecordLength(hmSbf)    'Get Sbf size
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", PostRep
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf) 'btrRecordLength(hmLcf)    'Get Lcf size
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", PostRep
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmRwf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRwf, "", sgDBPath & "Rwf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rwf.Btr)", PostRep
    On Error GoTo 0
    imRwfRecLen = Len(tmRwf)
    hmRdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rdf.Btr)", PostRep
    On Error GoTo 0
    imRdfRecLen = Len(tmRdf) 'btrRecordLength(hmChf)    'Get Chf size
    hmRvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rvf.Btr)", PostRep
    On Error GoTo 0
    imRvfRecLen = Len(tmRvf) 'btrRecordLength(hmChf)    'Get Chf size

    mDatePop
    ilRet = gObtainVef()
    ilRet = gObtainAdvt()
    ''mVehPop
    'PostRep.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    ''gCenterModalForm PostRep
    gCenterForm PostRep
    'Traffic!plcHelp.Caption = ""
    mInitBox
    lacTotals.Visible = False
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
'*             Created:9/02/93       By:D. LeVine      *
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
    Dim llMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long

    imLBCtrls = 1
    flTextHeight = pbcPostRep.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcPostRep.Move 120, 555, pbcPostRep.Width + vbcPostRep.Width + fgPanelAdj, pbcPostRep.Height + fgPanelAdj
    pbcPostRep.Move plcPostRep.Left + fgBevelX, plcPostRep.Top + fgBevelY
    vbcPostRep.Move pbcPostRep.Left + pbcPostRep.Width, pbcPostRep.Top + 15
    pbcArrow.Move plcPostRep.Left - pbcArrow.Width - 15    'Vehicle
    plcInfo.Move plcPostRep.Left + (plcPostRep.Width - plcInfo.Width) / 2, plcPostRep.Top + plcPostRep.Height - 60
    pbcKey.Move plcPostRep.Left, plcPostRep.Top
    lacCalDates.Move plcPostRep.Left, plcPostRep.Top + plcPostRep.Height + 60
    imcInsert.Move lacCalDates.Left + 120, lacCalDates.Top + lacCalDates.Height + 45
    If (igPostType = 5) Or (igPostType = 6) Then
        'Contract
        gSetCtrl tmCtrls(imCONTRACTINDEX), 30, 375, 630, fgBoxGridH
        'Line #
        gSetCtrl tmCtrls(imLINEINDEX), 675, tmCtrls(imCONTRACTINDEX).fBoxY, 300, fgBoxGridH
        'Advertiser, Vehicle and Daypart
        gSetCtrl tmCtrls(imADVTVEHDAYPARTINDEX), 990, tmCtrls(imCONTRACTINDEX).fBoxY, 2655, fgBoxGridH
        gSetCtrl tmCtrls(imLENGTHINDEX), 3660, tmCtrls(imCONTRACTINDEX).fBoxY, 240, fgBoxGridH
        'Ordered Number of Spots
        gSetCtrl tmCtrls(imONOSPOTSINDEX), 3915, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'Ordered Gross
        gSetCtrl tmCtrls(imOGROSSINDEX), 4320, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
        'Aired Number of Spots: Posted Prior number
        gSetCtrl tmCtrls(imANOSPOTSPRIORINDEX), 5160, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'Aired Number of Spots: Posted Current month
        gSetCtrl tmCtrls(imANOSPOTSCURRINDEX), 5565, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'Spot difference
        gSetCtrl tmCtrls(imDIFFSPOTSINDEX), 5970, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'Gross difference
        gSetCtrl tmCtrls(imDIFFGROSSINDEX), 6375, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
        'Bonus Number of Spots
        gSetCtrl tmCtrls(imBONUSPREVINDEX), 7215, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'Difference Gross
        gSetCtrl tmCtrls(imBONUSCURRINDEX), 7625, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'Bonus Number of Spots
        gSetCtrl tmCtrls(imNNOSPOTSINDEX), 8025, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'Difference Gross
        gSetCtrl tmCtrls(imNNOBONUSINDEX), 8430, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
    ElseIf igPostType = 4 Then
        'Contract
        gSetCtrl tmCtrls(imCONTRACTINDEX), 30, 375, 630, fgBoxGridH
        'Cash/Trade
        gSetCtrl tmCtrls(imCashTradeIndex), 675, tmCtrls(imCONTRACTINDEX).fBoxY, 240, fgBoxGridH
        'Advertiser
        gSetCtrl tmCtrls(imAdvtIndex), 930, tmCtrls(imCONTRACTINDEX).fBoxY, 1455, fgBoxGridH
        'Vehicle
        gSetCtrl tmCtrls(imVEHICLEINDEX), 2400, tmCtrls(imCONTRACTINDEX).fBoxY, 1455, fgBoxGridH
        'Prior Spots
        gSetCtrl tmCtrls(imPRIORNOSPOTSINDEX), 3870, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'Ordered Number of Spots
        gSetCtrl tmCtrls(imONOSPOTSINDEX), 4275, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'Spot Price
        gSetCtrl tmCtrls(imOPRICEINDEX), 4680, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
        'Ordered Gross
        gSetCtrl tmCtrls(imOGROSSINDEX), 5520, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
        'Aired Number of Spots
        gSetCtrl tmCtrls(imANOSPOTSINDEX), 6360, tmCtrls(imCONTRACTINDEX).fBoxY, 405, fgBoxGridH
        'Missed
        gSetCtrl tmCtrls(imMISSEDINDEX), 6780, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'MG
        gSetCtrl tmCtrls(imMGINDEX), 7185, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'Bonus Number of Spots
        gSetCtrl tmCtrls(imBONUSINDEX), 7590, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
        'Difference Gross
        gSetCtrl tmCtrls(imDGROSSINDEX), 7995, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
    ElseIf igPostType = 3 Then
        'Check
        gSetCtrl tmCtrls(imCHECKINDEX), 30, 375, 195, fgBoxGridH
        'Contract
        gSetCtrl tmCtrls(imCONTRACTINDEX), 240, tmCtrls(imCHECKINDEX).fBoxY, 660, fgBoxGridH
        'Cash/Trade
        gSetCtrl tmCtrls(imCashTradeIndex), 915, tmCtrls(imCHECKINDEX).fBoxY, 240, fgBoxGridH
        'Advertiser
        gSetCtrl tmCtrls(imAdvtIndex), 1170, tmCtrls(imCHECKINDEX).fBoxY, 3480, fgBoxGridH
        'Prior Spots
        gSetCtrl tmCtrls(imPRIORNOSPOTSINDEX), 4665, tmCtrls(imCHECKINDEX).fBoxY, 390, fgBoxGridH
        'Ordered Number of Spots
        gSetCtrl tmCtrls(imONOSPOTSINDEX), 5070, tmCtrls(imCHECKINDEX).fBoxY, 390, fgBoxGridH
        'Ordered Price
        gSetCtrl tmCtrls(imOPRICEINDEX), 5475, tmCtrls(imCHECKINDEX).fBoxY, 825, fgBoxGridH
        'Ordered Gross
        gSetCtrl tmCtrls(imOGROSSINDEX), 6315, tmCtrls(imCHECKINDEX).fBoxY, 825, fgBoxGridH
        'Received
        gSetCtrl tmCtrls(imRECEIVEDINDEX), 7155, tmCtrls(imCHECKINDEX).fBoxY, 825, fgBoxGridH
        'Posted
        gSetCtrl tmCtrls(imPOSTEDINDEX), 7995, tmCtrls(imCHECKINDEX).fBoxY, 825, fgBoxGridH
    Else
        If tgSpf.sPostCalAff = "C" Then
            'Contract
            gSetCtrl tmCtrls(imCONTRACTINDEX), 30, 375, 630, fgBoxGridH
            'Cash/Trade
            gSetCtrl tmCtrls(imCashTradeIndex), 675, tmCtrls(imCONTRACTINDEX).fBoxY, 240, fgBoxGridH
            'Advertiser
            If imAdvtIndex > 0 Then
                gSetCtrl tmCtrls(imAdvtIndex), 930, tmCtrls(imCONTRACTINDEX).fBoxY, 2115, fgBoxGridH
            End If
            If imVEHICLEINDEX > 0 Then
                gSetCtrl tmCtrls(imVEHICLEINDEX), 930, tmCtrls(imCONTRACTINDEX).fBoxY, 2115, fgBoxGridH
            End If
            'Prior Spots
            gSetCtrl tmCtrls(imPRIORNOSPOTSINDEX), 3060, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Ordered Number of Spots
            gSetCtrl tmCtrls(imONOSPOTSINDEX), 3465, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Ordered Price
            gSetCtrl tmCtrls(imOPRICEINDEX), 3870, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
            'Ordered Gross
            gSetCtrl tmCtrls(imOGROSSINDEX), 4710, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
            'Prev Number of Spots
            gSetCtrl tmCtrls(imAPREVNOSPOTSINDEX), 5550, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Aired Number of Spots
            gSetCtrl tmCtrls(imANOSPOTSINDEX), 5955, tmCtrls(imCONTRACTINDEX).fBoxY, 405, fgBoxGridH
            'Aired Number of Spots in next stanard period
            gSetCtrl tmCtrls(imANEXTNOSPOTSINDEX), 6375, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Missed
            gSetCtrl tmCtrls(imMISSEDINDEX), 6780, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'MG
            gSetCtrl tmCtrls(imMGINDEX), 7185, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Bonus Number of Spots
            gSetCtrl tmCtrls(imBONUSINDEX), 7590, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Difference Gross
            gSetCtrl tmCtrls(imDGROSSINDEX), 7995, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
        ElseIf igPostType = 1 And tgSpf.sPostCalAff = "W" Then
            'Contract
            gSetCtrl tmCtrls(imCONTRACTINDEX), 30, 375, 630, fgBoxGridH
            'Cash/Trade
            gSetCtrl tmCtrls(imCashTradeIndex), 675, tmCtrls(imCONTRACTINDEX).fBoxY, 240, fgBoxGridH
            'Advertiser
            If imAdvtIndex > 0 Then
                'gSetCtrl tmCtrls(imAdvtIndex), 930, tmCtrls(imCONTRACTINDEX).fBoxY, 2925, fgBoxGridH
                gSetCtrl tmCtrls(imAdvtIndex), 930, tmCtrls(imCONTRACTINDEX).fBoxY, 2400, fgBoxGridH
            End If
            If imVEHICLEINDEX > 0 Then
                'gSetCtrl tmCtrls(imVEHICLEINDEX), 930, tmCtrls(imCONTRACTINDEX).fBoxY, 2925, fgBoxGridH
                gSetCtrl tmCtrls(imVEHICLEINDEX), 930, tmCtrls(imCONTRACTINDEX).fBoxY, 2400, fgBoxGridH
            End If
            'Prior Spots
            gSetCtrl tmCtrls(imPRIORNOSPOTSINDEX), 3345, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Ordered Number of Spots
            gSetCtrl tmCtrls(imONOSPOTSINDEX), 3750, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Ordered Price
            gSetCtrl tmCtrls(imOPRICEINDEX), 4155, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
            'Ordered Gross
            gSetCtrl tmCtrls(imOGROSSINDEX), 4995, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
            'Station Invoice Number
            gSetCtrl tmCtrls(imINVOICENOINDEX), 5835, tmCtrls(imCONTRACTINDEX).fBoxY, 510, fgBoxGridH
            'Aired Number of Spots
            gSetCtrl tmCtrls(imANOSPOTSINDEX), 6360, tmCtrls(imCONTRACTINDEX).fBoxY, 405, fgBoxGridH
            'Missed
            gSetCtrl tmCtrls(imMISSEDINDEX), 6780, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'MG
            gSetCtrl tmCtrls(imMGINDEX), 7185, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Bonus Number of Spots
            gSetCtrl tmCtrls(imBONUSINDEX), 7590, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Difference Gross
            gSetCtrl tmCtrls(imDGROSSINDEX), 7995, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
            'If tgSpf.sPostCalAff = "W" Then
                gSetCtrl tmWklyCtrls(WKLYDATESINDEX), 30, 225, 1200, fgBoxGridH
                gSetCtrl tmWklyCtrls(WKLYORDERNOINDEX), 1245, tmWklyCtrls(WKLYDATESINDEX).fBoxY, 720, fgBoxGridH
                gSetCtrl tmWklyCtrls(WKLYAIRNOINDEX), 1980, tmWklyCtrls(WKLYDATESINDEX).fBoxY, 720, fgBoxGridH
                gSetCtrl tmWklyCtrls(WKLYCARRIEDNOINDEX), 2715, tmWklyCtrls(WKLYDATESINDEX).fBoxY, 720, fgBoxGridH
                gSetCtrl tmWklyCtrls(WKLYTOTALINDEX), 3450, tmWklyCtrls(WKLYDATESINDEX).fBoxY, 720, fgBoxGridH
            'End If
        Else
            'Contract
            gSetCtrl tmCtrls(imCONTRACTINDEX), 30, 375, 630, fgBoxGridH
            'Cash/Trade
            gSetCtrl tmCtrls(imCashTradeIndex), 675, tmCtrls(imCONTRACTINDEX).fBoxY, 240, fgBoxGridH
            'Advertiser
            If imAdvtIndex > 0 Then
                gSetCtrl tmCtrls(imAdvtIndex), 930, tmCtrls(imCONTRACTINDEX).fBoxY, 2925, fgBoxGridH
            End If
            If imVEHICLEINDEX > 0 Then
                gSetCtrl tmCtrls(imVEHICLEINDEX), 930, tmCtrls(imCONTRACTINDEX).fBoxY, 2925, fgBoxGridH
            End If
            'Prior Spots
            gSetCtrl tmCtrls(imPRIORNOSPOTSINDEX), 3870, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Ordered Number of Spots
            gSetCtrl tmCtrls(imONOSPOTSINDEX), 4275, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Ordered Price
            gSetCtrl tmCtrls(imOPRICEINDEX), 4680, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
            'Ordered Gross
            gSetCtrl tmCtrls(imOGROSSINDEX), 5520, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
            'Aired Number of Spots
            gSetCtrl tmCtrls(imANOSPOTSINDEX), 6360, tmCtrls(imCONTRACTINDEX).fBoxY, 405, fgBoxGridH
            'Missed
            gSetCtrl tmCtrls(imMISSEDINDEX), 6780, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'MG
            gSetCtrl tmCtrls(imMGINDEX), 7185, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Bonus Number of Spots
            gSetCtrl tmCtrls(imBONUSINDEX), 7590, tmCtrls(imCONTRACTINDEX).fBoxY, 390, fgBoxGridH
            'Difference Gross
            gSetCtrl tmCtrls(imDGROSSINDEX), 7995, tmCtrls(imCONTRACTINDEX).fBoxY, 825, fgBoxGridH
            If tgSpf.sPostCalAff = "W" Then
                gSetCtrl tmWklyCtrls(WKLYDATESINDEX), 30, 225, 1200, fgBoxGridH
                gSetCtrl tmWklyCtrls(WKLYORDERNOINDEX), 1245, tmWklyCtrls(WKLYDATESINDEX).fBoxY, 720, fgBoxGridH
                gSetCtrl tmWklyCtrls(WKLYAIRNOINDEX), 1980, tmWklyCtrls(WKLYDATESINDEX).fBoxY, 720, fgBoxGridH
                gSetCtrl tmWklyCtrls(WKLYCARRIEDNOINDEX), 2715, tmWklyCtrls(WKLYDATESINDEX).fBoxY, 720, fgBoxGridH
                gSetCtrl tmWklyCtrls(WKLYTOTALINDEX), 3450, tmWklyCtrls(WKLYDATESINDEX).fBoxY, 720, fgBoxGridH
            End If
        
        End If
    End If
    If (igPostType <> 5) And (igPostType <> 6) Then
        If imAdvtIndex > 0 Then
            imTOTALINDEX = imAdvtIndex
        Else
            imTOTALINDEX = imVEHICLEINDEX
        End If
    Else
        imTOTALINDEX = imADVTVEHDAYPARTINDEX
    End If
    llMax = 0
    For ilLoop = imLBCtrls To imMaxIndex Step 1
        tmCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxW)
        Do While (tmCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmCtrls(ilLoop).fBoxW = tmCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxX)
            Do While (tmCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmCtrls(ilLoop).fBoxX = tmCtrls(ilLoop).fBoxX + 1
            Loop
            If tmCtrls(ilLoop).fBoxX > 90 Then
                Do
                    If tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 < tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 > tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    pbcPostRep.Picture = LoadPicture("")
    pbcPostRep.Width = llMax
    plcPostRep.Width = llMax + vbcPostRep.Width + 2 * fgBevelX + 15
    lacFrame.Width = llMax - 15
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    If (igPostType = 5) Or (igPostType = 6) Then
        cmcDone.Left = Me.Width / 2 - cmcDone.Width / 2
    ElseIf igPostType <> 4 Then
        cmcDone.Left = (PostRep.Width - 3 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    Else
        cmcDone.Left = (PostRep.Width - 5 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    End If
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
    cmcUpdate.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    If igPostType = 4 Then
        cmcImport.Left = cmcUpdate.Left + cmcUpdate.Width + ilSpaceBetweenButtons
        cmcClear.Left = cmcImport.Left + cmcImport.Width + ilSpaceBetweenButtons
    End If
    cmcDone.Top = PostRep.Height - (3 * cmcDone.Height) / 2
    cmcCancel.Top = cmcDone.Top
    cmcUpdate.Top = cmcDone.Top
    cmcImport.Top = cmcDone.Top
    cmcClear.Top = cmcDone.Top
    imcTrash.Top = cmcDone.Top + cmcDone.Height - imcTrash.Height
    imcTrash.Left = PostRep.Width - (3 * imcTrash.Width) / 2
    imcInsert.Top = imcTrash.Top
    lacCalDates.Top = imcInsert.Top - lacCalDates.Height
    ckcAll.Top = lacCalDates.Top
    llAdjTop = ckcAll.Top - plcPostRep.Top - 45 - tmCtrls(1).fBoxY
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    Do While plcPostRep.Top + llAdjTop + 2 * fgBevelY + 240 < ckcAll.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    plcPostRep.Height = llAdjTop + 2 * fgBevelY
    pbcPostRep.Left = plcPostRep.Left + fgBevelX
    pbcPostRep.Top = plcPostRep.Top + fgBevelY
    pbcPostRep.Height = plcPostRep.Height - 2 * fgBevelY
    vbcPostRep.Left = pbcPostRep.Left + pbcPostRep.Width + 15
    vbcPostRep.Top = pbcPostRep.Top
    vbcPostRep.Height = pbcPostRep.Height

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMarketPop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Market combobox       *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Sub mMarketPop()
'
'   mMarketPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim slNameCode As String
    Dim ilSortCode As Integer
    Dim ilLoop As Integer
    Dim llLen As Long
    Dim slStr As String

    cbcNames.Clear
    ilSortCode = 0
    ReDim tmMktCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    ilRet = gObtainMnfForType("H3", slStr, tgMkMnf())
    For ilLoop = LBound(tgMkMnf) To UBound(tgMkMnf) - 1 Step 1
        If Trim$(tgMkMnf(ilLoop).sRPU) = "Y" Then
            slName = Trim$(tgMkMnf(ilLoop).sName)
            slName = slName & "\" & Trim$(str$(tgMkMnf(ilLoop).iCode))
            tmMktCode(ilSortCode).sKey = slName
            If ilSortCode >= UBound(tmMktCode) Then
                ReDim Preserve tmMktCode(0 To UBound(tmMktCode) + 100) As SORTCODE
            End If
            ilSortCode = ilSortCode + 1
        End If
    Next ilLoop
    ReDim Preserve tmMktCode(0 To ilSortCode) As SORTCODE
    If UBound(tmMktCode) - 1 > 0 Then
        ArraySortTyp fnAV(tmMktCode(), 0), UBound(tmMktCode), 0, LenB(tmMktCode(0)), 0, LenB(tmMktCode(0).sKey), 0
    End If
    llLen = 0
    For ilLoop = 0 To UBound(tmMktCode) - 1 Step 1
        slNameCode = tmMktCode(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet = CP_MSG_NONE Then
            slName = Trim$(slName)
            If Not gOkAddStrToListBox(slName, llLen, True) Then
                Exit For
            End If
            cbcNames.AddItem slName  'Add ID to list box
        End If
    Next ilLoop
    Exit Sub

    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMerge                          *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Merge SBF record into Save     *
'*                      images                         *
'*                                                     *
'*******************************************************
Function mMerge(slSource As String) As Integer
'
'   slSource(I)-  I=Import; F=File (sbf record)
'
    Dim slStr As String
    Dim ilRet As Integer    'Return status
    Dim ilLoop As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilIndex As Integer
    Dim ilNewRow As Integer
    Dim ilPass As Integer
    Dim ilSPass As Integer
    Dim ilEPass As Integer
    Dim ilFound As Integer
    Dim ilAdf As Integer
    Dim ilVef As Integer
    Dim slSpotRate As String
    Dim ilWk As Integer
    Dim ilLenAcqMatch As Integer
    Dim ilMatch As Integer

    ilFound = False
    'Pass zero- match on price
    'Pass 1- bypass price test
    If tmSbf.sInserted <> "Y" Then
        For ilPass = 0 To 1 Step 1
            For ilRow = LBONE To UBound(smSave, 2) - 1 Step 1
                If (InStr(1, smShow(imTOTALINDEX, ilRow), "Total:", 1) <= 0) Then
                    If ilPass = 0 Then
                        If (igPostType = 5) Or (igPostType = 6) Then
                            If (tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.iLineNo = imSave(7, ilRow)) Then
                                ilMatch = True
                            Else
                                ilMatch = False
                            End If
                        Else
                            If (tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.iAirVefCode = imSave(2, ilRow)) And (tmSbf.sCashTrade = smSave(1, ilRow)) And ((tmSbf.lSpotPrice = lmSave(3, ilRow)) Or (ilPass = 1)) Then
                                ilMatch = True
                            Else
                                ilMatch = False
                            End If
                        End If
                    Else
                        If (igPostType = 5) Or (igPostType = 6) Then
                            ilMatch = False
                        Else
                            If (tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.iAirVefCode = imSave(2, ilRow)) And (tmSbf.sCashTrade = smSave(1, ilRow)) And ((tmSbf.lSpotPrice = lmSave(3, ilRow)) Or (ilPass = 1)) Then
                                ilMatch = True
                            Else
                                ilMatch = False
                            End If
                        End If
                    End If
                    'If (tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.iAirVefCode = imSave(2, ilRow)) And (tmSbf.sCashTrade = smSave(1, ilRow)) And ((tmSbf.lSpotPrice = lmSave(3, ilRow)) Or (ilPass = 1)) Then
                    If ilMatch Then
                        ilLenAcqMatch = True
                        If (Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER Then
                            If imSave(6, ilRow) <> tmSbf.iSpotLen Then
                                ilLenAcqMatch = False
                            End If
                        End If
                        '6/7/15: replaced acquisition from site override with Barter in system options
                        'If ((Asc(tgSpf.sOverrideOptions) And SPACQUISITION) = SPACQUISITION) Or ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
                        If ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
                            If lmSave(4, ilRow) <> tmSbf.lAcquisitionCost Then
                                ilLenAcqMatch = False
                            End If
                        End If
                        If ilLenAcqMatch Then
                            ilFound = True
                            If slSource = "F" Then
                                lmSave(2, ilRow) = tmSbf.lCode
                            End If
                            If Trim$(smInfo(1, ilRow)) = "C" Then
                                If slSource = "F" Then
                                    If tmSbf.sPostStatus = "P" Then
                                        smSave(5, ilRow) = Trim$(str$(tmSbf.iAirNoSpots))
                                        smSave(16, ilRow) = Trim$(str$(tmSbf.iCalCarryOver))
                                    Else
                                        smSave(5, ilRow) = ""
                                        smSave(16, ilRow) = ""
                                    End If
                                Else
                                    smSave(5, ilRow) = Trim$(str$(tmSbf.iAirNoSpots))
                                    smSave(16, ilRow) = Trim$(str$(tmSbf.iCalCarryOver))
                                End If
                                If (igPostType <> 5) And (igPostType <> 6) Then
                                    smSave(6, ilRow) = ""   'gLongToStrDec(tmSbf.lGross, 2)
                                End If
                                smSave(7, ilRow) = ""   'Trim$(Str$(tmSbf.iBonusNoSpots))
                            Else
                                If tmSbf.sPostStatus = "P" Then
                                    smSave(5, ilRow) = gAddStr(smSave(5, ilRow), Trim$(str$(tmSbf.iAirNoSpots)))
                                    smSave(16, ilRow) = gAddStr(smSave(16, ilRow), Trim$(str$(tmSbf.iCalCarryOver)))
                                End If
                                If (igPostType <> 5) And (igPostType <> 6) Then
                                    smSave(6, ilRow) = ""   'gAddStr(smSave(6, ilRow), gLongToStrDec(tmSbf.lGross, 2))
                                End If
                                smSave(7, ilRow) = ""   'gAddStr(smSave(7, ilRow), Trim$(Str$(tmSbf.iBonusNoSpots)))
                            End If
                            If slSource = "I" Then
                                smSave(15, ilRow) = ""
                            Else
                                smSave(15, ilRow) = smSave(5, ilRow)
                            End If
                            smSave(8, ilRow) = tmSbf.sBilled
                            smSave(17, ilRow) = tmSbf.sBarterPaid
                            lmSave(5, ilRow) = tmSbf.lRefInvNo
                            If imANOSPOTSINDEX > 0 Then
                                gSetShow pbcPostRep, smSave(5, ilRow), tmCtrls(imANOSPOTSINDEX)
                                smShow(imANOSPOTSINDEX, ilRow) = tmCtrls(imANOSPOTSINDEX).sShow
                            End If
                            If imANEXTNOSPOTSINDEX > 0 Then
                                gSetShow pbcPostRep, smSave(16, ilRow), tmCtrls(imANEXTNOSPOTSINDEX)
                                smShow(imANEXTNOSPOTSINDEX, ilRow) = tmCtrls(imANEXTNOSPOTSINDEX).sShow
                            End If
                            If (tmSbf.sPostStatus = "R") Or (tmSbf.sPostStatus = "P") Then
                                imSave(5, ilRow) = True
                            Else
                                imSave(5, ilRow) = False
                            End If
                            If (igPostType = 5) Or (igPostType = 6) Then
                                imSave(8, ilRow) = imSave(8, ilRow) + tmSbf.iBonusNoSpots
                                imSave(9, ilRow) = imSave(9, ilRow) + tmSbf.iCalCarryBonus
                                If tmSbf.iAirNoSpots > 0 Then
                                    smSave(6, ilRow) = gLongToStrDec(tmSbf.lGross + gStrDecToLong(smSave(6, ilRow), 2), 2)
                                End If
                            End If
                            gUnpackDate tmSbf.iRecDate(0), tmSbf.iRecDate(1), smSave(11, ilRow)
                            gUnpackDate tmSbf.iPostDate(0), tmSbf.iPostDate(1), smSave(12, ilRow)
                            If imRECEIVEDINDEX > 0 Then
                                gSetShow pbcPostRep, smSave(11, ilRow), tmCtrls(imRECEIVEDINDEX)
                                smShow(imRECEIVEDINDEX, ilRow) = tmCtrls(imRECEIVEDINDEX).sShow
                            End If
                            If imPOSTEDINDEX > 0 Then
                                gSetShow pbcPostRep, smSave(12, ilRow), tmCtrls(imPOSTEDINDEX)
                                smShow(imPOSTEDINDEX, ilRow) = tmCtrls(imPOSTEDINDEX).sShow
                            End If
                            'Get Weekly vaules
                            If tgSpf.sPostCalAff = "W" Then
                                tmRwfSrchKey1.lSbfCode = tmSbf.lCode
                                ilRet = btrGetEqual(hmRwf, tmRwf, imRwfRecLen, tmRwfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    For ilWk = 1 To 5 Step 1
                                        smSave(18 + ilWk - 1, ilRow) = Trim$(str$(tmRwf.iWkNoSpots(ilWk - 1)))
                                        smSave(23 + ilWk - 1, ilRow) = Trim$(str$(tmRwf.iWkNoCarried(ilWk - 1)))
                                    Next ilWk
                                End If
                            End If
                            'gSetShow pbcPostRep, smSave(6, ilRow), tmCtrls(imAGROSSINDEX)
                            'smShow(imAGROSSINDEX, ilRow) = tmCtrls(imAGROSSINDEX).sShow
                            'gSetShow pbcPostRep, smSave(7, ilRow), tmCtrls(imABONUSINDEX)
                            'smShow(imABONUSINDEX, ilRow) = tmCtrls(imABONUSINDEX).sShow
                            ''If tmSbf.lCode = 0 Then
                            ''    smInfo(1, ilRow) = "I"
                            ''Else
                            ''    smInfo(1, ilRow) = "F"
                            ''End If
                            If slSource = "F" Then
                                If tmSbf.sInserted = "Y" Then
                                    smInfo(1, ilRow) = "S"
                                Else
                                    smInfo(1, ilRow) = slSource
                                End If
                            Else
                                smInfo(1, ilRow) = slSource
                            End If
                            gUnpackDate tmSbf.iExportDate(0), tmSbf.iExportDate(1), smInfo(2, ilRow)
                            gUnpackDate tmSbf.iImportDate(0), tmSbf.iImportDate(1), smInfo(3, ilRow)
                            gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), smInfo(4, ilRow)
                            smInfo(5, ilRow) = Trim$(str$(tmSbf.iCombineID))
                            smInfo(6, ilRow) = Trim$(str$(tmSbf.lRefInvNo))
                            '12/17/06-Change to tax by agency or vehicle
                            'smInfo(7, ilRow) = gLongToStrDec(tmSbf.lTax1, 2)
                            'smInfo(8, ilRow) = gLongToStrDec(tmSbf.lTax2, 2)
                            smInfo(9, ilRow) = Trim$(str$(tmSbf.iNoItems))
                            smInfo(10, ilRow) = gLongToStrDec(tmSbf.lOGross, 2)
                            smInfo(11, ilRow) = gIntToStrDec(tmSbf.iCommPct, 2)
                            gUnpackDate tmSbf.iPrintInvDate(0), tmSbf.iPrintInvDate(1), smInfo(12, ilRow)
                            Exit For
                        End If
                    End If
                End If
            Next ilRow
            If ilFound Then
                Exit For
            End If
        Next ilPass
    Else
        ilFound = False
    End If
    If Not ilFound Then
        'Add in above total record for contract
        ilSPass = 0
        If (igPostType <> 2) Then
            ilEPass = 2
        Else
            'If by advertiser, don't add sbf unless contract matched
            If slSource = "F" Then
                ilEPass = 1
            Else
                ilEPass = 2
            End If
        End If
        ilFound = False
        For ilPass = ilSPass To ilEPass Step 1
            For ilRow = LBONE To UBound(smSave, 2) - 1 Step 1
                If ((tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.sCashTrade = smSave(1, ilRow)) And (ilPass = 0)) Or ((tmSbf.lChfCode = lmSave(1, ilRow)) And (ilPass = 1)) Or ((ilRow = UBound(smSave, 2) - 2) And (ilPass = 2)) Then
                    If (((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) And (imSave(6, ilRow) = tmSbf.iSpotLen)) Or ((Asc(tgSpf.sUsingFeatures2) And BARTER) <> BARTER) Then
                        ilFound = True
                        'Continue search until total record
                        For ilLoop = ilRow + 1 To UBound(smSave, 2) - 1 Step 1
                            If (InStr(1, smShow(imTOTALINDEX, ilLoop), "Total:", 1) > 0) Or (lmSave(1, ilRow) <> lmSave(1, ilLoop)) Then
                                If ilPass = 0 Then
                                    ilNewRow = ilLoop
                                Else
                                    ilNewRow = ilLoop + 1
                                    If InStr(1, smShow(imTOTALINDEX, ilLoop), "Grand Total:", 1) > 0 Then
                                        ilNewRow = ilNewRow - 1
                                    End If
                                End If
                                'Move all records from and including ilLoop dowm one
                                For ilIndex = UBound(smSave, 2) To ilNewRow Step -1
                                    For ilCol = LBONE To UBound(smSave, 1) Step 1
                                        smSave(ilCol, ilIndex) = smSave(ilCol, ilIndex - 1)
                                    Next ilCol
                                    For ilCol = LBONE To UBound(imSave, 1) Step 1
                                        imSave(ilCol, ilIndex) = imSave(ilCol, ilIndex - 1)
                                    Next ilCol
                                    For ilCol = LBONE To UBound(lmSave, 1) Step 1
                                        lmSave(ilCol, ilIndex) = lmSave(ilCol, ilIndex - 1)
                                    Next ilCol
                                    For ilCol = LBONE To UBound(smShow, 1) Step 1
                                        smShow(ilCol, ilIndex) = smShow(ilCol, ilIndex - 1)
                                    Next ilCol
                                    For ilCol = LBONE To UBound(smInfo, 1) Step 1
                                        smInfo(ilCol, ilIndex) = smInfo(ilCol, ilIndex - 1)
                                    Next ilCol
                                Next ilIndex
                                'Add row
                                If (ilPass = 0) Or (igPostType = 1) Or (igPostType = 3) Then
                                    ReDim Preserve smSave(0 To 33, 0 To UBound(smSave, 2) + 1) As String
                                    ReDim Preserve imSave(0 To 10, 0 To UBound(imSave, 2) + 1) As Integer
                                    ReDim Preserve lmSave(0 To 5, 0 To UBound(lmSave, 2) + 1) As Long
                                    ReDim Preserve smShow(0 To 10, 0 To UBound(smShow, 2) + 1) As String * 40
                                    ReDim Preserve smInfo(0 To 13, 0 To UBound(smInfo, 2) + 1) As String * 12
                                Else
                                    ReDim Preserve smSave(0 To 33, 0 To UBound(smSave, 2) + 2) As String
                                    ReDim Preserve imSave(0 To 10, 0 To UBound(imSave, 2) + 2) As Integer
                                    ReDim Preserve lmSave(0 To 5, 0 To UBound(lmSave, 2) + 2) As Long
                                    ReDim Preserve smShow(0 To 10, 0 To UBound(smShow, 2) + 2) As String * 40
                                    ReDim Preserve smInfo(0 To 13, 0 To UBound(smInfo, 2) + 2) As String * 12
                                End If
                                If (ilPass = 0) Or (ilPass = 1) Then
                                    For ilCol = LBONE To UBound(smSave, 1) Step 1
                                        smSave(ilCol, ilNewRow) = smSave(ilCol, ilNewRow - 1)
                                    Next ilCol
                                    For ilCol = LBONE To UBound(imSave, 1) Step 1
                                        imSave(ilCol, ilNewRow) = imSave(ilCol, ilNewRow - 1)
                                    Next ilCol
                                    For ilCol = LBONE To UBound(lmSave, 1) Step 1
                                        lmSave(ilCol, ilNewRow) = lmSave(ilCol, ilNewRow - 1)
                                    Next ilCol
                                    For ilCol = LBONE To UBound(smShow, 1) Step 1
                                        smShow(ilCol, ilNewRow) = smShow(ilCol, ilNewRow - 1)
                                    Next ilCol
                                    For ilCol = LBONE To UBound(smInfo, 1) Step 1
                                        smInfo(ilCol, ilNewRow) = smInfo(ilCol, ilNewRow - 1)
                                    Next ilCol
                                Else
                                    For ilCol = LBONE To UBound(smSave, 1) Step 1
                                        smSave(ilCol, ilNewRow) = ""
                                    Next ilCol
                                    For ilCol = LBONE To UBound(imSave, 1) Step 1
                                        imSave(ilCol, ilNewRow) = 0
                                    Next ilCol
                                    For ilCol = LBONE To UBound(lmSave, 1) Step 1
                                        lmSave(ilCol, ilNewRow) = 0
                                    Next ilCol
                                    For ilCol = LBONE To UBound(smShow, 1) Step 1
                                        smShow(ilCol, ilNewRow) = ""
                                    Next ilCol
                                    For ilCol = LBONE To UBound(smInfo, 1) Step 1
                                        smInfo(ilCol, ilNewRow) = ""
                                    Next ilCol
                                End If
                                'Set values into new at ilLoop
                                smSave(1, ilNewRow) = tmSbf.sCashTrade
                                If ilPass = 0 Then
                                    smSave(2, ilNewRow) = smSave(2, ilLoop - 1)
                                End If
                                smSave(3, ilNewRow) = ""   'Trim$(Str$(tmSbf.iNoItems))
                                smSave(4, ilNewRow) = ""
                                smSave(5, ilNewRow) = Trim$(str$(tmSbf.iAirNoSpots))
                                smSave(16, ilNewRow) = Trim$(str$(tmSbf.iCalCarryOver))
                                If slSource = "I" Then
                                    smSave(15, ilNewRow) = ""
                                Else
                                    smSave(15, ilNewRow) = smSave(5, ilNewRow)
                                End If
                                smSave(6, ilNewRow) = gLongToStrDec(tmSbf.lGross, 2)
                                smSave(7, ilNewRow) = ""    'Trim$(Str$(tmSbf.iBonusNoSpots))
                                smSave(8, ilNewRow) = tmSbf.sBilled
                                smSave(17, ilNewRow) = tmSbf.sBarterPaid
                                lmSave(5, ilNewRow) = tmSbf.lRefInvNo
                                If (igPostType = 5) Or (igPostType = 6) Then
                                    imSave(8, ilNewRow) = tmSbf.iBonusNoSpots
                                    imSave(9, ilNewRow) = tmSbf.iCalCarryBonus
                                End If
                                tmChfSrchKey.lCode = tmSbf.lChfCode
                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    imSave(1, ilNewRow) = tmChf.iAdfCode
                                    slStr = Trim$(str$(tmChf.lCntrNo))
                                    gSetShow pbcPostRep, slStr, tmCtrls(imCONTRACTINDEX)
                                    smShow(imCONTRACTINDEX, ilNewRow) = tmCtrls(imCONTRACTINDEX).sShow
                                    If ilPass <> 0 Then
                                        smSave(2, ilNewRow) = gIntToStrDec(tmChf.iPctTrade, 0)
                                    End If
                                Else
                                    imSave(1, ilNewRow) = -1
                                    slStr = "Missing:" & Trim$(str$(tmSbf.lChfCode))
                                    gSetShow pbcPostRep, slStr, tmCtrls(imCONTRACTINDEX)
                                    smShow(imCONTRACTINDEX, ilNewRow) = tmCtrls(imCONTRACTINDEX).sShow
                                    If ilPass <> 0 Then
                                        smSave(2, ilNewRow) = "0"
                                    End If
                                End If
                                imSave(2, ilNewRow) = tmSbf.iAirVefCode
                                imSave(3, ilNewRow) = 0
                                If igPostType = 2 Then
                                    For ilCol = LBound(tmVehicle) To UBound(tmVehicle) - 1 Step 1
                                        slNameCode = tmVehicle(ilCol).sKey
                                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                        If imSave(2, ilNewRow) = Val(slCode) Then
                                            imSave(3, ilNewRow) = ilCol
                                            Exit For
                                        End If
                                    Next ilCol
                                End If
                                lmSave(1, ilNewRow) = tmSbf.lChfCode
                                'lmSave(2, ilNewRow) = tmSbf.lCode
                                If slSource = "F" Then
                                    lmSave(2, ilNewRow) = tmSbf.lCode
                                Else
                                    lmSave(2, ilNewRow) = 0
                                End If
                                lmSave(3, ilNewRow) = tmSbf.lSpotPrice
                                If (tmSbf.sPostStatus = "R") Or (tmSbf.sPostStatus = "P") Then
                                    imSave(5, ilRow) = True
                                Else
                                    imSave(5, ilRow) = False
                                End If
                                gUnpackDate tmSbf.iRecDate(0), tmSbf.iRecDate(1), smSave(11, ilRow)
                                gUnpackDate tmSbf.iPostDate(0), tmSbf.iPostDate(1), smSave(12, ilRow)

                                slStr = smSave(1, ilNewRow)
                                If imCashTradeIndex > 0 Then
                                    gSetShow pbcPostRep, slStr, tmCtrls(imCashTradeIndex)
                                    smShow(imCashTradeIndex, ilNewRow) = tmCtrls(imCashTradeIndex).sShow
                                End If
                                'Advertiser
                                slStr = "Missing"
                                'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                                '    If tgCommAdf(ilAdf).iCode = imSave(1, ilNewRow) Then
                                    ilAdf = gBinarySearchAdf(imSave(1, ilNewRow))
                                    If ilAdf <> -1 Then
                                        'slStr = Trim$(tgCommAdf(ilAdf).sName)
                                        If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
                                            slStr = Trim$(tgCommAdf(ilAdf).sName) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID)
                                        Else
                                            slStr = Trim$(tgCommAdf(ilAdf).sName)
                                        End If
                                '        Exit For
                                    End If
                                'Next ilAdf
                                If imAdvtIndex > 0 Then
                                    gSetShow pbcPostRep, slStr, tmCtrls(imAdvtIndex)
                                    smShow(imAdvtIndex, ilNewRow) = tmCtrls(imAdvtIndex).sShow
                                End If
                                'Vehicle
                                slStr = "Missing"
                                'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                '    If tgMVef(ilVef).iCode = imSave(2, ilNewRow) Then
                                    ilVef = gBinarySearchVef(imSave(2, ilNewRow))
                                    If ilVef <> -1 Then
                                        slStr = Trim$(tgMVef(ilVef).sName)
                                '        Exit For
                                    End If
                                'Next ilVef
                                If imVEHICLEINDEX > 0 Then
                                    gSetShow pbcPostRep, slStr, tmCtrls(imVEHICLEINDEX)
                                    smShow(imVEHICLEINDEX, ilNewRow) = tmCtrls(imVEHICLEINDEX).sShow
                                End If
                                gSetShow pbcPostRep, smSave(3, ilNewRow), tmCtrls(imONOSPOTSINDEX)
                                smShow(imONOSPOTSINDEX, ilNewRow) = tmCtrls(imONOSPOTSINDEX).sShow
                                If imOPRICEINDEX > 0 Then
                                    slSpotRate = gLongToStrDec(lmSave(3, ilNewRow), 2)
                                    gSetShow pbcPostRep, slSpotRate, tmCtrls(imOPRICEINDEX)
                                    smShow(imOPRICEINDEX, ilNewRow) = tmCtrls(imOPRICEINDEX).sShow
                                End If
                                gSetShow pbcPostRep, smSave(4, ilNewRow), tmCtrls(imOGROSSINDEX)
                                smShow(imOGROSSINDEX, ilNewRow) = tmCtrls(imOGROSSINDEX).sShow
                                If imANOSPOTSINDEX > 0 Then
                                    gSetShow pbcPostRep, smSave(5, ilNewRow), tmCtrls(imANOSPOTSINDEX)
                                    smShow(imANOSPOTSINDEX, ilNewRow) = tmCtrls(imANOSPOTSINDEX).sShow
                                End If
                                If imANEXTNOSPOTSINDEX > 0 Then
                                    gSetShow pbcPostRep, smSave(16, ilNewRow), tmCtrls(imANEXTNOSPOTSINDEX)
                                    smShow(imANEXTNOSPOTSINDEX, ilNewRow) = tmCtrls(imANEXTNOSPOTSINDEX).sShow
                                End If
                                If imRECEIVEDINDEX > 0 Then
                                    gSetShow pbcPostRep, smSave(11, ilNewRow), tmCtrls(imRECEIVEDINDEX)
                                    smShow(imRECEIVEDINDEX, ilNewRow) = tmCtrls(imRECEIVEDINDEX).sShow
                                End If
                                If imPOSTEDINDEX > 0 Then
                                    gSetShow pbcPostRep, smSave(12, ilNewRow), tmCtrls(imPOSTEDINDEX)
                                    smShow(imPOSTEDINDEX, ilNewRow) = tmCtrls(imPOSTEDINDEX).sShow
                                End If
                                'gSetShow pbcPostRep, smSave(6, ilNewRow), tmCtrls(imAGROSSINDEX)
                                'smShow(imAGROSSINDEX, ilNewRow) = tmCtrls(imAGROSSINDEX).sShow
                                'gSetShow pbcPostRep, smSave(7, ilNewRow), tmCtrls(imABONUSINDEX)
                                'smShow(imABONUSINDEX, ilNewRow) = tmCtrls(imABONUSINDEX).sShow
                                imSave(4, ilNewRow) = False
                                For ilVef = 0 To UBound(imMktVefCode) - 1 Step 1
                                    If imSave(2, ilNewRow) = imMktVefCode(ilVef) Then
                                        imSave(4, ilNewRow) = True
                                        Exit For
                                    End If
                                Next ilVef
                                If (igPostType = 4) Or (igPostType = 2) Then
                                    If ilPass <> 0 Then
                                        'For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                                        '    smShow(ilCol, ilNewRow + 1) = ""
                                        'Next ilCol
                                        'For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                                        '    smInfo(ilCol, ilNewRow + 1) = ""
                                        'Next ilCol
                                        For ilCol = LBONE To UBound(smSave, 1) Step 1
                                            smSave(ilCol, ilNewRow + 1) = ""
                                        Next ilCol
                                        For ilCol = LBONE To UBound(imSave, 1) Step 1
                                            imSave(ilCol, ilNewRow + 1) = 0
                                        Next ilCol
                                        For ilCol = LBONE To UBound(lmSave, 1) Step 1
                                            lmSave(ilCol, ilNewRow + 1) = 0
                                        Next ilCol
                                        For ilCol = LBONE To UBound(smShow, 1) Step 1
                                            smShow(ilCol, ilNewRow + 1) = ""
                                        Next ilCol
                                        For ilCol = LBONE To UBound(smInfo, 1) Step 1
                                            smInfo(ilCol, ilNewRow + 1) = ""
                                        Next ilCol
                                        smSave(1, ilNewRow + 1) = ""
                                        smSave(2, ilNewRow + 1) = ""
                                        smSave(3, ilNewRow + 1) = ""
                                        smSave(4, ilNewRow + 1) = ""
                                        smSave(5, ilNewRow + 1) = ""
                                        smSave(15, ilNewRow + 1) = ""
                                        smSave(6, ilNewRow + 1) = ""
                                        smSave(7, ilNewRow + 1) = ""
                                        smSave(8, ilNewRow + 1) = "N"
                                        smSave(9, ilNewRow + 1) = ""
                                        smSave(10, ilNewRow + 1) = ""
                                        smSave(16, ilNewRow + 1) = ""
                                        smSave(17, ilNewRow + 1) = "N"
                                        smSave(18, ilNewRow + 1) = ""
                                        smSave(19, ilNewRow + 1) = ""
                                        smSave(20, ilNewRow + 1) = ""
                                        smSave(21, ilNewRow + 1) = ""
                                        smSave(22, ilNewRow + 1) = ""
                                        smSave(23, ilNewRow + 1) = ""
                                        smSave(24, ilNewRow + 1) = ""
                                        smSave(25, ilNewRow + 1) = ""
                                        smSave(26, ilNewRow + 1) = ""
                                        smSave(27, ilNewRow + 1) = ""
                                        smSave(28, ilNewRow + 1) = ""
                                        smSave(29, ilNewRow + 1) = ""
                                        smSave(30, ilNewRow + 1) = ""
                                        smSave(31, ilNewRow + 1) = ""
                                        smSave(32, ilNewRow + 1) = ""
                                        smSave(33, ilNewRow + 1) = ""
                                        gSetShow pbcPostRep, smSave(3, ilNewRow + 1), tmCtrls(imONOSPOTSINDEX)
                                        smShow(imONOSPOTSINDEX, ilNewRow + 1) = tmCtrls(imONOSPOTSINDEX).sShow
                                        gSetShow pbcPostRep, smSave(4, ilNewRow + 1), tmCtrls(imOGROSSINDEX)
                                        smShow(imOGROSSINDEX, ilNewRow + 1) = tmCtrls(imOGROSSINDEX).sShow
                                        If imANOSPOTSINDEX > 0 Then
                                            gSetShow pbcPostRep, smSave(5, ilNewRow + 1), tmCtrls(imANOSPOTSINDEX)
                                            smShow(imANOSPOTSINDEX, ilNewRow + 1) = tmCtrls(imANOSPOTSINDEX).sShow
                                        End If
                                        If imANEXTNOSPOTSINDEX > 0 Then
                                            gSetShow pbcPostRep, smSave(16, ilNewRow + 1), tmCtrls(imANEXTNOSPOTSINDEX)
                                            smShow(imANEXTNOSPOTSINDEX, ilNewRow + 1) = tmCtrls(imANEXTNOSPOTSINDEX).sShow
                                        End If
                                        'gSetShow pbcPostRep, smSave(6, ilNewRow + 1), tmCtrls(imAGROSSINDEX)
                                        'smShow(imAGROSSINDEX, ilNewRow + 1) = tmCtrls(imAGROSSINDEX).sShow
                                        'gSetShow pbcPostRep, smSave(7, ilNewRow + 1), tmCtrls(imABONUSINDEX)
                                        'smShow(imABONUSINDEX, ilNewRow + 1) = tmCtrls(imABONUSINDEX).sShow
                                        imSave(2, ilNewRow + 1) = 0
                                        imSave(3, ilNewRow + 1) = -1
                                        imSave(4, ilNewRow + 1) = True
                                        If imSave(1, ilNewRow) = -1 Then
                                            slStr = "# Missing"
                                        Else
                                            slStr = "Total: " & Trim$(str$(tmChf.lCntrNo))
                                        End If
                                        gSetShow pbcPostRep, slStr, tmCtrls(imTOTALINDEX)
                                        smShow(imTOTALINDEX, ilNewRow + 1) = tmCtrls(imTOTALINDEX).sShow
                                        smInfo(1, ilNewRow + 1) = "T"
                                    End If
                                End If
                                ''If tmSbf.lCode = 0 Then
                                ''    smInfo(1, ilNewRow) = "I"
                                ''Else
                                ''    smInfo(1, ilNewRow) = "F"
                                ''End If
                                'smInfo(1, ilNewRow) = slSource
                                If slSource = "F" Then
                                    If tmSbf.sInserted = "Y" Then
                                        smInfo(1, ilNewRow) = "S"
                                    Else
                                        smInfo(1, ilNewRow) = slSource
                                    End If
                                Else
                                    smInfo(1, ilRow) = slSource
                                End If
                                gUnpackDate tmSbf.iExportDate(0), tmSbf.iExportDate(1), smInfo(2, ilNewRow)
                                gUnpackDate tmSbf.iImportDate(0), tmSbf.iImportDate(1), smInfo(3, ilNewRow)
                                gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), smInfo(4, ilNewRow)
                                smInfo(5, ilNewRow) = Trim$(str$(tmSbf.iCombineID))
                                smInfo(6, ilNewRow) = Trim$(str$(tmSbf.lRefInvNo))
                                '12/17/06-Change to tax by agency or vehicle
                                'smInfo(7, ilNewRow) = gLongToStrDec(tmSbf.lTax1, 2)
                                'smInfo(8, ilNewRow) = gLongToStrDec(tmSbf.lTax2, 2)
                                smInfo(9, ilNewRow) = Trim$(str$(tmSbf.iNoItems))
                                smInfo(10, ilNewRow) = gLongToStrDec(tmSbf.lOGross, 2)
                                smInfo(11, ilNewRow) = gIntToStrDec(tmSbf.iCommPct, 2)
                                gUnpackDate tmSbf.iPrintInvDate(0), tmSbf.iPrintInvDate(1), smInfo(12, ilNewRow)
                                Exit For
                            End If
                        Next ilLoop
                    End If
                End If
                If ilFound Then
                    Exit For
                End If
            Next ilRow
            If ilFound Then
                Exit For
            End If
        Next ilPass
    End If
    mMerge = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move controls values to record *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilRowNo As Integer)
'
'   mMoveCtrlToRec
'   Where:
'
    Dim ilRet As Integer

    tmSbf.lCode = lmSave(2, ilRowNo)
    tmSbf.lChfCode = lmSave(1, ilRowNo)
    If Trim$(smInfo(4, ilRowNo)) <> "" Then
        gPackDate Trim$(smInfo(4, ilRowNo)), tmSbf.iDate(0), tmSbf.iDate(1)
    Else
        tmChfSrchKey.lCode = tmSbf.lChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If tmChf.sBillCycle = "C" Then
                gPackDate smEndCal, tmSbf.iDate(0), tmSbf.iDate(1)
            Else
                gPackDate smEndStd, tmSbf.iDate(0), tmSbf.iDate(1)
            End If
        Else
            gPackDate smEndStd, tmSbf.iDate(0), tmSbf.iDate(1)
        End If
    End If
    If Trim$(smInfo(12, ilRowNo)) <> "" Then
        gPackDate Trim$(smInfo(12, ilRowNo)), tmSbf.iPrintInvDate(0), tmSbf.iPrintInvDate(1)
    Else
        gPackDate "", tmSbf.iPrintInvDate(0), tmSbf.iPrintInvDate(1)
    End If
    tmSbf.sTranType = "T"
    If Trim$(smInfo(1, ilRowNo)) = "S" Then
        tmSbf.sInserted = "Y"
    Else
        tmSbf.sInserted = "N"
    End If
    tmSbf.iBillVefCode = imSave(2, ilRowNo)
    tmSbf.iMnfItem = 0
    If Trim$(smInfo(9, ilRowNo)) <> "" Then
        tmSbf.iNoItems = Val(Trim$(smInfo(9, ilRowNo)))
    Else
        tmSbf.iNoItems = 0  'Val(smSave(3, ilRowNo))
    End If
    tmSbf.lGross = gStrDecToLong(smSave(6, ilRowNo), 2)
    'tmSbf.sUnitName = ""
    tmSbf.sDescr = ""
    tmSbf.sAgyComm = "Y"
    tmSbf.sSlsComm = "Y"
    '12/17/06-Change to tax by agency or vehicle
    'tmSbf.sSlsTax = "Y"
    tmSbf.iTrfCode = 0
    tmSbf.sBilled = smSave(8, ilRowNo)
    tmSbf.sBarterPaid = smSave(17, ilRowNo)
    tmSbf.sCashTrade = smSave(1, ilRowNo)
    tmSbf.iAirVefCode = imSave(2, ilRowNo)
    tmSbf.iAirNoSpots = Val(smSave(5, ilRowNo))
    If tgSpf.sPostCalAff = "C" Then
        tmSbf.iCalCarryOver = Val(smSave(16, ilRowNo))
    Else
        tmSbf.iCalCarryOver = 0
    End If
    tmSbf.iBonusNoSpots = 0 'Val(smSave(7, ilRowNo))
    '12/17/06-Change to tax by agency or vehicle
    'tmSbf.lTax1 = gStrDecToLong(Trim$(smInfo(7, ilRowNo)), 2)
    'tmSbf.lTax2 = gStrDecToLong(Trim$(smInfo(8, ilRowNo)), 2)
    If Trim$(smInfo(3, ilRowNo)) <> "" Then
        gPackDate Trim$(smInfo(3, ilRowNo)), tmSbf.iImportDate(0), tmSbf.iImportDate(1)
    Else
        gPackDate smNowDate, tmSbf.iImportDate(0), tmSbf.iImportDate(1)
    End If
    If Trim$(smInfo(2, ilRowNo)) <> "" Then
        gPackDate Trim$(smInfo(2, ilRowNo)), tmSbf.iExportDate(0), tmSbf.iExportDate(1)
    Else
        gPackDate smNowDate, tmSbf.iExportDate(0), tmSbf.iExportDate(1)
    End If
    If (igPostType = 1) Then
        tmSbf.lRefInvNo = Val(Trim$(lmSave(5, ilRowNo)))
    Else
        tmSbf.lRefInvNo = Val(Trim$(smInfo(6, ilRowNo)))
    End If
    tmSbf.iCombineID = Val(Trim$(smInfo(5, ilRowNo)))
    tmSbf.lOGross = gStrDecToLong(Trim$(smInfo(10, ilRowNo)), 2)
    tmSbf.iCommPct = gStrDecToInt(Trim$(smInfo(11, ilRowNo)), 2)
    If (igPostType = 3) Then    'Post received
        If imSave(5, ilRowNo) Then
            tmSbf.sPostStatus = "R"
            If Trim$(smSave(11, ilRowNo)) <> "" Then
                gPackDate Trim$(smSave(11, ilRowNo)), tmSbf.iRecDate(0), tmSbf.iRecDate(1)
            Else
                gPackDate smNowDate, tmSbf.iRecDate(0), tmSbf.iRecDate(1)
            End If
            If Trim$(smSave(12, ilRowNo)) <> "" Then
                gPackDate Trim$(smSave(12, ilRowNo)), tmSbf.iPostDate(0), tmSbf.iPostDate(1)
                tmSbf.sPostStatus = "P"
            Else
                gPackDate "", tmSbf.iPostDate(0), tmSbf.iPostDate(1)
            End If
        Else
            tmSbf.sPostStatus = ""
            gPackDate "", tmSbf.iRecDate(0), tmSbf.iRecDate(1)
            gPackDate "", tmSbf.iPostDate(0), tmSbf.iPostDate(1)
        End If
    Else
        If Trim$(smSave(5, ilRowNo)) <> "" Then
            If Trim$(smSave(5, ilRowNo)) <> Trim$(smSave(15, ilRowNo)) Then
                tmSbf.sPostStatus = "P"
                If Trim$(smSave(11, ilRowNo)) <> "" Then
                    gPackDate Trim$(smSave(11, ilRowNo)), tmSbf.iRecDate(0), tmSbf.iRecDate(1)
                Else
                    'Don't set received date if not not set this way it is know if record can be removed
                    'if spots set to blank
                    'gPackDate smNowDate, tmSbf.iRecDate(0), tmSbf.iRecDate(1)
                    gPackDate "", tmSbf.iRecDate(0), tmSbf.iRecDate(1)
                End If
                gPackDate smNowDate, tmSbf.iPostDate(0), tmSbf.iPostDate(1)
            Else
                'Record not altered
                If Trim$(smSave(11, ilRowNo)) <> "" Then
                    tmSbf.sPostStatus = "R"
                Else
                    tmSbf.sPostStatus = ""
                End If
                gPackDate Trim$(smSave(11, ilRowNo)), tmSbf.iRecDate(0), tmSbf.iRecDate(1)
                If Trim$(smSave(12, ilRowNo)) <> "" Then
                    tmSbf.sPostStatus = "P"
                End If
                gPackDate Trim$(smSave(12, ilRowNo)), tmSbf.iPostDate(0), tmSbf.iPostDate(1)
            End If
        Else
            tmSbf.sPostStatus = ""
            If Trim$(smSave(11, ilRowNo)) <> "" Then
                tmSbf.sPostStatus = "R"
                gPackDate Trim$(smSave(11, ilRowNo)), tmSbf.iRecDate(0), tmSbf.iRecDate(1)
            Else
                gPackDate "", tmSbf.iRecDate(0), tmSbf.iRecDate(1)
            End If
            'Remove posting status as spots removed
            'If Trim$(smSave(12, ilRowNo)) <> "" Then
            '    tmSbf.sPostStatus = "P"
            '    gPackDate Trim$(smSave(12, ilRowNo)), tmSbf.iPostDate(0), tmSbf.iPostDate(1)
            'Else
                gPackDate "", tmSbf.iPostDate(0), tmSbf.iPostDate(1)
            'End If
        End If
    End If
    smSave(15, ilRowNo) = smSave(5, ilRowNo)
    'Spot Price
    tmSbf.lSpotPrice = lmSave(3, ilRowNo)
    tmSbf.iMissCarryOver = Val(gAddStr(gAddStr(smSave(10, ilRowNo), smSave(14, ilRowNo)), smSave(13, ilRowNo)))
    If tmSbf.iMissCarryOver > 0 Then
        tmSbf.iMissCarryOver = 0
    End If
    If Asc(tgSpf.sUsingFeatures2) And BARTER = BARTER Then
        tmSbf.iSpotLen = imSave(6, ilRowNo)
    Else
        tmSbf.iSpotLen = 0
    End If
    tmSbf.lAcquisitionCost = lmSave(4, ilRowNo)
    tmSbf.iTrfCode = 0      'NTR Tax TrfCode
    If (igPostType <> 5) And (igPostType <> 6) Then
        tmSbf.iLineNo = 0
    End If
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mProcFlight                     *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Sub mObtainCntrInfo(llChfCode As Long)
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim ilPass As Integer
    Dim ilAdf As Integer
    Dim ilClf As Integer
    Dim ilVef As Integer
    Dim ilInsertLine As Integer
    Dim slSFlightDate As String
    Dim slEFlightDate As String
    Dim ilIncludeFlight As Integer
    Dim slTotalNoPerWk As String
    Dim slTotalRate As String
    Dim slCTotalNoPerWk As String
    Dim slCTotalRate As String
    Dim slPctTrade As String
    Dim ilAddTo As Integer
    Dim ilLoop As Integer
    Dim ilCff As Integer
    Dim ilStartRowNo As Integer
    Dim ilCol As Integer
    Dim ilNoPasses As Integer
    Dim slSpotRate As String
    Dim ilPos As Integer
    Dim slAdvtAbbr As String
    Dim slVehicle As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slDaypartTimes As String
    Dim slDaypartDays As String
    Dim ilDay As Integer
    Dim slGenARDate As String
    Dim blEDI As Boolean
    Dim ilAgf As Integer

    If igPostType = 4 Then
        ilNoPasses = 0  '1
    Else
        ilNoPasses = 0
    End If
    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llChfCode, False, tgChfRep, tgClfRep(), tgCffRep())
    If ilRet Then
        blEDI = False
        ilAdf = gBinarySearchAdf(tgChfRep.iAdfCode)
        If ilAdf <> -1 Then
            If (tgCommAdf(ilAdf).sBillAgyDir = "D") Then
                If tgCommAdf(ilAdf).iArfInvCode > 0 Then
                    blEDI = True
                End If
            Else
                ilAgf = gBinarySearchAgf(tgChfRep.iAgfCode)
                If ilAgf <> -1 Then
                    If tgCommAgf(ilAgf).iArfInvCode > 0 Then
                        blEDI = True
                    End If
                End If
            End If
        End If
        For ilPass = 0 To ilNoPasses Step 1
            ilStartRowNo = UBound(smSave, 2)
            ilInsertLine = True
            ilRowNo = UBound(smSave, 2)
            'Contract Number
            lmSave(1, ilRowNo) = tgChfRep.lCode
            lmSave(2, ilRowNo) = 0
            slStr = Trim$(str$(tgChfRep.lCntrNo))
            gSetShow pbcPostRep, slStr, tmCtrls(imCONTRACTINDEX)
            smShow(imCONTRACTINDEX, ilRowNo) = tmCtrls(imCONTRACTINDEX).sShow
            'Cash/Trade flag
            slPctTrade = gIntToStrDec(tgChfRep.iPctTrade, 0)
            If ilNoPasses = 1 Then
                If (ilPass = 0) And (tgChfRep.iPctTrade <> 100) Then
                    slStr = "C"
                ElseIf (ilPass = 1) And (tgChfRep.iPctTrade <> 0) Then
                    slStr = "T"
                Else
                    ilInsertLine = False
                End If
            Else
                If tgChfRep.iPctTrade = 0 Then
                    slStr = "C"
                ElseIf tgChfRep.iPctTrade = 100 Then
                    slStr = "T"
                Else
                    slStr = "B"
                End If
            End If
            If ilInsertLine Then
                smSave(1, ilRowNo) = slStr
                smSave(2, ilRowNo) = slPctTrade
                If imCashTradeIndex >= 0 Then
                    gSetShow pbcPostRep, slStr, tmCtrls(imCashTradeIndex)
                    smShow(imCashTradeIndex, ilRowNo) = tmCtrls(imCashTradeIndex).sShow
                End If
                'Advertiser
                slStr = "Missing"
                imSave(1, ilRowNo) = tgChfRep.iAdfCode
                'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                '    If tgCommAdf(ilAdf).iCode = tgChfRep.iAdfCode Then
                    ilAdf = gBinarySearchAdf(tgChfRep.iAdfCode)
                    If ilAdf <> -1 Then
                        'slStr = Trim$(tgCommAdf(ilAdf).sName)
                        If (igPostType <> 5) And (igPostType <> 6) Then
                            If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
                                slStr = Trim$(tgCommAdf(ilAdf).sName) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID)
                            Else
                                slStr = Trim$(tgCommAdf(ilAdf).sName)
                            End If
                            If tgCommAdf(ilAdf).sAllowRepMG = "N" Then
                                slStr = "*" & slStr
                            End If
                        Else
                            If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
                                slStr = Trim$(tgCommAdf(ilAdf).sAbbr) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID)
                            Else
                                slStr = Trim$(tgCommAdf(ilAdf).sAbbr)
                            End If
                            If tgCommAdf(ilAdf).sAllowRepMG = "N" Then
                                slStr = "*" & slStr
                            End If
                            slAdvtAbbr = slStr
                        End If
                '        Exit For
                    End If
                'Next ilAdf
                If imAdvtIndex > 0 Then
                    gSetShow pbcPostRep, slStr, tmCtrls(imAdvtIndex)
                    smShow(imAdvtIndex, ilRowNo) = tmCtrls(imAdvtIndex).sShow
                ElseIf imADVTVEHDAYPARTINDEX > 0 Then
                    'gSetShow pbcPostRep, slAdvtAbbr, tmCtrls(imADVTVEHDAYPARTINDEX)
                    'smShow(imADVTVEHDAYPARTINDEX, ilRowNo) = tmCtrls(imADVTVEHDAYPARTINDEX).sShow
                End If
                For ilClf = LBound(tgClfRep) To UBound(tgClfRep) - 1 Step 1
                    ilRowNo = UBound(smSave, 2)
                    ilInsertLine = False
                    For ilVef = LBound(imMktVefCode) To UBound(imMktVefCode) - 1 Step 1
                        If tgClfRep(ilClf).ClfRec.iVefCode = imMktVefCode(ilVef) Then
                            ilInsertLine = True
                            Exit For
                        End If
                    Next ilVef
                    If (ilInsertLine) And ((igPostType = 5) Or (igPostType = 6)) And (Not blEDI) Then
                        If (tgSpf.sPostCalAff <> "N") Then
                            ilVef = gBinarySearchVef(tgClfRep(ilClf).ClfRec.iVefCode)
                            If ilVef <> -1 Then
                                If tgMVef(ilVef).iNrfCode <= 0 Then
                                    ilInsertLine = False
                                End If
                            End If
                        End If
                    End If
                    If (ilInsertLine) And ((tgClfRep(ilClf).ClfRec.sType = "S") Or (tgClfRep(ilClf).ClfRec.sType = "H")) Then
                        slGenARDate = mGetGenARDate(tgChfRep.lCntrNo, tgClfRep(ilClf).ClfRec.iVefCode)
                        imSave(7, ilRowNo) = tgClfRep(ilClf).ClfRec.iLine
                        'Vehicle
                        slStr = "Missing"
                        imSave(2, ilRowNo) = tgClfRep(ilClf).ClfRec.iVefCode
                        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If tgMVef(ilVef).iCode = tgClfRep(ilClf).ClfRec.iVefCode Then
                            ilVef = gBinarySearchVef(tgClfRep(ilClf).ClfRec.iVefCode)
                            If ilVef <> -1 Then
                                slStr = Trim$(tgMVef(ilVef).sName)
                        '        Exit For
                            End If
                            slVehicle = slStr
                        'Next ilVef
                        slDaypartDays = String(7, " ")
                        If imVEHICLEINDEX > 0 Then
                            gSetShow pbcPostRep, slStr, tmCtrls(imVEHICLEINDEX)
                            smShow(imVEHICLEINDEX, ilRowNo) = tmCtrls(imVEHICLEINDEX).sShow
                        ElseIf imADVTVEHDAYPARTINDEX > 0 Then
                            'Get Daypart information
                            If ((tgClfRep(ilClf).ClfRec.iStartTime(0) <> 1) Or (tgClfRep(ilClf).ClfRec.iStartTime(1) <> 0)) Then
                                gUnpackTime tgClfRep(ilClf).ClfRec.iStartTime(0), tgClfRep(ilClf).ClfRec.iStartTime(1), "A", "1", slStartTime
                                gUnpackTime tgClfRep(ilClf).ClfRec.iEndTime(0), tgClfRep(ilClf).ClfRec.iEndTime(1), "A", "1", slEndTime
                                slDaypartTimes = slStartTime & "-" & slEndTime
                            Else
                                tmRdfSrchKey.iCode = tgClfRep(ilClf).ClfRec.iRdfCode  ' Daypart File Code
                                ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    'gUnpackTime tmRdf.iStartTime(0, 7), tmRdf.iStartTime(1, 7), "A", "1", slStartTime
                                    'gUnpackTime tmRdf.iEndTime(0, 7), tmRdf.iEndTime(1, 7), "A", "1", slEndTime
                                    gUnpackTime tmRdf.iStartTime(0, 6), tmRdf.iStartTime(1, 6), "A", "1", slStartTime
                                    gUnpackTime tmRdf.iEndTime(0, 6), tmRdf.iEndTime(1, 6), "A", "1", slEndTime
                                    slDaypartTimes = slStartTime & "-" & slEndTime
                                End If
                            End If
                            'slStr = Trim$(slAdvtAbbr) & ", " & slVehicle & ", " & slDaypartTimes
                            'gSetShow pbcPostRep, slStr, tmCtrls(imADVTVEHDAYPARTINDEX)
                            'smShow(imADVTVEHDAYPARTINDEX, ilRowNo) = tmCtrls(imADVTVEHDAYPARTINDEX).sShow
                        End If
                        imSave(4, ilRowNo) = False
                        For ilVef = 0 To UBound(imMktVefCode) - 1 Step 1
                            If imSave(2, ilRowNo) = imMktVefCode(ilVef) Then
                                imSave(4, ilRowNo) = True
                                Exit For
                            End If
                        Next ilVef
                        slTotalNoPerWk = "0"
                        slTotalRate = "0"
                        ''Test if vehicle match, later test if spot price match
                        'ilAddTo = False
                        'For ilLoop = ilStartRowNo To ilRowNo - 1 Step 1
                        '    If imSave(2, ilLoop) = imSave(2, ilRowNo) Then
                        '        ilAddTo = True
                        '        ilRowNo = ilLoop
                        '        slTotalNoPerWk = smSave(3, ilRowNo)
                        '        slTotalRate = smSave(4, ilRowNo)
                        '        Exit For
                        '    End If
                        'Next ilLoop
                        If ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Or (igPostType = 5) Or (igPostType = 6) Then
                            imSave(6, ilRowNo) = tgClfRep(ilClf).ClfRec.iLen
                        Else
                            imSave(6, ilRowNo) = 0
                        End If
                        lmSave(4, ilRowNo) = tgClfRep(ilClf).ClfRec.lAcquisitionCost
                        smSave(33, ilRowNo) = ""
                        tmCxfSrchKey.lCode = tgClfRep(ilClf).ClfRec.lCxfCode
                        If tmCxfSrchKey.lCode <> 0 Then
                            tmCxf.sComment = ""
                            imCxfRecLen = Len(tmCxf) '5027
                            ilRet = gCXFGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                slStr = gStripChr0(tmCxf.sComment)
                                'If tmCxf.iStrLen > 0 Then
                                If slStr <> "" Then
                                    smSave(33, ilRowNo) = Trim$(Left$(slStr, 100)) 'Trim$(Left$(tmCxf.sComment, 100))
                                    'Replace 0d 0a with blank
                                    ilPos = InStr(smSave(33, ilRowNo), sgCR & sgLF)
                                    Do While ilPos > 0
                                        Mid$(smSave(33, ilRowNo), ilPos, 2) = "  "
                                        ilPos = InStr(smSave(33, ilRowNo), sgCR & sgLF)
                                    Loop
                                End If
                            End If
                        End If
                        ilCff = tgClfRep(ilClf).iFirstCff
                        Do While ilCff <> -1
                            ilRowNo = UBound(smSave, 2)
                            slTotalNoPerWk = "0"
                            slTotalRate = "0"
                            'Test if vehicle match, later test if spot price match
                            ilAddTo = False
                            For ilLoop = ilStartRowNo To ilRowNo - 1 Step 1
                                'If imSave(2, ilLoop) = imSave(2, ilRowNo) Then
                                If (igPostType <> 5) And (igPostType <> 6) Then
                                    If imSave(2, ilLoop) = tgClfRep(ilClf).ClfRec.iVefCode Then
                                        If (lmSave(3, ilLoop) = tgCffRep(ilCff).CffRec.lActPrice) Then
                                            If (((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) And (imSave(6, ilLoop) = imSave(6, ilRowNo))) Or ((Asc(tgSpf.sUsingFeatures2) And BARTER) <> BARTER) Then
                                                ilAddTo = True
                                                ilRowNo = ilLoop
                                                slTotalNoPerWk = smSave(3, ilRowNo)
                                                slTotalRate = smSave(4, ilRowNo)
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Else
                                    If imSave(7, ilLoop) = tgClfRep(ilClf).ClfRec.iLine Then
                                        ilAddTo = True
                                        ilRowNo = ilLoop
                                        slTotalNoPerWk = smSave(3, ilRowNo)
                                        slTotalRate = smSave(4, ilRowNo)
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                            If Not ilAddTo Then
                                For ilLoop = 0 To 4 Step 1
                                    smSave(18 + ilLoop, ilRowNo) = ""
                                    smSave(23 + ilLoop, ilRowNo) = ""
                                    smSave(28 + ilLoop, ilRowNo) = ""
                                Next ilLoop
                            End If
                            gUnpackDate tgCffRep(ilCff).CffRec.iStartDate(0), tgCffRep(ilCff).CffRec.iStartDate(1), slSFlightDate
                            gUnpackDate tgCffRep(ilCff).CffRec.iEndDate(0), tgCffRep(ilCff).CffRec.iEndDate(1), slEFlightDate
                            ilIncludeFlight = True
                            If tgChfRep.sBillCycle = "C" Then
                                If (gDateValue(slSFlightDate) > lmEndCal) Or (gDateValue(slEFlightDate) < lmStartCal) Then
                                    ilIncludeFlight = False
                                End If
                            Else
                                If (gDateValue(slSFlightDate) > lmEndStd) Or (gDateValue(slEFlightDate) < lmStartStd) Then
                                    ilIncludeFlight = False
                                End If
                            End If
                            'Test if CBS
                            If gDateValue(slEFlightDate) < gDateValue(slSFlightDate) Then
                                ilIncludeFlight = False
                            End If
                            If ilIncludeFlight Then
                                For ilDay = 0 To 6 Step 1
                                    If tgCffRep(ilCff).CffRec.iDay(ilDay) > 0 Then
                                         If Trim$(Mid$(slDaypartDays, ilDay + 1, 1)) = "" Then
                                            Mid$(slDaypartDays, ilDay + 1, 1) = "Y"
                                        ElseIf Trim$(Mid$(slDaypartDays, ilDay + 1, 1)) = "N" Then
                                            Mid$(slDaypartDays, ilDay + 1, 1) = "M"
                                        End If
                                   Else
                                        If Trim$(Mid$(slDaypartDays, ilDay + 1, 1)) = "" Then
                                            Mid$(slDaypartDays, ilDay + 1, 1) = "N"
                                        ElseIf Trim$(Mid$(slDaypartDays, ilDay + 1, 1)) = "Y" Then
                                            Mid$(slDaypartDays, ilDay + 1, 1) = "M"
                                        End If
                                    End If
                                Next ilDay
                                If (imADVTVEHDAYPARTINDEX > 0) And (Trim$(slDaypartDays) <> "") Then
                                    slStr = Trim$(slAdvtAbbr) & ", " & slVehicle & ", " & slDaypartTimes & " " & slDaypartDays
                                    gSetShow pbcPostRep, slStr, tmCtrls(imADVTVEHDAYPARTINDEX)
                                    smShow(imADVTVEHDAYPARTINDEX, ilRowNo) = tmCtrls(imADVTVEHDAYPARTINDEX).sShow
                                End If
                                mProcFlight ilCff, slSFlightDate, slEFlightDate, ilPass, slPctTrade, slSpotRate, slTotalNoPerWk, slTotalRate, ilRowNo
                                If slTotalNoPerWk <> "0" Then
                                    smSave(3, ilRowNo) = slTotalNoPerWk
                                    smSave(4, ilRowNo) = slTotalRate
                                    If igPostType = 4 Then
                                        smSave(5, ilRowNo) = slTotalNoPerWk
                                        smSave(6, ilRowNo) = slTotalRate
                                        smSave(16, ilRowNo) = ""
                                    Else
                                        smSave(5, ilRowNo) = ""
                                        smSave(6, ilRowNo) = ""
                                        smSave(16, ilRowNo) = ""
                                    End If
                                    smSave(7, ilRowNo) = ""
                                    smSave(8, ilRowNo) = "N"
                                    smSave(17, ilRowNo) = "N"
                                    smSave(9, ilRowNo) = ""
                                    smSave(10, ilRowNo) = ""
                                    gSetShow pbcPostRep, slTotalNoPerWk, tmCtrls(imONOSPOTSINDEX)
                                    smShow(imONOSPOTSINDEX, ilRowNo) = tmCtrls(imONOSPOTSINDEX).sShow
                                    gSetShow pbcPostRep, slTotalRate, tmCtrls(imOGROSSINDEX)
                                    smShow(imOGROSSINDEX, ilRowNo) = tmCtrls(imOGROSSINDEX).sShow
                                    If (imANOSPOTSINDEX > 0) And (igPostType = 4) Then
                                        gSetShow pbcPostRep, slTotalNoPerWk, tmCtrls(imANOSPOTSINDEX)
                                        smShow(imANOSPOTSINDEX, ilRowNo) = tmCtrls(imANOSPOTSINDEX).sShow
                                    End If
                                    lmSave(3, ilRowNo) = tgCffRep(ilCff).CffRec.lActPrice
                                    If imOPRICEINDEX > 0 Then
                                        gSetShow pbcPostRep, slSpotRate, tmCtrls(imOPRICEINDEX)
                                        smShow(imOPRICEINDEX, ilRowNo) = tmCtrls(imOPRICEINDEX).sShow
                                    End If
                                    smSave(11, ilRowNo) = ""    'Received Date
                                    smSave(12, ilRowNo) = ""    'Posted Date
                                    smSave(15, ilRowNo) = ""    'Original aired spots
                                    If imRECEIVEDINDEX > 0 Then
                                        smShow(imRECEIVEDINDEX, ilRowNo) = ""
                                    End If
                                    If imPOSTEDINDEX > 0 Then
                                        smShow(imPOSTEDINDEX, ilRowNo) = ""
                                    End If
                                    'gSetShow pbcPostRep, slTotalRate, tmCtrls(imAGROSSINDEX)
                                    'smShow(imAGROSSINDEX, ilRowNo) = tmCtrls(imAGROSSINDEX).sShow
                                    If Not ilAddTo Then
                                        'smShow(imABONUSINDEX, ilRowNo) = ""
                                        'smShow(imDNOSPOTSINDEX, ilRowNo) = ""
                                        'smShow(imDGROSSINDEX, ilRowNo) = ""
                                        For ilCol = LBONE To UBound(smInfo, 1) Step 1
                                            smInfo(ilCol, ilRowNo) = ""
                                        Next ilCol
                                        smInfo(1, ilRowNo) = "C"
                                        smInfo(13, ilRowNo) = slGenARDate
                                        ReDim Preserve smSave(0 To 33, 0 To ilRowNo + 1) As String
                                        ReDim Preserve imSave(0 To 10, 0 To ilRowNo + 1) As Integer
                                        ReDim Preserve lmSave(0 To 5, 0 To ilRowNo + 1) As Long
                                        ReDim Preserve smShow(0 To 10, 0 To ilRowNo + 1) As String * 40
                                        ReDim Preserve smInfo(0 To 13, 0 To ilRowNo + 1) As String * 12
                                        'imSave(1, ilRowNo + 1) = imSave(1, ilRowNo)
                                        'imSave(2, ilRowNo + 1) = imSave(2, ilRowNo)
                                        'lmSave(1, ilRowNo + 1) = lmSave(1, ilRowNo)
                                        'lmSave(2, ilRowNo + 1) = lmSave(2, ilRowNo)
                                        'lmSave(3, ilRowNo + 1) = lmSave(3, ilRowNo)
                                        'smSave(1, ilRowNo + 1) = smSave(1, ilRowNo)
                                        'smSave(2, ilRowNo + 1) = smSave(2, ilRowNo)
                                        'smShow(imCONTRACTINDEX, ilRowNo + 1) = smShow(imCONTRACTINDEX, ilRowNo)
                                        'smShow(imCASHTRADEINDEX, ilRowNo + 1) = smShow(imCASHTRADEINDEX, ilRowNo)
                                        'smShow(imADVTINDEX, ilRowNo + 1) = smShow(imADVTINDEX, ilRowNo)
                                        For ilCol = LBONE To UBound(smSave, 1) Step 1
                                            smSave(ilCol, ilRowNo + 1) = smSave(ilCol, ilRowNo)
                                        Next ilCol
                                        For ilCol = LBONE To UBound(imSave, 1) Step 1
                                            imSave(ilCol, ilRowNo + 1) = imSave(ilCol, ilRowNo)
                                        Next ilCol
                                        For ilCol = LBONE To UBound(lmSave, 1) Step 1
                                            lmSave(ilCol, ilRowNo + 1) = lmSave(ilCol, ilRowNo)
                                        Next ilCol
                                        For ilCol = LBONE To UBound(smShow, 1) Step 1
                                            smShow(ilCol, ilRowNo + 1) = smShow(ilCol, ilRowNo)
                                        Next ilCol
                                        For ilCol = LBONE To UBound(smInfo, 1) Step 1
                                            smInfo(ilCol, ilRowNo + 1) = smInfo(ilCol, ilRowNo)
                                        Next ilCol
                                        ilRowNo = ilRowNo + 1
                                    End If
                                End If
                            End If
                            ilCff = tgCffRep(ilCff).iNextCff
                        Loop
                    End If
                Next ilClf
                If ((igPostType = 5) Or (igPostType = 6)) And (ilStartRowNo < UBound(smSave, 2)) Then
                    For ilLoop = ilStartRowNo To UBound(smSave, 2) - 1 Step 1
                        lmSave(3, ilLoop) = gStrDecToLong(gDivStr(smSave(4, ilLoop), smSave(3, ilLoop)), 2)
                    Next ilLoop
                End If
                If ilStartRowNo < UBound(smSave, 2) Then
                    'Add Total line
                    slCTotalNoPerWk = "0"
                    slCTotalRate = "0"
                    For ilLoop = ilStartRowNo To UBound(smSave, 2) - 1 Step 1
                        slCTotalNoPerWk = gAddStr(slCTotalNoPerWk, smSave(3, ilLoop))
                        If InStr(RTrim$(smSave(4, ilLoop)), ".") > 0 Then
                            slCTotalRate = gAddStr(slCTotalRate, smSave(4, ilLoop))
                        End If
                    Next ilLoop
                    ilRowNo = UBound(smSave, 2)
                    For ilCol = LBONE To UBound(smShow, 1) Step 1
                        smShow(ilCol, ilRowNo) = ""
                    Next ilCol
                    For ilCol = LBONE To UBound(smInfo, 1) Step 1
                        smInfo(ilCol, ilRowNo) = ""
                    Next ilCol
                    smSave(1, ilRowNo) = ""
                    smSave(2, ilRowNo) = ""
                    smSave(3, ilRowNo) = slCTotalNoPerWk
                    smSave(4, ilRowNo) = slCTotalRate
                    If igPostType = 4 Then
                        smSave(5, ilRowNo) = slCTotalNoPerWk
                        smSave(6, ilRowNo) = slCTotalRate
                        smSave(16, ilRowNo) = ""
                    Else
                        smSave(5, ilRowNo) = ""
                        smSave(6, ilRowNo) = ""
                        smSave(16, ilRowNo) = ""
                    End If
                    smSave(7, ilRowNo) = ""
                    smSave(8, ilRowNo) = "N"
                    smSave(9, ilRowNo) = ""
                    smSave(10, ilRowNo) = ""
                    smSave(17, ilRowNo) = "N"
                    gSetShow pbcPostRep, slCTotalNoPerWk, tmCtrls(imONOSPOTSINDEX)
                    smShow(imONOSPOTSINDEX, ilRowNo) = tmCtrls(imONOSPOTSINDEX).sShow
                    gSetShow pbcPostRep, slCTotalRate, tmCtrls(imOGROSSINDEX)
                    smShow(imOGROSSINDEX, ilRowNo) = tmCtrls(imOGROSSINDEX).sShow
                    If (imANOSPOTSINDEX > 0) And (igPostType = 4) Then
                        gSetShow pbcPostRep, slCTotalNoPerWk, tmCtrls(imANOSPOTSINDEX)
                        smShow(imANOSPOTSINDEX, ilRowNo) = tmCtrls(imANOSPOTSINDEX).sShow
                    End If
                    'gSetShow pbcPostRep, slCTotalRate, tmCtrls(imAGROSSINDEX)
                    'smShow(AGROSSINDEX, ilRowNo) = tmCtrls(imAGROSSINDEX).sShow
                    If imVEHICLEINDEX > 0 Then
                        slStr = Trim$(str$(UBound(smSave, 2) - ilStartRowNo))
                        gSetShow pbcPostRep, slStr, tmCtrls(imVEHICLEINDEX)
                        smShow(imVEHICLEINDEX, ilRowNo) = tmCtrls(imVEHICLEINDEX).sShow
                    End If
                    imSave(2, ilRowNo) = 0
                    imSave(3, ilRowNo) = -1
                    If (igPostType = 4) Or (igPostType = 2) Or (igPostType = 5) Or (igPostType = 6) Then
                        slStr = "Total: " & Trim$(str$(tgChfRep.lCntrNo))
                        gSetShow pbcPostRep, slStr, tmCtrls(imTOTALINDEX)
                        smShow(imTOTALINDEX, ilRowNo) = tmCtrls(imTOTALINDEX).sShow
                        smInfo(1, ilRowNo) = "T"
                        ReDim Preserve smSave(0 To 33, 0 To ilRowNo + 1) As String
                        ReDim Preserve imSave(0 To 10, 0 To ilRowNo + 1) As Integer
                        ReDim Preserve lmSave(0 To 5, 0 To ilRowNo + 1) As Long
                        ReDim Preserve smShow(0 To 10, 0 To ilRowNo + 1) As String * 40
                        ReDim Preserve smInfo(0 To 13, 0 To ilRowNo + 1) As String * 12
                        For ilCol = LBONE To UBound(smSave, 1) Step 1
                            smSave(ilCol, ilRowNo + 1) = ""
                        Next ilCol
                        For ilCol = LBONE To UBound(imSave, 1) Step 1
                            imSave(ilCol, ilRowNo + 1) = 0
                        Next ilCol
                        For ilCol = LBONE To UBound(lmSave, 1) Step 1
                            lmSave(ilCol, ilRowNo + 1) = 0
                        Next ilCol
                        For ilCol = LBONE To UBound(smShow, 1) Step 1
                            smShow(ilCol, ilRowNo + 1) = ""
                        Next ilCol
                        For ilCol = LBONE To UBound(smInfo, 1) Step 1
                            smInfo(ilCol, ilRowNo + 1) = ""
                        Next ilCol
                    End If
                End If
            End If
        Next ilPass
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:D. Smith       *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Function mOpenMsgFile(slMsgFile As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim ilRet As Integer
    Dim ilPos As Integer

    ilRet = 0
    'On Error GoTo mOpenMsgFileErr:
    ilPos = InStr(1, slMsgFile, ".")
    If ilPos > 0 Then
        'slToFile = Left$(slMsgFile, ilPos) & "T" & Mid$(slMsgFile, ilPos + 2)
        If InStr(1, slMsgFile, "Inv", 1) > 0 Then
            slToFile = Left$(slMsgFile, ilPos - 4) & Mid$(slMsgFile, ilPos + 1, 2) & ".Txt"
        Else
            slToFile = Left$(slMsgFile, ilPos - 4) & Mid$(slMsgFile, ilPos + 1, 2) & Mid$(slMsgFile, ilPos - 1, 1) & ".Txt"
        End If
    Else
        slToFile = slMsgFile & ".Txt"
    End If
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        Kill slToFile
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, ""
    slMsgFile = slToFile
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mProcFlight                     *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Sub mProcFlight(ilCff As Integer, slSFlightDate As String, slEFlightDate As String, ilPass As Integer, slPctTrade As String, slSpotRate As String, slTotalNoPerWk As String, slTotalRate As String, ilRowNo As Integer)
'
'   Where
'       ilCff(I)- Flight record index
'       slSFlightDate(I)- Flight Start date
'       slEFlightDate(I)- Flight End Date
'       slSpotRate(O)- Spot Rate
'       slTotalNoPerWk(O)- Running Total number of spots per week
'       slTotalRate(I/O)- Ordered Total $'s
'

    Dim llDate As Long
    Dim ilDay As Integer
    Dim llSDate As Long
    Dim slRate As String
    Dim ilWkNo As Integer

    'Get flight rate
    Select Case tgCffRep(ilCff).CffRec.sPriceType
        Case "T"    'True
            slRate = gLongToStrDec(tgCffRep(ilCff).CffRec.lActPrice, 2)
            'Remove separate records for trade and cash
            'If (ilPass = 0) And (Val(slPctTrade) <> 0) Then
            '    slRate = gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100")
            'ElseIf (ilPass = 1) And (Val(slPctTrade) <> 100) Then
            '    slRate = gSubStr(RTrim$(slRate), gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100"))
            'End If
        Case "N"    'No Charge
            slRate = "N/C"
        Case "M"    'MG Line
            slRate = "MG"
        Case "B"    'Bonus
            slRate = "Bonus"
        Case "S"    'Spinoff
            slRate = "Spinoff"
        Case "P"    'Package
            slRate = gLongToStrDec(tgCffRep(ilCff).CffRec.lActPrice, 2)
            'Remove separate records for trade and cash
            'If (ilPass = 0) And (Val(slPctTrade) <> 0) Then
            '    slRate = gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100")
            'ElseIf (ilPass = 1) And (Val(slPctTrade) <> 100) Then
            '    slRate = gSubStr(RTrim$(slRate), gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100"))
            'End If
        Case "R"    'Recapturable
            slRate = "Recapturable"
        Case "A"    'ADU
            slRate = "ADU"
    End Select
    'Ignore rates for all type except True and Package
    If (tgCffRep(ilCff).CffRec.sPriceType <> "T") And (tgCffRep(ilCff).CffRec.sPriceType <> "P") Then
        slRate = "0.00"
    End If
    slSpotRate = slRate
    If (tgCffRep(ilCff).CffRec.sDyWk <> "D") Then    'Weekly
        If tgChfRep.sBillCycle = "C" Then
            llDate = gDateValue(slSFlightDate)
            Do While llDate <= gDateValue(slEFlightDate)
                If llDate < lmStartCal Then
                    If llDate + 6 >= lmStartCal Then
                        slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRep(ilCff).CffRec.iXSpotsWk)))
                        If InStr(RTrim$(slRate), ".") > 0 Then
                            slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRep(ilCff).CffRec.iXSpotsWk))))
                        End If
                    End If
                ElseIf (llDate <= lmEndCal) Then
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRep(ilCff).CffRec.iSpotsWk)))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRep(ilCff).CffRec.iSpotsWk))))
                    End If
                Else
                    Exit Do
                End If
                llDate = gDateValue(gObtainNextMonday(Format$(llDate + 1, "m/d/yy")))
            Loop
        Else
            llDate = gDateValue(slSFlightDate)
            Do While llDate <= gDateValue(slEFlightDate)
                If (llDate >= lmStartStd) And (llDate <= lmEndStd) Then
                    If tgSpf.sPostCalAff = "W" Then
                        ilWkNo = (llDate - lmStartStd) \ 7
                        smSave(28 + ilWkNo, ilRowNo) = Trim$(str$(Val(smSave(28 + ilWkNo, ilRowNo)) + tgCffRep(ilCff).CffRec.iSpotsWk))
                    End If
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRep(ilCff).CffRec.iSpotsWk)))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRep(ilCff).CffRec.iSpotsWk))))
                    End If
                End If
                If llDate > lmEndStd Then
                    Exit Do
                End If
                llDate = gDateValue(gObtainNextMonday(Format$(llDate + 1, "m/d/yy")))
            Loop
        End If
    Else    'Daily
        If tgChfRep.sBillCycle = "C" Then
            If gDateValue(slSFlightDate) >= lmStartCal Then
                llSDate = gDateValue(slSFlightDate)
            Else
                llSDate = lmStartCal
            End If
            For llDate = llSDate To gDateValue(slEFlightDate) Step 1
                If (llDate >= lmStartCal) And (llDate <= lmEndCal) Then
                    ilDay = gWeekDayLong(llDate)
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRep(ilCff).CffRec.iDay(ilDay))))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRep(ilCff).CffRec.iDay(ilDay)))))
                    End If
                End If
                If llDate >= lmEndCal Then
                    Exit For
                End If
            Next llDate
        Else
            If gDateValue(slSFlightDate) >= lmStartStd Then
                llSDate = gDateValue(slSFlightDate)
            Else
                llSDate = lmStartStd
            End If
            For llDate = llSDate To gDateValue(slEFlightDate) Step 1
                If (llDate >= lmStartStd) And (llDate <= lmEndStd) Then
                    ilDay = gWeekDayLong(llDate)
                    If tgSpf.sPostCalAff = "W" Then
                        ilWkNo = (llDate - lmStartStd) \ 7
                        smSave(28 + ilWkNo, ilRowNo) = Trim$(str$(Val(smSave(28 + ilWkNo, ilRowNo)) + tgCffRep(ilCff).CffRec.iDay(ilDay)))
                    End If
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRep(ilCff).CffRec.iDay(ilDay))))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRep(ilCff).CffRec.iDay(ilDay)))))
                    End If
                End If
                If llDate >= lmEndStd Then
                    Exit For
                End If
            Next llDate
        End If
    End If
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mReadImportFile                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Function mReadImportFile(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilLoop As Integer
    Dim ilMatch As Integer
    Dim ilError As Integer
    Dim ilErrorLogged As Integer
    Dim slMsg As String
    Dim ilPos As Integer
    Dim slVefName As String
    Dim slImptName As String
    Dim slVehicleName As String
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer

    ilRet = 0
    'On Error GoTo mReadImportFileErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        mReadImportFile = False
        Exit Function
    End If
    ilErrorLogged = False
    smNowDate = Format$(gNow(), "m/d/yy")
    Err.Clear
    Do
        ilRet = 0
        'On Error GoTo mReadImportFileErr:
        If EOF(hmFrom) Then
            Exit Do
        End If
        Line Input #hmFrom, slLine
        On Error GoTo 0
        ilRet = Err.Number
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Close hmFrom
            mReadImportFile = False
            Exit Function
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                gParseCDFields slLine, False, smFieldValues()    'Change case
                For ilLoop = UBound(smFieldValues) - 1 To LBound(smFieldValues) Step -1
                    smFieldValues(ilLoop + 1) = Trim$(smFieldValues(ilLoop))
                Next ilLoop
                smFieldValues(0) = ""
                If Trim$(smFieldValues(2)) <> "" Then
                    'Test if fields 5 and 6 are still enclosed in quotes- if so remove
                    ilPos = InStr(1, smFieldValues(5), """", 1)
                    If ilPos = 1 Then
                        smFieldValues(5) = right$(smFieldValues(5), Len(smFieldValues(5)) - 1)
                        smFieldValues(5) = Left$(smFieldValues(5), Len(smFieldValues(5)) - 1)
                    End If
                    ilPos = InStr(1, smFieldValues(6), """", 1)
                    If ilPos = 1 Then
                        smFieldValues(6) = right$(smFieldValues(6), Len(smFieldValues(6)) - 1)
                        smFieldValues(6) = Left$(smFieldValues(6), Len(smFieldValues(6)) - 1)
                    End If
                    ilPos = InStr(1, smFieldValues(19), """", 1)
                    If ilPos = 1 Then
                        smFieldValues(19) = right$(smFieldValues(19), Len(smFieldValues(19)) - 1)
                        smFieldValues(19) = Left$(smFieldValues(19), Len(smFieldValues(19)) - 1)
                    End If
                    'Make SBF record, then merge
                    ilError = False
                    slMsg = "Contract # " & smFieldValues(1)
                    'tmSbf.lCode = 0
                    tmChfSrchKey1.lCntrNo = Val(smFieldValues(1))
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = Val(smFieldValues(1))) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
                        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = Val(smFieldValues(1))) Then
                        tmSbf.lChfCode = tmChf.lCode
                    Else
                        ilError = True
                        slMsg = slMsg & " Missing"
                    End If
                        'Test that date date request date
                    If gDateValue(smFieldValues(2)) <> lmEndStd Then
                        slMsg = slMsg & ", Bill Date " & smFieldValues(2) & " not matching requested date"
                        ilError = True
                    End If
                    gPackDate smFieldValues(2), tmSbf.iDate(0), tmSbf.iDate(1)
                    tmSbf.sTranType = "T"
                    tmSbf.sPostStatus = "P"
                    ilMatch = False
                    'slVehicleName = smFieldValues(5)
                    slVehicleName = smFieldValues(19)   'Line vehicle
                    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(slVehicleName), 1) = 0 Then
                            tmSbf.iBillVefCode = tgMVef(ilLoop).iCode
                            ilMatch = True
                            Exit For
                        End If
                        If StrComp(mRemoveBlanks(tgMVef(ilLoop).sName), mRemoveBlanks(slVehicleName), 1) = 0 Then
                            tmSbf.iBillVefCode = tgMVef(ilLoop).iCode
                            ilMatch = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilMatch Then
                        For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            slVefName = mRemoveBlanks(tgMVef(ilLoop).sName)
                            slImptName = mRemoveBlanks(slVehicleName)
                            ilPos = InStr(1, slVefName, slImptName, vbTextCompare)
                            If ilPos > 0 Then
                                tmSbf.iBillVefCode = tgMVef(ilLoop).iCode
                                ilMatch = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If Not ilMatch Then
                        slImptName = mRemoveBlanks(slVehicleName)
                        ilPos1 = InStr(1, slImptName, "-", vbTextCompare)
                        If ilPos1 > 0 Then
                            ilPos2 = InStr(ilPos1 + 1, slImptName, "-", vbTextCompare)
                            If ilPos2 > 0 Then
                                slImptName = Mid(slImptName, ilPos1 + 1)
                                For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                    slVefName = mRemoveBlanks(tgMVef(ilLoop).sName)
                                    ilPos = InStr(1, slVefName, slImptName, vbTextCompare)
                                    If ilPos > 0 Then
                                        tmSbf.iBillVefCode = tgMVef(ilLoop).iCode
                                        ilMatch = True
                                        Exit For
                                    End If
                                Next ilLoop
                            End If
                        End If

                    End If
                    If Not ilMatch Then
                        slMsg = slMsg & ", " & slVehicleName & " Bill Vehicle Missing"
                        ilError = True
                    Else
                        ilMatch = False
                        For ilLoop = 0 To UBound(imMktVefCode) - 1 Step 1
                            If tmSbf.iBillVefCode = imMktVefCode(ilLoop) Then
                                ilMatch = True
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilMatch Then
                            slMsg = slMsg & ", " & slVehicleName & " Bill Vehicle not in Market"
                            ilError = True
                        End If
                    End If
                    tmSbf.iNoItems = Val(smFieldValues(11))
                    tmSbf.lGross = 0    'gStrDecToLong(smFieldValues(7), 2)
                    tmSbf.sBilled = "N"
                    tmSbf.sCashTrade = smFieldValues(4)
                    ilMatch = False
                    'slVehicleName = smFieldValues(6)
                    slVehicleName = smFieldValues(19)   'Line vehicle
                    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(slVehicleName), 1) = 0 Then
                            tmSbf.iAirVefCode = tgMVef(ilLoop).iCode
                            ilMatch = True
                            Exit For
                        End If
                        If StrComp(mRemoveBlanks(tgMVef(ilLoop).sName), mRemoveBlanks(slVehicleName), 1) = 0 Then
                            tmSbf.iAirVefCode = tgMVef(ilLoop).iCode
                            ilMatch = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilMatch Then
                        For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            slVefName = mRemoveBlanks(tgMVef(ilLoop).sName)
                            slImptName = mRemoveBlanks(slVehicleName)
                            ilPos = InStr(1, slVefName, slImptName, vbTextCompare)
                            If ilPos > 0 Then
                                tmSbf.iAirVefCode = tgMVef(ilLoop).iCode
                                ilMatch = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If Not ilMatch Then
                        slMsg = slMsg & ", " & slVehicleName & " Air Vehicle Missing"
                        ilError = True
                    Else
                        ilMatch = False
                        For ilLoop = 0 To UBound(imMktVefCode) - 1 Step 1
                            If tmSbf.iAirVefCode = imMktVefCode(ilLoop) Then
                                ilMatch = True
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilMatch Then
                            slMsg = slMsg & ", " & slVehicleName & " Air Vehicle not in Market"
                            ilError = True
                        End If
                    End If
                    tmSbf.iAirNoSpots = Val(smFieldValues(12)) + Val(smFieldValues(13))
                    tmSbf.iBonusNoSpots = 0 'Val(smFieldValues(13))
                    '12/17/06-Change to tax by agency or vehicle
                    'tmSbf.lTax1 = gStrDecToLong(smFieldValues(9), 2)
                    'tmSbf.lTax2 = gStrDecToLong(smFieldValues(10), 2)
                    gPackDate smNowDate, tmSbf.iImportDate(0), tmSbf.iImportDate(1)
                    gPackDate smFieldValues(15), tmSbf.iExportDate(0), tmSbf.iExportDate(1)
                    gPackDate "", tmSbf.iRecDate(0), tmSbf.iRecDate(1)
                    gPackDate "", tmSbf.iPostDate(0), tmSbf.iPostDate(1)
                    tmSbf.lRefInvNo = Val(smFieldValues(3))
                    tmSbf.iCombineID = Val(smFieldValues(14))
                    tmSbf.lOGross = gStrDecToLong(smFieldValues(16), 2)
                    tmSbf.iCommPct = gStrDecToInt(smFieldValues(17), 2)
                    tmSbf.lSpotPrice = gStrDecToLong(smFieldValues(18), 2)
                    If Not ilError Then
                        ilRet = mMerge("I")
                        If ilRet Then
                            imChg = True
                        End If
                    Else
                        Print #hmMsg, slMsg
                        ilErrorLogged = True
                    End If
                End If
            End If
        End If
    Loop Until ilEof
    Close hmFrom
    If ilErrorLogged Then
        mReadImportFile = False
    Else
        mReadImportFile = True
    End If
    mSetCommands
    Exit Function
'mReadImportFileErr:
'    ilRet = Err.Number
'    Resume Next

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadSbfRec                     *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read item bill records         *
'*                                                     *
'*******************************************************
Private Function mReadSbfRec(ilTestForSbf As Integer) As Integer
'
'   iRet = mReadSbfRec
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilVeh As Integer
    Dim llDate As Long

    If (imMarketIndex >= 0) And (imInvDateIndex >= 0) Then
        imSbfRecLen = Len(tmSbf)
        tmSbfSrchKey2.sTranType = "T"
        'tmSbfSrchKey2.iDate(0) = 0
        'tmSbfSrchKey2.iDate(1) = 0
        gPackDate smStartStd, tmSbfSrchKey2.iDate(0), tmSbfSrchKey2.iDate(1)
        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.sTranType = "T")
            If ((igPostType <> 5) And (igPostType <> 6) And (tmSbf.iLineNo = 0)) Or ((igPostType = 5) And (tmSbf.iLineNo <> 0)) Or ((igPostType = 6) And (tmSbf.iLineNo <> 0)) Then
                gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
                If llDate > lmEndStd Then
                    Exit Do
                End If
                For ilVeh = 0 To UBound(imMktVefCode) - 1 Step 1
                    If tmSbf.iAirVefCode = imMktVefCode(ilVeh) Then
                        If ilTestForSbf Then
                            mReadSbfRec = True
                            Exit Function
                        End If
                        imChg = False
                        ilRet = mMerge("F")
                        Exit For
                    End If
                Next ilVeh
            End If
            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    If ilTestForSbf Then
        mReadSbfRec = False
    Else
        mReadSbfRec = True
    End If
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mRecomputeTotals                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Recompute all totals           *
'*                                                     *
'*******************************************************
Sub mRecomputeTotals()
    Dim slONoSpots As String
    Dim slOGross As String
    Dim slANoSpots As String
    Dim slAGross As String
    Dim slMissedNoSpots As String
    Dim slMGNoSpots As String
    Dim slBonusNoSpots As String
    Dim slCalCarry As String
    Dim slPrior As String
    Dim slCalPrev As String
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim ilAdj As Integer
    Dim ilComputeMG As Integer
    Dim ilAdf As Integer
    Dim ilCalCarryBonus As Integer
    Dim ilCurrBonus As Integer
    Dim ilPriorBonus As Integer

    ilStartRow = LBONE  ' LBound(smSave, 2)
    ilEndRow = ilStartRow + 1
    If ilEndRow >= UBound(smSave, 2) Then
        Exit Sub
    End If
    If igPostType = 4 Then
        ilAdj = 1
    Else
        ilAdj = 0
    End If
    Do
        If (InStr(1, smShow(imTOTALINDEX, ilEndRow), "Total:", 1) > 0) Or (InStr(1, smShow(imTOTALINDEX, ilEndRow), "Grand Total:", 1) > 0) Then
            slONoSpots = "0"
            slOGross = "0.00"
            slANoSpots = "0"
            slAGross = "0.00"
            slMissedNoSpots = "0"
            slMGNoSpots = "0"
            slBonusNoSpots = "0"
            slCalCarry = "0"
            slCalPrev = "0"
            slPrior = "0"
            ilCalCarryBonus = 0
            ilCurrBonus = 0
            ilPriorBonus = 0
            For ilRow = ilStartRow To ilEndRow - 1 Step 1
                'Compute # Missed
                slStr = ""
                If Trim$(smInfo(1, ilRow)) <> "S" Then
                    If (smSave(5, ilRow) <> "") Then
                        slStr = smSave(5, ilRow)
                    Else
                        slStr = "0"
                    End If
                    If (tgSpf.sPostCalAff = "C") And (smSave(9, ilRow) <> "") Then
                        slStr = gAddStr(slStr, smSave(9, ilRow))
                    End If
                    If (slStr <> "") Then
                        slStr = gSubStr(slStr, smSave(3, ilRow))
                        If Val(slStr) >= 0 Then
                            slStr = ""
                        End If
                    End If
                End If
                smSave(14, ilRow) = slStr
                'Compute # MGs
                ilComputeMG = True
                'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                '    If tgCommAdf(ilAdf).iCode = imSave(1, ilRow) Then
                    ilAdf = gBinarySearchAdf(imSave(1, ilRow))
                    If ilAdf <> -1 Then
                        If tgCommAdf(ilAdf).sAllowRepMG = "N" Then
                            ilComputeMG = False
                        End If
                '        Exit For
                    End If
                'Next ilAdf

                slStr = ""
                If Trim$(smInfo(1, ilRow)) <> "S" Then
                    If (smSave(5, ilRow) <> "") Then
                        slStr = smSave(5, ilRow)
                    End If
                    If (tgSpf.sPostCalAff = "C") And (smSave(9, ilRow) <> "") Then
                        slStr = gAddStr(slStr, smSave(9, ilRow))
                    End If
                    If ilComputeMG Then
                        If (slStr <> "") Then
                            slStr = gSubStr(slStr, smSave(3, ilRow))
                            If Val(slStr) > 0 Then
                                If Val(smSave(10, ilRow)) < 0 Then
                                    If Val(slStr) > Abs(Val(smSave(10, ilRow))) Then
                                        slStr = str$(Abs(Val(smSave(10, ilRow))))
                                    End If
                                Else
                                    slStr = ""
                                End If
                            Else
                                slStr = ""
                            End If
                        End If
                    Else
                        slStr = ""
                    End If
                End If
                smSave(13, ilRow) = slStr
                'Compute Bonus
                slStr = ""
                If Trim$(smInfo(1, ilRow)) <> "S" Then
                    If (smSave(5, ilRow) <> "") Then
                        slStr = smSave(5, ilRow)
                    End If
                    If (tgSpf.sPostCalAff = "C") And (smSave(9, ilRow) <> "") Then
                        slStr = gAddStr(slStr, smSave(9, ilRow))
                    End If
                    If (slStr <> "") Then
                        slStr = gSubStr(slStr, smSave(3, ilRow))
                        If Val(slStr) > 0 Then
                            If Val(smSave(10, ilRow)) <= 0 Then
                                If ilComputeMG Then
                                    If Val(slStr) > Abs(Val(smSave(10, ilRow))) Then
                                        slStr = str$(Val(slStr) - Abs(Val(smSave(10, ilRow))))
                                    Else
                                        slStr = ""
                                    End If
                                Else
                                    slStr = slStr
                                End If
                            Else
                                slStr = ""
                            End If
                        Else
                            slStr = ""
                        End If
                    End If
                End If
                smSave(7, ilRow) = slStr
                'Compute Aired Gross
                If (igPostType <> 5) And (igPostType <> 6) Then
                    If Trim$(smInfo(1, ilRow)) = "S" Then
                        slStr = smSave(5, ilRow)
                        smSave(6, ilRow) = gLongToStrDec(Val(slStr) * lmSave(3, ilRow), 2)
                    Else
                        slStr = gAddStr(smSave(3, ilRow), smSave(14, ilRow))
                        slStr = gAddStr(slStr, smSave(13, ilRow))
                        smSave(6, ilRow) = gLongToStrDec(Val(slStr) * lmSave(3, ilRow), 2)
                    End If
                End If
                'Compute running values
                slONoSpots = gAddStr(slONoSpots, smSave(3, ilRow))
                slOGross = gAddStr(slOGross, smSave(4, ilRow))
                slANoSpots = gAddStr(slANoSpots, smSave(5, ilRow))
                slMissedNoSpots = gAddStr(slMissedNoSpots, smSave(14, ilRow))
                slMGNoSpots = gAddStr(slMGNoSpots, smSave(13, ilRow))
                slBonusNoSpots = gAddStr(slBonusNoSpots, smSave(7, ilRow))
                slAGross = gAddStr(slAGross, smSave(6, ilRow))
                slCalCarry = gAddStr(slCalCarry, smSave(16, ilRow))
                slCalPrev = gAddStr(slCalPrev, smSave(9, ilRow))
                slPrior = gAddStr(slPrior, smSave(10, ilRow))
                ilCalCarryBonus = ilCalCarryBonus + imSave(9, ilRow)
                ilCurrBonus = ilCurrBonus + imSave(8, ilRow)
                ilPriorBonus = ilPriorBonus + imSave(10, ilRow)
            Next ilRow
            smSave(3, ilEndRow) = slONoSpots
            smSave(4, ilEndRow) = slOGross
            smSave(5, ilEndRow) = slANoSpots
            smSave(6, ilEndRow) = slAGross
            smSave(14, ilEndRow) = slMissedNoSpots
            smSave(13, ilEndRow) = slMGNoSpots
            smSave(7, ilEndRow) = slBonusNoSpots
            smSave(16, ilEndRow) = slCalCarry
            smSave(9, ilEndRow) = slCalPrev
            smSave(10, ilEndRow) = slPrior
            imSave(8, ilEndRow) = ilCurrBonus
            imSave(9, ilEndRow) = ilCalCarryBonus
            imSave(10, ilEndRow) = ilPriorBonus
            gSetShow pbcPostRep, slONoSpots, tmCtrls(imONOSPOTSINDEX)
            smShow(imONOSPOTSINDEX, ilEndRow) = tmCtrls(imONOSPOTSINDEX).sShow
            gSetShow pbcPostRep, slOGross, tmCtrls(imOGROSSINDEX)
            smShow(imOGROSSINDEX, ilEndRow) = tmCtrls(imOGROSSINDEX).sShow
            If imANOSPOTSINDEX > 0 Then
                gSetShow pbcPostRep, slANoSpots, tmCtrls(imANOSPOTSINDEX)
                smShow(imANOSPOTSINDEX, ilEndRow) = tmCtrls(imANOSPOTSINDEX).sShow
            End If
            If imANEXTNOSPOTSINDEX > 0 Then
                gSetShow pbcPostRep, slCalCarry, tmCtrls(imANEXTNOSPOTSINDEX)
                smShow(imANEXTNOSPOTSINDEX, ilEndRow) = tmCtrls(imANEXTNOSPOTSINDEX).sShow
            End If
            If imAPREVNOSPOTSINDEX > 0 Then
                gSetShow pbcPostRep, slCalPrev, tmCtrls(imAPREVNOSPOTSINDEX)
                smShow(imAPREVNOSPOTSINDEX, ilEndRow) = tmCtrls(imAPREVNOSPOTSINDEX).sShow
            End If
            If imPRIORNOSPOTSINDEX > 0 Then
                gSetShow pbcPostRep, slPrior, tmCtrls(imPRIORNOSPOTSINDEX)
                smShow(imPRIORNOSPOTSINDEX, ilEndRow) = tmCtrls(imPRIORNOSPOTSINDEX).sShow
            End If

            'gSetShow pbcPostRep, slAGross, tmCtrls(imAGROSSINDEX)
            'smShow(imAGROSSINDEX, ilEndRow) = tmCtrls(imAGROSSINDEX).sShow
            'gSetShow pbcPostRep, slBonusNoSpots, tmCtrls(imABONUSINDEX)
            'smShow(imABONUSINDEX, ilEndRow) = tmCtrls(imABONUSINDEX).sShow
            If InStr(smShow(imTOTALINDEX, ilEndRow), ",") <= 0 Then
                If imVEHICLEINDEX > 0 Then
                    If igPostType = 2 Then
                        slStr = Trim$(smShow(imTOTALINDEX, ilEndRow)) '& ", " & Trim$(Str$(ilEndRow - ilStartRow))
                        gSetShow pbcPostRep, slStr, tmCtrls(imTOTALINDEX)
                        smShow(imTOTALINDEX, ilEndRow) = tmCtrls(imTOTALINDEX).sShow
                    Else
                        slStr = Trim$(str$(ilEndRow - ilStartRow))
                        gSetShow pbcPostRep, slStr, tmCtrls(imVEHICLEINDEX)
                        smShow(imVEHICLEINDEX, ilEndRow) = tmCtrls(imVEHICLEINDEX).sShow
                    End If
                Else
                    slStr = Trim$(smShow(imTOTALINDEX, ilEndRow)) '& ", " & Trim$(Str$(ilEndRow - ilStartRow))
                    gSetShow pbcPostRep, slStr, tmCtrls(imTOTALINDEX)
                    smShow(imTOTALINDEX, ilEndRow) = tmCtrls(imTOTALINDEX).sShow
                End If
            End If
            ilStartRow = ilEndRow + 1
            ilEndRow = ilStartRow + 1
        Else
            ilEndRow = ilEndRow + 1
        End If
    Loop While ilEndRow < UBound(smSave, 2) - ilAdj
    mGrandTotal
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mRemoveAirCount                 *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: If no sbf and no import set    *
'*                      set air counts to zero         *
'*                                                     *
'*******************************************************
Sub mRemoveAirCount(ilRemoveAll As Integer)
    Dim ilLoop As Integer

    For ilLoop = LBONE To UBound(smSave, 2) - 1 Step 1
        If ((lmSave(2, ilLoop) = 0) And (Trim$(smInfo(1, ilLoop)) = "C")) Or (ilRemoveAll) Then
            smSave(5, ilLoop) = "0"
            smSave(6, ilLoop) = "0.00"
            smSave(7, ilLoop) = "0"
            smSave(16, ilLoop) = "0"
            If imANOSPOTSINDEX > 0 Then
                gSetShow pbcPostRep, smSave(5, ilLoop), tmCtrls(imANOSPOTSINDEX)
                smShow(imANOSPOTSINDEX, ilLoop) = tmCtrls(imANOSPOTSINDEX).sShow
            End If
            If imANEXTNOSPOTSINDEX > 0 Then
                gSetShow pbcPostRep, smSave(16, ilLoop), tmCtrls(imANEXTNOSPOTSINDEX)
                smShow(imANEXTNOSPOTSINDEX, ilLoop) = tmCtrls(imANEXTNOSPOTSINDEX).sShow
            End If
            'gSetShow pbcPostRep, smSave(6, ilLoop), tmCtrls(imAGROSSINDEX)
            'smShow(imAGROSSINDEX, ilLoop) = tmCtrls(imAGROSSINDEX).sShow
            'gSetShow pbcPostRep, smSave(7, ilLoop), tmCtrls(imABONUSINDEX)
            'smShow(imABONUSINDEX, ilLoop) = tmCtrls(imABONUSINDEX).sShow
        End If
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
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
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim slMsg As String
    Dim ilWk As Integer
    Dim ilAnyWkPosted As Integer
    Dim tlSbf As SBF

    Dim tlSbf1 As MOVEREC
    Dim tlSbf2 As MOVEREC


    mWklySetShow imWklyBoxNo
    mSetShow imBoxNo
    If mTestFields() = NO Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    For ilLoop = LBONE To UBound(smSave, 2) - 1 Step 1
        If (smSave(8, ilLoop) <> "Y") And (InStr(1, smShow(imTOTALINDEX, ilLoop), "Total:", 1) = 0) And (smSave(17, ilLoop) <> "Y") Then
            mMoveCtrlToRec ilLoop
            If Trim$(tmSbf.sPostStatus) <> "" Then
                Do  'Loop until record updated or added
                    If tmSbf.lCode = 0 Then 'New selected
                        tmSbf.lCode = 0
                        tmSbf.iCalCarryBonus = 0
                        tmSbf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                        ilRet = btrInsert(hmSbf, tmSbf, imSbfRecLen, INDEXKEY0)
                        slMsg = "mSaveRec (btrInsert: Posting)"
                    Else 'Old record-Update
                        slMsg = "mSaveRec (btrGetEqual: Posting)"
                        tmSbfSrchKey1.lCode = tmSbf.lCode
                        ilRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, slMsg, PostRep
                        On Error GoTo 0
                        LSet tlSbf1 = tlSbf
                        LSet tlSbf2 = tmSbf
                        If StrComp(tlSbf1.sChar, tlSbf2.sChar, 0) <> 0 Then
                            Do
                                tmSbf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                                ilRet = btrUpdate(hmSbf, tmSbf, imSbfRecLen)
                                If ilRet = BTRV_ERR_CONFLICT Then
                                    tmSbfSrchKey1.lCode = tmSbf.lCode
                                    ilCRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            slMsg = "mSaveRec (btrUpdate: Posting)"
                        Else
                            ilRet = BTRV_ERR_NONE
                        End If
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, PostRep
                On Error GoTo 0
                lmSave(2, ilLoop) = tmSbf.lCode
                If tgSpf.sPostCalAff = "W" Then
                    tmRwfSrchKey1.lSbfCode = tmSbf.lCode
                    ilRet = btrGetEqual(hmRwf, tmRwf, imRwfRecLen, tmRwfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        Do
                            For ilWk = 0 To 4 Step 1
                                If Trim$(smSave(18 + ilWk, ilLoop)) <> "" Then
                                    'tmRwf.iWkNoSpots(ilWk + 1) = Val(smSave(18 + ilWk, ilLoop))
                                    tmRwf.iWkNoSpots(ilWk) = Val(smSave(18 + ilWk, ilLoop))
                                Else
                                    'tmRwf.iWkNoSpots(ilWk + 1) = 0
                                    tmRwf.iWkNoSpots(ilWk) = 0
                                End If
                                If Trim$(smSave(23 + ilWk, ilLoop)) <> "" Then
                                    'tmRwf.iWkNoCarried(ilWk + 1) = Val(smSave(23 + ilWk, ilLoop))
                                    tmRwf.iWkNoCarried(ilWk) = Val(smSave(23 + ilWk, ilLoop))
                                Else
                                    'tmRwf.iWkNoCarried(ilWk + 1) = 0
                                    tmRwf.iWkNoCarried(ilWk) = 0
                                End If
                                If Trim$(smSave(28 + ilWk, ilLoop)) <> "" Then
                                    'tmRwf.iWkOrderSpotNo(ilWk + 1) = Val(smSave(28 + ilWk, ilLoop))
                                    tmRwf.iWkOrderSpotNo(ilWk) = Val(smSave(28 + ilWk, ilLoop))
                                Else
                                    'tmRwf.iWkOrderSpotNo(ilWk + 1) = 0
                                    tmRwf.iWkOrderSpotNo(ilWk) = 0
                                End If
                           Next ilWk
                            ilRet = btrUpdate(hmRwf, tmRwf, imRwfRecLen)
                            If ilRet = BTRV_ERR_CONFLICT Then
                                tmRwfSrchKey1.lSbfCode = tmSbf.lCode
                                ilCRet = btrGetEqual(hmRwf, tmRwf, imRwfRecLen, tmRwfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        slMsg = "mSaveRec (btrUpdate: Posting by Week)"
                    Else
                        tmRwf.lCode = 0
                        tmRwf.lSbfCode = tmSbf.lCode
                        ilAnyWkPosted = False
                        For ilWk = 0 To 4 Step 1
                            If Trim$(smSave(18 + ilWk, ilLoop)) <> "" Then
                                ilAnyWkPosted = True
                                'tmRwf.iWkNoSpots(ilWk + 1) = Val(smSave(18 + ilWk, ilLoop))
                                tmRwf.iWkNoSpots(ilWk) = Val(smSave(18 + ilWk, ilLoop))
                            Else
                                'tmRwf.iWkNoSpots(ilWk + 1) = 0
                                tmRwf.iWkNoSpots(ilWk) = 0
                            End If
                            If Trim$(smSave(23 + ilWk, ilLoop)) <> "" Then
                                ilAnyWkPosted = True
                                'tmRwf.iWkNoCarried(ilWk + 1) = Val(smSave(23 + ilWk, ilLoop))
                                tmRwf.iWkNoCarried(ilWk) = Val(smSave(23 + ilWk, ilLoop))
                            Else
                                'tmRwf.iWkNoCarried(ilWk + 1) = 0
                                tmRwf.iWkNoCarried(ilWk) = 0
                            End If
                            If Trim$(smSave(28 + ilWk, ilLoop)) <> "" Then
                                'tmRwf.iWkOrderSpotNo(ilWk + 1) = Val(smSave(28 + ilWk, ilLoop))
                                tmRwf.iWkOrderSpotNo(ilWk) = Val(smSave(28 + ilWk, ilLoop))
                            Else
                                'tmRwf.iWkOrderSpotNo(ilWk + 1) = 0
                                tmRwf.iWkOrderSpotNo(ilWk) = 0
                            End If
                        Next ilWk
                        If ilAnyWkPosted Then
                            tmRwf.sUnused = ""
                            ilRet = btrInsert(hmRwf, tmRwf, imRwfRecLen, INDEXKEY0)
                            slMsg = "mSaveRec (btrInsert: Posting by Week)"
                        Else
                            ilRet = BTRV_ERR_NONE
                        End If
                    End If
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, PostRep
                    On Error GoTo 0
                End If
                mUpdateRvfRefInvNo tmSbf
            Else
                If tmSbf.lCode <> 0 Then 'New selected
                    slMsg = "mSaveRec (btrGetEqual: Posting)"
                    tmSbfSrchKey1.lCode = tmSbf.lCode
                    ilRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, PostRep
                    On Error GoTo 0
                    Do
                        ilRet = btrDelete(hmSbf)
                        If ilRet = BTRV_ERR_CONFLICT Then
                            tmSbfSrchKey1.lCode = tmSbf.lCode
                            ilCRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        End If
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    slMsg = "mSaveRec (btrDelete: Posting)"
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, PostRep
                    On Error GoTo 0
                    lmSave(2, ilLoop) = 0
                    If tgSpf.sPostCalAff = "W" Then
                        tmRwfSrchKey1.lSbfCode = tmSbf.lCode
                        ilRet = btrGetEqual(hmRwf, tmRwf, imRwfRecLen, tmRwfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            Do
                                ilRet = btrDelete(hmRwf)
                                If ilRet = BTRV_ERR_CONFLICT Then
                                    tmRwfSrchKey1.lSbfCode = tmSbf.lCode
                                    ilCRet = btrGetEqual(hmRwf, tmRwf, imRwfRecLen, tmRwfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            slMsg = "mSaveRec (btrDelete: Posting by Week)"
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, slMsg, PostRep
                            On Error GoTo 0
                        End If
                    End If
                End If
            End If
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(lmSbfDel) - 1 Step 1
        slMsg = "mSaveRec (btrGetEqual for Delete: Posting)"
        tmSbfSrchKey1.lCode = lmSbfDel(ilLoop)
        ilRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, PostRep
        On Error GoTo 0
        Do
            ilRet = btrDelete(hmSbf)
            If ilRet = BTRV_ERR_CONFLICT Then
                tmSbfSrchKey1.lCode = lmSbfDel(ilLoop)
                ilCRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        slMsg = "mSaveRec (btrDelete: Posting)"
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, PostRep
        On Error GoTo 0
        If tgSpf.sPostCalAff = "W" Then
            tmRwfSrchKey1.lSbfCode = lmSbfDel(ilLoop)
            ilRet = btrGetEqual(hmRwf, tmRwf, imRwfRecLen, tmRwfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                Do
                    ilRet = btrDelete(hmRwf)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        tmRwfSrchKey1.lSbfCode = lmSbfDel(ilLoop)
                        ilCRet = btrGetEqual(hmRwf, tmRwf, imRwfRecLen, tmRwfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                slMsg = "mSaveRec (btrDelete: Posting by Week)"
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, PostRep
                On Error GoTo 0
            End If
        End If
    Next ilLoop
    ReDim lmSbfDel(0 To 0) As Long
    imChg = False
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
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
    If imChg Then
        If ilAsk Then
            slMess = "Save Changes"
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
            If ilRes = vbNo Then
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
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    'Update button set if all mandatory fields have data and any field altered
    If imChg Then
        cmcUpdate.Enabled = True
        cbcInvDate.Enabled = False
        cbcNames.Enabled = False
        cmcClear.Enabled = False
    Else
        cbcInvDate.Enabled = True
        cbcNames.Enabled = True
        cmcUpdate.Enabled = False
    End If
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
    igManUnload = YES
    Unload PostRep
    igManUnload = NO
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the Vehicle box      *
'*                                                     *
'*******************************************************
Private Sub mVehPop(ilPopNames As Integer)
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilLoop As Integer
    'Dim llVehType As Long
    Dim slNameCode As String
    Dim slName As String
    Dim ilUpper As Integer

    'llVehType = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH
    If ilPopNames = 1 Then
        cbcNames.Clear
        ReDim tmVehicle(0 To UBound(igRepVefCode)) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            ilUpper = 0
            For ilLoop = LBound(igRepVefCode) To UBound(igRepVefCode) - 1 Step 1
        '        If tgMVef(ilVef).iCode = igRepVefCode(ilLoop) Then
                ilVef = gBinarySearchVef(igRepVefCode(ilLoop))
                'Dick: Excluded Rep-Net vehicles when posting by Count
                If (ilVef <> -1) And (igPostType = 1) Then
                    If tgMVef(ilVef).iNrfCode > 0 Then
                        ilVef = -1
                    End If
                End If
                If ilVef <> -1 Then
                    tmVehicle(ilUpper).sKey = tgMVef(ilVef).sName & "\" & Trim$(str$(tgMVef(ilVef).iCode))
                    ilUpper = ilUpper + 1
                    'ReDim Preserve tmVehicle(0 To UBound(tmVehicle) + 1) As SORTCODE 'VB list box clear (list box used to retain code number so record can be found)
        '            Exit For
                End If
            Next ilLoop
            ReDim Preserve tmVehicle(0 To ilUpper) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        'Next ilVef
        If UBound(tmVehicle) - 1 > 0 Then
            'ArraySortTyp tmVehicle(), tmVehicle(0), UBound(tmVehicle), 0, Len(tmVehicle(0)), 0, Len(tmVehicle(0).sKey), 0
            ArraySortTyp fnAV(tmVehicle(), 0), UBound(tmVehicle), 0, LenB(tmVehicle(0)), 0, LenB(tmVehicle(0).sKey), 0
        End If
        For ilLoop = 0 To UBound(tmVehicle) - 1 Step 1
            slNameCode = tmVehicle(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            slName = Trim$(slName)
            cbcNames.AddItem slName
        Next ilLoop
    ElseIf ilPopNames = 0 Then
        lbcVehicle.Clear
        ReDim tmVehicle(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If tgMVef(ilVef).iMnfVehGp3Mkt = imMktCode Then
                'If mTestVehType(ilVehType, tgMVef(ilVef)) Then
                '    tmVehicle(UBound(tmVehicle)).sKey = tgMVef(ilVef).sName & "\" & Trim$(Str$(tgMVef(ilVef).iCode))
                '    ReDim Preserve tmVehicle(0 To UBound(tmVehicle) + 1) As SORTCODE 'VB list box clear (list box used to retain code number so record can be found)
                'End If
                For ilLoop = LBound(igMktVefCode) To UBound(igMktVefCode) - 1 Step 1
                    If tgMVef(ilVef).iCode = igMktVefCode(ilLoop) Then
                        tmVehicle(UBound(tmVehicle)).sKey = tgMVef(ilVef).sName & "\" & Trim$(str$(tgMVef(ilVef).iCode))
                        ReDim Preserve tmVehicle(0 To UBound(tmVehicle) + 1) As SORTCODE 'VB list box clear (list box used to retain code number so record can be found)
                        Exit For
                    End If
                Next ilLoop
            End If
        Next ilVef
        If UBound(tmVehicle) - 1 > 0 Then
            'ArraySortTyp tmVehicle(), tmVehicle(0), UBound(tmVehicle), 0, Len(tmVehicle(0)), 0, Len(tmVehicle(0).sKey), 0
            ArraySortTyp fnAV(tmVehicle(), 0), UBound(tmVehicle), 0, LenB(tmVehicle(0)), 0, LenB(tmVehicle(0).sKey), 0
        End If
        For ilLoop = 0 To UBound(tmVehicle) - 1 Step 1
            slNameCode = tmVehicle(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            slName = Trim$(slName)
            lbcVehicle.AddItem slName  'Add ID to list box
        Next ilLoop
    Else
        lbcVehicle.Clear
        ReDim tmVehicle(0 To UBound(igRepVefCode)) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            ilUpper = 0
            For ilLoop = LBound(igRepVefCode) To UBound(igRepVefCode) - 1 Step 1
        '        If tgMVef(ilVef).iCode = igRepVefCode(ilLoop) Then
                ilVef = gBinarySearchVef(igRepVefCode(ilLoop))
                If (ilVef <> -1) And (igPostType = 2) Then
                    If tgMVef(ilVef).iNrfCode > 0 Then
                        ilVef = -1
                    End If
                End If
                If ilVef <> -1 Then
                    tmVehicle(ilUpper).sKey = tgMVef(ilVef).sName & "\" & Trim$(str$(tgMVef(ilVef).iCode))
                    ilUpper = ilUpper + 1
                    'ReDim Preserve tmVehicle(0 To UBound(tmVehicle) + 1) As SORTCODE 'VB list box clear (list box used to retain code number so record can be found)
        '            Exit For
                End If
            Next ilLoop
            ReDim Preserve tmVehicle(0 To ilUpper) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        'Next ilVef
        If UBound(tmVehicle) - 1 > 0 Then
            'ArraySortTyp tmVehicle(), tmVehicle(0), UBound(tmVehicle), 0, Len(tmVehicle(0)), 0, Len(tmVehicle(0).sKey), 0
            ArraySortTyp fnAV(tmVehicle(), 0), UBound(tmVehicle), 0, LenB(tmVehicle(0)), 0, LenB(tmVehicle(0).sKey), 0
        End If
        For ilLoop = 0 To UBound(tmVehicle) - 1 Step 1
            slNameCode = tmVehicle(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            slName = Trim$(slName)
            lbcVehicle.AddItem slName
        Next ilLoop
    End If
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub pbcArrow_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub

Sub pbcArrow_KeyUp(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub pbcClickFocus_GotFocus()
    mWklySetShow imWklyBoxNo
    imWklyBoxNo = -1
    imWklyRowNo = -1
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcSTab_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilFound As Integer

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = -1  'Set-right to left
    ilBox = imBoxNo
    ilRow = imRowNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imTabDirection = 0  'Set-Left to right
                imSettingValue = True
                If tmcClick.Enabled Then
                    If (igPostType = 5) Or (igPostType = 6) Then
                        cmcDone.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                If UBound(smSave, 2) <= 1 Then
                    If (igPostType = 5) Or (igPostType = 6) Then
                        cmcDone.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                If InStr(1, smShow(imTOTALINDEX, 1), "Total:", 1) > 0 Then
                    If (igPostType = 5) Or (igPostType = 6) Then
                        cmcDone.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                vbcPostRep.Value = vbcPostRep.Min
                If UBound(smSave, 2) <= vbcPostRep.LargeChange + 1 Then 'was <=
                    vbcPostRep.Max = LBONE  'LBound(smSave, 2)
                Else
                    vbcPostRep.Max = UBound(smSave, 2) - vbcPostRep.LargeChange ' - 1
                End If
                imRowNo = 1
                Do While (imRowNo < UBound(smSave, 2)) And (smSave(8, imRowNo) = "Y")
                    imRowNo = imRowNo + 1
                    If imRowNo > vbcPostRep.Value + vbcPostRep.LargeChange Then
                        imSettingValue = True
                        vbcPostRep.Value = vbcPostRep.Value + 1
                    End If
                Loop
                If (imRowNo = UBound(smSave, 2)) Then
                    If (igPostType = 5) Or (igPostType = 6) Then
                        cmcDone.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                If igPostType = 1 And tgSpf.sPostCalAff = "W" Then
                    ilBox = imINVOICENOINDEX
                Else
                    ilBox = imANOSPOTSINDEX
                End If
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            Case imVEHICLEINDEX 'Name (first control within header)
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    mSetShow imBoxNo
                End If
                If imANEXTNOSPOTSINDEX <= 0 Then
                    ilBox = imANOSPOTSINDEX 'ABONUSINDEX
                Else
                    ilBox = imANEXTNOSPOTSINDEX
                End If
                If imRowNo <= 1 Then
                    imBoxNo = -1
                    imRowNo = -1
                    pbcArrow.Visible = False
                    lacFrame.Visible = False
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imRowNo = imRowNo - 1
                If imRowNo < vbcPostRep.Value Then
                    imSettingValue = True
                    vbcPostRep.Value = vbcPostRep.Value - 1
                End If
                'imBoxNo = ilBox
                'mEnableBox ilBox
                'Exit Sub
            Case imINVOICENOINDEX
                If imRowNo <= 1 Then
                    mSetShow imBoxNo
                    imBoxNo = -1
                    imRowNo = -1
                    pbcArrow.Visible = False
                    lacFrame.Visible = False
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imRowNo = imRowNo - 1
                If imRowNo < vbcPostRep.Value Then
                    imSettingValue = True
                    vbcPostRep.Value = vbcPostRep.Value - 1
                End If
                ilBox = imANOSPOTSINDEX
            Case imANOSPOTSINDEX 'Name (first control within header)
                If tgSpf.sPostCalAff = "W" Then
                    mWklySetShow imWklyBoxNo
                    If imWklyBoxNo = WKLYCARRIEDNOINDEX Then
                        imWklyBoxNo = WKLYAIRNOINDEX
                        mWklyEnableBox imWklyBoxNo
                        Exit Sub
                    Else
                        imWklyRowNo = imWklyRowNo - 1
                        If (imWklyRowNo >= 1) Then
                            imWklyBoxNo = WKLYCARRIEDNOINDEX
                            mWklyEnableBox imWklyBoxNo
                            Exit Sub
                        End If
                    End If
                End If
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    mSetShow imBoxNo
                    ilBox = imANOSPOTSINDEX
                End If
                If imANEXTNOSPOTSINDEX <= 0 Then
                    If igPostType = 1 And tgSpf.sPostCalAff = "W" Then
                        mSetShow imBoxNo
                        ilBox = imINVOICENOINDEX
                        imBoxNo = ilBox
                        mEnableBox ilBox
                    Else
                        ilBox = imANOSPOTSINDEX 'ABONUSINDEX
                    End If
                Else
                    ilBox = imANEXTNOSPOTSINDEX
                End If
                If imRowNo <= 1 Then
                    imBoxNo = -1
                    imRowNo = -1
                    pbcArrow.Visible = False
                    lacFrame.Visible = False
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imRowNo = imRowNo - 1
                If imRowNo < vbcPostRep.Value Then
                    imSettingValue = True
                    vbcPostRep.Value = vbcPostRep.Value - 1
                End If
                'imBoxNo = ilBox
                'mEnableBox ilBox
                'Exit Sub
            Case Else
                ilBox = ilBox - 1
        End Select
        If (smSave(8, imRowNo) = "Y") Or (InStr(1, smShow(imTOTALINDEX, imRowNo), "Total:", 1) > 0) Then
            ilFound = False
        End If
    Loop While Not ilFound
    If (imRowNo = ilRow) Then
        mSetShow imBoxNo
    End If
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilFound As Integer

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = 0  'Set-Left to right
    ilBox = imBoxNo
    ilRow = imRowNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imTabDirection = -1  'Set-Right to left
                imRowNo = UBound(smSave, 2)
                If imRowNo = LBONE Then
                    imRowNo = -1
                    If (igPostType = 5) Or (igPostType = 6) Then
                        cmcDone.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                If InStr(1, smShow(imTOTALINDEX, imRowNo), "Total:", 1) > 0 Then
                    imRowNo = -1
                    If (igPostType = 5) Or (igPostType = 6) Then
                        cmcDone.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                imSettingValue = True
                If imRowNo <= vbcPostRep.LargeChange + 1 Then
                    vbcPostRep.Value = vbcPostRep.Min
                Else
                    vbcPostRep.Value = imRowNo - vbcPostRep.LargeChange
                End If
                If (igPostType = 1 And tgSpf.sPostCalAff = "W") Then
                    ilBox = imINVOICENOINDEX
                Else
                    ilBox = imANOSPOTSINDEX
                End If
            Case 0
                If (igPostType = 1 And tgSpf.sPostCalAff = "W") And (InStr(1, smShow(imTOTALINDEX, imRowNo), "Total:", 1) <= 0) Then
                    ilBox = imINVOICENOINDEX
                Else
                    ilBox = imANOSPOTSINDEX
                End If
            Case imVEHICLEINDEX
                If (Trim$(smInfo(1, ilRow)) = "S") Then
                    ilBox = imOPRICEINDEX
                Else
                    ilBox = imANOSPOTSINDEX
                End If
            Case imOPRICEINDEX
                ilBox = imANOSPOTSINDEX
            Case imANOSPOTSINDEX 'Last control
                If (smSave(8, imRowNo) <> "Y") And (InStr(1, smShow(imTOTALINDEX, imRowNo), "Total:", 1) <= 0) Then
                    If tgSpf.sPostCalAff = "W" Then
                        mWklySetShow imWklyBoxNo
                        If imWklyBoxNo = WKLYAIRNOINDEX Then
                            imWklyBoxNo = WKLYCARRIEDNOINDEX
                            mWklyEnableBox imWklyBoxNo
                            Exit Sub
                        Else
                            imWklyRowNo = imWklyRowNo + 1
                            If (imWklyRowNo < 5) Or ((imWklyRowNo = 5) And (Trim$(smWklyDates(5)) <> "")) Then
                                imWklyBoxNo = WKLYAIRNOINDEX
                                mWklyEnableBox imWklyBoxNo
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                If imANEXTNOSPOTSINDEX <= 0 Then
                    If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                        mSetShow imBoxNo
                        If mTestSaveFields(imRowNo) = NO Then
                            mEnableBox imBoxNo
                            Exit Sub
                        End If
                    End If
                    imRowNo = imRowNo + 1
                    If imRowNo > vbcPostRep.Value + vbcPostRep.LargeChange Then
                        imSettingValue = True
                        vbcPostRep.Value = vbcPostRep.Value + 1
                    End If
                    If imRowNo >= UBound(smSave, 2) Then
                        mSetCommands
                        imBoxNo = 0
                        If (igPostType = 5) Or (igPostType = 6) Then
                            cmcDone.SetFocus
                        Else
                            cmcCancel.SetFocus
                        End If
                        Exit Sub
                    End If
                    If igPostType = 1 And tgSpf.sPostCalAff = "W" Then
                        ilBox = imINVOICENOINDEX
                        imBoxNo = ilBox
                        mEnableBox ilBox
                        Exit Sub
                    End If
                Else
                    ilBox = imANEXTNOSPOTSINDEX
                End If
            Case imANEXTNOSPOTSINDEX 'Last control
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    mSetShow imBoxNo
                    If mTestSaveFields(imRowNo) = NO Then
                        mEnableBox imBoxNo
                        Exit Sub
                    End If
                End If
                imRowNo = imRowNo + 1
                If imRowNo > vbcPostRep.Value + vbcPostRep.LargeChange Then
                    imSettingValue = True
                    vbcPostRep.Value = vbcPostRep.Value + 1
                End If
                If imRowNo >= UBound(smSave, 2) Then
                    mSetCommands
                    imBoxNo = 0
                    If (igPostType = 5) Or (igPostType = 6) Then
                        cmcDone.SetFocus
                    Else
                        cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                ilBox = imANOSPOTSINDEX
            Case Else
                ilBox = ilBox + 1
        End Select
        If (smSave(8, imRowNo) = "Y") Or (InStr(1, smShow(imTOTALINDEX, imRowNo), "Total:", 1) > 0) Then
            ilFound = False
        End If
    Loop While Not ilFound
    If (imRowNo = ilRow) Then
        mSetShow imBoxNo
    End If
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcPostRep_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Sub pbcPostRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    imButton = Button
    'If Button = 2 Then  'Right Mouse
        ilCompRow = vbcPostRep.LargeChange + 1
        If UBound(smSave, 2) > ilCompRow - 1 Then
            ilMaxRow = ilCompRow
        Else
            ilMaxRow = UBound(smSave, 2) - 1
        End If
        For ilRow = 1 To ilMaxRow Step 1
            For ilBox = imLBCtrls To imMaxIndex Step 1
                If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                    If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                        imButtonRow = ilRow + vbcPostRep.Value - 1
                        mShowInfo
                        Exit Sub
                    End If
                End If
            Next ilBox
        Next ilRow
    'End If
End Sub
Sub pbcPostRep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer

    If imIgnoreRightMove Then
        Exit Sub
    End If
    'pbcPostRep.ToolTipText = ""
    'If Button <> 2 Then  'Right Mouse
    '    imIgnoreRightMove = True
    '    ilCompRow = vbcPostRep.LargeChange + 1
    '    If UBound(smSave, 2) > ilCompRow - 1 Then
    '        ilMaxRow = ilCompRow
    '    Else
    '        ilMaxRow = UBound(smSave, 2) - 1
    '    End If
    '    For ilRow = 1 To ilMaxRow Step 1
    '        If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(1).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(1).fBoxY + tmCtrls(1).fBoxH)) Then
    '            ilCompRow = ilRow + vbcPostRep.Value - 1
    '            If ((igPostType = 1) Or (igPostType = 2)) And (Trim$(smInfo(1, ilCompRow)) <> "T") Then
    '                pbcPostRep.ToolTipText = smSave(33, ilCompRow)
    '            End If
    '            Exit For
    '        End If
    '    Next ilRow
    '    imIgnoreRightMove = False
    '    Exit Sub
    'End If
    If (imBoxNo < imLBCtrls) Or (imBoxNo > imMaxIndex) Then
        imButton = Button
        imIgnoreRightMove = True
        ilCompRow = vbcPostRep.LargeChange + 1
        If UBound(smSave, 2) > ilCompRow - 1 Then
            ilMaxRow = ilCompRow
        Else
            ilMaxRow = UBound(smSave, 2) - 1
        End If
        For ilRow = 1 To ilMaxRow Step 1
            For ilBox = imLBCtrls To imMaxIndex Step 1
                If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                    If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                        imButtonRow = ilRow + vbcPostRep.Value - 1
                        mShowInfo
                        imIgnoreRightMove = False
                        Exit Sub
                    End If
                End If
            Next ilBox
        Next ilRow
    End If
    plcInfo.Visible = False
    imIgnoreRightMove = False
End Sub
Private Sub pbcPostRep_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim ilRepRow As Integer
    Dim ilRet As Integer

    'If Button = 2 Then
    '    plcInfo.Visible = False
    '    Exit Sub
    'End If
    ilCompRow = vbcPostRep.LargeChange + 1
    If UBound(smSave, 2) > ilCompRow - 1 Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smSave, 2) - 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To imMaxIndex Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcPostRep.Value - 1
                    mWklySetShow imWklyBoxNo
                    mSetShow imBoxNo
                    If ilRowNo > UBound(smSave, 2) - 1 Then
                        Beep
                        Exit Sub
                    End If
                    If Trim$(smSave(8, ilRowNo)) = "Y" Then    'If billed disallow change
                        If (igPostType <> 5) And (igPostType <> 6) Then
                            If tgSpf.sPostCalAff <> "W" Then
                                Beep
                                Exit Sub
                            Else
                                If (ilBox <> imANOSPOTSINDEX) And (ilBox <> imINVOICENOINDEX) Then
                                    Beep
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                    If Trim$(smShow(imCONTRACTINDEX, ilRowNo)) = "" Then
                        Beep
                        Exit Sub
                    End If
                    If (Trim$(smInfo(1, ilRowNo)) = "T") Then
                        Beep
                        Exit Sub
                    End If
                    If igPostType = 3 Then
                        imSetAll = False
                        ckcAll.Value = vbUnchecked
                        imSetAll = True
                        imChg = True
                        imSave(5, ilRowNo) = Not imSave(5, ilRowNo)
                        pbcPostRep_Paint
                        mSetCommands
                    ElseIf (igPostType = 5) Or (igPostType = 6) Then
                        plcInfo.Visible = False
                        lgRepSpotChfCode = lmSave(1, ilRowNo)
                        igRepSpotLineNo = imSave(7, ilRowNo)
                        lgRepSpotStartCal = lmStartCal
                        lgRepSpotEndCal = lmEndCal
                        lgRepSpotStartStd = lmStartStd
                        lgRepSpotEndStd = lmEndStd
                        igRepSpotBilled = False
                        If Trim$(smSave(8, ilRowNo)) = "Y" Then    'If billed disallow change
                            igRepSpotBilled = True
                        End If
                        ReDim tgRepSpotLineInfo(0 To 0) As REPSPOTLINEINFO
                        For ilRepRow = LBONE To UBound(imSave, 2) Step 1
                            If lmSave(1, ilRepRow) = lgRepSpotChfCode Then
                                If (Trim$(smInfo(1, ilRepRow)) <> "T") Then
                                    tgRepSpotLineInfo(UBound(tgRepSpotLineInfo)).iLineNo = imSave(7, ilRepRow)
                                    'tgRepSpotLineInfo(UBound(tgRepSpotLineInfo)).iCarry = Val(smSave(10, ilRepRow))
                                    'tgRepSpotLineInfo(UBound(tgRepSpotLineInfo)).iCalCarry = Val(smSave(9, ilRepRow))
                                    ReDim Preserve tgRepSpotLineInfo(0 To UBound(tgRepSpotLineInfo) + 1) As REPSPOTLINEINFO
                                End If
                            End If
                        Next ilRepRow
                        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, lgRepSpotChfCode, False, tgChfRep, tgClfRep(), tgCffRep())
                        If ilRet Then
                            PostRepTimes.Show vbModal
                            'Refresh list
                            DoEvents
                            tmcClick_Timer
                        End If
                        Erase tgRepSpotLineInfo
                    Else
                        If (Trim$(smInfo(1, ilRowNo)) <> "S") Then
                            'If (ilBox <> AGROSSINDEX) And (ilBox <> imANOSPOTSINDEX) And (ilBox <> ABONUSINDEX) Then
                            If (ilBox <> imANOSPOTSINDEX) And (ilBox <> imINVOICENOINDEX) And (ilBox <> imANEXTNOSPOTSINDEX) Then
                                'If (Trim$(smInfo(1, ilRowNo)) <> "S") Or (lmSave(2, ilRowNo) <> 0) Or (ilBox <> imVEHICLEINDEX) Then
                                    imRowNo = ilRowNo
                                    imBoxNo = 0
                                    lacFrame.Move 0, tmCtrls(imCONTRACTINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15) - 30
                                    lacFrame.Visible = True
                                    pbcArrow.Move pbcArrow.Left, plcPostRep.Top + tmCtrls(imCONTRACTINDEX).fBoxY + (imRowNo - vbcPostRep.Value) * (fgBoxGridH + 15) + 45
                                    pbcArrow.Visible = True
                                    pbcArrow.SetFocus
                                    Exit Sub
                                'End If
                            End If
                        End If
                        imRowNo = ilRowNo
                        imBoxNo = ilBox
                        mEnableBox ilBox
                    End If
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcPostRep_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    mPaintRepTitle

    slFontName = pbcPostRep.FontName
    flFontSize = pbcPostRep.FontSize
    llColor = pbcPostRep.ForeColor
'    pbcPostRep.FontBold = False
'    pbcPostRep.FontSize = 7
'    pbcPostRep.FontName = "Arial"
'    pbcPostRep.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
'    pbcPostRep.ForeColor = BLUE
'    If igPostType = 1 Then
'        pbcPostRep.CurrentX = tmCtrls(imAdvtIndex).fBoxX + fgBoxInsetX
'        pbcPostRep.CurrentY = 15 '- 30'+ fgBoxInsetY
'        pbcPostRep.Print "Advertiser"
'    ElseIf igPostType = 2 Then
'        pbcPostRep.CurrentX = tmCtrls(imVEHICLEINDEX).fBoxX + fgBoxInsetX
'        pbcPostRep.CurrentY = 15 '- 30'+ fgBoxInsetY
'        pbcPostRep.Print "Vehicle"
'    End If
'    pbcPostRep.ForeColor = llColor
'    pbcPostRep.FontSize = flFontSize
'    pbcPostRep.FontName = slFontName
'    pbcPostRep.FontSize = flFontSize
'    pbcPostRep.FontBold = True

    ilStartRow = vbcPostRep.Value  'Top location
    ilEndRow = vbcPostRep.Value + vbcPostRep.LargeChange
    If ilEndRow > UBound(smSave, 2) - 1 Then
        ilEndRow = UBound(smSave, 2) - 1 'Don't include blank row
    End If
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBCtrls To imMaxIndex Step 1
            'If (ilBox = VEHICLEINDEX) Then
            '    If (lmSave(1, ilRow) > 0) Or (lmSave(2, ilRow) > 0) Then
            '        gPaintArea pbcPostRep, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
            '    Else
            '        gPaintArea pbcPostRep, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
            '    End If
            'End If
            If smSave(8, ilRow) = "Y" Then    'If billed- override any other color
                pbcPostRep.ForeColor = DARKGREEN
            End If
            If (imSave(4, ilRow) = False) And (ilBox = imVEHICLEINDEX) And ((Trim$(smInfo(1, ilRow)) = "I") Or (Trim$(smInfo(1, ilRow)) = "F") Or (Trim$(smInfo(1, ilRow)) = "C")) Then
                pbcPostRep.ForeColor = RED
            End If
            pbcPostRep.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcPostRep.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            If (ilBox = imMISSEDINDEX) Then
                'slStr = ""
                ''If (smSave(5, ilRow) <> "") And (smSave(7, ilRow) <> "") Then
                ''    slStr = gAddStr(smSave(5, ilRow), smSave(7, ilRow))
                ''ElseIf (smSave(5, ilRow) <> "") Then
                ''    slStr = smSave(5, ilRow)
                ''ElseIf (smSave(7, ilRow) <> "") Then
                ''    slStr = smSave(7, ilRow)
                ''End If
                'If (smSave(5, ilRow) <> "") Then
                '    slStr = smSave(5, ilRow)
                'End If
                'If (slStr <> "") Then
                '    slStr = gSubStr(slStr, smSave(3, ilRow))
                '    If Val(slStr) >= 0 Then
                '        slStr = ""
                '    End If
                'End If
                slStr = smSave(14, ilRow)
                gSetShow pbcPostRep, slStr, tmCtrls(imMISSEDINDEX)
                slStr = tmCtrls(imMISSEDINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imMGINDEX) Then
                'slStr = ""
                'If (smSave(5, ilRow) <> "") Then
                '    slStr = smSave(5, ilRow)
                'End If
                'If (slStr <> "") Then
                '    slStr = gSubStr(slStr, smSave(3, ilRow))
                '    If Val(slStr) > 0 Then
                '        If Val(smSave(10, ilRow)) < 0 Then
                '            If Val(slStr) > Abs(Val(smSave(10, ilRow))) Then
                '                slStr = Str$(Abs(Val(smSave(10, ilRow))))
                '            End If
                '        Else
                '            slStr = ""
                '        End If
                '    Else
                '        slStr = ""
                '    End If
                'End If
                slStr = smSave(13, ilRow)
                gSetShow pbcPostRep, slStr, tmCtrls(imMGINDEX)
                slStr = tmCtrls(imMGINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imBONUSINDEX) Then
                'slStr = ""
                'If (smSave(5, ilRow) <> "") Then
                '    slStr = smSave(5, ilRow)
                'End If
                'If (slStr <> "") Then
                '    slStr = gSubStr(slStr, smSave(3, ilRow))
                '    If Val(slStr) > 0 Then
                '        If Val(smSave(10, ilRow)) <= 0 Then
                '            If Val(slStr) > Abs(Val(smSave(10, ilRow))) Then
                '                slStr = Str$(Val(slStr) - Abs(Val(smSave(10, ilRow))))
                '            Else
                '                slStr = ""
                '            End If
                '        Else
                '            slStr = ""
                '        End If
                '    Else
                '        slStr = ""
                '    End If
                'End If
                slStr = smSave(7, ilRow)
                gSetShow pbcPostRep, slStr, tmCtrls(imBONUSINDEX)
                slStr = tmCtrls(imBONUSINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imDGROSSINDEX) Then
                slStr = ""
                If smSave(6, ilRow) <> "" Then
                    slStr = gSubStr(smSave(6, ilRow), smSave(4, ilRow))
                End If
                gSetShow pbcPostRep, slStr, tmCtrls(imDGROSSINDEX)
                slStr = tmCtrls(imDGROSSINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imANOSPOTSPRIORINDEX) Then
                slStr = Trim$(smSave(9, ilRow))
                gSetShow pbcPostRep, slStr, tmCtrls(imANOSPOTSPRIORINDEX)
                slStr = tmCtrls(imANOSPOTSPRIORINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imANOSPOTSCURRINDEX) Then
                slStr = Trim$(smSave(5, ilRow))
                gSetShow pbcPostRep, slStr, tmCtrls(imANOSPOTSCURRINDEX)
                slStr = tmCtrls(imANOSPOTSCURRINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imDIFFSPOTSINDEX) Then
                slStr = gSubStr(gAddStr(smSave(9, ilRow), smSave(5, ilRow)), smSave(3, ilRow))
                gSetShow pbcPostRep, slStr, tmCtrls(imDIFFSPOTSINDEX)
                slStr = tmCtrls(imDIFFSPOTSINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imDIFFGROSSINDEX) Then
                slStr = gSubStr(smSave(6, ilRow), smSave(4, ilRow))
                gSetShow pbcPostRep, slStr, tmCtrls(imDIFFGROSSINDEX)
                slStr = tmCtrls(imDIFFGROSSINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imBONUSPREVINDEX) Then
                slStr = Trim$(str$(imSave(10, ilRow)))
                gSetShow pbcPostRep, slStr, tmCtrls(imBONUSPREVINDEX)
                slStr = tmCtrls(imBONUSPREVINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imBONUSCURRINDEX) Then
                slStr = Trim$(str$(imSave(8, ilRow)))
                gSetShow pbcPostRep, slStr, tmCtrls(imBONUSCURRINDEX)
                slStr = tmCtrls(imBONUSCURRINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imNNOSPOTSINDEX) Then
                slStr = Trim$(smSave(16, ilRow))
                gSetShow pbcPostRep, slStr, tmCtrls(imNNOSPOTSINDEX)
                slStr = tmCtrls(imNNOSPOTSINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imNNOBONUSINDEX) Then
                slStr = Trim$(str$(imSave(9, ilRow)))
                gSetShow pbcPostRep, slStr, tmCtrls(imNNOBONUSINDEX)
                slStr = tmCtrls(imNNOBONUSINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imAdvtIndex) And (igPostType = 2) Then
                'Don't show name
            ElseIf (ilBox = imVEHICLEINDEX) And (igPostType = 1) Then
                'Don't show name
            ElseIf (ilBox = imLINEINDEX) And ((igPostType = 5) Or (igPostType = 6)) And ((Trim$(smInfo(1, ilRow)) <> "T")) Then
                slStr = Trim$(str$(imSave(7, ilRow)))
                gSetShow pbcPostRep, slStr, tmCtrls(imLINEINDEX)
                slStr = tmCtrls(imLINEINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf (ilBox = imLENGTHINDEX) And ((igPostType = 5) Or (igPostType = 6)) And ((Trim$(smInfo(1, ilRow)) <> "T")) Then
                slStr = Trim$(str$(imSave(6, ilRow)))
                gSetShow pbcPostRep, slStr, tmCtrls(imLENGTHINDEX)
                slStr = tmCtrls(imLENGTHINDEX).sShow
                pbcPostRep.Print slStr
            ElseIf ilBox = imCHECKINDEX Then
                pbcPostRep.FontName = "Monotype Sorts"
                pbcPostRep.FontBold = False
                'pbcPostRep.CurrentX = tmRBCtrls(1).fBoxX + fgBoxInsetX - 30
                'pbcPostRep.CurrentY = tmRBCtrls(1).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) + 15 '- 60'+ fgBoxInsetY
                gPaintArea pbcPostRep, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
                pbcPostRep.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX - 30
                pbcPostRep.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) + 15 '+ fgBoxInsetY
                If imSave(5, ilRow) = True Then
                    pbcPostRep.Print " 4"
                Else
                    pbcPostRep.Print "   "
                End If
                pbcPostRep.FontName = slFontName
                pbcPostRep.FontBold = True
            ElseIf (ilBox = imINVOICENOINDEX) Then
                'If (smShow(imCONTRACTINDEX, ilRow) <> "") And (InStr(1, smShow(imTOTALINDEX, ilRow), "Total:", 1) <= 0) Then
                    gPaintArea pbcPostRep, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
                'Else
                '    gPaintArea pbcPostRep, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                'End If
                pbcPostRep.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
                pbcPostRep.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                If lmSave(5, ilRow) > 0 Then
                    slStr = Trim$(str$(lmSave(5, ilRow)))
                    gSetShow pbcPostRep, slStr, tmCtrls(imINVOICENOINDEX)
                    slStr = tmCtrls(imINVOICENOINDEX).sShow
                Else
                    slStr = ""
                End If
                pbcPostRep.Print slStr
            ElseIf (Trim$(smInfo(1, ilRow)) = "S") Then
                slStr = Trim$(smShow(ilBox, ilRow))
                'If (ilBox = imOPRICEINDEX) Or (ilBox = imOGROSSINDEX) Or (ilBox = imVEHICLEINDEX) Then
                If (ilBox = imVEHICLEINDEX) Or (ilBox = imOPRICEINDEX) Then
                    gPaintArea pbcPostRep, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
                    pbcPostRep.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
                    pbcPostRep.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                End If
                pbcPostRep.Print slStr
            Else
                slStr = Trim$(smShow(ilBox, ilRow))
                If ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) And (ilBox = imOPRICEINDEX) And ((Trim$(smInfo(1, ilRow)) <> "T")) Then
                    'slStr = Trim$(Str$(imSave(6, ilRow))) & """" & " " & slStr
                    gSetShow pbcPostRep, slStr, tmCtrls(imOPRICEINDEX)
                    slStr = tmCtrls(imOPRICEINDEX).sShow
                End If
                pbcPostRep.Print slStr
            End If
            If (ilBox = imVEHICLEINDEX) Then
                pbcPostRep.ForeColor = llColor
            End If
        Next ilBox
        pbcPostRep.ForeColor = llColor
    Next ilRow
End Sub

Private Sub pbcWeekly_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilBox As Integer

    For ilRow = 1 To 5 Step 1
        If Trim$(smWklyDates(ilRow)) <> "" Then
            For ilBox = imLBWklyCtrls To UBound(tmWklyCtrls) Step 1
                If (X >= tmWklyCtrls(ilBox).fBoxX) And (X <= (tmWklyCtrls(ilBox).fBoxX + tmWklyCtrls(ilBox).fBoxW)) Then
                    If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmWklyCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmWklyCtrls(ilBox).fBoxY + tmWklyCtrls(ilBox).fBoxH)) Then
                        If (Trim$(smSave(8, imRowNo)) = "Y") Or (Trim$(smSave(17, imRowNo)) = "Y") Then    'If billed disallow change
                            Beep
                            Exit Sub
                        End If
                        mWklySetShow imWklyBoxNo
                        If (ilBox <= WKLYDATESINDEX) Or (ilBox >= WKLYTOTALINDEX) Then
                            Beep
                            Exit Sub
                        End If
                        imWklyRowNo = ilRow
                        imWklyBoxNo = ilBox
                        mWklyEnableBox ilBox
                        Exit Sub
                    End If
                End If
            Next ilBox
        End If
    Next ilRow
End Sub

Private Sub pbcWeekly_Paint()
    Dim ilRow As Integer
    Dim ilBox As Integer
    Dim ilRowTotal As Integer
    Dim ilOrderColTotal As Integer
    Dim ilAirColTotal As Integer
    Dim ilCarriedColTotal As Integer
    Dim ilTotal As Integer
    Dim slStr As String
    Dim llColor As Long

    If (imRowNo < LBONE) Or (imRowNo > UBound(smSave, 2)) Then
        Exit Sub
    End If
    ilTotal = 0
    ilAirColTotal = 0
    ilCarriedColTotal = 0
    ilOrderColTotal = 0
    pbcWeekly.Cls
    llColor = pbcWeekly.ForeColor
    If (smSave(17, imRowNo) = "Y") Then    'If billed- override any other color
        pbcWeekly.ForeColor = DARKGREEN
    End If
    For ilRow = 1 To 5 Step 1
        ilRowTotal = 0
        If Trim$(smWklyDates(ilRow)) <> "" Then
            For ilBox = imLBWklyCtrls To UBound(tmWklyCtrls) Step 1
                pbcWeekly.CurrentX = tmWklyCtrls(ilBox).fBoxX + fgBoxInsetX
                pbcWeekly.CurrentY = tmWklyCtrls(ilBox).fBoxY + (ilRow - 1) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                Select Case ilBox
                    Case WKLYDATESINDEX
                        slStr = Trim$(smWklyDates(ilRow))
                        gSetShow pbcWeekly, slStr, tmWklyCtrls(WKLYDATESINDEX)
                        slStr = tmWklyCtrls(WKLYDATESINDEX).sShow
                        pbcWeekly.Print slStr
                    Case WKLYORDERNOINDEX
                        If Trim$(smSave(28 + ilRow - 1, imRowNo)) <> "" Then
                            ilOrderColTotal = ilOrderColTotal + Val(smSave(28 + ilRow - 1, imRowNo))
                        End If
                        slStr = Trim$(smSave(28 + ilRow - 1, imRowNo))
                        gSetShow pbcWeekly, slStr, tmWklyCtrls(WKLYORDERNOINDEX)
                        slStr = tmWklyCtrls(WKLYORDERNOINDEX).sShow
                        pbcWeekly.Print slStr
                    Case WKLYAIRNOINDEX
                        If Trim$(smSave(18 + ilRow - 1, imRowNo)) <> "" Then
                            ilRowTotal = ilRowTotal + Val(smSave(18 + ilRow - 1, imRowNo))
                            ilAirColTotal = ilAirColTotal + Val(smSave(18 + ilRow - 1, imRowNo))
                        End If
                        slStr = Trim$(smSave(18 + ilRow - 1, imRowNo))
                        gSetShow pbcWeekly, slStr, tmWklyCtrls(WKLYAIRNOINDEX)
                        slStr = tmWklyCtrls(WKLYAIRNOINDEX).sShow
                        pbcWeekly.Print slStr
                    Case WKLYCARRIEDNOINDEX
                        If Trim$(smSave(23 + ilRow - 1, imRowNo)) <> "" Then
                            ilRowTotal = ilRowTotal + Val(smSave(23 + ilRow - 1, imRowNo))
                            ilCarriedColTotal = ilCarriedColTotal + Val(smSave(23 + ilRow - 1, imRowNo))
                        End If
                        slStr = Trim$(smSave(23 + ilRow - 1, imRowNo))
                        gSetShow pbcWeekly, slStr, tmWklyCtrls(WKLYCARRIEDNOINDEX)
                        slStr = tmWklyCtrls(WKLYCARRIEDNOINDEX).sShow
                        pbcWeekly.Print slStr
                    Case WKLYTOTALINDEX
                        ilTotal = ilTotal + ilRowTotal
                        slStr = Trim$(str$(ilRowTotal))
                        gSetShow pbcWeekly, slStr, tmWklyCtrls(WKLYTOTALINDEX)
                        slStr = tmWklyCtrls(WKLYTOTALINDEX).sShow
                        pbcWeekly.Print slStr
                End Select
            Next ilBox
        End If
    Next ilRow
    pbcWeekly.ForeColor = llColor
    pbcWeekly.CurrentX = tmWklyCtrls(WKLYORDERNOINDEX).fBoxX + fgBoxInsetX
    pbcWeekly.CurrentY = tmWklyCtrls(WKLYORDERNOINDEX).fBoxY + (4) * (fgBoxGridH + 15) - 30 + 225
    slStr = Trim$(str$(ilOrderColTotal))
    gSetShow pbcWeekly, slStr, tmWklyCtrls(WKLYORDERNOINDEX)
    slStr = tmWklyCtrls(WKLYORDERNOINDEX).sShow
    pbcWeekly.Print slStr
    pbcWeekly.CurrentX = tmWklyCtrls(WKLYAIRNOINDEX).fBoxX + fgBoxInsetX
    pbcWeekly.CurrentY = tmWklyCtrls(WKLYAIRNOINDEX).fBoxY + (4) * (fgBoxGridH + 15) - 30 + 225
    slStr = Trim$(str$(ilAirColTotal))
    gSetShow pbcWeekly, slStr, tmWklyCtrls(WKLYAIRNOINDEX)
    slStr = tmWklyCtrls(WKLYAIRNOINDEX).sShow
    pbcWeekly.Print slStr
    pbcWeekly.CurrentX = tmWklyCtrls(WKLYCARRIEDNOINDEX).fBoxX + fgBoxInsetX
    pbcWeekly.CurrentY = tmWklyCtrls(WKLYCARRIEDNOINDEX).fBoxY + (4) * (fgBoxGridH + 15) - 30 + 225
    slStr = Trim$(str$(ilCarriedColTotal))
    gSetShow pbcWeekly, slStr, tmWklyCtrls(WKLYCARRIEDNOINDEX)
    slStr = tmWklyCtrls(WKLYCARRIEDNOINDEX).sShow
    pbcWeekly.Print slStr
    pbcWeekly.CurrentX = tmWklyCtrls(WKLYTOTALINDEX).fBoxX + fgBoxInsetX
    pbcWeekly.CurrentY = tmWklyCtrls(WKLYTOTALINDEX).fBoxY + (4) * (fgBoxGridH + 15) - 30 + 225
    slStr = Trim$(str$(ilTotal))
    gSetShow pbcWeekly, slStr, tmWklyCtrls(WKLYTOTALINDEX)
    slStr = tmWklyCtrls(WKLYTOTALINDEX).sShow
    pbcWeekly.Print slStr
End Sub

Private Sub plcPostRep_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcPostRep_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub plcSelect_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSelect_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub

Private Sub plcWeekly_Paint()
    If (imRowNo < LBONE) Or (imRowNo > UBound(smSave, 2)) Then
        Exit Sub
    End If
    If (Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER Then
        plcWeekly.CurrentX = 0
        plcWeekly.CurrentY = 0
        plcWeekly.Print "Spot Length" & str$(imSave(6, imRowNo)) & "s"
    End If
End Sub


Private Sub tmcClick_Timer()
    Dim ilRet As Integer
    Dim ilSBFFound As Integer
    Dim slMsgFile As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer
    Dim ilSbfExist As Integer
    Dim ilVef As Integer

    Screen.MousePointer = vbHourglass
    tmcClick.Enabled = False
    pbcPostRep.Cls
    mClearCtrlFields

    ilSBFFound = False
    mBuildDate
    ReDim imMktVefCode(0 To 0) As Integer
    pbcPostRep.Cls
    If imMarketIndex >= 0 Then
        If (igPostType = 5) And ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then      'Rep Spot Times
            slNameCode = tmNetNames(imMarketIndex).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imMktCode = Val(slCode)
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                'If (tgMVef(ilLoop).iNrfCode = imMktCode) Then
                If (tgMVef(ilLoop).iNrfCode = imMktCode) Or ((tgSpf.sAEDII = "Y") And ((Asc(tgSpf.sUsingFeatures8) And REPBYDT) = REPBYDT) And (tgMVef(ilLoop).iNrfCode = 0) And (tgMVef(ilLoop).sType = "R")) Then
                    imMktVefCode(UBound(imMktVefCode)) = tgMVef(ilLoop).iCode
                    ReDim Preserve imMktVefCode(0 To UBound(imMktVefCode) + 1) As Integer
                End If
            Next ilLoop
        ElseIf igPostType = 4 Then  'Cluster- only vehicles with selected market
            slNameCode = tmMktCode(imMarketIndex).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imMktCode = Val(slCode)
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If tgMVef(ilLoop).iMnfVehGp3Mkt = imMktCode Then
                    imMktVefCode(UBound(imMktVefCode)) = tgMVef(ilLoop).iCode
                    ReDim Preserve imMktVefCode(0 To UBound(imMktVefCode) + 1) As Integer
                End If
            Next ilLoop
        ElseIf (igPostType = 3) Or (igPostType = 1) Or ((igPostType = 5) And ((Asc(tgSpf.sAutoType2) And RN_REP) <> RN_REP)) Or (igPostType = 6) Then    'Post Received or Post by vehicle, only selected vehicle
            ReDim imMktVefCode(0 To 1) As Integer
            slNameCode = tmVehicle(imMarketIndex).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imMktVefCode(0) = Val(slCode)
            'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '    If tgMVef(ilLoop).iCode = imMktVefCode(0) Then
                ilLoop = gBinarySearchVef(imMktVefCode(0))
                If ilLoop <> -1 Then
                    '9/11/06- Place vehicle code into field instead of market reference.
                    '         This field is only used to know that contracts can be read mGetCntr
                    '         It is checked that value exist, it does not carry what is in the field
                    'imMktCode = tgMVef(ilLoop).iMnfVehGp3Mkt
                    imMktCode = tgMVef(ilLoop).iCode
            '        Exit For
                End If
            'Next ilLoop
        ElseIf igPostType = 2 Then  'Post by Advertiser, all Rep vehicles allowed
            ReDim imMktVefCode(0 To UBound(igRepVefCode)) As Integer
            For ilLoop = LBound(igRepVefCode) To UBound(igRepVefCode) - 1 Step 1
                ilVef = gBinarySearchVef(igRepVefCode(ilLoop))
                If ilVef <> -1 Then
                    If tgMVef(ilVef).iNrfCode = 0 Then
                        imMktVefCode(ilLoop) = igRepVefCode(ilLoop)
                    End If
                End If
            Next ilLoop
            'Place advertiser code into imMktCode variable
            slNameCode = tmAdvertiser(imMarketIndex).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imMktCode = Val(slCode)
        Else
            imMktCode = -1
        End If
    Else
        imMktCode = -1
    End If
    igBrowserReturn = 0
    If (imInvDateIndex >= 0) And (imMarketIndex >= 0) Then
        If Not mReadSbfRec(True) Then
            ilSbfExist = False
            'Remove automatic import- Jim request on 4/26/02
            'Screen.MousePointer = vbDefault
            ''gObtainYearMonthDayStr smStartStd, True, slFYear, slFMonth, slFDay
            'gObtainYearMonthDayStr smEndStd, True, slFYear, slFMonth, slFDay
            'slFMonth = Left$(cbcInvDate.List(imInvDateIndex), 3)
            'igBrowserType = 7  'Mask
            '''sgBrowseMaskFile = "F" & Right$(slFYear, 2) & slFMonth & slFDay & "?.I??"
            ''sgBrowseMaskFile = "?" & right$(slFYear, 2) & slFMonth & slFDay & "?.I??"
            'sgBrowseMaskFile = slFMonth & Right$(slFYear, 2) & "In?.??"
            'sgBrowserTitle = "Import for " & cbcNames.List(imMarketIndex)
            'Browser.Show vbModal
            'sgBrowserTitle = ""
            cmcClear.Enabled = False
        Else
            igBrowserReturn = 0
            ilSbfExist = True
            'cmcClear.Enabled = True
        End If
    Else
        igBrowserReturn = 0
    End If
    Screen.MousePointer = vbHourglass
    'Get contracts for market
    mGetCntr
    If UBound(smSave, 2) > LBONE Then
        If igPostType <> 4 Then  '4=Cluster- only vehicles with selected market
            imChg = True    'Set as changed so that contracts can be saved
        End If
    End If
    'Populate vehicle list box
    If igPostType = 4 Then  'Cluster- only vehicles with selected market
        mVehPop 0
    End If
    'Get previously entered SBF for market
    If (imInvDateIndex >= 0) And (imMarketIndex >= 0) Then
        mAddGrandTotalLine
    End If
    ilRet = mReadSbfRec(False)
    'Get carry from previous month
    mGetCarry
    pbcPostRep.Cls
    'If (imInvDateIndex >= 0) And (imMarketIndex >= 0) Then
    '    mAddGrandTotalLine
    'End If
    If igBrowserReturn = 1 Then
        slMsgFile = sgBrowserFile
        'If InStr(slMsgFile, ":") = 0 Then
        If (InStr(slMsgFile, ":") = 0) And (Left$(slMsgFile, 2) <> "\\") Then
            slMsgFile = sgImportPath & slMsgFile
        End If
        ilRet = mOpenMsgFile(slMsgFile)
        If Not ilRet Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Print #hmMsg, "Import " & sgBrowserFile & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        mRemoveAirCount True
        pbcPostRep.Cls
        ilRet = mReadImportFile(sgBrowserFile)
        If ilRet Then
            Print #hmMsg, "Import Finish Successfully"
            Close #hmMsg
        Else
            Print #hmMsg, "** Import Errors or Terminated " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            MsgBox "See " & slMsgFile & " for errors related to Rejected Records"
        End If
    Else
        'Retain air count Jim request on 4/24/02 along with not importing
        ''Remove Aired count if not previously defined and not imported
        'mRemoveAirCount False
        pbcPostRep.Cls
    End If
    'pbcPostRep.Cls
    'Compute totals
    mRecomputeTotals
    vbcPostRep.Min = LBONE  'LBound(smSave, 2)
    If UBound(smSave, 2) <= vbcPostRep.LargeChange Then
        vbcPostRep.Max = LBONE  'LBound(smSave, 2)
    Else
        vbcPostRep.Max = UBound(smSave, 2) - vbcPostRep.LargeChange
    End If
    If vbcPostRep.Value = vbcPostRep.Min Then
        pbcPostRep_Paint
    Else
        vbcPostRep.Value = vbcPostRep.Min
    End If
    If (imInvDateIndex >= 0) And (imMarketIndex >= 0) Then
        cmcImport.Enabled = True
    Else
        cmcImport.Enabled = False
    End If
    If imChg Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    If (ilSbfExist) And (igPostType = 4) Then
        cmcClear.Enabled = True
        For ilLoop = LBONE To UBound(smSave, 2) - 1 Step 1
            If (Trim$(smInfo(1, ilLoop)) = "I") Or (Trim$(smInfo(1, ilLoop)) = "F") Or (Trim$(smInfo(1, ilLoop)) = "S") Then
                If Trim$(smSave(8, ilLoop)) = "Y" Then
                    cmcClear.Enabled = False
                    Exit For
                End If
            End If
        Next ilLoop
    Else
        cmcClear.Enabled = False
    End If
    mShowCalDates
    Screen.MousePointer = vbDefault
End Sub

Private Sub vbcPostRep_Change()
    If imSettingValue Then
        pbcPostRep.Cls
        pbcPostRep_Paint
        imSettingValue = False
    Else
        mWklySetShow imWklyBoxNo
        mSetShow imBoxNo
        pbcPostRep.Cls
        pbcPostRep_Paint
        mEnableBox imBoxNo
    End If
End Sub
Private Sub vbcPostRep_DragDrop(Source As control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub vbcPostRep_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Function mRemoveBlanks(slInStr As String) As String
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slChar As String * 1

    slStr = ""
    For ilLoop = 1 To Len(slInStr) Step 1
        slChar = Mid$(slInStr, ilLoop, 1)
        If slChar <> " " Then
            slStr = slStr & slChar
        End If
    Next ilLoop
    mRemoveBlanks = slStr
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetCarry                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get spots carried from previous*
'*                      month                          *
'*                                                     *
'*******************************************************
Private Sub mGetCarry()
    Dim ilRet As Integer    'Return status
    Dim llDate As Long
    Dim llStartStd As Long
    Dim llEndStd As Long
    Dim slDate As String
    Dim ilRow As Integer
    Dim ilPass As Integer
    Dim ilFound As Integer

    If (imMarketIndex >= 0) And (imInvDateIndex >= 0) Then
        slDate = gDecOneDay(smStartStd)
        slDate = gObtainStartStd(slDate)
        llStartStd = gDateValue(slDate)
        slDate = gObtainEndStd(slDate)
        llEndStd = gDateValue(slDate)
        tmSbfSrchKey2.sTranType = "T"
        'tmSbfSrchKey2.iDate(0) = 0
        'tmSbfSrchKey2.iDate(1) = 0
        gPackDateLong llStartStd, tmSbfSrchKey2.iDate(0), tmSbfSrchKey2.iDate(1)
        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.sTranType = "T")
            gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
            If llDate > llEndStd Then
                Exit Do
            End If
            If (tmSbf.iMissCarryOver < 0) Or ((tmSbf.iCalCarryOver > 0) And (tgSpf.sPostCalAff = "C") And (igPostType <= 2)) Or ((tmSbf.iCalCarryOver > 0) And ((igPostType = 5) Or (igPostType = 6))) Or ((tmSbf.iCalCarryBonus > 0) And ((igPostType = 5) Or (igPostType = 6))) Then
                ilFound = False
                For ilPass = 0 To 1 Step 1
                    For ilRow = LBONE To UBound(smSave, 2) - 1 Step 1
                        If (InStr(1, smShow(imTOTALINDEX, ilRow), "Total:", 1) <= 0) Then
                            If (igPostType = 5) Or (igPostType = 6) Then
                                If (tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.iAirVefCode = imSave(2, ilRow)) And (tmSbf.sCashTrade = smSave(1, ilRow)) And (tmSbf.iLineNo = imSave(7, ilRow)) Then
                                    If (tmSbf.iMissCarryOver < 0) Then
                                        If Trim$(smSave(10, ilRow)) = "" Then
                                            smSave(10, ilRow) = Trim$(str$(tmSbf.iMissCarryOver))
                                        Else
                                            smSave(10, ilRow) = Trim$(str$(Val(smSave(10, ilRow)) + tmSbf.iMissCarryOver))
                                        End If
                                    End If
                                    If (tmSbf.iCalCarryOver > 0) Then
                                        If Trim$(smSave(9, ilRow)) = "" Then
                                            smSave(9, ilRow) = Trim$(str$(tmSbf.iCalCarryOver))
                                        Else
                                            smSave(9, ilRow) = Trim$(str$(Val(smSave(9, ilRow)) + tmSbf.iCalCarryOver))
                                        End If
                                    End If
                                    If (tmSbf.iCalCarryBonus > 0) Then
                                        If imSave(10, ilRow) = 0 Then
                                            imSave(10, ilRow) = tmSbf.iCalCarryBonus
                                        Else
                                            imSave(10, ilRow) = imSave(10, ilRow) + tmSbf.iCalCarryBonus
                                        End If
                                    End If
                                    ilFound = True
                                    Exit For
                                End If
                            Else
                                If (tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.iAirVefCode = imSave(2, ilRow)) And (tmSbf.sCashTrade = smSave(1, ilRow)) And ((tmSbf.lSpotPrice = lmSave(3, ilRow)) Or (ilPass = 1)) Then
                                    If (tmSbf.iMissCarryOver < 0) Then
                                        If Trim$(smSave(10, ilRow)) = "" Then
                                            smSave(10, ilRow) = Trim$(str$(tmSbf.iMissCarryOver))
                                        Else
                                            smSave(10, ilRow) = Trim$(str$(Val(smSave(10, ilRow)) + tmSbf.iMissCarryOver))
                                        End If
                                    End If
                                    If ((tmSbf.iCalCarryOver > 0) And (tgSpf.sPostCalAff = "C") And (igPostType <= 2)) Then
                                        If Trim$(smSave(9, ilRow)) = "" Then
                                            smSave(9, ilRow) = Trim$(str$(tmSbf.iCalCarryOver))
                                        Else
                                            smSave(9, ilRow) = Trim$(str$(Val(smSave(9, ilRow)) + tmSbf.iCalCarryOver))
                                        End If
                                    End If
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next ilRow
                    If ilFound Then
                        Exit For
                    End If
                Next ilPass
            End If
            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        For ilRow = LBONE To UBound(smSave, 2) - 1 Step 1
            If (InStr(1, smShow(imTOTALINDEX, ilRow), "Total:", 1) <= 0) Then
                If imAPREVNOSPOTSINDEX > 0 Then
                    gSetShow pbcPostRep, smSave(9, ilRow), tmCtrls(imAPREVNOSPOTSINDEX)
                    smShow(imAPREVNOSPOTSINDEX, ilRow) = tmCtrls(imAPREVNOSPOTSINDEX).sShow
                End If
                If imPRIORNOSPOTSINDEX > 0 Then
                    gSetShow pbcPostRep, smSave(10, ilRow), tmCtrls(imPRIORNOSPOTSINDEX)
                    smShow(imPRIORNOSPOTSINDEX, ilRow) = tmCtrls(imPRIORNOSPOTSINDEX).sShow
                End If
            End If
        Next ilRow
    End If
    Exit Sub

    On Error GoTo 0
    Exit Sub

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mShowCalDates                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show Calendar dates related to *
'*                      standard dates                 *
'*                                                     *
'*******************************************************
Private Sub mShowCalDates()
    Dim ilAdf As Integer

    If (tgSpf.sPostCalAff = "C") And (igPostType <= 2) Or (igPostType = 5) Or (igPostType = 6) Then
        'Compute previous month dates and next month dates
        lacCalDates.Caption = ""
        'Carry from previous month: StartStd to StartCal-1
        If lmStartStd < lmStartCal Then
            lacCalDates.Caption = "Aired Prev: " & Format$(lmStartStd, "mm/dd") & "-" & Format$(lmStartCal - 1, "mm/dd")
        Else
            lacCalDates.Caption = "Aired Prev: No Dates"
        End If
        'Part of next month std: EndStd+1 To EndCal
        If lmEndStd < lmEndCal Then
            lacCalDates.Caption = lacCalDates.Caption & " Aired Next: " & Format$(lmEndStd + 1, "mm/dd") & "-" & Format$(lmEndCal, "mm/dd")
        Else
            lacCalDates.Caption = lacCalDates.Caption & " Aired Next: No Dates"
        End If
        If igPostType = 1 Then
            lacCalDates.Caption = lacCalDates.Caption & ".   * indicates no MGs allowed"
        Else
            'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
            '    If tgCommAdf(ilAdf).iCode = imMktCode Then
                ilAdf = gBinarySearchAdf(imMktCode)
                If ilAdf <> -1 Then
                    If tgCommAdf(ilAdf).sAllowRepMG = "N" Then
                        lacCalDates.Caption = lacCalDates.Caption & ".   No MGs allowed"
                    Else
                        lacCalDates.Caption = lacCalDates.Caption & ".   MGs allowed"
                    End If
            '        Exit For
                End If
            'Next ilAdf
        End If
        lacCalDates.Visible = True
    ElseIf (igPostType = 5) Or (igPostType = 6) Then
        'Compute previous month dates and next month dates
        lacCalDates.Caption = ""
        'Carry from previous month: StartStd to StartCal-1
        If lmStartStd < lmStartCal Then
            lacCalDates.Caption = "Prior: " & Format$(lmStartStd, "mm/dd") & "-" & Format$(lmStartCal - 1, "mm/dd")
        Else
            lacCalDates.Caption = "Prior: No Dates"
        End If
        'Part of next month std: EndStd+1 To EndCal
        If lmEndStd < lmEndCal Then
            lacCalDates.Caption = lacCalDates.Caption & " Next: " & Format$(lmEndStd + 1, "mm/dd") & "-" & Format$(lmEndCal, "mm/dd")
        Else
            lacCalDates.Caption = lacCalDates.Caption & " Next: No Dates"
        End If
        If igPostType = 1 Then
            lacCalDates.Caption = lacCalDates.Caption & ".   * indicates no MGs allowed"
        Else
            'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
            '    If tgCommAdf(ilAdf).iCode = imMktCode Then
                ilAdf = gBinarySearchAdf(imMktCode)
                If ilAdf <> -1 Then
                    If tgCommAdf(ilAdf).sAllowRepMG = "N" Then
                        lacCalDates.Caption = lacCalDates.Caption & ".   No MGs allowed"
                    Else
                        lacCalDates.Caption = lacCalDates.Caption & ".   MGs allowed"
                    End If
            '        Exit For
                End If
            'Next ilAdf
        End If
        lacCalDates.Visible = True
    Else
        If igPostType = 3 Then
            lacCalDates.Visible = False
        ElseIf igPostType = 2 Then
            lacCalDates.Visible = False
            'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
            '    If tgCommAdf(ilAdf).iCode = imMktCode Then
                ilAdf = gBinarySearchAdf(imMktCode)
                If ilAdf <> -1 Then
                    If tgCommAdf(ilAdf).sAllowRepMG = "N" Then
                        lacCalDates.Caption = "No MGs allowed"
                    Else
                        lacCalDates.Caption = "MGs allowed"
                    End If
                    lacCalDates.Visible = True
            '        Exit For
                End If
            'Next ilAdf
        Else
            lacCalDates.Caption = "* in front of advertiser name indicates No MGs allowed"
            lacCalDates.Visible = True
        End If
    End If
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mWklyEnableBox                  *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mWklyEnableBox(ilBoxNo As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************

'
'   mDmEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (imRowNo < LBONE) Or (imRowNo > UBound(smSave, 2)) Then
        Exit Sub
    End If
    If ilBoxNo < WKLYAIRNOINDEX Or (ilBoxNo > WKLYCARRIEDNOINDEX) Then
        Exit Sub
    End If
    If (Trim$(smSave(8, imRowNo)) = "Y") Or (Trim$(smSave(17, imRowNo)) = "Y") Then    'If billed disallow change
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case WKLYAIRNOINDEX
            edcWkly.Width = tmCtrls(imANOSPOTSINDEX).fBoxW
            gMoveTableCtrl pbcWeekly, edcWkly, tmWklyCtrls(WKLYAIRNOINDEX).fBoxX, tmWklyCtrls(WKLYAIRNOINDEX).fBoxY + (imWklyRowNo - 1) * (fgBoxGridH + 15)
            edcWkly.Text = Trim$(smSave(18 + imWklyRowNo - 1, imRowNo))
            edcWkly.Visible = True  'Set visibility
            edcWkly.SetFocus
        Case WKLYCARRIEDNOINDEX
            edcWkly.Width = tmCtrls(imANOSPOTSINDEX).fBoxW
            gMoveTableCtrl pbcWeekly, edcWkly, tmWklyCtrls(WKLYCARRIEDNOINDEX).fBoxX, tmWklyCtrls(WKLYCARRIEDNOINDEX).fBoxY + (imWklyRowNo - 1) * (fgBoxGridH + 15)
            edcWkly.Text = Trim$(smSave(23 + imWklyRowNo - 1, imRowNo))
            edcWkly.Visible = True  'Set visibility
            edcWkly.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mWklySetChg                     *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mWklySetChg(ilBoxNo As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilIndex                                                 *
'******************************************************************************************

'
'   mDmSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'

    If (imRowNo < LBONE) Or (imRowNo > UBound(smSave, 2)) Then
        Exit Sub
    End If
    If ilBoxNo < WKLYAIRNOINDEX Or (ilBoxNo > WKLYCARRIEDNOINDEX) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case WKLYAIRNOINDEX
            If (Trim$(smSave(18 + imWklyRowNo - 1, imRowNo)) = "") And (Trim$(edcWkly.Text) <> "") Then
                imChg = True
            Else
                If Val(edcWkly.Text) <> Val(smSave(18 + imWklyRowNo - 1, imRowNo)) Then
                    imChg = True
                End If
            End If
        Case WKLYCARRIEDNOINDEX
            If (Trim$(smSave(23 + imWklyRowNo - 1, imRowNo)) = "") And (Trim$(edcWkly.Text) <> "") Then
                imChg = True
            Else
                If Val(edcWkly.Text) <> Val(smSave(23 + imWklyRowNo - 1, imRowNo)) Then
                    imChg = True
                End If
            End If
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mWklySetShow                    *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mWklySetShow(ilBoxNo As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                       slStr                                                   *
'******************************************************************************************

'
'   mDmSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    If (imRowNo < LBONE) Or (imRowNo > UBound(smSave, 2)) Then
        Exit Sub
    End If
    If (imWklyRowNo < 1) Or (imWklyRowNo > 5) Then
        Exit Sub
    End If
    If ilBoxNo < WKLYAIRNOINDEX Or (ilBoxNo > WKLYCARRIEDNOINDEX) Then
        Exit Sub
    End If
    If (smSave(8, imRowNo) = "Y") Or (smSave(17, imRowNo) = "Y") Then
        Exit Sub
    End If
    mWklySetChg ilBoxNo
    Select Case ilBoxNo 'Branch on box type (control)
        Case WKLYAIRNOINDEX
            edcWkly.Visible = False
            smSave(18 + imWklyRowNo - 1, imRowNo) = Trim$(edcWkly.Text)
        Case WKLYCARRIEDNOINDEX
            edcWkly.Visible = False
            smSave(23 + imWklyRowNo - 1, imRowNo) = Trim$(edcWkly.Text)
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintLnTitle                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint Header Titles            *
'*                                                     *
'*******************************************************
Private Sub mPaintRepTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer
    Dim llWidth As Long

    ilHalfY = tmCtrls(1).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop

    llColor = pbcPostRep.ForeColor
    slFontName = pbcPostRep.FontName
    flFontSize = pbcPostRep.FontSize
    ilFillStyle = pbcPostRep.FillStyle
    llFillColor = pbcPostRep.FillColor
    pbcPostRep.ForeColor = BLUE
    pbcPostRep.FontBold = False
    pbcPostRep.FontSize = 7
    pbcPostRep.FontName = "Arial"
    pbcPostRep.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    If igPostType = 4 Then
        pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX - 15, 15)-Step(tmCtrls(imCONTRACTINDEX).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX, 30)-Step(tmCtrls(imCONTRACTINDEX).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imCONTRACTINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Contract"
        pbcPostRep.Line (tmCtrls(imCashTradeIndex).fBoxX - 15, 15)-Step(tmCtrls(imCashTradeIndex).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imCashTradeIndex).fBoxX, 30)-Step(tmCtrls(imCashTradeIndex).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imCashTradeIndex).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "C/T"
        pbcPostRep.Line (tmCtrls(imAdvtIndex).fBoxX - 15, 15)-Step(tmCtrls(imAdvtIndex).fBoxW + 15, tmCtrls(imAdvtIndex).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imAdvtIndex).fBoxX, 30)-Step(tmCtrls(imAdvtIndex).fBoxW - 15, tmCtrls(imAdvtIndex).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imAdvtIndex).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Advertiser"
        pbcPostRep.Line (tmCtrls(imVEHICLEINDEX).fBoxX - 15, 15)-Step(tmCtrls(imVEHICLEINDEX).fBoxW + 15, tmCtrls(imVEHICLEINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imVEHICLEINDEX).fBoxX, 30)-Step(tmCtrls(imVEHICLEINDEX).fBoxW - 15, tmCtrls(imVEHICLEINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imVEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Vehicle"
        llWidth = tmCtrls(imANOSPOTSINDEX).fBoxX - tmCtrls(imPRIORNOSPOTSINDEX).fBoxX
        pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX - 15, 15)-Step(llWidth, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imPRIORNOSPOTSINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Ordered") 'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Ordered"
        pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imPRIORNOSPOTSINDEX).fBoxW + 15, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imPRIORNOSPOTSINDEX).fBoxW - 15, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imPRIORNOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Prior"
        pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW + 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW - 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imONOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Spot"
        pbcPostRep.Line (tmCtrls(imOPRICEINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imOPRICEINDEX).fBoxW + 15, tmCtrls(imOPRICEINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imOPRICEINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imOPRICEINDEX).fBoxW - 15, tmCtrls(imOPRICEINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imOPRICEINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Price"
        pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imOGROSSINDEX).fBoxW + 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imOGROSSINDEX).fBoxW - 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imOGROSSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Gross"
        pbcPostRep.Line (tmCtrls(imANOSPOTSINDEX).fBoxX - 15, 15)-Step(tmCtrls(imANOSPOTSINDEX).fBoxW + 15, tmCtrls(imANOSPOTSINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.CurrentX = tmCtrls(imANOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Aired"
        pbcPostRep.CurrentX = tmCtrls(imANOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Spot"

        llWidth = tmCtrls(imDGROSSINDEX).fBoxX + tmCtrls(imDGROSSINDEX).fBoxW - tmCtrls(imMISSEDINDEX).fBoxX
        pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX - 15, 15)-Step(llWidth + 15, tmCtrls(imMISSEDINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imMISSEDINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imMISSEDINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Results") 'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Results"
        pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imMISSEDINDEX).fBoxW + 15, tmCtrls(imMISSEDINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imMISSEDINDEX).fBoxW - 15, tmCtrls(imMISSEDINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imMISSEDINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Miss"
        pbcPostRep.Line (tmCtrls(imMGINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imMGINDEX).fBoxW + 15, tmCtrls(imMGINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imMGINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imMGINDEX).fBoxW - 15, tmCtrls(imMGINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imMGINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "MG"
        pbcPostRep.Line (tmCtrls(imBONUSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imBONUSINDEX).fBoxW + 15, tmCtrls(imBONUSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imBONUSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imBONUSINDEX).fBoxW - 15, tmCtrls(imBONUSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imBONUSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Bonus"
        pbcPostRep.Line (tmCtrls(imDGROSSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imDGROSSINDEX).fBoxW + 15, tmCtrls(imDGROSSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imDGROSSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imDGROSSINDEX).fBoxW - 15, tmCtrls(imDGROSSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imDGROSSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Difference"


    ElseIf igPostType = 3 Then
        pbcPostRep.Line (tmCtrls(imCHECKINDEX).fBoxX - 15, 15)-Step(tmCtrls(imCHECKINDEX).fBoxW + 15, tmCtrls(imCHECKINDEX).fBoxY - 30), BLUE, B
        'pbcPostRep.Line (tmCtrls(imCHECKINDEX).fBoxX, 30)-Step(tmCtrls(imCHECKINDEX).fBoxW - 15, tmCtrls(imCHECKINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imCHECKINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 45 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.FontName = "Monotype Sorts"
        pbcPostRep.FontBold = False
        pbcPostRep.Print " 4"
        pbcPostRep.FontBold = False
        pbcPostRep.FontSize = 7
        pbcPostRep.FontName = "Arial"
        pbcPostRep.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX - 15, 15)-Step(tmCtrls(imCONTRACTINDEX).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX, 30)-Step(tmCtrls(imCONTRACTINDEX).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imCONTRACTINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Contract"
        pbcPostRep.Line (tmCtrls(imCashTradeIndex).fBoxX - 15, 15)-Step(tmCtrls(imCashTradeIndex).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imCashTradeIndex).fBoxX, 30)-Step(tmCtrls(imCashTradeIndex).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imCashTradeIndex).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "C/T"
        pbcPostRep.Line (tmCtrls(imAdvtIndex).fBoxX - 15, 15)-Step(tmCtrls(imAdvtIndex).fBoxW + 15, tmCtrls(imAdvtIndex).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imAdvtIndex).fBoxX, 30)-Step(tmCtrls(imAdvtIndex).fBoxW - 15, tmCtrls(imAdvtIndex).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imAdvtIndex).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Advertiser"
        llWidth = tmCtrls(imRECEIVEDINDEX).fBoxX - tmCtrls(imPRIORNOSPOTSINDEX).fBoxX
        pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX - 15, 15)-Step(llWidth, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imPRIORNOSPOTSINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Ordered") 'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Ordered"
        pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imPRIORNOSPOTSINDEX).fBoxW + 15, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imPRIORNOSPOTSINDEX).fBoxW - 15, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imPRIORNOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Prior"
        pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW + 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW - 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imONOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Spot"
        pbcPostRep.Line (tmCtrls(imOPRICEINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imOPRICEINDEX).fBoxW + 15, tmCtrls(imOPRICEINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imOPRICEINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imOPRICEINDEX).fBoxW - 15, tmCtrls(imOPRICEINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imOPRICEINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Price"
        pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imOGROSSINDEX).fBoxW + 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imOGROSSINDEX).fBoxW - 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imOGROSSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Gross"
        pbcPostRep.Line (tmCtrls(imRECEIVEDINDEX).fBoxX - 15, 15)-Step(tmCtrls(imRECEIVEDINDEX).fBoxW + 15, tmCtrls(imRECEIVEDINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imRECEIVEDINDEX).fBoxX, 30)-Step(tmCtrls(imRECEIVEDINDEX).fBoxW - 15, tmCtrls(imRECEIVEDINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imRECEIVEDINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Received"
        pbcPostRep.Line (tmCtrls(imPOSTEDINDEX).fBoxX - 15, 15)-Step(tmCtrls(imPOSTEDINDEX).fBoxW + 15, tmCtrls(imPOSTEDINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imPOSTEDINDEX).fBoxX, 30)-Step(tmCtrls(imPOSTEDINDEX).fBoxW - 15, tmCtrls(imPOSTEDINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imPOSTEDINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Posted"
    ElseIf (igPostType = 5) Or (igPostType = 6) Then
        pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX - 15, 15)-Step(tmCtrls(imCONTRACTINDEX).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX, 30)-Step(tmCtrls(imCONTRACTINDEX).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imCONTRACTINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Contract"
        pbcPostRep.Line (tmCtrls(imLINEINDEX).fBoxX - 15, 15)-Step(tmCtrls(imLINEINDEX).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imLINEINDEX).fBoxX, 30)-Step(tmCtrls(imLINEINDEX).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imLINEINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Line"
        pbcPostRep.Line (tmCtrls(imADVTVEHDAYPARTINDEX).fBoxX - 15, 15)-Step(tmCtrls(imADVTVEHDAYPARTINDEX).fBoxW + 15, tmCtrls(imADVTVEHDAYPARTINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imADVTVEHDAYPARTINDEX).fBoxX, 30)-Step(tmCtrls(imADVTVEHDAYPARTINDEX).fBoxW - 15, tmCtrls(imADVTVEHDAYPARTINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imADVTVEHDAYPARTINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Advertiser, Vehicle, Daypart"
        pbcPostRep.Line (tmCtrls(imLENGTHINDEX).fBoxX - 15, 15)-Step(tmCtrls(imLENGTHINDEX).fBoxW + 15, tmCtrls(imLENGTHINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imLENGTHINDEX).fBoxX, 30)-Step(tmCtrls(imLENGTHINDEX).fBoxW - 15, tmCtrls(imLENGTHINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imLENGTHINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Len"
        llWidth = tmCtrls(imANOSPOTSPRIORINDEX).fBoxX - tmCtrls(imONOSPOTSINDEX).fBoxX
        pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX - 15, 15)-Step(llWidth, tmCtrls(imONOSPOTSINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imONOSPOTSINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imONOSPOTSINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Ordered") / 2 'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Ordered"
        pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW + 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW - 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imONOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Spot"
        pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imOGROSSINDEX).fBoxW + 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imOGROSSINDEX).fBoxW - 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imOGROSSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Gross"

        llWidth = tmCtrls(imDIFFSPOTSINDEX).fBoxX - tmCtrls(imANOSPOTSPRIORINDEX).fBoxX
        pbcPostRep.Line (tmCtrls(imANOSPOTSPRIORINDEX).fBoxX - 15, 15)-Step(llWidth, tmCtrls(imANOSPOTSPRIORINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imANOSPOTSPRIORINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imANOSPOTSPRIORINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imANOSPOTSPRIORINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Paid Spot") / 2 'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Paid Spot"
        pbcPostRep.Line (tmCtrls(imANOSPOTSPRIORINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imANOSPOTSPRIORINDEX).fBoxW + 15, tmCtrls(imANOSPOTSPRIORINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imANOSPOTSPRIORINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imANOSPOTSPRIORINDEX).fBoxW - 15, tmCtrls(imANOSPOTSPRIORINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imANOSPOTSPRIORINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Prior"
        pbcPostRep.Line (tmCtrls(imANOSPOTSCURRINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imANOSPOTSCURRINDEX).fBoxW + 15, tmCtrls(imANOSPOTSCURRINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imANOSPOTSCURRINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imANOSPOTSCURRINDEX).fBoxW - 15, tmCtrls(imANOSPOTSCURRINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imANOSPOTSCURRINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Curr"

        llWidth = tmCtrls(imBONUSPREVINDEX).fBoxX - tmCtrls(imDIFFSPOTSINDEX).fBoxX
        pbcPostRep.Line (tmCtrls(imDIFFSPOTSINDEX).fBoxX - 15, 15)-Step(llWidth, tmCtrls(imDIFFSPOTSINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imDIFFSPOTSINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imDIFFSPOTSINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imDIFFSPOTSINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Difference") / 2 'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Difference"
        pbcPostRep.Line (tmCtrls(imDIFFSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imDIFFSPOTSINDEX).fBoxW + 15, tmCtrls(imDIFFSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imDIFFSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imDIFFSPOTSINDEX).fBoxW - 15, tmCtrls(imDIFFSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imDIFFSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Spot"
        pbcPostRep.Line (tmCtrls(imDIFFGROSSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imDIFFGROSSINDEX).fBoxW + 15, tmCtrls(imDIFFGROSSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imDIFFGROSSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imDIFFGROSSINDEX).fBoxW - 15, tmCtrls(imDIFFGROSSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imDIFFGROSSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Gross"

        llWidth = tmCtrls(imNNOSPOTSINDEX).fBoxX - tmCtrls(imBONUSPREVINDEX).fBoxX
        pbcPostRep.Line (tmCtrls(imBONUSPREVINDEX).fBoxX - 15, 15)-Step(llWidth, tmCtrls(imBONUSPREVINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imBONUSPREVINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imBONUSPREVINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imBONUSPREVINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Bonus") / 2 'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Bonus"
        pbcPostRep.Line (tmCtrls(imBONUSPREVINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imBONUSPREVINDEX).fBoxW + 15, tmCtrls(imBONUSPREVINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imBONUSPREVINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imBONUSPREVINDEX).fBoxW - 15, tmCtrls(imBONUSPREVINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imBONUSPREVINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Prior"
        pbcPostRep.Line (tmCtrls(imBONUSCURRINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imBONUSCURRINDEX).fBoxW + 15, tmCtrls(imBONUSCURRINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imBONUSCURRINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imBONUSCURRINDEX).fBoxW - 15, tmCtrls(imBONUSCURRINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imBONUSCURRINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Curr"

        llWidth = tmCtrls(imNNOSPOTSINDEX).fBoxW + tmCtrls(imNNOBONUSINDEX).fBoxW + 30
        pbcPostRep.Line (tmCtrls(imNNOSPOTSINDEX).fBoxX - 15, 15)-Step(llWidth, tmCtrls(imNNOSPOTSINDEX).fBoxY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imNNOSPOTSINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imNNOSPOTSINDEX).fBoxY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imNNOSPOTSINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Next Month") / 2 'fgBoxInsetX
        pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPostRep.Print "Next Month"
        pbcPostRep.Line (tmCtrls(imNNOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imNNOSPOTSINDEX).fBoxW + 15, tmCtrls(imNNOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imNNOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imNNOSPOTSINDEX).fBoxW - 15, tmCtrls(imNNOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imNNOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Spot"
        pbcPostRep.Line (tmCtrls(imNNOBONUSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imNNOBONUSINDEX).fBoxW + 15, tmCtrls(imNNOBONUSINDEX).fBoxY - ilHalfY - 30), BLUE, B
        pbcPostRep.Line (tmCtrls(imNNOBONUSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imNNOBONUSINDEX).fBoxW - 15, tmCtrls(imNNOBONUSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
        pbcPostRep.CurrentX = tmCtrls(imNNOBONUSINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPostRep.CurrentY = ilHalfY + 15
        pbcPostRep.Print "Bonus"
    Else
        If tgSpf.sPostCalAff = "C" Then
            pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX - 15, 15)-Step(tmCtrls(imCONTRACTINDEX).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX, 30)-Step(tmCtrls(imCONTRACTINDEX).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imCONTRACTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Contract"
            pbcPostRep.Line (tmCtrls(imCashTradeIndex).fBoxX - 15, 15)-Step(tmCtrls(imCashTradeIndex).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imCashTradeIndex).fBoxX, 30)-Step(tmCtrls(imCashTradeIndex).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imCashTradeIndex).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "C/T"
            If imAdvtIndex > 0 Then
                pbcPostRep.Line (tmCtrls(imAdvtIndex).fBoxX - 15, 15)-Step(tmCtrls(imAdvtIndex).fBoxW + 15, tmCtrls(imAdvtIndex).fBoxY - 30), BLUE, B
                pbcPostRep.Line (tmCtrls(imAdvtIndex).fBoxX, 30)-Step(tmCtrls(imAdvtIndex).fBoxW - 15, tmCtrls(imAdvtIndex).fBoxY - 60), LIGHTYELLOW, BF
                pbcPostRep.CurrentX = tmCtrls(imAdvtIndex).fBoxX + 15  'fgBoxInsetX
                pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
                pbcPostRep.Print "Advertiser"
            Else
                pbcPostRep.Line (tmCtrls(imVEHICLEINDEX).fBoxX - 15, 15)-Step(tmCtrls(imVEHICLEINDEX).fBoxW + 15, tmCtrls(imVEHICLEINDEX).fBoxY - 30), BLUE, B
                pbcPostRep.Line (tmCtrls(imVEHICLEINDEX).fBoxX, 30)-Step(tmCtrls(imVEHICLEINDEX).fBoxW - 15, tmCtrls(imVEHICLEINDEX).fBoxY - 60), LIGHTYELLOW, BF
                pbcPostRep.CurrentX = tmCtrls(imVEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
                pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
                pbcPostRep.Print "Vehicle"
           End If
            llWidth = tmCtrls(imAPREVNOSPOTSINDEX).fBoxX - tmCtrls(imPRIORNOSPOTSINDEX).fBoxX
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX - 15, 15)-Step(llWidth, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imPRIORNOSPOTSINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Ordered") 'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Ordered"
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imPRIORNOSPOTSINDEX).fBoxW + 15, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imPRIORNOSPOTSINDEX).fBoxW - 15, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imPRIORNOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Prior"
            pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW + 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW - 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imONOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Spot"
            pbcPostRep.Line (tmCtrls(imOPRICEINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imOPRICEINDEX).fBoxW + 15, tmCtrls(imOPRICEINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imOPRICEINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imOPRICEINDEX).fBoxW - 15, tmCtrls(imOPRICEINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imOPRICEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Price"
            pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imOGROSSINDEX).fBoxW + 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imOGROSSINDEX).fBoxW - 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imOGROSSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Gross"
            llWidth = tmCtrls(imMISSEDINDEX).fBoxX - tmCtrls(imAPREVNOSPOTSINDEX).fBoxX
            pbcPostRep.Line (tmCtrls(imAPREVNOSPOTSINDEX).fBoxX - 15, 15)-Step(llWidth, tmCtrls(imAPREVNOSPOTSINDEX).fBoxY - 30), BLUE, B
            'pbcPostRep.Line (tmCtrls(imAPREVNOSPOTSINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imAPREVNOSPOTSINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imAPREVNOSPOTSINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Aired") 'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Aired"
            pbcPostRep.Line (tmCtrls(imAPREVNOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imAPREVNOSPOTSINDEX).fBoxW + 15, tmCtrls(imAPREVNOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imAPREVNOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imAPREVNOSPOTSINDEX).fBoxW - 15, tmCtrls(imAPREVNOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imAPREVNOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Prev"
            pbcPostRep.Line (tmCtrls(imANOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imANOSPOTSINDEX).fBoxW + 15, tmCtrls(imANOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            'pbcPostRep.Line (tmCtrls(imANOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imANOSPOTSINDEX).fBoxW - 15, tmCtrls(imANOSPOTSINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imANOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Spot"
            pbcPostRep.Line (tmCtrls(imANEXTNOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imANEXTNOSPOTSINDEX).fBoxW + 15, tmCtrls(imANEXTNOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            'pbcPostRep.Line (tmCtrls(imANEXTNOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imANEXTNOSPOTSINDEX).fBoxW - 15, tmCtrls(imANEXTNOSPOTSINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imANEXTNOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Next"

            llWidth = tmCtrls(imDGROSSINDEX).fBoxX + tmCtrls(imDGROSSINDEX).fBoxW - tmCtrls(imMISSEDINDEX).fBoxX
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX - 15, 15)-Step(llWidth + 15, tmCtrls(imMISSEDINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imMISSEDINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imMISSEDINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Results") 'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Results"
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imMISSEDINDEX).fBoxW + 15, tmCtrls(imMISSEDINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imMISSEDINDEX).fBoxW - 15, tmCtrls(imMISSEDINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imMISSEDINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Miss"
            pbcPostRep.Line (tmCtrls(imMGINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imMGINDEX).fBoxW + 15, tmCtrls(imMGINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imMGINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imMGINDEX).fBoxW - 15, tmCtrls(imMGINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imMGINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "MG"
            pbcPostRep.Line (tmCtrls(imBONUSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imBONUSINDEX).fBoxW + 15, tmCtrls(imBONUSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imBONUSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imBONUSINDEX).fBoxW - 15, tmCtrls(imBONUSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imBONUSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Bonus"
            pbcPostRep.Line (tmCtrls(imDGROSSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imDGROSSINDEX).fBoxW + 15, tmCtrls(imDGROSSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imDGROSSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imDGROSSINDEX).fBoxW - 15, tmCtrls(imDGROSSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imDGROSSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Difference"
        ElseIf igPostType = 1 And tgSpf.sPostCalAff = "W" Then
            pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX - 15, 15)-Step(tmCtrls(imCONTRACTINDEX).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX, 30)-Step(tmCtrls(imCONTRACTINDEX).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imCONTRACTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Contract"

            pbcPostRep.Line (tmCtrls(imCashTradeIndex).fBoxX - 15, 15)-Step(tmCtrls(imCashTradeIndex).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imCashTradeIndex).fBoxX, 30)-Step(tmCtrls(imCashTradeIndex).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imCashTradeIndex).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "C/T"
            If imAdvtIndex > 0 Then
                pbcPostRep.Line (tmCtrls(imAdvtIndex).fBoxX - 15, 15)-Step(tmCtrls(imAdvtIndex).fBoxW + 15, tmCtrls(imAdvtIndex).fBoxY - 30), BLUE, B
                pbcPostRep.Line (tmCtrls(imAdvtIndex).fBoxX, 30)-Step(tmCtrls(imAdvtIndex).fBoxW - 15, tmCtrls(imAdvtIndex).fBoxY - 60), LIGHTYELLOW, BF
                pbcPostRep.CurrentX = tmCtrls(imAdvtIndex).fBoxX + 15  'fgBoxInsetX
                pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
                pbcPostRep.Print "Advertiser"
            Else
                pbcPostRep.Line (tmCtrls(imVEHICLEINDEX).fBoxX - 15, 15)-Step(tmCtrls(imVEHICLEINDEX).fBoxW + 15, tmCtrls(imVEHICLEINDEX).fBoxY - 30), BLUE, B
                pbcPostRep.Line (tmCtrls(imVEHICLEINDEX).fBoxX, 30)-Step(tmCtrls(imVEHICLEINDEX).fBoxW - 15, tmCtrls(imVEHICLEINDEX).fBoxY - 60), LIGHTYELLOW, BF
                pbcPostRep.CurrentX = tmCtrls(imVEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
                pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
                pbcPostRep.Print "Vehicle"
            End If
            'llWidth = tmCtrls(imANOSPOTSINDEX).fBoxX - tmCtrls(imPRIORNOSPOTSINDEX).fBoxX
            llWidth = tmCtrls(imINVOICENOINDEX).fBoxX - tmCtrls(imPRIORNOSPOTSINDEX).fBoxX
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX - 15, 15)-Step(llWidth, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imPRIORNOSPOTSINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Ordered") 'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Ordered"
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imPRIORNOSPOTSINDEX).fBoxW + 15, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imPRIORNOSPOTSINDEX).fBoxW - 15, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imPRIORNOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Prior"
            pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW + 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW - 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imONOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Spot"
            pbcPostRep.Line (tmCtrls(imOPRICEINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imOPRICEINDEX).fBoxW + 15, tmCtrls(imOPRICEINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imOPRICEINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imOPRICEINDEX).fBoxW - 15, tmCtrls(imOPRICEINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imOPRICEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Price"
            pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imOGROSSINDEX).fBoxW + 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imOGROSSINDEX).fBoxW - 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imOGROSSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Gross"
            
            pbcPostRep.Line (tmCtrls(imINVOICENOINDEX).fBoxX - 15, 15)-Step(tmCtrls(imINVOICENOINDEX).fBoxW + 15, tmCtrls(imINVOICENOINDEX).fBoxY - 30), BLUE, B
            'pbcPostRep.Line (tmCtrls(imINVOICENOINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imINVOICENOINDEX).fBoxW - 15, tmCtrls(imINVOICENOINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imINVOICENOINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = 15
            pbcPostRep.Print "Station"
            pbcPostRep.CurrentX = tmCtrls(imINVOICENOINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Invoice #"
            
            pbcPostRep.Line (tmCtrls(imANOSPOTSINDEX).fBoxX - 15, 15)-Step(tmCtrls(imANOSPOTSINDEX).fBoxW + 15, tmCtrls(imANOSPOTSINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.CurrentX = tmCtrls(imANOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Aired"
            pbcPostRep.CurrentX = tmCtrls(imANOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Spot"

            llWidth = tmCtrls(imDGROSSINDEX).fBoxX + tmCtrls(imDGROSSINDEX).fBoxW - tmCtrls(imMISSEDINDEX).fBoxX
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX - 15, 15)-Step(llWidth + 15, tmCtrls(imMISSEDINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imMISSEDINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imMISSEDINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Results") 'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Results"
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imMISSEDINDEX).fBoxW + 15, tmCtrls(imMISSEDINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imMISSEDINDEX).fBoxW - 15, tmCtrls(imMISSEDINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imMISSEDINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Miss"
            pbcPostRep.Line (tmCtrls(imMGINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imMGINDEX).fBoxW + 15, tmCtrls(imMGINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imMGINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imMGINDEX).fBoxW - 15, tmCtrls(imMGINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imMGINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "MG"
            pbcPostRep.Line (tmCtrls(imBONUSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imBONUSINDEX).fBoxW + 15, tmCtrls(imBONUSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imBONUSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imBONUSINDEX).fBoxW - 15, tmCtrls(imBONUSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imBONUSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Bonus"
            pbcPostRep.Line (tmCtrls(imDGROSSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imDGROSSINDEX).fBoxW + 15, tmCtrls(imDGROSSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imDGROSSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imDGROSSINDEX).fBoxW - 15, tmCtrls(imDGROSSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imDGROSSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Difference"
        Else
            pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX - 15, 15)-Step(tmCtrls(imCONTRACTINDEX).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imCONTRACTINDEX).fBoxX, 30)-Step(tmCtrls(imCONTRACTINDEX).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imCONTRACTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Contract"
            pbcPostRep.Line (tmCtrls(imCashTradeIndex).fBoxX - 15, 15)-Step(tmCtrls(imCashTradeIndex).fBoxW + 15, tmCtrls(imCONTRACTINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imCashTradeIndex).fBoxX, 30)-Step(tmCtrls(imCashTradeIndex).fBoxW - 15, tmCtrls(imCONTRACTINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imCashTradeIndex).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "C/T"
            If imAdvtIndex > 0 Then
                pbcPostRep.Line (tmCtrls(imAdvtIndex).fBoxX - 15, 15)-Step(tmCtrls(imAdvtIndex).fBoxW + 15, tmCtrls(imAdvtIndex).fBoxY - 30), BLUE, B
                pbcPostRep.Line (tmCtrls(imAdvtIndex).fBoxX, 30)-Step(tmCtrls(imAdvtIndex).fBoxW - 15, tmCtrls(imAdvtIndex).fBoxY - 60), LIGHTYELLOW, BF
                pbcPostRep.CurrentX = tmCtrls(imAdvtIndex).fBoxX + 15  'fgBoxInsetX
                pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
                pbcPostRep.Print "Advertiser"
            Else
                pbcPostRep.Line (tmCtrls(imVEHICLEINDEX).fBoxX - 15, 15)-Step(tmCtrls(imVEHICLEINDEX).fBoxW + 15, tmCtrls(imVEHICLEINDEX).fBoxY - 30), BLUE, B
                pbcPostRep.Line (tmCtrls(imVEHICLEINDEX).fBoxX, 30)-Step(tmCtrls(imVEHICLEINDEX).fBoxW - 15, tmCtrls(imVEHICLEINDEX).fBoxY - 60), LIGHTYELLOW, BF
                pbcPostRep.CurrentX = tmCtrls(imVEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
                pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
                pbcPostRep.Print "Vehicle"
            End If
            llWidth = tmCtrls(imANOSPOTSINDEX).fBoxX - tmCtrls(imPRIORNOSPOTSINDEX).fBoxX
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX - 15, 15)-Step(llWidth, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imPRIORNOSPOTSINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Ordered") 'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Ordered"
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imPRIORNOSPOTSINDEX).fBoxW + 15, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imPRIORNOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imPRIORNOSPOTSINDEX).fBoxW - 15, tmCtrls(imPRIORNOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imPRIORNOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Prior"
            pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW + 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imONOSPOTSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imONOSPOTSINDEX).fBoxW - 15, tmCtrls(imONOSPOTSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imONOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Spot"
            pbcPostRep.Line (tmCtrls(imOPRICEINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imOPRICEINDEX).fBoxW + 15, tmCtrls(imOPRICEINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imOPRICEINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imOPRICEINDEX).fBoxW - 15, tmCtrls(imOPRICEINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imOPRICEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Price"
            pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imOGROSSINDEX).fBoxW + 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imOGROSSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imOGROSSINDEX).fBoxW - 15, tmCtrls(imOGROSSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imOGROSSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Gross"
            pbcPostRep.Line (tmCtrls(imANOSPOTSINDEX).fBoxX - 15, 15)-Step(tmCtrls(imANOSPOTSINDEX).fBoxW + 15, tmCtrls(imANOSPOTSINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.CurrentX = tmCtrls(imANOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Aired"
            pbcPostRep.CurrentX = tmCtrls(imANOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Spot"

            llWidth = tmCtrls(imDGROSSINDEX).fBoxX + tmCtrls(imDGROSSINDEX).fBoxW - tmCtrls(imMISSEDINDEX).fBoxX
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX - 15, 15)-Step(llWidth + 15, tmCtrls(imMISSEDINDEX).fBoxY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX, 30)-Step(llWidth - 30, tmCtrls(imMISSEDINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imMISSEDINDEX).fBoxX + llWidth / 2 - pbcPostRep.TextWidth("Results") 'fgBoxInsetX
            pbcPostRep.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcPostRep.Print "Results"
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imMISSEDINDEX).fBoxW + 15, tmCtrls(imMISSEDINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imMISSEDINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imMISSEDINDEX).fBoxW - 15, tmCtrls(imMISSEDINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imMISSEDINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Miss"
            pbcPostRep.Line (tmCtrls(imMGINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imMGINDEX).fBoxW + 15, tmCtrls(imMGINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imMGINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imMGINDEX).fBoxW - 15, tmCtrls(imMGINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imMGINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "MG"
            pbcPostRep.Line (tmCtrls(imBONUSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imBONUSINDEX).fBoxW + 15, tmCtrls(imBONUSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imBONUSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imBONUSINDEX).fBoxW - 15, tmCtrls(imBONUSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imBONUSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Bonus"
            pbcPostRep.Line (tmCtrls(imDGROSSINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmCtrls(imDGROSSINDEX).fBoxW + 15, tmCtrls(imDGROSSINDEX).fBoxY - ilHalfY - 30), BLUE, B
            pbcPostRep.Line (tmCtrls(imDGROSSINDEX).fBoxX, ilHalfY + 30)-Step(tmCtrls(imDGROSSINDEX).fBoxW - 15, tmCtrls(imDGROSSINDEX).fBoxY - ilHalfY - 60), LIGHTYELLOW, BF
            pbcPostRep.CurrentX = tmCtrls(imDGROSSINDEX).fBoxX + 15  'fgBoxInsetX
            pbcPostRep.CurrentY = ilHalfY + 15
            pbcPostRep.Print "Difference"
        End If
    End If

    ilLineCount = 0
    llTop = tmCtrls(1).fBoxY
    Do
        For ilLoop = imLBCtrls To imMaxIndex Step 1
            If (igPostType = 5) Or (igPostType = 6) Then
                If (ilLoop <> imANOSPOTSCURRINDEX) And (ilLoop <> imBONUSCURRINDEX) And (ilLoop <> imNNOSPOTSINDEX) And (ilLoop <> imNNOBONUSINDEX) Then
                    pbcPostRep.FillStyle = 0 'Solid
                    pbcPostRep.FillColor = LIGHTYELLOW
                Else
                    pbcPostRep.FillStyle = 0 'Solid
                    pbcPostRep.FillColor = LIGHTBLUE
                End If
                pbcPostRep.Line (tmCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
                pbcPostRep.FillStyle = ilFillStyle
                pbcPostRep.FillColor = llFillColor
            Else
                If (ilLoop <> imCHECKINDEX) And (ilLoop <> imANOSPOTSINDEX) And (ilLoop <> imINVOICENOINDEX) And (ilLoop <> imANEXTNOSPOTSINDEX) Then
                    pbcPostRep.FillStyle = 0 'Solid
                    pbcPostRep.FillColor = LIGHTYELLOW
                End If
                pbcPostRep.Line (tmCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
                If (ilLoop <> imCHECKINDEX) And (ilLoop <> imANOSPOTSINDEX) And (ilLoop <> imINVOICENOINDEX) And (ilLoop <> imANEXTNOSPOTSINDEX) Then
                    pbcPostRep.FillStyle = ilFillStyle
                    pbcPostRep.FillColor = llFillColor
                End If
            End If
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmCtrls(1).fBoxH + 15
    Loop While llTop + tmCtrls(1).fBoxH < pbcPostRep.Height
    vbcPostRep.LargeChange = ilLineCount - 1
    pbcPostRep.FontSize = flFontSize
    pbcPostRep.FontName = slFontName
    pbcPostRep.FontSize = flFontSize
    pbcPostRep.ForeColor = llColor
    pbcPostRep.FontBold = True
End Sub

Private Sub mNetNamePop()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slName As String

    ReDim tmNetNames(0 To 0) As SORTCODE
    cbcNames.Clear
    ilRet = btrGetFirst(hmNrf, tmNrf, imNrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
            If tmNrf.sType = "N" Then
                tmNetNames(UBound(tmNetNames)).sKey = tmNrf.sName & "\" & Trim$(str$(tmNrf.iCode))
                ReDim Preserve tmNetNames(0 To UBound(tmNetNames) + 1) As SORTCODE
            End If
        End If
        ilRet = btrGetNext(hmNrf, tmNrf, imNrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    If UBound(tmNetNames) - 1 > 0 Then
        ArraySortTyp fnAV(tmNetNames(), 0), UBound(tmNetNames), 0, LenB(tmNetNames(0)), 0, LenB(tmNetNames(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tmNetNames) - 1 Step 1
        slNameCode = tmNetNames(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        slName = Trim$(slName)
        cbcNames.AddItem slName
    Next ilLoop
End Sub


Function mGetGenARDate(llCntrNo As Long, ilVefCode As Integer) As String
    Dim ilRet As Integer
    Dim llTranDate As Long
    Dim slDate As String
    
    mGetGenARDate = ""
    tmRvfSrchKey4.lCntrNo = llCntrNo
    gPackDateLong lmEndStd, tmRvfSrchKey4.iTranDate(0), tmRvfSrchKey4.iTranDate(1)
    ilRet = btrGetEqual(hmRvf, tmRvf, imRvfRecLen, tmRvfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (tmRvf.lCntrNo = llCntrNo)
        gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llTranDate
        If llTranDate <> lmEndStd Then
            Exit Do
        End If
        If tmRvf.iBillVefCode = ilVefCode Then
            gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slDate
            mGetGenARDate = slDate
            Exit Do
        End If
        ilRet = btrGetNext(hmRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Function

Private Sub mUpdateRvfRefInvNo(tlSbf As SBF)
    Dim ilRet As Integer
    Dim llTranDate As Long
    Dim slDate As String
    
    If (igPostType = 1) And (tgSpf.sPostCalAff = "W") Then
        tmChfSrchKey.lCode = tmSbf.lChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tmRvfSrchKey4.lCntrNo = tmChf.lCntrNo
            gPackDateLong lmEndStd, tmRvfSrchKey4.iTranDate(0), tmRvfSrchKey4.iTranDate(1)
            ilRet = btrGetEqual(hmRvf, tmRvf, imRvfRecLen, tmRvfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While (ilRet = BTRV_ERR_NONE) And (tmRvf.lCntrNo = tmChf.lCntrNo)
                gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llTranDate
                If llTranDate <> lmEndStd Then
                    Exit Do
                End If
                If (tmRvf.iBillVefCode = tmSbf.iBillVefCode) And (tmRvf.sTranType = "IN") Then
                    tmRvf.lRefInvNo = tmSbf.lRefInvNo
                    ilRet = btrUpdate(hmRvf, tmRvf, imRvfRecLen)
                End If
                ilRet = btrGetNext(hmRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    End If
End Sub
