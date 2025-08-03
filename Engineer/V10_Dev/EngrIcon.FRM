VERSION 5.00
Begin VB.Form EngrIcon 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4020
   ClientLeft      =   1140
   ClientTop       =   2685
   ClientWidth     =   7365
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
   ScaleHeight     =   4020
   ScaleWidth      =   7365
   Begin VB.PictureBox pbcWhiteAlert 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5940
      Picture         =   "EngrIcon.frx":0000
      ScaleHeight     =   270
      ScaleWidth      =   945
      TabIndex        =   1
      Top             =   2655
      Width           =   945
   End
   Begin VB.PictureBox pbcRedAlert 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5475
      Picture         =   "EngrIcon.frx":0482
      ScaleHeight     =   270
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   3150
      Width           =   945
   End
   Begin VB.Image imcSpecPrt 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   4035
      Picture         =   "EngrIcon.frx":0904
      Top             =   3060
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcPrinter 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3165
      Picture         =   "EngrIcon.frx":11CE
      Top             =   3060
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcExport 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2625
      Picture         =   "EngrIcon.frx":1A98
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcIconDown 
      Appearance      =   0  'Flat
      DragIcon        =   "EngrIcon.frx":2362
      Height          =   180
      Left            =   4290
      Top             =   705
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconUp 
      Appearance      =   0  'Flat
      DragIcon        =   "EngrIcon.frx":266C
      Height          =   180
      Left            =   3390
      Top             =   765
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2490
      Picture         =   "EngrIcon.frx":2976
      Top             =   1410
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcMag 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1995
      Picture         =   "EngrIcon.frx":3240
      Top             =   765
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcCamera 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   660
      Picture         =   "EngrIcon.frx":354A
      Top             =   1290
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   4575
      Picture         =   "EngrIcon.frx":3854
      Top             =   2175
      Width           =   480
   End
   Begin VB.Image imcIconReplace 
      Appearance      =   0  'Flat
      DragIcon        =   "EngrIcon.frx":3B5E
      Height          =   180
      Left            =   3630
      Top             =   165
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcTrashOpened 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   795
      Picture         =   "EngrIcon.frx":3E68
      Top             =   675
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcTrashClosed 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1410
      Picture         =   "EngrIcon.frx":4172
      Top             =   735
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcCalc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5505
      Picture         =   "EngrIcon.frx":447C
      Top             =   2250
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imcCSLogo 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3495
      Picture         =   "EngrIcon.frx":465E
      Top             =   1470
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcFire 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1380
      Picture         =   "EngrIcon.frx":4F28
      Top             =   2535
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcFireOut 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   780
      Picture         =   "EngrIcon.frx":5232
      Top             =   2445
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcBoxOpened 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   720
      Picture         =   "EngrIcon.frx":553C
      Top             =   1815
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcBoxClosed 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1380
      Picture         =   "EngrIcon.frx":5846
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcHideUp 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   795
      Picture         =   "EngrIcon.frx":5B50
      Top             =   3060
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcHideDn 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1425
      Picture         =   "EngrIcon.frx":5E5A
      Top             =   3075
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcArrow 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   435
      Picture         =   "EngrIcon.frx":6164
      Top             =   3105
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1305
      Picture         =   "EngrIcon.frx":646E
      Top             =   1290
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcArrowUp 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   975
      Picture         =   "EngrIcon.frx":6778
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcArrowDown 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1665
      Picture         =   "EngrIcon.frx":6A82
      Top             =   3630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcIconStd 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   5460
      Top             =   165
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconTrash 
      Appearance      =   0  'Flat
      DragIcon        =   "EngrIcon.frx":6D8C
      Height          =   195
      Left            =   5835
      Top             =   165
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imcIconLink 
      Appearance      =   0  'Flat
      DragIcon        =   "EngrIcon.frx":7096
      Height          =   180
      Left            =   6420
      Top             =   165
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconMove 
      Appearance      =   0  'Flat
      DragIcon        =   "EngrIcon.frx":73A0
      Height          =   195
      Left            =   6105
      Top             =   165
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imcIconDrag 
      Appearance      =   0  'Flat
      DragIcon        =   "EngrIcon.frx":76AA
      Height          =   180
      Left            =   4800
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconSwap 
      Appearance      =   0  'Flat
      DragIcon        =   "EngrIcon.frx":79B4
      Height          =   180
      Left            =   4380
      Top             =   135
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconDwnArrow 
      Appearance      =   0  'Flat
      DragIcon        =   "EngrIcon.frx":7CBE
      Height          =   195
      Left            =   6660
      Top             =   150
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imcIconChg 
      Appearance      =   0  'Flat
      DragIcon        =   "EngrIcon.frx":7FC8
      Height          =   180
      Left            =   5145
      Top             =   135
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconInsert 
      Appearance      =   0  'Flat
      DragIcon        =   "EngrIcon.frx":82D2
      Height          =   180
      Left            =   4065
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "EngrIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EngrIcon.Frm
'
' Release: 1.0
'
' Description:
'   This file contains Icons
'
'   Control Name       Icon File
'   imcIconDwnArrow    DwnAr_Wt.Ico
'   imcIconChg         Chg_Wt.Ico
'   imcIconLink        Link_Wt.Ico
'   imcIconInsert      InsertWt.Ico
'   imcIconSwap        Swap_Wt.Ico
'   imcIconMove        Move_Wt.Ico
'   imcIconDrag        Drag_Wt.Ico
'   imcIconStd         None
'   imcIconTrash       Not Used at this time
Option Explicit
Option Compare Text
Private Sub Form_Unload(Cancel As Integer)
    Set EngrIcon = Nothing   'Remove data segment
End Sub

