VERSION 5.00
Begin VB.Form IconTraf 
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
      Weight          =   400
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
   Begin VB.PictureBox pbcMonitor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3930
      ScaleHeight     =   210
      ScaleWidth      =   330
      TabIndex        =   2
      Top             =   2610
      Width           =   360
   End
   Begin VB.PictureBox pbcWhiteAlert 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   270
      Left            =   5445
      Picture         =   "ICONTRAF.frx":0000
      ScaleHeight     =   270
      ScaleWidth      =   945
      TabIndex        =   1
      Top             =   3165
      Width           =   945
   End
   Begin VB.PictureBox pbcRedAlert 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   270
      Left            =   5475
      Picture         =   "ICONTRAF.frx":0482
      ScaleHeight     =   270
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   3525
      Width           =   945
   End
   Begin VB.Image imcIconOverbook 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":0904
      Height          =   195
      Left            =   5730
      Top             =   1710
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imcEmail 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3960
      Picture         =   "ICONTRAF.frx":0C0E
      Top             =   1275
      Width           =   480
   End
   Begin VB.Image imcIconJoin 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":14D8
      Height          =   195
      Left            =   5115
      Top             =   825
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imcIconAdd 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":17E2
      Height          =   195
      Left            =   6510
      Top             =   840
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imcIconFill 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":1AEC
      Height          =   195
      Left            =   5715
      Top             =   735
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imcSpecPrt 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3930
      Picture         =   "ICONTRAF.frx":1DF6
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcPrinter 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3165
      Picture         =   "ICONTRAF.frx":26C0
      Top             =   3060
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcExport 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2625
      Picture         =   "ICONTRAF.frx":2F8A
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcIconDown 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":3854
      Height          =   180
      Left            =   4290
      Top             =   705
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconUp 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":3B5E
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
      Picture         =   "ICONTRAF.frx":3E68
      Top             =   1410
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcMag 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1995
      Picture         =   "ICONTRAF.frx":4732
      Top             =   765
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcCamera 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   660
      Picture         =   "ICONTRAF.frx":4A3C
      Top             =   1290
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3600
      Picture         =   "ICONTRAF.frx":4D46
      Top             =   2220
      Width           =   480
   End
   Begin VB.Image imcIconReplace 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":5050
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
      Picture         =   "ICONTRAF.frx":535A
      Top             =   675
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcTrashClosed 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1410
      Picture         =   "ICONTRAF.frx":5664
      Top             =   735
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcCalc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4785
      Picture         =   "ICONTRAF.frx":596E
      Top             =   3450
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imcCSLogo 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3495
      Picture         =   "ICONTRAF.frx":5B50
      Top             =   1470
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcFire 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1380
      Picture         =   "ICONTRAF.frx":641A
      Top             =   2535
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcFireOut 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   780
      Picture         =   "ICONTRAF.frx":6724
      Top             =   2445
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcBoxOpened 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   720
      Picture         =   "ICONTRAF.frx":6A2E
      Top             =   1815
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcBoxClosed 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1380
      Picture         =   "ICONTRAF.frx":6D38
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcHideUp 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   795
      Picture         =   "ICONTRAF.frx":7042
      Top             =   3060
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcHideDn 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1425
      Picture         =   "ICONTRAF.frx":734C
      Top             =   3075
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcArrow 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   435
      Picture         =   "ICONTRAF.frx":7656
      Top             =   3105
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1305
      Picture         =   "ICONTRAF.frx":7960
      Top             =   1290
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcArrowUp 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   975
      Picture         =   "ICONTRAF.frx":7C6A
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcArrowDown 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1665
      Picture         =   "ICONTRAF.frx":7F74
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
      DragIcon        =   "ICONTRAF.frx":827E
      Height          =   195
      Left            =   5835
      Top             =   165
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imcIconLink 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":8588
      Height          =   180
      Left            =   6420
      Top             =   165
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconMove 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":8892
      Height          =   195
      Left            =   6105
      Top             =   165
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imcIconDrag 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":8B9C
      Height          =   180
      Left            =   4800
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconSwap 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":8EA6
      Height          =   180
      Left            =   4380
      Top             =   135
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconDwnArrow 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":91B0
      Height          =   195
      Left            =   6660
      Top             =   150
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imcIconChg 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":94BA
      Height          =   180
      Left            =   5145
      Top             =   135
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconInsert 
      Appearance      =   0  'Flat
      DragIcon        =   "ICONTRAF.frx":97C4
      Height          =   180
      Left            =   4065
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "IconTraf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ICONTRAF.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: IconTraf.Frm
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
    Set IconTraf = Nothing   'Remove data segment
End Sub

