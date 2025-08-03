VERSION 5.00
Begin VB.Form CopyInv 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
   ClientLeft      =   255
   ClientTop       =   1440
   ClientWidth     =   8235
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
   ScaleHeight     =   5715
   ScaleWidth      =   8235
   Begin VB.ListBox lbcCntr 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "Copyinv.frx":0000
      Left            =   2865
      List            =   "Copyinv.frx":0002
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1215
      Visible         =   0   'False
      Width           =   5085
   End
   Begin V81Traffic.CSI_RTFEdit edcHistComment 
      Height          =   705
      Left            =   6420
      TabIndex        =   39
      Top             =   900
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1244
      Text            =   $"Copyinv.frx":0004
   End
   Begin V81Traffic.CSI_RTFEdit edcComment 
      Height          =   855
      Left            =   570
      TabIndex        =   18
      Top             =   3615
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1508
      Text            =   $"Copyinv.frx":0086
      FontName        =   ""
      FontSize        =   0
   End
   Begin VB.CommandButton cmcDupl 
      Appearance      =   0  'Flat
      Caption         =   "D&uplicate"
      Height          =   285
      Left            =   3600
      TabIndex        =   29
      Top             =   5325
      Width           =   1050
   End
   Begin VB.ListBox lbcDuplInvCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   5535
      Sorted          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   -30
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   6
      Left            =   7815
      Top             =   5115
   End
   Begin VB.ListBox lbcAnn 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   765
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2385
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcTapeDisp 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   705
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.ListBox lbcCartDisp 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   690
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2175
      Visible         =   0   'False
      Width           =   2340
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
      Left            =   8010
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4245
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
      Left            =   7995
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3855
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
      Left            =   8010
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4605
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox pbcYN 
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
      Left            =   165
      ScaleHeight     =   210
      ScaleWidth      =   870
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ListBox lbcProd 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3150
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1635
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcComp 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   2025
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1350
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcComp 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   1
      Left            =   5175
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcLen 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4830
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   960
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
      Left            =   1155
      Picture         =   "Copyinv.frx":0108
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   915
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
      Left            =   135
      MaxLength       =   20
      TabIndex        =   10
      Top             =   915
      Visible         =   0   'False
      Width           =   1020
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
      Height          =   105
      Left            =   0
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   405
      Width           =   105
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   270
      TabIndex        =   26
      Top             =   5325
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   1380
      TabIndex        =   27
      Top             =   5325
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   2490
      TabIndex        =   28
      Top             =   5325
      Width           =   1050
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4710
      TabIndex        =   30
      Top             =   5325
      Width           =   1050
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
      Height          =   285
      Left            =   5820
      TabIndex        =   31
      Top             =   5325
      Width           =   1050
   End
   Begin VB.PictureBox pbcInv 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   315
      Picture         =   "Copyinv.frx":0202
      ScaleHeight     =   1845
      ScaleWidth      =   7590
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   900
      Width           =   7590
      Begin VB.PictureBox plcCover 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   330
         Left            =   555
         ScaleHeight     =   330
         ScaleWidth      =   2505
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
         Width           =   2505
      End
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   3675
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3675
   End
   Begin VB.PictureBox plcDupl 
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
      Height          =   2325
      Left            =   255
      ScaleHeight     =   2265
      ScaleWidth      =   7650
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2880
      Width           =   7710
      Begin VB.PictureBox pbcDupl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1860
         Left            =   30
         Picture         =   "Copyinv.frx":2E474
         ScaleHeight     =   1860
         ScaleWidth      =   7590
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   390
         Width           =   7590
      End
      Begin VB.ComboBox cbcDuplInv 
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
         Left            =   1545
         TabIndex        =   24
         Top             =   60
         Width           =   6090
      End
      Begin VB.Label lacDupl 
         Appearance      =   0  'Flat
         Caption         =   "Previous Use"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   90
         TabIndex        =   23
         Top             =   75
         Width           =   1335
      End
   End
   Begin VB.PictureBox plcInv 
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
      Height          =   1980
      Left            =   255
      ScaleHeight     =   1920
      ScaleWidth      =   7650
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   840
      Width           =   7710
   End
   Begin VB.PictureBox plcCopyInv 
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   255
      ScaleHeight     =   465
      ScaleWidth      =   7665
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   7725
      Begin VB.OptionButton rbcPurged 
         Caption         =   "Never Activated"
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
         Left            =   1275
         TabIndex        =   41
         Top             =   255
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.PictureBox pbcSort 
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
         Left            =   4095
         ScaleHeight     =   210
         ScaleWidth      =   705
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.OptionButton rbcPurged 
         Caption         =   "All Inventory"
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
         Left            =   2655
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.OptionButton rbcPurged 
         Caption         =   "Purged Only"
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
         Left            =   1275
         TabIndex        =   3
         Top             =   30
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.ComboBox cbcInv 
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
         Left            =   4905
         TabIndex        =   6
         Top             =   30
         Width           =   2760
      End
      Begin VB.ComboBox cbcMedia 
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
         Top             =   30
         Width           =   1215
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
      Height          =   105
      Left            =   0
      ScaleHeight     =   105
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   660
      Width           =   60
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
      Height          =   135
      Left            =   15
      ScaleHeight     =   135
      ScaleWidth      =   60
      TabIndex        =   21
      Top             =   2655
      Width           =   60
   End
   Begin VB.CommandButton cmcImport 
      Appearance      =   0  'Flat
      Caption         =   "&Import"
      Height          =   285
      Left            =   6930
      TabIndex        =   32
      Top             =   5325
      Width           =   1050
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   5145
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "CopyInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Copyinv.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CopyInv.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Copy Inventory input screen code
Option Explicit
Option Compare Text
'12/10/14: Handle case where initially using copy Inventory, then switch to ISCI
Private smUseCartNo As String
Private smSvUseCartNo As String

Dim tmMediaCode() As SORTCODE
Dim smMediaCodeTag As String
Dim tmAnnCode() As SORTCODE
Dim smAnnCodeTag As String
Dim tmCopyCntrCode() As SORTCODE
Dim smCopyCntrCodeTag As String
Dim tmCtrls(0 To 23)  As FIELDAREA
Dim imLBCtrls As Integer
Dim tmDuplCtrls(0 To 23)  As FIELDAREA
Dim smSave(0 To 8) As String    'Index 1=Inventory #; 2=Cut #; 3=Reel #; 4=Length
                                '5=Product; 6=ISCI; 7=Creative title; 8=Action Date
Dim imSave(0 To 9) As Integer   'Index 1=Comp #1; 2=Comp # 2; 3=Purged(0=Active;1=Purged;2=History); 4=Cart Disposition (0=N/A; 1=Save; 2=Purge;3=Ask);
                                '5=Tape disposition (0=N/A;1=Return;2=Destroy;3=Ask); 6=Number times aired;
                                '7=Tape in House(0-No;1=Yes); 8=Tape approved(0=No;1=Yes); 9=Announcer
Dim smDateSave(0 To 5) As String    '1=Entered date; 2=Last used date; 3=Earliest rotation date
                                    '4=Latest rotation date; 5= Purged date
Dim lmCntrNo As Long
Dim smComment As String
Dim smCommentTextOnly As String
Dim smOrigComment As String
Dim smOrigComp0 As String
Dim smOrigComp1 As String
Dim smOrigAnn As String
Dim smOrigPurge As String
Dim smOrigLen As String
Dim smOrigRotDate As String
Dim smOrigSentDate As String
Dim smScreenCaption As String
Dim imSortCart As Integer   '0=Last Used Date; 1=Cart #
Dim imMcfCodeForSort As Integer
'Copy inventory file
Dim hmCif As Integer 'Copy inventory file handle
Dim tmCif As CIF        'CIF record image
Dim tmDuplCif As CIF
Dim tmCifSrchKey As LONGKEY0    'CIF key record image
Dim tmCifSrchKey1 As CIFKEY1    'CIF key record image
Dim imCifRecLen As Integer        'CIF record length
Dim imCifIndex As Integer
Dim imDuplIndex As Integer
Dim imBypassPurge As Integer    'If rotation is in future- disallow purge being set
'Copy script or comment file
Dim hmCsf As Integer 'Copy script or comment file handle
Dim tmCsf As CSF        'CSF record image
Dim tmDuplCsf As CSF
Dim tmCsfSrchKey As LONGKEY0    'CSF key record image
Dim imCsfRecLen As Integer        'CSF record length
'Copy Usage
Dim tmCuf As CUF            'CUF record image
Dim tmCufSrchKey As CUFKEY0  'CUF key record image
Dim tmCufSrchKey1 As CUFKEY1  'CUF key record image
Dim hmCuf As Integer        'CUF Handle
Dim imCufRecLen As Integer      'CUF record length

'Advertiser file
Dim hmAdf As Integer 'Advertiser file handle
Dim tmAdf As ADF        'ADF record image
Dim tmAdfSrchKey As INTKEY0    'ADF key record image
Dim imAdfRecLen As Integer        'ADF record length
'Copy product/ISCI file
Dim hmCpf As Integer 'Copy product/ISCI file handle
Dim tmCpf As CPF        'CPF record image
Dim tmDuplCpf As CPF
Dim tmCpfSrchKey As LONGKEY0    'CPF key record image
Dim tmCpfSrchKey1 As CPFKEY1    'CPF key record image
Dim imCpfRecLen As Integer        'CPF record length
'Media code file
Dim hmMcf As Integer 'Media file handle
Dim tmMcf As MCF        'MCF record image
Dim tmMcfSrchKey As INTKEY0    'MCF key record image
Dim imMcfRecLen As Integer        'MCF record length
Dim imMcfIndex As Integer
'Media code file
Dim hmMef As Integer 'Media file handle
Dim tmMef As MEF        'MCF record image
Dim tmMefSrchKey1 As MEFKEY1    'MCF key record image
Dim imMefRecLen As Integer        'MCF record length
'Prduct code file
Dim hmPrf As Integer 'Product file handle
Dim tmPrf As PRF        'PRF record image
Dim imPrfRecLen As Integer
'Short Title file
Dim hmSif As Integer 'Product file handle
Dim tmSif As SIF        'PRF record image
Dim imSifRecLen As Integer
'Contract
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey As LONGKEY0  'CHF key record image
Dim tmChfSrchKey1 As CHFKEY1  'CHF key record image
Dim hmCHF As Integer        'CHF Handle
Dim imCHFRecLen As Integer      'CHF record length
'Comments
Dim hmCxf As Integer            'Comments file handle
Dim tmCxf As CXF               'CXF record image
Dim tmCxfSrchKey As LONGKEY0     'CXF key record image
Dim imCxfRecLen As Integer         'CXF record length

'Record Locks
Dim lmLock1RecCode As Long
Dim hmRlf As Integer

'Dim tmRec As LPOPREC
Dim imInvStatus As Integer  '0=Active; 4=Purged; 5=History
Dim imProcMode As Integer   '0=New; <>0=Change
Dim imCbcDropDown As Integer
Dim imMaxNoCtrls As Integer 'Set to Contract Address(19) or Tax (28)
Dim imBoxNo As Integer   'Current Agency Box
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imChgModeMedia As Integer    'Change mode status (so change not entered when in change)
Dim imChgModeInv As Integer    'Change mode status (so change not entered when in change)
Dim imChgModeDuplInv As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imComboBoxIndexMedia As Integer
Dim imComboBoxIndexInv As Integer
Dim imComboBoxIndexDuplInv As Integer
Dim imComboBoxIndex As Integer
Dim imFirstActivate As Integer
Dim imFirstFocusMedia As Integer
Dim imFirstFocusInv As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer 'True = Processing arrow key- retain current list box visibly
                                'False= Make list box invisible
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imIgnorePurgeSetting As Integer
Dim imLen() As Integer
Dim imDLen As Integer   'Default length
Dim imUpdateAllowed As Integer    'User can update records

Dim tmInvNameCode() As SORTCODE
Dim smInvNameCodeTag As String

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Const INVNOINDEX = 1    'Inventory number control/field
Const CUTINDEX = 2      'Cut control/field
Const REELINDEX = 3     'Reel control/field
Const LENINDEX = 4      'Len control/field
Const ANNINDEX = 5      'Annoucer control/field
Const COMPINDEX = 6     'Competitive control/field
Const PRODINDEX = 8     'Product control/field
Const ISCIINDEX = 9     'ISCI control/field
Const PURGEINDEX = 10   'Purged control/field
Const CARTINDEX = 11    'Cart disposition control/field
Const TAPEINDEX = 12    'Tape disposition control/field
Const NOAIRINDEX = 13   'Number of times aired control/field
Const TITLEINDEX = 14   'Creative title control/field
Const TAPEININDEX = 15  'Tape in House control/field
Const TAPEAPPINDEX = 16 'Tape approved control/field
Const INVSENTDATEINDEX = 17 'Inventory vCreative Action Date
Const SCRIPTINDEX = 18  'Scripted or comment control/field
Const ENTEREDINDEX = 19 'Date entered
Const LASTUSEDINDEX = 20 'Last date used
Const EARLROTINDEX = 21 'Earliest rotation date
Const LATROTINDEX = 22  'Latest rotation date
Const PDATEINDEX = 23   'Purged date
Private Sub cbcDuplInv_Change()
    If imChgModeDuplInv = False Then
        imChgModeDuplInv = True
        If cbcDuplInv.Text <> "" Then
            gManLookAhead cbcDuplInv, imBSMode, imComboBoxIndexDuplInv
        End If
        imDuplIndex = cbcDuplInv.ListIndex
        pbcDupl.Cls
        'mPaintTitle 1
        'mPaintCopyInvTitle pbcDupl
        mReadDuplCif
        mMoveDupl
        pbcDupl_Paint
        Screen.MousePointer = vbDefault
        imChgModeDuplInv = False
        imBypassSetting = False
    End If
End Sub
Private Sub cbcDuplInv_Click()
    cbcDuplInv_Change
End Sub
Private Sub cbcDuplInv_GotFocus()
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    imComboBoxIndexDuplInv = imDuplIndex
    gCtrlGotFocus cbcDuplInv
    Exit Sub
End Sub
Private Sub cbcDuplInv_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcDuplInv_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcDuplInv.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcInv_Change()
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    If imChgModeInv = False Then
        imChgModeInv = True
        If ((Trim$(tmMcf.sReuse) = "N") And (imProcMode = 0)) Or (smUseCartNo = "N") Then
            imBypassSetting = True
            Screen.MousePointer = vbHourglass  'Wait
            ilRet = gOptionLookAhead(cbcInv, imBSMode, slStr)
            If ilRet = 0 Then
                imCifIndex = cbcInv.ListIndex
                mReadCif SETFORREADONLY
            Else
                If ilRet = 1 Then
                    cbcInv.ListIndex = 0
                End If
                ilRet = 1   'Clear fields as no match name found
                imCifIndex = 0
            End If
            pbcInv.Cls
            If ilRet = 0 Then
                mMoveRecToCtrl
            Else
                mClearCtrlFields
                If slStr <> "[New]" Then
                    If smUseCartNo <> "N" Then
                        smSave(1) = slStr
                    Else
                        smSave(6) = slStr
                    End If
                End If
            End If
        Else
            Screen.MousePointer = vbHourglass  'Wait
            If cbcInv.Text <> "" Then
                gManLookAhead cbcInv, imBSMode, imComboBoxIndexInv
            End If
            imCifIndex = cbcInv.ListIndex
            pbcInv.Cls
            mClearCtrlFields
            If Not mBlockInventory() Then
                Screen.MousePointer = vbDefault
                imChgModeInv = False
                imBypassSetting = False
                mSetCommands
                cbcInv.ListIndex = -1
                Exit Sub
            End If
            mReadCif SETFORREADONLY
            If imProcMode <> 0 Then 'lgCopyInvCifCode > 0 Then    'Change mode
                mDuplInvPop False
                mMoveRecToCtrl
            Else    'New mode
                mDuplInvPop True
                mClearPartOfCopy
                mMoveRecToCtrl
            End If
        End If
        If imProcMode <> 0 Then
            edcDropDown.Text = Trim$(tmCif.sName)
        End If
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            mInitSetShow ilLoop  'Set show strings
        Next ilLoop
        pbcInv_Paint
        Screen.MousePointer = vbDefault
        imChgModeInv = False
        imBypassSetting = False
        mSetCommands
    End If
End Sub
Private Sub cbcInv_Click()
    cbcInv_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcInv_DropDown()
    tmcClick.Interval = 300 'Delay processing encase double click
    tmcClick.Enabled = True
    imCbcDropDown = True
End Sub
Private Sub cbcInv_GotFocus()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocusInv Then
        imFirstFocusInv = False
        If igCopyInvCallSource <> CALLNONE Then  'If from advertiser or contract- set name and branch to control
            If imProcMode <> 0 Then 'lgCopyInvCifCode > 0 Then
                ilIndex = -1
                For ilLoop = 0 To UBound(tmInvNameCode) - 1 Step 1 'lbcInvCode.ListCount - 1 Step 1
                    slNameCode = tmInvNameCode(ilLoop).sKey   'lbcInvCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If ilRet = CP_MSG_NONE Then
                        If tmCif.lCode = Val(slCode) Then
                            'If smUseCartNo <> "N" Then
                            '    If Trim$(tmMcf.sReuse) = "N" Then
                            '        ilIndex = ilLoop + 1
                            '    Else
                            '        ilIndex = ilLoop
                            '    End If
                            'Else
                            '    ilIndex = ilLoop + 1
                            'End If
                            If cbcInv.ListCount > 0 Then
                                If cbcInv.List(0) = "[New]" Then
                                    ilIndex = ilLoop + 1
                                Else
                                    ilIndex = ilLoop
                                End If
                            Else
                                ilIndex = ilLoop
                            End If
                            cbcInv.ListIndex = ilIndex
                            DoEvents
                            cmcDone.SetFocus
                            Exit Sub
                        End If
                    End If
                Next ilLoop
                If ilIndex < 0 Then
                    ilIndex = ilIndex
                End If
            Else
                ilIndex = 0
            End If
            If cbcInv.ListCount > 0 Then
                cbcInv.ListIndex = ilIndex
            End If
        End If
    End If
    imComboBoxIndexInv = imCifIndex
    gCtrlGotFocus cbcInv
    Exit Sub
End Sub
Private Sub cbcInv_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcInv_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcInv.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcMedia_Change()
    If imChgModeMedia = False Then
        imChgModeMedia = True
        If cbcMedia.Text <> "" Then
            gManLookAhead cbcMedia, imBSMode, imComboBoxIndexMedia
        End If
        imMcfIndex = cbcMedia.ListIndex
        pbcInv.Cls
        Screen.MousePointer = vbHourglass
        mClearCtrlFields
        mReadMcf
        mInvPop
        mSetCommands
        Screen.MousePointer = vbDefault
        imChgModeMedia = False
    End If
End Sub
Private Sub cbcMedia_Click()
    cbcMedia_Change
End Sub
Private Sub cbcMedia_GotFocus()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocusMedia Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocusMedia = False
        If smUseCartNo = "N" Then
            If cbcInv.Enabled Then
                cbcInv.SetFocus
            Else
                cmcCancel.SetFocus
            End If
            Exit Sub
        End If
        If igCopyInvCallSource <> CALLNONE Then  'If from advertiser or contract- set name and branch to control
            If imProcMode <> 0 Then 'lgCopyInvCifCode > 0 Then
                ilIndex = -1
                For ilLoop = 0 To UBound(tmMediaCode) - 1 Step 1  'lbcMediaCode.ListCount - 1 Step 1
                    slNameCode = tmMediaCode(ilLoop).sKey    'lbcMediaCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If ilRet = CP_MSG_NONE Then
                        If tmCif.iMcfCode = Val(slCode) Then
                            ilIndex = ilLoop
                            cbcMedia.ListIndex = ilIndex
                            DoEvents
                            cbcInv.SetFocus
                            cbcMedia.Enabled = False
                            Exit Sub
                        End If
                    End If
                Next ilLoop
                If ilIndex < 0 Then
                    ilIndex = ilIndex
                End If
            Else
                ilIndex = 0
            End If
            cbcMedia.ListIndex = ilIndex
            If imProcMode <> 0 Then 'Change mode
                DoEvents
                If cbcInv.Enabled Then
                    cbcInv.SetFocus
                    Exit Sub
                End If
            Else
                If cbcMedia.ListCount = 1 Then
                    DoEvents
                    If cbcInv.Enabled Then
                        cbcInv.SetFocus
                    Else
                        pbcClickFocus.SetFocus
                    End If
                    Exit Sub
                End If
            End If
        End If
    End If
    imComboBoxIndexMedia = imMcfIndex
    gCtrlGotFocus cbcMedia
    Exit Sub
End Sub
Private Sub cbcMedia_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcMedia_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcMedia.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    mSetShow imBoxNo
    If imProcMode = 0 Then  'lgCopyInvCifCode = 0 Then    'New mode
        If mSaveRecChg(True) = False Then
            If Not imTerminate Then
                mEnableBox imBoxNo
                Exit Sub
            Else
                cmcCancel_Click
                Exit Sub
            End If
        End If
    Else
        If mSaveRecChg(False) = False Then
            If Not imTerminate Then
                mEnableBox imBoxNo
                Exit Sub
            Else
                cmcCancel_Click
                Exit Sub
            End If
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case REELINDEX   'Reel #
            lbcCntr.Visible = Not lbcCntr.Visible
        Case LENINDEX
            lbcLen.Visible = Not lbcLen.Visible
        Case ANNINDEX
            lbcAnn.Visible = Not lbcAnn.Visible
        Case COMPINDEX
            lbcComp(0).Visible = Not lbcComp(0).Visible
        Case COMPINDEX + 1
            lbcComp(1).Visible = Not lbcComp(1).Visible
        Case PRODINDEX
            lbcProd.Visible = Not lbcProd.Visible
        Case CARTINDEX
            lbcCartDisp.Visible = Not lbcCartDisp.Visible
        Case TAPEINDEX
            lbcTapeDisp.Visible = Not lbcTapeDisp.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDupl_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer

    If smUseCartNo = "N" Then
        Exit Sub
    End If
    sgISCI = smSave(6)
    sgCreativeTitle = smSave(7)
    igSortCart = imSortCart
    imSvSelectedIndex = imCifIndex
    slName = cbcInv.List(imCifIndex)
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    tgCif = tmCif
    CopyDupl.Show vbModal
    imBoxNo = -1
    'Must reset display so altered flag is cleared and setcommand will turn select on
    If Trim$(tmMcf.sReuse) = "N" Then
        If (imCifIndex < 1) And (imProcMode = 0) Then
            cbcInv.ListIndex = 0
        Else
            cbcInv.Text = slName
        End If
    Else
        If imCifIndex < 0 Then
            If cbcInv.ListCount > 0 Then
                cbcInv.ListIndex = 0
            Else
                mSetCommands
                cmcDone.SetFocus
                Exit Sub
            End If
        Else
            cbcInv.Text = slName
        End If
    End If
    cbcInv_Change    'Call change so picture area repainted
    mSetCommands
    cbcInv.SetFocus
End Sub
Private Sub cmcErase_GotFocus()
    'Code commented out in mSetCommands since erase not coded
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcImport_Click()
    If smUseCartNo <> "N" Then
        igCopyInvMcfCode = tmMcf.iCode
    Else
        igCopyInvMcfCode = 0
    End If
    ImptCopy.Show vbModal
    If smUseCartNo <> "N" Then
        pbcInv.Cls
        Screen.MousePointer = vbHourglass
        smInvNameCodeTag = ""
        mInvPop
        mSetCommands
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmcUndo_Click()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilNewInv As Integer
    Dim ilRet As Integer
    
    ilRet = mFreeBlock()
    
    ilNewInv = False
    If (Trim$(tmMcf.sReuse) = "N") Or (tgSpf.sUseCartNo = "N") Then
        If (imCifIndex < 1) And (imProcMode = 0) Then
            ilNewInv = True
        End If
    Else
        If imProcMode = 0 Then  'lgCopyInvCifCode = 0 Then    'New mode
            ilNewInv = True
        End If
    End If
    If Not ilNewInv Then
        mClearCtrlFields
        mReadCif SETFORREADONLY
        If Trim$(tmMcf.sReuse) <> "N" Then
            mDuplInvPop False
        End If
        mMoveRecToCtrl
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            mInitSetShow ilLoop  'Set show strings
        Next ilLoop
        If imProcMode <> 0 Then
            edcDropDown.Text = Trim$(tmCif.sName)
        End If
        pbcInv.Cls
        pbcInv_Paint
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcInv.Cls
    'mPaintTitle 0
    mPaintCopyInvTitle pbcInv
    cbcInv.ListIndex = 0
    cbcInv_Change    'Call change so picture area repainted
    mSetCommands
    cbcInv.SetFocus
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub cmcUndo_GotFocus()
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUpdate_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    Dim ilRet As Integer
    
    imSvSelectedIndex = imCifIndex
    If smUseCartNo <> "N" Then
        slName = cbcInv.List(imCifIndex)
    Else
        slName = smSave(6)
    End If
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    ilRet = mFreeBlock()
    If smUseCartNo <> "N" Then
        'mMediaPop
    Else
        mInvPop   'Repop incase isci changed
    End If
    imBoxNo = -1
    'Must reset display so altered flag is cleared and setcommand will turn select on
    If Trim$(tmMcf.sReuse) = "N" Then
        If (imCifIndex < 1) And (imProcMode = 0) Then
            cbcInv.ListIndex = 0
        Else
            cbcInv.Text = slName
        End If
    Else
        If imCifIndex < 0 Then
            If cbcInv.ListCount > 0 Then
                cbcInv.ListIndex = 0
            Else
                mSetCommands
                cmcDone.SetFocus
                Exit Sub
            End If
        Else
            cbcInv.Text = slName
        End If
    End If
    cbcInv_Change    'Call change so picture area repainted
    mSetCommands
    cbcInv.SetFocus
End Sub
Private Sub cmcUpdate_GotFocus()
    mSetShow imBoxNo    'Remove focus
    imBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcComment_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcComment_LostFocus()
    smComment = edcComment.Text ' This text includes the RTF format codes as well.
    smCommentTextOnly = edcComment.TextOnly ' For display purposes.
End Sub

Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNo
        Case INVNOINDEX
            smSave(1) = Trim$(edcDropDown.Text)
        Case CUTINDEX
            smSave(2) = Trim$(edcDropDown.Text)
        Case REELINDEX
            smSave(3) = Trim$(edcDropDown.Text)
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcCntr, imBSMode, slStr)
        Case LENINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcLen, imBSMode, imComboBoxIndex
        Case ANNINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcAnn, imBSMode, slStr)
            If ilRet = 1 Then
                lbcAnn.ListIndex = 1
            End If
        Case COMPINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcComp(0), imBSMode, slStr)
            If ilRet = 1 Then
                lbcComp(0).ListIndex = 1
            End If
        Case COMPINDEX + 1
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcComp(1), imBSMode, slStr)
            If ilRet = 1 Then
                lbcComp(1).ListIndex = 1
            End If
        Case PRODINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcProd, imBSMode, slStr)
            If ilRet = 1 Then   'input was ""
                lbcProd.ListIndex = 0
            End If
            If edcDropDown.Text <> "[None]" Then
                smSave(5) = edcDropDown.Text
            Else
                smSave(5) = ""
            End If
        Case ISCIINDEX
            smSave(6) = Trim$(edcDropDown.Text)
        Case NOAIRINDEX
            imSave(6) = Val(edcDropDown.Text)
        Case TITLEINDEX
            smSave(7) = Trim$(edcDropDown.Text)
        Case CARTINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcCartDisp, imBSMode, imComboBoxIndex
        Case TAPEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcTapeDisp, imBSMode, imComboBoxIndex
        Case INVSENTDATEINDEX
            smSave(8) = Trim$(edcDropDown.Text)
    End Select
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case COMPINDEX
            If lbcComp(0).ListCount = 1 Then
                lbcComp(0).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case COMPINDEX + 1
            If lbcComp(1).ListCount = 1 Then
                lbcComp(1).ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
    End Select
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    '2/3/16: Disallow forward slash
    
    'If Not gCheckKeyAscii(ilKey) Then
    'TTP 10829 JJB - 2023-09-15 changed the following function from gCheckKeyAsciiIncludeSlash to mCheckKeyAsciiIncludeSlash (local to form) so that
    '                           it won't apply to other process that may constrain key values outside of this form.
    If Not mCheckKeyAsciiIncludeSlash(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Function mCheckKeyAsciiIncludeSlash(ilKeyAscii As Integer) As Integer
    '92=Backslash (\);47=Forwardslash (/); 94=Caret (^); 124=Verical Bar (|); 91=Sq Bracket ([); 93=Sq Bracket (]); 59=Semi-colon(;); 35=pound (#); 37=Percent (%)
    'If (ilKeyAscii < 32) Or (ilKeyAscii > 126) Or (ilKeyAscii = 92) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Then
    'Allow LF, CR and Backspace.  Just remove test for < 32
    
    'TTP 10829 JJB - 2023-09-15 Added apostrophe to disallowed list
    '                                      "                    \                    /                    ^                    |                     [                    ]                    ;                     #                   %                    +                    '
    If (ilKeyAscii > 126) Or (ilKeyAscii = 34) Or (ilKeyAscii = 92) Or (ilKeyAscii = 47) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Or (ilKeyAscii = 35) Or (ilKeyAscii = 37) Or (ilKeyAscii = 43) Or (ilKeyAscii = 39) Then
        Beep
        mCheckKeyAsciiIncludeSlash = False
        Exit Function
    End If
    
    mCheckKeyAsciiIncludeSlash = True
End Function

Public Function mRemoveIllegalPastedChar(slOldName As String, Optional slExclude As String = "") As String
   
    Dim slTempName As String
    slTempName = slOldName
    
    If Trim$(slTempName) = "[New]" Or Trim$(slTempName) = "[None]" Or Trim$(slTempName) = "N/A" Or Trim$(slTempName) = "[N/A]" Then
        mRemoveIllegalPastedChar = slTempName
        Exit Function
    End If
    
    'TTP 10829 JJB - 2023-09-15 Added apostrophe to disallowed list
    If InStr(1, slExclude, "'") <= 0 Then slTempName = Replace(slTempName, "'", "")
    If InStr(1, slExclude, "/") <= 0 Then slTempName = Replace(slTempName, "/", "")
    If InStr(1, slExclude, "\") <= 0 Then slTempName = Replace(slTempName, "\", "")
    'If InStr(1, slExclude, "%") <= 0 Then slTempName = Replace(slTempName, "%", "-")
    'If InStr(1, slExclude, "*") <= 0 Then slTempName = Replace(slTempName, "*", "-")
    'If InStr(1, slExclude, ":") <= 0 Then slTempName = Replace(slTempName, ":", "-")
    If InStr(1, slExclude, "|") <= 0 Then slTempName = Replace(slTempName, "|", "")
    If InStr(1, slExclude, """") <= 0 Then slTempName = Replace(slTempName, """", "")
    'If InStr(1, slExclude, ".") <= 0 Then slTempName = Replace(slTempName, ".", "-")
    'If InStr(1, slExclude, "<") <= 0 Then slTempName = Replace(slTempName, "<", "-")
    'If InStr(1, slExclude, ">") <= 0 Then slTempName = Replace(slTempName, ">", "-")
    If InStr(1, slExclude, "^") <= 0 Then slTempName = Replace(slTempName, "^", "")
    If InStr(1, slExclude, "[") <= 0 Then slTempName = Replace(slTempName, "[", "")
    If InStr(1, slExclude, "]") <= 0 Then slTempName = Replace(slTempName, "]", "")
    If InStr(1, slExclude, ";") <= 0 Then slTempName = Replace(slTempName, ";", "")
    If InStr(1, slExclude, "#") <= 0 Then slTempName = Replace(slTempName, "#", "")
    If InStr(1, slExclude, "%") <= 0 Then slTempName = Replace(slTempName, "%", "")
    If InStr(1, slExclude, "+") <= 0 Then slTempName = Replace(slTempName, "+", "")
    
    mRemoveIllegalPastedChar = slTempName
    
End Function

Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imBoxNo
            Case REELINDEX
                gProcessArrowKey Shift, KeyCode, lbcCntr, imLbcArrowSetting
            Case LENINDEX
                gProcessArrowKey Shift, KeyCode, lbcLen, imLbcArrowSetting
            Case ANNINDEX
                gProcessArrowKey Shift, KeyCode, lbcAnn, imLbcArrowSetting
            Case COMPINDEX
                gProcessArrowKey Shift, KeyCode, lbcComp(0), imLbcArrowSetting
            Case COMPINDEX + 1
                gProcessArrowKey Shift, KeyCode, lbcComp(1), imLbcArrowSetting
            Case PRODINDEX
                gProcessArrowKey Shift, KeyCode, lbcProd, imLbcArrowSetting
            Case CARTINDEX
                gProcessArrowKey Shift, KeyCode, lbcCartDisp, imLbcArrowSetting
            Case TAPEINDEX
                gProcessArrowKey Shift, KeyCode, lbcTapeDisp, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub

Private Sub edcDropDown_LostFocus()
    '9760  '4 22 20 removed And (imBoxNo <> PRODINDEX)
    
     'TTP 10829 JJB - 2023-09-15 changed the following function from gRemoveIllegalPastedChar to mRemoveIllegalPastedChar (local to form) so that
    '                           it won't apply to other process that may constrain key values outside of this form.
    If (imBoxNo <> ANNINDEX) And (imBoxNo <> COMPINDEX) Then
        edcDropDown.Text = mRemoveIllegalPastedChar(edcDropDown.Text)
    End If
    Select Case imBoxNo
        Case INVNOINDEX
            smSave(1) = Trim$(edcDropDown.Text)
        Case CUTINDEX
            smSave(2) = Trim$(edcDropDown.Text)
        Case REELINDEX
            smSave(3) = Trim$(edcDropDown.Text)
        Case PRODINDEX
            If edcDropDown.Text <> "[None]" Then
                smSave(5) = edcDropDown.Text
            Else
                smSave(5) = ""
            End If
        Case ISCIINDEX
            smSave(6) = Trim$(edcDropDown.Text)
        Case NOAIRINDEX
            imSave(6) = Val(edcDropDown.Text)
        Case TITLEINDEX
            smSave(7) = Trim$(edcDropDown.Text)
        Case INVSENTDATEINDEX
            smSave(8) = Trim$(edcDropDown.Text)
    End Select
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
            Case PRODINDEX, ANNINDEX, COMPINDEX, COMPINDEX + 1
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
        End Select
        imDoubleClickName = False
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
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
    If (igWinStatus(COPYJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcInv.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcInv.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    CopyInv.Refresh
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
        If cbcInv.Enabled Then
            cbcInv.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            cbcInv.Enabled = True
        End If
    End If

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = (((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        If fmAdjFactorW < 1# Then
            fmAdjFactorW = 1#
        Else
            Me.Width = ((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        End If
        'fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        'Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
        fmAdjFactorH = 1#
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
        mSetShow imBoxNo
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            mEnableBox imBoxNo
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    '12/10/14
    tgSpf.sUseCartNo = smSvUseCartNo

    Erase tmInvNameCode
    
    If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
        Erase tmCopyCntrCode
    End If

    Erase tmMediaCode
    Erase tmAnnCode
    Erase imLen
    ilRet = btrClose(hmMef)
    btrDestroy hmMef
    ilRet = btrClose(hmRlf)
    btrDestroy hmRlf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCsf)
    btrDestroy hmCsf
    ilRet = btrClose(hmCuf)
    btrDestroy hmCuf
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmCxf)
    btrDestroy hmCxf
    
    Set CopyInv = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcAnn_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcAnn, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcAnn_DblClick()
    imCbcDropDown = False
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcAnn_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcAnn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcAnn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcAnn, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcCartDisp_Click()
    gProcessLbcClick lbcCartDisp, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcCartDisp_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcCntr_Click()
    gProcessLbcClick lbcCntr, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcCntr_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcComp_Click(Index As Integer)
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcComp(Index), edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcComp_DblClick(Index As Integer)
    imCbcDropDown = False
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcComp_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcComp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcComp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcComp(Index), edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcLen_Click()
    gProcessLbcClick lbcLen, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcLen_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcProd_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcProd, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcProd_DblClick()
    imCbcDropDown = False
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcProd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcProd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcProd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcProd, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcTapeDisp_Click()
    gProcessLbcClick lbcTapeDisp, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcTapeDisp_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mActiveToHistory                *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Change Active to History       *
'*                      Call only if New Inventory     *
'*                                                     *
'*******************************************************
Private Sub mActiveToHistory()
    Dim tlCif As CIF
    Dim ilRet As Integer
    Dim llRecPos As Long
    Dim ilCRet As Integer
    If tmCif.sPurged <> "A" Then
        Exit Sub
    End If
    If smUseCartNo = "N" Then
        Exit Sub
    End If
    tmCifSrchKey1.iMcfCode = tmCif.iMcfCode
    tmCifSrchKey1.sName = tmCif.sName
    tmCifSrchKey1.sCut = tmCif.sCut
    ilRet = btrGetGreaterOrEqual(hmCif, tlCif, imCifRecLen, tmCifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlCif.iMcfCode = tmCif.iMcfCode) And (tlCif.sName = tmCif.sName) And (tlCif.sCut = tmCif.sCut)
        If (tlCif.lCode <> tmCif.lCode) Then
            If tlCif.sPurged <> "H" Then
                ilRet = btrGetPosition(hmCif, llRecPos)
                Do
                    'tmRec = tlCif
                    'ilRet = gGetByKeyForUpdate("CIF", hmCif, tmRec)
                    'tlCif = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    ilRet = MsgBox("Task Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
                    '    Exit Sub
                    'End If
                    tlCif.sPurged = "H"
                    ilRet = btrUpdate(hmCif, tlCif, imCifRecLen)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        ilCRet = btrGetDirect(hmCif, tlCif, imCifRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = MsgBox("Task Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            End If
        End If
        ilRet = btrGetNext(hmCif, tlCif, imCifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAnnBranch                      *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Announc*
'*                      and process communication      *
'*                      back from announcer            *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mAnnBranch() As Integer
'
'   ilRet = mAnnBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcAnn, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mAnnBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(ANNOUNCERNAMESLIST)) Then
    '    imDoubleClickName = False
    '    mAnnBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "A"
    igMNmCallSource = CALLSOURCEADVERTISER
    If edcDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "CopyInv^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "CopyInv^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "CopyInv^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "CopyInv^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'CopyInv.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'CopyInv.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mAnnBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcAnn.Clear
        smAnnCodeTag = ""
        mAnnPop
        If imTerminate Then
            mAnnBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcAnn
        sgMNmName = ""
        If gLastFound(lbcAnn) > 0 Then
            imChgMode = True
            lbcAnn.ListIndex = gLastFound(lbcAnn)
            edcDropDown.Text = lbcAnn.List(lbcAnn.ListIndex)
            imChgMode = False
            mAnnBranch = False
            mSetChg imBoxNo
        Else
            imChgMode = True
            lbcAnn.ListIndex = 1
            edcDropDown.Text = lbcAnn.List(1)
            imChgMode = False
            mSetChg imBoxNo
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAnnPop                        *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate announcer list        *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mAnnPop()
'
'   mAnnPop
'   Where:
'
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slAnn As String      'announcer name, saved to determine if changed
    Dim ilAnn As Integer      'announcer name, saved to determine if changed
    'Repopulate if required- if sales source changed by another user while in this screen
    ilFilter(0) = CHARFILTER
    slFilter(0) = "A"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilAnn = lbcAnn.ListIndex
    If ilAnn > 1 Then
        slAnn = lbcAnn.List(ilAnn)
    End If
    'ilRet = gIMoveListBox(CopyInv, lbcAnn, lbcAnnCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(CopyInv, lbcAnn, tmAnnCode(), smAnnCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAnnPopErr
        gCPErrorMsg ilRet, "mAnnPop (gIMoveListBox)", CopyInv
        On Error GoTo 0
        lbcAnn.AddItem "[None]", 0
        lbcAnn.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilAnn > 1 Then
            gFindMatch slAnn, 2, lbcAnn
            If gLastFound(lbcAnn) > 1 Then
                lbcAnn.ListIndex = gLastFound(lbcAnn)
            Else
                lbcAnn.ListIndex = -1
            End If
        Else
            lbcAnn.ListIndex = ilAnn
        End If
        imChgMode = False
    End If
    Exit Sub
mAnnPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCheckStatus                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Check status                   *
'*                      New Status  Check              *
'*                      Active      Disallow if another*
'*                                  exist as Purged or *
'*                                  active             *
'*                      History     Disallow unless    *
'*                                  another exist as   *
'*                                  Active or Purged   *
'*                      Purged      Disallow if another*
'*                                  exist as Purged or *
'*                                  active             *
'*                                                     *
'*******************************************************
Private Function mCheckStatus() As Integer
    Dim tlCif As CIF
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFound As Integer
    If smUseCartNo = "N" Then
        mCheckStatus = True
        Exit Function
    End If
    If Trim$(tmMcf.sReuse) = "N" Then
        mCheckStatus = True
        Exit Function
    End If
    If imCifIndex < 0 Then
        slCode = "0"
    Else
        slNameCode = tmInvNameCode(imCifIndex).sKey   'lbcInvCode.List(imCifIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
    End If
    ilFound = False
    tmCifSrchKey1.iMcfCode = tmCif.iMcfCode
    tmCifSrchKey1.sName = tmCif.sName
    tmCifSrchKey1.sCut = tmCif.sCut
    ilRet = btrGetGreaterOrEqual(hmCif, tlCif, imCifRecLen, tmCifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlCif.iMcfCode = tmCif.iMcfCode) And (tlCif.sName = tmCif.sName) And (tlCif.sCut = tmCif.sCut)
        If (tlCif.lCode <> Val(slCode)) Then
            If imSave(3) = 1 Then   'Purged
                If (tlCif.sPurged = "P") Then
                    Beep
                    MsgBox "This inventory already has another marked as Purged, only one allowed", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    mCheckStatus = False
                    Exit Function
                End If
                If (tlCif.sPurged = "A") Then
                    Beep
                    MsgBox "This inventory already has another marked as Active, only one allowed", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    mCheckStatus = False
                    Exit Function
                End If
            ElseIf imSave(3) = 2 Then   'History
                If (tlCif.sPurged = "P") Or (tlCif.sPurged = "A") Then
                    ilFound = True
                    Exit Do
                End If
            Else    'Active (purged is Ok as it will be marked as history)
                If (tlCif.sPurged = "A") Then
                    Beep
                    MsgBox "This inventory already has another marked as Active, only one allowed", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    mCheckStatus = False
                    Exit Function
                End If
            End If
        End If
        ilRet = btrGetNext(hmCif, tlCif, imCifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (imSave(3) = 2) And (Not ilFound) Then
        Beep
        MsgBox "One piece of inventory must be marked as Purged or Active", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        mCheckStatus = False
        Exit Function
    End If
    mCheckStatus = True
    Exit Function
End Function
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
    Dim ilBoxNo As Integer

    For ilLoop = LBound(smSave) To UBound(smSave) Step 1
        smSave(ilLoop) = ""
    Next ilLoop
    For ilLoop = LBound(smDateSave) To UBound(smDateSave) Step 1
        smDateSave(ilLoop) = ""
    Next ilLoop
    For ilLoop = LBound(imSave) To UBound(imSave) Step 1
        imSave(ilLoop) = -1
    Next ilLoop
    For ilBoxNo = LBound(tmDuplCtrls) To UBound(tmDuplCtrls) Step 1
        tmDuplCtrls(ilBoxNo).sShow = ""
    Next ilBoxNo
    lmCntrNo = -1
    lbcLen.ListIndex = -1
    lbcAnn.ListIndex = -1
    lbcComp(0).ListIndex = -1
    lbcComp(1).ListIndex = -1
    lbcProd.ListIndex = -1
    smOrigComp0 = ""
    smOrigComp1 = ""
    smOrigAnn = ""
    smOrigPurge = ""
    smOrigRotDate = ""
    smComment = ""
    smCommentTextOnly = ""
    smOrigComment = ""
    smOrigSentDate = ""
    smOrigLen = ""
    tmCpf.lCode = 0
    tmCpf.sName = ""
    tmCpf.sISCI = ""
    tmCpf.sCreative = ""
    tmCsf.lCode = 0
    tmCsf.sComment = ""
    'tmCsf.iStrLen = 0
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    'mPaintTitle 0
    'mPaintTitle 1
    mPaintCopyInvTitle pbcInv
    mPaintCopyInvTitle pbcDupl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearPartOfCopy                *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear copy from Cif, Csp, Cpf  *
'*                      Used for duplcating, if len=0  *
'*                      then retain auto code          *
'*                                                     *
'*******************************************************
Private Sub mClearPartOfCopy()
    If tmCif.iLen <> 0 Then
        tmCif.lCode = 0 'Insert duplicate record
    End If
    'Retain media, Inventory name and cut
    tmCif.sReel = ""
    tmCif.iLen = 0
    tmCif.sPurged = ""
    tmCif.iMnfAnn = 0
    tmCif.iMnfComp(0) = 0
    tmCif.iMnfComp(1) = 0
    tmCif.sCartDisp = ""
    tmCif.sTapeDisp = ""
    'Retain NoTimesAir
    tmCif.sHouse = ""
    tmCif.sCleared = ""
    tmCif.iDateEntrd(0) = 0
    tmCif.iDateEntrd(1) = 0
    tmCif.iPurgeDate(0) = 0
    tmCif.iPurgeDate(1) = 0
    tmCif.iUsedDate(0) = 0
    tmCif.iUsedDate(1) = 0
    tmCif.iRotStartDate(0) = 0
    tmCif.iRotStartDate(1) = 0
    tmCif.iRotEndDate(0) = 0
    tmCif.iRotEndDate(1) = 0
    tmCif.sPrint = "N"
    tmCsf.lCode = 0
    tmCsf.sComment = ""
    'tmCsf.iStrLen = 0

    tmCpf.lCode = 0
    tmCpf.sName = ""
    tmCpf.sISCI = ""
    tmCpf.sCreative = ""
    tmCpf.iRotEndDate(0) = 0
    tmCpf.iRotEndDate(1) = 0

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCompBranch                     *
'*                                                     *
'*             Created:5/8/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      competitive and process             *
'*                      communication back from        *
'*                      competitive                    *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mCompBranch(ilIndex As Integer) As Integer
'
'   ilRet = mCompBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcDropDown, lbcComp(ilIndex), imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mCompBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(COMPETITIVESLIST)) Then
    '    imDoubleClickName = False
    '    mCompBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "C"
    igMNmCallSource = CALLSOURCEADVERTISER
    If edcDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "CopyInv^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "CopyInv^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "CopyInv^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "CopyInv^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'CopyInv.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'CopyInv.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mCompBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcComp(ilIndex).Clear
        sgCompCodeTag = ""
        sgCompMnfStamp = ""
        mCompPop
        If imTerminate Then
            mCompBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcComp(ilIndex)
        sgMNmName = ""
        If gLastFound(lbcComp(ilIndex)) > 0 Then
            imChgMode = True
            lbcComp(ilIndex).ListIndex = gLastFound(lbcComp(ilIndex))
            edcDropDown.Text = lbcComp(ilIndex).List(lbcComp(ilIndex).ListIndex)
            imChgMode = False
            mCompBranch = False
            mSetChg imBoxNo
        Else
            imChgMode = True
            lbcComp(ilIndex).ListIndex = 1
            edcDropDown.Text = lbcComp(ilIndex).List(1)
            imChgMode = False
            mSetChg imBoxNo
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCompPop                        *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate competitive list      *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mCompPop()
'
'   mCompPop
'   Where:
'
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    ReDim slComp(0 To 1) As String      'Competitive name, saved to determine if changed
    ReDim ilComp(0 To 1) As Integer      'Competitive name, saved to determine if changed
    'Repopulate if required- if sales source changed by another user while in this screen
    ilFilter(0) = CHARFILTER
    slFilter(0) = "C"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilComp(0) = lbcComp(0).ListIndex
    ilComp(1) = lbcComp(1).ListIndex
    If ilComp(0) > 1 Then
        slComp(0) = lbcComp(0).List(ilComp(0))
    End If
    If ilComp(1) > 1 Then
        slComp(1) = lbcComp(1).List(ilComp(1))
    End If
    If lbcComp(0).ListCount <> lbcComp(1).ListCount Then
        lbcComp(0).Clear
    End If
    'ilRet = gIMoveListBox(CopyInv, lbcComp(0), lbcCompCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(CopyInv, lbcComp(0), tgCompCode(), sgCompCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mCompPopErr
        gCPErrorMsg ilRet, "mCompPop (gIMoveListBox)", CopyInv
        On Error GoTo 0
        lbcComp(0).AddItem "[None]", 0
        lbcComp(0).AddItem "[New]", 0  'Force as first item on list
        lbcComp(1).Clear
        For ilLoop = lbcComp(0).ListCount - 1 To 0 Step -1
            lbcComp(1).AddItem lbcComp(0).List(ilLoop), 0
        Next ilLoop
        imChgMode = True
        If ilComp(0) > 1 Then
            gFindMatch slComp(0), 2, lbcComp(0)
            If gLastFound(lbcComp(0)) > 1 Then
                lbcComp(0).ListIndex = gLastFound(lbcComp(0))
            Else
                lbcComp(0).ListIndex = -1
            End If
        Else
            lbcComp(0).ListIndex = ilComp(0)
        End If
        If ilComp(1) > 1 Then
            gFindMatch slComp(1), 2, lbcComp(1)
            If gLastFound(lbcComp(1)) > 1 Then
                lbcComp(1).ListIndex = gLastFound(lbcComp(1))
            Else
                lbcComp(1).ListIndex = -1
            End If
        Else
            lbcComp(1).ListIndex = ilComp(1)
        End If
        imChgMode = False
    End If
    Exit Sub
mCompPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDuplInvPop                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the inventory         *
'*                      box with duplicated #'s        *
'*                                                     *
'*******************************************************
Private Sub mDuplInvPop(ilIncludeCif As Integer)
    Dim tlCif As CIF
    Dim tlAdf As ADF
    Dim ilRet As Integer
    Dim slRotStartDate As String
    Dim slRotEndDate As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim slEnteredDate As String
    Dim slInvStatus As String
    Dim slAdvtName As String

    lbcDuplInvCode.Clear
    cbcDuplInv.Clear
    pbcDupl.Cls
    'mPaintTitle 1
    mPaintCopyInvTitle pbcDupl
    If smUseCartNo = "N" Then
        Exit Sub
    End If
    tmCifSrchKey1.iMcfCode = tmCif.iMcfCode
    tmCifSrchKey1.sName = tmCif.sName
    tmCifSrchKey1.sCut = tmCif.sCut
    ilRet = btrGetGreaterOrEqual(hmCif, tlCif, imCifRecLen, tmCifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlCif.iMcfCode = tmCif.iMcfCode) And (tlCif.sName = tmCif.sName) And (tlCif.sCut = tmCif.sCut)
        If ilIncludeCif Or (tlCif.lCode <> tmCif.lCode) Then
            If tlCif.sPurged = "P" Then
                slInvStatus = " Purged"
            ElseIf tlCif.sPurged = "H" Then
                slInvStatus = " History"
            Else
                slInvStatus = ""
            End If
            'Obtain advertiser
            tmAdfSrchKey.iCode = tlCif.iAdfCode
            ilRet = btrGetEqual(hmAdf, tlAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDate tlCif.iRotStartDate(0), tlCif.iRotStartDate(1), slRotStartDate
                gUnpackDate tlCif.iRotEndDate(0), tlCif.iRotEndDate(1), slRotEndDate
                If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then
                    slAdvtName = Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID)
                Else
                    slAdvtName = Trim$(tlAdf.sName)
                End If
                If slRotStartDate <> "" Then
                    lbcDuplInvCode.AddItem slAdvtName & " " & slRotStartDate & "-" & slRotEndDate & slInvStatus & "\" & Trim$(str$(tlCif.lCode)) & "\" & Trim$(tlCif.sPurged)
                Else
                    gUnpackDate tlCif.iDateEntrd(0), tlCif.iDateEntrd(1), slEnteredDate
                    lbcDuplInvCode.AddItem slAdvtName & " " & slEnteredDate & slInvStatus & "\" & Trim$(str$(tlCif.lCode)) & "\" & Trim$(tlCif.sPurged)
                End If
            End If
        End If
        ilRet = btrGetNext(hmCif, tlCif, imCifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilIndex = -1
    For ilLoop = 0 To lbcDuplInvCode.ListCount - 1 Step 1
        slNameCode = lbcDuplInvCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        cbcDuplInv.AddItem slName
        If tmCif.lCode = Val(slCode) Then
            ilIndex = ilLoop
        End If
    Next ilLoop
    If ilIndex >= 0 Then
        cbcDuplInv.ListIndex = ilIndex
    Else
        If lbcDuplInvCode.ListCount > 0 Then
            cbcDuplInv.ListIndex = 0
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilHeight                                                                              *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim sl1Len As String
    Dim sl2Len As String
    Dim ilFound As Integer
    Dim ilPos As Integer

    If ilBoxNo < imLBCtrls Or ilBoxNo > imMaxNoCtrls Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case INVNOINDEX 'Name
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 5
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            imChgMode = True
            edcDropDown.Text = smSave(1)
            imChgMode = False
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case CUTINDEX   'Cut #
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 1
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            imChgMode = True
            edcDropDown.Text = smSave(2)
            imChgMode = False
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case REELINDEX   'Reel #
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                edcDropDown.MaxLength = 0
            Else
                edcDropDown.MaxLength = 10
            End If
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                lbcCntr.height = gListBoxHeight(lbcCntr.ListCount, 10)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcCntr.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            End If
            imChgMode = True
            edcDropDown.Text = smSave(3)
            imChgMode = False
            edcDropDown.Visible = True  'Set visibility
            If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                ilPos = InStr(1, smSave(3), "-", vbTextCompare)
                If ilPos > 0 Then
                    lmCntrNo = Val(Left(smSave(3), ilPos - 1))
                Else
                    ilPos = InStr(1, smSave(3), " ", vbTextCompare)
                    If ilPos > 0 Then
                        lmCntrNo = Val(Left(smSave(3), ilPos - 1))
                    Else
                        lmCntrNo = Val(smSave(3))
                    End If
                End If
                lbcCntr.Visible = True
                cmcDropDown.Visible = True
            End If
            edcDropDown.SetFocus
        Case LENINDEX   'Length
            lbcLen.height = gListBoxHeight(lbcLen.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 4
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcLen.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            gFindMatch smSave(4), 0, lbcLen
            If gLastFound(lbcLen) >= 0 Then
                imChgMode = True
                lbcLen.ListIndex = gLastFound(lbcLen)
                edcDropDown.Text = lbcLen.List(lbcLen.ListIndex)
                imChgMode = False
            Else
                ilFound = False
                If igCopyInvRotLen > 0 Then
                    gFindMatch Trim$(str$(igCopyInvRotLen)), 0, lbcLen
                    If gLastFound(lbcLen) >= 0 Then
                        ilFound = True
                    End If
                End If
                If ilFound Then
                    imChgMode = True
                    lbcLen.ListIndex = gLastFound(lbcLen)
                    edcDropDown.Text = lbcLen.List(lbcLen.ListIndex)
                    imChgMode = False
                Else
                    gFindMatch Trim$(str$(imDLen)), 0, lbcLen
                    If gLastFound(lbcLen) >= 0 Then
                        imChgMode = True
                        lbcLen.ListIndex = gLastFound(lbcLen)
                        edcDropDown.Text = lbcLen.List(lbcLen.ListIndex)
                        imChgMode = False
                    Else
                        sl1Len = "30"
                        sl2Len = "60"
                        If sgCopyForBBs = "Y" Then
                            sl1Len = "5"
                            sl2Len = "10"
                        End If
                        gFindMatch sl1Len, 0, lbcLen
                        If gLastFound(lbcLen) >= 0 Then
                            imChgMode = True
                            lbcLen.ListIndex = gLastFound(lbcLen)
                            edcDropDown.Text = lbcLen.List(lbcLen.ListIndex)
                            imChgMode = False
                        Else
                            gFindMatch sl2Len, 0, lbcLen
                            If gLastFound(lbcLen) >= 0 Then
                                imChgMode = True
                                lbcLen.ListIndex = gLastFound(lbcLen)
                                edcDropDown.Text = lbcLen.List(lbcLen.ListIndex)
                                imChgMode = False
                            Else
                                lbcLen.ListIndex = -1
                                edcDropDown.Text = ""
                            End If
                        End If
                    End If
                End If
            End If
            imComboBoxIndex = lbcLen.ListIndex
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ANNINDEX 'announcer
            mAnnPop
            If imTerminate Then
                Exit Sub
            End If
            lbcAnn.height = gListBoxHeight(lbcAnn.ListCount, 7)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcAnn.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            If imSave(9) >= 1 Then
                imChgMode = True
                lbcAnn.ListIndex = imSave(9)
                edcDropDown.Text = lbcAnn.List(lbcAnn.ListIndex)
                imChgMode = False
            Else
                imChgMode = True
                lbcAnn.ListIndex = 1   '[None]
                edcDropDown.Text = lbcAnn.List(lbcAnn.ListIndex)
                imChgMode = False
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case COMPINDEX 'Competitive
            mCompPop
            If imTerminate Then
                Exit Sub
            End If
            lbcComp(0).height = gListBoxHeight(lbcComp(0).ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcComp(0).Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            If imSave(1) >= 0 Then
                imChgMode = True
                lbcComp(0).ListIndex = imSave(1)
                edcDropDown.Text = lbcComp(0).List(lbcComp(0).ListIndex)
                imChgMode = False
            Else
                imChgMode = True
                lbcComp(0).ListIndex = 1   '[None]
                edcDropDown.Text = lbcComp(0).List(lbcComp(0).ListIndex)
                imChgMode = False
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case COMPINDEX + 1 'Competitive
            mCompPop
            If imTerminate Then
                Exit Sub
            End If
            lbcComp(1).height = gListBoxHeight(lbcComp(1).ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20  'tgSpf.iAProd
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcComp(1).Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            If imSave(2) >= 0 Then
                imChgMode = True
                lbcComp(1).ListIndex = imSave(2)
                edcDropDown.Text = lbcComp(1).List(lbcComp(1).ListIndex)
                imChgMode = False
            Else
                imChgMode = True
                lbcComp(1).ListIndex = 1   '[None]
                edcDropDown.Text = lbcComp(1).List(lbcComp(1).ListIndex)
                imChgMode = False
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PRODINDEX
            mProdPop
            If imTerminate Then
                Exit Sub
            End If
            lbcProd.height = gListBoxHeight(lbcProd.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 35  'tgSpf.iAProd
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcProd.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            gFindMatch smSave(5), 1, lbcProd
            If gLastFound(lbcProd) >= 1 Then
                imChgMode = True
                lbcProd.ListIndex = gLastFound(lbcProd)
                edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
                imChgMode = False
            Else
                If smSave(5) <> "" Then
                    imChgMode = True
                    lbcProd.ListIndex = -1
                    edcDropDown.Text = smSave(5)
                    imChgMode = False
                Else
                    imChgMode = True
                    lbcProd.ListIndex = 0   '[None]
                    edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
                    imChgMode = False
                End If
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ISCIINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            If ((Asc(tgSpf.sUsingFeatures9) And LIMITISCI) = LIMITISCI) Then
                edcDropDown.MaxLength = 15
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            imChgMode = True
            edcDropDown.Text = smSave(6)
            imChgMode = False
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case PURGEINDEX
            If imSave(3) < 0 Then
                imSave(3) = 0    'Active
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcInv, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case CARTINDEX
            If imSave(4) < 0 Then
                If (tgSpf.sUseCartNo <> "N") And (smUseCartNo <> "N") Then
                    Select Case tmMcf.sCartDisp
                        Case "N"
                            imSave(4) = 0
                        Case "S"
                            imSave(4) = 1
                        Case "P"
                            imSave(4) = 2
                        Case "A"
                            imSave(4) = 3
                        Case Else
                            imSave(4) = 3    'Ask
                    End Select
                Else
                    imSave(4) = 2
                End If
            End If
            lbcCartDisp.height = gListBoxHeight(lbcCartDisp.ListCount, 4)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 17
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcCartDisp.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            If imSave(4) < 0 Then
                lbcCartDisp.ListIndex = 0
            Else
                lbcCartDisp.ListIndex = imSave(4)
            End If
            imComboBoxIndex = lbcCartDisp.ListIndex
            If lbcCartDisp.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcCartDisp.List(lbcCartDisp.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TAPEINDEX
            If imSave(5) < 0 Then
                '2/2/12: Replaced mcfTapeDisp with mcfSuppressOnExport. Use default of N
                'Select Case tmMcf.sTapeDisp
                '    Case "N"
                '        imSave(5) = 0
                '    Case "R"
                '        imSave(5) = 1
                '    Case "D"
                '        imSave(5) = 2
                '    Case "A"
                '        imSave(5) = 3
                '    Case Else
                '        imSave(5) = 2    'Destroy
                'End Select
                imSave(5) = 0
            End If
            lbcTapeDisp.height = gListBoxHeight(lbcTapeDisp.ListCount, 4)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 17
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcTapeDisp.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            If imSave(5) < 0 Then
                lbcTapeDisp.ListIndex = 0
            Else
                lbcTapeDisp.ListIndex = imSave(5)
            End If
            imComboBoxIndex = lbcTapeDisp.ListIndex
            If lbcTapeDisp.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcTapeDisp.List(lbcTapeDisp.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case NOAIRINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 4
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            imChgMode = True
            If imSave(6) = -1 Then
                edcDropDown.Text = "0"
            Else
                edcDropDown.Text = Trim$(str$(imSave(6)))
            End If
            imChgMode = False
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case TITLEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 30
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            imChgMode = True
            'If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
            '    If smSave(7) = "" Then
            '        smSave(7) = Trim$(tmMef.sPrefix) & smSave(1) & Trim$(tmMef.sSuffix)
            '    End If
            'End If
            edcDropDown.Text = smSave(7)
            imChgMode = False
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case TAPEININDEX
            If imSave(7) < 0 Then
                imSave(7) = 1    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcInv, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case TAPEAPPINDEX
            If imSave(8) < 0 Then
                If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                    imSave(8) = 0    'No
                Else
                    imSave(8) = 1    'Yes
                End If
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcInv, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case INVSENTDATEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 30
            gMoveFormCtrl pbcInv, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            imChgMode = True
            edcDropDown.Text = smSave(8)
            imChgMode = False
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case SCRIPTINDEX
            edcComment.MaxLength = 5000
            edcComment.Top = plcCopyInv.Top
            edcComment.Left = plcInv.Left
            edcComment.Width = plcInv.Width
            ' edcComment.Height = cmcDone.Top - plcCopyInv.Top - 120
            edcComment.height = cmcDone.Top + cmcDone.height - plcCopyInv.Top
            imChgMode = True
            edcComment.Text = smComment
            imChgMode = False
            edcComment.Visible = True  'Set visibility
            edcComment.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
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
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilIndex As Integer
    Dim slAdvtName As String

    imTerminate = False
    imFirstActivate = True
    lmLock1RecCode = -1
    imLBCtrls = 1

    Screen.MousePointer = vbHourglass
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    '12/10/14
    smSvUseCartNo = tgSpf.sUseCartNo
    mInitBox
    CopyInv.height = cmcDone.Top + 5 * cmcDone.height / 3
    gCenterStdAlone CopyInv
    imMcfCodeForSort = -1
    '12/10/14: Handle case where initially entered copy by inventory, how using ISCI only
    hmCif = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "CIF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CIF.Btr)", CopyInv
    On Error GoTo 0
    imCifRecLen = Len(tmCif)  'Get and save CIF record length
    '12/10/14: Code moved here to handle cart numbers not in use any longer
    smUseCartNo = sgUseCartNo
    If lgCopyInvCifCode > 0 Then    'Change mode
        tmCifSrchKey.lCode = lgCopyInvCifCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If tmCif.iMcfCode > 0 Then
                smUseCartNo = "Y"
                tgSpf.sUseCartNo = "Y"
            End If
        End If
    End If
    If smUseCartNo <> "N" Then
        cbcMedia.Visible = True
        plcDupl.Visible = True
        plcCover.Visible = False
        cmcImport.Enabled = False
    Else
        cbcMedia.Visible = False
        plcDupl.Visible = False
        plcCover.Visible = True
        tmMcf.sScript = "Y"         'Allow scripts
        If lgCopyInvCifCode > 0 Then    'Change mode
            cmcImport.Enabled = False
        Else
            cmcImport.Enabled = True
        End If
    End If
    ReDim tmMediaCode(0 To 0) As SORTCODE
    'CopyInv.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imFirstFocusMedia = True
    imFirstFocusInv = True
    imSelectedIndex = -1
    imMcfIndex = -1
    imCifIndex = -1
    imDuplIndex = -1
    lmCntrNo = -1
    imBoxNo = -1 'Initialize current Box to N/A
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imChgMode = False
    imChgModeMedia = False
    imChgModeInv = False
    imChgModeDuplInv = False
    imBSMode = False
    imBypassSetting = False
    imCbcDropDown = False
    imMaxNoCtrls = SCRIPTINDEX  '17
    mClearCtrlFields
    lbcCartDisp.AddItem "N/A"
    lbcCartDisp.AddItem "Save"
    lbcCartDisp.AddItem "Purge"
    lbcCartDisp.AddItem "Ask after Expired" 'Default
    lbcTapeDisp.AddItem "N/A"
    lbcTapeDisp.AddItem "Return"
    lbcTapeDisp.AddItem "Destroy"       'Default
    lbcTapeDisp.AddItem "Ask after Expired"
    mLenPop
'    hmCif = CBtrvTable(TWOHANDLES) 'CBtrvObj()
'    ilRet = btrOpen(hmCif, "", sgDBPath & "CIF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: CIF.Btr)", CopyInv
'    On Error GoTo 0
'    imCifRecLen = Len(tmCif)  'Get and save CIF record length
    imProcMode = 0  'New mode
    If lgCopyInvCifCode > 0 Then    'Change mode
        imProcMode = 1
        tmCifSrchKey.lCode = lgCopyInvCifCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrGetEqual: ADF.Btr)", CopyInv
        On Error GoTo 0
        rbcPurged(0).Visible = False
        rbcPurged(1).Visible = False
        rbcPurged(2).Visible = False
        pbcSort.Visible = False
        cmcDupl.Caption = "D&uplicate"
        cbcMedia.Visible = False
    Else
        rbcPurged(0).Visible = True
        rbcPurged(1).Visible = True
        rbcPurged(2).Visible = True
        imIgnorePurgeSetting = True
        rbcPurged(0).Value = True
        pbcSort.Visible = True
        imIgnorePurgeSetting = False
        cmcDupl.Caption = "Save-D&upl"
    End If
    hmCsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCsf, "", sgDBPath & "CSF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CSF.Btr)", CopyInv
    On Error GoTo 0
    imCsfRecLen = Len(tmCsf)  'Get and save CSF record length
    hmCuf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCuf, "", sgDBPath & "CUF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CUF.Btr)", CopyInv
    On Error GoTo 0
    imCufRecLen = Len(tmCuf)  'Get and save CSF record length
    hmCpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "CPF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CPF.Btr)", CopyInv
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)  'Get and save CPF record length
    hmPrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "PRF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: PRF.Btr)", CopyInv
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)  'Get and save CSF record length
    hmSif = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSif, "", sgDBPath & "SIF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: SIF.Btr)", CopyInv
    On Error GoTo 0
    imSifRecLen = Len(tmSif)  'Get and save CSF record length
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "MCF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: MCF.Btr)", CopyInv
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)  'Get and save MCF record length
    hmMef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMef, "", sgDBPath & "MEF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: MEF.Btr)", CopyInv
    On Error GoTo 0
    imMefRecLen = Len(tmMef)  'Get and save MCF record length
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "ADF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: ADF.Btr)", CopyInv
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)  'Get and save ADF record length
    
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "CHF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CHF.Btr)", CopyInv
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)  'Get and save ADF record length
    
    hmCxf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCxf, "", sgDBPath & "CXF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CXF.Btr)", CopyInv
    On Error GoTo 0
    imCxfRecLen = Len(tmCxf)  'Get and save ADF record length
    
    'Record Locks
    hmRlf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRlf, "", sgDBPath & "Rlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rlf.Btr)", CopyInv
    On Error GoTo 0
    
    smScreenCaption = "Copy Inventory"
    If igCopyInvAdfCode > 0 Then
        tmAdfSrchKey.iCode = igCopyInvAdfCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrGetEqual: ADF.Btr)", CopyInv
        On Error GoTo 0
        If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
            slAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
        Else
            slAdvtName = Trim$(tmAdf.sName)
        End If
        ilPos = InStr(slAdvtName, "&")
        If ilPos > 0 Then
            smScreenCaption = smScreenCaption & "- " & Left$(Trim$(slAdvtName), ilPos - 1) & "&&" & Mid$(Trim$(slAdvtName), ilPos + 1)
        Else
            smScreenCaption = smScreenCaption & "- " & Trim$(slAdvtName)
        End If
        If imProcMode > 0 Then
            If imInvStatus = 4 Then
                smScreenCaption = smScreenCaption & " (Purged)"
            ElseIf imInvStatus = 5 Then
                smScreenCaption = smScreenCaption & " (History)"
            Else
                smScreenCaption = smScreenCaption & " (Active)"
            End If
        End If
    Else
        'Obtain event type/name
    End If
    Screen.MousePointer = vbHourglass  'Wait
    lbcComp(0).Clear 'Force list box to be populated
    lbcComp(1).Clear 'Force list box to be populated
    mCompPop
    If imTerminate Then
        Exit Sub
    End If
    lbcAnn.Clear 'Force list box to be populated
    mAnnPop
    If imTerminate Then
        Exit Sub
    End If
    lbcProd.Clear 'Force list box to be populated
    mProdPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    cbcMedia.Clear  'Force list to be populated

    ReDim tmInvNameCode(0 To 0) As SORTCODE
    If smUseCartNo <> "N" Then
        mMediaPop
        If imProcMode = 1 Then
            For ilLoop = LBound(tmMediaCode) To UBound(tmMediaCode) - 1 Step 1
                slNameCode = tmMediaCode(ilLoop).sKey    'lbcMediaCode.List(imMcfIndex)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If tmCif.iMcfCode = Val(Trim$(slCode)) Then
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    For ilIndex = 0 To cbcMedia.ListCount - 1 Step 1
                        If StrComp(slName, cbcMedia.List(ilIndex), 1) = 0 Then
                            imChgModeMedia = True
                            cbcMedia.ListIndex = ilIndex
                            imMcfIndex = cbcMedia.ListIndex
                            mReadMcf
                            mInvPop
                            imChgModeMedia = False
                            Exit For
                        End If
                    Next ilIndex
                    Exit For
                End If
            Next ilLoop
        End If
    Else
        tmMcf.iCode = 0
        tmMcf.sReuse = "N"
        mInvPop
    End If
    If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
        mCntrPop
    End If
    If Not imTerminate Then
        'cbcMedia.ListIndex = 0 'This will generate a select_change event
        'mSetCommands
    End If
    plcScreen_Paint
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
'*             Created:6/1/93        By:D. LeVine      *
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
    Dim ilGap As Integer
    Dim llMax As Long
    
    flTextHeight = pbcInv.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcInv.Move 255, 840, pbcInv.Width + fgPanelAdj, pbcInv.height + fgPanelAdj
    pbcInv.Move plcInv.Left + fgBevelX, plcInv.Top + fgBevelY
    plcDupl.Move plcInv.Left, 2880
    'Inventory
    gSetCtrl tmCtrls(INVNOINDEX), 30, 30, 1245, fgBoxStH
    'Cut
    gSetCtrl tmCtrls(CUTINDEX), 1290, tmCtrls(INVNOINDEX).fBoxY, 1245, fgBoxStH
    tmCtrls(CUTINDEX).iReq = False
    'Reel
    gSetCtrl tmCtrls(REELINDEX), 2550, tmCtrls(INVNOINDEX).fBoxY, 1245, fgBoxStH
    tmCtrls(REELINDEX).iReq = False
    'Length
    gSetCtrl tmCtrls(LENINDEX), 3815, tmCtrls(INVNOINDEX).fBoxY, 1245, fgBoxStH
    'Annoucer
    gSetCtrl tmCtrls(ANNINDEX), 5070, tmCtrls(INVNOINDEX).fBoxY, 2505, fgBoxStH
    tmCtrls(ANNINDEX).iReq = False
    'Competitive
    gSetCtrl tmCtrls(COMPINDEX), 30, tmCtrls(INVNOINDEX).fBoxY + fgStDeltaY, 1245, fgBoxStH
    tmCtrls(COMPINDEX).iReq = False
    gSetCtrl tmCtrls(COMPINDEX + 1), 1290, tmCtrls(COMPINDEX).fBoxY, 1245, fgBoxStH
    tmCtrls(COMPINDEX + 1).iReq = False
    'Product
    gSetCtrl tmCtrls(PRODINDEX), 2550, tmCtrls(COMPINDEX).fBoxY, 2505, fgBoxStH
    tmCtrls(PRODINDEX).iReq = False
    'ISCI
    gSetCtrl tmCtrls(ISCIINDEX), 5070, tmCtrls(COMPINDEX).fBoxY, 2505, fgBoxStH
    If smUseCartNo <> "N" Then
        tmCtrls(ISCIINDEX).iReq = False
    End If
    'Purged
    gSetCtrl tmCtrls(PURGEINDEX), 30, tmCtrls(COMPINDEX).fBoxY + fgStDeltaY, 1245, fgBoxStH
    'tmCtrls(PURGEINDEX).iReq = False
    'Cart disposition
    gSetCtrl tmCtrls(CARTINDEX), 1290, tmCtrls(PURGEINDEX).fBoxY, 1245, fgBoxStH
    'tmCtrls(CARTINDEX).iReq = False 'If not defined- set to No
    'Tape disposition
    gSetCtrl tmCtrls(TAPEINDEX), 2550, tmCtrls(PURGEINDEX).fBoxY, 1245, fgBoxStH
    'tmCtrls(TAPEINDEX).iReq = False
    'Number times aired
    gSetCtrl tmCtrls(NOAIRINDEX), 3815, tmCtrls(PURGEINDEX).fBoxY, 1245, fgBoxStH
    tmCtrls(NOAIRINDEX).iReq = False
    'Creative title
    gSetCtrl tmCtrls(TITLEINDEX), 5070, tmCtrls(PURGEINDEX).fBoxY, 2505, fgBoxStH
    tmCtrls(TITLEINDEX).iReq = False
    'Tape in
    gSetCtrl tmCtrls(TAPEININDEX), 30, tmCtrls(PURGEINDEX).fBoxY + fgStDeltaY, 1245, fgBoxStH
    tmCtrls(TAPEININDEX).iReq = False
    'Tape approved
    gSetCtrl tmCtrls(TAPEAPPINDEX), 1290, tmCtrls(TAPEININDEX).fBoxY, 1245, fgBoxStH
    tmCtrls(TAPEAPPINDEX).iReq = False
    'Inventory sent
    gSetCtrl tmCtrls(INVSENTDATEINDEX), 2550, tmCtrls(TAPEININDEX).fBoxY, 1245, fgBoxStH
    tmCtrls(INVSENTDATEINDEX).iReq = False
    'Script or comment
    gSetCtrl tmCtrls(SCRIPTINDEX), 3810, tmCtrls(TAPEININDEX).fBoxY, 3765, fgBoxStH
    tmCtrls(SCRIPTINDEX).iReq = False
    'Date entered
    gSetCtrl tmCtrls(ENTEREDINDEX), 30, 1500, 1245, fgBoxStH
    tmCtrls(LASTUSEDINDEX).iReq = False
    'Last date used
    gSetCtrl tmCtrls(LASTUSEDINDEX), 1290, tmCtrls(ENTEREDINDEX).fBoxY, 1245, fgBoxStH
    tmCtrls(LASTUSEDINDEX).iReq = False
    'Earliest rotation date
    gSetCtrl tmCtrls(EARLROTINDEX), 2550, tmCtrls(ENTEREDINDEX).fBoxY, 1875, fgBoxStH
    tmCtrls(EARLROTINDEX).iReq = False
    'Latest rotation date
    gSetCtrl tmCtrls(LATROTINDEX), 4440, tmCtrls(ENTEREDINDEX).fBoxY, 1740, fgBoxStH
    tmCtrls(LATROTINDEX).iReq = False
    'Purged date
    gSetCtrl tmCtrls(PDATEINDEX), 6195, tmCtrls(ENTEREDINDEX).fBoxY, 1380, fgBoxStH
    tmCtrls(PDATEINDEX).iReq = False
    
    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        If tmCtrls(ilLoop).fBoxX >= 0 Then
            tmCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxW)
            Do While (tmCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmCtrls(ilLoop).fBoxW = tmCtrls(ilLoop).fBoxW + 1
            Loop
            tmCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxX)
            Do While (tmCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmCtrls(ilLoop).fBoxX = tmCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmCtrls(ilLoop).fBoxX > 90 Then
                    If tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 < tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 > tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    
    pbcInv.Picture = LoadPicture("")
    pbcDupl.Picture = LoadPicture("")
    pbcInv.Width = llMax
    plcInv.Width = llMax + 2 * fgBevelX + 15
    pbcDupl.Width = llMax
    plcDupl.Width = llMax + 2 * fgBevelX + 15
    
    'cbcSelect.Left = plcBkgd.Left + plcBkgd.Width - cbcSelect.Width
    'lacCode.Left = plcBkgd.Left + plcBkgd.Width - lacCode.Width
    
    ilGap = cmcCancel.Left - (cmcDone.Left + cmcDone.Width)
    
    cmcDupl.Left = CopyInv.Width / 2 - cmcDupl.Width / 2
    cmcUpdate.Left = cmcDupl.Left - cmcUpdate.Width - ilGap
    cmcCancel.Left = cmcUpdate.Left - cmcCancel.Width - ilGap
    cmcDone.Left = cmcCancel.Left - cmcDone.Width - ilGap
    
    cmcErase.Left = cmcDupl.Left + cmcDupl.Width + ilGap
    cmcUndo.Left = cmcErase.Left + cmcErase.Width + ilGap
    cmcImport.Left = cmcUndo.Left + cmcUndo.Width + ilGap
    
    plcCover.Width = tmCtrls(INVNOINDEX).fBoxW + tmCtrls(CUTINDEX).fBoxW
    plcCover.Left = pbcInv.Left - plcInv.Left + tmCtrls(INVNOINDEX).fBoxX - 30
    plcCover.Top = pbcInv.Top - plcInv.Top + tmCtrls(INVNOINDEX).fBoxY - 30
    plcCover.height = tmCtrls(INVNOINDEX).fBoxH - 15
    
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmDuplCtrls(ilLoop).fBoxX = tmCtrls(ilLoop).fBoxX
        tmDuplCtrls(ilLoop).fBoxY = tmCtrls(ilLoop).fBoxY
        tmDuplCtrls(ilLoop).fBoxW = tmCtrls(ilLoop).fBoxW
        tmDuplCtrls(ilLoop).fBoxH = tmCtrls(ilLoop).fBoxH
        tmDuplCtrls(ilLoop).iAlign = tmCtrls(ilLoop).iAlign
        tmDuplCtrls(ilLoop).sShow = ""
    Next ilLoop
    If lgCopyInvCifCode > 0 Then    'Change mode
        plcCopyInv.Width = cbcInv.Width + 2 * fgBevelX + 15
        cbcInv.Left = fgBevelX
    End If
    plcCopyInv.Left = plcInv.Left + plcInv.Width - plcCopyInv.Width
    
    pbcInv.BackColor = WHITE
    
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitSetShow                    *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mInitSetShow(ilBoxNo As Integer)
'
'   mInitSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim ilPos As Integer
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case INVNOINDEX 'Name
            gSetShow pbcInv, smSave(1), tmCtrls(ilBoxNo)
        Case CUTINDEX   'Cut #
            gSetShow pbcInv, smSave(2), tmCtrls(ilBoxNo)
        Case REELINDEX   'Reel #
            gSetShow pbcInv, smSave(3), tmCtrls(ilBoxNo)
        Case LENINDEX   'Length
            gSetShow pbcInv, smSave(4), tmCtrls(ilBoxNo)
        Case ANNINDEX 'announcer
            If imSave(9) <= 1 Then
                slStr = ""
            Else
                slStr = lbcAnn.List(imSave(9))
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case COMPINDEX 'Competitive
            If imSave(1) <= 1 Then
                slStr = ""
            Else
                slStr = lbcComp(0).List(imSave(1))
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case COMPINDEX + 1 'Competitive
            If imSave(2) <= 1 Then
                slStr = ""
            Else
                slStr = lbcComp(1).List(imSave(2))
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case PRODINDEX
            gSetShow pbcInv, smSave(5), tmCtrls(ilBoxNo)
        Case ISCIINDEX
            gSetShow pbcInv, smSave(6), tmCtrls(ilBoxNo)
        Case PURGEINDEX
            If imSave(3) = 0 Then
                slStr = "Active"
            ElseIf imSave(3) = 1 Then
                slStr = "Purged"
            ElseIf imSave(3) = 2 Then
                slStr = "History"
            Else
                slStr = ""
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case CARTINDEX
            If imSave(4) < 0 Then
                slStr = ""
            Else
                slStr = lbcCartDisp.List(imSave(4))
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case TAPEINDEX
            If imSave(5) < 0 Then
                slStr = ""
            Else
                slStr = lbcTapeDisp.List(imSave(5))
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case NOAIRINDEX
            slStr = Trim$(str$(imSave(6)))
            If (imSave(6) = -1) Or ((imSave(6) = 0) And (cbcInv.ListIndex = -1)) Then
                slStr = ""
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case TITLEINDEX
            gSetShow pbcInv, smSave(7), tmCtrls(ilBoxNo)
        Case TAPEININDEX
            If imSave(7) = 0 Then
                slStr = "No"
            ElseIf imSave(7) = 1 Then
                slStr = "Yes"
            Else
                slStr = ""
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case TAPEAPPINDEX
            If ((Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT) Then
                If imSave(8) = 0 Then
                    slStr = "Not Sent"
                ElseIf imSave(8) = 1 Then
                    slStr = "Produced"
                ElseIf (imSave(8) = 2) Then
                    slStr = "Sent"
                ElseIf (imSave(8) = 3) Then
                    slStr = "Hold"
                Else
                    slStr = ""
                End If
            Else
                If imSave(8) = 0 Then
                    slStr = "No"
                ElseIf imSave(8) = 1 Then
                    slStr = "Yes"
                Else
                    slStr = ""
                End If
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case INVSENTDATEINDEX
            If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                gSetShow pbcInv, smSave(8), tmCtrls(ilBoxNo)
            End If
        Case SCRIPTINDEX
            slStr = Left$(smCommentTextOnly, 80)
            ilPos = InStr(slStr, sgLF)
            If ilPos = 2 Then
                slStr = Mid$(slStr, ilPos + 1)
            End If
            ilPos = InStr(slStr, sgCR)
            If ilPos > 0 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case ENTEREDINDEX
            gSetShow pbcInv, smDateSave(1), tmCtrls(ilBoxNo)
        Case LASTUSEDINDEX
            gSetShow pbcInv, smDateSave(2), tmCtrls(ilBoxNo)
        Case EARLROTINDEX
            gSetShow pbcInv, smDateSave(3), tmCtrls(ilBoxNo)
        Case LATROTINDEX
            gSetShow pbcInv, smDateSave(4), tmCtrls(ilBoxNo)
        Case PDATEINDEX
            gSetShow pbcInv, smDateSave(5), tmCtrls(ilBoxNo)
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInvPop                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the inventory         *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mInvPop()
'
'   mInvPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slName As String
    Dim ilIndex As Integer
    Dim slNowDate As String

    ilRet = mFreeBlock()
    
    If smUseCartNo = "N" Then
        'If (lbcInvCode.ListCount > 0) And (cbcInv.ListCount > 0) Then
        '    Exit Sub
        'End If
        'lbcInvCode.Clear
        smInvNameCodeTag = ""
        ReDim tmInvNameCode(0 To 0) As SORTCODE
        cbcInv.Clear

        If imProcMode <> 0 Then 'Change mode
            'ilRet = gPopCopyForAdvtBox(CopyInv, igCopyInvAdfCode, 4, imInvStatus, cbcInv, lbcInvCode)
            ilRet = gPopCopyForAdvtBox(CopyInv, igCopyInvAdfCode, 4, imInvStatus + &H400, cbcInv, tmInvNameCode(), smInvNameCodeTag)
        Else
            If rbcPurged(0).Value Then
                ilRet = gPopCopyForAdvtBox(CopyInv, igCopyInvAdfCode, 4, 4, cbcInv, tmInvNameCode(), smInvNameCodeTag)
            Else
                ilRet = gPopCopyForAdvtBox(CopyInv, igCopyInvAdfCode, 4, 1 + &H100, cbcInv, tmInvNameCode(), smInvNameCodeTag)
            End If
        End If
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mInvPopErr
            gCPErrorMsg ilRet, "mInvPop (gIMoveListBox)", CopyInv
            On Error GoTo 0
            If imProcMode = 0 Then
                cbcInv.AddItem "[New]", 0  'Force as first item on list
            End If
        End If
        imBypassPurge = False
        Exit Sub
    End If
    If imMcfIndex < 0 Then
        Exit Sub
    End If
    If imProcMode <> 0 Then 'lgCopyInvCifCode > 0 Then    'Change mode- only populate with specified inventory
        'If (lbcInvCode.ListCount > 0) And (cbcInv.ListCount > 0) Then
        If (UBound(tmInvNameCode) > 0) And (cbcInv.ListCount > 0) Then
            Exit Sub
        End If
        'lbcInvCode.Clear
        smInvNameCodeTag = ""
        ReDim tmInvNameCode(0 To 0) As SORTCODE
        cbcInv.Clear

        ilRet = gPopCopyForAdvtBox(CopyInv, igCopyInvAdfCode, 0, imInvStatus + &H400, cbcInv, tmInvNameCode(), smInvNameCodeTag)
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mInvPopErr
            gCPErrorMsg ilRet, "mInvPop (gIMoveListBox)", CopyInv
            On Error GoTo 0
            'If Trim$(tmMcf.sReuse) = "N" Then
            '    cbcInv.AddItem "[New]", 0  'Force as first item on list
            'End If
        End If
        'gUnpackDateForSort tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slRotDate

        'If slRotDate = "" Then
        '    slRotDate = "000000"
        'End If
        'If tmCif.sPurged = "Y" Then
        '    slName = slRotDate 'slPurgeDate
         '   slName = "A" & slName
        'Else
        '    slName = slRotDate
        '    slName = "Z" & slName
        'End If
        'If Trim$(tmCif.sCut) = "" Then
        '    slInvName = Trim$(tmCif.sName)
        'Else
        '    slInvName = Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut)
        'End If
        'If tmCif.sPurged = "Y" Then
        '    slInvName = slInvName & "/Purged"
        'End If
        'slName = slName & "|" & slInvName & "\" & Trim$(Str$(tmCif.lCode))
        'lbcInvCode.AddItem slName
        'cbcInv.AddItem slInvName
        'gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slRotDate
        'If slRotDate <> "" Then
        '    slDate = Format$(gNow(), "m/d/yy")
        '    If gDateValue(slRotDate) < gDateValue(slDate) Then
        '        imBypassPurge = False
        '    Else
        '        imBypassPurge = True
        '    End If
        'Else
            imBypassPurge = False
        'End If
        Exit Sub
    End If
    ilIndex = cbcInv.ListIndex
    If ilIndex >= 0 Then
        slName = cbcInv.List(ilIndex)
    End If
    slNowDate = Format$(gNow(), "m/d/yy")
    If rbcPurged(0).Value Then
        ilRet = gPopCopyForMediaBox(CopyInv, tmMcf.iCode, slNowDate, True, True, imSortCart, cbcInv, tmInvNameCode(), smInvNameCodeTag)    'lbcInvCode)
    ElseIf rbcPurged(2).Value Then  'Never actived
        ilRet = gPopCopyForMediaBox(CopyInv, tmMcf.iCode, "-1", True, True, imSortCart, cbcInv, tmInvNameCode(), smInvNameCodeTag)    'lbcInvCode)
    Else
        ilRet = gPopCopyForMediaBox(CopyInv, tmMcf.iCode, slNowDate, True, False, imSortCart, cbcInv, tmInvNameCode(), smInvNameCodeTag)  'lbcInvCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mInvPopErr
        gCPErrorMsg ilRet, "mInvPop (gIMoveListBox)", CopyInv
        On Error GoTo 0
        If Trim$(tmMcf.sReuse) = "N" Then
            cbcInv.AddItem "[New]", 0  'Force as first item on list
        End If
        If ilIndex >= 0 Then
            gFindMatch slName, 0, cbcInv
            If gLastFound(cbcInv) >= 0 Then
                cbcInv.ListIndex = gLastFound(cbcInv)
            Else
                cbcInv.ListIndex = -1
            End If
        Else
            cbcInv.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mInvPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInvPurgeOK                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test that Purge status has     *
'*                      not changed to "A"             *
'*                                                     *
'*******************************************************
Private Function mInvPurgeOK() As Integer
    Dim tlCif As CIF
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slDuplNameCode As String
    Dim slDuplCode As String
    Dim slDuplPurged As String
    If smUseCartNo = "N" Then
        mInvPurgeOK = True
        Exit Function
    End If
    If (imSave(3) > 0) Then 'If not "Active"- exit
        mInvPurgeOK = True
        Exit Function
    End If
    If Trim$(tmMcf.sReuse) = "N" Then
        If (imCifIndex < 1) And (imProcMode = 0) Then
            mInvPurgeOK = True
            Exit Function
        End If
        If Trim$(cbcInv.List(0)) = "[New]" Then
            slNameCode = tmInvNameCode(imCifIndex - 1).sKey   'lbcInvCode.List(imCifIndex - 1)
        Else
            slNameCode = tmInvNameCode(imCifIndex).sKey   'lbcInvCode.List(imCifIndex - 1)
        End If
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
    Else
        If imCifIndex < 0 Then
            slCode = "0"
        Else
            slNameCode = tmInvNameCode(imCifIndex).sKey   'lbcInvCode.List(imCifIndex)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
        End If
    End If
    If Trim$(tmMcf.sReuse) = "N" Then
        tmCifSrchKey1.iMcfCode = tmMcf.iCode    'tmCif.iMcfCode
        tmCifSrchKey1.sName = smSave(1) 'tmCif.sName
        tmCifSrchKey1.sCut = smSave(2)  'tmCif.sCut
    Else
        tmCifSrchKey1.iMcfCode = tmMcf.iCode    'tmCif.iMcfCode
        tmCifSrchKey1.sName = tmCif.sName
        tmCifSrchKey1.sCut = tmCif.sCut
    End If
    ilRet = btrGetGreaterOrEqual(hmCif, tlCif, imCifRecLen, tmCifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlCif.iMcfCode = tmCif.iMcfCode) And (tlCif.sName = tmCif.sName) And (tlCif.sCut = tmCif.sCut)
        'Find Previous inventory
        ilFound = False
        For ilLoop = 0 To lbcDuplInvCode.ListCount - 1 Step 1
            slDuplNameCode = lbcDuplInvCode.List(ilLoop)
            ilRet = gParseItem(slDuplNameCode, 2, "\", slDuplCode)
            ilRet = gParseItem(slDuplNameCode, 3, "\", slDuplPurged)
            If tlCif.lCode = Val(slDuplCode) Then
                ilFound = True
                If (Trim$(slDuplPurged) <> Trim$(tlCif.sPurged)) And (Trim$(tlCif.sPurged) = "A") Then
                    Beep
                    MsgBox "Inventory altered by another user, select another", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    mInvPurgeOK = False
                    Exit Function
                End If
            End If
        Next ilLoop
        If Not ilFound Then
            If (tlCif.lCode = Val(slCode)) Then
                ilFound = True
                If (Trim$(tlCif.sPurged) = "A") And (smOrigPurge <> "A") Then
                    Beep
                    MsgBox "Inventory altered by another user, select another", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    mInvPurgeOK = False
                    Exit Function
                End If
            End If
        End If

        If (Not ilFound) And (tlCif.sPurged = "A") And (tlCif.lCode <> Val(slCode)) Then
            Beep
            MsgBox "Inventory added by another user, select another", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
            mInvPurgeOK = False
            Exit Function
        End If
        ilRet = btrGetNext(hmCif, tlCif, imCifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mInvPurgeOK = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestISCI                       *
'*                                                     *
'*             Created:4/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test is ISCI is unique          *
'*                                                     *
'*******************************************************
Private Function mISCIOk(sISCI As String) As Integer
    Dim tlCpf As CPF
    Dim ilRet As Integer
    Dim hlCif As Integer        'Cif handle
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilCifRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlCif As CIF
    Dim ilOffSet As Integer
    Dim ilTest As Integer
    Dim tlLTypeBuff As POPLCODE   'Type field record
    mISCIOk = True
    If Trim$(sISCI) = "" Then
        Exit Function
    End If
    hlCif = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        On Error GoTo mISCIOkErr
        gBtrvErrorMsg ilRet, "mISCIOk (btrOpen):" & "Cif.Btr", CopyInv
        On Error GoTo 0
        Exit Function
    End If
    ilCifRecLen = Len(tlCif) 'btrRecordLength(hlAdf)  'Get and save record length
    tmCpfSrchKey1.sISCI = sISCI 'smSave(6)
    ilRet = btrGetEqual(hmCpf, tlCpf, imCpfRecLen, tmCpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (Trim$(tlCpf.sISCI) = Trim$(sISCI))
        If tmCpf.lCode <> tlCpf.lCode Then
            'Test if inventory that is referencing this CPF is Purged or
            'History- if so, then ISCI Ok
            ilExtLen = Len(tlCif)  'Extract operation record size
            llNoRec = gExtNoRec(ilExtLen)  'btrRecords(hlCif) 'Obtain number of records
            btrExtClear hlCif   'Clear any previous extend operation
            ilRet = btrGetFirst(hlCif, tlCif, ilCifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_END_OF_FILE Then
                ilRet = btrClose(hlCif)
                btrDestroy hlCif
                Exit Function
            End If
            If ilRet <> BTRV_ERR_NONE Then
                On Error GoTo mISCIOkErr
                gBtrvErrorMsg ilRet, "mISCIOk (btrGetFirst):" & "Cif.Btr", CopyInv
                On Error GoTo 0
                Exit Function
            End If
            Call btrExtSetBounds(hlCif, llNoRec, -1, "UC", "CIF", "") 'Set extract limits (all records including first)
            tlLTypeBuff.lCode = tlCpf.lCode
            ilOffSet = gFieldOffset("Cif", "CIFCPFCODE")
            ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlLTypeBuff, 4)
            ilOffSet = 0
            ilRet = btrExtAddField(hlCif, ilOffSet, Len(tlCif))  'Extract the whole record
            ilRet = btrExtGetNext(hlCif, tlCif, ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                    On Error GoTo mISCIOkErr
                    gBtrvErrorMsg ilRet, "mISCIOk (btrExtGetNext):" & "Cif.Btr", CopyInv
                    On Error GoTo 0
                    Exit Function
                End If
                ilExtLen = Len(tlCif)  'Extract operation record size
                'ilRet = btrExtGetFirst(hlCif, tlCifExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCif, tlCif, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    ilTest = True
'                    If smUseCartNo <> "N" Then
                        If tmMcf.iCode <> tlCif.iMcfCode Then
                            If igCopyInvAdfCode = tlCif.iAdfCode Then
                                ilTest = False
                            End If
                        End If
'                    End If
                    If (tlCif.sPurged = "A") And (ilTest) Then
                        ilRet = btrClose(hlCif)
                        btrDestroy hlCif
                        mISCIOk = False
                        Exit Function
                    End If
                    ilRet = btrExtGetNext(hlCif, tlCif, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlCif, tlCif, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        ilRet = btrGetNext(hmCpf, tlCpf, imCpfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = btrClose(hlCif)
    btrDestroy hlCif
    mISCIOk = True
    Exit Function
mISCIOkErr:
    ilRet = btrClose(hlCif)
    btrDestroy hlCif
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mLenPop                         *
'*                                                     *
'*             Created:7/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection length  *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mLenPop()
    Dim ilLoop As Integer
    Dim ilLen As Integer
    Dim ilIndex As Integer
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim ilMaxLen As Integer
    Dim ilMinLen As Integer

    lbcLen.Clear
    ilLen = 0
    ilMaxLen = -1
    ilMinLen = 32000
    ReDim imLen(0 To 0) As Integer
    For ilLoop = 0 To UBound(tgVpf) Step 1
        ilFound = False
        'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgMVef(ilIndex).iCode = tgVpf(ilLoop).iVefKCode Then
            ilIndex = gBinarySearchVef(tgVpf(ilLoop).iVefKCode)
            If ilIndex <> -1 Then
                'Rep added to handle clients that are rep[ only and want to keep copy inventory
                If ((tgMVef(ilIndex).sType = "C") Or (tgMVef(ilIndex).sType = "G") Or (tgMVef(ilIndex).sType = "S") Or (tgMVef(ilIndex).sType = "R")) And (tgMVef(ilIndex).sState <> "D") Then
                    ilFound = True
                End If
        '        Exit For
            End If
        'Next ilIndex
        If ilFound Then
            For ilIndex = LBound(tgVpf(ilLoop).iSLen) To UBound(tgVpf(ilLoop).iSLen) Step 1
                If tgVpf(ilLoop).iSLen(ilIndex) > 0 Then
                    ilFound = False
                    For ilTest = 0 To UBound(imLen) - 1 Step 1
                        If imLen(ilTest) = tgVpf(ilLoop).iSLen(ilIndex) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilTest
                    'If billboards, bypass lengths greater or equal to 30 sec
                    If sgCopyForBBs = "Y" Then
                        If tgVpf(ilLoop).iSLen(ilIndex) >= 30 Then
                            ilFound = True
                        End If
                    End If
                    If Not ilFound Then
                        imDLen = tgVpf(ilLoop).iSDLen
                        If tgVpf(ilLoop).iSLen(ilIndex) < ilMinLen Then
                            ilMinLen = tgVpf(ilLoop).iSLen(ilIndex)
                        End If
                        If tgVpf(ilLoop).iSLen(ilIndex) > ilMaxLen Then
                            ilMaxLen = tgVpf(ilLoop).iSLen(ilIndex)
                        End If
                        imLen(UBound(imLen)) = tgVpf(ilLoop).iSLen(ilIndex)
                        ReDim Preserve imLen(0 To UBound(imLen) + 1) As Integer
                    End If
                End If
            Next ilIndex
        End If
    Next ilLoop
    'For ilLoop = LBound(tgSpf.iSLen) To UBound(tgSpf.iSLen) Step 1
    '    If tgSpf.iSLen(ilLoop) <> 0 Then
    '        lbcLen.AddItem Trim$(Str$(tgSpf.iSLen(ilLoop)))
    '    End If
    'Next ilLoop

    'Sort by length
    For ilLen = ilMinLen To ilMaxLen Step 1
        For ilLoop = LBound(imLen) To UBound(imLen) - 1 Step 1
            If ilLen = imLen(ilLoop) Then
                lbcLen.AddItem Trim$(str$(imLen(ilLoop)))
                Exit For
            End If
        Next ilLoop
    Next ilLen
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMediaPop                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the media combo       *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mMediaPop()
'
'   mMediaPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim slName As String
    Dim ilIndex As Integer
    If smUseCartNo = "N" Then
        ReDim tmMediaCode(0 To 0) As SORTCODE
        Exit Sub
    End If
    ilIndex = cbcMedia.ListIndex
    If ilIndex >= 0 Then
        slName = cbcMedia.List(ilIndex)
    End If
    ilFilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffSet(0) = 0
    'ilRet = gIMoveListBox(CopyInv, cbcMedia, lbcMediaCode, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(CopyInv, cbcMedia, tmMediaCode(), smMediaCodeTag, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mMediaPopErr
        gCPErrorMsg ilRet, "mMediaPop (gIMoveListBox)", CopyInv
        On Error GoTo 0
'        cbcMedia.AddItem "[New]", 0  'Force as first item on list
        If ilIndex >= 0 Then
            gFindMatch slName, 0, cbcMedia
            If gLastFound(cbcMedia) >= 0 Then
                cbcMedia.ListIndex = gLastFound(cbcMedia)
            Else
                cbcMedia.ListIndex = -1
            End If
        Else
            cbcMedia.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mMediaPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilTestChg As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPos                         slTemp                        slStrToFind               *
'*  ilLenOfStrToFind                                                                      *
'******************************************************************************************

'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    Dim ilLoop As Integer
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'code number
    Dim slNameCode As String
    Dim slNowDate As String
    Dim ilPos As Integer

    slNowDate = Format$(gNow(), "m/d/yy")
    If smUseCartNo <> "N" Then
        If Trim$(tmMcf.sReuse) = "N" Then
            If (imCifIndex = 0) And (imProcMode = 0) Then  'New
                tmCif.lCode = 0
                tmCif.sName = smSave(1)
            End If
            If Not ilTestChg Or tmCtrls(INVNOINDEX).iChg Then
                tmCif.sName = smSave(1)
            End If
            If Not ilTestChg Or tmCtrls(CUTINDEX).iChg Then
                tmCif.sCut = smSave(2)
            End If
        End If
        tmCif.iMcfCode = tmMcf.iCode
    Else
        If (imCifIndex = 0) And (imProcMode = 0) Then  'New
            tmCif.lCode = 0
        End If
        tmCif.sName = ""
        tmCif.sCut = ""
        tmCif.iMcfCode = 0
    End If
    tmCif.iEtfCode = igCopyInvEtfCode
    tmCif.iEnfCode = igCopyInvEnfCode
    tmCif.iAdfCode = igCopyInvAdfCode 'From rotation input screen
    If Not ilTestChg Or tmCtrls(REELINDEX).iChg Then
        tmCif.sReel = smSave(3)
        If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
            ilPos = InStr(1, smSave(3), " ", vbTextCompare)
            If ilPos > 0 Then
                tmCif.sReel = Left(smSave(3), ilPos - 1)
            End If
        End If
    End If
    If Not ilTestChg Or tmCtrls(LENINDEX).iChg Then
        tmCif.iLen = Val(smSave(4))
    End If
    'tmCif.lCpfCode set within save
    For ilLoop = 0 To 1 Step 1
        If Not ilTestChg Or tmCtrls(COMPINDEX + ilLoop).iChg Then
            If imSave(1 + ilLoop) >= 2 Then
                slNameCode = tgCompCode(imSave(1 + ilLoop) - 2).sKey   'lbcCompCode.List(imSave(1 + ilLoop) - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveCtrlToRecErr
                gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", CopyInv
                On Error GoTo 0
                slCode = Trim$(slCode)
                tmCif.iMnfComp(ilLoop) = CInt(slCode)
            Else
                tmCif.iMnfComp(ilLoop) = 0
            End If
        End If
    Next ilLoop
    If Not imBypassPurge Then
        If Not ilTestChg Or tmCtrls(PURGEINDEX).iChg Then
            If imSave(3) = 1 Then
                tmCif.sPurged = "P"
            ElseIf imSave(3) = 2 Then
                tmCif.sPurged = "H"
            Else
                tmCif.sPurged = "A"
            End If
        End If
    Else
        If imSave(3) = -1 Then
            tmCif.sPurged = "A"
        End If
    End If
    If Not ilTestChg Or tmCtrls(CARTINDEX).iChg Then
        If imSave(4) = 1 Then
            tmCif.sCartDisp = "S"
        ElseIf imSave(4) = 2 Then
            tmCif.sCartDisp = "P"
        ElseIf imSave(4) = 3 Then
            tmCif.sCartDisp = "A"
        Else
            tmCif.sCartDisp = "N"
        End If
    End If
    If Not ilTestChg Or tmCtrls(TAPEINDEX).iChg Then
        If imSave(5) = 1 Then
            tmCif.sTapeDisp = "R"
        ElseIf imSave(5) = 2 Then
            tmCif.sTapeDisp = "D"
        ElseIf imSave(5) = 3 Then
            tmCif.sTapeDisp = "A"
        Else
            tmCif.sTapeDisp = "N"
        End If
    End If
    If Not ilTestChg Or tmCtrls(NOAIRINDEX).iChg Then
        tmCif.iNoTimesAir = imSave(6)
    End If
    If Not ilTestChg Or tmCtrls(TAPEININDEX).iChg Then
        If imSave(7) = 1 Then
            tmCif.sHouse = "Y"
        Else
            tmCif.sHouse = "N"
        End If
    End If
    If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
        If Not ilTestChg Or tmCtrls(TAPEAPPINDEX).iChg Then
            If imSave(8) = 1 Then
                If (tmCif.sCleared <> "Y") Or (tmCtrls(INVSENTDATEINDEX).iChg) Then
                    If (tmCtrls(INVSENTDATEINDEX).iChg) And (smSave(8) <> "") And (smSave(8) <> "1/1/1970") Then
                        gPackDate smSave(8), tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    Else
                        gPackDate slNowDate, tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    End If
                End If
                tmCif.sCleared = "Y"
            ElseIf imSave(8) = 2 Then
                If (tmCif.sCleared <> "S") Or (tmCtrls(INVSENTDATEINDEX).iChg) Then
                    If (tmCtrls(INVSENTDATEINDEX).iChg) And (smSave(8) <> "") And (smSave(8) <> "1/1/1970") Then
                        gPackDate smSave(8), tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    Else
                        gPackDate slNowDate, tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    End If
                End If
                tmCif.sCleared = "S"
            ElseIf imSave(8) = 3 Then
                If (tmCif.sCleared <> "H") Or (tmCtrls(INVSENTDATEINDEX).iChg) Then
                    If (tmCtrls(INVSENTDATEINDEX).iChg) And (smSave(8) <> "") And (smSave(8) <> "1/1/1970") Then
                        gPackDate smSave(8), tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    Else
                        gPackDate slNowDate, tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    End If
                End If
                tmCif.sCleared = "H"
            Else
                If (tmCif.sCleared <> "N") Or (tmCtrls(INVSENTDATEINDEX).iChg) Then
                    If (tmCtrls(INVSENTDATEINDEX).iChg) And (smSave(8) <> "") And (smSave(8) <> "1/1/1970") Then
                        gPackDate smSave(8), tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    Else
                        gPackDate slNowDate, tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    End If
                End If
                tmCif.sCleared = "N"
            End If
        End If
    Else
        If Not ilTestChg Or tmCtrls(TAPEAPPINDEX).iChg Then
            If imSave(8) = 1 Then
                tmCif.sCleared = "Y"
            Else
                tmCif.sCleared = "N"
            End If
        End If
    End If
    If Not ilTestChg Or tmCtrls(ANNINDEX).iChg Then
        If imSave(9) >= 2 Then
            slNameCode = tmAnnCode(imSave(9) - 2).sKey   'lbcAnnCode.List(imSave(9) - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", CopyInv
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmCif.iMnfAnn = CInt(slCode)
        Else
            tmCif.iMnfAnn = 0
        End If
    End If
    
    If ((Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT) Then
        If Not tmCtrls(TAPEAPPINDEX).iChg Then
            If Not ilTestChg Or tmCtrls(INVSENTDATEINDEX).iChg Then
                If (smSave(8) = "") Or (smSave(8) = "1/1/1970") Then
                    gPackDate slNowDate, tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    tmCif.sCleared = "N"
                Else
                    gPackDate smSave(8), tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    'If gDateValue(smOrigSentDate) <> gDateValue(smSave(8)) Then
                    '    tmCif.sCleared = "Y"
                    'End If
                End If
            End If
        End If
    Else
        gPackDate "1/1/1970", tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
    End If
    
    If Not ilTestChg Or tmCtrls(SCRIPTINDEX).iChg Then
'        ' Strip out all language identifiers.
'        It turns out this is not what is causing the double blank page issue.
'        I'm leaving this code here though just in case we need to strip out any other control codes in the future.
'        slStrToFind = "\deflangfe1033"
'        ilLenOfStrToFind = Len(slStrToFind)
'        ilPos = InStr(1, smComment, slStrToFind)
'        If ilPos > 0 Then
'            slTemp = Left(smComment, ilPos - 1)
'            slTemp = slTemp + Mid(smComment, ilPos + ilLenOfStrToFind, Len(smComment))
'            smComment = slTemp
'        End If
'        slStrToFind = "\deflangfe1033"
'        ilLenOfStrToFind = Len(slStrToFind)
'        ilPos = InStr(1, smComment, slStrToFind)
'        If ilPos > 0 Then
'            slTemp = Left(smComment, ilPos - 1)
'            slTemp = slTemp + Mid(smComment, ilPos + ilLenOfStrToFind, Len(smComment))
'            smComment = slTemp
'        End If
        'tmCsf.iStrLen = Len(smComment)
        tmCsf.sType = "S"
        tmCsf.iAdfCode = igCopyInvAdfCode
        tmCsf.sComment = Trim$(smComment) & Chr$(0) '& Chr$(0) 'sgTB
    End If
    Exit Sub
mMoveCtrlToRecErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveDupl                       *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveDupl()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilBoxNo As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    For ilBoxNo = imLBCtrls To UBound(tmDuplCtrls) Step 1
        Select Case ilBoxNo 'Branch on box type (control)
            Case INVNOINDEX 'Name
                slStr = Trim$(tmDuplCif.sName)
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case CUTINDEX   'Cut #
                slStr = Trim$(tmDuplCif.sCut)
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case REELINDEX   'Reel #
                slStr = Trim$(tmDuplCif.sReel)
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case LENINDEX   'Length
                slStr = Trim$(str$(tmDuplCif.iLen))
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case ANNINDEX 'announcer
                slStr = ""
                If tmDuplCif.iMnfAnn > 0 Then
                    For ilLoop = 0 To UBound(tmAnnCode) - 1 Step 1 'lbcAnnCode.ListCount - 1 Step 1
                        slNameCode = tmAnnCode(ilLoop).sKey   'lbcAnnCode.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = tmDuplCif.iMnfAnn Then
                            ilRet = gParseItem(slNameCode, 1, "\", slStr)
                            Exit For
                        End If
                    Next ilLoop
                End If
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case COMPINDEX 'Competitive
                slStr = ""
                If tmDuplCif.iMnfComp(0) > 0 Then
                    For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1 'lbcCompCode.ListCount - 1 Step 1
                        slNameCode = tgCompCode(ilLoop).sKey    'tgCompCode(imSave(1 + ilLoop) - 2).sKey   'lbcCompCode.List(imSave(1 + ilLoop) - 2)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = tmDuplCif.iMnfComp(0) Then
                            ilRet = gParseItem(slNameCode, 1, "\", slStr)
                            Exit For
                        End If
                    Next ilLoop
                End If
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case COMPINDEX + 1 'Competitive
                slStr = ""
                If tmDuplCif.iMnfComp(1) > 0 Then
                    For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1 'lbcCompCode.ListCount - 1 Step 1
                        slNameCode = tgCompCode(ilLoop).sKey    'tgCompCode(imSave(1 + ilLoop) - 2).sKey   'lbcCompCode.List(imSave(1 + ilLoop) - 2)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = tmDuplCif.iMnfComp(1) Then
                            ilRet = gParseItem(slNameCode, 1, "\", slStr)
                            Exit For
                        End If
                    Next ilLoop
                End If
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case PRODINDEX
                slStr = Trim$(tmDuplCpf.sName)
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case ISCIINDEX
                slStr = Trim$(tmDuplCpf.sISCI)
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case PURGEINDEX
                If tmDuplCif.sPurged = "P" Then
                    slStr = "Purged"
                ElseIf tmDuplCif.sPurged = "H" Then
                    slStr = "History"
                Else
                    slStr = "Active"
                End If
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case CARTINDEX
                Select Case tmDuplCif.sCartDisp
                    Case "S"
                        ilIndex = 1
                    Case "P"
                        ilIndex = 2
                    Case "A"
                        ilIndex = 3
                    Case Else
                        ilIndex = 0
                End Select
                If ilIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcCartDisp.List(ilIndex)
                End If
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case TAPEINDEX
                Select Case tmDuplCif.sTapeDisp
                    Case "R"
                        ilIndex = 1
                    Case "D"
                        ilIndex = 2
                    Case "A"
                        ilIndex = 3
                    Case Else
                        ilIndex = 0
                End Select
                If ilIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcTapeDisp.List(ilIndex)
                End If
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case NOAIRINDEX
                slStr = Trim$(str$(tmDuplCif.iNoTimesAir))
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case TITLEINDEX
                slStr = Trim$(tmDuplCpf.sCreative)
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case TAPEININDEX
                If imSave(7) = 0 Then
                    slStr = "No"
                ElseIf imSave(7) = 1 Then
                    slStr = "Yes"
                Else
                    slStr = ""
                End If
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case TAPEAPPINDEX
                If ((Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT) Then
                    If imSave(8) = 0 Then
                        slStr = "Not Sent"
                    ElseIf imSave(8) = 1 Then
                        slStr = "Produced"
                    ElseIf (imSave(8) = 2) Then
                        slStr = "Sent"
                    ElseIf (imSave(8) = 3) Then
                        slStr = "Hold"
                    Else
                        slStr = ""
                    End If
                Else
                    If imSave(8) = 0 Then
                        slStr = "No"
                    ElseIf imSave(8) = 1 Then
                        slStr = "Yes"
                    Else
                        slStr = ""
                    End If
                End If
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case SCRIPTINDEX
                slStr = gStripChr0(tmDuplCsf.sComment)
                'If tmDuplCsf.iStrLen > 0 Then
                If slStr <> "" Then
                '    slStr = Trim$(Left$(tmDuplCsf.sComment, tmDuplCsf.iStrLen))
                    edcHistComment.SetText slStr
                    'slStr = edcHistComment.TextOnly
                    slStr = Left$(edcHistComment.TextOnly, 80)
                    ilPos = InStr(slStr, sgLF)
                    If ilPos = 2 Then
                        slStr = Mid$(slStr, ilPos + 1)
                    End If
                    ilPos = InStr(slStr, sgCR)
                    If ilPos > 0 Then
                        slStr = Left$(slStr, ilPos - 1)
                    End If
                'Else
                '    slStr = ""
                End If
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case ENTEREDINDEX
                gUnpackDate tmDuplCif.iDateEntrd(0), tmDuplCif.iDateEntrd(1), slStr
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case LASTUSEDINDEX
                gUnpackDate tmDuplCif.iUsedDate(0), tmDuplCif.iUsedDate(1), slStr
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case EARLROTINDEX
                gUnpackDate tmDuplCif.iRotStartDate(0), tmDuplCif.iRotStartDate(1), slStr
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case LATROTINDEX
                gUnpackDate tmDuplCif.iRotEndDate(0), tmDuplCif.iRotEndDate(1), slStr
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
            Case PDATEINDEX
                gUnpackDate tmDuplCif.iPurgeDate(0), tmDuplCif.iPurgeDate(1), slStr
                gSetShow pbcInv, slStr, tmDuplCtrls(ilBoxNo)
        End Select
    Next ilBoxNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer    'Return call status
    Dim slStr As String
    If smUseCartNo <> "N" Then
        smSave(1) = Trim$(tmCif.sName)
        smSave(2) = Trim$(tmCif.sCut)
    Else
        smSave(1) = ""
        smSave(2) = ""
    End If
    smSave(3) = Trim$(tmCif.sReel)
    If tmCif.iLen <> 0 Then
        smSave(4) = Trim$(str$(tmCif.iLen))
        gFindMatch smSave(4), 0, lbcLen    'Determine if name exist
        If gLastFound(lbcLen) <> -1 Then   'Name found
            imChgMode = True
            lbcLen.ListIndex = gLastFound(lbcLen)
            imChgMode = False
        End If
    Else
        smSave(4) = ""
    End If
    smOrigLen = smSave(4)
    smOrigPurge = Trim$(tmCif.sPurged)
    gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), smOrigRotDate
    If smOrigRotDate = "" Then
        smOrigRotDate = ": Unused"
    Else
        smOrigRotDate = ": " & smOrigRotDate
    End If
    imSave(9) = -1
    smOrigAnn = ""
    If tmCif.iMnfAnn > 0 Then
        For ilLoop = 0 To UBound(tmAnnCode) - 1 Step 1 'lbcAnnCode.ListCount - 1 Step 1
            slNameCode = tmAnnCode(ilLoop).sKey  'lbcAnnCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmCif.iMnfAnn Then
                imSave(9) = ilLoop + 2
                imChgMode = True
                lbcAnn.ListIndex = imSave(9)
                imChgMode = False
                ilRet = gParseItem(slNameCode, 1, "\", smOrigAnn)
                Exit For
            End If
        Next ilLoop
    End If
    'Read in product/ISCI
    smSave(5) = Trim$(tmCpf.sName)
    smSave(6) = Trim$(tmCpf.sISCI)
    smSave(7) = Trim$(tmCpf.sCreative)
    smSave(8) = ""
    If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
        If tmCif.sCleared = "Y" Or tmCif.sCleared = "S" Then
            gUnpackDate tmCif.iInvSentDate(0), tmCif.iInvSentDate(1), smSave(8)
            If gDateValue(smSave(8)) = gDateValue("1/1/1970") Then
                smSave(8) = ""
            End If
        End If
    End If
    smOrigSentDate = smSave(8)
    imSave(1) = -1
    smOrigComp0 = ""
    If tmCif.iMnfComp(0) > 0 Then
        For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1 'lbcCompCode.ListCount - 1 Step 1
            slNameCode = tgCompCode(ilLoop).sKey   'lbcCompCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmCif.iMnfComp(0) Then
                imSave(1) = ilLoop + 2
                ilRet = gParseItem(slNameCode, 1, "\", smOrigComp0)
                imChgMode = True
                lbcComp(0).ListIndex = imSave(1)
                imChgMode = False
                Exit For
            End If
        Next ilLoop
    End If
    imSave(2) = -1
    If tmCif.iMnfComp(1) > 0 Then
        For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1 'lbcCompCode.ListCount - 1 Step 1
            slNameCode = tgCompCode(ilLoop).sKey   'lbcCompCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmCif.iMnfComp(1) Then
                imSave(2) = ilLoop + 2
                ilRet = gParseItem(slNameCode, 1, "\", smOrigComp1)
                imChgMode = True
                lbcComp(1).ListIndex = imSave(1)
                imChgMode = False
                Exit For
            End If
        Next ilLoop
    End If
    If Trim$(tmCif.sPurged) = "P" Then
        imSave(3) = 1
    ElseIf Trim$(tmCif.sPurged) = "H" Then
        imSave(3) = 2
    ElseIf Trim$(tmCif.sPurged) = "A" Then
        imSave(3) = 0
    Else
        imSave(3) = -1
    End If
    Select Case Trim$(tmCif.sCartDisp)
        Case "N"
            imSave(4) = 0
        Case "S"
            imSave(4) = 1
        Case "P"
            imSave(4) = 2
        Case "A"
            imSave(4) = 3
        Case Else
            imSave(4) = -1
    End Select
    imChgMode = True
    lbcCartDisp.ListIndex = imSave(4)
    imChgMode = False
    Select Case Trim$(tmCif.sTapeDisp)
        Case "N"
            imSave(5) = 0
        Case "R"
            imSave(5) = 1
        Case "D"
            imSave(5) = 2
        Case "A"
            imSave(5) = 3
        Case Else
            imSave(5) = -1
    End Select
    imChgMode = True
    lbcTapeDisp.ListIndex = imSave(5)
    imChgMode = False
    imSave(6) = tmCif.iNoTimesAir
    If Trim$(tmCif.sHouse) = "Y" Then
        imSave(7) = 1
    ElseIf Trim$(tmCif.sHouse) = "N" Then
        imSave(7) = 0
    Else
        imSave(7) = -1
    End If
    
    If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
        If Trim$(tmCif.sCleared) = "Y" Then
            imSave(8) = 1
        ElseIf Trim$(tmCif.sCleared) = "N" Then
            imSave(8) = 0
        ElseIf (Trim$(tmCif.sCleared) = "S") Then
            imSave(8) = 2
        ElseIf (Trim$(tmCif.sCleared) = "H") Then
            imSave(8) = 3
        Else
            imSave(8) = -1
        End If
    Else
        If Trim$(tmCif.sCleared) = "Y" Then
            imSave(8) = 1
        ElseIf Trim$(tmCif.sCleared) = "N" Then
            imSave(8) = 0
        Else
            imSave(8) = -1
        End If
    End If
    slStr = gStripChr0(tmCsf.sComment)
    'If tmCsf.iStrLen > 0 Then
    If slStr <> "" Then
        smComment = slStr   'Trim$(Left$(tmCsf.sComment, tmCsf.iStrLen))
        edcComment.MaxLength = 5000
        edcComment.SetText (smComment)
        smCommentTextOnly = edcComment.TextOnly
    'Else
    '    smComment = ""
    End If
    smOrigComment = smComment
    gUnpackDate tmCif.iDateEntrd(0), tmCif.iDateEntrd(1), smDateSave(1)
    gUnpackDate tmCif.iUsedDate(0), tmCif.iUsedDate(1), smDateSave(2)
    gUnpackDate tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), smDateSave(3)
    gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), smDateSave(4)
    gUnpackDate tmCif.iPurgeDate(0), tmCif.iPurgeDate(1), smDateSave(5)
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName() As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim llCode As Long
    Dim ilRet As Integer
    Dim tlCif As CIF

    If smUseCartNo = "N" Then
        mOKName = True
        Exit Function
    End If
    If Trim$(tmMcf.sReuse) <> "N" Then
        mOKName = True
        Exit Function
    End If
    If smSave(1) <> "" Then    'Test name
        If smSave(2) <> "" Then
            slStr = smSave(1) & "-" & smSave(2)
        Else
            slStr = smSave(1)
        End If
        If imProcMode = 0 Then
            If smOrigPurge = "P" Then
                slStr = slStr & ": Purged"
            Else
                slStr = slStr & smOrigRotDate
            End If
        Else
            slStr = Trim$(tmMcf.sName) & slStr
        End If
        slStr = Trim$(slStr)
        gFindMatch slStr, 0, cbcInv    'Determine if name exist
        If gLastFound(cbcInv) <> -1 Then   'Name found
            If gLastFound(cbcInv) <> imCifIndex Then
                If slStr = cbcInv.List(gLastFound(cbcInv)) Then
                    Beep
                    MsgBox "Name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = 1
                    mEnableBox imBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
        If ((imCifIndex < 1) And (imProcMode = 0)) Or ((imCifIndex < 1) And (Trim$(cbcInv.List(0)) = "[New]")) Then
            llCode = 0
        Else
            If Trim$(cbcInv.List(0)) = "[New]" Then
                slNameCode = tmInvNameCode(imCifIndex - 1).sKey   'lbcInvCode.List(imCifIndex - 1)
            Else
                slNameCode = tmInvNameCode(imCifIndex).sKey
            End If
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            llCode = Val(slCode)
        End If
        'Test if name in use by this same advertiser.
        tmCifSrchKey1.iMcfCode = tmMcf.iCode
        tmCifSrchKey1.sName = smSave(1)
        tmCifSrchKey1.sCut = smSave(2)
        ilRet = btrGetGreaterOrEqual(hmCif, tlCif, imCifRecLen, tmCifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tlCif.iMcfCode = tmMcf.iCode) And (Trim$(tlCif.sName) = Trim$(smSave(1))) And (Trim$(tlCif.sCut) = Trim$(smSave(2)))
            If (igCopyInvAdfCode = tlCif.iAdfCode) And (tlCif.sPurged = "A") And (tlCif.lCode <> llCode) Then
                Beep
                MsgBox "Name already defined and active for this Advertiser, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                mSetShow imBoxNo
                mSetChg imBoxNo
                imBoxNo = 1
                mEnableBox imBoxNo
                mOKName = False
                Exit Function
            End If
            ilRet = btrGetNext(hmCif, tlCif, imCifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop

    End If
    mOKName = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintTitle                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint Title                    *
'*                                                     *
'*******************************************************
Private Sub mPaintTitle(ilImage As Integer)
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    If ilImage = 0 Then
        llColor = pbcInv.ForeColor
        slFontName = pbcInv.FontName
        flFontSize = pbcInv.FontSize
        pbcInv.ForeColor = BLUE
        pbcInv.FontBold = False
        pbcInv.FontSize = 7
        pbcInv.FontName = "Arial"
        pbcInv.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcInv.CurrentX = tmCtrls(PRODINDEX).fBoxX + 15  'fgBoxInsetX
        pbcInv.CurrentY = tmCtrls(PRODINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If tgSpf.sUseProdSptScr = "P" Then  'Short Title
            pbcInv.Print "Short Title"
        Else
            pbcInv.Print "Product"
        End If
        pbcInv.CurrentX = tmCtrls(TITLEINDEX).fBoxX + 15  'fgBoxInsetX
        pbcInv.CurrentY = tmCtrls(TITLEINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        'If (Asc(tgSpf.sUsingFeatures10) And WegenerIPump) <> WegenerIPump Then
            pbcInv.Print "Creative Title"
        'Else
        '    pbcInv.Print "Audio Name Title"
        'End If
        pbcInv.CurrentX = tmCtrls(TAPEAPPINDEX).fBoxX + 15  'fgBoxInsetX
        pbcInv.CurrentY = tmCtrls(TAPEAPPINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
            pbcInv.Print "Copy Produced"
        Else
            If tgSpf.sTapeShowForm = "C" Then  'Tape Carted or Approved
                pbcInv.Print "Tape Carted"
            Else
                pbcInv.Print "Tape Approved"
            End If
        End If
        pbcInv.CurrentX = tmCtrls(INVSENTDATEINDEX).fBoxX + 15  'fgBoxInsetX
        pbcInv.CurrentY = tmCtrls(INVSENTDATEINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
            pbcInv.Print "Action Date"
        Else
            pbcInv.Print ""
        End If
        pbcInv.CurrentX = tmCtrls(REELINDEX).fBoxX + 15  'fgBoxInsetX
        pbcInv.CurrentY = tmCtrls(REELINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
            pbcInv.Print "Contract #"
        Else
            pbcInv.Print "Reel #"
        End If
        pbcInv.FontSize = flFontSize
        pbcInv.FontName = slFontName
        pbcInv.FontSize = flFontSize
        pbcInv.ForeColor = llColor
        pbcInv.FontBold = True
    Else
        llColor = pbcDupl.ForeColor
        slFontName = pbcDupl.FontName
        flFontSize = pbcDupl.FontSize
        pbcDupl.ForeColor = BLUE
        pbcDupl.FontBold = False
        pbcDupl.FontSize = 7
        pbcDupl.FontName = "Arial"
        pbcDupl.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcDupl.CurrentX = tmDuplCtrls(PRODINDEX).fBoxX + 15  'fgBoxInsetX
        pbcDupl.CurrentY = tmDuplCtrls(PRODINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If tgSpf.sUseProdSptScr = "P" Then  'Short Title
            pbcDupl.Print "Short Title"
        Else
            pbcDupl.Print "Product"
        End If
        pbcDupl.CurrentX = tmDuplCtrls(TITLEINDEX).fBoxX + 15  'fgBoxInsetX
        pbcDupl.CurrentY = tmDuplCtrls(TITLEINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcDupl.Print "Creative Title"
        pbcDupl.CurrentX = tmDuplCtrls(TAPEAPPINDEX).fBoxX + 15  'fgBoxInsetX
        pbcDupl.CurrentY = tmDuplCtrls(TAPEAPPINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
            pbcDupl.Print "Copy Produced"
        Else
            If tgSpf.sTapeShowForm = "C" Then  'Tape Carted
                pbcDupl.Print "Tape Carted"
            Else
                pbcDupl.Print "Tape Approved"
            End If
        End If
        pbcDupl.CurrentX = tmCtrls(INVSENTDATEINDEX).fBoxX + 15  'fgBoxInsetX
        pbcDupl.CurrentY = tmCtrls(INVSENTDATEINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
            pbcDupl.Print "Action Date"
        Else
            pbcDupl.Print ""
        End If
        pbcDupl.CurrentX = tmCtrls(REELINDEX).fBoxX + 15  'fgBoxInsetX
        pbcDupl.CurrentY = tmCtrls(REELINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
            pbcDupl.Print "Contract #"
        Else
            pbcDupl.Print "Reel #"
        End If
        pbcDupl.FontSize = flFontSize
        pbcDupl.FontName = slFontName
        pbcDupl.FontSize = flFontSize
        pbcDupl.ForeColor = llColor
        pbcDupl.FontBold = True
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    slCommand = sgCommandStr    'Command$
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    '    imShowHelpMsg = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
    '    imShowHelpMsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone CopyInv, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igCopyInvCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    'Copy^Test^NOHELP\Guide\5\6810\26\0\0\0
    '    'Copy^Prod^NOHELP\Guide\5\188\37\0\0\0
    '    igCopyInvCallSource = 5 'CALLNONE
    '    lgCopyInvCifCode = 2173 '3108    '6810'1=change'0=new
    '    igCopyInvAdfCode = 81   '26'46'3 (5=800 800Len)
    '    igCopyInvEtfCode = 0
    '    igCopyInvEnfCode = 0
    '    imInvStatus = 0 '0=Active; 4=Purged; 5=History
    '    Exit Sub
    'End If
    If igCopyInvCallSource <> CALLNONE Then  'If advertiser code number
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            lgCopyInvCifCode = Val(slStr)   'Change mode
        Else
            lgCopyInvCifCode = 0    'New mode
        End If
        'Either by advertiser or event type/event name
        ilRet = gParseItem(slCommand, 5, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            igCopyInvAdfCode = Val(slStr)
        Else
            igCopyInvAdfCode = 0
        End If
        ilRet = gParseItem(slCommand, 6, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            igCopyInvEtfCode = Val(slStr)
        Else
            igCopyInvEtfCode = 0
        End If
        ilRet = gParseItem(slCommand, 7, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            igCopyInvEnfCode = Val(slStr)
        Else
            igCopyInvEnfCode = 0
        End If
        If lgCopyInvCifCode > 0 Then
            ilRet = gParseItem(slCommand, 8, "\", slStr)
            If ilRet = CP_MSG_NONE Then
                imInvStatus = Val(slStr)
                If imInvStatus = 1 Then
                    imInvStatus = 4
                ElseIf imInvStatus = 2 Then
                    imInvStatus = 5
                Else
                    imInvStatus = 0
                End If
            Else
                imInvStatus = 0 'Active
            End If
        Else
            imInvStatus = 0 'Active
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mProdBranch                     *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      advertiser product and process *
'*                      communication back from        *
'*                      advertiser product             *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mProdBranch() As Integer
'
'   ilRet = mProdBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    'If Not gWinRoom(igNoExeWinRes(ADVTPRODEXE)) Then
    '    imDoubleClickName = False
    '    mProdBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
'        ilRet = gOptionalLookAhead(edcDropDown, lbcProd, imBSMode, slStr)
    If (Not imDoubleClickName) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mProdBranch = False
        Exit Function
    End If
    'Screen.MousePointer = vbHourGlass  'Wait
    igAdvtProdCallSource = CALLSOURCEADVERTISER
    'sgAdvtProdName = Trim$(tmAdf.sName) 'cbcSelect.List(imSelectedIndex)
    If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
        sgAdvtProdName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & "/Direct"
    Else
        sgAdvtProdName = Trim$(tmAdf.sName)
    End If
    If edcDropDown.Text = "[New]" Then
        sgAdvtProdName = sgAdvtProdName & "\" & " "
    Else
        sgAdvtProdName = sgAdvtProdName & "\" & Trim$(edcDropDown.Text)
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "CopyInv^Test\" & sgUserName & "\" & Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
        Else
            slStr = "CopyInv^Prod\" & sgUserName & "\" & Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "CopyInv^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    '    Else
    '        slStr = "CopyInv^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    '    End If
    'End If
    sgCommandStr = slStr
    If tgSpf.sUseProdSptScr = "P" Then
        'lgShellRet = Shell(sgExePath & "ShtTitle.Exe " & slStr, 1)
        ShtTitle.Show vbModal
    Else
        'lgShellRet = Shell(sgExePath & "AdvtProd.Exe " & slStr, 1)
        AdvtProd.Show vbModal
    End If
    'CopyInv.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgAdvtProdName)
    igAdvtProdCallSource = Val(sgAdvtProdName)
    ilParse = gParseItem(slStr, 2, "\", sgAdvtProdName)
    'CopyInv.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mProdBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igAdvtProdCallSource = CALLDONE Then  'Done
        igAdvtProdCallSource = CALLNONE
'        gSetMenuState True
        lbcProd.Clear
        sgProdCodeTag = ""
        mProdPop
        If imTerminate Then
            mProdBranch = False
            Exit Function
        End If
        gFindMatch sgAdvtProdName, 1, lbcProd
        sgAdvtProdName = ""
        If gLastFound(lbcProd) > 0 Then
            imChgMode = True
            lbcProd.ListIndex = gLastFound(lbcProd)
            edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
            imChgMode = False
            mProdBranch = False
            mSetChg PRODINDEX
        Else
            imChgMode = True
            lbcProd.ListIndex = -1
            edcDropDown.Text = sgAdvtProdName
            imChgMode = False
            mSetChg PRODINDEX
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igAdvtProdCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igAdvtProdCallSource = CALLNONE
        sgAdvtProdName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igAdvtProdCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igAdvtProdCallSource = CALLNONE
        sgAdvtProdName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mProdPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate advertiser product    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mProdPop()
'
'   mProdPop
'   Where:
'       igCopyInvAdfCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    If igCopyInvAdfCode <= 0 Then
        If lbcProd.ListCount <= 0 Then
            lbcProd.AddItem "[None]", 0  'Force as first item on list
        End If
        Exit Sub
    End If
    ilIndex = lbcProd.ListIndex
    If ilIndex > 0 Then
        slName = lbcProd.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtProdBox(CopyInv, igCopyInvAdfCode, lbcProd, lbcProdCode)
    If tgSpf.sUseProdSptScr = "P" Then
        ilRet = gPopShortTitleBox(CopyInv, igCopyInvAdfCode, lbcProd, tgProdCode(), sgProdCodeTag)
    Else
        ilRet = gPopAdvtProdBox(CopyInv, igCopyInvAdfCode, lbcProd, tgProdCode(), sgProdCodeTag)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mProdPopErr
        gCPErrorMsg ilRet, "mProdPop (gPopAdvtProdBox)", CopyInv
        On Error GoTo 0
        lbcProd.AddItem "[None]", 0  'Force as first item on list
'        lbcProd.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcProd
            If gLastFound(lbcProd) > 0 Then
                lbcProd.ListIndex = gLastFound(lbcProd)
            Else
                lbcProd.ListIndex = -1
            End If
        Else
            lbcProd.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mProdPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadCif                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Sub mReadCif(ilForUpdate As Integer)
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    Dim slDate As String
    Dim slRotDate As String
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilIndex As Integer
    
    imBypassPurge = True
    If smUseCartNo <> "N" Then
        If Trim$(tmMcf.sReuse) = "N" Then
            If ((imCifIndex < 1) And (imProcMode = 0)) Or ((imCifIndex < 0) And (imProcMode <> 0)) Then
                tmCif.sName = ""
                tmCif.sCut = ""
                tmCif.iLen = 0
                tmCif.iNoTimesAir = 0
                mClearPartOfCopy
                Exit Sub
            End If
            If cbcInv.List(0) = "[New]" Then
                slNameCode = tmInvNameCode(imCifIndex - 1).sKey   'lbcInvCode.List(imCifIndex - 1)
            Else
                slNameCode = tmInvNameCode(imCifIndex).sKey   'lbcInvCode.List(imCifIndex - 1)
            End If
        Else
            If imCifIndex < 0 Then
                tmCif.sName = ""
                tmCif.sCut = ""
                tmCif.iLen = 0
                mClearPartOfCopy
                Exit Sub
            End If
            slNameCode = tmInvNameCode(imCifIndex).sKey   'lbcInvCode.List(imCifIndex)
        End If
    Else
        'If imCifIndex < 1 Then
        If ((imCifIndex < 1) And (imProcMode = 0)) Or ((imCifIndex < 0) And (imProcMode <> 0)) Then
            tmCif.sName = ""
            tmCif.sCut = ""
            tmCif.iLen = 0
            tmCif.iNoTimesAir = 0
            mClearPartOfCopy
            Exit Sub
        End If
        If cbcInv.List(0) = "[New]" Then
            slNameCode = tmInvNameCode(imCifIndex - 1).sKey   'lbcInvCode.List(imCifIndex - 1)
        Else
            slNameCode = tmInvNameCode(imCifIndex).sKey   'lbcInvCode.List(imCifIndex - 1)
        End If
    End If
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadCifErr
    gCPErrorMsg ilRet, "mReadCifErr (gParseItem field 2)", CopyInv
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmCifSrchKey.lCode = CLng(slCode)
    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadCifErr
    gBtrvErrorMsg ilRet, "mReadCifErr (btrGetEqual: Copy Inventory)", CopyInv
    On Error GoTo 0
    If tmCif.lcpfCode > 0 Then
        tmCpfSrchKey.lCode = tmCif.lcpfCode
        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
        If ilRet <> BTRV_ERR_NONE Then
            tmCpf.sName = ""
            tmCpf.sISCI = ""
            tmCpf.sCreative = ""
        End If
    Else
        tmCpf.lCode = 0
        tmCpf.sName = ""
        tmCpf.sISCI = ""
        tmCpf.sCreative = ""
    End If
    tmCsfSrchKey.lCode = tmCif.lCsfCode
    If tmCif.lCsfCode <> 0 Then
        tmCsf.sComment = ""
        imCsfRecLen = Len(tmCsf) '5011
        ilRet = btrGetEqual(hmCsf, tmCsf, imCsfRecLen, tmCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
        If ilRet <> BTRV_ERR_NONE Then
            tmCsf.lCode = 0
            tmCsf.sComment = ""
            'tmCsf.iStrLen = 0
        End If
    Else
        tmCsf.lCode = 0
        tmCsf.sComment = ""
        'tmCsf.iStrLen = 0
    End If
    gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slRotDate
    If slRotDate <> "" Then
        slDate = Format$(gNow(), "m/d/yy")
        If gDateValue(slRotDate) < gDateValue(slDate) Then
            imBypassPurge = False
        Else
            imBypassPurge = True
        End If
    Else
        imBypassPurge = False
    End If
    For ilLoop = LBound(tmMediaCode) To UBound(tmMediaCode) - 1 Step 1
        slNameCode = tmMediaCode(ilLoop).sKey    'lbcMediaCode.List(imMcfIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If tmCif.iMcfCode = Val(Trim$(slCode)) Then
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            For ilIndex = 0 To cbcMedia.ListCount - 1 Step 1
                If StrComp(slName, cbcMedia.List(ilIndex), 1) = 0 Then
                    imChgModeMedia = True
                    cbcMedia.ListIndex = ilIndex
                    imMcfIndex = cbcMedia.ListIndex
                    mReadMcf
                    imChgModeMedia = False
                    Exit For
                End If
            Next ilIndex
            Exit For
        End If
    Next ilLoop
    Exit Sub
mReadCifErr:
    On Error GoTo 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadDuplCif                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Sub mReadDuplCif()
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    If imDuplIndex < 0 Then
        Exit Sub
    End If
    slNameCode = lbcDuplInvCode.List(imDuplIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadDuplCifErr
    gCPErrorMsg ilRet, "mReadDuplCifErr (gParseItem field 2)", CopyInv
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmCifSrchKey.lCode = CLng(slCode)
    ilRet = btrGetEqual(hmCif, tmDuplCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mReadDuplCifErr
    gBtrvErrorMsg ilRet, "mReadDuplCifErr (btrGetEqual: Copy Inventory)", CopyInv
    On Error GoTo 0
    If tmDuplCif.lcpfCode > 0 Then
        tmCpfSrchKey.lCode = tmDuplCif.lcpfCode
        ilRet = btrGetEqual(hmCpf, tmDuplCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmDuplCpf.sName = ""
            tmDuplCpf.sISCI = ""
            tmDuplCpf.sCreative = ""
        End If
    Else
        tmDuplCpf.sName = ""
        tmDuplCpf.sISCI = ""
        tmDuplCpf.sCreative = ""
    End If
    tmCsfSrchKey.lCode = tmDuplCif.lCsfCode
    If tmDuplCif.lCsfCode <> 0 Then
        tmDuplCsf.sComment = ""
        imCsfRecLen = Len(tmDuplCsf) '5011
        ilRet = btrGetEqual(hmCsf, tmDuplCsf, imCsfRecLen, tmCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmDuplCsf.lCode = 0
            tmDuplCsf.sComment = ""
            'tmDuplCsf.iStrLen = 0
        End If
    Else
        tmDuplCsf.lCode = 0
        tmDuplCsf.sComment = ""
        'tmDuplCsf.iStrLen = 0
    End If
    Exit Sub
mReadDuplCifErr:
    On Error GoTo 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadMcf                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Sub mReadMcf()
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    If smUseCartNo <> "N" Then
        If imMcfIndex < 0 Then
            Exit Sub
        End If
        slNameCode = tmMediaCode(imMcfIndex).sKey    'lbcMediaCode.List(imMcfIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mReadMcfErr
        gCPErrorMsg ilRet, "mReadMcfErr (gParseItem field 2)", CopyInv
        On Error GoTo 0
        slCode = Trim$(slCode)
        tmMcfSrchKey.iCode = CInt(slCode)
        ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mReadMcfErr
        gBtrvErrorMsg ilRet, "mReadMcfErr (btrGetEqual: Media Code)", CopyInv
        On Error GoTo 0
        tmMefSrchKey1.iMcfCode = CInt(slCode)
        ilRet = btrGetEqual(hmMef, tmMef, imMefRecLen, tmMefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmMef.iCode = 0
            tmMef.sPrefix = ""
            tmMef.sSuffix = ""
            tmMef.sEventType = ""
            tmMef.sNetworkID = ""
            tmMef.sNameSpace = ""
        End If
        cbcDuplInv.Clear
        lbcDuplInvCode.Clear
        imDuplIndex = -1
        If Trim$(tmMcf.sReuse) = "N" Then
            plcDupl.Enabled = False
            pbcDupl.Enabled = False
            cbcDuplInv.Enabled = False
        Else
            plcDupl.Enabled = True
            pbcDupl.Enabled = True
            cbcDuplInv.Enabled = True
        End If
        If imMcfCodeForSort <> tmMcf.iCode Then
            If tmMcf.sSortCart = "C" Then
                imSortCart = 1
            Else
                imSortCart = 0
            End If
            imMcfCodeForSort = tmMcf.iCode
        End If
        pbcSort.Cls
        pbcSort_Paint
    Else
        plcDupl.Enabled = False
        pbcDupl.Enabled = False
        cbcDuplInv.Enabled = False
    End If
    Exit Sub
mReadMcfErr:
    On Error GoTo 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
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
    Dim slName As String    'Name
    Dim ilRet As Integer
    Dim ilErrRet As Integer
    Dim slMsg As String
    Dim ilNewInv As Integer
    Dim tlCif As CIF
    Dim ilCifRet As Integer
    Dim slDate As String
    Dim llSvCpfCode As Long
    Dim llSvCsfCode As Long
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilAddNew As Integer

    mSetShow imBoxNo
    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    If Not mCheckStatus() Then
        mSaveRec = False
        Exit Function
    End If
    If Not mInvPurgeOK() Then
        Screen.MousePointer = vbHourglass
        pbcInv.Cls
        mClearCtrlFields
        mInvPop
        Screen.MousePointer = vbDefault
        mSaveRec = False
        mSetCommands
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    gGetSyncDateTime slSyncDate, slSyncTime
    ilNewInv = False
'    If smUseCartNo <> "N" Then
'        If Trim$(tmMcf.sReuse) = "N" Then
'            If imCifIndex < 1 Then
'                ilNewInv = True
'            End If
'        Else
'            If imProcMode = 0 Then  'lgCopyInvCifCode <= 0 Then
'                ilNewInv = True
'            End If
'        End If
'    Else
'        If imCifIndex < 1 Then
'            ilNewInv = True
'        End If
'    End If
    If smUseCartNo <> "N" Then
        If Trim$(tmMcf.sReuse) = "N" Then
            If (imCifIndex < 1) And (imProcMode = 0) Then
                ilNewInv = True
            End If
        Else
            If imProcMode = 0 Then  'lgCopyInvCifCode <= 0 Then
                ilNewInv = True
            End If
        End If
    Else
        If (imCifIndex < 1) And (imProcMode = 0) Then
            ilNewInv = True
        End If
    End If
    Do  'Loop until record updated or added
        If Not ilNewInv Then
            mReadCif SETFORWRITE
        Else
            If smUseCartNo <> "N" Then
                If (Trim$(tmMcf.sReuse) <> "N") And (tmCif.lCode > 0) Then 'Update
                    mReadCif SETFORWRITE
                End If
            End If
        End If
        mMoveCtrlToRec True
        imCsfRecLen = Len(tmCsf) '- Len(tmCsf.sComment) + Len(Trim$(tmCsf.sComment)) ' + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
        If ilNewInv Then 'New selected
            'If Len(Trim$(tmCsf.sComment)) > 2 Then  'imCsfRecLen - 2 > 9 Then '-2 so the control character at the end is not counted
            If gStripChr0(tmCsf.sComment) <> "" Then
                tmCsf.lCode = 0 'Autoincrement
                ilRet = btrInsert(hmCsf, tmCsf, imCsfRecLen, INDEXKEY0)
            Else
                tmCsf.lCode = 0
                ilRet = BTRV_ERR_NONE
            End If
            slMsg = "mSaveRec (btrInsert: Comment)"
        Else 'Old record-Update
            'If Len(Trim$(tmCsf.sComment)) > 2 Then  'imCsfRecLen - 2 > 9 Then '-2 so the control character at the end is not counted
            If gStripChr0(tmCsf.sComment) <> "" Then
                If tmCsf.lCode = 0 Then
                    tmCsf.lCode = 0 'Autoincrement
                    ilRet = btrInsert(hmCsf, tmCsf, imCsfRecLen, INDEXKEY0)
                Else
                    ilRet = btrUpdate(hmCsf, tmCsf, imCsfRecLen)
                End If
            Else
                If tmCif.lCsfCode <> 0 Then
                    ilRet = btrDelete(hmCsf)
                End If
                tmCsf.lCode = 0
            End If
            slMsg = "mSaveRec (btrUpdate: Comment)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, CopyInv
    On Error GoTo 0
    tmCif.lCsfCode = tmCsf.lCode
    If igCopyInvAdfCode > 0 Then
        If Trim$(smSave(5)) <> "" Then
            gFindMatch smSave(5), 0, lbcProd
            If gLastFound(lbcProd) < 0 Then
                If tgSpf.sUseProdSptScr = "P" Then
                    tmSif.lCode = 0
                    tmSif.iAdfCode = igCopyInvAdfCode
                    tmSif.sName = smSave(5)
                    tmSif.sState = "A"
                    tmSif.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                    tmSif.iRemoteID = tgUrf(0).iRemoteUserID
                    tmSif.lAutoCode = tmSif.lCode
                    ilRet = btrInsert(hmSif, tmSif, imSifRecLen, INDEXKEY0)
                    slMsg = "mSaveRec (btrInsert:Short Title)"
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, CopyInv
                    On Error GoTo 0
                    Do
                        tmSif.iRemoteID = tgUrf(0).iRemoteUserID
                        tmSif.lAutoCode = tmSif.lCode
                        gPackDate slSyncDate, tmSif.iSyncDate(0), tmSif.iSyncDate(1)
                        gPackTime slSyncTime, tmSif.iSyncTime(0), tmSif.iSyncTime(1)
                        ilRet = btrUpdate(hmSif, tmSif, imSifRecLen)
                        slMsg = "mSaveRec (btrUpdate:Short Title)"
                    Loop While ilRet = BTRV_ERR_CONFLICT
                Else
                    Do  'Loop until record updated or added
                        tmPrf.lCode = 0
                        tmPrf.iAdfCode = igCopyInvAdfCode
                        tmPrf.sName = smSave(5)
                        tmPrf.iMnfComp(0) = 0
                        tmPrf.iMnfComp(1) = 0
                        tmPrf.iMnfExcl(0) = 0
                        tmPrf.iMnfExcl(1) = 0
                        tmPrf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                        tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                        tmPrf.lAutoCode = tmPrf.lCode
                        ilRet = btrInsert(hmPrf, tmPrf, imPrfRecLen, INDEXKEY0)
                        slMsg = "mSaveRec (btrInsert:Product)"
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, CopyInv
                    On Error GoTo 0
                    'If tgSpf.sRemoteUsers = "Y" Then
                        Do
                            'tmPrfSrchKey0.lCode = tmPrf.lCode
                            'ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                            'slMsg = "mSaveRec (btrGetEqual:Product)"
                            'On Error GoTo mSaveRecErr
                            'gBtrvErrorMsg ilRet, slMsg, CopyInv
                            'On Error GoTo 0
                            tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                            tmPrf.lAutoCode = tmPrf.lCode
                            tmPrf.iSourceID = tgUrf(0).iRemoteUserID
                            gPackDate slSyncDate, tmPrf.iSyncDate(0), tmPrf.iSyncDate(1)
                            gPackTime slSyncTime, tmPrf.iSyncTime(0), tmPrf.iSyncTime(1)
                            ilRet = btrUpdate(hmPrf, tmPrf, imPrfRecLen)
                            slMsg = "mSaveRec (btrUpdate:Product)"
                        Loop While ilRet = BTRV_ERR_CONFLICT
                    'End If
                End If
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, CopyInv
                On Error GoTo 0
            End If
        End If
        'Build product/ISCI
        If (smSave(5) <> "") Or (smSave(6) <> "") Or (smSave(7) <> "") Then
            If (ilNewInv) Or (tmCif.lcpfCode = 0) Then
                tmCpf.lCode = 0
                tmCpf.sName = smSave(5)
                tmCpf.sISCI = smSave(6)
                tmCpf.sCreative = smSave(7)
                tmCpf.iRotEndDate(0) = 0
                tmCpf.iRotEndDate(1) = 0
                ilRet = btrInsert(hmCpf, tmCpf, imCpfRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert: Cpf)"
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, CopyInv
                On Error GoTo 0
                tmCif.lcpfCode = tmCpf.lCode
            Else
                Do
                    tmCpf.sName = smSave(5)
                    tmCpf.sISCI = smSave(6)
                    tmCpf.sCreative = smSave(7)
                    ilRet = btrUpdate(hmCpf, tmCpf, imCpfRecLen)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        tmCpfSrchKey.lCode = tmCif.lcpfCode
                        ilErrRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                slMsg = "mSaveRec (btrUpdate: Cpf)"
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, CopyInv
                On Error GoTo 0
            End If
        Else
            tmCif.lcpfCode = 0
        End If
    Else
        tmCif.lcpfCode = 0
    End If
    llSvCsfCode = tmCif.lCsfCode
    llSvCpfCode = tmCif.lcpfCode
    Do  'Loop until record updated or added
        If ilNewInv Then 'New selected
            If (Trim$(tmMcf.sReuse) = "N") Or (tmCif.lCode = 0) Or (smUseCartNo = "N") Then
                tmCif.lCode = 0  'Autoincrement
                If tmCif.sPurged = "P" Then
                    slDate = Format$(gNow(), "m/d/yy")
                    gPackDate slDate, tmCif.iPurgeDate(0), tmCif.iPurgeDate(1)
                ElseIf tmCif.sPurged = "H" Then
                    slDate = Format$(gNow(), "m/d/yy")
                    gPackDate slDate, tmCif.iPurgeDate(0), tmCif.iPurgeDate(1)
                Else
                    tmCif.iPurgeDate(0) = 0
                    tmCif.iPurgeDate(1) = 0
                End If
                slDate = Format$(gNow(), "m/d/yy")
                gPackDate slDate, tmCif.iDateEntrd(0), tmCif.iDateEntrd(1)
                tmCif.iUsedDate(0) = 0
                tmCif.iUsedDate(1) = 0
                tmCif.iRotStartDate(0) = 0
                tmCif.iRotStartDate(1) = 0
                tmCif.iRotEndDate(0) = 0
                tmCif.iRotEndDate(1) = 0
                tmCif.iUrfCode = tgUrf(0).iCode
                ilRet = btrInsert(hmCif, tmCif, imCifRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert: Cif)"
            Else
                If tmCif.sPurged = "P" Then
                    slDate = Format$(gNow(), "m/d/yy")
                    gPackDate slDate, tmCif.iPurgeDate(0), tmCif.iPurgeDate(1)
                ElseIf tmCif.sPurged = "H" Then
                    slDate = Format$(gNow(), "m/d/yy")
                    gPackDate slDate, tmCif.iPurgeDate(0), tmCif.iPurgeDate(1)
                Else
                    tmCif.iPurgeDate(0) = 0
                    tmCif.iPurgeDate(1) = 0
                End If
                slDate = Format$(gNow(), "m/d/yy")
                gPackDate slDate, tmCif.iDateEntrd(0), tmCif.iDateEntrd(1)
                tmCif.iUsedDate(0) = 0
                tmCif.iUsedDate(1) = 0
                tmCif.iRotStartDate(0) = 0
                tmCif.iRotStartDate(1) = 0
                tmCif.iRotEndDate(0) = 0
                tmCif.iRotEndDate(1) = 0
                tmCif.iUrfCode = tgUrf(0).iCode
                ilRet = btrUpdate(hmCif, tmCif, imCifRecLen)
                slMsg = "mSaveRec (btrInsert: Cif)"
            End If
        Else 'Old record-Update
            If (smOrigPurge = "P") And (tmCif.sPurged <> "P") Then
                'If tmCif.sPurged = "A" Then
                    'tmCif.iPurgeDate(0) = 0
                    'tmCif.iPurgeDate(1) = 0
                'End If
                'tmCif.iUsedDate(0) = 0
                'tmCif.iUsedDate(1) = 0
                'tmCif.iRotStartDate(0) = 0
                'tmCif.iRotStartDate(1) = 0
                'tmCif.iRotEndDate(0) = 0
                'tmCif.iRotEndDate(1) = 0
            End If
            If (smOrigPurge <> "P") And (tmCif.sPurged = "P") Then
                slDate = Format$(gNow(), "m/d/yy")
                gPackDate slDate, tmCif.iPurgeDate(0), tmCif.iPurgeDate(1)
            End If
            If (smOrigPurge <> "H") And (tmCif.sPurged = "H") Then
                slDate = Format$(gNow(), "m/d/yy")
                gPackDate slDate, tmCif.iPurgeDate(0), tmCif.iPurgeDate(1)
            End If
            tmCif.iUrfCode = tgUrf(0).iCode
            If smUseCartNo <> "N" Then
                If Trim$(tmMcf.sReuse) = "N" Then
                    'Need to delete in case Name changed
                    ilRet = btrDelete(hmCif)
                    ilRet = btrInsert(hmCif, tmCif, imCifRecLen, INDEXKEY0)
                Else
                    ilRet = btrUpdate(hmCif, tmCif, imCifRecLen)
                End If
            Else
                ilRet = btrUpdate(hmCif, tmCif, imCifRecLen)
            End If
            slMsg = "mSaveRec (btrUpdate: Cif)"
            If ilRet = BTRV_ERR_CONFLICT Then
                tmCifSrchKey.lCode = tmCif.lCode
                ilCifRet = btrGetEqual(hmCif, tlCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                mMoveCtrlToRec True
                tmCif.lCsfCode = llSvCsfCode
                tmCif.lcpfCode = llSvCpfCode
            End If
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, CopyInv
    On Error GoTo 0
    If ilNewInv Then
        mClearCuf
        mActiveToHistory
        If smUseCartNo <> "N" Then
            If Trim$(tmMcf.sReuse) = "N" Then
                ilAddNew = False
                If cbcInv.ListCount > 0 Then
                    If cbcInv.List(0) = "[New]" Then
                        ilAddNew = True
                        cbcInv.RemoveItem 0
                    End If
                End If
                If Trim$(tmCif.sCut) <> "" Then
                    slName = Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut)
                Else
                    slName = Trim$(tmCif.sName)
                End If
                cbcInv.AddItem slName, 0
                If ilAddNew Then
                    cbcInv.AddItem "[New]", 0
                End If
                slName = slName & "\" & LTrim$(str$(tmCif.lCode))
                'lbcInvCode.AddItem slName, 0
                ReDim Preserve tmInvNameCode(0 To UBound(tmInvNameCode) + 1) As SORTCODE
                For ilLoop = UBound(tmInvNameCode) To 1 Step -1
                    tmInvNameCode(ilLoop).sKey = tmInvNameCode(ilLoop - 1).sKey
                Next ilLoop
                tmInvNameCode(0).sKey = slName
            Else
                cbcInv.RemoveItem imCifIndex
                'lbcInvCode.RemoveItem imCifIndex
                gRemoveItemFromSortCode imCifIndex, tmInvNameCode()
                imCifIndex = -1
            End If
        Else
            ilAddNew = False
            If cbcInv.ListCount > 0 Then
                If cbcInv.List(0) = "[New]" Then
                    ilAddNew = True
                    cbcInv.RemoveItem 0
                End If
            End If
            slName = Trim$(tmCpf.sISCI)
            cbcInv.AddItem slName, 0
            If ilAddNew Then
                cbcInv.AddItem "[New]", 0
            End If
            slName = slName & "\" & LTrim$(str$(tmCif.lCode))
            'lbcInvCode.AddItem slName, 0
            ReDim Preserve tmInvNameCode(0 To UBound(tmInvNameCode) + 1) As SORTCODE
            For ilLoop = UBound(tmInvNameCode) To 1 Step -1
                tmInvNameCode(ilLoop).sKey = tmInvNameCode(ilLoop - 1).sKey
            Next ilLoop
            tmInvNameCode(0).sKey = slName
            imCifIndex = -1
        End If
    Else
        If Trim$(tmMcf.sReuse) = "N" Then  'Update inventory number of changed
        End If
    End If
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
'*      Procedure Name:mSaveRecChg                      *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
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
    Dim ilAltered As Integer
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If mTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO Then
        If ilAltered = YES Then
            If ilAsk Then
                If ((imCifIndex > 0) And (imProcMode = 0)) Or ((imCifIndex >= 0) And (imProcMode = 1)) Then
                    slMess = "Save Changes to " & cbcInv.List(imCifIndex)
                Else
                    If smUseCartNo <> "N" Then
                        If Trim$(smSave(2)) <> "" Then
                            slMess = "Add " & smSave(1) & "-" & smSave(2)
                        Else
                            slMess = "Add " & smSave(1)
                        End If
                    Else
                        slMess = "Add " & smSave(6)
                    End If
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcInv_Paint
                    Exit Function
                End If
                If ilRes = vbYes Then
                    ilRes = mSaveRec()
                    mSaveRecChg = ilRes
                    Exit Function
                End If
                If ilRes = vbNo Then
                    cbcInv.ListIndex = 0
                End If
            Else
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
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
    Dim slStr As String
    Dim slStr1 As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > imMaxNoCtrls Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case INVNOINDEX 'Name
            gSetChgFlag tmCif.sName, edcDropDown, tmCtrls(ilBoxNo)
        Case CUTINDEX   'Cut #
            gSetChgFlag tmCif.sCut, edcDropDown, tmCtrls(ilBoxNo)
        Case REELINDEX   'Reel #
            gSetChgFlag tmCif.sReel, edcDropDown, tmCtrls(ilBoxNo)
        Case LENINDEX   'Length
            gSetChgFlag Trim$(str$(tmCif.iLen)), lbcLen, tmCtrls(ilBoxNo)
        Case ANNINDEX 'Announcer
            gSetChgFlag smOrigAnn, lbcAnn, tmCtrls(ilBoxNo)
        Case COMPINDEX 'Competitive
            gSetChgFlag smOrigComp0, lbcComp(0), tmCtrls(ilBoxNo)
        Case COMPINDEX + 1 'Competitive
            gSetChgFlag smOrigComp1, lbcComp(1), tmCtrls(ilBoxNo)
        Case PRODINDEX
            gSetChgFlagStr Trim$(tmCpf.sName), smSave(5), tmCtrls(ilBoxNo)
        Case ISCIINDEX
            gSetChgFlag Trim$(tmCpf.sISCI), edcDropDown, tmCtrls(ilBoxNo)
        Case PURGEINDEX
        Case CARTINDEX
            Select Case Trim$(tmCif.sCartDisp)
                Case "N"
                    slStr = lbcCartDisp.List(0)
                Case "S"
                    slStr = lbcCartDisp.List(1)
                Case "P"
                    slStr = lbcCartDisp.List(2)
                Case "A"
                    slStr = lbcCartDisp.List(3)
                Case Else
                    slStr = ""
            End Select
            '12/26/13:N/A is allowed answer
            'gSetChgFlag slStr, lbcCartDisp, tmCtrls(ilBoxNo)
            If imSave(4) = 1 Then
                slStr1 = lbcCartDisp.List(1)
            ElseIf imSave(4) = 2 Then
                slStr1 = lbcCartDisp.List(2)
            ElseIf imSave(4) = 3 Then
                slStr1 = lbcCartDisp.List(3)
            ElseIf imSave(4) = 0 Then
                slStr1 = lbcCartDisp.List(0)
            Else
                slStr1 = ""
            End If
            'gSetChgFlagStr Trim$(slStr), Trim(slStr1), tmCtrls(ilBoxNo)
            If slStr <> slStr1 Then
                tmCtrls(ilBoxNo).iChg = True
            Else
                tmCtrls(ilBoxNo).iChg = True
            End If
        Case TAPEINDEX
            Select Case Trim$(tmCif.sTapeDisp)
                Case "N"
                    slStr = lbcTapeDisp.List(0)
                Case "R"
                    slStr = lbcTapeDisp.List(1)
                Case "D"
                    slStr = lbcTapeDisp.List(2)
                Case "A"
                    slStr = lbcTapeDisp.List(3)
                Case Else
                    slStr = ""
            End Select
            '12/26/13:N/A is allowed answer
            'gSetChgFlag slStr, lbcTapeDisp, tmCtrls(ilBoxNo)
            If imSave(5) = 1 Then
                slStr1 = lbcTapeDisp.List(1)
            ElseIf imSave(5) = 2 Then
                slStr1 = lbcTapeDisp.List(2)
            ElseIf imSave(5) = 3 Then
                slStr1 = lbcTapeDisp.List(3)
            ElseIf imSave(5) = 0 Then
                slStr1 = lbcTapeDisp.List(0)
            Else
                slStr1 = ""
            End If
            'gSetChgFlagStr Trim$(slStr), Trim(slStr1), tmCtrls(ilBoxNo)
            If slStr <> slStr1 Then
                tmCtrls(ilBoxNo).iChg = True
            Else
                tmCtrls(ilBoxNo).iChg = True
            End If
        Case NOAIRINDEX
            gSetChgFlag Trim$(str$(tmCif.iNoTimesAir)), edcDropDown, tmCtrls(ilBoxNo)
        Case TITLEINDEX
            gSetChgFlag Trim$(tmCpf.sCreative), edcDropDown, tmCtrls(ilBoxNo)
        Case TAPEININDEX
        Case TAPEAPPINDEX
        Case INVSENTDATEINDEX
            gSetChgFlag smOrigSentDate, edcDropDown, tmCtrls(ilBoxNo)
        Case SCRIPTINDEX
            gSetChgFlag smOrigComment, edcComment, tmCtrls(ilBoxNo)
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
    If imBypassSetting Then
        Exit Sub
    End If
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If smUseCartNo <> "N" Then
        If (imMcfIndex < 0) Or (imCifIndex < 0) Then
            cbcMedia.Enabled = True
            cbcInv.Enabled = True
            pbcInv.Enabled = False  'Disallow mouse
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
        Else
            pbcInv.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            If Not ilAltered Then
                If (imProcMode <> 0) And (Not imFirstFocusMedia) Then
                    cbcMedia.Enabled = False
                Else
                    cbcMedia.Enabled = True
                End If
                cbcInv.Enabled = True
                'pbcInv.Enabled = True
            Else
                cbcMedia.Enabled = False
                cbcInv.Enabled = False
                'pbcInv.Enabled = False
            End If
        End If
    Else
        If (imCifIndex < 0) Then
            cbcMedia.Enabled = True
            cbcInv.Enabled = True
            pbcInv.Enabled = False  'Disallow mouse
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
        Else
            pbcInv.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            If Not ilAltered Then
                cbcInv.Enabled = True
                'pbcInv.Enabled = True
            Else
                cbcInv.Enabled = False
                'pbcInv.Enabled = False
            End If
        End If
    End If
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) Then
        cmcUpdate.Enabled = True
        If (imProcMode = 0) And (smUseCartNo <> "N") Then
            cmcDupl.Enabled = True
        End If
    Else
        cmcUpdate.Enabled = False
        If (imProcMode = 0) And (smUseCartNo <> "N") Then
            cmcDupl.Enabled = False
        End If
    End If
    'Revert button set if any field changed
    If (ilAltered) Then
        cmcUndo.Enabled = True
        cmcImport.Enabled = False
    Else
        cmcUndo.Enabled = False
        If imProcMode = 0 Then
            If (smUseCartNo <> "N") Then
                If Trim$(tmMcf.sReuse) = "N" Then
                    cmcImport.Enabled = False
                Else
                    cmcImport.Enabled = True
                End If
            Else
                cmcImport.Enabled = True
            End If
        Else
            cmcImport.Enabled = False
        End If
    End If
    If (smUseCartNo <> "N") Then
        If imProcMode <> 0 Then
            If ilAltered Then
                cmcDupl.Enabled = False
            Else
                If cbcMedia.ListCount <= 1 Then
                    cmcDupl.Enabled = False
                Else
                    If Trim$(tmMcf.sReuse) = "N" Then
                        If imCifIndex < 1 Then
                            cmcDupl.Enabled = False
                        Else
                            cmcDupl.Enabled = True
                        End If
                    Else
                        cmcDupl.Enabled = True
                    End If
                End If
            End If
        End If
    Else
        cmcDupl.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    'If smUseCartNo <> "N" Then
    '    If Trim$(tmMcf.sReuse) = "N" Then
    '        If (imCifIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) Then
    '            cmcErase.Enabled = True
    '        Else
    '            cmcErase.Enabled = False
    '        End If
    '    Else
    '        cmcErase.Enabled = False
    '    End If
    'Else
    '    If (imCifIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) Then
    '        cmcErase.Enabled = True
    '    Else
    '        cmcErase.Enabled = False
    '    End If
    'End If
    If imProcMode = 0 Then
        rbcPurged(0).Enabled = cbcInv.Enabled
        rbcPurged(1).Enabled = cbcInv.Enabled
        rbcPurged(2).Enabled = cbcInv.Enabled
        pbcSort.Enabled = cbcInv.Enabled
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxNoCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case INVNOINDEX 'Name
            edcDropDown.SetFocus
        Case CUTINDEX   'Cut #
            edcDropDown.SetFocus
        Case REELINDEX   'Reel #
            edcDropDown.SetFocus
        Case LENINDEX   'Length
            edcDropDown.SetFocus
        Case ANNINDEX 'announcer
            edcDropDown.SetFocus
        Case COMPINDEX 'Competitive
            edcDropDown.SetFocus
        Case COMPINDEX + 1 'Competitive
            edcDropDown.SetFocus
        Case PRODINDEX
            edcDropDown.SetFocus
        Case ISCIINDEX
            edcDropDown.SetFocus
        Case PURGEINDEX
            pbcYN.SetFocus
        Case CARTINDEX
            edcDropDown.SetFocus
        Case TAPEINDEX
            edcDropDown.SetFocus
        Case NOAIRINDEX
            edcDropDown.SetFocus
        Case TITLEINDEX
            edcDropDown.SetFocus
        Case TAPEININDEX
            pbcYN.SetFocus
        Case TAPEAPPINDEX
            pbcYN.SetFocus
        Case SCRIPTINDEX
            edcComment.SetFocus
    End Select
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
    Dim ilPos As Integer
    Dim llCntrNo As Long
    
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxNoCtrls) Then
        Exit Sub
    End If

    If edcDropDown.Visible Then
        '12/9/15
        If (ilBoxNo <> CARTINDEX) And (ilBoxNo <> TAPEINDEX) Then
            '12/5/16: Check product
            'If ((ilBoxNo <> PRODINDEX) And (ilBoxNo <> ANNINDEX) And (ilBoxNo <> COMPINDEX) And (ilBoxNo <> COMPINDEX + 1)) Or (edcDropDown.Text <> "[None]") Then
            If ((ilBoxNo <> ANNINDEX) And (ilBoxNo <> COMPINDEX) And (ilBoxNo <> COMPINDEX + 1)) Or (edcDropDown.Text <> "[None]") Then
                slStr = gReplaceIllegalCharacters(edcDropDown.Text)
                edcDropDown.Text = slStr
            End If
        End If
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case INVNOINDEX 'Name
            edcDropDown.Visible = False  'Set visibility
            smSave(1) = edcDropDown.Text
            gSetShow pbcInv, smSave(1), tmCtrls(ilBoxNo)
        Case CUTINDEX   'Cut #
            edcDropDown.Visible = False  'Set visibility
            smSave(2) = edcDropDown.Text
            gSetShow pbcInv, smSave(2), tmCtrls(ilBoxNo)
        Case REELINDEX   'Reel #
            lbcCntr.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                ilPos = InStr(1, edcDropDown.Text, "-", vbTextCompare)
                If ilPos > 0 Then
                    llCntrNo = Val(Left(edcDropDown.Text, ilPos - 1))
                Else
                    ilPos = InStr(1, edcDropDown.Text, " ", vbTextCompare)
                    If ilPos > 0 Then
                        llCntrNo = Val(Left(edcDropDown.Text, ilPos - 1))
                    Else
                        llCntrNo = Val(edcDropDown.Text)
                    End If
                End If
                If lmCntrNo <> llCntrNo Then
                    lmCntrNo = llCntrNo
                    mGetInternalComment lmCntrNo
                    tmCtrls(SCRIPTINDEX).iChg = True
                End If
                smSave(3) = edcDropDown.Text    'Trim$(Str$(llCntrNo))
            Else
                smSave(3) = edcDropDown.Text
            End If
            gSetShow pbcInv, smSave(3), tmCtrls(ilBoxNo)
        Case LENINDEX   'Length
            lbcLen.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcLen.ListIndex < 0 Then
                smSave(4) = ""
            Else
                smSave(4) = lbcLen.List(lbcLen.ListIndex)
            End If
            gSetShow pbcInv, smSave(4), tmCtrls(ilBoxNo)
        Case ANNINDEX 'announcer
            lbcAnn.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            imSave(9) = lbcAnn.ListIndex
            If lbcAnn.ListIndex <= 1 Then
                slStr = ""
            Else
                slStr = lbcAnn.List(lbcAnn.ListIndex)
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case COMPINDEX 'Competitive
            lbcComp(0).Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            imSave(1) = lbcComp(0).ListIndex
            If lbcComp(0).ListIndex <= 1 Then
                slStr = ""
            Else
                slStr = lbcComp(0).List(lbcComp(0).ListIndex)
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case COMPINDEX + 1 'Competitive
            lbcComp(1).Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            imSave(2) = lbcComp(1).ListIndex
            If lbcComp(1).ListIndex <= 1 Then
                slStr = ""
            Else
                slStr = lbcComp(1).List(lbcComp(1).ListIndex)
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case PRODINDEX
            lbcProd.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            'If lbcProd.ListIndex <= 0 Then
            '    smSave(5) = ""
            'Else
            '    smSave(5) = lbcProd.List(lbcProd.ListIndex)
            'End If
            gSetShow pbcInv, smSave(5), tmCtrls(ilBoxNo)
        Case ISCIINDEX
            edcDropDown.Visible = False  'Set visibility
            smSave(6) = Trim$(edcDropDown.Text)
            gSetShow pbcInv, smSave(6), tmCtrls(ilBoxNo)
        Case PURGEINDEX
            pbcYN.Visible = False  'Set visibility
            If imSave(3) = 0 Then
                slStr = "Active"
            ElseIf imSave(3) = 1 Then
                slStr = "Purged"
            ElseIf imSave(3) = 2 Then
                slStr = "History"
            Else
                slStr = ""
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case CARTINDEX
            lbcCartDisp.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            imSave(4) = lbcCartDisp.ListIndex
            If lbcCartDisp.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcCartDisp.List(lbcCartDisp.ListIndex)
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case TAPEINDEX
            lbcCartDisp.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            imSave(5) = lbcTapeDisp.ListIndex
            If lbcTapeDisp.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcTapeDisp.List(lbcTapeDisp.ListIndex)
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case NOAIRINDEX
            edcDropDown.Visible = False  'Set visibility
            imSave(6) = Val(edcDropDown.Text)
            slStr = edcDropDown.Text
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case TITLEINDEX
            edcDropDown.Visible = False  'Set visibility
            smSave(7) = Trim$(edcDropDown.Text)
            gSetShow pbcInv, smSave(7), tmCtrls(ilBoxNo)
        Case TAPEININDEX
            pbcYN.Visible = False  'Set visibility
            If imSave(7) = 0 Then
                slStr = "No"
            ElseIf imSave(7) = 1 Then
                slStr = "Yes"
            Else
                slStr = ""
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case TAPEAPPINDEX
            pbcYN.Visible = False  'Set visibility
            If ((Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT) Then
                If imSave(8) = 0 Then
                    slStr = "Not Sent"
                ElseIf imSave(8) = 1 Then
                    slStr = "Produced"
                ElseIf (imSave(8) = 2) Then
                    slStr = "Sent"
                ElseIf (imSave(8) = 3) Then
                    slStr = "Hold"
                Else
                    slStr = ""
                End If
            Else
                If imSave(8) = 0 Then
                    slStr = "No"
                ElseIf imSave(8) = 1 Then
                    slStr = "Yes"
                Else
                    slStr = ""
                End If
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
        Case INVSENTDATEINDEX
            edcDropDown.Visible = False  'Set visibility
            smSave(8) = Trim$(edcDropDown.Text)
            gSetShow pbcInv, smSave(8), tmCtrls(ilBoxNo)
        Case SCRIPTINDEX
            edcComment.Visible = False  'Set visibility
            smCommentTextOnly = Trim$(edcComment.TextOnly)
            If smCommentTextOnly <> "" Then
                smComment = edcComment.Text
            Else
                smComment = ""
            End If
            slStr = Left$(smCommentTextOnly, 80)
            ilPos = InStr(slStr, sgLF)
            If ilPos = 2 Then
                slStr = Mid$(slStr, ilPos + 1)
            End If
            ilPos = InStr(slStr, sgCR)
            If ilPos > 0 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            gSetShow pbcInv, slStr, tmCtrls(ilBoxNo)
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
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

    On Error Resume Next

    ilRet = mFreeBlock()

    smInvNameCodeTag = ""

    sgDoneMsg = Trim$(str$(igCopyInvCallSource)) & "\" & str$(lgCopyInvCifCode)
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload CopyInv
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:4/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
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
    Dim slStr As String
    If smUseCartNo <> "N" Then
        If (ilCtrlNo = INVNOINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedStr(smSave(1), "", "Inventory # must be specified", tmCtrls(INVNOINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = INVNOINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = CUTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedStr(smSave(2), "", "Cut # must be specified", tmCtrls(CUTINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = CUTINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = REELINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
            If gFieldDefinedStr(smSave(3), "", "Contract # must be specified", tmCtrls(REELINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = REELINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        Else
            If gFieldDefinedStr(smSave(3), "", "Reel # must be specified", tmCtrls(REELINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = REELINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = LENINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcLen, "", "Length must be specified", tmCtrls(LENINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = LENINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
        If ilCtrlNo = TESTALLCTRLS Then
            If (Val(smOrigLen) <> 0) And (Val(smSave(4)) <> Val(smOrigLen)) And (ilCtrlNo = TESTALLCTRLS) Then
                If smOrigRotDate <> ": Unused" Then
                    If (ilState And SHOWMSG) = SHOWMSG Then 'Only when saving
                        ilRes = MsgBox("Changing Spot Length will cause this inventory in previously defined rotations to be bypassed", vbOKCancel + vbExclamation, "Warning")
                        If ilRes = vbCancel Then
                            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                                imBoxNo = LENINDEX
                            End If
                            mTestFields = NO
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If
    If (ilCtrlNo = ANNINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcAnn, "", "Announcer must be specified", tmCtrls(ANNINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ANNINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = COMPINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcComp(0), "", "Competitive must be specified", tmCtrls(COMPINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMPINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = COMPINDEX + 1) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcComp(1), "", "Competitive must be specified", tmCtrls(COMPINDEX + 1).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMPINDEX + 1
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PRODINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSave(5), "", "Product must be specified", tmCtrls(PRODINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PRODINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = ISCIINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSave(6), "", "ISCI must be specified", tmCtrls(ISCIINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ISCIINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
        If (ilCtrlNo = TESTALLCTRLS) And ((ilState And SHOWMSG) = SHOWMSG) Then
            If Not mISCIOk(smSave(6)) Then
                If (ilState And SHOWMSG) = SHOWMSG Then
                    ilRes = MsgBox("ISCI must be a unique #", vbOKOnly + vbExclamation, "Incomplete")
                End If
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = ISCIINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = PURGEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        Select Case imSave(3)
            Case 0
                slStr = "Active"
            Case 1
                slStr = "Purged"
            Case 2
                slStr = "History"
            Case Else
                slStr = ""
        End Select
        If gFieldDefinedStr(slStr, "", "Status must be specified", tmCtrls(PURGEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PURGEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = CARTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcCartDisp, "", "Cart Disposition must be specified", tmCtrls(CARTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = CARTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TAPEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcTapeDisp, "", "Tape Disposition must be specified", tmCtrls(TAPEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TAPEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = NOAIRINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        slStr = Trim$(str$(imSave(6)))
        If gFieldDefinedStr(slStr, "", "# Times Aired must be specified", tmCtrls(NOAIRINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NOAIRINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TITLEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSave(7), "", "Creative Title must be specified", tmCtrls(TITLEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TITLEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TAPEININDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        Select Case imSave(7)
            Case 0
                slStr = "No"
            Case 1
                slStr = "Yes"
            Case Else
                slStr = ""
        End Select
        If gFieldDefinedStr(slStr, "", "Tape in House must be specified", tmCtrls(TAPEININDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TAPEININDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TAPEAPPINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
            Select Case imSave(8)
                Case 0
                    slStr = "Not Sent"
                Case 1
                    slStr = "Produced"
                Case 2
                    slStr = "Sent"
                Case 3
                    slStr = "Hold"
                Case Else
                    slStr = ""
            End Select
            If gFieldDefinedStr(slStr, "", "Copy Produced must be specified", tmCtrls(TAPEAPPINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = TAPEAPPINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        Else
            Select Case imSave(8)
                Case 0
                    slStr = "No"
                Case 1
                    slStr = "Yes"
                Case Else
                    slStr = ""
            End Select
            If tgSpf.sTapeShowForm = "C" Then  'Tape Carted or Approved
                If gFieldDefinedStr(slStr, "", "Tape Carted must be specified", tmCtrls(TAPEAPPINDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = TAPEAPPINDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            Else
                If gFieldDefinedStr(slStr, "", "Tape Approved must be specified", tmCtrls(TAPEAPPINDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = TAPEAPPINDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            End If
        End If
    End If
    If (ilCtrlNo = INVSENTDATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSave(7), "", "Action Date must be specified", tmCtrls(TITLEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = INVSENTDATEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SCRIPTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smComment, "", "Script/Comment must be specified", tmCtrls(SCRIPTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SCRIPTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    mTestFields = YES
End Function
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcDupl_Paint()
    Dim ilBox As Integer

    'mPaintTitle 1
    mPaintCopyInvTitle pbcDupl
    For ilBox = imLBCtrls To UBound(tmDuplCtrls) Step 1
        pbcDupl.CurrentX = tmDuplCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcDupl.CurrentY = tmDuplCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcDupl.Print tmDuplCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    'Handle double names- if drop down selected the index is changed to the
    'first name without any events- forces back so change occurs
    If cbcInv.ListIndex <> imCifIndex Then
        cbcInv_Change
        cbcInv.SetFocus
        Exit Sub
    End If
    For ilBox = imLBCtrls To imMaxNoCtrls Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                If smUseCartNo <> "N" Then
                    If (ilBox = INVNOINDEX) And (Trim$(tmMcf.sReuse) <> "N") Then
                        mSetFocus imBoxNo
                        Beep
                        Exit Sub
                    End If
                    If (ilBox = CUTINDEX) And (Trim$(tmMcf.sReuse) <> "N") Then
                        mSetFocus imBoxNo
                        Beep
                        Exit Sub
                    End If
                Else
                    If (ilBox = INVNOINDEX) Then
                        mSetFocus imBoxNo
                        Beep
                        Exit Sub
                    End If
                    If (ilBox = CUTINDEX) Then
                        mSetFocus imBoxNo
                        Beep
                        Exit Sub
                    End If
                End If
                If (ilBox = INVSENTDATEINDEX) And (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) <> VCREATIVEEXPORT Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = SCRIPTINDEX) And (tmMcf.sScript <> "Y") Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If (ilBox = PURGEINDEX) And (imBypassPurge) Then
                    mSetFocus imBoxNo
                    Beep
                    Exit Sub
                End If
                If igCopyInvAdfCode <= 0 Then
                    If (ilBox = PRODINDEX) Or (ilBox = ISCIINDEX) Or (ilBox = TITLEINDEX) Then
                        mSetFocus imBoxNo
                        Beep
                        Exit Sub
                    End If
                End If
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcInv_Paint()
    Dim ilBox As Integer
    'mPaintTitle 0
    mPaintCopyInvTitle pbcInv
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        pbcInv.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcInv.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcInv.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub

Private Sub pbcSort_KeyPress(KeyAscii As Integer)
    Dim ilChg As Integer
    ilChg = False
    If KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If imSortCart <> 0 Then
            'tmCtrls(imBoxNo).iChg = True
            ilChg = True
        End If
        imSortCart = 0
        pbcSort_Paint
    ElseIf KeyAscii = Asc("C") Or (KeyAscii = Asc("c")) Then
        If imSortCart <> 1 Then
            'tmCtrls(imBoxNo).iChg = True
            ilChg = True
        End If
        imSortCart = 1
        pbcSort_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imSortCart = 0 Then
            'tmCtrls(imBoxNo).iChg = True
            ilChg = True
            imSortCart = 1
            pbcSort_Paint
        ElseIf imSortCart = 1 Then
            'tmCtrls(imBoxNo).iChg = True
            ilChg = True
            imSortCart = 0
            pbcSort_Paint
        End If
    End If
    If ilChg Then
        If rbcPurged(1).Value Then
            rbcPurged_Click 1
        ElseIf rbcPurged(2).Value Then
            rbcPurged_Click 2
        Else
            rbcPurged_Click 0
        End If
    End If
End Sub

Private Sub pbcSort_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imSortCart = 0 Then
        'tmCtrls(imBoxNo).iChg = True
        imSortCart = 1
    ElseIf imSortCart = 1 Then
        'tmCtrls(imBoxNo).iChg = True
        imSortCart = 0
    Else
        imSortCart = 0
    End If
    pbcSort_Paint
    If rbcPurged(1).Value Then
        rbcPurged_Click 1
    ElseIf rbcPurged(2).Value Then
        rbcPurged_Click 2
    Else
        rbcPurged_Click 0
    End If
End Sub

Private Sub pbcSort_Paint()
    pbcSort.Cls
    pbcSort.CurrentX = fgBoxInsetX
    pbcSort.CurrentY = 0 'fgBoxInsetY
    If imSortCart = 0 Then
        pbcSort.Print "Date"
    ElseIf imSortCart = 1 Then
        pbcSort.Print "Cart"
    Else
        pbcSort.Print "    "
    End If
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    'Handle double names- if drop down selected the index is changed to the
    'first name without any events- forces back so change occurs
    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    If smUseCartNo <> "N" Then
        If cbcInv.ListIndex <> imCifIndex Then
            cbcInv_Change
            cbcInv.SetFocus
            Exit Sub
        End If
    End If
    If imBoxNo = ANNINDEX Then
        If mAnnBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = PRODINDEX Then
        If mProdBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = COMPINDEX Then
        If mCompBranch(0) Then
            Exit Sub
        End If
    End If
    If imBoxNo = COMPINDEX + 1 Then
        If mCompBranch(1) Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxNoCtrls) Then
        If (imBoxNo <> INVNOINDEX) Or (Not cbcInv.Enabled) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    imTabDirection = -1  'Set-right to left
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imTabDirection = 0  'Set-Left to right
                If smUseCartNo <> "N" Then
                    If Trim$(tmMcf.sReuse) = "N" Then
                        If imCifIndex = 0 Then  'New
                            ilBox = INVNOINDEX
                            mSetCommands
                        Else
                            mSetChg 1
                            ilBox = CUTINDEX
                        End If
                    Else
                        mSetChg 1
                        ilBox = REELINDEX
                    End If
                Else
                    ilBox = REELINDEX
                End If
            Case INVNOINDEX 'Name (first control within header)
                mSetShow imBoxNo
                imBoxNo = -1
                If cbcInv.Enabled Then
                    cbcInv.SetFocus
                    Exit Sub
                End If
                cmcDone.SetFocus
                Exit Sub
            Case REELINDEX
                If smUseCartNo <> "N" Then
                    If Trim$(tmMcf.sReuse) = "N" Then
                        ilBox = ilBox - 1
                    Else
                        mSetShow imBoxNo
                        imBoxNo = -1
                        If cbcInv.Enabled Then
                            cbcInv.SetFocus
                            Exit Sub
                        End If
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                Else
                    mSetShow imBoxNo
                    imBoxNo = -1
                    If cbcInv.Enabled Then
                        cbcInv.SetFocus
                        Exit Sub
                    End If
                    cmcDone.SetFocus
                    Exit Sub
                End If
            Case COMPINDEX
                If lbcAnn.ListCount <= 2 Then
                    imChgMode = True
                    lbcAnn.ListIndex = 1
                    imChgMode = False
                    ilFound = False
                End If
                ilBox = ANNINDEX
            Case PRODINDEX
                If igCopyInvAdfCode <= 0 Then
                    ilFound = False
                End If
                ilBox = ilBox - 1
            Case ISCIINDEX
                If igCopyInvAdfCode <= 0 Then
                    ilFound = False
                End If
                ilBox = ilBox - 1
            Case PURGEINDEX
                If igCopyInvAdfCode <= 0 Then
                    ilFound = False
                End If
                ilBox = ilBox - 1
            Case CARTINDEX
                If imBypassPurge Then
                    ilFound = False
                End If
                ilBox = ilBox - 1
            Case TITLEINDEX
                If igCopyInvAdfCode <= 0 Then
                    ilFound = False
                End If
                ilBox = ilBox - 1
            Case TAPEININDEX
                If igCopyInvAdfCode <= 0 Then
                    ilFound = False
                End If
                ilBox = ilBox - 1
            Case SCRIPTINDEX
                If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) <> VCREATIVEEXPORT Then
                    ilFound = False
                End If
                ilBox = ilBox - 1
            Case Else
                ilBox = ilBox - 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer

    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    If imBoxNo = ANNINDEX Then
        If mAnnBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = PRODINDEX Then
        If mProdBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = COMPINDEX Then
        If mCompBranch(0) Then
            Exit Sub
        End If
    End If
    If imBoxNo = COMPINDEX + 1 Then
        If mCompBranch(1) Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxNoCtrls) Then
        If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Shift tab from button
                imTabDirection = -1  'Set-Right to left
                If tmMcf.sScript = "Y" Then
                    ilBox = SCRIPTINDEX
                Else
                    imBoxNo = TAPEAPPINDEX
                End If
            Case LENINDEX
                If lbcAnn.ListCount <= 2 Then
                    imChgMode = True
                    lbcAnn.ListIndex = 1
                    imChgMode = False
                    ilFound = False
                End If
                ilBox = ANNINDEX
            Case COMPINDEX + 1
                If igCopyInvAdfCode <= 0 Then
                    ilFound = False
                End If
                ilBox = ilBox + 1
            Case PRODINDEX
                If igCopyInvAdfCode <= 0 Then
                    ilFound = False
                End If
                ilBox = ilBox + 1
            Case ISCIINDEX
                If igCopyInvAdfCode <= 0 Then
                    ilFound = False
                End If
                If imBypassPurge Then
                    ilFound = False
                End If
                ilBox = ilBox + 1
            Case NOAIRINDEX
                If igCopyInvAdfCode <= 0 Then
                    ilFound = False
                End If
                ilBox = ilBox + 1
            Case TITLEINDEX
                If igCopyInvAdfCode <= 0 Then
                    ilFound = False
                End If
                ilBox = ilBox + 1
            Case TAPEAPPINDEX
                If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) <> VCREATIVEEXPORT Then
                    ilFound = False
                End If
                ilBox = ilBox + 1
            Case INVSENTDATEINDEX
                If tmMcf.sScript = "Y" Then
                    ilBox = ilBox + 1
                Else
                    mSetShow imBoxNo
                    imBoxNo = -1
                    If (imProcMode = 0) And (cmcUpdate.Enabled) Then    'New mode
                    'If (lgCopyInvCifCode = 0) And (cmcUpdate.Enabled) Then    'New mode
                        cmcUpdate.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                End If
            Case SCRIPTINDEX
                mSetShow imBoxNo
                imBoxNo = -1
                'If (lgCopyInvCifCode = 0) And (cmcUpdate.Enabled) Then    'New mode
                If (imProcMode = 0) And (cmcUpdate.Enabled) Then    'New mode
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
            Case Else
                ilBox = ilBox + 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcYN_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    If imBoxNo = PURGEINDEX Then
        If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
            If imSave(3) <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(3) = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("P") Or (KeyAscii = Asc("p")) Then
            If imSave(3) <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(3) = 1
            pbcYN_Paint
        ElseIf KeyAscii = Asc("H") Or (KeyAscii = Asc("h")) Then
            If imSave(3) <> 2 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(3) = 2
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSave(3) = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imSave(3) = 1
                pbcYN_Paint
            ElseIf imSave(3) = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imSave(3) = 2
                pbcYN_Paint
            ElseIf imSave(3) = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imSave(3) = 0
                pbcYN_Paint
            End If
        End If
    ElseIf imBoxNo = TAPEININDEX Then
        If KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imSave(7) <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(7) = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imSave(7) <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(7) = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSave(7) = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imSave(7) = 1
                pbcYN_Paint
            ElseIf imSave(7) = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imSave(7) = 0
                pbcYN_Paint
            End If
        End If
    ElseIf imBoxNo = TAPEAPPINDEX Then
        If ((Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT) Then
            If (KeyAscii = Asc("N")) Or (KeyAscii = Asc("n")) Then
                If imSave(8) <> 0 Then
                    tmCtrls(imBoxNo).iChg = True
                End If
                imSave(8) = 0
                pbcYN_Paint
            ElseIf (KeyAscii = Asc("P")) Or (KeyAscii = Asc("p")) Then
                If imSave(8) <> 1 Then
                    tmCtrls(imBoxNo).iChg = True
                End If
                imSave(8) = 1
                pbcYN_Paint
            ElseIf (KeyAscii = Asc("S")) Or (KeyAscii = Asc("s")) Then
                If imSave(8) <> 2 Then
                    tmCtrls(imBoxNo).iChg = True
                End If
                imSave(8) = 2
                pbcYN_Paint
            ElseIf (KeyAscii = Asc("H")) Or (KeyAscii = Asc("h")) Then
                If imSave(8) <> 3 Then
                    tmCtrls(imBoxNo).iChg = True
                End If
                imSave(8) = 3
                pbcYN_Paint
            End If
            If KeyAscii = Asc(" ") Then
                If imSave(8) = 0 Then
                    tmCtrls(imBoxNo).iChg = True
                    imSave(8) = 1
                    pbcYN_Paint
                ElseIf imSave(8) = 1 Then
                    tmCtrls(imBoxNo).iChg = True
                    imSave(8) = 2
                    pbcYN_Paint
                ElseIf imSave(8) = 2 Then
                    tmCtrls(imBoxNo).iChg = True
                    imSave(8) = 3
                    pbcYN_Paint
                Else
                    tmCtrls(imBoxNo).iChg = True
                    imSave(8) = 0
                    pbcYN_Paint
                End If
            End If
        Else
            If KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
                If imSave(8) <> 0 Then
                    tmCtrls(imBoxNo).iChg = True
                End If
                imSave(8) = 0
                pbcYN_Paint
            ElseIf KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
                If imSave(8) <> 1 Then
                    tmCtrls(imBoxNo).iChg = True
                End If
                imSave(8) = 1
                pbcYN_Paint
            End If
            If KeyAscii = Asc(" ") Then
                If imSave(8) = 0 Then
                    tmCtrls(imBoxNo).iChg = True
                    imSave(8) = 1
                    pbcYN_Paint
                ElseIf imSave(8) = 1 Then
                    tmCtrls(imBoxNo).iChg = True
                    imSave(8) = 0
                    pbcYN_Paint
                End If
            End If
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = PURGEINDEX Then
        If imSave(3) = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(3) = 1
            pbcYN_Paint
        ElseIf imSave(3) = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(3) = 2
            pbcYN_Paint
        ElseIf imSave(3) = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(3) = 0
            pbcYN_Paint
        End If
    ElseIf imBoxNo = TAPEININDEX Then
        If imSave(7) = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(7) = 1
            pbcYN_Paint
        ElseIf imSave(7) = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(7) = 0
            pbcYN_Paint
        End If
    ElseIf imBoxNo = TAPEAPPINDEX Then
        If ((Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT) Then
            If imSave(8) = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imSave(8) = 1
                pbcYN_Paint
            ElseIf imSave(8) = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imSave(8) = 2
                pbcYN_Paint
            ElseIf imSave(8) = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imSave(8) = 3
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imSave(8) = 0
                pbcYN_Paint
            End If
        Else
            If imSave(8) = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imSave(8) = 1
                pbcYN_Paint
            ElseIf imSave(8) = 1 Then
                imSave(8) = 0
                pbcYN_Paint
            End If
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    If imBoxNo = PURGEINDEX Then
        If imSave(3) = 0 Then
            pbcYN.Print "Active"
        ElseIf imSave(3) = 1 Then
            pbcYN.Print "Purged"
        ElseIf imSave(3) = 2 Then
            pbcYN.Print "History"
        Else
            pbcYN.Print "   "
        End If
    ElseIf imBoxNo = TAPEININDEX Then
        If imSave(7) = 0 Then
            pbcYN.Print "No"
        ElseIf imSave(7) = 1 Then
            pbcYN.Print "Yes"
        Else
            pbcYN.Print "   "
        End If
    ElseIf imBoxNo = TAPEAPPINDEX Then
        If ((Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT) Then
            If imSave(8) = 0 Then
                pbcYN.Print "Not Sent"
            ElseIf imSave(8) = 1 Then
                pbcYN.Print "Produced"
            ElseIf (imSave(8) = 2) Then
                pbcYN.Print "Sent"
            ElseIf (imSave(8) = 3) Then
                pbcYN.Print "Hold"
            Else
                pbcYN.Print "   "
            End If
        Else
            If imSave(8) = 0 Then
                pbcYN.Print "No"
            ElseIf imSave(8) = 1 Then
                pbcYN.Print "Yes"
            Else
                pbcYN.Print "   "
            End If
        End If
    End If
End Sub
Private Sub plcCopyInv_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcCover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSetFocus imBoxNo
    Beep
    Exit Sub
End Sub
Private Sub plcDupl_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcInv_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcPurged_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcPurged(Index).Value
    'End of coded added
    If imIgnorePurgeSetting = True Then
        Exit Sub
    End If
    If Value Then
        Screen.MousePointer = vbHourglass
        pbcInv.Cls
        lbcDuplInvCode.Clear
        cbcDuplInv.Clear
        pbcDupl.Cls
        mClearCtrlFields
        mInvPop     'cbcInv_Change will call mDuplInvPop
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub rbcPurged_GotFocus(Index As Integer)
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If imCbcDropDown Then
        imCbcDropDown = False
        If cbcInv.ListIndex <> imCifIndex Then
            cbcInv_Change
            'cbcInv.SetFocus
            Exit Sub
        End If
        Exit Sub
    End If
    Select Case imBoxNo
        Case PRODINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcProd, edcDropDown, imChgMode, imLbcArrowSetting
        Case ANNINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcAnn, edcDropDown, imChgMode, imLbcArrowSetting
        Case COMPINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcComp(0), edcDropDown, imChgMode, imLbcArrowSetting
        Case COMPINDEX + 1
            imLbcArrowSetting = False
            gProcessLbcClick lbcComp(1), edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub

Private Sub mClearCuf()
    Dim ilRet As Integer
    
    tmCufSrchKey1.lCifCode = tmCif.lCode
    ilRet = btrGetEqual(hmCuf, tmCuf, imCufRecLen, tmCufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmCuf.lCifCode = tmCif.lCode)
        ilRet = btrDelete(hmCuf)
        tmCufSrchKey1.lCifCode = tmCif.lCode
        ilRet = btrGetEqual(hmCuf, tmCuf, imCufRecLen, tmCufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
End Sub

Private Sub mGetInternalComment(llCntrNo As Long)
    Dim ilRet As Integer
    Dim slStr As String
    
    slStr = ""
    imCHFRecLen = Len(tmChf)
    tmChfSrchKey1.lCntrNo = llCntrNo
    tmChfSrchKey1.iCntRevNo = 32000
    tmChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llCntrNo) And (tmChf.sSchStatus <> "F")
        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llCntrNo) And (tmChf.sSchStatus = "F") Then
        If tmChf.lCxfInt > 0 Then
            imCxfRecLen = Len(tmCxf)
            tmCxfSrchKey.lCode = tmChf.lCxfInt
            ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If ilRet = BTRV_ERR_NONE Then
                If tmCxf.sShSpot = "Y" Then
                    slStr = gStripChr0(tmCxf.sComment)
                End If
            End If
        End If
    End If
    smComment = slStr   'Trim$(Left$(tmCsf.sComment, tmCsf.iStrLen))
    edcComment.MaxLength = 5000
    edcComment.SetText (smComment)
    smCommentTextOnly = edcComment.TextOnly
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mCntrPop                        *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mCntrPop()
'
'   mCntrPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim ilAdfCode As Integer
    Dim slNameCode As String  'Name and code
    Dim slCode As String    'Code number
    Dim ilCurrent As Integer
    Dim ilVehCode As Integer
    Dim llTodayDate As Long
    Dim ilAAS As Integer
    Dim ilShow As Integer
    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilState As Integer
    
    ilAdfCode = igCopyInvAdfCode

    Screen.MousePointer = vbHourglass  'Wait
    llTodayDate = gDateValue(gNow())
    ilCurrent = 0   'Current (1=All)
    ilVehCode = -1
    ilAAS = 0
    slCntrStatus = "HO"
    slCntrType = ""
    ilShow = 2
    ilState = 1
    ilRet = gPopCntrForAASWithRotBox(CopyInv, ilAAS, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, Copy!lbcRot, lbcCntr, tmCopyCntrCode(), smCopyCntrCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mCntrPopErr
        gCPErrorMsg ilRet, "mCntrPop (gPopCntrForAASBox)", Copy
        On Error GoTo 0
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
mCntrPopErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Function mBlockInventory() As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slUserName As String
    
    ilRet = mFreeBlock()
    
    If cbcInv.List(0) = "[New]" Then
        If imCifIndex <= 0 Then
            mBlockInventory = True
            Exit Function
        End If
        slNameCode = tmInvNameCode(imCifIndex - 1).sKey   'lbcInvCode.List(imCifIndex - 1)
    Else
        If imCifIndex < 0 Then
            mBlockInventory = True
            Exit Function
        End If
        slNameCode = tmInvNameCode(imCifIndex).sKey   'lbcInvCode.List(imCifIndex - 1)
    End If
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    lmLock1RecCode = gCreateLockRec(hmRlf, "M", "", Val(slCode), False, slUserName)
    If lmLock1RecCode = 0 Then
        mBlockInventory = False
        ilRet = MsgBox("Copy Inventory item currently being modified by " & slUserName & ", select a different inventory item", vbOKOnly + vbInformation, "Block")
        Exit Function
    End If
    mBlockInventory = True
    Exit Function
End Function

Private Function mFreeBlock() As Integer
    Dim ilRet As Integer
    
    If lmLock1RecCode <= 0 Then
        mFreeBlock = True
        Exit Function
    End If
    ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLock1RecCode)
    If ilRet = False Then
        mFreeBlock = False
    Else
        lmLock1RecCode = -1
        mFreeBlock = True
    End If
    Exit Function
End Function

Private Sub mPaintCopyInvTitle(pbcCopy As PictureBox)
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer

    llColor = pbcCopy.ForeColor
    slFontName = pbcCopy.FontName
    flFontSize = pbcCopy.FontSize
    pbcCopy.ForeColor = BLUE
    pbcCopy.FontBold = False
    pbcCopy.FontSize = 7
    pbcCopy.FontName = "Arial"
    pbcCopy.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    ''For ilLoop = LBound(tmCtrls) To UBound(tmCtrls) Step 1
    'For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
    For ilLoop = INVNOINDEX To PDATEINDEX Step 1
        If ilLoop >= ENTEREDINDEX Then
            pbcCopy.Line (tmCtrls(ilLoop).fBoxX + 15, tmCtrls(ilLoop).fBoxY)-Step(tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW - 15, tmCtrls(ilLoop).fBoxH - 45), LIGHTERYELLOW, BF
        ElseIf pbcCopy Is pbcDupl Then
            pbcCopy.Line (tmCtrls(ilLoop).fBoxX + 15, tmCtrls(ilLoop).fBoxY)-Step(tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW - 15, tmCtrls(ilLoop).fBoxH - 45), LIGHTERYELLOW, BF
        End If
        If ilLoop = COMPINDEX Then
            pbcCopy.Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + tmCtrls(ilLoop + 1).fBoxW + 30, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
        ElseIf ilLoop = COMPINDEX + 1 Then
        Else
            pbcCopy.Line (tmCtrls(ilLoop).fBoxX - 15, tmCtrls(ilLoop).fBoxY - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
        End If
        pbcCopy.CurrentX = tmCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
        pbcCopy.CurrentY = tmCtrls(ilLoop).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        Select Case ilLoop
            Case INVNOINDEX
                If smUseCartNo <> "N" Then
                    pbcCopy.Print "Inventory #"
                End If
            Case CUTINDEX
                If smUseCartNo <> "N" Then
                    pbcCopy.Print "Cut Character"
                End If
            Case REELINDEX
                If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                    pbcCopy.Print "Contract #"
                Else
                    pbcCopy.Print "Reel #"
                End If
            Case LENINDEX
                pbcCopy.Print "Length"
            Case ANNINDEX
                pbcCopy.Print "Announcer"
            Case COMPINDEX
                pbcCopy.Print "Competitives: 1st/2nd"
            Case PRODINDEX
                If tgSpf.sUseProdSptScr = "P" Then  'Short Title
                    pbcCopy.Print "Short Title"
                Else
                    pbcCopy.Print "Product"
                End If
            Case ISCIINDEX
                pbcCopy.Print "ISCI Code"
            Case PURGEINDEX
                pbcCopy.Print "Status"
            Case CARTINDEX
                pbcCopy.Print "Cart Disposition"
            Case TAPEINDEX
                pbcCopy.Print "Tape Disposition"
            Case NOAIRINDEX
                pbcCopy.Print "# Times Aired"
            Case TITLEINDEX
                pbcCopy.Print "Creative Title"
            Case TAPEININDEX
                pbcCopy.Print "Tape in House"
            Case TAPEAPPINDEX
                If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                    pbcCopy.Print "Copy Produced"
                Else
                    If tgSpf.sTapeShowForm = "C" Then  'Tape Carted or Approved
                        pbcCopy.Print "Tape Carted"
                    Else
                        pbcCopy.Print "Tape Approved"
                    End If
                End If
            'D.S. 07/16/19 added case statement below for vCreative
            Case INVSENTDATEINDEX
                If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                    pbcCopy.Print "Action Date"
                Else
                    pbcCopy.Print ""
                End If
            Case SCRIPTINDEX
                pbcCopy.Print "Script/Comment"
            Case ENTEREDINDEX
                pbcCopy.Print "Date Entered"
            Case LASTUSEDINDEX
                pbcCopy.Print "Last Date Used"
            Case EARLROTINDEX
                pbcCopy.Print "Earliest Rotation Start Date"
            Case LATROTINDEX
                pbcCopy.Print "Latest Rotation End Date"
            Case PDATEINDEX
                pbcCopy.Print "Purged/History Date"
        End Select
    Next ilLoop
    
    
    pbcCopy.FontSize = flFontSize
    pbcCopy.FontName = slFontName
    pbcCopy.FontSize = flFontSize
    pbcCopy.ForeColor = llColor
    pbcCopy.FontBold = True


End Sub

