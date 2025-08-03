VERSION 5.00
Begin VB.Form StdPkg 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5700
   ClientLeft      =   240
   ClientTop       =   2985
   ClientWidth     =   13740
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5700
   ScaleWidth      =   13740
   Begin VB.PictureBox plcACT1Settings 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   990
      Left            =   8280
      Picture         =   "Stdpkg.frx":0000
      ScaleHeight     =   930
      ScaleWidth      =   1950
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   2010
      Begin VB.TextBox edcACT1SettingT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "No"
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox edcACT1SettingS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "No"
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox edcACT1SettingC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "No"
         Top             =   620
         Width           =   705
      End
      Begin VB.TextBox edcACT1SettingF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "No"
         Top             =   620
         Width           =   705
      End
   End
   Begin VB.ListBox lbcPDFName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6180
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1455
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.FileListBox lbcPDFFile 
      Height          =   285
      Left            =   8205
      Pattern         =   "*.PDF"
      TabIndex        =   10
      Top             =   5325
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.PictureBox pbcProgrammatic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3375
      ScaleHeight     =   210
      ScaleWidth      =   1020
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmcProof 
      Appearance      =   0  'Flat
      Caption         =   "&Proof"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   1
      Left            =   6540
      TabIndex        =   31
      Top             =   5325
      Width           =   945
   End
   Begin VB.Timer tmcInit 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1830
      Top             =   5250
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   6075
      TabIndex        =   1
      Top             =   75
      Width           =   4200
   End
   Begin VB.ListBox lbcVehGp3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5490
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1005
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox pbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   255
      Picture         =   "Stdpkg.frx":5F32
      ScaleHeight     =   1065
      ScaleWidth      =   2265
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox pbcStartNew 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   9270
      ScaleHeight     =   120
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   330
      Width           =   45
   End
   Begin VB.ListBox lbcDemo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7110
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   1
      Left            =   5340
      TabIndex        =   30
      Top             =   5325
      Width           =   945
   End
   Begin VB.PictureBox pbcInvTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4575
      ScaleHeight     =   210
      ScaleWidth      =   1170
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1080
      Top             =   5235
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
      Left            =   1575
      Picture         =   "Stdpkg.frx":106F8
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcSpecDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2205
      MaxLength       =   20
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox pbcAlter 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4395
      ScaleHeight     =   210
      ScaleWidth      =   1020
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   300
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pbcPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3465
      ScaleHeight     =   210
      ScaleWidth      =   1035
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox pbcPkgTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   60
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   26
      Top             =   5280
      Width           =   135
   End
   Begin VB.PictureBox pbcPkgSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   -45
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   20
      Top             =   285
      Width           =   105
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   2
      Left            =   4125
      TabIndex        =   29
      Top             =   5325
      Width           =   945
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   1
      Left            =   2895
      TabIndex        =   28
      Top             =   5325
      Width           =   945
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   285
      MaxLength       =   10
      TabIndex        =   23
      Top             =   2580
      Visible         =   0   'False
      Width           =   1290
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
      Left            =   1605
      Picture         =   "Stdpkg.frx":107F2
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2580
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   60
      Picture         =   "Stdpkg.frx":108EC
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   390
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5385
      Width           =   75
   End
   Begin VB.PictureBox pbcSpecTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   135
      TabIndex        =   19
      Top             =   1020
      Width           =   135
   End
   Begin VB.PictureBox pbcSpecSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   60
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   3
      Top             =   1050
      Width           =   105
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   -15
      ScaleHeight     =   240
      ScaleWidth      =   1650
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1650
   End
   Begin VB.PictureBox pbcSpec 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   375
      Left            =   240
      Picture         =   "Stdpkg.frx":10BF6
      ScaleHeight     =   375
      ScaleWidth      =   13320
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   13320
   End
   Begin VB.PictureBox plcSpec 
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   195
      ScaleHeight     =   435
      ScaleWidth      =   13260
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   13325
   End
   Begin VB.PictureBox pbcPkg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   3975
      Left            =   270
      Picture         =   "Stdpkg.frx":20478
      ScaleHeight     =   3975
      ScaleWidth      =   8370
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1170
      Width           =   8370
      Begin VB.Label lacCover 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0FFFF&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3825
         TabIndex        =   37
         Top             =   3750
         Visible         =   0   'False
         Width           =   4560
      End
      Begin VB.Label lacPkgFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   36
         Top             =   750
         Visible         =   0   'False
         Width           =   8385
      End
   End
   Begin VB.PictureBox plcPkg 
      ForeColor       =   &H00000000&
      Height          =   4095
      Left            =   225
      ScaleHeight     =   4035
      ScaleWidth      =   8670
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1110
      Width           =   8730
      Begin VB.VScrollBar vbcPkg 
         Height          =   3975
         LargeChange     =   17
         Left            =   8400
         Min             =   1
         TabIndex        =   27
         Top             =   30
         Value           =   1
         Width           =   240
      End
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4905
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2955
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5415
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2925
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5295
      TabIndex        =   33
      Top             =   2850
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   300
      Picture         =   "Stdpkg.frx":448BA
      Top             =   285
      Width           =   480
   End
End
Attribute VB_Name = "StdPkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Stdpkg.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: StdPkg.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Standard Package input screen code
'
Option Explicit
Option Compare Text
Dim hmDrf As Integer
Dim hmMnf As Integer
'Vehicle
Dim imAddPkg As Integer     'Force vef and vpf to be re-read
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
Dim tmPkgVehicle() As SORTCODE
Dim smPkgVehicleTag As String
Dim smOrigVehGp3 As String
Dim smOrigDemo As String
'Vehicle Options
Dim hmVpf As Integer        'Vehicle file handle
Dim tmVpf As VPF            'VPF record image
Dim tmVpfSrchKey As VPFKEY0 'VPF key record image
Dim imVpfRecLen As Integer  'VPF record length
'Vehicle Features
Dim hmVff As Integer        'Vehicle file handle
Dim tmVff As VFF            'VFF record image
Dim tmVffSrchKey As INTKEY0 'VFF key record image
Dim tmVffSrchKey1 As INTKEY0
Dim imVffRecLen As Integer  'VFF record length
'Standard Package Vechcle
Dim hmPvf As Integer        'Standard Vehicle file handle
Dim tmPvf() As PVF            'PVF record image
Dim tmTPvf As PVF
Dim tmPvfSrchKey As LONGKEY0 'PVF key record image
Dim imPvfRecLen As Integer  'PVF record length
Dim hmDnf As Integer            'Multiname file handle
Dim imDnfRecLen As Integer      'MNF record length
Dim tmDnfSrchKey0 As INTKEY0
Dim tmDnf As DNF
Dim tmSvDnf() As DNF

'Demo Plus
Dim hmDpf As Integer        'Demo Plus handle
'Research Estimate
Dim hmDef As Integer
Dim hmRaf As Integer
'Specification Area
Dim tmSpecCtrls(0 To 12) As FIELDAREA 'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
Dim imLBSpecCtrls As Integer
Dim imSpecBoxNo As Integer
Dim smSpecSave(0 To 7) As String  'Values saved (1=Name; 2=Demo; 3=$ Index;4=Market; 5=Sales Brochure, 6=ACT1Code, 7=ACT1Setting)
Dim imSpecSave(0 To 5) As Integer 'Values saved (1=Price; 2=InvTime; 3=Alter Hidden; 4=Alter Name; 5=Programmatic)
Dim tmVehGp3Code() As SORTCODE
Dim smVehGp3CodeTag As String
Dim imVehGp3ChgMode As Integer
'Package Vehicle Areas
Dim tmPkgCtrls(0 To 10)  As FIELDAREA    'Time/Days
Dim imLBPkgCtrls As Integer
Dim imPkgBoxNo As Integer
Dim imPkgRowNo As Integer
Dim imPkgChg As Integer
Dim smTShow(0 To 7) As String
Dim tmPBDP() As RCPBDPGEN
Dim tmRdf As RDF
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imUpdateAllowed As Integer
Dim imLbcArrowSetting As Integer
Dim imTabDirection As Integer
Dim imDirProcess As Integer
Dim imDoubleClickName As Integer
Dim imLbcMouseDown As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imFirstFocus As Integer
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imComboBoxIndex As Integer
Dim imSettingValue As Integer
Dim imFirstTimeSelect As Integer
Dim bmProgrammaticAllowed As Boolean
Dim smOrigProgrammaticAllowed As String
Dim smOrigSalesBrochure As String

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Const LBONE = 1

Const AVGINDEX = RCAVGRATEINDEX 'Also in RateCard.Frm and CPMPkg screens

Const NAMEINDEX = 1          'Name control/field
Const PRICEINDEX = 2      'Price control/field
Const INVTIMEINDEX = 3         'Package Inv Time control/field
Const PROGRAMMATICINDEX = 4
Const SALESBROCHUREINDEX = 5
Const ALTERNAMEINDEX = 6     'Alter control/field
Const ALTERHIDDENINDEX = 7        'Alter control/field
Const DEMOINDEX = 8
Const MKTINDEX = 9
Const DOLLARINDEX = 10
Const ACT1CODEINDEX = 11
Const ACT1SETTINGINDEX = 12

Const PKGVEHINDEX = 1       'Vehicle control/field
Const PKGDPINDEX = 2        'Daypart control/field
Const PKGPRICEINDEX = 3     'Price control/field
Const PKGBOOKNAMEINDEX = 4
Const PKGRATINGINDEX = 5    'Rating spot control/field
Const PKGAUDINDEX = 6       'Audience
Const PKGCPPINDEX = 7       'CPP
Const PKGCPMINDEX = 8       'CPM
Const PKGSPOTINDEX = 9      '# Spots
Const PKGPERCENTINDEX = 10   'Percent
Dim smPVFType As String
Dim smPcfType As String
Dim bmMixedFound As Boolean
'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
'Hide cursor in inputs for ACT1Settings boxes
Private Declare Function ShowCaret Lib "user32" (ByVal HWnd As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal HWnd As Long) As Long
Dim bmAct1Allowed As Boolean
Dim bmAct1Enabled As Boolean

Private Sub cbcSelect_Change()
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode Then  'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    mClearCtrlFields
    If ilRet = 0 Then
        ilIndex = cbcSelect.ListIndex
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cbcSelectErr
        End If
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    pbcSpec.Cls
    pbcPkg.Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
        mInitShow
        If lbcDemo.ListIndex > 0 Then
            mGetPkgAud -1
        End If
        mGetTotals
    Else
        imSelectedIndex = 0
        If slStr <> "[New]" Then
            edcSpecDropDown.MaxLength = 20
            edcSpecDropDown.Text = slStr
            mSpecSetShow NAMEINDEX
        End If
    End If
    imFirstTimeSelect = True
    pbcSpec_Paint
    pbcPkg_Paint
    imChgMode = False
    imBypassSetting = False
    mSetCommands
    Screen.MousePointer = vbDefault
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    Exit Sub
End Sub

Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub

Private Sub cbcSelect_DropDown()
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
End Sub

Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = -1
    imPkgRowNo = -1
    pbcArrow.Visible = False
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'mInitDDE
        If igDPNameCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            '5/30/19: Replaced settiong ListIndex = 0 in mInit to adviod multi-calss to create screen
            'If sgDPName = "" Then
            '    cbcSelect.ListIndex = 0
            'Else
            '    cbcSelect.Text = sgDPName    'New name
            'End If
            'cbcSelect_Change
            If sgDPName <> "" Then
                mSetCommands
                gFindMatch sgDPName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    'DL 6/19/03: Unable to setfocus to control
                    'This was is caused with the code added to force the scroll bar to show
                    'Form_activate: Me.Visible=False, followed by DoEvents, followed by me.Visible = true
                    'cmcDone.SetFocus
                    Exit Sub
                End If
            End If
            'If pbcSpecSTab.Enabled Then
            '    pbcSpecSTab.SetFocus
            'Else
            '    cmcCancel.SetFocus
            'End If
            '3/13/06: Force repaint
            vbcPkg.Visible = False
            DoEvents
            vbcPkg.Visible = True
            Exit Sub
        End If
    End If
    slSvText = cbcSelect.Text
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields 'Make sure all fields cleared
        imFirstTimeSelect = True
        pbcStartNew.SetFocus
        Exit Sub
    End If
    gCtrlGotFocus cbcSelect
    If (slSvText = "") Or (slSvText = "[New]") Then
        cbcSelect.ListIndex = 0
        cbcSelect_Change    'Call change so picture area repainted
    Else
        gFindMatch slSvText, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                cbcSelect_Change    'Call change so picture area repainted
                imPopReqd = False
            End If
        Else
            cbcSelect.ListIndex = 0
            mClearCtrlFields
            cbcSelect_Change    'Call change so picture area repainted
        End If
    End If
End Sub

Private Sub cmcCancel_Click()
    If igDPNameCallSource <> CALLNONE Then
        igDPNameCallSource = CALLCANCELLED
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = -1
    imPkgRowNo = -1
    pbcArrow.Visible = False
    gCtrlGotFocus cmcCancel
End Sub

Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If igDPNameCallSource <> CALLNONE Then
        sgDPName = smSpecSave(1) 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgDPName = "[New]"
            If Not imTerminate Then
                mSpecEnableBox imSpecBoxNo
                Exit Sub
            Else
                cmcCancel_Click
                Exit Sub
            End If
        End If
    Else
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mSpecEnableBox imSpecBoxNo
            Exit Sub
        End If
    End If
    If igDPNameCallSource <> CALLNONE Then
        If sgDPName = "[New]" Then
            igDPNameCallSource = CALLCANCELLED
        Else
            igDPNameCallSource = CALLDONE
        End If
        mTerminate
        Exit Sub
    End If
    mTerminate
End Sub

Private Sub cmcDone_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = -1
    imPkgRowNo = -1
    pbcArrow.Visible = False
    gCtrlGotFocus cmcDone
End Sub

Private Sub cmcDropDown_Click()
    Select Case imPkgBoxNo
        Case PKGSPOTINDEX
        Case PKGPERCENTINDEX
    End Select
    edcSpecDropDown.SelStart = 0
    edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    edcSpecDropDown.SetFocus
End Sub

Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcProof_Click()
    Dim hlProof As Integer
    Dim ilRet As Integer
    Dim slToFile As String
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim ilLoop As Integer
    Dim slDateTime As String
    
    '8/1/18: Check for illegal characters in name
    'slToFile = sgExportPath & Trim$(smSpecSave(1)) & ".csv"
    slToFile = sgExportPath & gFileNameFilter(Trim$(smSpecSave(1))) & ".csv"
    ilRet = 0
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        Kill slToFile
    End If
    ilRet = 0
    ilRet = gFileOpen(slToFile, "Output", hlProof)
    If ilRet <> 0 Then
        MsgBox "Open " & slToFile & ", Error #" & Str(err.Number), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
        Exit Sub
    End If
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        If Trim$(smPkgSave(1, ilLoop)) <> "" Then
            If Val(Trim$(smPkgSave(1, ilLoop))) > 0 Then
                ilVef = gBinarySearchVef(tmPBDP(ilLoop).iVefCode)
                If ilVef <> -1 Then
                    ilRdf = gBinarySearchRdf(tmPBDP(ilLoop).iRdfCode)
                    If ilRdf <> -1 Then
                        Print #hlProof, Trim$(tgMVef(ilVef).sName) & "," & Trim$(tgMRdf(ilRdf).sName)
                    End If
                End If
            End If
        End If
    Next ilLoop
    Close #hlProof
    MsgBox "Create file: " & slToFile, vbOKOnly + vbApplicationModal, "File Saved"
    Exit Sub
End Sub

Private Sub cmcSpecDropDown_Click()
    Select Case imSpecBoxNo
        Case NAMEINDEX
        Case PRICEINDEX
        Case INVTIMEINDEX
        Case PROGRAMMATICINDEX
        Case SALESBROCHUREINDEX
            lbcPDFName.Visible = Not lbcPDFName.Visible
        Case ALTERNAMEINDEX
        Case ALTERHIDDENINDEX
        Case DEMOINDEX
            lbcDemo.Visible = Not lbcDemo.Visible
        Case MKTINDEX
            lbcVehGp3.Visible = Not lbcVehGp3.Visible
        Case DOLLARINDEX
        Case ACT1CODEINDEX
        Case ACT1SETTINGINDEX
    End Select
    edcSpecDropDown.SelStart = 0
    edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    edcSpecDropDown.SetFocus
End Sub

Private Sub cmcSpecDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcUpdate_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    Dim ilLoop As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slName = cbcSelect.Text   'Save name
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mSpecEnableBox imSpecBoxNo
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imSpecBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
    ilCode = tmVef.iCode
    cbcSelect.Clear
    smPkgVehicleTag = ""
    mPopulate
    For ilLoop = 0 To UBound(tmPkgVehicle) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
        slNameCode = tmPkgVehicle(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If Val(slCode) = ilCode Then
            If cbcSelect.ListIndex = ilLoop + 1 Then
                cbcSelect_Change
            Else
                cbcSelect.ListIndex = ilLoop + 1
            End If
            Exit For
        End If
    Next ilLoop
    Screen.MousePointer = vbDefault
    mSetCommands
    If cbcSelect.Enabled = True Then cbcSelect.SetFocus
End Sub

'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
Private Sub edcACT1SettingC_Click()
    If edcACT1SettingC.Text = "No" Then
        edcACT1SettingC.Text = "Yes"
    Else
        edcACT1SettingC.Text = "No"
    End If
End Sub

Private Sub edcACT1SettingC_GotFocus()
    HideCaret edcACT1SettingC.HWnd
    edcACT1SettingC.BackColor = &HFFFF00
End Sub

Private Sub edcACT1SettingC_KeyPress(KeyAscii As Integer)
    If edcACT1SettingC.Text = "No" Then
        edcACT1SettingC.Text = "Yes"
    Else
        edcACT1SettingC.Text = "No"
    End If
    If UCase(Chr(KeyAscii)) = "Y" Then edcACT1SettingC.Text = "Yes"
    If UCase(Chr(KeyAscii)) = "N" Then edcACT1SettingC.Text = "No"
End Sub

Private Sub edcACT1SettingC_LostFocus()
    edcACT1SettingC.BackColor = &HFFFFFF
End Sub

Private Sub edcACT1SettingF_Click()
    If edcACT1SettingF.Text = "No" Then
        edcACT1SettingF.Text = "Yes"
    Else
        edcACT1SettingF.Text = "No"
    End If
End Sub

Private Sub edcACT1SettingF_GotFocus()
    HideCaret edcACT1SettingF.HWnd
    edcACT1SettingF.BackColor = &HFFFF00
End Sub

Private Sub edcACT1SettingF_KeyPress(KeyAscii As Integer)
    If edcACT1SettingF.Text = "No" Then
        edcACT1SettingF.Text = "Yes"
    Else
        edcACT1SettingF.Text = "No"
    End If
    If UCase(Chr(KeyAscii)) = "Y" Then edcACT1SettingF.Text = "Yes"
    If UCase(Chr(KeyAscii)) = "N" Then edcACT1SettingF.Text = "No"
End Sub

Private Sub edcACT1SettingF_LostFocus()
    edcACT1SettingF.BackColor = &HFFFFFF
End Sub

Private Sub edcACT1SettingS_Click()
    If edcACT1SettingS.Text = "No" Then
        edcACT1SettingS.Text = "Yes"
    Else
        edcACT1SettingS.Text = "No"
    End If
End Sub

Private Sub edcACT1SettingS_GotFocus()
    HideCaret edcACT1SettingS.HWnd
    edcACT1SettingS.BackColor = &HFFFF00
End Sub

Private Sub edcACT1SettingS_KeyPress(KeyAscii As Integer)
    If edcACT1SettingS.Text = "No" Then
        edcACT1SettingS.Text = "Yes"
    Else
        edcACT1SettingS.Text = "No"
    End If
    If UCase(Chr(KeyAscii)) = "Y" Then edcACT1SettingS.Text = "Yes"
    If UCase(Chr(KeyAscii)) = "N" Then edcACT1SettingS.Text = "No"
End Sub

Private Sub edcACT1SettingS_LostFocus()
    edcACT1SettingS.BackColor = &HFFFFFF
End Sub

Private Sub edcACT1SettingT_Click()
    If edcACT1SettingT.Text = "No" Then
        edcACT1SettingT.Text = "Yes"
    Else
        edcACT1SettingT.Text = "No"
    End If
End Sub

Private Sub edcACT1SettingT_GotFocus()
    HideCaret edcACT1SettingT.HWnd
    edcACT1SettingT.BackColor = &HFFFF00
End Sub

Private Sub edcACT1SettingT_KeyPress(KeyAscii As Integer)
    If edcACT1SettingT.Text = "No" Then
        edcACT1SettingT.Text = "Yes"
    Else
        edcACT1SettingT.Text = "No"
    End If
    If UCase(Chr(KeyAscii)) = "Y" Then edcACT1SettingT.Text = "Yes"
    If UCase(Chr(KeyAscii)) = "N" Then edcACT1SettingT.Text = "No"
End Sub

Private Sub edcACT1SettingT_LostFocus()
    edcACT1SettingT.BackColor = &HFFFFFF
End Sub

Private Sub edcDropDown_Change()
    Select Case imPkgBoxNo
        Case PKGSPOTINDEX
        Case PKGPERCENTINDEX
    End Select
End Sub

Private Sub edcDropDown_GotFocus()
    Select Case imPkgBoxNo
        Case PKGSPOTINDEX
        Case PKGPERCENTINDEX
    End Select
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imPkgBoxNo
        Case PKGSPOTINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case PKGPERCENTINDEX
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
            If gCompNumberStr(slStr, "100.00") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub

Private Sub edcSpecDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imSpecBoxNo
        Case NAMEINDEX
        Case PRICEINDEX
        Case INVTIMEINDEX
        Case PROGRAMMATICINDEX
        Case SALESBROCHUREINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcSpecDropDown, lbcPDFName, imBSMode, slStr)
            If ilRet = 1 Then
                lbcPDFName.ListIndex = 0
            End If
            imLbcArrowSetting = False
        Case ALTERNAMEINDEX
        Case ALTERHIDDENINDEX
        Case DEMOINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcSpecDropDown, lbcDemo, imBSMode, imComboBoxIndex
        Case MKTINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcSpecDropDown, lbcVehGp3, imBSMode, slStr)
            If ilRet = 1 Then
                lbcVehGp3.ListIndex = 0
            End If
            imLbcArrowSetting = False
        Case DOLLARINDEX
        Case ACT1CODEINDEX
        Case ACT1SETTINGINDEX
    End Select
    imLbcArrowSetting = False
End Sub

Private Sub edcSpecDropDown_DblClick()
    'imDoubleClickName = True    'Double click event foolowed by mouse up
    If imSpecBoxNo = MKTINDEX Then
        imDoubleClickName = True
    End If
End Sub

Private Sub edcSpecDropDown_GotFocus()
    Select Case imSpecBoxNo
        Case NAMEINDEX
        Case PRICEINDEX
        Case INVTIMEINDEX
        Case PROGRAMMATICINDEX
        Case SALESBROCHUREINDEX
        Case ALTERNAMEINDEX
        Case ALTERHIDDENINDEX
        Case DEMOINDEX
            If lbcDemo.ListCount = 1 Then
                lbcDemo.ListIndex = 0
            End If
        Case DOLLARINDEX
        Case ACT1CODEINDEX
        Case ACT1SETTINGINDEX
    End Select
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSpecDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcSpecDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim ilPos As Integer
    Dim slStr As String
    Dim slComp As String
    If imSpecBoxNo = DOLLARINDEX Then
        ilPos = InStr(edcSpecDropDown.SelText, ".")
        If ilPos = 0 Then
            ilPos = InStr(edcSpecDropDown.Text, ".")    'Disallow multi-decimal points
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
        slStr = edcSpecDropDown.Text
        slStr = Left$(slStr, edcSpecDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSpecDropDown.SelStart - edcSpecDropDown.SelLength)
        slComp = "99.99"
        If gCompNumberStr(slStr, slComp) > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Else
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
    End If
End Sub

Private Sub edcSpecDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KeyDown) Then
        imDirProcess = KeyCode 'mDirection 0
        pbcSpecTab.SetFocus
        Exit Sub
    End If
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imSpecBoxNo
            Case NAMEINDEX
            Case PRICEINDEX
            Case INVTIMEINDEX
            Case PROGRAMMATICINDEX
            Case SALESBROCHUREINDEX
                gProcessArrowKey Shift, KeyCode, lbcPDFName, imLbcArrowSetting
            Case ALTERNAMEINDEX
            Case ALTERHIDDENINDEX
            Case DEMOINDEX
                gProcessArrowKey Shift, KeyCode, lbcDemo, imLbcArrowSetting
            Case MKTINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehGp3, imLbcArrowSetting
            Case DOLLARINDEX
            Case ACT1CODEINDEX
            Case ACT1SETTINGINDEX
        End Select
        edcSpecDropDown.SelStart = 0
        edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    End If
End Sub

Private Sub edcSpecDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imSpecBoxNo
            Case NAMEINDEX
            Case PRICEINDEX
            Case INVTIMEINDEX
            Case PROGRAMMATICINDEX
            Case SALESBROCHUREINDEX
            Case ALTERNAMEINDEX
            Case ALTERHIDDENINDEX
            Case DEMOINDEX
            Case MKTINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSpecSTab.SetFocus
                Else
                    pbcSpecTab.SetFocus
                End If
                Exit Sub
            Case DOLLARINDEX
            Case ACT1CODEINDEX
            Case ACT1SETTINGINDEX

        End Select
        imDoubleClickName = False
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
    If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcSpec.Enabled = False
        pbcPkg.Enabled = False
        pbcSpecSTab.Enabled = False
        pbcSpecTab.Enabled = False
        pbcPkgSTab.Enabled = False
        pbcPkgTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcSpec.Enabled = True
        pbcPkg.Enabled = True
        pbcSpecSTab.Enabled = True
        pbcSpecTab.Enabled = True
        pbcPkgSTab.Enabled = True
        pbcPkgTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    StdPkg.Refresh
    vbcPkg.Visible = False
    DoEvents
    vbcPkg.Visible = True
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
    '11/4/21 - JW - Fix Per Dan found some gremlins
    'mSpecSetShow imSpecBoxNo
    'imSpecBoxNo = -1
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If (cbcSelect.Enabled) And ((imSpecBoxNo > 0) Or (imPkgBoxNo > 0)) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imSpecBoxNo > 0 Then
            mSpecEnableBox imSpecBoxNo
        ElseIf imPkgBoxNo > 0 Then
            mPkgEnableBox imPkgBoxNo
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
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = (((lgPercentAdjW - 10) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = ((lgPercentAdjW - 10) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
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
        mSpecSetShow imSpecBoxNo
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            mSpecEnableBox imSpecBoxNo
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    Cancel = 0
    Erase tmSvDnf
    Erase tmPvf
    Erase tmPBDP
    Erase smPkgShow
    Erase smPkgSave
    Erase tmPkgVehicle
    Erase tmVehGp3Code
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    ilRet = btrClose(hmVff)
    btrDestroy hmVff
    ilRet = btrClose(hmPvf)
    btrDestroy hmPvf
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmDpf)
    btrDestroy hmDpf
    ilRet = btrClose(hmDef)
    btrDestroy hmDef
    ilRet = btrClose(hmRaf)
    btrDestroy hmRaf
    ilRet = btrClose(hmDnf)
    btrDestroy hmDnf
    Set StdPkg = Nothing   'Remove data segment
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

Private Sub lbcDemo_Click()
    gProcessLbcClick lbcDemo, edcSpecDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcDemo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcPDFName_Click()
    gProcessLbcClick lbcPDFName, edcSpecDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcPDFName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcVehGp3_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcVehGp3, edcSpecDropDown, imVehGp3ChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcVehGp3_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcVehGp3_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcVehGp3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcVehGp3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcVehGp3, edcSpecDropDown, imVehGp3ChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSpecSTab.SetFocus
        Else
            pbcSpecTab.SetFocus
        End If
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mAdjPrices                      *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Adjust average price by $ index*
'*                                                     *
'*******************************************************
Private Sub mAdjPrices()
    Dim ilLoop As Integer
    Dim llDollarIndex As Long
    Dim slStr As String
    llDollarIndex = gStrDecToLong(smSpecSave(3), 2)
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        tmPBDP(ilLoop).lAvgPrice = (CSng(tmPBDP(ilLoop).lSvAvgPrice) * llDollarIndex) / 100
        slStr = gLongToStrDec(tmPBDP(ilLoop).lAvgPrice, 0)
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGPRICEINDEX)
        smPkgShow(PKGPRICEINDEX, ilLoop) = tmPkgCtrls(PKGPRICEINDEX).sShow
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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
    Dim slStr As String
    imPkgChg = False
    lbcDemo.ListIndex = -1
    smOrigDemo = ""
    lbcVehGp3.ListIndex = -1
    smOrigVehGp3 = ""
    For ilLoop = LBound(smSpecSave) To UBound(smSpecSave) Step 1
        smSpecSave(ilLoop) = ""
    Next ilLoop
    For ilLoop = LBound(imSpecSave) To UBound(imSpecSave) Step 1
        imSpecSave(ilLoop) = -1
    Next ilLoop
    For ilLoop = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        tmSpecCtrls(ilLoop).sShow = ""
        tmSpecCtrls(ilLoop).iChg = False
    Next ilLoop
    tmVef.sName = ""
    tmVef.sStdPrice = ""
    tmVef.sStdInvTime = ""
    tmVef.sStdAlter = ""
    tmVef.sStdAlterName = ""
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        tmPBDP(ilLoop).sKey = "1" & tmPBDP(ilLoop).sSvKey
        tmPBDP(ilLoop).lAvgPrice = tmPBDP(ilLoop).lSvAvgPrice
        tmPBDP(ilLoop).lAvgAud = 0
        tmPBDP(ilLoop).iAvgRating = 0
        tmPBDP(ilLoop).lCPP = 0
        tmPBDP(ilLoop).lCPM = 0
        tmPBDP(ilLoop).lPop = 0
    Next ilLoop
    smOrigProgrammaticAllowed = ""
    smOrigSalesBrochure = ""
    mInitShowFields
    For ilLoop = LBONE To UBound(smPkgSave, 2) - 1 Step 1
        slStr = ""
        smPkgSave(1, ilLoop) = slStr
        slStr = ""
        smPkgSave(2, ilLoop) = slStr
    Next ilLoop
    For ilLoop = LBound(smTShow) To UBound(smTShow) Step 1
        smTShow(ilLoop) = ""
    Next ilLoop
    vbcPkg.Value = vbcPkg.Min
    ReDim tmPvf(0 To 0) As PVF
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mDemoPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Potential Code        *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mDemoPop()
'
'   mDemoPop
'   Where:
'
    Dim ilRet As Integer
    Dim slDemo As String      'Demo name, saved to determine if changed
    Dim ilIndex As Integer      'Demo name, saved to determine if changed
    ilIndex = lbcDemo.ListIndex
    If ilIndex > 0 Then
        slDemo = lbcDemo.List(ilIndex)
    End If
    ilRet = gPopMnfPlusFieldsBox(StdPkg, lbcDemo, tgDemoCode(), sgDemoCodeTag, "D")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mDemoPopErr
        gCPErrorMsg ilRet, "mDemoPop (gPopMnfPlusFieldsBox)", StdPkg
        On Error GoTo 0
        lbcDemo.AddItem "[None]", 0
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slDemo, 1, lbcDemo
            If gLastFound(lbcDemo) > 0 Then
                lbcDemo.ListIndex = gLastFound(lbcDemo)
            Else
                lbcDemo.ListIndex = -1
            End If
        Else
            lbcDemo.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mDemoPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetPkgAud                      *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get Package Audience            *
'*                                                     *
'*******************************************************
Private Sub mGetPkgAud(ilIndexTlPBDp As Integer)
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim ilTime As Integer
    Dim ilDay As Integer
    Dim ilMnfDemo As Integer
    Dim ilDnfCode As Integer
    Dim ilMnfSocEco As Integer
    Dim ilRdfCode As Integer
    Dim llPop As Long
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    ReDim ilDays(0 To 6) As Integer
    ReDim llWkSpotCount(0 To 0) As Long
    ReDim llWkActPrice(0 To 0) As Long
    ReDim llWkAvgAud(0 To 0) As Long
    ReDim llWkPopEst(0 To 0) As Long
    Dim llAvgAud As Long
    'Dim llLnCost As Long
    Dim dlLnCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim llAvgAudAvg As Long
    ReDim ilWkRating(0 To 0) As Integer
    Dim ilLnAvgRating As Integer
    ReDim llWkGrImp(0 To 0) As Long
    ReDim llWkGRP(0 To 0) As Long
    Dim llLnGRP As Long
    Dim llLnGrImp As Long
    Dim llCPP As Long
    Dim llCPM As Long
    Dim slStr As String
    Dim llDate As Long
    Dim llPopEst As Long
    Dim ilSLoop As Integer
    Dim ilELoop As Integer
    Dim llRafCode As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long

    llOvStartTime = 0
    llOvEndTime = 0
    If lbcDemo.ListIndex <= 0 Then
        ilMnfDemo = 0
    Else
        slNameCode = tgDemoCode(lbcDemo.ListIndex - 1).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilMnfDemo = Val(slCode)
    End If
    If (ilIndexTlPBDp >= LBONE) And (ilIndexTlPBDp < UBound(tmPBDP)) Then
        ilSLoop = ilIndexTlPBDp
        ilELoop = ilIndexTlPBDp + 1
    Else
        ilSLoop = LBONE 'LBound(tmPBDP)
        ilELoop = UBound(tmPBDP)
    End If
    For ilLoop = ilSLoop To ilELoop - 1 Step 1
        'Vehicle
        ilVefCode = tmPBDP(ilLoop).iVefCode
        ilRdfCode = tmPBDP(ilLoop).iRdfCode
        ilDnfCode = -1
        ilVef = gBinarySearchVef(tmPBDP(ilLoop).iVefCode)
        If ilVef <> -1 Then
            ilDnfCode = tgMVef(ilVef).iDnfCode
        End If
        ilMnfSocEco = 0
        If (ilDnfCode > 0) And (ilMnfDemo > 0) Then
            ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, ilMnfSocEco, ilMnfDemo, llPop)
        Else
            llPop = 0
        tmPBDP(ilLoop).lPop = llPop
        End If
        'Build record into tmPBDP
        For ilDay = 0 To 6 Step 1
            ilDays(ilDay) = False
        Next ilDay
        For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
            If tmPBDP(ilLoop).iRdfCode = tgMRdf(ilRdf).iCode Then
                tmRdf = tgMRdf(ilRdf)
                For ilTime = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1  'Row
                    If (tmRdf.iStartTime(0, ilTime) <> 1) Or (tmRdf.iStartTime(1, ilTime) <> 0) Then
                        For ilDay = 1 To 7 Step 1
                            If tmRdf.sWkDays(ilTime, ilDay - 1) = "Y" Then
                                ilDays(ilDay - 1) = True
                            End If
                        Next ilDay
                    End If
                Next ilTime
                Exit For
            End If
        Next ilRdf
        If (ilDnfCode > 0) And (ilVefCode > 0) And (ilMnfDemo > 0) Then
            llDate = 0
            llRafCode = 0
            ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, ilVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, tmRdf.iCode, llOvStartTime, llOvEndTime, ilDays(), "S", llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
        Else
            llAvgAud = 0
        End If
        tmPBDP(ilLoop).lAvgAud = llAvgAud

        'Get Rating
        'Get avg audience
        'Get CPP, CPM
        If Trim$(smPkgSave(1, ilLoop)) <> "" Then
            llWkSpotCount(0) = Val(Trim$(smPkgSave(1, ilLoop)))
        Else
            llWkSpotCount(0) = 0
        End If
        llWkActPrice(0) = 100 * tmPBDP(ilLoop).lAvgPrice
        llWkAvgAud(0) = llAvgAud
        llWkPopEst(0) = llPopEst
        gAvgAudToLnResearch "1", False, llPop, llWkPopEst(), llWkSpotCount(), llWkActPrice(), llWkAvgAud(), dlLnCost, llAvgAudAvg, ilWkRating(), ilLnAvgRating, llWkGrImp(), llLnGrImp, llWkGRP(), llLnGRP, llCPP, llCPM, llPopEst
        tmPBDP(ilLoop).iAvgRating = ilLnAvgRating
        tmPBDP(ilLoop).lGrImp = llLnGrImp
        tmPBDP(ilLoop).lGRP = llLnGRP
        tmPBDP(ilLoop).lCPP = llCPP
        tmPBDP(ilLoop).lCPM = llCPM
        tmPBDP(ilLoop).lPop = llPop
    Next ilLoop
    For ilLoop = ilSLoop To ilELoop - 1 Step 1
        If lbcDemo.ListIndex <= 0 Then
            slStr = ""
        Else
            slStr = gIntToStrDec(tmPBDP(ilLoop).iAvgRating, 1)
        End If
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGRATINGINDEX)
        smPkgShow(PKGRATINGINDEX, ilLoop) = tmPkgCtrls(PKGRATINGINDEX).sShow
        If lbcDemo.ListIndex <= 0 Then
            slStr = ""
        Else
            'slStr = Trim$(Str$(tmPBDP(ilLoop).lAvgAud))
            If tgSpf.sSAudData = "H" Then
                slStr = gLongToStrDec(tmPBDP(ilLoop).lAvgAud, 1)
            ElseIf tgSpf.sSAudData = "N" Then
                slStr = gLongToStrDec(tmPBDP(ilLoop).lAvgAud, 2)
            ElseIf tgSpf.sSAudData = "U" Then
                slStr = gLongToStrDec(tmPBDP(ilLoop).lAvgAud, 3)
            Else
                slStr = Trim$(Str$(tmPBDP(ilLoop).lAvgAud))
            End If
        End If
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGAUDINDEX)
        smPkgShow(PKGAUDINDEX, ilLoop) = tmPkgCtrls(PKGAUDINDEX).sShow
        If lbcDemo.ListIndex <= 0 Then
            slStr = ""
        Else
            slStr = Trim$(Str$(tmPBDP(ilLoop).lCPP))
        End If
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGCPPINDEX)
        smPkgShow(PKGCPPINDEX, ilLoop) = tmPkgCtrls(PKGCPPINDEX).sShow
        If lbcDemo.ListIndex <= 0 Then
            slStr = ""
        Else
            slStr = gLongToStrDec(tmPBDP(ilLoop).lCPM, 2)
        End If
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGCPMINDEX)
        smPkgShow(PKGCPMINDEX, ilLoop) = tmPkgCtrls(PKGCPMINDEX).sShow
    Next ilLoop
    pbcPkg_Paint
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetTotals                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine totals               *
'*                                                     *
'*******************************************************
Private Sub mGetTotals()
    Dim ilCount As Integer
    Dim ilLoop As Integer
    Dim llPop As Long
    'Dim llTCost As Long
    Dim dlTCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim ilAvgRate As Integer
    Dim ilTNoSpots As Integer
    Dim llTAud As Long
    Dim llTGrImp As Long
    Dim llTGRP As Long
    Dim llTCPP As Long
    Dim llTCPM As Long
    Dim slStr As String
    Dim slTotalPct As String
    Dim llLnSpots As Long
    Dim llAvgAud As Long
    ilCount = 0
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        If Trim$(smPkgSave(1, ilLoop)) <> "" Then
            If Val(Trim$(smPkgSave(1, ilLoop))) > 0 Then
                ilCount = ilCount + 1
            End If
        End If
    Next ilLoop
    slTotalPct = "0.0"
    If ilCount <= 0 Then
        For ilLoop = LBound(smTShow) To UBound(smTShow) Step 1
            smTShow(ilLoop) = ""
        Next ilLoop
    Else
        ilCount = ilCount - 1
        ReDim llCost(0 To ilCount) As Long
        ReDim ilAvgRating(0 To ilCount) As Integer
        ReDim llGrImp(0 To ilCount) As Long
        ReDim llGRP(0 To ilCount) As Long
        ilCount = 0
        llTAud = 0
        ilTNoSpots = 0
        For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
            If Trim$(smPkgSave(1, ilLoop)) <> "" Then
                If Val(Trim$(smPkgSave(1, ilLoop))) > 0 Then
                    llTAud = llTAud + tmPBDP(ilLoop).lAvgAud
                    ilTNoSpots = ilTNoSpots + Val(Trim$(smPkgSave(1, ilLoop)))
                    llCost(ilCount) = 100 * Val(Trim$(smPkgSave(1, ilLoop))) * tmPBDP(ilLoop).lAvgPrice
                    ilAvgRating(ilCount) = tmPBDP(ilLoop).iAvgRating
                    llGrImp(ilCount) = tmPBDP(ilLoop).lGrImp
                    llGRP(ilCount) = tmPBDP(ilLoop).lGRP
                    If ilCount = 0 Then
                        llPop = tmPBDP(ilLoop).lPop
                    Else
                        If llPop <> tmPBDP(ilLoop).lPop Then
                            llPop = 0
                        End If
                    End If
                    If imSpecSave(1) = 2 Then
                        slTotalPct = gAddStr(Trim$(smPkgSave(2, ilLoop)), slTotalPct)
                    End If
                    ilCount = ilCount + 1
                End If
            End If
        Next ilLoop
        llLnSpots = 1   'ilTNoSpots
        'gResearchTotals "1", False, llPop, llCost(), llGrImp(), llGRP(), llLnSpots, llTCost, ilAvgRate, llTGrImp, llTGRP, llTCPP, llTCPM, llAvgAud
        gResearchTotals "1", False, llPop, llCost(), llGrImp(), llGRP(), llLnSpots, dlTCost, ilAvgRate, llTGrImp, llTGRP, llTCPP, llTCPM, llAvgAud 'TTP 10439 - Rerate 21,000,000
        'slStr = gLongToStrDec(llTCost / 100, 0)
        slStr = gDblToStrDec(dlTCost / 100, 0) 'TTP 10439 - Rerate 21,000,000
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGPRICEINDEX)
        smTShow(1) = tmPkgCtrls(PKGPRICEINDEX).sShow
        slStr = gIntToStrDec(ilAvgRate, 1)
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGRATINGINDEX)
        smTShow(2) = tmPkgCtrls(PKGRATINGINDEX).sShow
        'slStr = Trim$(Str$(llTAud))
        If tgSpf.sSAudData = "H" Then
            slStr = gLongToStrDec(llTGrImp, 1)
        ElseIf tgSpf.sSAudData = "N" Then
            slStr = gLongToStrDec(llTGrImp, 2)
        ElseIf tgSpf.sSAudData = "U" Then
            slStr = gLongToStrDec(llTGrImp, 3)
        Else
            slStr = Trim$(Str$(llTGrImp))
        End If
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGAUDINDEX)
        smTShow(3) = tmPkgCtrls(PKGAUDINDEX).sShow
        slStr = Trim$(Str$(llTCPP))
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGCPPINDEX)
        smTShow(4) = tmPkgCtrls(PKGCPPINDEX).sShow
        slStr = gLongToStrDec(llTCPM, 2)
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGCPMINDEX)
        smTShow(5) = tmPkgCtrls(PKGCPMINDEX).sShow
        slStr = Trim$(Str$(ilTNoSpots))
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGSPOTINDEX)
        smTShow(6) = tmPkgCtrls(PKGSPOTINDEX).sShow
        If imSpecSave(1) = 2 Then
            slStr = slTotalPct
        Else
            slStr = ""
        End If
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGPERCENTINDEX)
        smTShow(7) = tmPkgCtrls(PKGPERCENTINDEX).sShow
    End If
    lacCover.Visible = True
    lacCover.Visible = False
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Screen.MousePointer = vbHourglass
    imLBSpecCtrls = 1
    imLBPkgCtrls = 1
    imAddPkg = False
    imFirstActivate = True
    imcKey.Picture = IconTraf!imcKey.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    igDPAltered = False
    imTerminate = False
    imBypassSetting = False
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imChgMode = False
    imBSMode = False
    imPopReqd = False
    imFirstFocus = True
    imSelectedIndex = -1
    imVehGp3ChgMode = False
    imPkgBoxNo = -1
    imSpecBoxNo = -1
    imPkgChg = False
    imSettingValue = False
    imFirstTimeSelect = True
    hmVef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", StdPkg
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVpf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vpf.Btr)", StdPkg
    On Error GoTo 0
    imVpfRecLen = Len(tmVpf)
    hmVff = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVff, "", sgDBPath & "Vff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vff.Btr)", StdPkg
    On Error GoTo 0
    imVffRecLen = Len(tmVff)
    ReDim tmPvf(0 To 0) As PVF
    hmPvf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmPvf, "", sgDBPath & "Pvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pvf.Btr)", StdPkg
    On Error GoTo 0
    imPvfRecLen = Len(tmPvf(0))
    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "DRF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: DRF.Btr)", StdPkg
    On Error GoTo 0
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "MNF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: MNF.Btr)", StdPkg
    On Error GoTo 0
    hmDpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dpf.Btr)", StdPkg
    On Error GoTo 0
    ' setup global variable for Demo Plus file (to see if any exists)
    lgDpfNoRecs = btrRecords(hmDpf)
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = -1
    End If
    hmDef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Def.Btr)", StdPkg
    On Error GoTo 0
    hmRaf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", StdPkg
    
    hmDnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", StdPkg
    imDnfRecLen = Len(tmDnf)
    ReDim tmSvDnf(0 To 0) As DNF
    
    On Error GoTo 0
    'TTP 10325 - JW 11/3/21 - Moved above spec drawing so that it can color yellow
    If (Asc(tgSaf(0).sFeatures5) And PROGRAMMATICALLOWED) = PROGRAMMATICALLOWED Then 'Programmatic Buy Allowed
        bmProgrammaticAllowed = True
    Else
        bmProgrammaticAllowed = False
    End If
    'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
    If (Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES Then 'ACT1 Enabled
        bmAct1Enabled = True
    Else
        bmAct1Enabled = False
    End If
    
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainRdf(sgMRdfStamp, tgMRdf())
    mDemoPop
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    mInitPkg
    lbcVehGp3.Clear
    mVehGp3Pop
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If Not imTerminate Then
        '5/30/19: moved here from cbcSelect_GotFocus to aviod multi-calls to create screen
        'cbcSelect.ListIndex = 0 'This will generate a select_change event
        If igDPNameCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgDPName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgDPName    'New name
            End If
        End If
        mSetCommands
    End If
    
    mPDFPop
    Screen.MousePointer = vbHourglass  'Wait
    gCenterStdAlone StdPkg
    tmcInit.Enabled = True
    
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
    Dim flTextHeight As Single  'Standard text height
    Dim llMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long
    Dim ilLoop As Integer
    Dim ilBox As Integer
    Dim ilXPos As Single
    Dim ilYPos As Single
    Dim ilWidth As Single
    
    '-------------------------------------
    'Header Grid (plcSpec)
    'TTP 10325 - JW 11/3/21 - Add ACT1 Code and settings
    flTextHeight = pbcSpec.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcSpec.Move 240, 540, pbcSpec.Width + fgPanelAdj, pbcSpec.Height + fgPanelAdj
    pbcSpec.Move plcSpec.Left + fgBevelX, plcSpec.Top + fgBevelY
    'Package Name
    ilXPos = 30: ilYPos = 30: ilWidth = 2400
    gSetCtrl tmSpecCtrls(NAMEINDEX), ilXPos, ilYPos, ilWidth, fgBoxStH
    'Price
    ilXPos = tmSpecCtrls(NAMEINDEX).fBoxX + tmSpecCtrls(NAMEINDEX).fBoxW + 15:  ilWidth = 1035
    gSetCtrl tmSpecCtrls(PRICEINDEX), ilXPos, ilYPos, ilWidth, fgBoxStH
    'Inv Time
    ilXPos = tmSpecCtrls(PRICEINDEX).fBoxX + tmSpecCtrls(PRICEINDEX).fBoxW + 15:  ilWidth = 930
    gSetCtrl tmSpecCtrls(INVTIMEINDEX), 3495, ilYPos, ilWidth, fgBoxStH
    'Programatic
    ilXPos = tmSpecCtrls(INVTIMEINDEX).fBoxX + tmSpecCtrls(INVTIMEINDEX).fBoxW + 15:  ilWidth = 930
    gSetCtrl tmSpecCtrls(PROGRAMMATICINDEX), ilXPos, ilYPos, ilWidth, fgBoxStH
    tmSpecCtrls(PROGRAMMATICINDEX).iReq = False
    'Sales Brochure
    ilXPos = tmSpecCtrls(PROGRAMMATICINDEX).fBoxX + tmSpecCtrls(PROGRAMMATICINDEX).fBoxW + 15:  ilWidth = 1020
    gSetCtrl tmSpecCtrls(SALESBROCHUREINDEX), ilXPos, ilYPos, ilWidth, fgBoxStH
    tmSpecCtrls(SALESBROCHUREINDEX).iReq = False
    'Alter Name
    ilXPos = tmSpecCtrls(SALESBROCHUREINDEX).fBoxX + tmSpecCtrls(SALESBROCHUREINDEX).fBoxW + 15:  ilWidth = 840
    gSetCtrl tmSpecCtrls(ALTERNAMEINDEX), ilXPos, ilYPos, ilWidth, fgBoxStH
    'Alter Hidden
    ilXPos = tmSpecCtrls(ALTERNAMEINDEX).fBoxX + tmSpecCtrls(ALTERNAMEINDEX).fBoxW + 15:  ilWidth = 840
    gSetCtrl tmSpecCtrls(ALTERHIDDENINDEX), ilXPos, ilYPos, ilWidth, fgBoxStH
    'Demo
    ilXPos = tmSpecCtrls(ALTERHIDDENINDEX).fBoxX + tmSpecCtrls(ALTERHIDDENINDEX).fBoxW + 15:  ilWidth = 1035
    gSetCtrl tmSpecCtrls(DEMOINDEX), ilXPos, ilYPos, ilWidth, fgBoxStH
    'Market
    ilXPos = tmSpecCtrls(DEMOINDEX).fBoxX + tmSpecCtrls(DEMOINDEX).fBoxW + 15:  ilWidth = 1515
    gSetCtrl tmSpecCtrls(MKTINDEX), ilXPos, ilYPos, ilWidth, fgBoxStH
    If tgSpf.sMktBase <> "Y" Then
        tmSpecCtrls(MKTINDEX).iReq = False
    End If
    '$ Index
    ilXPos = tmSpecCtrls(MKTINDEX).fBoxX + tmSpecCtrls(MKTINDEX).fBoxW + 15:  ilWidth = 690
    gSetCtrl tmSpecCtrls(DOLLARINDEX), ilXPos, ilYPos, ilWidth, fgBoxStH
    'ACT1CODEINDEX
    ilXPos = tmSpecCtrls(DOLLARINDEX).fBoxX + tmSpecCtrls(DOLLARINDEX).fBoxW + 15:  ilWidth = 1245
    gSetCtrl tmSpecCtrls(ACT1CODEINDEX), ilXPos, ilYPos, ilWidth, fgBoxStH
    'ACT1SETTINGINDEX
    ilXPos = tmSpecCtrls(ACT1CODEINDEX).fBoxX + tmSpecCtrls(ACT1CODEINDEX).fBoxW + 15:  ilWidth = 555
    gSetCtrl tmSpecCtrls(ACT1SETTINGINDEX), ilXPos, ilYPos, ilWidth, fgBoxStH
    
    '10/25/14: One pixel removed from top and left side when using macromedia fireworks
    For ilBox = LBound(tmSpecCtrls) To UBound(tmSpecCtrls) Step 1
        tmSpecCtrls(ilBox).fBoxX = tmSpecCtrls(ilBox).fBoxX - 15
        tmSpecCtrls(ilBox).fBoxY = tmSpecCtrls(ilBox).fBoxY - 15
    Next ilBox
    
    '-------------------------------------
    'Detail Grid (plcPkg)
    plcPkg.Move 225, 1110, pbcPkg.Width + vbcPkg.Width + fgPanelAdj, pbcPkg.Height + fgPanelAdj
    pbcPkg.Move plcPkg.Left + fgBevelX, plcPkg.Top + fgBevelY
    pbcArrow.Move plcPkg.Left - pbcArrow.Width - 15
    'Vehicle
    gSetCtrl tmPkgCtrls(PKGVEHINDEX), 30, 225, 1710, fgBoxGridH
    'Daypart
    gSetCtrl tmPkgCtrls(PKGDPINDEX), 1755, tmPkgCtrls(PKGVEHINDEX).fBoxY, 1545, fgBoxGridH
    'Price
    gSetCtrl tmPkgCtrls(PKGPRICEINDEX), 3315, tmPkgCtrls(PKGVEHINDEX).fBoxY, 780, fgBoxGridH
    'Book Name
    gSetCtrl tmPkgCtrls(PKGBOOKNAMEINDEX), 4110, tmPkgCtrls(PKGVEHINDEX).fBoxY, 1095, fgBoxGridH
    'Rating
    gSetCtrl tmPkgCtrls(PKGRATINGINDEX), 5220, tmPkgCtrls(PKGVEHINDEX).fBoxY, 510, fgBoxGridH
    'Audience
    gSetCtrl tmPkgCtrls(PKGAUDINDEX), 5745, tmPkgCtrls(PKGVEHINDEX).fBoxY, 630, fgBoxGridH
    'CPP
    gSetCtrl tmPkgCtrls(PKGCPPINDEX), 6390, tmPkgCtrls(PKGVEHINDEX).fBoxY, 510, fgBoxGridH
    'CPM
    gSetCtrl tmPkgCtrls(PKGCPMINDEX), 6915, tmPkgCtrls(PKGVEHINDEX).fBoxY, 510, fgBoxGridH
    'Spots
    gSetCtrl tmPkgCtrls(PKGSPOTINDEX), 7440, tmPkgCtrls(PKGVEHINDEX).fBoxY, 450, fgBoxGridH
    'Percent
    gSetCtrl tmPkgCtrls(PKGPERCENTINDEX), 7905, tmPkgCtrls(PKGVEHINDEX).fBoxY, 450, fgBoxGridH
    tmPkgCtrls(PKGPERCENTINDEX).iReq = False
    
    llMax = 0
    For ilLoop = imLBPkgCtrls To UBound(tmPkgCtrls) Step 1
        tmPkgCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmPkgCtrls(ilLoop).fBoxW)
        Do While (tmPkgCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmPkgCtrls(ilLoop).fBoxW = tmPkgCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmPkgCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmPkgCtrls(ilLoop).fBoxX)
            Do While (tmPkgCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmPkgCtrls(ilLoop).fBoxX = tmPkgCtrls(ilLoop).fBoxX + 1
            Loop
            If (tmPkgCtrls(ilLoop).fBoxX > 90) Then
                Do
                    If tmPkgCtrls(ilLoop - 1).fBoxX + tmPkgCtrls(ilLoop - 1).fBoxW + 15 < tmPkgCtrls(ilLoop).fBoxX Then
                        tmPkgCtrls(ilLoop - 1).fBoxW = tmPkgCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmPkgCtrls(ilLoop - 1).fBoxX + tmPkgCtrls(ilLoop - 1).fBoxW + 15 > tmPkgCtrls(ilLoop).fBoxX Then
                        tmPkgCtrls(ilLoop - 1).fBoxW = tmPkgCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmPkgCtrls(ilLoop).fBoxX + tmPkgCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmPkgCtrls(ilLoop).fBoxX + tmPkgCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop

    pbcPkg.Picture = LoadPicture("")
    pbcPkg.Width = llMax
    plcPkg.Width = llMax + vbcPkg.Width + 2 * fgBevelX + 15
    Me.Width = plcPkg.Width + 2 * plcPkg.Left
    cbcSelect.Left = plcSpec.Left + plcSpec.Width - cbcSelect.Width
    lacPkgFrame.Width = llMax - 15
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    cmcDone.Left = (StdPkg.Width - 4 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
    cmcUpdate.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    cmcProof.Left = cmcUpdate.Left + cmcUpdate.Width + ilSpaceBetweenButtons
    cmcDone.Top = StdPkg.Height - (3 * cmcDone.Height) / 2
    cmcCancel.Top = cmcDone.Top
    cmcUpdate.Top = cmcDone.Top
    cmcProof.Top = cmcDone.Top

    llAdjTop = cmcDone.Top - plcSpec.Top - plcSpec.Height - 120 - tmPkgCtrls(1).fBoxH
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    Do While plcPkg.Top + llAdjTop + 2 * fgBevelY + 240 < cmcDone.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    plcPkg.Height = llAdjTop + 2 * fgBevelY
    pbcPkg.Left = plcPkg.Left + fgBevelX
    pbcPkg.Top = plcPkg.Top + fgBevelY
    pbcPkg.Height = plcPkg.Height - 2 * fgBevelY
    vbcPkg.Left = plcPkg.Width - vbcPkg.Width - fgBevelX - 30
    vbcPkg.Top = fgBevelY
    vbcPkg.Height = pbcPkg.Height
    lacCover.Left = tmPkgCtrls(PKGPRICEINDEX).fBoxX
    lacCover.Top = pbcPkg.Height - lacCover.Height - 15
    lacCover.Width = tmPkgCtrls(PKGPERCENTINDEX).fBoxX + tmPkgCtrls(PKGPERCENTINDEX).fBoxW - tmPkgCtrls(PKGPRICEINDEX).fBoxX
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitPkg                        *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Initialize Rate card Items for  *
'*                     Package                         *
'*                                                     *
'*******************************************************

'**********************************************************
'Rules to use for active Podcast Medium Vehicles (1/20/22)
'**********************************************************
'Rate card Screen:
'**********************************************************
'  Always show active podcast medium vehicles
'  sGMedium = "P"
'**********************************************************
'Std Pkg Screen: (you are here)
'**********************************************************
'  Show podcast medium vehicles when vehicle has programming
'    Vpf.sGMedium = "P" = PodCast
'    LTF_Lbrary_Title WHERE LtfVefCode
'**********************************************************
'CPM Pkg Screen:
'**********************************************************
'  Show podcast medium vehicles when it has an ad server..
'  vendor defined in vehicle options
'    Vpf.sGMedium = "P" (PodCast)
'
'    pvfType="C" =Podcast Ad Server (CPM only)
'    Vff.iAvfCode <> 0 (has Ad Server)
'
'    CpmPkg button visible when sFeatures8=PODADSERVER
'**********************************************************
Private Sub mInitPkg()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slVehName As String
    Dim ilVefCode As Integer
    Dim slDPName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilBypass As Integer
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim ilRdfCode As Integer
    Dim slStr As String
    Dim llUpper As Long
    Dim slVpfMedium As String
    ReDim tmPBDP(0 To 1) As RCPBDPGEN
    ReDim smPkgShow(0 To 10, 0 To UBound(tmPBDP)) As String * 30
    ReDim smPkgSave(0 To 2, 0 To UBound(tmPBDP)) As String * 10
    Dim slSQLQuery As String
    Dim pvf_rst As ADODB.Recordset
    
    For ilLoop = LBound(smPkgShow, 1) To UBound(smPkgShow, 1) Step 1
        For ilIndex = LBound(smPkgShow, 2) To UBound(smPkgShow, 2) Step 1
            smPkgShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    For ilLoop = LBound(smPkgSave, 1) To UBound(smPkgSave, 1) Step 1
        For ilIndex = LBound(smPkgSave, 2) To UBound(smPkgSave, 2) Step 1
            smPkgShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    ilRet = 0
    On Error GoTo mInitPkgErr
    llUpper = LBound(tmRifRec)
    On Error GoTo 0
    If ilRet = 0 Then
        For ilLoop = LBONE To UBound(tmRifRec) - 1 Step 1
            'Vehicle
            gFindMatch Trim$(smRCSave(1, ilLoop)), 0, RateCard!lbcVehicle
            ilIndex = gLastFound(RateCard!lbcVehicle)
            If ilIndex >= 0 Then
                slNameCode = tgRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
                ilRet = gParseItem(slNameCode, 1, "\", slVehName)
                ilRet = gParseItem(slVehName, 3, "|", slVehName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                'Test if Package vehicle- if so, bypass
                ilBypass = False
                slVpfMedium = ""
                ilVef = gBinarySearchVef(ilVefCode)
                If ilVef <> -1 Then
                    If (tgMVef(ilVef).sType = "P") Then 'Package Vehicles
                        If (tgMVef(ilVef).lPvfCode = 0) Then
                            ilBypass = True
                        Else
                            'its a Valid Package, now check PVF
                            '2/23/21 Exclude PVF - pvfType="C" =Podcast Ad Server (CPM only)
                            slSQLQuery = "Select pvfType from PVF_Package_Vehicle Where pvfCode = " & tgMVef(ilVef).lPvfCode
                            Set pvf_rst = gSQLSelectCall(slSQLQuery)
                            If Not pvf_rst.EOF Then
                                If pvf_rst!pvfType = "C" Then
                                    ilBypass = True
                                End If
                            End If
                        End If
                    End If
                    '2/23/21 - Only include podcast vehicle if programming defined (test for ltf).
                    If ilBypass = False Then
                        ilVpf = gBinarySearchVpf(ilVefCode)
                        If ilVpf <> -1 Then
                            If tgVpf(ilVpf).sGMedium = "P" Then 'PodCast
                                'If ((Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER) Then
                                    If gExistLtf(ilVefCode) = False Then 'No Programming
                                        ilBypass = True 'Exclude
                                    End If
                                'End If
                            End If
                        End If
                    End If
                    'Exclude games
                    '6/25/12: Include games
                    If (tgMVef(ilVef).sType = "G") Then
                    '    ilBypass = True
                    End If
                End If
            Else
                ilVefCode = -1
                slVehName = "Missing"
                ilBypass = False
            End If
            If Not ilBypass Then
                'Daypart
                gFindMatch Trim$(smRCSave(2, ilLoop)), 0, RateCard!lbcDPName
                ilIndex = gLastFound(RateCard!lbcDPName)
                If ilIndex >= 0 Then
                    slNameCode = RateCard!lbcDPNameCode.List(ilIndex)
                    ilRet = gParseItem(slNameCode, 1, "\", slDPName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilRdfCode = Val(slCode)
                Else
                    ilRdfCode = -1
                    slDPName = "Missing"
                End If
                'Build record into tmPBDP
                tmPBDP(UBound(tmPBDP)).sSvKey = tmRifRec(ilLoop).sKey
                tmPBDP(UBound(tmPBDP)).iRdfCode = ilRdfCode
                tmPBDP(UBound(tmPBDP)).sVehName = slVehName
                tmPBDP(UBound(tmPBDP)).sDPName = slDPName   'mMakePrgName(ilRdfCode)
                tmPBDP(UBound(tmPBDP)).iVefCode = ilVefCode
                slStr = Trim$(smRCShow(AVGINDEX, ilLoop))
                gUnformatStr slStr, UNFMTDEFAULT, slStr
                tmPBDP(UBound(tmPBDP)).lAvgPrice = gStrDecToLong(slStr, 0)
                tmPBDP(UBound(tmPBDP)).lSvAvgPrice = gStrDecToLong(slStr, 0)
                tmPBDP(UBound(tmPBDP)).lAvgAud = 0
                tmPBDP(UBound(tmPBDP)).iAvgRating = 0
                tmPBDP(UBound(tmPBDP)).lGrImp = 0
                tmPBDP(UBound(tmPBDP)).lGRP = 0
                tmPBDP(UBound(tmPBDP)).lCPP = 0
                tmPBDP(UBound(tmPBDP)).lCPM = 0
                tmPBDP(UBound(tmPBDP)).lPop = 0
                tmPBDP(UBound(tmPBDP)).iVehDormant = imRCSave(9, ilLoop)
                tmPBDP(UBound(tmPBDP)).iDPDormant = imRCSave(10, ilLoop)
                tmPBDP(UBound(tmPBDP)).iPkgVeh = imRCSave(11, ilLoop)
                tmPBDP(UBound(tmPBDP)).sMedium = slVpfMedium
                ReDim Preserve tmPBDP(0 To UBound(tmPBDP) + 1) As RCPBDPGEN
            End If
        Next ilLoop
    End If
    ReDim smPkgShow(0 To 10, 0 To UBound(tmPBDP)) As String * 30
    ReDim smPkgSave(0 To 2, 0 To UBound(tmPBDP)) As String * 10
    For ilLoop = LBound(smPkgShow, 1) To UBound(smPkgShow, 1) Step 1
        For ilIndex = LBound(smPkgShow, 2) To UBound(smPkgShow, 2) Step 1
            smPkgShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    For ilLoop = LBound(smPkgSave, 1) To UBound(smPkgSave, 1) Step 1
        For ilIndex = LBound(smPkgSave, 2) To UBound(smPkgSave, 2) Step 1
            smPkgShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    'mInitShowFields
    imSettingValue = True
    vbcPkg.Min = LBONE  'LBound(tmPBDP)
    imSettingValue = True
    If UBound(tmPBDP) - 1 <= vbcPkg.LargeChange + 1 Then ' + 1 Then
        vbcPkg.Max = LBONE  'LBound(tmPBDP)
    Else
        vbcPkg.Max = UBound(tmPBDP) - vbcPkg.LargeChange
    End If
    imSettingValue = True
    vbcPkg.Value = vbcPkg.Min
    pbcPkg_Paint
    
    On Error Resume Next
    pvf_rst.Close
    
    Exit Sub
mInitPkgErr:
    ilRet = 1
    Resume Next
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
Private Sub mInitShow()
    Dim ilBoxNo As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
    For ilBoxNo = NAMEINDEX To ACT1SETTINGINDEX
        Select Case ilBoxNo 'Branch on box type (control)
            Case NAMEINDEX
                slStr = smSpecSave(1)
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case PRICEINDEX
                Select Case imSpecSave(1)
                    Case 0
                        slStr = "Rate"
                    Case 1
                        slStr = "Audience"
                    Case 2
                        slStr = "Percent"
                    Case 3
                        slStr = "Spot Count"
                    Case Else
                        slStr = ""
                End Select
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case INVTIMEINDEX
                Select Case imSpecSave(2)
                    Case 0
                        slStr = "Real"
                    Case 1
                        slStr = "Virtual"   '"Generate"
                    Case 2
                        slStr = "Equal"   '"Generate"
                    Case Else
                        slStr = ""
                End Select
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case PROGRAMMATICINDEX 'Avail
                Select Case imSpecSave(5)
                    Case 0
                        slStr = "Yes"
                    Case 1
                        slStr = "No"
                    Case Else
                        slStr = ""
                End Select
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case SALESBROCHUREINDEX
                slStr = smSpecSave(5)
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case ALTERNAMEINDEX 'Avail
                Select Case imSpecSave(4)
                    Case 0
                        slStr = "Yes"
                    Case 1
                        slStr = "No"
                    Case Else
                        slStr = ""
                End Select
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case ALTERHIDDENINDEX 'Avail
                Select Case imSpecSave(3)
                    Case 0
                        slStr = "Yes"
                    Case 1
                        slStr = "No"
                    Case 2
                        slStr = "Cmmt/Audio only"
                    Case 3
                        slStr = "Rate only"
                    Case Else
                        slStr = ""
                End Select
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case DEMOINDEX
                slStr = smSpecSave(2)
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case MKTINDEX
                slStr = smSpecSave(4)
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case DOLLARINDEX
                slStr = smSpecSave(3)
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case ACT1CODEINDEX
                slStr = smSpecSave(6)
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case ACT1SETTINGINDEX
                slStr = smSpecSave(7)
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        End Select
    Next ilBoxNo
    For ilLoop = LBONE To UBound(smPkgSave, 2) - 1 Step 1
        slStr = Trim$(smPkgSave(1, ilLoop))
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGSPOTINDEX)
        smPkgShow(PKGSPOTINDEX, ilLoop) = tmPkgCtrls(PKGSPOTINDEX).sShow
        slStr = Trim$(smPkgSave(2, ilLoop))
        gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGPERCENTINDEX)
        smPkgShow(PKGPERCENTINDEX, ilLoop) = tmPkgCtrls(PKGPERCENTINDEX).sShow
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitShowFields                 *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show for rate card  *
'*                      fields                         *
'*                                                     *
'*******************************************************
Private Sub mInitShowFields()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llLoop As Long
    Dim ilVef As Integer
    Dim ilRet As Integer
    
    If UBound(tmPBDP) > 1 Then
        For llLoop = LBound(tmPBDP) To UBound(tmPBDP) - 1 Step 1
            tmPBDP(llLoop) = tmPBDP(llLoop + 1)
        Next llLoop
        ReDim Preserve tmPBDP(0 To UBound(tmPBDP) - 1) As RCPBDPGEN
        ArraySortTyp fnAV(tmPBDP(), 0), UBound(tmPBDP), 0, LenB(tmPBDP(0)), 0, LenB(tmPBDP(0).sKey), 0
        ReDim Preserve tmPBDP(0 To UBound(tmPBDP) + 1) As RCPBDPGEN
        For llLoop = UBound(tmPBDP) - 1 To LBound(tmPBDP) Step -1
            tmPBDP(llLoop + 1) = tmPBDP(llLoop)
        Next llLoop
    End If
    
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        slStr = tmPBDP(ilLoop).sVehName
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGVEHINDEX)
        smPkgShow(PKGVEHINDEX, ilLoop) = tmPkgCtrls(PKGVEHINDEX).sShow
        slStr = tmPBDP(ilLoop).sDPName
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGDPINDEX)
        smPkgShow(PKGDPINDEX, ilLoop) = tmPkgCtrls(PKGDPINDEX).sShow
        slStr = gLongToStrDec(tmPBDP(ilLoop).lAvgPrice, 0)
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGPRICEINDEX)
        smPkgShow(PKGPRICEINDEX, ilLoop) = tmPkgCtrls(PKGPRICEINDEX).sShow
        slStr = ""
        ilVef = gBinarySearchVef(tmPBDP(ilLoop).iVefCode)
        If ilVef <> -1 Then
            slStr = mGetBookName(tgMVef(ilVef).iDnfCode)
        End If
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGBOOKNAMEINDEX)
        smPkgShow(PKGBOOKNAMEINDEX, ilLoop) = tmPkgCtrls(PKGBOOKNAMEINDEX).sShow
        slStr = ""
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGRATINGINDEX)
        smPkgShow(PKGRATINGINDEX, ilLoop) = tmPkgCtrls(PKGRATINGINDEX).sShow
        slStr = ""
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGAUDINDEX)
        smPkgShow(PKGAUDINDEX, ilLoop) = tmPkgCtrls(PKGAUDINDEX).sShow
        slStr = ""
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGCPPINDEX)
        smPkgShow(PKGCPPINDEX, ilLoop) = tmPkgCtrls(PKGCPPINDEX).sShow
        slStr = ""
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGCPMINDEX)
        smPkgShow(PKGCPMINDEX, ilLoop) = tmPkgCtrls(PKGCPMINDEX).sShow
        slStr = ""
        smPkgSave(1, ilLoop) = slStr
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGSPOTINDEX)
        smPkgShow(PKGSPOTINDEX, ilLoop) = tmPkgCtrls(PKGSPOTINDEX).sShow
        slStr = ""
        smPkgSave(2, ilLoop) = slStr
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGPERCENTINDEX)
        smPkgShow(PKGPERCENTINDEX, ilLoop) = tmPkgCtrls(PKGPERCENTINDEX).sShow
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitVef                        *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Inialize vehicle               *
'*                                                     *
'*******************************************************
Private Sub mInitVef()
    gInitVef tmVef
    tmVef.sType = "P"
    tmVef.sState = "A"
    tmVef.sExportRAB = "N"
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
'   mMoveCtrlToRec iTest
'   Where:
'
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim slStr As String
    tmVef.sName = smSpecSave(1)    'Name
    If imSpecSave(1) = 0 Then
        tmVef.sStdPrice = "R"
    ElseIf imSpecSave(1) = 1 Then
        tmVef.sStdPrice = "A"
    ElseIf imSpecSave(1) = 2 Then
        tmVef.sStdPrice = "P"
    ElseIf imSpecSave(1) = 3 Then
        tmVef.sStdPrice = "S"
    Else
        tmVef.sStdPrice = "R"
    End If
    If imSpecSave(2) = 0 Then
        tmVef.sStdInvTime = "A"
    ElseIf imSpecSave(2) = 2 Then
        tmVef.sStdInvTime = "E"
    Else
        tmVef.sStdInvTime = "O"
    End If
    If imSpecSave(3) = 0 Then
        tmVef.sStdAlter = "Y"
    ElseIf imSpecSave(3) = 2 Then
        tmVef.sStdAlter = "C"
    ElseIf imSpecSave(3) = 3 Then
        tmVef.sStdAlter = "R"
    Else
        tmVef.sStdAlter = "N"
    End If
    If imSpecSave(4) = 0 Then
        tmVef.sStdAlterName = "Y"
    Else
        tmVef.sStdAlterName = "N"
    End If
    gFindMatch smSpecSave(2), 1, lbcDemo
    If gLastFound(lbcDemo) > 0 Then
        ilIndex = gLastFound(lbcDemo)
        slNameCode = tgDemoCode(ilIndex - 1).sKey 'lbcAvailCode.List(ilIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveCtrlToRecErr
        gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", StdPkg
        On Error GoTo 0
        slCode = Trim$(slCode)
        tmVef.iMnfDemo = CInt(slCode)
    Else
        tmVef.iMnfDemo = 0
    End If
    imVehGp3ChgMode = True
    tmVef.iMnfVehGp3Mkt = 0
    slStr = smSpecSave(4)
    gFindMatch slStr, 2, lbcVehGp3
    If gLastFound(lbcVehGp3) > 1 Then
        slNameCode = tmVehGp3Code(gLastFound(lbcVehGp3) - 2).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        tmVef.iMnfVehGp3Mkt = Val(slCode)
    End If
    imVehGp3ChgMode = False
    tmVef.iStdIndex = gStrDecToInt(smSpecSave(3), 2)    'Dollar Index
    ReDim tmPvf(0 To 0) As PVF
    ilIndex = LBound(tmPvf(0).iVefCode)
    
    'load tmPvf
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        If Trim$(smPkgSave(1, ilLoop)) <> "" Then
            If Val(Trim$(smPkgSave(1, ilLoop))) > 0 Then
                tmPvf(UBound(tmPvf)).iVefCode(ilIndex) = tmPBDP(ilLoop).iVefCode
                tmPvf(UBound(tmPvf)).iRdfCode(ilIndex) = tmPBDP(ilLoop).iRdfCode
                tmPvf(UBound(tmPvf)).iNoSpot(ilIndex) = Val(Trim$(smPkgSave(1, ilLoop)))
                tmPvf(UBound(tmPvf)).sType = smPcfType
                If imSpecSave(1) = 2 Then
                    tmPvf(UBound(tmPvf)).iPctRate(ilIndex) = gStrDecToLong(Trim$(smPkgSave(2, ilLoop)), 2)
                Else
                    tmPvf(UBound(tmPvf)).iPctRate(ilIndex) = 0
                End If
                ilIndex = ilIndex + 1
                If ilIndex > UBound(tmPvf(0).iVefCode) Then
                    ReDim Preserve tmPvf(0 To UBound(tmPvf) + 1) As PVF
                    ilIndex = LBound(tmPvf(0).iVefCode)
                End If
            End If
        End If
    Next ilLoop
    If ilIndex > LBound(tmPvf(0).iVefCode) Then
        ReDim Preserve tmPvf(0 To UBound(tmPvf) + 1) As PVF
    End If
    Exit Sub
mMoveCtrlToRecErr:
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
Private Sub mMoveRecToCtrl()
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilPvf As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilVff As Integer
    
    smSpecSave(1) = Trim$(tmVef.sName)
    If tmVef.sStdPrice = "R" Then
        imSpecSave(1) = 0
    ElseIf tmVef.sStdPrice = "A" Then
        imSpecSave(1) = 1
    ElseIf tmVef.sStdPrice = "P" Then
        imSpecSave(1) = 2
    ElseIf tmVef.sStdPrice = "S" Then
        imSpecSave(1) = 3
    Else
        imSpecSave(1) = -1
    End If
    If tmVef.sStdInvTime = "A" Then
        imSpecSave(2) = 0
    ElseIf tmVef.sStdInvTime = "O" Then
        imSpecSave(2) = 1
    ElseIf tmVef.sStdInvTime = "E" Then
        imSpecSave(2) = 2
    Else
        imSpecSave(2) = -1
    End If
    If tmVef.sStdAlter = "Y" Then
        imSpecSave(3) = 0
    ElseIf tmVef.sStdAlter = "N" Then
        imSpecSave(3) = 1
    ElseIf tmVef.sStdAlter = "C" Then
        imSpecSave(3) = 2
    ElseIf tmVef.sStdAlter = "R" Then
        imSpecSave(3) = 3
    Else
        imSpecSave(3) = -1
    End If
    If tmVef.sStdAlterName = "Y" Then
        imSpecSave(4) = 0
    ElseIf tmVef.sStdAlterName = "N" Then
        imSpecSave(4) = 1
    Else
        imSpecSave(4) = -1
    End If
    
    'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
    'ACT1 Lineup Code and settings section should probably only be visible/available when ACT1 Enabled via Site Setting and "vef Alter Hidden" = ("n" or = "c")  and "vef alter name" = "n"
    bmAct1Allowed = False
    If (tmVef.sStdAlter = "N" Or tmVef.sStdAlter = "C") And tmVef.sStdAlterName = "N" And bmAct1Enabled Then
        bmAct1Allowed = True
    End If
    
    smSpecSave(2) = ""
    lbcDemo.ListIndex = -1
    slRecCode = Trim$(Str$(tmVef.iMnfDemo))
    For ilTest = 0 To UBound(tgDemoCode) - 1 Step 1 'lbcAvailCode.ListCount - 1 Step 1
        slNameCode = tgDemoCode(ilTest).sKey  'lbcAvailCode.List(ilTest)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", StdPkg
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcDemo.ListIndex = ilTest + 1
            smSpecSave(2) = lbcDemo.List(ilTest + 1)
            Exit For
        End If
    Next ilTest
    smSpecSave(2) = Trim$(smSpecSave(2))
    smOrigDemo = smSpecSave(2)
    mVehGp3Pop
    smSpecSave(4) = ""
    If tmVef.iMnfVehGp3Mkt > 0 Then
        For ilVef = 0 To UBound(tmVehGp3Code) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
            slNameCode = tmVehGp3Code(ilVef).sKey   'lbcVehGpCode.List(ilVef)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmVef.iMnfVehGp3Mkt Then
                lbcVehGp3.ListIndex = ilVef + 2
                smSpecSave(4) = lbcVehGp3.List(ilVef + 2)
                Exit For
            End If
        Next ilVef
    End If
    smSpecSave(4) = Trim$(smSpecSave(4))
    smOrigVehGp3 = smSpecSave(4)
    smSpecSave(3) = gIntToStrDec(tmVef.iStdIndex, 2)
    For ilTest = LBONE To UBound(tmPBDP) - 1 Step 1
        tmPBDP(ilTest).sKey = "1" & tmPBDP(ilTest).sSvKey
    Next ilTest
    For ilPvf = LBound(tmPvf) To UBound(tmPvf) - 1 Step 1
        For ilLoop = LBound(tmPvf(ilPvf).iVefCode) To UBound(tmPvf(ilPvf).iVefCode) Step 1
            If tmPvf(ilPvf).iVefCode(ilLoop) > 0 Then
                For ilTest = LBONE To UBound(tmPBDP) - 1 Step 1
                    If (tmPvf(ilPvf).iVefCode(ilLoop) = tmPBDP(ilTest).iVefCode) And (tmPvf(ilPvf).iRdfCode(ilLoop) = tmPBDP(ilTest).iRdfCode) Then
                        If tmPvf(ilPvf).iNoSpot(ilLoop) > 0 Then
                            tmPBDP(ilTest).sKey = "0" & tmPBDP(ilTest).sSvKey
                        End If
                        Exit For
                    End If
                Next ilTest
            End If
        Next ilLoop
    Next ilPvf
    mAdjPrices
    smOrigProgrammaticAllowed = ""
    smOrigSalesBrochure = ""
    ilVff = gBinarySearchVff(tmVef.iCode)
    If ilVff <> -1 Then
        If tgVff(ilVff).sPrgmmaticAllow = "Y" Then
            imSpecSave(5) = 0
            smOrigProgrammaticAllowed = "Y"
        Else
            imSpecSave(5) = 1
            smOrigProgrammaticAllowed = "N"
        End If
        smSpecSave(5) = Trim$(tgVff(ilVff).sSalesBrochure)
    Else
        imSpecSave(5) = -1
        smSpecSave(5) = ""
    End If
    smOrigSalesBrochure = smSpecSave(5)
    
    mInitShowFields
    For ilPvf = LBound(tmPvf) To UBound(tmPvf) - 1 Step 1
        For ilLoop = LBound(tmPvf(ilPvf).iVefCode) To UBound(tmPvf(ilPvf).iVefCode) Step 1
            If tmPvf(ilPvf).iVefCode(ilLoop) > 0 Then
                For ilTest = LBONE To UBound(tmPBDP) - 1 Step 1
                    If (tmPvf(ilPvf).iVefCode(ilLoop) = tmPBDP(ilTest).iVefCode) And (tmPvf(ilPvf).iRdfCode(ilLoop) = tmPBDP(ilTest).iRdfCode) Then
                        smPkgSave(1, ilTest) = Trim$(Str$(tmPvf(ilPvf).iNoSpot(ilLoop)))
                        If imSpecSave(1) = 2 Then
                            smPkgSave(2, ilTest) = gIntToStrDec(tmPvf(ilPvf).iPctRate(ilLoop), 2)
                        Else
                            smPkgSave(2, ilTest) = ""
                        End If
                        smSpecSave(6) = Trim(tmPvf(ilPvf).sACT1LineupCode)
                        smSpecSave(7) = ""
                        If tmPvf(ilPvf).sACT1StoredTime = "T" Then smSpecSave(7) = smSpecSave(7) & "T"
                        If tmPvf(ilPvf).sACT1StoredSpots = "S" Then smSpecSave(7) = smSpecSave(7) & "S"
                        If tmPvf(ilPvf).sACT1StoreClearPct = "C" Then smSpecSave(7) = smSpecSave(7) & "C"
                        If tmPvf(ilPvf).sACT1DaypartFilter = "F" Then smSpecSave(7) = smSpecSave(7) & "F"
                        Exit For
                    End If
                Next ilTest
            End If
        Next ilLoop
    Next ilPvf

    Exit Sub
mMoveRecToCtrlErr:
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
Private Function mOKName()
    Dim slStr As String
    Dim ilVef As Integer
    If smSpecSave(1) <> "" Then    'Test name
        slStr = Trim$(smSpecSave(1))
        'gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If StrComp(slStr, Trim$(tgMVef(ilVef).sName), 1) = 0 Then
                If (imSelectedIndex = 0) Or (tgMVef(ilVef).iCode <> tmVef.iCode) Then
                    Beep
                    MsgBox "Standard Package Name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    imSpecBoxNo = NAMEINDEX
                    mSpecEnableBox imSpecBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        Next ilVef
    End If
    mOKName = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mPkgEnableBox                   *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mPkgEnableBox(ilBoxNo As Integer)
'
'   mPkgEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBPkgCtrls Or ilBoxNo > UBound(tmPkgCtrls) Then
        Exit Sub
    End If
    If (imPkgRowNo < vbcPkg.Value) Or (imPkgRowNo >= vbcPkg.Value + vbcPkg.LargeChange + 1) Then
        pbcArrow.Visible = False
        lacPkgFrame.Visible = False
        Exit Sub
    End If
    lacPkgFrame.Move 0, tmPkgCtrls(PKGVEHINDEX).fBoxY + (imPkgRowNo - vbcPkg.Value) * (fgBoxGridH + 15) - 30
    lacPkgFrame.Visible = True

    pbcArrow.Visible = False
    pbcArrow.Move plcPkg.Left - pbcArrow.Width - 15, plcPkg.Top + tmPkgCtrls(PKGVEHINDEX).fBoxY + (imPkgRowNo - vbcPkg.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case PKGSPOTINDEX 'Start/End
            edcDropDown.Width = tmPkgCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 3
            gMoveTableCtrl pbcPkg, edcDropDown, tmPkgCtrls(PKGSPOTINDEX).fBoxX, tmPkgCtrls(PKGSPOTINDEX).fBoxY + (imPkgRowNo - vbcPkg.Value) * (fgBoxGridH + 15)
            edcDropDown.Text = Trim$(smPkgSave(1, imPkgRowNo))
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case PKGPERCENTINDEX 'Start/End
            edcDropDown.Width = tmPkgCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 7
            gMoveTableCtrl pbcPkg, edcDropDown, tmPkgCtrls(PKGPERCENTINDEX).fBoxX, tmPkgCtrls(PKGPERCENTINDEX).fBoxY + (imPkgRowNo - vbcPkg.Value) * (fgBoxGridH + 15)
            edcDropDown.Text = Trim$(smPkgSave(2, imPkgRowNo))
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPkgSetFocus                    *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mPkgSetFocus(ilBoxNo As Integer)
'
'   mPkgSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBPkgCtrls Or ilBoxNo > UBound(tmPkgCtrls) Then
        Exit Sub
    End If
    If (imPkgRowNo < vbcPkg.Value) Or (imPkgRowNo >= vbcPkg.Value + vbcPkg.LargeChange + 1) Then
        pbcArrow.Visible = False
        lacPkgFrame.Visible = False
        Exit Sub
    End If

    pbcArrow.Visible = False
    pbcArrow.Move plcPkg.Left - pbcArrow.Width - 15, plcPkg.Top + tmPkgCtrls(1).fBoxY + (imPkgRowNo - 1) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case PKGSPOTINDEX 'Start/End
            If edcDropDown.Visible = True Then edcDropDown.SetFocus
        Case PKGPERCENTINDEX 'Start/End
            edcDropDown.SetFocus
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPkgSetShow                      *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mPkgSetShow(ilBoxNo As Integer)
'
'   mPkgSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    pbcArrow.Visible = False
    lacPkgFrame.Visible = False
    If ilBoxNo < imLBPkgCtrls Or ilBoxNo > UBound(tmPkgCtrls) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case PKGSPOTINDEX 'Vehicle
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            gSetShow pbcPkg, slStr, tmPkgCtrls(ilBoxNo)
            smPkgShow(PKGSPOTINDEX, imPkgRowNo) = tmPkgCtrls(ilBoxNo).sShow
            If Trim$(smPkgSave(1, imPkgRowNo)) <> edcDropDown.Text Then
                imPkgChg = True
                smPkgSave(1, imPkgRowNo) = edcDropDown.Text
                mGetPkgAud imPkgRowNo
                mGetTotals
            End If
        Case PKGPERCENTINDEX 'Vehicle
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            gSetShow pbcPkg, slStr, tmPkgCtrls(ilBoxNo)
            smPkgShow(PKGPERCENTINDEX, imPkgRowNo) = tmPkgCtrls(ilBoxNo).sShow
            If Trim$(smPkgSave(2, imPkgRowNo)) <> edcDropDown.Text Then
                imPkgChg = True
                smPkgSave(2, imPkgRowNo) = edcDropDown.Text
                mGetPkgAud imPkgRowNo
                mGetTotals
            End If
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(StdPkg, VEHSTDPKG + ACTIVEVEH + DORMANTVEH, cbcSelect, tmPkgVehicle(), smPkgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox: StdPkg)", StdPkg
        On Error GoTo 0
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
        imPopReqd = True
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer)
'
'   iRet = mReadRec()
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    Dim llPvfCode As Long
    ilRet = 0
    slNameCode = tmPkgVehicle(ilSelectIndex - 1).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", StdPkg
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmVefSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual-VEF)", StdPkg
    On Error GoTo 0
    llPvfCode = tmVef.lPvfCode
    ReDim tmPvf(0 To 0) As PVF
    Do While llPvfCode > 0
        tmPvfSrchKey.lCode = llPvfCode
        ilRet = btrGetEqual(hmPvf, tmPvf(UBound(tmPvf)), imPvfRecLen, tmPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mReadRecErr
        'JW 1/21/22-KB 336 - Std. Package Screen: Warning "Package has been damaged" when PVF BTRV_ERR_KEY_NOT_FOUND
        If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
            MsgBox "This Package has been damaged!" & vbCrLf & "The package vehicles could not be loaded..." & vbCrLf & "Contact Couterpoint Support to help recover the package.", vbCritical + vbOKOnly, "Warning"
        Else
            gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", StdPkg
        End If
        On Error GoTo 0
        llPvfCode = tmPvf(UBound(tmPvf)).lLkPvfCode
        ReDim Preserve tmPvf(0 To UBound(tmPvf) + 1) As PVF
    Loop
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
'*             Created:6/04/93       By:D. LeVine      *
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
    Dim slStamp As String   'Date/Time stamp for file
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim llPvfCode As Long
    Dim ilLen As Integer
    Dim ilPvf As Integer
    Dim ilVef As Integer
    Dim ilLenTest As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim ilFirstFd As Integer
    Dim ilPvf1 As Integer
    Dim ilVef1 As Integer
    Dim ilLoop1 As Integer
    Dim ilVpf As Integer
    Dim slMedium As String
    
    'For AT + PodcastSpot Mix Testing
    Dim liATCount
    Dim liSpotCount
    Dim liFound
    mSpecSetShow imSpecBoxNo
    mPkgSetShow imPkgBoxNo
    If mTestSaveFields(SHOWMSG, True) = NO Then
        mSaveRec = False
        Exit Function
    End If
    ilFound = False
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        If Trim$(smPkgSave(1, ilLoop)) <> "" Then
            If Val(Trim$(smPkgSave(1, ilLoop))) > 0 Then
                ilFound = True
                Exit For
            End If
        End If
    Next ilLoop
    If Not ilFound Then
        ilRet = MsgBox("At least one Vehicle must be associated with the Package", vbOKOnly + vbExclamation, "Incomplete")
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    gGetSyncDateTime slSyncDate, slSyncTime
    ilRet = btrBeginTrans(hmVef, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault    'Default
        ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "VEF.Btr")
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                GoTo mSaveRecErr
            End If
        Else
            mInitVef
        End If
        '-----------------------------------
        'Check if We're trying to save a Mixed set of Vehicles: (AT and Spot)
        'mixed AT + POD Spots NOT ALLOWED
        liATCount = 0
        liSpotCount = 0
        bmMixedFound = False
        'get the Vehicle's Medium to determine if this is a AT or PodSpot vehicle.
        For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
            If Trim$(smPkgSave(1, ilLoop)) <> "" Then
                If Val(Trim$(smPkgSave(1, ilLoop))) > 0 Then
                    If tmPBDP(ilLoop).iVefCode > 0 Then
                        ilVpf = gBinarySearchVpf(tmPBDP(ilLoop).iVefCode)
                        If ilVpf <> -1 Then
                            slMedium = tgVpf(ilVpf).sGMedium
                            
                        End If
                        If slMedium = "P" Then
                            'PodSpot
                            smPcfType = "S"
                            liSpotCount = liSpotCount + 1
                        Else
                            'AirTime
                            smPcfType = "A"
                            liATCount = liATCount + 1
                        End If
                        If liSpotCount > 0 And liATCount > 0 Then
                            smPcfType = "M"
                            bmMixedFound = True
                            Exit For
                        End If
                    End If
                End If
            End If
        Next ilLoop
        '2/25/21: Mix question changed to Ad Server Tab View Only.
        '         Air time and Podcast spot allowed within the same package
        'If (Asc(tgSaf(0).sFeatures8) And PODADSERVERVIEWONLY) <> PODADSERVERVIEWONLY Then
        '    If bmMixedFound = True Then
        '        Screen.MousePointer = vbDefault    'Default
        '        ilRet = MsgBox("Mixing Air and Podcast Spot Vehicles is not allowed per Site Rule: 'Mix Airtime and Podcast'", vbOKOnly + vbExclamation, "Incomplete")
        '        mSaveRec = False
        '        Exit Function
        '    End If
        'End If
        
        '-----------------------------------
        'Delete PVF, then Create PVF
        If imSelectedIndex <> 0 Then 'New selected
            For ilLoop = LBound(tmPvf) To UBound(tmPvf) - 1 Step 1
                Do
                    tmPvfSrchKey.lCode = tmPvf(ilLoop).lCode
                    ilRet = btrGetEqual(hmPvf, tmTPvf, imPvfRecLen, tmPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilCRet = btrAbortTrans(hmPvf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
                        mSaveRec = False
                        Exit Function
                    End If
                    ilRet = btrDelete(hmPvf)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmPvf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
                    mSaveRec = False
                    Exit Function
                End If
            Next ilLoop
        End If
        ''-----------------------------------
        ''Get input spot /% values
        mMoveCtrlToRec
        llPvfCode = 0

        '-----------------------------------
        'Save
        For ilLoop = UBound(tmPvf) - 1 To LBound(tmPvf) Step -1
            tmPvf(ilLoop).lCode = 0
            tmPvf(ilLoop).sName = smSpecSave(1)
            'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
            tmPvf(ilLoop).sACT1LineupCode = smSpecSave(6)
            If InStr(1, smSpecSave(7), "T") > 0 Then
                tmPvf(ilLoop).sACT1StoredTime = "T"
            Else
                tmPvf(ilLoop).sACT1StoredTime = ""
            End If
            If InStr(1, smSpecSave(7), "S") > 0 Then
                tmPvf(ilLoop).sACT1StoredSpots = "S"
            Else
                tmPvf(ilLoop).sACT1StoredSpots = ""
            End If
            If InStr(1, smSpecSave(7), "C") > 0 Then
                tmPvf(ilLoop).sACT1StoreClearPct = "C"
            Else
                tmPvf(ilLoop).sACT1StoreClearPct = ""
            End If
            If InStr(1, smSpecSave(7), "F") > 0 Then
                tmPvf(ilLoop).sACT1DaypartFilter = "F"
            Else
                tmPvf(ilLoop).sACT1DaypartFilter = ""
            End If
            tmPvf(ilLoop).lLkPvfCode = llPvfCode
            ilRet = btrInsert(hmPvf, tmPvf(ilLoop), imPvfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmPvf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
                mSaveRec = False
                Exit Function
            End If
            llPvfCode = tmPvf(ilLoop).lCode
        Next ilLoop
        If imSelectedIndex = 0 Then 'New selected
            imAddPkg = True
            tmVef.iCode = 0  'Autoincrement
            tmVef.lPvfCode = llPvfCode
            tmVef.iRemoteID = tgUrf(0).iRemoteUserID
            tmVef.iAutoCode = tmVef.iCode
            ilRet = btrInsert(hmVef, tmVef, imVefRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert:StdPkg)"
        Else 'Old record-Update
            tmVef.lPvfCode = llPvfCode
            'tmVef.iSourceID = tgUrf(0).iRemoteUserID
            'gPackDate slSyncDate, tmVef.iSyncDate(0), tmVef.iSyncDate(1)
            'gPackTime slSyncTime, tmVef.iSyncTime(0), tmVef.iSyncTime(1)
            ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
            slMsg = "mSaveRec (btrUpdate:StdPkg)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        ilCRet = btrAbortTrans(hmPvf)
        Screen.MousePointer = vbDefault    'Default
        ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
        mSaveRec = False
        Exit Function
    End If
    If imSelectedIndex = 0 Then 'New selected
        Do
            tmVef.iRemoteID = tgUrf(0).iRemoteUserID
            tmVef.iAutoCode = tmVef.iCode
            ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
            slMsg = "mSaveRec (btrUpdate:StdPkg)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmPvf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
            mSaveRec = False
            Exit Function
        End If
        sgMVefStamp = ""
        ilRet = gObtainVef()
    Else
        ilRet = gBinarySearchVef(tmVef.iCode)
        If ilRet <> -1 Then
            tgMVef(ilRet) = tmVef
        End If
    End If
    ilRet = gVpfFind(StdPkg, tmVef.iCode)
    'Update lengths
    ilFirstFd = False
    tmVpfSrchKey.iVefKCode = tmVef.iCode
    ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        If smPVFType = "S" Then
            tmVpf.sGMedium = "P" 'Podcast Spots
        Else
            tmVpf.sGMedium = "N" 'Radio Net (Mix or AT Vehicles)
        End If
        For ilLen = LBound(tmVpf.iSLen) To UBound(tmVpf.iSLen) Step 1
            tmVpf.iSLen(ilLen) = 0
        Next ilLen
        ilIndex = LBound(tmVpf.iSLen)
        For ilPvf = LBound(tmPvf) To UBound(tmPvf) - 1 Step 1
            For ilVef = LBound(tmPvf(ilPvf).iVefCode) To UBound(tmPvf(ilPvf).iVefCode) Step 1
                ilLoop = gBinarySearchVpf(tmPvf(ilPvf).iVefCode(ilVef))
                If ilLoop <> -1 Then
                    ilFirstFd = True
                    For ilLen = LBound(tgVpf(ilLoop).iSLen) To UBound(tgVpf(ilLoop).iSLen) Step 1
                        If tgVpf(ilLoop).iSLen(ilLen) <> 0 Then
                            ilFound = False
                            For ilLenTest = LBound(tmVpf.iSLen) To UBound(tmVpf.iSLen) Step 1
                                If tmVpf.iSLen(ilLenTest) = tgVpf(ilLoop).iSLen(ilLen) Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLenTest
                            If Not ilFound Then
                                'Test if in all other vehicles- if not don't add
                                ilFound = True
                                For ilPvf1 = LBound(tmPvf) To UBound(tmPvf) - 1 Step 1
                                    For ilVef1 = LBound(tmPvf(ilPvf1).iVefCode) To UBound(tmPvf(ilPvf1).iVefCode) Step 1
                                        ilLoop1 = gBinarySearchVpf(tmPvf(ilPvf1).iVefCode(ilVef1))
                                        If ilLoop1 <> -1 Then
                                            ilFound = False
                                            For ilLenTest = LBound(tmVpf.iSLen) To UBound(tmVpf.iSLen) Step 1
                                                If tgVpf(ilLoop).iSLen(ilLen) = tgVpf(ilLoop1).iSLen(ilLenTest) Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilLenTest
                                        End If
                                        If Not ilFound Then
                                            Exit For
                                        End If
                                    Next ilVef1
                                    If Not ilFound Then
                                        Exit For
                                    End If
                                Next ilPvf1
                                If Not ilFound Then
                                    ilFound = True
                                Else
                                    ilFound = False
                                End If
                            End If
                            If Not ilFound Then
                                tmVpf.iSLen(ilIndex) = tgVpf(ilLoop).iSLen(ilLen)
                                If ilIndex = LBound(tmVpf.iSLen) Then
                                    tmVpf.iSDLen = tgVpf(ilLoop).iSDLen
                                End If
                                ilIndex = ilIndex + 1
                                If ilIndex > UBound(tmVpf.iSLen) Then
                                    Exit For
                                End If
                            End If
                        End If
                    Next ilLen
                End If
                
                If ilFirstFd Then
                    Exit For
                End If
            Next ilVef
            
            If ilFirstFd Then
                Exit For
            End If
        Next ilPvf
        ilRet = btrUpdate(hmVpf, tmVpf, imVpfRecLen)
        ilRet = gBinarySearchVpf(tmVpf.iVefKCode)
        If ilRet <> -1 Then
            tgVpf(ilRet) = tmVpf
        End If
    End If
    
    ilFirstFd = False
    ilRet = mVffReadRec(tmVef.iCode)
    If ilRet Then
        If bmProgrammaticAllowed Then
            If imSpecSave(5) = 0 Then
                tmVff.sPrgmmaticAllow = "Y"
            Else
                tmVff.sPrgmmaticAllow = "N"
            End If
            If smSpecSave(5) <> "[None]" Then
                tmVff.sSalesBrochure = smSpecSave(5)
            Else
                tmVff.sSalesBrochure = ""
            End If
        Else
            tmVff.sPrgmmaticAllow = "N"
            tmVff.sSalesBrochure = ""
        End If
        ilRet = btrUpdate(hmVff, tmVff, imVffRecLen)
        sgVffStamp = "~"
        ilRet = gVffRead()
    End If
    'Update lengths
    ilRet = btrEndTrans(hmPvf)
    gFileChgdUpdate "vef.btr", False
    gFileChgdUpdate "vpf.btr", False

    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function
mSaveRecErr:
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
'*             Created:6/04/93       By:D. LeVine      *
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
    ilAltered = gAnyFieldChgd(tmSpecCtrls(), TESTALLCTRLS)
    If mTestSaveFields(NOMSG, True) = YES Then  'No Then
        If (ilAltered = YES) Or (imPkgChg = True) Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & smSpecSave(1)
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
                If ilRes = vbNo Then
                    cbcSelect.ListIndex = 0
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
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
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
    Dim ilAltered As Integer
    Dim ilLoop As Integer
    
    If imBypassSetting Then
        Exit Sub
    End If
    ilAltered = gAnyFieldChgd(tmSpecCtrls(), TESTALLCTRLS)
    'Update button set if all mandatory fields have data and any field altered
    If (mTestSaveFields(NOMSG, False) = YES) And ((ilAltered = YES) Or (imPkgChg = True)) Then
        If imUpdateAllowed Then
            cmcUpdate.Enabled = True
        Else
            cmcUpdate.Enabled = False
        End If
    Else
        cmcUpdate.Enabled = False
    End If
    If Not ilAltered Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
    cmcProof.Enabled = False
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        If Trim$(smPkgSave(1, ilLoop)) <> "" Then
            If Val(Trim$(smPkgSave(1, ilLoop))) > 0 Then
                cmcProof.Enabled = True
                Exit For
            End If
        End If
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecEnableBox                  *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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
    Dim slStr As String
    If (ilBoxNo < imLBSpecCtrls) Or (ilBoxNo > UBound(tmSpecCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX
            edcSpecDropDown.Width = tmSpecCtrls(ilBoxNo).fBoxW
            If tgSpf.iVehLen <= 40 Then
                edcSpecDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcSpecDropDown.MaxLength = 20
            End If
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            imChgMode = True
            edcSpecDropDown.Text = smSpecSave(1)
            imChgMode = False
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case PRICEINDEX
            If imSpecSave(1) < 0 Then
                imSpecSave(1) = 0    'Yes
                mSpecSetChg imSpecBoxNo, True
            End If
            pbcPrice.Width = tmSpecCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSpec, pbcPrice, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            pbcPrice_Paint
            pbcPrice.Visible = True
            pbcPrice.SetFocus
        Case INVTIMEINDEX
            If imSpecSave(2) < 0 Then
                If tgSpf.sCPkAired = "Y" Then
                    imSpecSave(2) = 0    'Yes
                ElseIf tgSpf.sCPkOrdered = "Y" Then
                    imSpecSave(2) = 1
                ElseIf tgSpf.sCPkEqual = "Y" Then
                    imSpecSave(2) = 2
                End If
                tmSpecCtrls(INVTIMEINDEX).iChg = True
                mSpecSetChg imSpecBoxNo, True
            End If
            pbcInvTime.Width = tmSpecCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSpec, pbcInvTime, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            pbcInvTime_Paint
            pbcInvTime.Visible = True
            pbcInvTime.SetFocus
        Case PROGRAMMATICINDEX
            If imSpecSave(5) < 0 Then
                imSpecSave(5) = 0    'Yes
                tmSpecCtrls(PROGRAMMATICINDEX).iChg = True
                mSpecSetChg imSpecBoxNo, True
            End If
            pbcProgrammatic.Width = tmSpecCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSpec, pbcProgrammatic, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            pbcProgrammatic_Paint
            pbcProgrammatic.Visible = True
            pbcProgrammatic.SetFocus
        Case SALESBROCHUREINDEX
            lbcPDFName.Height = gListBoxHeight(lbcPDFName.ListCount, 10)
            edcSpecDropDown.Width = 3 * tmSpecCtrls(ilBoxNo).fBoxW - cmcSpecDropDown.Width
            edcSpecDropDown.MaxLength = 20
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            lbcPDFName.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.Height
            lbcPDFName.Width = edcSpecDropDown.Width + cmcSpecDropDown.Width
            imChgMode = True
            gFindMatch smSpecSave(5), 0, lbcPDFName
            If gLastFound(lbcPDFName) >= 0 Then
                lbcPDFName.ListIndex = gLastFound(lbcPDFName)
                edcSpecDropDown.Text = lbcPDFName.List(lbcPDFName.ListIndex)
            Else
                If lbcPDFName.ListCount > 1 Then
                    lbcPDFName.ListIndex = 0
                    edcSpecDropDown.Text = lbcPDFName.List(lbcPDFName.ListIndex)
                Else
                    lbcPDFName.ListIndex = -1
                    edcSpecDropDown.Text = ""
                End If
            End If
            imChgMode = False
            imComboBoxIndex = lbcPDFName.ListIndex
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case ALTERNAMEINDEX
            If imSpecSave(4) < 0 Then
                imSpecSave(4) = 0    'Yes
                tmSpecCtrls(ALTERNAMEINDEX).iChg = True
                mSpecSetChg imSpecBoxNo, True
            End If
            pbcAlter.Width = tmSpecCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSpec, pbcAlter, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            pbcAlter_Paint
            pbcAlter.Visible = True
            pbcAlter.SetFocus
        Case ALTERHIDDENINDEX
            If imSpecSave(3) < 0 Then
                imSpecSave(3) = 0    'Yes
                tmSpecCtrls(ALTERHIDDENINDEX).iChg = True
                mSpecSetChg imSpecBoxNo, True
            End If
            pbcAlter.Width = tmSpecCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSpec, pbcAlter, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            pbcAlter_Paint
            pbcAlter.Visible = True
            pbcAlter.SetFocus
        Case DEMOINDEX
            lbcDemo.Height = gListBoxHeight(lbcDemo.ListCount, 10)
            edcSpecDropDown.Width = tmSpecCtrls(ilBoxNo).fBoxW - cmcSpecDropDown.Width
            edcSpecDropDown.MaxLength = 20
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            lbcDemo.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.Height
            imChgMode = True
            gFindMatch smSpecSave(2), 0, lbcDemo
            If gLastFound(lbcDemo) >= 0 Then
                lbcDemo.ListIndex = gLastFound(lbcDemo)
                edcSpecDropDown.Text = lbcDemo.List(lbcDemo.ListIndex)
            Else
                If lbcDemo.ListCount > 1 Then
                    lbcDemo.ListIndex = 0
                    edcSpecDropDown.Text = lbcDemo.List(lbcDemo.ListIndex)
                Else
                    lbcDemo.ListIndex = -1
                    edcSpecDropDown.Text = ""
                End If
            End If
            imChgMode = False
            imComboBoxIndex = lbcDemo.ListIndex
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case MKTINDEX
            mVehGp3Pop
            If imTerminate Then
                Exit Sub
            End If
            lbcVehGp3.Height = gListBoxHeight(lbcVehGp3.ListCount, 10)
            edcSpecDropDown.Width = tmSpecCtrls(ilBoxNo).fBoxW - cmcSpecDropDown.Width
            edcSpecDropDown.MaxLength = 20
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            imVehGp3ChgMode = True
            slStr = smSpecSave(4)
            gFindMatch slStr, 1, lbcVehGp3
            If gLastFound(lbcVehGp3) >= 1 Then
                lbcVehGp3.ListIndex = gLastFound(lbcVehGp3)
                edcSpecDropDown.Text = lbcVehGp3.List(lbcVehGp3.ListIndex)
            Else
                If lbcVehGp3.ListCount > 1 Then
                    lbcVehGp3.ListIndex = 1
                    edcSpecDropDown.Text = lbcVehGp3.List(1)
                Else
                    lbcVehGp3.ListIndex = 0
                    edcSpecDropDown.Text = lbcVehGp3.List(0)
                End If
            End If
            imVehGp3ChgMode = False
            lbcVehGp3.Move edcSpecDropDown.Left + edcSpecDropDown.Width + cmcSpecDropDown.Width - lbcVehGp3.Width, edcSpecDropDown.Top + edcSpecDropDown.Height
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case DOLLARINDEX
            edcSpecDropDown.Width = tmSpecCtrls(ilBoxNo).fBoxW
            edcSpecDropDown.MaxLength = 5
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            imChgMode = True
            If smSpecSave(3) = "" Then
                edcSpecDropDown.Text = "1.00"
            Else
                edcSpecDropDown.Text = smSpecSave(3)
            End If
            imChgMode = False
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
            
        'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
        Case ACT1CODEINDEX
            edcSpecDropDown.Width = tmSpecCtrls(ilBoxNo).fBoxW
            edcSpecDropDown.MaxLength = 11
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            imChgMode = True
            If smSpecSave(6) = "" Then
                edcSpecDropDown.Text = ""
            Else
                edcSpecDropDown.Text = smSpecSave(6)
            End If
            imChgMode = False
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
            
        Case ACT1SETTINGINDEX
            plcACT1Settings.Visible = True
            edcDropDown.MaxLength = 4
            gMoveTableCtrl pbcSpec, plcACT1Settings, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            edcSpecDropDown.Text = Trim$(smSpecSave(7))
            If InStr(1, edcSpecDropDown.Text, "T") > 0 Then
                edcACT1SettingT.Text = "Yes"
            Else
                edcACT1SettingT.Text = "No"
            End If
            If InStr(1, edcSpecDropDown.Text, "S") > 0 Then
                edcACT1SettingS.Text = "Yes"
            Else
                edcACT1SettingS.Text = "No"
            End If
            If InStr(1, edcSpecDropDown.Text, "C") > 0 Then
                edcACT1SettingC.Text = "Yes"
            Else
                edcACT1SettingC.Text = "No"
            End If
            If InStr(1, edcSpecDropDown.Text, "F") > 0 Then
                edcACT1SettingF.Text = "Yes"
            Else
                edcACT1SettingF.Text = "No"
            End If
            edcACT1SettingT.SetFocus
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecSetChg                     *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mSpecSetChg(ilBoxNo As Integer, ilUseSave As Integer)
'
'   mSpecSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim slStr As String
    Dim slInitStr As String
    If ilBoxNo < imLBSpecCtrls Or (ilBoxNo > UBound(tmSpecCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            slStr = edcSpecDropDown.Text
            gSetChgFlagStr tmVef.sName, slStr, tmSpecCtrls(ilBoxNo)
        Case PRICEINDEX
            Select Case imSpecSave(1)
                Case 0
                    slStr = "R"
                Case 1
                    slStr = "A"
                Case 2
                    slStr = "P"
                Case 3
                    slStr = "S"
                Case Else
                    slStr = ""
            End Select
            gSetChgFlagStr tmVef.sStdPrice, slStr, tmSpecCtrls(ilBoxNo)
        Case INVTIMEINDEX
            Select Case imSpecSave(2)
                Case 0
                    slStr = "A"
                Case 1
                    slStr = "O"
                Case 2
                    slStr = "E"
                Case Else
                    slStr = ""
            End Select
            gSetChgFlagStr tmVef.sStdInvTime, slStr, tmSpecCtrls(ilBoxNo)
        Case PROGRAMMATICINDEX
            Select Case imSpecSave(5)
                Case 0
                    slStr = "Y"
                Case 1
                    slStr = "Y"
                Case Else
                    slStr = ""
            End Select
            gSetChgFlagStr smOrigProgrammaticAllowed, slStr, tmSpecCtrls(ilBoxNo)
        Case SALESBROCHUREINDEX
            gSetChgFlag smOrigSalesBrochure, lbcPDFName, tmSpecCtrls(ilBoxNo)
        Case ALTERNAMEINDEX
            Select Case imSpecSave(4)
                Case 0
                    slStr = "Y"
                Case 1
                    slStr = "N"
                Case Else
                    slStr = ""
            End Select
            gSetChgFlagStr tmVef.sStdAlterName, slStr, tmSpecCtrls(ilBoxNo)
        Case ALTERHIDDENINDEX
            Select Case imSpecSave(3)
                Case 0
                    slStr = "Y"
                Case 1
                    slStr = "N"
                Case 2
                    slStr = "C"
                Case 3
                    slStr = "R"
                Case Else
                    slStr = ""
            End Select
            gSetChgFlagStr tmVef.sStdAlter, slStr, tmSpecCtrls(ilBoxNo)
        Case DEMOINDEX
            gSetChgFlag smOrigDemo, lbcDemo, tmSpecCtrls(ilBoxNo)
        Case MKTINDEX
            gSetChgFlag smOrigVehGp3, lbcVehGp3, tmSpecCtrls(ilBoxNo)
        Case DOLLARINDEX 'Name
            slStr = edcSpecDropDown.Text
            slInitStr = gIntToStrDec(tmVef.iStdIndex, 2)
            gSetChgFlagStr slInitStr, slStr, tmSpecCtrls(ilBoxNo)
        
        'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
        Case ACT1CODEINDEX
            slStr = edcSpecDropDown.Text
            gSetChgFlagStr slInitStr, slStr, tmSpecCtrls(ilBoxNo)

        Case ACT1SETTINGINDEX
            slStr = edcSpecDropDown.Text
            gSetChgFlagStr slInitStr, slStr, tmSpecCtrls(ilBoxNo)

    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecSetFocus                   *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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
    If (ilBoxNo < imLBSpecCtrls) Or (ilBoxNo > UBound(tmSpecCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX
            edcSpecDropDown.SetFocus
        Case PRICEINDEX
            pbcPrice.SetFocus
        Case INVTIMEINDEX
            pbcInvTime.SetFocus
        Case PROGRAMMATICINDEX
            pbcProgrammatic.SetFocus
        Case SALESBROCHUREINDEX
            edcSpecDropDown.SetFocus
        Case ALTERNAMEINDEX
            pbcAlter.SetFocus
        Case ALTERHIDDENINDEX
            pbcAlter.SetFocus
        Case DEMOINDEX
            edcSpecDropDown.SetFocus
        Case MKTINDEX
            edcSpecDropDown.SetFocus
        Case DOLLARINDEX
            edcSpecDropDown.SetFocus
        Case ACT1CODEINDEX
            edcSpecDropDown.SetFocus
        Case ACT1SETTINGINDEX
            
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecSetShow                    *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
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
    If (ilBoxNo < imLBSpecCtrls) Or (ilBoxNo > UBound(tmSpecCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX
            edcSpecDropDown.Visible = False
            slStr = Trim$(edcSpecDropDown.Text)
            smSpecSave(1) = slStr
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case PRICEINDEX
            pbcPrice.Visible = False
            Select Case imSpecSave(1)
                Case 0
                    slStr = "Rate"
                Case 1
                    slStr = "Audience"
                Case 2
                    slStr = "Percent"
                Case 3
                    slStr = "Spot Count"
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case INVTIMEINDEX
            pbcInvTime.Visible = False
            Select Case imSpecSave(2)
                Case 0
                    slStr = "Real"
                Case 1
                    slStr = "Virtual"   '"Generate"
                Case 2
                    slStr = "Equal"   '"Generate"
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case PROGRAMMATICINDEX 'Avail
            pbcProgrammatic.Visible = False
            Select Case imSpecSave(5)
                Case 0
                    slStr = "Yes"
                Case 1
                    slStr = "No"
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case SALESBROCHUREINDEX
            lbcPDFName.Visible = False
            edcSpecDropDown.Visible = False
            cmcSpecDropDown.Visible = False
            slStr = edcSpecDropDown.Text
            smSpecSave(5) = slStr
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case ALTERNAMEINDEX 'Avail
            pbcAlter.Visible = False
            Select Case imSpecSave(4)
                Case 0
                    slStr = "Yes"
                Case 1
                    slStr = "No"
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case ALTERHIDDENINDEX 'Avail
            pbcAlter.Visible = False
            Select Case imSpecSave(3)
                Case -1
                    slStr = "Yes"
                Case 0
                    slStr = "Yes"
                Case 1
                    slStr = "No"
                Case 2
                    slStr = "Cmmt/Audio only"
                Case 3
                    slStr = "Rate only"
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case DEMOINDEX
            lbcDemo.Visible = False
            edcSpecDropDown.Visible = False
            cmcSpecDropDown.Visible = False
            If lbcDemo.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcDemo.List(lbcDemo.ListIndex)
            End If
            If smSpecSave(2) <> slStr Then
                pbcPkg.Cls
                smSpecSave(2) = slStr
                gSetShow pbcSpec, slStr, tmSpecCtrls(DEMOINDEX)
                mGetPkgAud -1
                mGetTotals
                pbcPkg_Paint
            End If
        Case MKTINDEX 'Market Name
            lbcVehGp3.Visible = False
            edcSpecDropDown.Visible = False
            cmcSpecDropDown.Visible = False
            slStr = edcSpecDropDown.Text
            smSpecSave(4) = slStr
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case DOLLARINDEX
            edcSpecDropDown.Visible = False
            slStr = Trim$(edcSpecDropDown.Text)
            If smSpecSave(3) <> slStr Then
                smSpecSave(3) = slStr
                gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
                pbcPkg.Cls
                mAdjPrices
                mGetPkgAud -1
                mGetTotals
                pbcPkg_Paint
            End If
            
        'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
        Case ACT1CODEINDEX
            edcSpecDropDown.Visible = False
            slStr = edcSpecDropDown.Text
            gSetShow pbcPkg, slStr, tmSpecCtrls(ACT1CODEINDEX)
            smSpecSave(6) = slStr
            
        Case ACT1SETTINGINDEX
            plcACT1Settings.Visible = False
            edcSpecDropDown.Visible = False
            edcSpecDropDown.Text = ""
            If edcACT1SettingT.Text = "Yes" Then edcSpecDropDown.Text = edcSpecDropDown.Text & "T"
            If edcACT1SettingS.Text = "Yes" Then edcSpecDropDown.Text = edcSpecDropDown.Text & "S"
            If edcACT1SettingC.Text = "Yes" Then edcSpecDropDown.Text = edcSpecDropDown.Text & "C"
            If edcACT1SettingF.Text = "Yes" Then edcSpecDropDown.Text = edcSpecDropDown.Text & "F"
            slStr = edcSpecDropDown.Text
            gSetShow pbcPkg, slStr, tmSpecCtrls(ACT1SETTINGINDEX)
            smSpecSave(7) = slStr

    End Select
    mSpecSetChg imSpecBoxNo, False
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
    Dim ilRet As Integer
    Screen.MousePointer = vbDefault
    If imAddPkg Then
'        sgMVefStamp = "~'"
'        ilRet = gObtainVef()
'        sgVpfStamp = "~"    'Force read
'        ilRet = gVpfRead()
    End If

    igManUnload = YES
    Unload StdPkg
    igManUnload = NO
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSaveFields                 *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields(ilMsg As Integer, ilSetBox As Integer) As Integer
'
'   iRet = mTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilLoop As Integer
    Dim slTotalPct As String
    If smSpecSave(1) = "" Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Name must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imSpecBoxNo = NAMEINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    If imSpecSave(1) < 0 Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Price type must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imSpecBoxNo = PRICEINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    If imSpecSave(2) < 0 Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Package Invoice Time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imSpecBoxNo = INVTIMEINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    If imSpecSave(4) < 0 Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Alter Name must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imSpecBoxNo = ALTERNAMEINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    If imSpecSave(3) < 0 Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Alter Hidden must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imSpecBoxNo = ALTERHIDDENINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If
    If smSpecSave(4) = "" Then
        If tgSpf.sMktBase = "Y" Then
            If ilMsg = SHOWMSG Then
                ilRes = MsgBox("Market must be specified", vbOKOnly + vbExclamation, "Incomplete")
            End If
            If ilSetBox Then
                imSpecBoxNo = MKTINDEX
            End If
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If imSpecSave(1) = 2 Then
        slTotalPct = "0.0"
        For ilLoop = LBONE To UBound(smPkgSave, 2) - 1 Step 1
            If Trim$(smPkgSave(2, ilLoop)) <> "" Then
                slTotalPct = gAddStr(Trim$(smPkgSave(2, ilLoop)), slTotalPct)
            End If
        Next ilLoop
        If gCompNumberStr(slTotalPct, "100.00") <> 0 Then
            If ilMsg = SHOWMSG Then
                ilRes = MsgBox("Percent Not Equal to 100", vbOKOnly + vbExclamation, "Incomplete")
            End If
            If ilSetBox Then
                imSpecBoxNo = PRICEINDEX
            End If
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    mTestSaveFields = YES
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp3Branch                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Vehicle group and process      *
'*                      communication back from        *
'*                      invoice sorting                *
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
Private Function mVehGp3Branch() As Integer
'
'   ilRet = mVehGp3Branch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    Dim ilEnable As Integer
    ilRet = gOptionalLookAhead(edcSpecDropDown, lbcVehGp3, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcSpecDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mVehGp3Branch = False
        Exit Function
    End If
    ilEnable = cbcSelect.Enabled
    cbcSelect.Enabled = False
    sgMnfCallType = "H"
    igMNmCallSource = CALLSOURCEVEHICLE
    If edcSpecDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    If igTestSystem Then
        slStr = "RateCard^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\3"
    Else
        slStr = "RateCard^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\3"
    End If

    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    cbcSelect.Enabled = ilEnable

    imDoubleClickName = False
    mVehGp3Branch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
        lbcVehGp3.Clear
        smVehGp3CodeTag = ""
        mVehGp3Pop
        If imTerminate Then
            mVehGp3Branch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 2, lbcVehGp3
        sgMNmName = ""
        If gLastFound(lbcVehGp3) > 0 Then
            imVehGp3ChgMode = True
            lbcVehGp3.ListIndex = gLastFound(lbcVehGp3)
            edcSpecDropDown.Text = lbcVehGp3.List(lbcVehGp3.ListIndex)
            imVehGp3ChgMode = False
            mVehGp3Branch = False
        Else
            imVehGp3ChgMode = True
            lbcVehGp3.ListIndex = 1
            edcSpecDropDown.Text = lbcVehGp3.List(1)
            imVehGp3ChgMode = False
            edcSpecDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mSpecEnableBox imSpecBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mSpecEnableBox imSpecBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp3Pop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Vehicle Group Code    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mVehGp3Pop()
'
'   mVehGpPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcVehGp3.ListIndex
    If ilIndex > 1 Then
        slName = lbcVehGp3.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gIMoveListBox(Vehicle, lbcVehGp, tmVehGpCode(), smVehGpCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gPopMnfPlusFieldsBox(StdPkg, lbcVehGp3, tmVehGp3Code(), smVehGp3CodeTag, "H3")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehGp3PopErr
        gCPErrorMsg ilRet, "mVehGp3Pop (gPopMnfPlusFieldsBox)", StdPkg
        On Error GoTo 0
        lbcVehGp3.AddItem "[None]", 0  'Force as first item on list
        lbcVehGp3.AddItem "[New]", 0  'Force as first item on list
        imVehGp3ChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcVehGp3
            If gLastFound(lbcVehGp3) >= 2 Then
                lbcVehGp3.ListIndex = gLastFound(lbcVehGp3)
            Else
                lbcVehGp3.ListIndex = -1
            End If
        Else
            lbcVehGp3.ListIndex = ilIndex
        End If
        imVehGp3ChgMode = False
    End If
    Exit Sub
mVehGp3PopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Private Sub pbcAlter_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcAlter_KeyPress(KeyAscii As Integer)
    Dim ilIndex As Integer
    If imSpecBoxNo = ALTERNAMEINDEX Then
        ilIndex = 4
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            imSpecSave(ilIndex) = 0
            pbcAlter_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            imSpecSave(ilIndex) = 1
            pbcAlter_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSpecSave(ilIndex) = 0 Then
                imSpecSave(ilIndex) = 1
                pbcAlter_Paint
            ElseIf imSpecSave(ilIndex) = 1 Then
                imSpecSave(ilIndex) = 0
                pbcAlter_Paint
            End If
        End If
    Else
        ilIndex = 3
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            imSpecSave(ilIndex) = 0
            pbcAlter_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            imSpecSave(ilIndex) = 1
            pbcAlter_Paint
        ElseIf KeyAscii = Asc("C") Or (KeyAscii = Asc("c")) Then
            imSpecSave(ilIndex) = 2
            pbcAlter_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSpecSave(ilIndex) = 0 Then
                imSpecSave(ilIndex) = 1
                pbcAlter_Paint
            ElseIf imSpecSave(ilIndex) = 1 Then
                imSpecSave(ilIndex) = 2
                pbcAlter_Paint
            ElseIf imSpecSave(ilIndex) = 2 Then
                imSpecSave(ilIndex) = 0
                pbcAlter_Paint
            End If
        End If
    End If
    mSpecSetChg imSpecBoxNo, True
End Sub

Private Sub pbcAlter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilIndex As Integer
    If imSpecBoxNo = ALTERNAMEINDEX Then
        ilIndex = 4
        If imSpecSave(ilIndex) = 0 Then
            imSpecSave(ilIndex) = 1
        Else
            imSpecSave(ilIndex) = 0
        End If
    Else
        ilIndex = 3
        If imSpecSave(ilIndex) = 0 Then
            imSpecSave(ilIndex) = 1
        ElseIf imSpecSave(ilIndex) = 1 Then
            imSpecSave(ilIndex) = 2
        ElseIf imSpecSave(ilIndex) = 2 Then
            imSpecSave(ilIndex) = 3
        Else
            imSpecSave(ilIndex) = 0
        End If
    End If
    pbcAlter_Paint
    mSpecSetChg imSpecBoxNo, True
End Sub

Private Sub pbcAlter_Paint()
    Dim ilIndex As Integer
    If imSpecBoxNo = ALTERNAMEINDEX Then
        ilIndex = 4
    Else
        ilIndex = 3
    End If
    pbcAlter.Cls
    pbcAlter.CurrentX = fgBoxInsetX
    pbcAlter.CurrentY = 0 'fgBoxInsetY
    If imSpecSave(ilIndex) = 0 Then
        pbcAlter.Print "Yes"
    ElseIf imSpecSave(ilIndex) = 1 Then
        pbcAlter.Print "No"
    ElseIf (imSpecSave(ilIndex) = 2) Then
        pbcAlter.Print "Cmmt/Audio only"
    ElseIf (imSpecSave(ilIndex) = 3) Then
        pbcAlter.Print "Rate only"
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
'11/4/21 - JW - Fix Per Dan found some gremlins
'    mSpecSetShow imSpecBoxNo
'    imSpecBoxNo = -1
'    mPkgSetShow imPkgBoxNo
'    imPkgBoxNo = -1
'    imPkgRowNo = -1
    pbcArrow.Visible = False
    If plcACT1Settings.Visible Then pbcSpecTab.SetFocus
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
    End If
End Sub

Private Sub pbcInvTime_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcInvTime_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("R")) Or (KeyAscii = Asc("r")) And (tgSpf.sCPkAired = "Y") Then
        imSpecSave(2) = 0
        pbcInvTime_Paint
    ElseIf KeyAscii = Asc("V") Or (KeyAscii = Asc("v")) And (tgSpf.sCPkOrdered = "Y") Then
        imSpecSave(2) = 1
        pbcInvTime_Paint
    ElseIf KeyAscii = Asc("E") Or (KeyAscii = Asc("e")) And (tgSpf.sCPkEqual = "Y") Then
        imSpecSave(2) = 2
        pbcInvTime_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imSpecSave(2) = 0 Then
            If tgSpf.sCPkOrdered = "Y" Then
                imSpecSave(2) = 1
            ElseIf tgSpf.sCPkEqual = "Y" Then
                imSpecSave(2) = 2
            End If
            pbcInvTime_Paint
        ElseIf imSpecSave(2) = 1 Then
            If tgSpf.sCPkEqual = "Y" Then
                imSpecSave(2) = 2
            ElseIf tgSpf.sCPkAired = "Y" Then
                imSpecSave(2) = 1
            End If
            pbcInvTime_Paint
        ElseIf imSpecSave(2) = 2 Then
            If tgSpf.sCPkAired = "Y" Then
                imSpecSave(2) = 0
            ElseIf tgSpf.sCPkOrdered = "Y" Then
                imSpecSave(2) = 1
            End If
            pbcInvTime_Paint
        End If
    End If
    mSpecSetChg imSpecBoxNo, True
End Sub

Private Sub pbcInvTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imSpecSave(2) = 0 Then
        If tgSpf.sCPkOrdered = "Y" Then
            imSpecSave(2) = 1
        ElseIf tgSpf.sCPkEqual = "Y" Then
            imSpecSave(2) = 2
        End If
    ElseIf imSpecSave(2) = 1 Then
        If tgSpf.sCPkEqual = "Y" Then
            imSpecSave(2) = 2
        ElseIf tgSpf.sCPkAired = "Y" Then
            imSpecSave(2) = 0
        End If
    Else
        If tgSpf.sCPkAired = "Y" Then
            imSpecSave(2) = 0
        ElseIf tgSpf.sCPkOrdered = "Y" Then
            imSpecSave(2) = 1
        ElseIf tgSpf.sCPkEqual = "Y" Then
            imSpecSave(2) = 2
        End If
    End If
    pbcInvTime_Paint
    mSpecSetChg imSpecBoxNo, True
End Sub

Private Sub pbcInvTime_Paint()
    pbcInvTime.Cls
    pbcInvTime.CurrentX = fgBoxInsetX
    pbcInvTime.CurrentY = 0 'fgBoxInsetY
    If imSpecSave(2) = 0 Then
        pbcInvTime.Print "Real"
    ElseIf imSpecSave(2) = 1 Then
        pbcInvTime.Print "Virtual"  '"Generate"
    ElseIf imSpecSave(2) = 2 Then
        pbcInvTime.Print "Equal"
    Else
        pbcInvTime.Print ""
    End If
End Sub

'11/4/21 - JW - Fix Per Dan found some gremlins
Private Sub pbcPkg_GotFocus()
    If plcACT1Settings.Visible Then pbcSpecTab.SetFocus
End Sub

Private Sub pbcPkg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim ilCompRow As Integer
    ilCompRow = vbcPkg.LargeChange + 1
    If UBound(smPkgSave, 2) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smPkgSave, 2) + 1  'UBound(tgBvfRec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = PKGSPOTINDEX To PKGPERCENTINDEX Step 1
            If (X >= tmPkgCtrls(ilBox).fBoxX) And (X <= (tmPkgCtrls(ilBox).fBoxX + tmPkgCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmPkgCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmPkgCtrls(ilBox).fBoxY + tmPkgCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcPkg.Value - 1
                    If ilRowNo >= UBound(smPkgSave, 2) Then
                        Beep
                        mPkgSetFocus imPkgBoxNo
                        Exit Sub
                    End If
                    If (ilBox = PKGPERCENTINDEX) And (imSpecSave(1) <> 2) Then
                        Beep
                        mPkgSetFocus imPkgBoxNo
                        Exit Sub
                    End If
                    mPkgSetShow imPkgBoxNo
                    imPkgRowNo = ilRow + vbcPkg.Value - 1
                    imPkgBoxNo = ilBox
                    mPkgEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mPkgSetFocus imPkgBoxNo
End Sub

Private Sub pbcPkg_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim llColor As Long
    Dim slStr As String

    mPaintPkgTitle
    ilStartRow = vbcPkg.Value '+ 1  'Top location
    ilEndRow = vbcPkg.Value + vbcPkg.LargeChange ' + 1
    If ilEndRow > UBound(smPkgSave, 2) Then
        ilEndRow = UBound(smPkgSave, 2) 'include blank row as it might have data
    End If
    llColor = pbcPkg.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBPkgCtrls To UBound(tmPkgCtrls) Step 1
            If (ilBox = PKGVEHINDEX) And (tmPBDP(ilRow).iPkgVeh > 0) Then
                pbcPkg.ForeColor = BLUE
            End If
            If (ilBox = PKGVEHINDEX) And (tmPBDP(ilRow).iVehDormant > 0) Then
                pbcPkg.ForeColor = Red
            End If
            If (ilBox = PKGDPINDEX) And (tmPBDP(ilRow).iDPDormant > 0) Then
                pbcPkg.ForeColor = Red
            End If
            If (ilBox = PKGVEHINDEX) And (tmPBDP(ilRow).sMedium = "P") Then
                'PODCAST SPOT, Show Vehicle Name in Italic font
                pbcPkg.FontItalic = True
            End If
            pbcPkg.CurrentX = tmPkgCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcPkg.CurrentY = tmPkgCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            slStr = Trim$(smPkgShow(ilBox, ilRow))
            pbcPkg.Print slStr
            pbcPkg.ForeColor = llColor
            pbcPkg.FontItalic = False
        Next ilBox
    Next ilRow
    For ilBox = LBONE To UBound(smTShow) Step 1
        If ilBox = LBONE Then
            pbcPkg.CurrentX = tmPkgCtrls(ilBox + PKGPRICEINDEX - LBONE).fBoxX + fgBoxInsetX
        Else
            pbcPkg.CurrentX = tmPkgCtrls(ilBox + PKGPRICEINDEX - LBONE + 1).fBoxX + fgBoxInsetX
        End If
        pbcPkg.CurrentY = tmPkgCtrls(1).fBoxY + (vbcPkg.LargeChange + 1) * (fgBoxGridH + 15) + 15
        pbcPkg.Print smTShow(ilBox)
    Next ilBox
End Sub

Private Sub pbcPkgSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcPkgSTab.HWnd Then
        Exit Sub
    End If
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    imTabDirection = -1 'Set- Right to left
    Select Case imPkgBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            imSettingValue = True
            vbcPkg.Value = 1
            imSettingValue = False
            imPkgRowNo = 1
            ilBox = PKGSPOTINDEX
            imPkgBoxNo = ilBox
            mPkgEnableBox ilBox
            Exit Sub
        Case PKGSPOTINDEX 'Name (first control within header)
            mPkgSetShow imPkgBoxNo
            If imPkgRowNo <= 1 Then
                If cbcSelect.Enabled Then
                    imPkgBoxNo = -1
                    cbcSelect.SetFocus
                End If
                'TTP 10325 - JW 11/3/21 - Wrap from PkgGrid to Spec Grid - Skip over ACT1 if Disabled
                If bmAct1Allowed = True Then
                    imSpecBoxNo = ACT1SETTINGINDEX + 1
                Else
                    imSpecBoxNo = DOLLARINDEX + 1
                End If
                pbcSpecSTab.SetFocus
                mSpecSetShow imSpecBoxNo
                Exit Sub
                ilBox = 1
            Else
                If imSpecSave(1) <> 2 Then
                    ilBox = PKGSPOTINDEX
                Else
                    ilBox = PKGPERCENTINDEX
                End If
                imPkgRowNo = imPkgRowNo - 1
                If imPkgRowNo < vbcPkg.Value Then
                    imSettingValue = True
                    vbcPkg.Value = vbcPkg.Value - 1
                    imSettingValue = False
                End If
                imPkgBoxNo = ilBox
                mPkgEnableBox ilBox
                Exit Sub
            End If
        Case Else
            ilBox = imPkgBoxNo - 1
    End Select
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = ilBox
    mPkgEnableBox ilBox
End Sub

Private Sub pbcPkgTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcPkgTab.HWnd Then
        Exit Sub
    End If
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    imTabDirection = 0 'Set- Left to right
    Select Case imPkgBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            imPkgRowNo = UBound(tmPBDP) - 1
            imSettingValue = True
            If imPkgRowNo <= vbcPkg.LargeChange + 1 Then
                vbcPkg.Value = 1
            Else
                vbcPkg.Value = imPkgRowNo - vbcPkg.LargeChange - 1
            End If
            imSettingValue = False
            If imSpecSave(1) <> 2 Then
                ilBox = PKGSPOTINDEX
            Else
                ilBox = PKGPERCENTINDEX
            End If
        Case 0
            ilBox = PKGSPOTINDEX
        Case PKGSPOTINDEX
            If imSpecSave(1) <> 2 Then
                mPkgSetShow imPkgBoxNo
                If imPkgRowNo + 1 >= UBound(tmPBDP) Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imPkgRowNo = imPkgRowNo + 1
                If imPkgRowNo > vbcPkg.Value + vbcPkg.LargeChange Then
                    imSettingValue = True
                    vbcPkg.Value = vbcPkg.Value + 1
                    imSettingValue = False
                End If
                ilBox = PKGSPOTINDEX
                imPkgBoxNo = ilBox
                mPkgEnableBox ilBox
                Exit Sub
            Else
                ilBox = PKGPERCENTINDEX
            End If

        Case PKGPERCENTINDEX
            mPkgSetShow imPkgBoxNo
            If imPkgRowNo + 1 >= UBound(tmPBDP) Then
                cmcDone.SetFocus
                Exit Sub
            End If
            imPkgRowNo = imPkgRowNo + 1
            If imPkgRowNo > vbcPkg.Value + vbcPkg.LargeChange Then
                imSettingValue = True
                vbcPkg.Value = vbcPkg.Value + 1
                imSettingValue = False
            End If
            ilBox = PKGSPOTINDEX
            imPkgBoxNo = ilBox
            mPkgEnableBox ilBox
            Exit Sub
        Case Else
            ilBox = imPkgBoxNo + 1
    End Select
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = ilBox
    mPkgEnableBox ilBox
End Sub

Private Sub pbcPrice_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcPrice_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("R")) Or (KeyAscii = Asc("r")) Then
        imSpecSave(1) = 0
        tmSpecCtrls(imSpecBoxNo).iChg = True
        pbcPrice_Paint
    ElseIf (KeyAscii = Asc("A") Or (KeyAscii = Asc("a"))) And (tgSpf.sCAudPkg = "Y") Then
        imSpecSave(1) = 1
        tmSpecCtrls(imSpecBoxNo).iChg = True
        pbcPrice_Paint
    ElseIf KeyAscii = Asc("P") Or (KeyAscii = Asc("p")) Then
        imSpecSave(1) = 2
        tmSpecCtrls(imSpecBoxNo).iChg = True
        pbcPrice_Paint
    ElseIf KeyAscii = Asc("S") Or (KeyAscii = Asc("s")) Then
        imSpecSave(1) = 3
        tmSpecCtrls(imSpecBoxNo).iChg = True
        pbcPrice_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imSpecSave(1) = 0 Then
            tmSpecCtrls(imSpecBoxNo).iChg = True
            If tgSpf.sCAudPkg = "Y" Then
                imSpecSave(1) = 1
            Else
                imSpecSave(1) = 2
            End If
            pbcPrice_Paint
        ElseIf imSpecSave(1) = 1 Then
            tmSpecCtrls(imSpecBoxNo).iChg = True
            imSpecSave(1) = 2
            pbcPrice_Paint
        ElseIf imSpecSave(1) = 2 Then
            tmSpecCtrls(imSpecBoxNo).iChg = True
            imSpecSave(1) = 3
            pbcPrice_Paint
        ElseIf imSpecSave(1) = 3 Then
            tmSpecCtrls(imSpecBoxNo).iChg = True
            imSpecSave(1) = 0
            pbcPrice_Paint
        End If
    End If
    mSpecSetChg imSpecBoxNo, True
End Sub

Private Sub pbcPrice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imSpecSave(1) = 0 Then
        If tgSpf.sCAudPkg = "Y" Then
            imSpecSave(1) = 1
        Else
            imSpecSave(1) = 2
        End If
    ElseIf imSpecSave(1) = 1 Then
        imSpecSave(1) = 2
    ElseIf imSpecSave(1) = 2 Then
        imSpecSave(1) = 3
    Else
        imSpecSave(1) = 0
    End If
    tmSpecCtrls(imSpecBoxNo).iChg = True
    pbcPrice_Paint
    mSpecSetChg imSpecBoxNo, True
End Sub

Private Sub pbcPrice_Paint()
    pbcPrice.Cls
    pbcPrice.CurrentX = fgBoxInsetX
    pbcPrice.CurrentY = 0 'fgBoxInsetY
    If imSpecSave(1) = 0 Then
        pbcPrice.Print "Rate"
    ElseIf imSpecSave(1) = 1 Then
        pbcPrice.Print "Audience"
    ElseIf imSpecSave(1) = 2 Then
        pbcPrice.Print "Percent"
    ElseIf imSpecSave(1) = 3 Then
        pbcPrice.Print "Spot Count"
    End If
End Sub

Private Sub pbcProgrammatic_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcProgrammatic_KeyPress(KeyAscii As Integer)
    Dim ilIndex As Integer
    ilIndex = 5
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        imSpecSave(ilIndex) = 0
        pbcProgrammatic_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        imSpecSave(ilIndex) = 1
        pbcProgrammatic_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imSpecSave(ilIndex) = 0 Then
            imSpecSave(ilIndex) = 1
            pbcProgrammatic_Paint
        ElseIf imSpecSave(ilIndex) = 1 Then
            imSpecSave(ilIndex) = 0
            pbcProgrammatic_Paint
        End If
    End If

    mSpecSetChg imSpecBoxNo, True
End Sub

Private Sub pbcProgrammatic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilIndex As Integer
    ilIndex = 5
    If imSpecSave(ilIndex) = 0 Then
        imSpecSave(ilIndex) = 1
    Else
        imSpecSave(ilIndex) = 0
    End If
    pbcProgrammatic_Paint
    mSpecSetChg imSpecBoxNo, True

End Sub

Private Sub pbcProgrammatic_Paint()
    Dim ilIndex As Integer
    ilIndex = 5
    pbcProgrammatic.Cls
    pbcProgrammatic.CurrentX = fgBoxInsetX
    pbcProgrammatic.CurrentY = 0 'fgBoxInsetY
    If imSpecSave(ilIndex) = 0 Then
        pbcProgrammatic.Print "Yes"
    ElseIf imSpecSave(ilIndex) = 1 Then
        pbcProgrammatic.Print "No"
    End If
End Sub

Private Sub pbcSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = NAMEINDEX To ACT1SETTINGINDEX Step 1
        If (X >= tmSpecCtrls(ilBox).fBoxX) And (X <= (tmSpecCtrls(ilBox).fBoxX + tmSpecCtrls(ilBox).fBoxW)) Then
            If (Y >= (tmSpecCtrls(ilBox).fBoxY)) And (Y <= (tmSpecCtrls(ilBox).fBoxY + tmSpecCtrls(ilBox).fBoxH)) Then
                If ((ilBox = PROGRAMMATICINDEX) Or (ilBox = SALESBROCHUREINDEX)) And (Not bmProgrammaticAllowed) Then
                    Beep
                    Exit Sub
                End If
                'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
                If ((ilBox = ACT1CODEINDEX) Or (ilBox = ACT1SETTINGINDEX)) And (Not bmAct1Allowed) Then
                    Beep
                    Exit Sub
                End If
                
                mSpecSetShow imSpecBoxNo
                imSpecBoxNo = ilBox
                mSpecEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSpecSetFocus imSpecBoxNo
End Sub

Private Sub pbcSpec_Paint()
    Dim ilBox As Integer
    Dim slStr As String
    Dim blReadOnly As Boolean
    pbcSpec.Cls
    
    'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
    bmAct1Allowed = False
    If (tmSpecCtrls(ALTERHIDDENINDEX).sShow = "No" Or tmSpecCtrls(ALTERHIDDENINDEX).sShow = "Cmmt/Au" Or tmSpecCtrls(ALTERHIDDENINDEX).sShow = "Rate Only") And tmSpecCtrls(ALTERNAMEINDEX).sShow = "No" And bmAct1Enabled = True Then bmAct1Allowed = True
    
    For ilBox = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        'gPaintArea pbcSpec, tmSpecCtrls(ilBox).fBoxX, tmSpecCtrls(ilBox).fBoxY + (ilRow - 1) * (fgBoxGridH + 15), tmSpecCtrls(ilBox).fBoxW - 15, tmSpecCtrls(ilBox).fBoxH - 15, WHITE
        slStr = tmSpecCtrls(ilBox).sShow
        
        blReadOnly = False
        If Not bmProgrammaticAllowed And (ilBox = SALESBROCHUREINDEX Or ilBox = PROGRAMMATICINDEX) Then blReadOnly = True
        If Not bmAct1Allowed And (ilBox = ACT1CODEINDEX Or ilBox = ACT1SETTINGINDEX) Then blReadOnly = True
        If blReadOnly Then
            pbcSpec.FillStyle = vbSolid
            pbcSpec.FillColor = LIGHTYELLOW
            pbcSpec.ForeColor = LIGHTYELLOW
            pbcSpec.Line (tmSpecCtrls(ilBox).fBoxX, tmSpecCtrls(ilBox).fBoxY + fgBoxInsetY + 30)-Step(tmSpecCtrls(ilBox).fBoxW - 15, tmSpecCtrls(ilBox).fBoxH - fgBoxInsetY - 40), LIGHTYELLOW, B
        End If
        '11/5/21 - JW - Wipe out Values for ACT1 code and Setting when it becomes ReadOnly per Jason Teams
        If (ilBox = ACT1CODEINDEX Or ilBox = ACT1SETTINGINDEX) And bmAct1Allowed = False Then
            tmSpecCtrls(ilBox).sShow = ""
            smSpecSave(6) = ""
            smSpecSave(7) = ""
        Else
            pbcSpec.CurrentX = tmSpecCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcSpec.CurrentY = tmSpecCtrls(ilBox).fBoxY + fgBoxInsetY
            pbcSpec.ForeColor = vbBlack
            pbcSpec.Print tmSpecCtrls(ilBox).sShow
        End If
    Next ilBox
End Sub

Private Sub pbcSpecSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSpecSTab.HWnd Then
        Exit Sub
    End If
    If (imSpecBoxNo = MKTINDEX) Then
        If mVehGp3Branch() Then
            Exit Sub
        End If
    End If
    imTabDirection = -1  'Set-Right to left
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = -1
    imPkgRowNo = -1
    Select Case imSpecBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                ilBox = 1
                mSetCommands
            Else
                ilBox = 2
            End If
        Case NAMEINDEX
            mSpecSetShow imSpecBoxNo
            imSpecBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            If (cmcUpdate.Enabled) And (igDPNameCallSource = CALLNONE) Then
                cmcUpdate.SetFocus
            Else
                cmcDone.SetFocus
            End If
            Exit Sub
        'TTP 10325 - JW 11/3/21 - ACT1 Default settings for Pkg Vehicle
        Case ACT1SETTINGINDEX + 1
            If bmAct1Allowed Then
                ilBox = ACT1SETTINGINDEX
            Else
                ilBox = INVTIMEINDEX
            End If
        Case ALTERNAMEINDEX
            If bmProgrammaticAllowed Then
                ilBox = SALESBROCHUREINDEX
            Else
                ilBox = INVTIMEINDEX
            End If
        Case Else
            ilBox = imSpecBoxNo - 1
    End Select
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = ilBox
    mSpecEnableBox ilBox
End Sub

Private Sub pbcSpecTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSpecTab.HWnd Then
        Exit Sub
    End If
    If (imSpecBoxNo = MKTINDEX) Then
        If mVehGp3Branch() Then
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = -1
    imPkgRowNo = -1
    Select Case imSpecBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            ilBox = ALTERHIDDENINDEX
        Case DOLLARINDEX 'Last control within header when ACT1 is disabled
            If bmAct1Allowed = False Then
                mSpecSetShow imSpecBoxNo
                pbcPkgSTab.SetFocus
                Exit Sub
            Else
                ilBox = imSpecBoxNo + 1
            End If
        Case ACT1SETTINGINDEX 'Last control within header
            mSpecSetShow imSpecBoxNo
            pbcPkgSTab.SetFocus
            Exit Sub
        Case 0
            ilBox = NAMEINDEX
        Case INVTIMEINDEX
            If bmProgrammaticAllowed Then
                ilBox = PROGRAMMATICINDEX
            Else
                ilBox = ALTERNAMEINDEX
            End If
        Case Else
            ilBox = imSpecBoxNo + 1
    End Select
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = ilBox
    mSpecEnableBox ilBox
End Sub

Private Sub pbcStartNew_GotFocus()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    If (Not imFirstTimeSelect) Then
        Exit Sub
    End If
    imFirstTimeSelect = False
    If (imSelectedIndex = 0) And (cbcSelect.ListCount > 1) Then
        igStdPkgReturn = 0
        igStdPkgModel = 0
        sgTmpSortTag = "S" 'Show Standard Package Vehicles
        SPModel.Show vbModal
        If (igStdPkgReturn = 1) And (igStdPkgModel > 0) Then    'Done
            For ilLoop = LBound(tmPkgVehicle) To UBound(tmPkgVehicle) - 1 Step 1
                slNameCode = tmPkgVehicle(ilLoop).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If CInt(slCode) = igStdPkgModel Then
                    pbcPkg.Cls
                    ilRet = mReadRec(ilLoop + 1, SETFORREADONLY)
                    tmVef.sName = ""
                    tmVef.iCode = 0
                    mMoveRecToCtrl
                    mInitShow
                    If lbcDemo.ListIndex > 0 Then
                        mGetPkgAud -1
                    End If
                    mGetTotals
                    ReDim tmPvf(0 To 0) As PVF
                    pbcSpec_Paint
                    pbcPkg_Paint
                    Exit For
                End If
            Next ilLoop
        End If
    End If
    '2/7/09: Added to handle case where focus can't be set
    On Error Resume Next
    imSpecBoxNo = 0
    mSpecSetShow 0
    pbcSpecTab.SetFocus
    On Error GoTo 0
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imSpecBoxNo
        Case MKTINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcVehGp3, edcSpecDropDown, imVehGp3ChgMode, imLbcArrowSetting
    End Select
End Sub

Private Sub tmcInit_Timer()
    tmcInit.Enabled = False
    vbcPkg.Visible = False
    DoEvents
    vbcPkg.Visible = True
End Sub

Private Sub vbcPkg_Change()
    If imSettingValue Then
        pbcPkg.Cls
        pbcPkg_Paint
        imSettingValue = False
    Else
        mPkgSetShow imPkgBoxNo
        pbcPkg.Cls
        pbcPkg_Paint
        If (igWinStatus(RATECARDSJOB) <> 1) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            mPkgEnableBox imPkgBoxNo
        End If
    End If
End Sub

Private Sub vbcPkg_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Standard Packages"
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
Private Sub mPaintPkgTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer

    llColor = pbcPkg.ForeColor
    slFontName = pbcPkg.FontName
    flFontSize = pbcPkg.FontSize
    ilFillStyle = pbcPkg.FillStyle
    llFillColor = pbcPkg.FillColor
    pbcPkg.ForeColor = BLUE
    pbcPkg.FontBold = False
    pbcPkg.FontSize = 7
    pbcPkg.FontName = "Arial"
    pbcPkg.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    pbcPkg.Line (tmPkgCtrls(PKGVEHINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGVEHINDEX).fBoxW + 15, tmPkgCtrls(PKGVEHINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGVEHINDEX).fBoxX, 30)-Step(tmPkgCtrls(PKGVEHINDEX).fBoxW - 15, tmPkgCtrls(PKGVEHINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.CurrentX = tmPkgCtrls(PKGVEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPkg.Print "Vehicle"
    pbcPkg.Line (tmPkgCtrls(PKGDPINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGDPINDEX).fBoxW + 15, tmPkgCtrls(PKGDPINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGDPINDEX).fBoxX, 30)-Step(tmPkgCtrls(PKGDPINDEX).fBoxW - 15, tmPkgCtrls(PKGDPINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.CurrentX = tmPkgCtrls(PKGDPINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPkg.Print "Daypart"
    pbcPkg.Line (tmPkgCtrls(PKGPRICEINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGPRICEINDEX).fBoxW + 15, tmPkgCtrls(PKGPRICEINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGPRICEINDEX).fBoxX, 30)-Step(tmPkgCtrls(PKGPRICEINDEX).fBoxW - 15, tmPkgCtrls(PKGPRICEINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.CurrentX = tmPkgCtrls(PKGPRICEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPkg.Print "Price"
    
    pbcPkg.Line (tmPkgCtrls(PKGBOOKNAMEINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGBOOKNAMEINDEX).fBoxW + 15, tmPkgCtrls(PKGBOOKNAMEINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGBOOKNAMEINDEX).fBoxX, 30)-Step(tmPkgCtrls(PKGBOOKNAMEINDEX).fBoxW - 15, tmPkgCtrls(PKGBOOKNAMEINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.CurrentX = tmPkgCtrls(PKGBOOKNAMEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPkg.Print "Book Name"
    
    pbcPkg.Line (tmPkgCtrls(PKGRATINGINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGRATINGINDEX).fBoxW + 15, tmPkgCtrls(PKGRATINGINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGRATINGINDEX).fBoxX, 30)-Step(tmPkgCtrls(PKGRATINGINDEX).fBoxW - 15, tmPkgCtrls(PKGRATINGINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.CurrentX = tmPkgCtrls(PKGRATINGINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15
    pbcPkg.Print "Rating"
    pbcPkg.Line (tmPkgCtrls(PKGAUDINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGAUDINDEX).fBoxW + 15, tmPkgCtrls(PKGAUDINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGAUDINDEX).fBoxX, 30)-Step(tmPkgCtrls(PKGAUDINDEX).fBoxW - 15, tmPkgCtrls(PKGAUDINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.CurrentX = tmPkgCtrls(PKGAUDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15
    pbcPkg.Print "Audience"
    pbcPkg.Line (tmPkgCtrls(PKGCPPINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGCPPINDEX).fBoxW + 15, tmPkgCtrls(PKGCPPINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGCPPINDEX).fBoxX, 30)-Step(tmPkgCtrls(PKGCPPINDEX).fBoxW - 15, tmPkgCtrls(PKGCPPINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.CurrentX = tmPkgCtrls(PKGCPPINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15
    pbcPkg.Print "CPP"
    pbcPkg.Line (tmPkgCtrls(PKGCPMINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGCPMINDEX).fBoxW + 15, tmPkgCtrls(PKGCPMINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGCPMINDEX).fBoxX, 30)-Step(tmPkgCtrls(PKGCPMINDEX).fBoxW - 15, tmPkgCtrls(PKGCPMINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.CurrentX = tmPkgCtrls(PKGCPMINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPkg.Print "CPM"
    pbcPkg.Line (tmPkgCtrls(PKGSPOTINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGSPOTINDEX).fBoxW + 15, tmPkgCtrls(PKGSPOTINDEX).fBoxH + 15), BLUE, B
    pbcPkg.CurrentX = tmPkgCtrls(PKGSPOTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPkg.Print "# Spots"
    pbcPkg.Line (tmPkgCtrls(PKGPERCENTINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGPERCENTINDEX).fBoxW + 15, tmPkgCtrls(PKGPERCENTINDEX).fBoxH + 15), BLUE, B
    pbcPkg.CurrentX = tmPkgCtrls(PKGPERCENTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15
    pbcPkg.Print "%"

    ilLineCount = 0
    llTop = tmPkgCtrls(1).fBoxY
    Do
        For ilLoop = imLBPkgCtrls To UBound(tmPkgCtrls) Step 1
            pbcPkg.Line (tmPkgCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmPkgCtrls(ilLoop).fBoxW + 15, tmPkgCtrls(ilLoop).fBoxH + 15), BLUE, B
            If (ilLoop < PKGSPOTINDEX) Then
                pbcPkg.Line (tmPkgCtrls(ilLoop).fBoxX, llTop)-Step(tmPkgCtrls(ilLoop).fBoxW - 15, tmPkgCtrls(ilLoop).fBoxH - 15), LIGHTYELLOW, BF
            End If
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmPkgCtrls(1).fBoxH + 15
    Loop While llTop + tmPkgCtrls(1).fBoxH + tmPkgCtrls(1).fBoxH + 30 < pbcPkg.Height
    vbcPkg.LargeChange = ilLineCount - 1
    llTop = llTop + 30
    pbcPkg.Line (tmPkgCtrls(PKGPRICEINDEX).fBoxX - 15, llTop)-Step(tmPkgCtrls(PKGPRICEINDEX).fBoxW + 15, tmPkgCtrls(PKGPRICEINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGPRICEINDEX).fBoxX, llTop + 15)-Step(tmPkgCtrls(PKGPRICEINDEX).fBoxW - 15, tmPkgCtrls(PKGPRICEINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.Line (tmPkgCtrls(PKGRATINGINDEX).fBoxX - 15, llTop)-Step(tmPkgCtrls(PKGRATINGINDEX).fBoxW + 15, tmPkgCtrls(PKGRATINGINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGRATINGINDEX).fBoxX, llTop + 15)-Step(tmPkgCtrls(PKGRATINGINDEX).fBoxW - 15, tmPkgCtrls(PKGRATINGINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.Line (tmPkgCtrls(PKGAUDINDEX).fBoxX - 15, llTop)-Step(tmPkgCtrls(PKGAUDINDEX).fBoxW + 15, tmPkgCtrls(PKGAUDINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGAUDINDEX).fBoxX, llTop + 15)-Step(tmPkgCtrls(PKGAUDINDEX).fBoxW - 15, tmPkgCtrls(PKGAUDINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.Line (tmPkgCtrls(PKGCPPINDEX).fBoxX - 15, llTop)-Step(tmPkgCtrls(PKGCPPINDEX).fBoxW + 15, tmPkgCtrls(PKGCPPINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGCPPINDEX).fBoxX, llTop + 15)-Step(tmPkgCtrls(PKGCPPINDEX).fBoxW - 15, tmPkgCtrls(PKGCPPINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.Line (tmPkgCtrls(PKGCPMINDEX).fBoxX - 15, llTop)-Step(tmPkgCtrls(PKGCPMINDEX).fBoxW + 15, tmPkgCtrls(PKGCPMINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGCPMINDEX).fBoxX, llTop + 15)-Step(tmPkgCtrls(PKGCPMINDEX).fBoxW - 15, tmPkgCtrls(PKGCPMINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.Line (tmPkgCtrls(PKGSPOTINDEX).fBoxX - 15, llTop)-Step(tmPkgCtrls(PKGSPOTINDEX).fBoxW + 15, tmPkgCtrls(PKGSPOTINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGSPOTINDEX).fBoxX, llTop + 15)-Step(tmPkgCtrls(PKGSPOTINDEX).fBoxW - 15, tmPkgCtrls(PKGSPOTINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.Line (tmPkgCtrls(PKGPERCENTINDEX).fBoxX - 15, llTop)-Step(tmPkgCtrls(PKGPERCENTINDEX).fBoxW + 15, tmPkgCtrls(PKGPERCENTINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGPERCENTINDEX).fBoxX, llTop + 15)-Step(tmPkgCtrls(PKGPERCENTINDEX).fBoxW - 15, tmPkgCtrls(PKGPERCENTINDEX).fBoxH - 15), LIGHTYELLOW, BF

    pbcPkg.FontSize = flFontSize
    pbcPkg.FontName = slFontName
    pbcPkg.FontSize = flFontSize
    pbcPkg.ForeColor = llColor
    pbcPkg.FontBold = True

    pbcPkg.CurrentX = tmPkgCtrls(1).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = llTop '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPkg.Print "Total:"
    lacCover.Top = llTop
End Sub

Private Sub mPDFPop()
    Dim ilIndex As Integer
    
    If right$(sgSalesBrochurePath, 1) = "\" Then
        lbcPDFFile.Path = Left$(sgSalesBrochurePath, Len(sgSalesBrochurePath) - 1)
    Else
        lbcPDFFile.Path = sgSalesBrochurePath
    End If
    For ilIndex = 0 To lbcPDFFile.ListCount - 1 Step 1
        lbcPDFName.AddItem lbcPDFFile.List(ilIndex)
    Next ilIndex
    lbcPDFName.AddItem "[None]", 0
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mVffReadRec                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mVffReadRec(ilVefCode As Integer) As Integer
'
'   iRet = mVffReadRec()
'   Where:
'       ilVefCode(I) - Vehicle Code
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    If ilVefCode > 0 Then
        tmVffSrchKey1.iCode = ilVefCode
        ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, tmVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Else
        mVffReadRec = False
        Exit Function
    End If
    If ilRet <> BTRV_ERR_NONE Then
        'Add Record
        tmVff.iCode = 0
        tmVff.iVefCode = ilVefCode  'tmVef.iCode
        tmVff.sGroupName = ""
        tmVff.sWegenerExportID = ""
        tmVff.sOLAExportID = ""
        tmVff.iLiveCompliantAdj = 5
        tmVff.iUstCode = 0
        tmVff.iUrfCode = tgUrf(0).iCode
        If ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
            tmVff.sXDXMLForm = "S"
        Else
            If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) Then
                tmVff.sXDXMLForm = "P"
            Else
                tmVff.sXDXMLForm = ""
            End If
        End If
        tmVff.sXDISCIPrefix = ""
        tmVff.sXDProgCodeID = ""
        tmVff.sXDSaveCF = "Y"
        tmVff.sXDSaveHDD = "N"
        tmVff.sXDSaveNAS = "N"
        tmVff.iCwfCode = 0
        tmVff.sAirWavePrgID = ""
        tmVff.sExportAirWave = ""
        tmVff.sExportNYESPN = ""
        tmVff.sPledgeVsAir = "N"
        tmVff.sFedDelivery(0) = ""
        tmVff.sFedDelivery(1) = ""
        tmVff.sFedDelivery(2) = ""
        tmVff.sFedDelivery(3) = ""
        tmVff.sFedDelivery(4) = ""
        gPackDate "1/1/1990", tmVff.iLastAffExptDate(0), tmVff.iLastAffExptDate(1)
        tmVff.sMoveSportToNon = "N"
        tmVff.sMoveSportToSport = "N"
        tmVff.sMoveNonToSport = "N"
        tmVff.sMergeTraffic = "S"
        tmVff.sMergeAffiliate = "S"
        tmVff.sMergeWeb = "S"
        tmVff.sPledgeByEvent = "N"
        tmVff.lPledgeHdVtfCode = 0
        tmVff.lPledgeFtVtfCode = 0
        tmVff.iPledgeClearance = 0
        tmVff.sExportEncoESPN = "N"
        tmVff.sWebName = ""
        tmVff.lSeasonGhfCode = 0
        tmVff.iMcfCode = 0
        tmVff.sExportAudio = "N"
        tmVff.sExportMP2 = "N"
        tmVff.sExportCnCSpot = "N"
        tmVff.sExportEnco = "N"
        tmVff.sExportCnCNetInv = "N"
        tmVff.sIPumpEventTypeOV = ""
        tmVff.sExportIPump = "N"
        tmVff.sAddr4 = ""
        tmVff.lBBOpenCefCode = 0
        tmVff.lBBCloseCefCode = 0
        tmVff.lBBBothCefCode = 0
        tmVff.sXDSISCIPrefix = ""
        tmVff.sXDSSaveCF = "Y"
        tmVff.sXDSSaveHDD = "N"
        tmVff.sXDSSaveNAS = "N"
        tmVff.sMGsOnWeb = "N"   '"Y"
        tmVff.sReplacementOnWeb = "N"   '"Y"
        tmVff.sExportMatrix = "N"
        tmVff.sSentToXDSStatus = "N"
        tmVff.sStationComp = "N"
        tmVff.sExportSalesForce = "N"
        tmVff.sExportEfficio = "N"
        tmVff.sExportJelli = "N"
        tmVff.sOnXMLInsertion = "N"
        tmVff.sOnInsertions = "N"
        tmVff.sPostLogSource = "N"
        tmVff.sExportTableau = "N"
        tmVff.sStationPassword = ""
        tmVff.sHonorZeroUnits = "N"
        tmVff.sHideCommOnLog = "N"
        tmVff.sHideCommOnWeb = "N"
        tmVff.iConflictWinLen = 0
        tmVff.sACT1LineupCode = ""
        tmVff.sPrgmmaticAllow = "N"
        tmVff.sSalesBrochure = ""
        tmVff.sCartOnWeb = "N"
        tmVff.sDefaultAudioType = "R"
        tmVff.iLogExptArfCode = 0
        tmVff.sASICallLetters = ""
        tmVff.sASIBand = ""
        tmVff.sExportCustom = "" 'TTP 9992
        ilRet = btrInsert(hmVff, tmVff, imVffRecLen, INDEXKEY0)
        On Error GoTo mVffReadRecErr
        gBtrvErrorMsg ilRet, "mVffReadRec (btrInsert)", StdPkg
        On Error GoTo 0
    End If
    mVffReadRec = True
    Exit Function
mVffReadRecErr:
    On Error GoTo 0
    mVffReadRec = False
    Exit Function
End Function

Private Function mGetBookName(ilDnfCode As Integer) As String
    Dim ilDnf As Integer
    Dim ilRet As Integer
    mGetBookName = ""
    If ilDnfCode <= 0 Then
        Exit Function
    End If
    For ilDnf = 0 To UBound(tmSvDnf) - 1 Step 1
        If ilDnfCode = tmSvDnf(ilDnf).iCode Then
            mGetBookName = Trim$(tmSvDnf(ilDnf).sBookName)
            Exit Function
        End If
    Next ilDnf
    
    tmDnfSrchKey0.iCode = ilDnfCode
    ilRet = btrGetEqual(hmDnf, tmSvDnf(UBound(tmSvDnf)), imDnfRecLen, tmDnfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        mGetBookName = Trim$(tmSvDnf(UBound(tmSvDnf)).sBookName)
        ReDim Preserve tmSvDnf(0 To UBound(tmSvDnf) + 1) As DNF
    End If
    Exit Function
End Function
