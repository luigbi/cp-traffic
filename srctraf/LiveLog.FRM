VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form LiveLog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6210
   ClientLeft      =   465
   ClientTop       =   1950
   ClientWidth     =   9480
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
   Icon            =   "LiveLog.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbcDateTab 
      Height          =   135
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   840
   End
   Begin V81LiveLog.CSI_Calendar csiLiveLogDate 
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Text            =   "2/8/2021"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   0   'False
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
      CSI_CalDateFormat=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin VB.ListBox lbcCntr 
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
      ItemData        =   "LiveLog.frx":08CA
      Left            =   8610
      List            =   "LiveLog.frx":08D1
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5505
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox lbcNotAired 
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
      ItemData        =   "LiveLog.frx":08DE
      Left            =   3480
      List            =   "LiveLog.frx":08E0
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1965
      Visible         =   0   'False
      Width           =   3210
   End
   Begin VB.ListBox lbcCopy 
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
      ItemData        =   "LiveLog.frx":08E2
      Left            =   3195
      List            =   "LiveLog.frx":08E4
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2685
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1410
      Left            =   1275
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
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
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "LiveLog.frx":08E6
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "LiveLog.frx":15A4
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
   Begin VB.CommandButton cmcDropdown 
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
      Left            =   2190
      Picture         =   "LiveLog.frx":18AE
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1005
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1260
      MaxLength       =   10
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1005
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox pbcBanner 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   9480
      Begin VB.Label lacBanner 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Live Log Date"
         Height          =   210
         Index           =   2
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   1425
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacCopyright 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2003-2011"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   7920
         TabIndex        =   31
         Top             =   405
         Width           =   630
      End
      Begin VB.Label lacHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Help"
         Height          =   240
         Left            =   8775
         TabIndex        =   28
         Top             =   600
         Width           =   675
      End
      Begin VB.Image imcLogo 
         Height          =   615
         Left            =   5430
         Picture         =   "LiveLog.frx":19A8
         Top             =   150
         Width           =   3180
      End
      Begin VB.Label lacBanner 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   585
         Width           =   5505
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacBanner 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   300
         Width           =   5505
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox pbcLog 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   9420
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   9450
      Begin VB.PictureBox pbcFocusStop 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   9255
         ScaleHeight     =   165
         ScaleWidth      =   135
         TabIndex        =   11
         Top             =   5250
         Width           =   135
      End
      Begin VB.CommandButton cmcChgFeed 
         Appearance      =   0  'Flat
         Caption         =   "&Change Feed"
         Enabled         =   0   'False
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
         Left            =   5715
         TabIndex        =   29
         Top             =   4950
         Width           =   1560
      End
      Begin VB.CommandButton cmcUndo 
         Appearance      =   0  'Flat
         Caption         =   "&Undo"
         Enabled         =   0   'False
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
         Left            =   3705
         TabIndex        =   27
         Top             =   4950
         Width           =   1560
      End
      Begin VB.CommandButton cmcAddSpot 
         Appearance      =   0  'Flat
         Caption         =   "&Add Bonus Spot"
         Enabled         =   0   'False
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
         Left            =   7650
         TabIndex        =   24
         Top             =   4950
         Width           =   1560
      End
      Begin VB.CommandButton cmcSignOff 
         Appearance      =   0  'Flat
         Caption         =   "Sign-&Off"
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
         Left            =   1665
         TabIndex        =   23
         Top             =   4950
         Width           =   1560
      End
      Begin VB.CommandButton cmcSignOn 
         Appearance      =   0  'Flat
         Caption         =   "&Sign-On"
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
         Left            =   60
         TabIndex        =   22
         Top             =   4950
         Width           =   1560
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLog 
         Height          =   3045
         Left            =   45
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   390
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   5371
         _Version        =   393216
         Cols            =   15
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   15
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLogged 
         Height          =   1290
         Left            =   30
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3480
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   2275
         _Version        =   393216
         Rows            =   1
         Cols            =   15
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   15
      End
      Begin VB.PictureBox pbcTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   165
         ScaleHeight     =   105
         ScaleWidth      =   75
         TabIndex        =   21
         Top             =   5280
         Width           =   75
      End
      Begin VB.PictureBox pbcSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   105
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   13
         Top             =   5220
         Width           =   105
      End
      Begin VB.Label lacSpacer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   15
         TabIndex        =   25
         Top             =   3330
         Width           =   9465
      End
   End
   Begin VB.PictureBox pbcVehicleSelection 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5205
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   9420
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1245
      Width           =   9450
      Begin VB.CommandButton cmcContinue 
         Appearance      =   0  'Flat
         Caption         =   "&Continue"
         Enabled         =   0   'False
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
         Left            =   3315
         TabIndex        =   8
         Top             =   4875
         Width           =   1050
      End
      Begin VB.CommandButton cmcExit 
         Appearance      =   0  'Flat
         Caption         =   "&Exit"
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
         Left            =   5130
         TabIndex        =   9
         Top             =   4890
         Width           =   1050
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdVehicleSelection 
         Height          =   4200
         Left            =   90
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   420
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7408
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   12
      End
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   5760
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   -30
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6150
      Width           =   75
   End
End
Attribute VB_Name = "LiveLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of LiveLog.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  smSyncDate                    smSyncTime                    hmMsg                     *
'*  tmClfSrchKey                  tmFCff                        tmCff                     *
'*  tmRdfSrchKey                  tmFSFSrchKey                  tmLlfSrchKey2             *
'*  tmLdfSrchKey                  tmMnfSrchKey                                            *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mInitBox                      mParseCmmdLine                mTerminate                *
'*  mOpenMsgFile                  mReadChfClfRdfRec                                       *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: LiveLog.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export Selling to Airing Links screen code
Option Explicit
Option Compare Text
Dim imFixing As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imComboBoxIndex As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imLbcArrowSetting As Integer
Dim lmNowDate As Long
Dim smNowDate As String
Dim lmEventDate As Long
Dim smEventDate As String
Dim smLastStdMnth As String
Dim lmLastStdMnth As Long
Dim smActualStartDate As String
Dim lmTimeWidth As Long
Dim lmLengthWidth As Long
Dim lmStationCode As Long
Dim imChgMode As Integer
Dim imCheckOnExit As Integer
Dim imAllowUndo As Integer  'Undo only allowed if not complete
Dim imPrgStartUrfCode As Integer
Dim imPrgEndUrfCode As Integer

Dim lmLibStartTime(0 To 1) As Long
Dim lmLibEndTime(0 To 1) As Long

Dim tmDelOrMissedSdf() As SDF

Dim imSignOn As Integer     '0=Have not sign on; 1=Sign-on
Dim imGameStart As Integer
Dim smGameStartTime As String 'Set if Replacing another user that signed on (imSignOn = 2
Dim smGameEndTime As String 'Set if Replacing another user that signed on (imSignOn = 2
Dim imVefCode As Integer
Dim imLLVefCode() As Integer
Dim smVefType As String
Dim imVpfIndex As Integer
Dim imGameNo As Integer
Dim smCurrentFeed As String
Dim imCurrentVefCode As Integer

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim imWhichGrid As Integer '0=Log; 1=Logged

Dim tmCopyCode() As SORTCODE
Dim smCopyCodeTag As String
Dim tmMissedCode() As SORTCODE
Dim smMissedCodeTag As String
Dim tmCntrSortCode() As SORTCODE
Dim smCntrSortCodeTag As String

'Spot summary file
Dim hmSsf As Integer        'Spot summary file handle
Dim tmSsf As SSF            'SSF record image
Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
Dim tmSsfSrchKey1 As SSFKEY1 'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2 'SSF key record image
Dim imSsfRecLen As Integer     'SSF record length
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim lmSsfDate As Long
Dim lmSsfRecPos As Long
Dim imSSFType As Integer    '0=Program; >0 = game number

Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey3 As LONGKEY0    'SDF record image (key 3)
Dim imSdfRecLen As Integer      'SDF record length
Dim tmSdf As SDF

Dim hmSmf As Integer            'Spot detail file handle
Dim imSmfRecLen As Integer      'SDF record length
Dim tmSmf As SMF

'Game Schedule
Dim hmGsf As Integer
Dim tmGsf As GSF
Dim tmGsfSrchKey3 As GSFKEY3
Dim imGsfRecLen As Integer

Dim hmGhf As Integer
Dim hmSxf As Integer

'Avail Name
Dim hmAnf As Integer        'Avail Name file handle
Dim tmAnf As ANF            'ANF record image
Dim tmAnfSrchKey As INTKEY0 'ANF key record image
Dim imAnfRecLen As Integer  'ANF record length

'Media code record information
Dim hmMcf As Integer        'Contract line file handle
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer  'MCF record length
Dim tmMcf As MCF            'MCF record image

'Rotation
Dim hmCrf As Integer            'Short Title file handle
Dim tmCrf As CRF               'SIF record image
Dim tmCrfSrchKey1 As CRFKEY1
Dim imCrfRecLen As Integer         'SIF record length
Dim tmOldCrf() As CRF

'Copy inventory record information
Dim hmCif As Integer        'Copy line file handle
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer  'CIF record length
Dim tmCif As CIF            'CIF record image

'Copy/Product
Dim hmCpf As Integer
Dim tmCpf As CPF
Dim tmCpfSrchKey As LONGKEY0 'CPEF key record image
Dim imCpfRecLen As Integer  'CPF record length

'Copy/Product
Dim hmCnf As Integer
Dim tmCnf As CNF
Dim tmCnfSrchKey As CNFKEY0 'CNF key record image
Dim imCnfRecLen As Integer  'CNF record length

'Library calendar
Dim hmLcf As Integer            'Log calendar library file handle
Dim tmLcf As LCF               'LCF record image-delete
Dim tmLcfSrchKey As LCFKEY0     'LCF key record image
Dim imLcfRecLen As Integer         'LCF record length

'Library version
Dim hmLvf As Integer            'Log library file handle
Dim tmLvf As LVF                'LVF record image
Dim tmLvfSrchKey As LONGKEY0     'LVF key record image
Dim imLvfRecLen As Integer         'LVF record length

'Library Title
Dim hmLtf As Integer            'Log title file handle
Dim tmLtf As LTF                'LTF record image
Dim tmLtfSrchKey As INTKEY0     'LTF key record image
Dim imLtfRecLen As Integer         'LTF record length

'Contract
Dim tmChf As CHF        'Chf record image
Dim tmChfSrchKey As LONGKEY0    'Chf key record image
Dim hmCHF As Integer    'Contract file handle
Dim imCHFRecLen As Integer        'CHF record length

'Contract line
Dim hmClf As Integer        'Contract line file handle
Dim tmClf As CLF            'CLF record image
Dim tmClfSrchKey As CLFKEY0 'CLF key record image
Dim imClfRecLen As Integer     'CLF record length
Dim tmAddClf() As CLFLIST

Dim tmFCff() As CFF
Dim tmCff As CFF

' Rate Card Programs/Times File
Dim hmRdf As Integer        'Rate Card Programs/Times file handle
Dim tmRdf As RDF            'RDF record image
Dim tmRdfSrchKey As INTKEY0 'RDF key record image
Dim imRdfRecLen As Integer     'RDF record length

'Feed
Dim hmFsf As Integer
Dim tmFsf As FSF            'FSF record image
Dim tmFSFSrchKey As LONGKEY0 'FSF key record image
Dim imFsfRecLen As Integer     'FSF record length

'Live Log
Dim tmLlf As LLF        'Llf record image
Dim tmLlfSrchKey As LONGKEY0    'Llf key record image
Dim tmLLFSrchKey1 As LLFKEY1
Dim hmLlf As Integer    'Live Log file handle
Dim imLlfRecLen As Integer        'LLF record length

'Live Log Detail
Dim tmLdf As LDF        'Llf record image
Dim tmLDFSrchKey1 As LDFKEY1
Dim hmLdf As Integer    'Live Log file handle
Dim imLdfRecLen As Integer        'LLF record length

'Missed Reason
Dim hmMnf As Integer        'Avail Name file handle
Dim tmMnf As MNF            'ANF record image
Dim imMnfRecLen As Integer  'ANF record length

Dim hmUrf As Integer
Dim tmUrfSrchKey As INTKEY0
Dim tmUrf As URF
Dim imUrfRecLen As Integer

'Feed Name
Dim hmFnf As Integer

'Product
Dim hmPrf As Integer

'Vehicle link
Dim tmVlf As VLF            'VLF record image
Dim hmVLF As Integer        'VLF Handle
Dim imVlfRecLen As Integer      'VLF record length
Dim tmVlfSrchKey0 As VLFKEY0 'VLF key record image
Dim imVlfDate0 As Integer
Dim imVlfDate1 As Integer
Dim imAirVefCode() As Integer

Dim tmLang() As MNF
Dim smLangTag As String

Dim tmTeam() As MNF
Dim smTeamTag As String

'6/9/14
Dim smEventTitle1 As String
Dim smEventTitle2 As String

Dim hmSaf As Integer
Dim tmSaf As SAF            'Schedule Attributes record image
Dim imSafRecLen As Integer

Const VEHICLENAMEINDEX = 0
Const LIBSTARTTIMEINDEX = 1
Const LIBLENGTHINDEX = 2
Const LIBEVENTINDEX = 3
Const STATIONCODEINDEX = 4
Const VEHICLECODEINDEX = 5
Const GAMENOINDEX = 6
Const GAMETEAMSINDEX = 7
Const GAMEFEEDINDEX = 8
Const GAMELANGINDEX = 9
Const LIBSORTINDEX = 10
Const FULLTEAMNAMES = 11

Const AIREDINDEX = 0
Const NOTAIREDINDEX = 1
Const CERTIYINDEX = 2
Const SCHDTIMEINDEX = 3
Const LENGTHINDEX = 4
Const TYPEINDEX = 5
Const ADVTPRODINDEX = 6
Const COPYINDEX = 7
Const SDFCODEINDEX = 8
Const SDFSORTINDEX = 9
Const POSITIONINDEX = 10
Const VEHICLETYPEINDEX = 11
Const ORIGTIMEINDEX = 12
Const ORIGCOPYINDEX = 13
Const BKQHINDEX = 14




Private Sub cmcAddSpot_Click()
    Dim ilRet As Integer
    Dim ilChf As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilAAS As Integer
    Dim slStatus As String
    Dim slCntrType As String
    Dim ilCurrent As Integer
    Dim ilHOType As Integer
    Dim ilShow As Integer
    Dim llChfCode As Long
    Dim ilLineNo As Integer
    Dim ilPass As Integer
    Dim llSpotTime As Long
    Dim llSdfRecPos As Long
    Dim ilFreedom As Integer
    Dim ilPriceLevel As Integer
    Dim slSchStatus As String
    Dim ilAvailIndex As Integer

    LlAddSpt.Show vbModal
    If igTerminateReturn = 1 Then
        'Add Spot
        ilLineNo = 0
        tmCifSrchKey.lCode = lgAddCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            llChfCode = mGetCntrFromRot(tmCif.iLen)
            If llChfCode > 0 Then
                ilLineNo = mChkCopyCntr(llChfCode, tmCif.iLen)
            End If
            If ilLineNo = 0 Then
                For ilPass = 0 To 1 Step 1
                    ilAAS = 0
                    slStatus = "HO"
                    slCntrType = ""
                    ilCurrent = ilPass
                    ilHOType = 1
                    ilShow = 0
                    ilRet = gPopCntrForAASBox(LiveLog, ilAAS, igAddAvtCode, slStatus, slCntrType, ilCurrent, ilHOType, ilShow, lbcCntr, tmCntrSortCode(), smCntrSortCodeTag)
                    For ilChf = LBound(tmCntrSortCode) To UBound(tmCntrSortCode) - 1 Step 1
                        slNameCode = tmCntrSortCode(ilChf).sKey
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        llChfCode = Val(slCode)
                        ilLineNo = mChkCopyCntr(llChfCode, tmCif.iLen)
                        If ilLineNo > 0 Then
                            Exit For
                        End If
                    Next ilChf
                    If ilLineNo > 0 Then
                        Exit For
                    End If
                Next ilPass
            End If
        End If
        If (llChfCode > 0) And (ilLineNo > 0) Then
            ilRet = mReadChfClfRdfRec(llChfCode, ilLineNo, 0)
            If ilRet Then
                'Make spot
                ilRet = mMakeUnschSpot(llChfCode, ilLineNo, 0, imVefCode)
                tmSdfSrchKey3.lCode = tmSdf.lCode
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                ilRet = btrGetPosition(hmSdf, llSdfRecPos)
                'Find avail
                llSpotTime = gTimeToLong(sgAddTime, False)
                ilAvailIndex = mAddSpotFindAvailIndex(llSpotTime, tmSdf.iLen)
                If ilAvailIndex <> -1 Then
                    'Book Spot
                    slSchStatus = "O"
                    If tmChf.sType = "S" Then
                        ilFreedom = 1060
                    ElseIf tmChf.sType = "M" Then
                        ilFreedom = 1050
                    ElseIf tmChf.sType = "Q" Then
                        ilFreedom = 1030
                    Else
                        ilFreedom = 1045
                    End If
                    ilPriceLevel = 0
                    ilRet = gBookSpot(slSchStatus, hmSdf, tmSdf, llSdfRecPos, ilFreedom, hmSsf, tmSsf, lmSsfRecPos, ilAvailIndex, -1, tmChf, tmClf, tmRdf, imVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, ilPriceLevel, False, hmGsf)
                    If (ilRet) Then
                        If (lgAddCopyCode > 0) Then
                            tmSdf.sPtType = "1"
                            tmSdf.lCopyCode = lgAddCopyCode
                        End If
                        gPackTime sgAddTime, tmSdf.iTime(0), tmSdf.iTime(1)
                        tmSdf.sAffChg = "Y"
                        tmSdf.iUrfCode = tgUrf(0).iCode
                        ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                        grdLogged.Redraw = False
                        tmAvail.ianfCode = 0
                        mAddSpotToGrid grdLogged, 1, "", ""
                        grdLogged.Redraw = True
                    Else
                        MsgBox "Unable to Add Bonus Spots", vbOKOnly + vbExclamation, "Bonus Not Saved"
                    End If
                Else
                    MsgBox "Unable to Add Bonus Spots", vbOKOnly + vbExclamation, "Bonus Not Saved"
                End If
            Else
                MsgBox "Unable to Add Bonus Spots", vbOKOnly + vbExclamation, "Bonus Not Saved"
            End If
        Else
            MsgBox "Unable to Add Bonus Spots as Contract could not be Found", vbOKOnly + vbExclamation, "Bonus Not Saved"
        End If
    End If
End Sub

Private Sub cmcAddSpot_GotFocus()
    mSetShow True
End Sub

Private Sub cmcChgFeed_Click()
    Dim ilRet As Integer
    Dim slEndTime As String
    Dim slXMid As String
    Dim slNowDate As String
    Dim slEventTitle1 As String
    Dim slEventTitle2 As String

    sgLLCurrentFeed = smCurrentFeed
    igLLCurrentVefCode = imCurrentVefCode
    LLFeed.Show vbModal
    If igLLFeedReturn Then
        tmLdf.lCode = 0
        tmLdf.sType = "F"
        tmLdf.sSubType = "E"
        gPackTime sgLLFeedEndTime, tmLdf.iTime(0), tmLdf.iTime(1)
        tmLdf.lLlfCode = tmLlf.lCode
        tmLdf.iUrfCode = tgUrf(0).iCode
        tmLdf.sCurrentFeed = smCurrentFeed
        slNowDate = gNow()
        tmLdf.lEnteredDate = gDateValue(Format$(slNowDate, "m/d/yy"))
        tmLdf.lEnteredTime = gTimeToLong(Format$(slNowDate, "h:mm:ssAM/PM"), False)
        tmLdf.sUnused = ""
        ilRet = btrInsert(hmLdf, tmLdf, imLdfRecLen, INDEXKEY0)
        'If smCurrentFeed = "H" Then
        '    smCurrentFeed = "V"
        'ElseIf smCurrentFeed = "V" Then
        '    smCurrentFeed = "H"
        'ElseIf smCurrentFeed = "N" Then
        '    smCurrentFeed = "V"
        'End If
        smCurrentFeed = sgLLCurrentFeed
        imCurrentVefCode = igLLCurrentVefCode
        tmLdf.lCode = 0
        tmLdf.sType = "F"
        tmLdf.sSubType = "S"
        gPackTime sgLLFeedStartTime, tmLdf.iTime(0), tmLdf.iTime(1)
        tmLdf.lLlfCode = tmLlf.lCode
        tmLdf.iUrfCode = tgUrf(0).iCode
        tmLdf.sCurrentFeed = smCurrentFeed
        tmLdf.lEnteredDate = gDateValue(Format$(slNowDate, "m/d/yy"))
        tmLdf.lEnteredTime = gTimeToLong(Format$(slNowDate, "h:mm:ssAM/PM"), False)
        tmLdf.sUnused = ""
        ilRet = btrInsert(hmLdf, tmLdf, imLdfRecLen, INDEXKEY0)
        gAddTimeLength Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)), Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBLENGTHINDEX)), "A", "1", slEndTime, slXMid
        'lacBanner(1).Caption = " " & Format$(lmNowDate, "m/d/yy") & " " & Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)) & "-" & slEndTime
        lacBanner(1).Caption = " " & Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)) & "-" & slEndTime
        If imGameNo > 0 Then
            If (Asc(tgSpf.sSportInfo) And USINGFEED) = USINGFEED Then
                gGetEventTitles imCurrentVefCode, slEventTitle1, slEventTitle2
                If smCurrentFeed = "V" Then
                    lacBanner(1).Caption = lacBanner(1).Caption & " Feed: " & slEventTitle1    'Visiting"
                ElseIf smCurrentFeed = "H" Then
                    lacBanner(1).Caption = lacBanner(1).Caption & " Feed: " & slEventTitle2    'Home"
                ElseIf smCurrentFeed = "N" Then
                    lacBanner(1).Caption = lacBanner(1).Caption & " Feed: National"
                End If
            End If
            If ((Asc(tgSpf.sSportInfo) And USINGLANG) = USINGLANG) And (Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMELANGINDEX)) <> "") Then
                lacBanner(1).Caption = lacBanner(1).Caption & " Language: " & Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMELANGINDEX))
            End If
        End If
    End If
End Sub

Private Sub cmcContinue_Click()
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRet As Integer
    Dim ilStatus As Integer
    Dim slLcfStartTime As String
    Dim llLcfStartTime As Long
    Dim llStartTime As Long
    Dim tlLlf As LLF

    If grdVehicleSelection.Row < grdVehicleSelection.FixedRows Then
        Exit Sub
    End If
    mCheckGG
    If igGGFlag = 0 Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imVefCode = Val(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, VEHICLECODEINDEX))
    imVpfIndex = gBinarySearchVpf(imVefCode)
    imGameNo = Val(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMENOINDEX))
    mBuildVehicleArray
    imSignOn = 0
    imGameStart = 0
    lmEnableRow = -1
    lmEnableCol = -1
    imWhichGrid = -1
    smGameStartTime = ""
    smGameEndTime = ""
    imAllowUndo = True
    ReDim tmDelOrMissedSdf(0 To 0) As SDF
    If grdVehicleSelection.CellForeColor <> vbBlack Then
        ilStatus = 0
        tlLlf.lEnteredDate = 0
        tlLlf.lEnteredTime = 0
        slLcfStartTime = grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)
        llLcfStartTime = gTimeToLong(slLcfStartTime, False)
        gPackDateLong lmEventDate, ilDate0, ilDate1   'lmNowDate, ilDate0, ilDate1
        tmLLFSrchKey1.iAirDate(0) = ilDate0
        tmLLFSrchKey1.iAirDate(1) = ilDate1
        tmLLFSrchKey1.iVefCode = imVefCode
        ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLLFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        Do While (ilRet = BTRV_ERR_NONE) And (tmLlf.iVefCode = imVefCode) And (tmLlf.iAirDate(0) = ilDate0) And (tmLlf.iAirDate(1) = ilDate1)
            If (tmLlf.iGameNo = imGameNo) Then
                gUnpackTimeLong tmLlf.iLcfStartTime(0), tmLlf.iLcfStartTime(1), False, llStartTime
                If llStartTime = llLcfStartTime Then
                    If tmLlf.sLogCompleted = "Y" Then
                        imAllowUndo = False
                        ilStatus = 2
                    Else
                        If ilStatus <> 2 Then
                            ilStatus = 1
                        End If
                    End If
                    'Save newest LLF
                    If ((tmLlf.lEnteredTime > tlLlf.lEnteredTime) And (tmLlf.lEnteredDate = tlLlf.lEnteredDate)) Or (tmLlf.lEnteredDate > tlLlf.lEnteredDate) Then
                        tlLlf = tmLlf
                    End If
                    tmLDFSrchKey1.lLlfCode = tmLlf.lCode
                    ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlf.lCode)
                        If (tmLdf.sType = "G") Or (tmLdf.sType = "P") Then
                            If tmLdf.sSubType = "S" Then
                                gUnpackTime tmLdf.iTime(0), tmLdf.iTime(1), "A", "1", smGameStartTime
                                smActualStartDate = Format$(tmLdf.lEnteredDate, "m/d/yy")
                            Else
                                gUnpackTime tmLdf.iTime(0), tmLdf.iTime(1), "A", "1", smGameEndTime
                            End If
                        End If
                        ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
            End If
            ilRet = btrGetNext(hmLlf, tmLlf, imLlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If ilStatus > 0 Then
            Screen.MousePointer = vbDefault
            tmLlf = tlLlf
            If ilStatus = 2 Then
                ilRet = MsgBox("Day completely Logged, Continue", vbYesNo + vbExclamation, "Log Done")
                If ilRet = vbNo Then
                    Exit Sub
                End If
                imSignOn = 2
            End If
            If ilStatus = 1 Then
                If tmLlf.sSignOffStatus <> "Y" Then
                    ilRet = MsgBox("Day partially Logged by " & Trim$(tmLlf.sName) & " in " & Trim$(tmLlf.sStudioID) & ", Replace", vbYesNo + vbQuestion, "Partial Log")
                    If ilRet = vbYes Then
                        imSignOn = 2
                    Else
                        Exit Sub
                    End If
                Else
                    imSignOn = 2
                End If
            End If
        End If
    Else
        tmLlf.lCode = 0
    End If
    If (imGameNo <= 0) Or ((Asc(tgSpf.sSportInfo) And USINGFEED) <> USINGFEED) Then
        cmcChgFeed.Visible = False
    Else
        cmcChgFeed.Visible = True
    End If
    pbcLog.Visible = True
    Form_Resize
    pbcVehicleSelection.Visible = False
    csiLiveLogDate.SetEnabled False
    mPopLog
    Screen.MousePointer = vbDefault

    cmcSignOn_Click
End Sub

Private Sub cmcDropDown_Click()
    Select Case lmEnableCol
        Case AIREDINDEX
            plcTme.Visible = Not plcTme.Visible
        Case NOTAIREDINDEX
            lbcNotAired.Visible = Not lbcNotAired.Visible
        Case COPYINDEX
            lbcCopy.Visible = Not lbcCopy.Visible
    End Select
    edcDropdown.SelStart = 0
    edcDropdown.SelLength = Len(edcDropdown.Text)
    edcDropdown.SetFocus
End Sub

Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcExit_Click()
    Unload LiveLog
End Sub



Private Sub cmcSignOff_Click()
    If mExitOk() Then
        'cmcChgFeed.Enabled = False
        'imCheckOnExit = False
        'cmcExit_Click
        imSignOn = 0
        mPopVehSel
        pbcVehicleSelection.Visible = True
        csiLiveLogDate.SetEnabled True
        Form_Resize
        pbcLog.Visible = False
    End If
    imCheckOnExit = True

End Sub

Private Sub cmcSignOff_GotFocus()
    mSetShow True
End Sub

Private Sub cmcSignOn_Click()
    Dim ilRet As Integer
    Dim slTime As String
    Dim llCol As Long
    Dim llRow As Long
    Dim slLcfStartTime As String
    Dim slNowDate As String
    Dim tlLlf As LLF

    sgSignOnMsg = "Sign-On: " & Trim$(lacBanner(0).Caption)
    If imSignOn = 1 Then
        sgReplaceMsg = "Replacing: " & sgLLName & " in Studio " & sgStudioID
    Else
        sgReplaceMsg = ""
    End If
    LLSignOn.Show vbModal
    If igTerminateReturn = 1 Then
        If imSignOn = 1 Then
            'Sign user off
            'Do
                'Do
                '    tmLDFSrchKey1.lLlfCode = tmLlf.lCode
                '    ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                '    Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlf.lCode)
                '        If (tmLdf.sType = "S") Then
                '            slTime = Format$(gNow(), "h:mm:ssAM/PM")
                '            gPackTime slTime, tmLdf.iEndTime(0), tmLdf.iEndTime(1)
                '            ilRet = btrUpdate(hmLdf, tmLdf, imLdfRecLen)
                '            Exit Do
                '        End If
                '        ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                '    Loop
                'Loop While ilRet = BTRV_ERR_CONFLICT
            tmLdf.lCode = 0
            tmLdf.sType = "S"
            tmLdf.sSubType = "E"
            slTime = Format$(gNow(), "h:mm:ssAM/PM")
            gPackTime slTime, tmLdf.iTime(0), tmLdf.iTime(1)
            tmLdf.lLlfCode = tmLlf.lCode
            tmLdf.iUrfCode = tgUrf(0).iCode
            tmLdf.sCurrentFeed = smCurrentFeed
            slNowDate = gNow()
            tmLdf.lEnteredDate = gDateValue(Format$(slNowDate, "m/d/yy"))
            tmLdf.lEnteredTime = gTimeToLong(Format$(slNowDate, "h:mm:ssAM/PM"), False)
            tmLdf.sUnused = ""
            ilRet = btrInsert(hmLdf, tmLdf, imLdfRecLen, INDEXKEY0)
            Do
                tmLlfSrchKey.lCode = tmLlf.lCode
                ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                If ilRet = BTRV_ERR_NONE Then
                    tmLlf.sSignOffStatus = "Y"
                    ilRet = btrUpdate(hmLlf, tmLlf, imLlfRecLen)
                Else
                    Exit Do
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
        tlLlf = tmLlf
        tmLlf.lCode = 0
        tmLlf.sName = sgLLName
        tmLlf.sStudioID = sgStudioID
        tmLlf.iVefCode = imVefCode
        tmLlf.iGameNo = imGameNo
        gPackDateLong lmEventDate, tmLlf.iAirDate(0), tmLlf.iAirDate(1)   'lmNowDate, tmLlf.iAirDate(0), tmLlf.iAirDate(1)
        'slTime = Format$(gNow(), "h:mm:ssAM/PM")
        'gPackTime slTime, tmLlf.iSignonTime(0), tmLlf.iSignonTime(1)
        'gPackTime "", tmLlf.iSignOffTime(0), tmLlf.iSignOffTime(1)
        'If imSignOn = 0 Then
        '    gPackTime "", tmLlf.iStartTime(0), tmLlf.iStartTime(1)
        '    gPackTime "", tmLlf.iEndTime(0), tmLlf.iEndTime(1)
        'End If
        tmLlf.sSignOffStatus = "N"
        tmLlf.sReplaceStatus = "N"
        tmLlf.lReplaceLlfCode = 0
        tmLlf.sLogCompleted = "N"
        slLcfStartTime = grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)
        gPackTime slLcfStartTime, tmLlf.iLcfStartTime(0), tmLlf.iLcfStartTime(1)
        slNowDate = gNow()
        tmLlf.lEnteredDate = gDateValue(Format$(slNowDate, "m/d/yy"))
        tmLlf.lEnteredTime = gTimeToLong(Format$(slNowDate, "h:mm:ssAM/PM"), False)
        ilRet = btrInsert(hmLlf, tmLlf, imLlfRecLen, INDEXKEY0)
        If (imSignOn = 1) Or (imSignOn = 2) Then
            Do
                tmLlfSrchKey.lCode = tlLlf.lCode
                ilRet = btrGetEqual(hmLlf, tlLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                If ilRet = BTRV_ERR_NONE Then
                    tlLlf.sReplaceStatus = "Y"
                    tlLlf.lReplaceLlfCode = tmLlf.lCode
                    ilRet = btrUpdate(hmLlf, tlLlf, imLlfRecLen)
                Else
                    Exit Do
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
        tmLdf.lCode = 0
        tmLdf.sType = "S"
        tmLdf.sSubType = "S"
        slTime = Format$(gNow(), "h:mm:ssAM/PM")
        gPackTime slTime, tmLdf.iTime(0), tmLdf.iTime(1)
        tmLdf.lLlfCode = tmLlf.lCode
        tmLdf.iUrfCode = tgUrf(0).iCode
        tmLdf.sCurrentFeed = smCurrentFeed
        tmLdf.lEnteredDate = gDateValue(Format$(slNowDate, "m/d/yy"))
        tmLdf.lEnteredTime = gTimeToLong(Format$(slNowDate, "h:mm:ssAM/PM"), False)
        tmLdf.sUnused = ""
        ilRet = btrInsert(hmLdf, tmLdf, imLdfRecLen, INDEXKEY0)
        llRow = grdLogged.Rows
        If llRow - 1 = grdLogged.FixedRows Then
            If grdLogged.TextMatrix(grdLogged.FixedRows, ADVTPRODINDEX) = "" Then
                llRow = grdLogged.FixedRows
            Else
                grdLogged.AddItem ""
            End If
        Else
            If llRow >= grdLogged.Rows Then
                grdLogged.AddItem ""
            End If
        End If
        grdLogged.Row = llRow
        grdLogged.Col = SCHDTIMEINDEX
        grdLogged.CellBackColor = LIGHTYELLOW
        grdLogged.Col = LENGTHINDEX
        grdLogged.CellBackColor = LIGHTYELLOW
        grdLogged.Col = TYPEINDEX
        grdLogged.CellBackColor = LIGHTYELLOW
        grdLogged.Col = ADVTPRODINDEX
        grdLogged.CellBackColor = LIGHTYELLOW
        For llCol = 0 To grdLogged.Cols - 1 Step 1
            grdLogged.Col = llCol
            grdLogged.CellAlignment = flexAlignLeftTop
        Next llCol
        grdLogged.TextMatrix(llRow, TYPEINDEX) = "Sign-On"
        grdLogged.TextMatrix(llRow, AIREDINDEX) = slTime
        grdLogged.TextMatrix(llRow, ADVTPRODINDEX) = sgLLName
        For llCol = 0 To grdLogged.Cols - 1 Step 1
            grdLogged.Col = llCol
            grdLogged.CellForeColor = DARKGREEN 'vbGreen
        Next llCol
        imSignOn = 1
        cmcSignOff.Enabled = True
        cmcSignOn.Enabled = True    'False
        cmcChgFeed.Enabled = True
        '4/29/11: Set day or game as Incomplete
        ilRet = mUpdateLCFAffPost(tmLlf, "I")
    End If
End Sub

Private Sub cmcSignOn_GotFocus()
    mSetShow True
End Sub

Private Sub cmcUndo_Click()
    Dim slCopy As String

    If imWhichGrid = 0 Then
        Select Case lmEnableCol
            Case AIREDINDEX
                edcDropdown.Text = ""
                mSetShow False
            Case NOTAIREDINDEX
                lbcNotAired.ListIndex = -1
                edcDropdown.Text = ""
                mSetShow False
            Case COPYINDEX
                slCopy = grdLog.TextMatrix(lmEnableRow, ORIGCOPYINDEX)
                gFindMatch slCopy, 0, lbcCopy
                If gLastFound(lbcCopy) >= 0 Then
                    lbcCopy.ListIndex = gLastFound(lbcCopy)
                    edcDropdown.Text = lbcCopy.List(lbcCopy.ListIndex)
                    mSetShow False
                End If
        End Select
    Else
        Select Case lmEnableCol
            Case AIREDINDEX
                edcDropdown.Text = ""
                mSetShow True
            Case NOTAIREDINDEX
                lbcNotAired.ListIndex = -1
                edcDropdown.Text = ""
                mSetShow True
            Case COPYINDEX
                slCopy = grdLogged.TextMatrix(lmEnableRow, ORIGCOPYINDEX)
                gFindMatch slCopy, 0, lbcCopy
                If gLastFound(lbcCopy) >= 0 Then
                    lbcCopy.ListIndex = gLastFound(lbcCopy)
                    edcDropdown.Text = lbcCopy.List(lbcCopy.ListIndex)
                    mSetShow False
                End If
        End Select
        pbcClickFocus.SetFocus
    End If
End Sub

Private Sub csiLiveLogDate_Change()
    tmcDelay.Enabled = False
    tmcDelay.Enabled = True
End Sub

Private Sub csiLiveLogDate_GotFocus()
    tmcDelay.Enabled = False
End Sub

Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilPosA As Integer
    Dim ilPosP As Integer

    Select Case lmEnableCol
        Case AIREDINDEX
            slStr = Trim$(edcDropdown.Text)
            If (InStr(1, slStr, "a", vbTextCompare) <= 0) And (InStr(1, slStr, "p", vbTextCompare) <= 0) And (Len(slStr) > 0) Then
                If Hour(Time$()) < 12 Then
                    slStr = slStr & "a"
                Else
                    slStr = slStr & "p"
                End If
                edcDropdown.Text = slStr
                edcDropdown.SelStart = Len(slStr) - 1
            Else
                If (InStr(1, slStr, "a", vbTextCompare) > 0) And (InStr(1, slStr, "p", vbTextCompare) > 0) And (Len(slStr) > 0) Then
                    ilPosA = InStr(1, slStr, "a", vbTextCompare)
                    ilPosP = InStr(1, slStr, "p", vbTextCompare)
                    If ilPosA < ilPosP Then
                        edcDropdown.Text = Left$(slStr, ilPosA)
                        edcDropdown.SelStart = ilPosA - 1
                    Else
                        edcDropdown.Text = Left$(slStr, ilPosP)
                        edcDropdown.SelStart = ilPosP - 1
                    End If
                End If
            End If
        Case NOTAIREDINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropdown, lbcNotAired, imBSMode, slStr)
            If ilRet = 1 Then
                'If lbcNotAired.ListCount > 0 Then
                '    lbcNotAired.ListIndex = 0
                'End If
            End If
        Case COPYINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropdown, lbcCopy, imBSMode, imComboBoxIndex
    End Select
    imLbcArrowSetting = False
End Sub

Private Sub edcDropDown_GotFocus()
    Select Case lmEnableCol
        Case AIREDINDEX
        Case NOTAIREDINDEX
        Case COPYINDEX
    End Select
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
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropdown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case lmEnableCol
        Case AIREDINDEX
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
        Case NOTAIREDINDEX
        Case COPYINDEX
    End Select
End Sub

Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case lmEnableCol
            Case AIREDINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case NOTAIREDINDEX
                gProcessArrowKey Shift, KeyCode, lbcNotAired, imLbcArrowSetting
            Case COPYINDEX
                gProcessArrowKey Shift, KeyCode, lbcCopy, imLbcArrowSetting
        End Select
        edcDropdown.SelStart = 0
        edcDropdown.SelLength = Len(edcDropdown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case lmEnableCol
            Case AIREDINDEX
            Case NOTAIREDINDEX
            Case COPYINDEX
        End Select
    End If
End Sub

Private Sub Form_Activate()
    sgDoneMsg = ""
End Sub


Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    mCheckGG
    If igGGFlag = 0 Then
        tmcTerminate.Enabled = True
    Else
        tmcDelay_Timer
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSetShow True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (pbcLog.Visible) And (imCheckOnExit) Then
        If Not mExitOk() Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim llWidth As Long
    If pbcVehicleSelection.Visible Then
        If LiveLog.Height - 2 * pbcBanner.Height - 120 < 0 Then
            Exit Sub
        End If
        pbcVehicleSelection.Height = LiveLog.Height - 2 * pbcBanner.Height - 120
        pbcVehicleSelection.Width = 0.95 * LiveLog.Width
        grdVehicleSelection.Height = pbcVehicleSelection.Height - 1.2 * cmcContinue.Height
        grdVehicleSelection.Width = pbcVehicleSelection.Width - 30 '0.9 * pbcVehicleSelection.Width
        pbcVehicleSelection.Top = pbcBanner.Top + pbcBanner.Height + 60     '(LiveLog.Height - pbcVehicleSelection.Height) / 2 - 240
        pbcVehicleSelection.Left = (LiveLog.Width - pbcVehicleSelection.Width) / 2 - 60
        grdVehicleSelection.Top = 0 '(pbcVehicleSelection.Height - grdVehicleSelection.Height) / 2 - 240
        grdVehicleSelection.Left = 0    '(pbcVehicleSelection.Width - grdVehicleSelection.Width) / 2 - 60
        cmcContinue.Move pbcVehicleSelection.Width / 2 - 2 * cmcContinue.Width, pbcVehicleSelection.Height - cmcContinue.Height - 30 'grdVehicleSelection.Top + grdVehicleSelection.Height + cmcContinue.Height / 2
        cmcExit.Move pbcVehicleSelection.Width / 2 + cmcContinue.Width, cmcContinue.Top
        mVehicleSelectionColumnWidths
        lacHelp.Left = pbcVehicleSelection.Left + pbcVehicleSelection.Width - lacHelp.Width
        lacBanner(0).Left = pbcVehicleSelection.Left
        lacBanner(1).Left = lacBanner(0).Left
        lacBanner(2).Left = lacBanner(0).Left
        csiLiveLogDate.Left = lacBanner(0).Left + lacBanner(2).Width
    End If
    If pbcLog.Visible Then
        If LiveLog.Height - 2 * pbcBanner.Height - 120 < 0 Then
            Exit Sub
        End If
        pbcLog.Height = LiveLog.Height - 2 * pbcBanner.Height - 120
        pbcLog.Width = 0.95 * LiveLog.Width
        grdLog.Height = 3 * (pbcLog.Height - 1.2 * cmcSignOn.Height) / 4
        grdLog.Width = pbcLog.Width - 30    '0.9 * pbcLog.Width
        grdLogged.Height = (pbcLog.Height - 1.2 * cmcSignOn.Height) / 4 - 180   '120
        grdLogged.Width = pbcLog.Width - 30    '0.9 * pbcLog.Width
        pbcLog.Top = pbcBanner.Top + pbcBanner.Height + 60     '(LiveLog.Height - pbcLog.Height) / 2 - 240
        pbcLog.Left = (LiveLog.Width - pbcLog.Width) / 2 - 60
        grdLog.Top = 0  '(pbcLog.Height - grdLog.Height) / 2 - 240
        grdLog.Left = 0 '(pbcLog.Width - grdLog.Width) / 2 - 60
        grdLogged.Top = grdLog.Top + grdLog.Height + 120    '60
        grdLogged.Left = 0
        cmcSignOn.Move grdLog.Left, pbcLog.Height - cmcSignOn.Height - 30   'grdLog.Top + grdLog.Height + cmcSignOn.Height / 2
        cmcAddSpot.Move grdLog.Left + grdLog.Width - cmcAddSpot.Width, cmcSignOn.Top
        If (imGameNo <= 0) Or ((Asc(tgSpf.sSportInfo) And USINGFEED) <> USINGFEED) Then
            llWidth = (cmcAddSpot.Left - (cmcSignOn.Left + cmcSignOn.Width) - cmcSignOff.Width) \ 3
        Else
            llWidth = (cmcAddSpot.Left - (cmcSignOn.Left + cmcSignOn.Width) - cmcSignOff.Width) \ 4
        End If
        If llWidth < 0 Then
            llWidth = 0
        End If
        cmcSignOff.Move cmcSignOn.Left + cmcSignOn.Width + llWidth, cmcSignOn.Top
        cmcUndo.Move cmcSignOff.Left + cmcSignOff.Width + llWidth, cmcSignOn.Top
        If (imGameNo > 0) And ((Asc(tgSpf.sSportInfo) And USINGFEED) = USINGFEED) Then
            cmcChgFeed.Move cmcUndo.Left + llWidth, cmcSignOn.Top
        End If
        mLogColumnWidths
        lacSpacer.Move 0, grdLog.Top + grdLog.Height, pbcLog.Width, grdLogged.Top - grdLog.Top - grdLog.Height
        If (lmEnableRow >= grdLog.FixedRows) And (lmEnableRow < grdLog.Rows) Then
            If grdLog.RowIsVisible(lmEnableRow) Then
                mSetFocus
            End If
        End If
        lacHelp.Left = pbcLog.Left + pbcLog.Width - lacHelp.Width
        lacBanner(0).Left = pbcLog.Left
        lacBanner(1).Left = lacBanner(0).Left
        lacBanner(2).Left = lacBanner(0).Left
        csiLiveLogDate.Left = lacBanner(0).Left + lacBanner(2).Width
    End If
    pbcBanner.Visible = False

    'imcLogo.Left = pbcBanner.Width - imcLogo.Width
    'lacHelp.Left = imcLogo.Left - lacHelp.Width
    'If imcLogo.Width + lacHelp.Width < pbcBanner.Width Then
    '    lacBanner(0).Width = pbcBanner.Width - imcLogo.Width - lacHelp.Width
    'Else
    '    lacBanner(0).Width = 0
    'End If
    'lacHelp.Left = pbcBanner.Width - lacHelp.Width
    imcLogo.Width = 2460
    imcLogo.Left = lacHelp.Left - imcLogo.Width - lacCopyright.Width - lacHelp.Width / 2
    lacCopyright.Left = imcLogo.Left + imcLogo.Width + 15
    lacCopyright.Caption = "2003-" & Year(Now)
    If imcLogo.Width + lacHelp.Width + lacCopyright.Width + lacHelp.Width / 2 < pbcBanner.Width Then
        lacBanner(0).Width = pbcBanner.Width - imcLogo.Width - lacHelp.Width - lacCopyright.Width - lacHelp.Width / 2
    Else
        lacBanner(0).Width = 0
    End If
    lacBanner(1).Width = lacBanner(0).Width
    pbcBanner.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    Erase igEventAdfCode
    
    Erase imLLVefCode
    Erase tmDelOrMissedSdf
    Erase tmCopyCode
    Erase tmMissedCode
    Erase tmCntrSortCode
    Erase tmOldCrf
    Erase tmFCff
    Erase tmLang
    Erase tmTeam
    Erase imAirVefCode
    
    btrDestroy hmVLF
    btrDestroy hmSsf
    btrDestroy hmSmf
    btrDestroy hmSdf
    btrDestroy hmGsf
    btrDestroy hmGhf
    btrDestroy hmSxf
    btrDestroy hmAnf
    btrDestroy hmMcf
    btrDestroy hmCrf
    btrDestroy hmCif
    btrDestroy hmCpf
    btrDestroy hmCnf
    btrDestroy hmLcf
    btrDestroy hmLvf
    btrDestroy hmLtf
    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmRdf
    btrDestroy hmFsf
    btrDestroy hmLlf
    btrDestroy hmLdf
    btrDestroy hmMnf
    btrDestroy hmFnf
    btrDestroy hmPrf
    
    gEraseGlobalVar True
    Set LiveLog = Nothing   'Remove data segment

    End
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'
'      2-02-01 Move Opens to gcmcFix; right before the
'           generation of summary records.  The timer
'           wasnt activated until after opening; therefore
'           files could remain open and the backup
'           aborted
'*                                                     *
'*******************************************************
Private Sub mInit()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        slStr                                                   *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mInitErr                                                                              *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slDate As String
    imTerminate = False
    'mParseCmmdLine
'Rm**    btrDestroy hgHlf
    Screen.MousePointer = vbHourglass
    mParseCmmdLine
    pbcDateTab.Left = -100 - pbcDateTab.Width
    imSignOn = 0
    imGameStart = 0
    imFixing = False
    imFirstFocus = True
    imBypassFocus = False
    lmEnableRow = -1
    lmEnableCol = -1
    imWhichGrid = -1
    lmTimeWidth = 1200
    lmLengthWidth = 900
    lmStationCode = 720
    imChgMode = False
    imCheckOnExit = True
    imAllowUndo = True
    ReDim imLLVefCode(0 To 0) As Integer
    ReDim imAirVefCode(0 To 0) As Integer
    ReDim igEventAdfCode(0 To 0) As Integer
    smLangTag = ""
    mLanguagePop
    smTeamTag = ""
    mTeamPop
    ReDim tmCopyCode(0 To 0) As SORTCODE
    smCopyCodeTag = ""
    ReDim tmMissedCode(0 To 0) As SORTCODE
    smMissedCodeTag = ""
    mMissedPop
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    smNowDate = slDate
    lmNowDate = gDateValue(slDate)
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), smLastStdMnth
    lmLastStdMnth = gDateValue(smLastStdMnth)

    'lacBanner(1) = " " & "Log Date " & Format$(slDate, "m/d/yy")

    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)

    hmSdf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSdfRecLen = Len(tmSdf)

    hmGsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imGsfRecLen = Len(tmGsf)

    hmGhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)

    hmSxf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)

    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imAnfRecLen = Len(tmAnf)

    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imMcfRecLen = Len(tmMcf)

    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCrfRecLen = Len(tmCrf)

    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCifRecLen = Len(tmCif)

    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCpfRecLen = Len(tmCpf)

    hmCnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCnfRecLen = Len(tmCnf)

    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imLcfRecLen = Len(tmLcf)

    hmLvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imLvfRecLen = Len(tmLvf)

    hmLtf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLtf, "", sgDBPath & "Ltf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imLtfRecLen = Len(tmLtf)

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCHFRecLen = Len(tmChf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imClfRecLen = Len(tmClf)

    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imRdfRecLen = Len(tmRdf)

    hmFsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imFsfRecLen = Len(tmFsf)

    hmLlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmLlf, "", sgDBPath & "Llf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imLlfRecLen = Len(tmLlf)

    hmLdf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmLdf, "", sgDBPath & "Ldf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imLdfRecLen = Len(tmLdf)

    hmMnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imMnfRecLen = Len(tmMnf)

    hmSmf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSmfRecLen = Len(tmSmf)

    hmFnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFnf, "", sgDBPath & "Fnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)

    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)

    hmVLF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imVlfRecLen = Len(tmVlf)

    ilRet = gObtainAdvt()

    ilRet = gObtainAvail()

    mVehicleSelectionColumnNames
    mVehicleSelectionColumnWidths
    'mPopVehSel
    csiLiveLogDate.Text = slDate
    mLogColumnNames

    If (Asc(tgSpf.sSportInfo) And USINGFEED) <> USINGFEED Then
        cmcChgFeed.Visible = False
    End If

    Me.Width = (Screen.Width - 360) / (640 * 15 / Me.Width)
    Me.Height = (Screen.Height - 500) / (480 * 15 / Me.Height)
    gCenterStdAlone Me

    If Not gRecLengthOk("Llf.btr", Len(tmLlf)) Then
        imTerminate = True
        tmcTerminate.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Not gRecLengthOk("Ldf.btr", Len(tmLdf)) Then
        imTerminate = True
        tmcTerminate.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If imTerminate Then
        tmcTerminate.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub











Private Sub mPopVehSel()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slFeed                                                                                *
'******************************************************************************************

    'Populate Vehicle selection box
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilAnyAdded As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim slStartTime As String
    Dim llStartTime As Long
    Dim slLength As String
    Dim llColor As Long
    Dim llPrevColor As Long
    Dim ilLastColSorted As Integer
    Dim ilLastSort As Integer
    Dim ilLcf As Integer
    Dim llLcfStartTime As Long
    Dim ilGameNo As Integer
    Dim slVisitTeam As String
    Dim slHomeTeam As String
    Dim slLanguage As String
    Dim slFeedSource As String
    Dim slFullName As String
    Dim llCurrentFeedDate As Long
    Dim llCurrentFeedTime As Long
    Dim llDate As Long
    Dim llTime As Long
    Dim slGameStatus As String
    Dim ilLogAdded As Integer
    Dim ilIndex As Integer
    Dim slLogStartTime As String
    Dim slLogEndTime As String
    Dim slStr As String
    Dim ilVff As Integer
    Dim slEventTitle1 As String
    Dim slEventTitle2 As String

    ilAnyAdded = False
    grdVehicleSelection.Redraw = False
    ilRet = gObtainVef()
    ilRet = gVpfRead()
    ilRet = gVffRead()
    'grdVehicleSelection.Rows = grdVehicleSelection.FixedRows + 1
    'llRow = grdVehicleSelection.FixedRows
    'For ilCol = 0 To grdVehicleSelection.Cols - 1 Step 1
    '    grdVehicleSelection.TextMatrix(llRow, ilCol) = ""
    'Next ilCol
    mClearVehSelGrid
    llRow = grdVehicleSelection.FixedRows
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        imVpfIndex = gBinarySearchVpf(tgMVef(ilVef).iCode)
        If tgMVef(ilVef).sType <> "L" Then
            If (tgMVef(ilVef).iVefCode > 0) Then
                For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                    If tgMVef(ilVef).iCode = tgVff(ilVff).iVefCode Then
                        If tgVff(ilVff).sMergeTraffic = "M" Then
                            imVpfIndex = -1
                        End If
                        Exit For
                    End If
                Next ilVff
                imVpfIndex = -1
            End If
        End If
        If imVpfIndex >= 0 Then
            'If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G")) And (tgVpf(imVpfIndex).sGenLog = "L") And (tgMVef(ilVef).sState <> "D") Then
            If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G") Or (tgMVef(ilVef).sType = "L")) And ((tgVpf(imVpfIndex).sGenLog = "L") Or (tgVpf(imVpfIndex).sGenLog = "A")) And (tgMVef(ilVef).sState <> "D") Then
                'Check if Todays date is an airing date
                ilLogAdded = False
                slLogStartTime = ""
                slLogEndTime = ""
                llPrevColor = vbBlack
                imVefCode = tgMVef(ilVef).iCode
                
                If tgVpf(imVpfIndex).sGMedium <> "M" And tgVpf(imVpfIndex).sGMedium <> "P" Then
                    '1/25/21 - Dont add Podcast CPM and Podcast Spot vehicles
                    mBuildVehicleArray
                End If
                ilIndex = LBound(imLLVefCode)
                gPackDateLong lmEventDate, ilDate0, ilDate1   'lmNowDate, ilDate0, ilDate1
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                tmSsfSrchKey2.iVefCode = imLLVefCode(ilIndex)    'tgMVef(ilVef).iCode
                tmSsfSrchKey2.iDate(0) = ilDate0
                tmSsfSrchKey2.iDate(1) = ilDate1
                ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                Do
                    If (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = imLLVefCode(ilIndex)) And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1) Then
                        ilGameNo = tmSsf.iType
                        slGameStatus = mGetGameStatus(imLLVefCode(ilIndex), ilGameNo)
                        tmLcfSrchKey.iType = ilGameNo
                        tmLcfSrchKey.sStatus = "C"
                        tmLcfSrchKey.iVefCode = imLLVefCode(ilIndex)
                        tmLcfSrchKey.iLogDate(0) = ilDate0
                        tmLcfSrchKey.iLogDate(1) = ilDate1
                        tmLcfSrchKey.iSeqNo = 1
                        ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                        If (ilRet = BTRV_ERR_NONE) And (slGameStatus <> "C") And (slGameStatus <> "P") Then
                            For ilLcf = LBound(tmLcf.lLvfCode) To UBound(tmLcf.lLvfCode) Step 1
                                If tmLcf.lLvfCode(ilLcf) > 0 Then
                                    tmLvfSrchKey.lCode = tmLcf.lLvfCode(ilLcf)
                                    ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                                    If ilRet = BTRV_ERR_NONE Then
                                        tmLtfSrchKey.iCode = tmLvf.iLtfCode
                                        ilRet = btrGetEqual(hmLtf, tmLtf, imLtfRecLen, tmLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                                        If ilRet = BTRV_ERR_NONE Then
                                            If llRow >= grdVehicleSelection.Rows Then
                                                grdVehicleSelection.AddItem ""
                                            End If
                                            grdVehicleSelection.Row = llRow
                                            llColor = vbBlack
                                            smCurrentFeed = ""
                                            imCurrentVefCode = 0
                                            llCurrentFeedDate = 0
                                            llCurrentFeedTime = 0
                                            gUnpackTime tmLcf.iTime(0, ilLcf), tmLcf.iTime(1, ilLcf), "A", "1", slStartTime
                                            gUnpackLength tmLvf.iLen(0), tmLvf.iLen(1), "3", False, slLength
                                            tmLLFSrchKey1.iAirDate(0) = ilDate0
                                            tmLLFSrchKey1.iAirDate(1) = ilDate1
                                            tmLLFSrchKey1.iVefCode = imVefCode  'tgMVef(ilVef).iCode
                                            ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLLFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                                            Do While (ilRet = BTRV_ERR_NONE) And (tmLlf.iVefCode = tgMVef(ilVef).iCode) And (tmLlf.iAirDate(0) = ilDate0) And (tmLlf.iAirDate(1) = ilDate1)
                                                If tmLlf.iGameNo = ilGameNo Then
                                                    If (Asc(tgSpf.sSportInfo) And USINGFEED) = USINGFEED Then
                                                        tmLDFSrchKey1.lLlfCode = tmLlf.lCode
                                                        ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                                        Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlf.lCode)
                                                            If (tmLdf.sType = "F") And (tmLdf.sSubType = "S") Then
                                                                llDate = tmLdf.lEnteredDate
                                                                llTime = tmLdf.lEnteredTime
                                                                If ((llTime > llCurrentFeedTime) And (llDate = llCurrentFeedDate)) Or (llDate > llCurrentFeedDate) Then
                                                                    smCurrentFeed = tmLdf.sCurrentFeed
                                                                    imCurrentVefCode = imVefCode
                                                                    llCurrentFeedDate = llDate
                                                                    llCurrentFeedTime = llTime
                                                                End If
                                                            End If
                                                            ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                                        Loop
                                                    End If
                                                    gUnpackTimeLong tmLlf.iLcfStartTime(0), tmLlf.iLcfStartTime(1), False, llLcfStartTime
                                                    If gTimeToLong(slStartTime, False) = llLcfStartTime Then
                                                        If tmLlf.sLogCompleted = "Y" Then
                                                            llColor = DARKGREEN 'vbGreen
                                                            'Exit Do
                                                        Else
                                                            If llColor <> DARKGREEN Then
                                                                llColor = vbRed
                                                            End If
                                                        End If
                                                        ''Exit Do
                                                    End If
                                                End If
                                                ilRet = btrGetNext(hmLlf, tmLlf, imLlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                            Loop
                                            If llPrevColor <> llColor Then
                                                If llColor = DARKGREEN Then
                                                    llPrevColor = llColor
                                                ElseIf llColor = vbRed And llPrevColor <> DARKGREEN Then
                                                    llPrevColor = llColor
                                                Else
                                                    If llPrevColor <> vbBlack Then
                                                        If llPrevColor = DARKGREEN Then
                                                            llColor = llPrevColor
                                                        ElseIf llPrevColor = vbRed And llColor <> DARKGREEN Then
                                                            llColor = llPrevColor
                                                        End If
                                                    Else
                                                        llPrevColor = llColor
                                                    End If
                                                End If
                                            End If
                                            grdVehicleSelection.Col = VEHICLENAMEINDEX
                                            grdVehicleSelection.CellForeColor = llColor
                                            grdVehicleSelection.TextMatrix(llRow, VEHICLENAMEINDEX) = Trim$(tgMVef(ilVef).sName)
                                            grdVehicleSelection.Col = LIBSTARTTIMEINDEX
                                            grdVehicleSelection.CellForeColor = llColor
                                            gUnpackTime tmLcf.iTime(0, ilLcf), tmLcf.iTime(1, ilLcf), "A", "1", slStartTime
                                            If tgMVef(ilVef).sType = "L" Then
                                                If slLogStartTime = "" Then
                                                    slLogStartTime = slStartTime
                                                Else
                                                    If gTimeToLong(slStartTime, False) < gTimeToLong(slLogStartTime, False) Then
                                                        slLogStartTime = slStartTime
                                                    End If
                        
                                                End If
                                                grdVehicleSelection.TextMatrix(llRow, LIBSTARTTIMEINDEX) = slLogStartTime
                                            Else
                                                grdVehicleSelection.TextMatrix(llRow, LIBSTARTTIMEINDEX) = slStartTime
                                            End If
                                            gUnpackLength tmLvf.iLen(0), tmLvf.iLen(1), "3", False, slLength
                                            grdVehicleSelection.Col = LIBLENGTHINDEX
                                            grdVehicleSelection.CellForeColor = llColor
                                            If tgMVef(ilVef).sType = "L" Then
                                                gAddTimeLength slStartTime, slLength, "A", "1", slStr, "N"
                                                If slLogEndTime = "" Then
                                                    slLogEndTime = slStr
                                                Else
                                                    If gTimeToLong(slStr, True) > gTimeToLong(slLogEndTime, True) Then
                                                        slLogEndTime = slStr
                                                    End If
                                                End If
                                                slStr = gCurrencyToLength(gTimeToCurrency(gFormatTimeLong(gTimeToLong(slLogEndTime, True) - gTimeToLong(slLogStartTime, False), "A", "1"), True))
                                                grdVehicleSelection.TextMatrix(llRow, LIBLENGTHINDEX) = slStr
                                            Else
                                                grdVehicleSelection.TextMatrix(llRow, LIBLENGTHINDEX) = slLength
                                            End If
                                            grdVehicleSelection.Col = LIBEVENTINDEX
                                            grdVehicleSelection.CellForeColor = llColor
                                            grdVehicleSelection.TextMatrix(llRow, LIBEVENTINDEX) = Trim$(tmLtf.sName)
                                            grdVehicleSelection.Col = STATIONCODEINDEX
                                            grdVehicleSelection.CellForeColor = llColor
                                            grdVehicleSelection.TextMatrix(llRow, STATIONCODEINDEX) = Trim$(tgMVef(ilVef).sCodeStn)
                                            grdVehicleSelection.TextMatrix(llRow, VEHICLECODEINDEX) = Trim$(tgMVef(ilVef).iCode)
                                            If tgMVef(ilVef).sType = "C" Then
                                                grdVehicleSelection.TextMatrix(llRow, GAMENOINDEX) = "0"
                                                grdVehicleSelection.TextMatrix(llRow, GAMETEAMSINDEX) = ""
                                                grdVehicleSelection.TextMatrix(llRow, GAMEFEEDINDEX) = ""
                                                grdVehicleSelection.TextMatrix(llRow, GAMELANGINDEX) = ""
                                            ElseIf tgMVef(ilVef).sType = "L" Then
                                                'grdVehicleSelection.TextMatrix(llRow, LIBSTARTTIMEINDEX) = "12am"
                                                'grdVehicleSelection.TextMatrix(llRow, LIBLENGTHINDEX) = "24h"
                                                grdVehicleSelection.TextMatrix(llRow, LIBEVENTINDEX) = ""
                                                grdVehicleSelection.TextMatrix(llRow, GAMENOINDEX) = "0"
                                                grdVehicleSelection.TextMatrix(llRow, GAMETEAMSINDEX) = ""
                                                grdVehicleSelection.TextMatrix(llRow, GAMEFEEDINDEX) = ""
                                                grdVehicleSelection.TextMatrix(llRow, GAMELANGINDEX) = ""
                                            Else
                                                grdVehicleSelection.TextMatrix(llRow, GAMENOINDEX) = Trim$(Str$(ilGameNo))
                                                mGetTeamAndLang tgMVef(ilVef).iCode, ilGameNo, slVisitTeam, slHomeTeam, slLanguage, slFeedSource, slFullName
                                                grdVehicleSelection.TextMatrix(llRow, GAMETEAMSINDEX) = slVisitTeam & " @" & slHomeTeam
                                                If llCurrentFeedTime > 0 Then
                                                    gGetEventTitles imCurrentVefCode, slEventTitle1, slEventTitle2
                                                    If smCurrentFeed = "H" Then
                                                        slFeedSource = slEventTitle1   '"Home"
                                                    ElseIf smCurrentFeed = "V" Then
                                                        slFeedSource = slEventTitle2   '"Visiting"
                                                    ElseIf smCurrentFeed = "N" Then
                                                        slFeedSource = "National"
                                                    End If
                                                End If
                                                grdVehicleSelection.TextMatrix(llRow, GAMEFEEDINDEX) = slFeedSource
                                                grdVehicleSelection.TextMatrix(llRow, GAMELANGINDEX) = slLanguage
                                                grdVehicleSelection.TextMatrix(llRow, FULLTEAMNAMES) = slFullName
                                            End If
                                            gUnpackTimeLong tmLcf.iTime(0, ilLcf), tmLcf.iTime(1, ilLcf), False, llStartTime
                                            slStartTime = Trim$(Str$(llStartTime))
                                            Do While Len(slStartTime) < 6
                                                slStartTime = "0" & slStartTime
                                            Loop
                                            grdVehicleSelection.TextMatrix(llRow, LIBSORTINDEX) = slStartTime & Trim$(tgMVef(ilVef).sName)
                                            ilAnyAdded = True
                                            ilLogAdded = True
                                            If tgMVef(ilVef).sType <> "L" Then
                                                llRow = llRow + 1
                                            End If
                                        End If
                                    End If
                                End If
                            Next ilLcf
                        End If
                    End If
                    If tgMVef(ilVef).sType = "G" Then
                        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                        ilRet = btrGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit Do
                        Else
                            If tmSsf.iVefCode <> tgMVef(ilVef).iCode Then
                                Exit Do
                            End If
                        End If
                    ElseIf tgMVef(ilVef).sType = "L" Then
                        'If ilLogAdded Then
                        '    Exit Do
                        'Else
                            ilIndex = ilIndex + 1
                            If ilIndex >= UBound(imLLVefCode) Then
                                If ilLogAdded Then
                                    llRow = llRow + 1
                                End If
                                Exit Do
                            End If
                            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                            tmSsfSrchKey2.iVefCode = imLLVefCode(ilIndex)    'tgMVef(ilVef).iCode
                            tmSsfSrchKey2.iDate(0) = ilDate0
                            tmSsfSrchKey2.iDate(1) = ilDate1
                            ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                        'End If
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
    Next ilVef
    grdVehicleSelection.ColWordWrapOption(VEHICLENAMEINDEX) = True
    If Not ilAnyAdded Then
        grdVehicleSelection.TextMatrix(grdVehicleSelection.FixedRows, VEHICLENAMEINDEX) = "No Live Logs"
    Else
        ilLastColSorted = -1
        ilLastSort = -1
        gGrid_SortByCol grdVehicleSelection, VEHICLENAMEINDEX, LIBSORTINDEX, ilLastColSorted, ilLastSort
    End If
    grdVehicleSelection.Row = 0
    grdVehicleSelection.Col = VEHICLECODEINDEX
    grdVehicleSelection.Redraw = True
End Sub

Private Sub mVehicleSelectionColumnNames()
    grdVehicleSelection.Row = 0
    grdVehicleSelection.Col = VEHICLENAMEINDEX
    grdVehicleSelection.CellFontBold = False
    grdVehicleSelection.CellFontName = "Arial"
    grdVehicleSelection.CellFontSize = 6.75
    grdVehicleSelection.CellForeColor = vbBlue
    grdVehicleSelection.TextMatrix(0, VEHICLENAMEINDEX) = "Vehicle Name"
    grdVehicleSelection.Col = LIBSTARTTIMEINDEX
    grdVehicleSelection.CellFontBold = False
    grdVehicleSelection.CellFontName = "Arial"
    grdVehicleSelection.CellFontSize = 6.75
    grdVehicleSelection.CellForeColor = vbBlue
    grdVehicleSelection.TextMatrix(0, LIBSTARTTIMEINDEX) = "Start Time"
    grdVehicleSelection.Col = LIBLENGTHINDEX
    grdVehicleSelection.CellFontBold = False
    grdVehicleSelection.CellFontName = "Arial"
    grdVehicleSelection.CellFontSize = 6.75
    grdVehicleSelection.CellForeColor = vbBlue
    grdVehicleSelection.TextMatrix(0, LIBLENGTHINDEX) = "Length"
    grdVehicleSelection.Col = LIBEVENTINDEX
    grdVehicleSelection.CellFontBold = False
    grdVehicleSelection.CellFontName = "Arial"
    grdVehicleSelection.CellFontSize = 6.75
    grdVehicleSelection.CellForeColor = vbBlue
    grdVehicleSelection.TextMatrix(0, LIBEVENTINDEX) = "Event"
    grdVehicleSelection.Col = STATIONCODEINDEX
    grdVehicleSelection.CellFontBold = False
    grdVehicleSelection.CellFontName = "Arial"
    grdVehicleSelection.CellFontSize = 6.75
    grdVehicleSelection.CellForeColor = vbBlue
    grdVehicleSelection.TextMatrix(0, STATIONCODEINDEX) = "Code"
    grdVehicleSelection.Col = GAMENOINDEX
    grdVehicleSelection.CellFontBold = False
    grdVehicleSelection.CellFontName = "Arial"
    grdVehicleSelection.CellFontSize = 6.75
    grdVehicleSelection.CellForeColor = vbBlue
    grdVehicleSelection.TextMatrix(0, GAMENOINDEX) = "Event No"
    grdVehicleSelection.Col = GAMETEAMSINDEX
    grdVehicleSelection.CellFontBold = False
    grdVehicleSelection.CellFontName = "Arial"
    grdVehicleSelection.CellFontSize = 6.75
    grdVehicleSelection.CellForeColor = vbBlue
    grdVehicleSelection.TextMatrix(0, GAMETEAMSINDEX) = "Teams"
    grdVehicleSelection.Col = GAMEFEEDINDEX
    grdVehicleSelection.CellFontBold = False
    grdVehicleSelection.CellFontName = "Arial"
    grdVehicleSelection.CellFontSize = 6.75
    grdVehicleSelection.CellForeColor = vbBlue
    grdVehicleSelection.TextMatrix(0, GAMEFEEDINDEX) = "Feed"
    grdVehicleSelection.Col = GAMELANGINDEX
    grdVehicleSelection.CellFontBold = False
    grdVehicleSelection.CellFontName = "Arial"
    grdVehicleSelection.CellFontSize = 6.75
    grdVehicleSelection.CellForeColor = vbBlue
    grdVehicleSelection.TextMatrix(0, GAMELANGINDEX) = "Language"

End Sub

Private Sub mVehicleSelectionColumnWidths()
    Dim ilCol As Integer
    Dim llWidth As Long
    Dim ilValue As Integer

    grdVehicleSelection.ColWidth(VEHICLECODEINDEX) = 0
    grdVehicleSelection.ColWidth(GAMENOINDEX) = 0
    grdVehicleSelection.ColWidth(LIBSORTINDEX) = 0
    grdVehicleSelection.ColWidth(FULLTEAMNAMES) = 0
    grdVehicleSelection.ColWidth(LIBSTARTTIMEINDEX) = lmTimeWidth    '0.051 * grdVehicleSelection.Width
    grdVehicleSelection.ColWidth(LIBLENGTHINDEX) = lmLengthWidth    '0.051 * grdVehicleSelection.Width
    grdVehicleSelection.ColWidth(STATIONCODEINDEX) = lmStationCode    '0.051 * grdVehicleSelection.Width
    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then 'Using Sports
        grdVehicleSelection.ColWidth(GAMETEAMSINDEX) = 0.11 * grdVehicleSelection.Width
        'Feed Source
        If (ilValue And USINGFEED) = USINGFEED Then
            grdVehicleSelection.ColWidth(GAMEFEEDINDEX) = 0.07 * grdVehicleSelection.Width
        Else
            grdVehicleSelection.ColWidth(GAMEFEEDINDEX) = 0
        End If
        'Language
        If (ilValue And USINGLANG) = USINGLANG Then
            grdVehicleSelection.ColWidth(GAMELANGINDEX) = 0.09 * grdVehicleSelection.Width
        Else
            grdVehicleSelection.ColWidth(GAMELANGINDEX) = 0
        End If
    Else
        grdVehicleSelection.ColWidth(GAMETEAMSINDEX) = 0
        grdVehicleSelection.ColWidth(GAMEFEEDINDEX) = 0
        grdVehicleSelection.ColWidth(GAMELANGINDEX) = 0
    End If
    llWidth = GRIDSCROLLWIDTH + 45
    For ilCol = 0 To grdVehicleSelection.Cols - 1 Step 1
        If (ilCol <> VEHICLENAMEINDEX) And (ilCol <> LIBEVENTINDEX) Then
            llWidth = llWidth + grdVehicleSelection.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdVehicleSelection.Width - llWidth
    grdVehicleSelection.ColWidth(LIBEVENTINDEX) = llWidth / 3
    llWidth = llWidth - grdVehicleSelection.ColWidth(LIBEVENTINDEX)
    grdVehicleSelection.ColWidth(VEHICLENAMEINDEX) = llWidth
End Sub

Private Sub grdLog_GotFocus()
    If imWhichGrid = 1 Then
        mSetShow True
    End If
End Sub

Private Sub grdLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer
    Dim llSdfCode As Long
    Dim llTstSdfCode As Long

    ilCol = grdLog.MouseCol
    ilRow = grdLog.MouseRow
    If ilCol < grdLog.FixedCols Then
        Exit Sub
    End If
    If ilRow < grdLog.FixedRows Then
        Exit Sub
    End If
    If imSignOn <> 1 Then
        cmcSignOn_Click
        Exit Sub
    End If
    If ilRow = grdLog.FixedRows Then
        If grdLog.TextMatrix(ilRow, TYPEINDEX) = "" Then
            Exit Sub
        End If
    End If
    If mReplaced() Then
        If cmcSignOff.Enabled And cmcSignOff.Visible Then
            cmcSignOff.SetFocus
        End If
        Exit Sub
    End If
    If (imGameStart <> 1) And (imGameStart <> 2) Then
        If (grdLog.TextMatrix(ilRow, TYPEINDEX) <> "Prg Start") And (grdLog.TextMatrix(ilRow, TYPEINDEX) <> "Event Start") Then
            If imSSFType > 0 Then
                MsgBox "Event Start Must Be Specified Prior To Logging Times", vbOKOnly + vbExclamation, "Start"
            Else
                MsgBox "Program Start Must Be Specified Prior To Logging Times", vbOKOnly + vbExclamation, "Start"
            End If
            For ilRow = grdLog.FixedRows To grdLog.Rows - 1 Step 1
                If (grdLog.TextMatrix(ilRow, TYPEINDEX) = "Prg Start") Or (grdLog.TextMatrix(ilRow, TYPEINDEX) = "Event Start") Then
                    grdLog.Row = ilRow
                    grdLog.Col = AIREDINDEX
                    imWhichGrid = 0
                    mEnableBox
                    Exit Sub
                End If
            Next ilRow
            Exit Sub
        End If
    End If
    If (lmEnableCol = AIREDINDEX) And (ilCol = NOTAIREDINDEX) And (ilRow = lmEnableRow) Then
        edcDropdown.Text = ""
    End If
    If (lmEnableCol = NOTAIREDINDEX) And (ilCol = AIREDINDEX) And (ilRow = lmEnableRow) Then
        edcDropdown.Text = ""
    End If
    If ilRow <> lmEnableRow Then
        llSdfCode = grdLog.TextMatrix(ilRow, SDFCODEINDEX)
        mSetShow True
        If ilRow = grdLog.FixedRows Then
            If grdLog.TextMatrix(ilRow, TYPEINDEX) = "" Then
                Exit Sub
            End If
        End If

        If ilRow >= grdLog.Rows Then
            ilRow = grdLog.Rows - 1
        End If
        If llSdfCode <> grdLog.TextMatrix(ilRow, SDFCODEINDEX) Then
            llTstSdfCode = grdLog.TextMatrix(ilRow, SDFCODEINDEX)
            Do While llSdfCode <> llTstSdfCode
                ilRow = ilRow - 1
                If ilRow < grdLog.FixedRows Then
                    Exit Sub
                End If
                llTstSdfCode = grdLog.TextMatrix(ilRow, SDFCODEINDEX)
            Loop
        End If
    Else
        mSetShow False
    End If
    If ilRow >= grdLog.Rows Then
        Exit Sub
    End If
    If grdLog.TextMatrix(ilRow, VEHICLETYPEINDEX) = "S" Then
        If (ilCol <> CERTIYINDEX) Then
            Exit Sub
        End If
    Else
        If (ilCol <> AIREDINDEX) And (ilCol <> NOTAIREDINDEX) And (ilCol <> COPYINDEX) Then
            Exit Sub
        End If
    End If
    If (grdLog.TextMatrix(ilRow, TYPEINDEX) = "Prg Start") Or (grdLog.TextMatrix(ilRow, TYPEINDEX) = "Event Start") Or (grdLog.TextMatrix(ilRow, TYPEINDEX) = "Prg Over") Or (grdLog.TextMatrix(ilRow, TYPEINDEX) = "Event Over") Then
        If ilCol <> AIREDINDEX Then
            Exit Sub
        End If
    End If
    If grdLog.TextMatrix(ilRow, VEHICLETYPEINDEX) = "S" Then
        grdLog.Col = ilCol
        grdLog.Row = ilRow
        lmEnableRow = grdLog.Row
        lmEnableCol = grdLog.Col
        imWhichGrid = 0
        mSetShow True
    Else
        grdLog.Col = ilCol
        grdLog.Row = ilRow
        imWhichGrid = 0
        mEnableBox
    End If
End Sub

Private Sub grdLog_Scroll()
    If (imCtrlVisible) And (grdLog.Row >= grdLog.FixedRows) And (grdLog.Col >= grdLog.FixedCols) And (imWhichGrid = 0) Then
        If grdLog.RowIsVisible(grdLog.Row) Then
            mSetFocus
        Else
            cmcDropdown.Visible = False
            edcDropdown.Visible = False
            lbcNotAired.Visible = False
            lbcCopy.Visible = False
            plcTme.Visible = False
        End If
    End If

End Sub

Private Sub grdLogged_GotFocus()
    If imWhichGrid = 0 Then
        mSetShow True
    End If
End Sub

Private Sub grdLogged_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer
    Dim llSdfCode As Long
    Dim llTstSdfCode As Long
    Dim slType As String
    Dim slTstType As String
    Dim ilRet As Integer

    ilCol = grdLogged.MouseCol
    ilRow = grdLogged.MouseRow
    If ilCol < grdLogged.FixedCols Then
        Exit Sub
    End If
    If ilRow < grdLogged.FixedRows Then
        Exit Sub
    End If
    If imSignOn <> 1 Then
        cmcSignOn_Click
        Exit Sub
    End If
    If mReplaced() Then
        If cmcSignOff.Enabled And cmcSignOff.Visible Then
            cmcSignOff.SetFocus
        End If
        Exit Sub
    End If
    If (grdLogged.TextMatrix(ilRow, TYPEINDEX) = "Sign-On") Or (grdLogged.TextMatrix(ilRow, TYPEINDEX) = "Sign-Off") Then
        Exit Sub
    End If
    If (grdLogged.TextMatrix(ilRow, TYPEINDEX) = "") Then
        Exit Sub
    End If
    If (lmEnableCol = AIREDINDEX) And (ilCol = NOTAIREDINDEX) And (ilRow = lmEnableRow) Then
        edcDropdown.Text = ""
    End If
    If (lmEnableCol = NOTAIREDINDEX) And (ilCol = AIREDINDEX) And (ilRow = lmEnableRow) Then
        edcDropdown.Text = ""
    End If
    If ilRow <> lmEnableRow Then
        llSdfCode = grdLogged.TextMatrix(ilRow, SDFCODEINDEX)
        slType = Trim$(grdLogged.TextMatrix(ilRow, TYPEINDEX))
        mSetShow True
        If ilRow = grdLogged.FixedRows Then
            If slType = "" Then
                Exit Sub
            End If
        End If
        If ilRow >= grdLogged.Rows Then
            ilRow = grdLogged.Rows - 1
        End If
        llTstSdfCode = grdLogged.TextMatrix(ilRow, SDFCODEINDEX)
        slTstType = Trim$(grdLogged.TextMatrix(ilRow, TYPEINDEX))
        If (llSdfCode <> llTstSdfCode) Or (StrComp(slType, slTstType, vbTextCompare) <> 0) Then
            llTstSdfCode = grdLogged.TextMatrix(ilRow, SDFCODEINDEX)
            slTstType = Trim$(grdLogged.TextMatrix(ilRow, TYPEINDEX))
            Do While (llSdfCode <> llTstSdfCode) Or (StrComp(slType, slTstType, vbTextCompare) <> 0)
                ilRow = ilRow - 1
                If ilRow < grdLogged.FixedRows Then
                    Exit Sub
                End If
                llTstSdfCode = grdLogged.TextMatrix(ilRow, SDFCODEINDEX)
                slTstType = Trim$(grdLogged.TextMatrix(ilRow, TYPEINDEX))
            Loop
        End If
    Else
        mSetShow False
    End If
    If ilRow >= grdLogged.Rows Then
        Exit Sub
    End If
    If grdLogged.TextMatrix(grdLogged.FixedRows, TYPEINDEX) = "" Then
        Exit Sub
    End If
    If grdLogged.TextMatrix(ilRow, VEHICLETYPEINDEX) = "S" Then
        If (ilCol <> CERTIYINDEX) Then
            Exit Sub
        End If
    Else
        If (ilCol <> AIREDINDEX) And (ilCol <> NOTAIREDINDEX) And (ilCol <> COPYINDEX) Then
            Exit Sub
        End If
    End If
    If (grdLogged.TextMatrix(ilRow, TYPEINDEX) = "Prg Start") Or (grdLogged.TextMatrix(ilRow, TYPEINDEX) = "Event Start") Or (grdLogged.TextMatrix(ilRow, TYPEINDEX) = "Prg Over") Or (grdLogged.TextMatrix(ilRow, TYPEINDEX) = "Event Over") Then
        If ilCol <> AIREDINDEX Then
            Exit Sub
        End If
        If (grdLogged.TextMatrix(ilRow, TYPEINDEX) = "Prg Start") Then
            ilRet = mFindLdf("P", "S")
        ElseIf (grdLogged.TextMatrix(ilRow, TYPEINDEX) = "Event Start") Then
            ilRet = mFindLdf("G", "S")
        ElseIf (grdLogged.TextMatrix(ilRow, TYPEINDEX) = "Prg Over") Then
            ilRet = mFindLdf("P", "E")
        ElseIf (grdLogged.TextMatrix(ilRow, TYPEINDEX) = "Event Over") Then
            ilRet = mFindLdf("G", "E")
        Else
            ilRet = False
        End If
        If Not ilRet Then
            Exit Sub
        End If
        If tmLdf.iUrfCode <> tgUrf(0).iCode Then
            Exit Sub
        End If
    Else
        llSdfCode = Val(grdLogged.TextMatrix(ilRow, SDFCODEINDEX))
        tmSdfSrchKey3.lCode = llSdfCode
        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            If grdLogged.TextMatrix(ilRow, VEHICLETYPEINDEX) <> "S" Then
                If tmSdf.iUrfCode <> tgUrf(0).iCode Then
                    Exit Sub
                End If
            End If
        Else
            Exit Sub
        End If
    End If
    If grdLogged.TextMatrix(ilRow, VEHICLETYPEINDEX) = "S" Then
        grdLogged.Col = ilCol
        grdLogged.Row = ilRow
        imWhichGrid = 1
        lmEnableRow = grdLogged.Row
        lmEnableCol = grdLogged.Col
        mSetShow True
    Else
        grdLogged.Col = ilCol
        grdLogged.Row = ilRow
        imWhichGrid = 1
        mEnableBox
    End If
End Sub

Private Sub grdLogged_Scroll()
    If (imCtrlVisible) And (grdLogged.Row >= grdLogged.FixedRows) And (grdLogged.Col >= grdLogged.FixedCols) And (imWhichGrid = 1) Then
        If grdLogged.RowIsVisible(grdLogged.Row) Then
            mSetFocus
        Else
            cmcDropdown.Visible = False
            edcDropdown.Visible = False
            lbcNotAired.Visible = False
            lbcCopy.Visible = False
            plcTme.Visible = False
        End If
    End If
End Sub

Private Sub grdVehicleSelection_Click()
    Dim llRow As Long


    If grdVehicleSelection.Row <> grdVehicleSelection.RowSel Then
        llRow = grdVehicleSelection.RowSel
        'grdVehicleSelection.Row = 0
        'grdVehicleSelection.Col = VEHICLECODEINDEX
        grdVehicleSelection.Row = llRow
        grdVehicleSelection.RowSel = llRow
        grdVehicleSelection.Col = VEHICLENAMEINDEX
        grdVehicleSelection.ColSel = STATIONCODEINDEX
    End If
    'If grdVehicleSelection.CellForeColor = DARKGREEN Then
    '    cmcContinue.Enabled = False
    '    grdVehicleSelection.Row = 0
    '    grdVehicleSelection.Col = VEHICLECODEINDEX
    'Else
        If grdVehicleSelection.TextMatrix(grdVehicleSelection.FixedRows, VEHICLENAMEINDEX) <> "No Live Logs" Then
            cmcContinue.Enabled = True
        End If
    'End If
End Sub

Private Sub mLogColumnNames()
    grdLog.Row = 0
    grdLog.Col = AIREDINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.TextMatrix(0, AIREDINDEX) = "Aired"
    grdLog.Col = NOTAIREDINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.TextMatrix(0, NOTAIREDINDEX) = "Not Aired"
    grdLog.Col = CERTIYINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.TextMatrix(0, CERTIYINDEX) = "Certify"
    grdLog.Col = SCHDTIMEINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.TextMatrix(0, SCHDTIMEINDEX) = "Sched Time"
    grdLog.Col = LENGTHINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.TextMatrix(0, LENGTHINDEX) = "Length"
    grdLog.Col = TYPEINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.TextMatrix(0, TYPEINDEX) = "Type"
    grdLog.Col = ADVTPRODINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.TextMatrix(0, ADVTPRODINDEX) = "Advertiser/Product"
    grdLog.Col = COPYINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.TextMatrix(0, COPYINDEX) = "Copy"
End Sub

Private Sub mLogColumnWidths()
    Dim ilCol As Integer
    Dim llWidth As Long

    grdLog.ColWidth(SDFCODEINDEX) = 0
    grdLog.ColWidth(SDFSORTINDEX) = 0
    grdLog.ColWidth(POSITIONINDEX) = 0
    grdLog.ColWidth(VEHICLETYPEINDEX) = 0
    grdLog.ColWidth(ORIGTIMEINDEX) = 0
    grdLog.ColWidth(ORIGCOPYINDEX) = 0
    grdLog.ColWidth(BKQHINDEX) = 0
    grdLog.ColWidth(AIREDINDEX) = lmTimeWidth
    grdLog.ColWidth(NOTAIREDINDEX) = 0.08 * grdLog.Width
    grdLog.ColWidth(SCHDTIMEINDEX) = lmTimeWidth
    grdLog.ColWidth(LENGTHINDEX) = lmLengthWidth
    grdLog.ColWidth(TYPEINDEX) = 0.14 * grdLog.Width
    grdLog.ColWidth(COPYINDEX) = 0.08 * grdLog.Width
    grdLog.ColWidth(CERTIYINDEX) = 0
    If UBound(imAirVefCode) > LBound(imAirVefCode) Then
        grdLog.ColWidth(CERTIYINDEX) = 0.08 * grdLog.Width
    End If

    llWidth = GRIDSCROLLWIDTH + 45
    For ilCol = 0 To grdLog.Cols - 1 Step 1
        If ilCol <> ADVTPRODINDEX Then
            llWidth = llWidth + grdLog.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdLog.Width - llWidth
    If llWidth > 0 Then
        grdLog.ColWidth(ADVTPRODINDEX) = llWidth
    End If

    For ilCol = 0 To grdLog.Cols - 1 Step 1
        grdLogged.ColWidth(ilCol) = grdLog.ColWidth(ilCol)
    Next ilCol


End Sub

Private Sub mPopLog()
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRet As Integer
    Dim ilEvt As Integer
    Dim ilSpot As Integer
    Dim slTime As String
    Dim slRunTime As String
    Dim llRow As Long
    Dim ilCol As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llStartTime(0 To 1) As Long
    Dim llEndTime(0 To 1) As Long
    Dim ilSpotAdded As Integer
    Dim slXMid As String
    Dim ilVef As Integer
    Dim ilLastColSorted As Integer
    Dim ilLastSort As Integer
    Dim slVehicleType As String
    Dim ilType As Integer
    Dim slEventTitle1 As String
    Dim slEventTitle2 As String

    If (grdVehicleSelection.Row <= grdVehicleSelection.FixedRows) And (grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, VEHICLENAMEINDEX) = "No Live Logs") Then
        Exit Sub
    End If
    'If grdVehicleSelection.CellForeColor <> vbBlack Then
    '    Exit Sub
    'End If
    ilSpotAdded = False
    grdLog.Redraw = False
    grdLog.Rows = grdLog.FixedRows + 1
    llRow = grdLog.FixedRows
    For ilCol = 0 To grdLog.Cols - 1 Step 1
        grdLog.TextMatrix(llRow, ilCol) = ""
    Next ilCol
    grdLog.Row = grdLog.FixedRows
    grdLog.Col = AIREDINDEX
    grdLog.CellBackColor = vbGreen
    grdLog.Col = NOTAIREDINDEX
    grdLog.CellBackColor = vbRed
    grdLog.Col = SCHDTIMEINDEX
    grdLog.CellBackColor = LIGHTYELLOW
    grdLog.Col = LENGTHINDEX
    grdLog.CellBackColor = LIGHTYELLOW
    grdLog.Col = TYPEINDEX
    grdLog.CellBackColor = LIGHTYELLOW
    grdLog.Col = ADVTPRODINDEX
    grdLog.CellBackColor = LIGHTYELLOW
    For ilCol = 0 To grdLog.Cols - 1 Step 1
        grdLog.Col = ilCol
        grdLog.CellAlignment = flexAlignLeftTop
    Next ilCol

    grdLogged.Redraw = False
    grdLogged.Rows = grdLogged.FixedRows + 1
    llRow = grdLogged.FixedRows
    For ilCol = 0 To grdLogged.Cols - 1 Step 1
        grdLogged.TextMatrix(llRow, ilCol) = ""
    Next ilCol
    grdLogged.Row = llRow
    grdLogged.Col = SCHDTIMEINDEX
    grdLogged.CellBackColor = LIGHTYELLOW
    grdLogged.Col = LENGTHINDEX
    grdLogged.CellBackColor = LIGHTYELLOW
    grdLogged.Col = TYPEINDEX
    grdLogged.CellBackColor = LIGHTYELLOW
    grdLogged.Col = ADVTPRODINDEX
    grdLogged.CellAlignment = flexAlignLeftTop
    grdLogged.CellBackColor = LIGHTYELLOW
    For ilCol = 0 To grdLog.Cols - 1 Step 1
        grdLogged.Col = ilCol
        grdLogged.CellAlignment = flexAlignLeftTop
    Next ilCol

    llRow = grdLog.FixedRows
    lacBanner(0).Caption = " " & Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, VEHICLENAMEINDEX)) & " " & Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, FULLTEAMNAMES))
    gAddTimeLength Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)), Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBLENGTHINDEX)), "A", "1", slEndTime, slXMid
    'lacBanner(1).Caption = " " & Format$(lmNowDate, "m/d/yy") & " " & Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)) & "-" & slEndTime
    lacBanner(1).Caption = " " & Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)) & "-" & slEndTime
    imVefCode = Val(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, VEHICLECODEINDEX))
    gGetEventTitles imVefCode, slEventTitle1, slEventTitle2
    imGameNo = Val(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMENOINDEX))
    If imGameNo > 0 Then
        If (Asc(tgSpf.sSportInfo) And USINGFEED) = USINGFEED Then
            If grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMEFEEDINDEX) = slEventTitle1 Then  '"Visiting" Then
                lacBanner(1).Caption = lacBanner(1).Caption & " Feed: " & slEventTitle1 'Visiting"
            ElseIf grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMEFEEDINDEX) = slEventTitle2 Then  '"Home" Then
                lacBanner(1).Caption = lacBanner(1).Caption & " Feed: " & slEventTitle2   '"Home"
            ElseIf grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMEFEEDINDEX) = "National" Then
                lacBanner(1).Caption = lacBanner(1).Caption & " Feed: National"
            End If
        End If
        If ((Asc(tgSpf.sSportInfo) And USINGLANG) = USINGLANG) And (Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMELANGINDEX)) <> "") Then
            lacBanner(1).Caption = lacBanner(1).Caption & " Language: " & Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMELANGINDEX))
        End If
    End If
    If grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMEFEEDINDEX) = slEventTitle1 Then  '"Visiting" Then
        smCurrentFeed = "V"
        imCurrentVefCode = imVefCode
    ElseIf grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMEFEEDINDEX) = slEventTitle2 Then  '"Home" Then
        smCurrentFeed = "H"
        imCurrentVefCode = imVefCode
    ElseIf grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMEFEEDINDEX) = "National" Then
        smCurrentFeed = "N"
        imCurrentVefCode = imVefCode
    Else
        smCurrentFeed = ""
        imCurrentVefCode = 0
    End If
    gPackDateLong lmEventDate, ilDate0, ilDate1   'lmNowDate, ilDate0, ilDate1
    slStartTime = grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)
    llStartTime(0) = gTimeToLong(slStartTime, False)
    llEndTime(0) = gTimeToLong(slEndTime, True) - 1
    lmLibStartTime(0) = llStartTime(0)
    lmLibEndTime(0) = llEndTime(0)
    If llEndTime(0) < llStartTime(0) Then
        llStartTime(1) = 0
        llEndTime(1) = llEndTime(0)
        llEndTime(0) = 86399
        lmLibStartTime(1) = 0
        lmLibEndTime(1) = lmLibEndTime(0)
        lmLibEndTime(0) = 86399
    Else
        llStartTime(1) = -1
        llEndTime(1) = -1
        lmLibStartTime(1) = -1
        lmLibEndTime(1) = -1
    End If
    mBuildVehicleArray
    
    ReDim igEventAdfCode(0 To 0) As Integer
    
    imSSFType = Val(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMENOINDEX))
    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
    ilVef = LBound(imLLVefCode)
    If imGameNo <= 0 Then
        If smVefType <> "L" Then
            tmSsfSrchKey.iType = 0 'slType-On Air
            tmSsfSrchKey.iVefCode = imLLVefCode(ilVef)  'imVefCode
            tmSsfSrchKey.iDate(0) = ilDate0
            tmSsfSrchKey.iDate(1) = ilDate1
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        Else
            Do
                tmSsfSrchKey2.iVefCode = imLLVefCode(ilVef) 'imVefCode
                tmSsfSrchKey2.iDate(0) = ilDate0
                tmSsfSrchKey2.iDate(1) = ilDate1
                ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    imGameNo = tmSsf.iType
                    Exit Do
                End If
                ilVef = ilVef + 1
                If ilVef >= UBound(imLLVefCode) Then
                    Exit Do
                End If
            Loop
        End If
    Else
        'tmSsfSrchKey1.iVefCode = imVefCode
        'tmSsfSrchKey1.iType = imGameNo
        'ilRet = gSSFGetEqualKey1(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
        tmSsfSrchKey2.iVefCode = imLLVefCode(ilVef) 'imVefCode
        tmSsfSrchKey2.iDate(0) = ilDate0
        tmSsfSrchKey2.iDate(1) = ilDate1
        ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = imLLVefCode(ilVef))
            If (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1) Then
                If tmSsf.iType = imGameNo Then
                    Exit Do
                End If
            Else
                Exit Do
            End If
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    End If
    If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = imGameNo) And (tmSsf.iVefCode = imLLVefCode(ilVef)) And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1) Then
        If (imSignOn <> 2) Or (smGameStartTime = "") Then
            If llRow >= grdLog.Rows Then
                grdLog.AddItem ""
            End If
            grdLog.Row = llRow
            grdLog.Col = AIREDINDEX
            grdLog.CellBackColor = vbGreen
            grdLog.Col = NOTAIREDINDEX
            grdLog.CellBackColor = vbRed
            grdLog.Col = SCHDTIMEINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = LENGTHINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = TYPEINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = ADVTPRODINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = COPYINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = CERTIYINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            For ilCol = 0 To grdLog.Cols - 1 Step 1
                grdLog.Col = ilCol
                grdLog.CellAlignment = flexAlignLeftTop
            Next ilCol
            grdLog.TextMatrix(llRow, SCHDTIMEINDEX) = grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)
            smGameStartTime = grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)
            grdLog.TextMatrix(llRow, LENGTHINDEX) = grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBLENGTHINDEX)
            If imSSFType > 0 Then
                grdLog.TextMatrix(llRow, TYPEINDEX) = "Event Start"
            Else
                grdLog.TextMatrix(llRow, TYPEINDEX) = "Prg Start"
            End If
            grdLog.TextMatrix(llRow, ADVTPRODINDEX) = Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, VEHICLENAMEINDEX))
            grdLog.TextMatrix(llRow, SDFCODEINDEX) = 0
            grdLog.TextMatrix(llRow, POSITIONINDEX) = 0
            grdLog.TextMatrix(llRow, VEHICLETYPEINDEX) = ""
            grdLog.TextMatrix(llRow, SDFSORTINDEX) = "000000"
            llRow = llRow + 1
        Else
            If grdLogged.FixedRows >= grdLogged.Rows Then
                grdLogged.AddItem ""
            End If
            grdLogged.Row = grdLogged.FixedRows
            grdLogged.Col = SCHDTIMEINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = LENGTHINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = TYPEINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = ADVTPRODINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = COPYINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = CERTIYINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            For ilCol = 0 To grdLogged.Cols - 1 Step 1
                grdLogged.Col = ilCol
                grdLogged.CellAlignment = flexAlignLeftTop
            Next ilCol
            grdLogged.TextMatrix(grdLogged.FixedRows, AIREDINDEX) = smGameStartTime
            grdLogged.TextMatrix(grdLogged.FixedRows, SCHDTIMEINDEX) = grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)
            grdLogged.TextMatrix(grdLogged.FixedRows, LENGTHINDEX) = grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBLENGTHINDEX)
            If imSSFType > 0 Then
                grdLogged.TextMatrix(grdLogged.FixedRows, TYPEINDEX) = "Event Start"
            Else
                grdLogged.TextMatrix(grdLogged.FixedRows, TYPEINDEX) = "Prg Start"
            End If
            grdLogged.TextMatrix(grdLogged.FixedRows, ADVTPRODINDEX) = Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, VEHICLENAMEINDEX))
            grdLogged.TextMatrix(grdLogged.FixedRows, SDFCODEINDEX) = 0
            grdLogged.TextMatrix(grdLogged.FixedRows, POSITIONINDEX) = 0
            grdLogged.TextMatrix(grdLogged.FixedRows, VEHICLETYPEINDEX) = ""
            grdLogged.TextMatrix(grdLogged.FixedRows, SDFSORTINDEX) = "000000"
            imGameStart = 1
        End If
        
        Do While (ilRet = BTRV_ERR_NONE) And ((tmSsf.iType = imGameNo) Or (smVefType = "L")) And (tmSsf.iVefCode = imLLVefCode(ilVef)) And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1)
            ilType = gBinarySearchVef(imLLVefCode(ilVef))
            If ilType <> -1 Then
                slVehicleType = tgMVef(ilType).sType
            Else
                slVehicleType = ""
            End If
            If slVehicleType = "S" Then
                'If Not mObtainVlfStartDate(imLLVefCode(ilVef), lmNowDate) Then
                If Not mObtainVlfStartDate(imLLVefCode(ilVef), lmEventDate) Then
                    tmSsf.iCount = 0
                End If
            End If
            For ilEvt = 1 To tmSsf.iCount Step 1
               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                    slRunTime = slTime
                    If ((gTimeToLong(slTime, False) >= llStartTime(0)) And (gTimeToLong(slTime, False) <= llEndTime(0))) Or ((gTimeToLong(slTime, False) >= llStartTime(1)) And (gTimeToLong(slTime, False) <= llEndTime(1))) Then
                        For ilSpot = ilEvt + 1 To ilEvt + tmAvail.iNoSpotsThis Step 1
                           LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpot)
                            tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            If ilRet = BTRV_ERR_NONE Then
                                If slVehicleType <> "S" Then
                                    If (Trim$(tmSdf.sAffChg) = "") Or (Trim$(tmSdf.sAffChg) = "N") Then
                                        ilSpotAdded = True
                                        mAddSpotToGrid grdLog, 0, slTime, slRunTime
                                    Else
                                        mAddSpotToGrid grdLogged, 1, slTime, slRunTime
                                    End If
                                    slRunTime = gFormatTimeLong(gTimeToLong(slRunTime, False) + tmSdf.iLen, "A", "1")
                                Else
                                    'ilRet = mTranslateSellToAir(tmSdf.iVefCode, lmNowDate, ilSpotAdded)
                                    ilRet = mTranslateSellToAir(tmSdf.iVefCode, lmEventDate, ilSpotAdded)
                                End If
                            End If
                        Next ilSpot
                    End If
                End If
            Next ilEvt
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            If (ilRet = BTRV_ERR_NONE) And ((tmSsf.iType = imGameNo) Or (smVefType = "L")) And (tmSsf.iVefCode = imLLVefCode(ilVef)) And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1) Then
            Else
                If smVefType = "L" Then
                    ilVef = ilVef + 1
                    Do
                        If ilVef < UBound(imLLVefCode) Then
                            tmSsfSrchKey2.iVefCode = imLLVefCode(ilVef) 'imVefCode
                            tmSsfSrchKey2.iDate(0) = ilDate0
                            tmSsfSrchKey2.iDate(1) = ilDate1
                            ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                            If ilRet = BTRV_ERR_NONE Then
                                Exit Do
                            End If
                            ilVef = ilVef + 1
                        Else
                            Exit Do
                        End If
                    Loop
                Else
                    Exit Do
                End If
            End If
        Loop
        If smVefType = "L" Then
            'Sort Time within Date
            ilLastColSorted = -1
            ilLastSort = -1
            gGrid_SortByCol grdLog, ADVTPRODINDEX, SDFSORTINDEX, ilLastColSorted, ilLastSort
        End If
        If (imSignOn <> 2) Or (smGameEndTime = "") Then
            If grdLog.TextMatrix(grdLog.Rows - 1, TYPEINDEX) <> "" Then
                grdLog.AddItem ""
            End If
            grdLog.Row = grdLog.Rows - 1
            llRow = grdLog.Row
            For ilCol = 0 To grdLog.Cols - 1 Step 1
                grdLog.Col = ilCol
                grdLog.CellAlignment = flexAlignLeftTop
            Next ilCol
            grdLog.Col = AIREDINDEX
            grdLog.CellBackColor = vbGreen
            grdLog.Col = NOTAIREDINDEX
            grdLog.CellBackColor = vbRed
            grdLog.Col = SCHDTIMEINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = LENGTHINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = TYPEINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = ADVTPRODINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = COPYINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = CERTIYINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            gAddTimeLength Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBSTARTTIMEINDEX)), Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, LIBLENGTHINDEX)), "A", "1", slEndTime, slXMid
            grdLog.TextMatrix(llRow, SCHDTIMEINDEX) = slEndTime
            smGameEndTime = slEndTime
            grdLog.TextMatrix(llRow, LENGTHINDEX) = ""
            If imSSFType > 0 Then
                grdLog.TextMatrix(llRow, TYPEINDEX) = "Event Over"
            Else
                grdLog.TextMatrix(llRow, TYPEINDEX) = "Prg Over"
            End If
            grdLog.TextMatrix(llRow, ADVTPRODINDEX) = Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, VEHICLENAMEINDEX))
            grdLog.TextMatrix(llRow, SDFCODEINDEX) = 0
        Else
            If grdLogged.TextMatrix(grdLogged.Rows - 1, TYPEINDEX) <> "" Then
                grdLogged.AddItem ""
            End If
            grdLogged.Row = grdLogged.Rows - 1
            llRow = grdLogged.Row
            For ilCol = 0 To grdLogged.Cols - 1 Step 1
                grdLogged.Col = ilCol
                grdLogged.CellAlignment = flexAlignLeftTop
            Next ilCol
            grdLogged.Col = SCHDTIMEINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = LENGTHINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = TYPEINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = ADVTPRODINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = COPYINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = CERTIYINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.TextMatrix(llRow, AIREDINDEX) = smGameEndTime
            grdLogged.TextMatrix(llRow, SCHDTIMEINDEX) = ""
            grdLogged.TextMatrix(llRow, LENGTHINDEX) = ""
            If imSSFType > 0 Then
                grdLogged.TextMatrix(llRow, TYPEINDEX) = "Event Over"
            Else
                grdLogged.TextMatrix(llRow, TYPEINDEX) = "Prg Over"
            End If
            grdLogged.TextMatrix(llRow, ADVTPRODINDEX) = Trim$(grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, VEHICLENAMEINDEX))
            grdLogged.TextMatrix(llRow, SDFCODEINDEX) = 0
        End If
    End If
'    If Not ilSpotAdded Then
'        If grdLog.TextMatrix(grdLog.Rows - 1, TYPEINDEX) <> "" Then
'            grdLog.AddItem ""
'        End If
'        grdLog.Row = grdLog.Rows - 1
'        grdLog.Col = SCHDTIMEINDEX
'        grdLog.CellBackColor = LIGHTYELLOW
'        grdLog.Col = LENGTHINDEX
'        grdLog.CellBackColor = LIGHTYELLOW
'        grdLog.Col = TYPEINDEX
'        grdLog.CellBackColor = LIGHTYELLOW
'        grdLog.Col = ADVTPRODINDEX
'        grdLog.CellBackColor = LIGHTYELLOW
'        For ilCol = 0 To grdLog.Cols - 1 Step 1
'            grdLog.Col = ilCol
'            grdLog.CellAlignment = flexAlignLeftTop
'        Next ilCol
'        grdLog.TextMatrix(grdLog.FixedRows, ADVTPRODINDEX) = "No Spots"
'    End If
    grdLogged.Redraw = True
    grdLog.Redraw = True
    If smVefType <> "L" Then
        If (imGameStart = 1) Or (imGameStart = 2) Then
            cmcAddSpot.Enabled = True
        Else
            cmcAddSpot.Enabled = False
        End If
    Else
        cmcAddSpot.Enabled = False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCopy                     *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Sub mObtainCopy(slCart As String, slISCI As String, slProduct As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************

'
'   mObtainCopy
'       Where:
'           tmSdf(I)- Spot record
'           slCart(O)- Carts (different zones separated by Chr(10))
'           slISCI(O)- ISCI (different zones separated by Chr(10))
'
    Dim ilRet As Integer
    slCart = ""
    slISCI = ""
    slProduct = ""
    tmCif.iMcfCode = 0
    If tmSdf.sPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        tmCifSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If tmCif.lcpfCode > 0 Then
                tmCpfSrchKey.lCode = tmCif.lcpfCode
                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    tmCpf.sISCI = ""
                    tmCpf.sName = ""
                End If
                slISCI = Trim$(tmCpf.sISCI)
                slProduct = Trim$(tmCpf.sName)
            End If
            If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                If tmCif.iMcfCode <> tmMcf.iCode Then
                    tmMcfSrchKey.iCode = tmCif.iMcfCode
                    ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        tmMcf.sName = ""
                    End If
                    slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                Else
                    slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                End If
            Else
                slCart = " "
            End If
        End If
    ElseIf tmSdf.sPtType = "2" Then  '  Combo Copy
    ElseIf tmSdf.sPtType = "3" Then  '  Time Zone Copy
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If imWhichGrid = 0 Then
        If mGridEnable(grdLog) Then
            mGridSetFocus grdLog
        End If
    ElseIf imWhichGrid = 1 Then
        If mGridEnable(grdLogged) Then
            mGridSetFocus grdLogged
        End If
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilRet As Integer

    If imWhichGrid = 0 Then
        ilRet = mGridSetFocus(grdLog)
    ElseIf imWhichGrid = 1 Then
        ilRet = mGridSetFocus(grdLogged)
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilMoveRow As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim ilRet As Integer

    If imWhichGrid = 0 Then
        ilRet = mGridSetShow(grdLog, ilMoveRow)
    ElseIf imWhichGrid = 1 Then
        ilRet = mGridSetShow(grdLogged, ilMoveRow)
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyPop                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection Name    *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mCopyPop()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mCopyPopErr                                                                           *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilAdvtCode As Integer
    Dim ilISCIProd As Integer
    Dim llSdfCode As Long
    Dim ilLength As Integer
    Dim slLive As String

    If tgSpf.sUseCartNo <> "N" Then
        ilISCIProd = 2
    Else
        ilISCIProd = 5
    End If
    ' Read SDF to get advertiser
    If imWhichGrid = 0 Then
        llSdfCode = Val(grdLog.TextMatrix(lmEnableRow, SDFCODEINDEX))
    Else
        llSdfCode = Val(grdLogged.TextMatrix(lmEnableRow, SDFCODEINDEX))
    End If
    tmSdfSrchKey3.lCode = llSdfCode
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    ilAdvtCode = tmSdf.iAdfCode
    ilLength = tmSdf.iLen
    slLive = ""
    'ilRet = gPopCopyForAdvtBox(PostLog, ilAdvtCode, ilISCIProd, 1, lbcCopyNm, lbcCopyNmCode)
    If (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
        ilRet = gPopCopyForAdvtBoxWithTest(LiveLog, ilAdvtCode, -1, ilLength, slLive, ilISCIProd, 0 + &H200, lbcCopy, tmCopyCode(), smCopyCodeTag, 0, 0)
    Else
        ilRet = gPopCopyForAdvtBoxWithTest(LiveLog, ilAdvtCode, -1, ilLength, slLive, ilISCIProd, 0, lbcCopy, tmCopyCode(), smCopyCodeTag, 0, 0)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        'On Error GoTo mCopyPopErr
        'gCPErrorMsg ilRet, "mCopyPop (gPopCopyForAdvtBox: LiveLog)", LiveLog
        'On Error GoTo 0
        lbcCopy.AddItem "[None]", 0   ' Add this at the top of the list
    End If
    Exit Sub
mCopyPopErr: 'VBC NR
    On Error GoTo 0
    'imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMissedPop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Missed list box if    *
'*                      required                       *
'*                                                     *
'*******************************************************
Private Sub mMissedPop()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mMissedPopErr                                                                         *
'******************************************************************************************

'
'   mMissedPop
'   Where:
'
    ReDim ilfilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim slDefaultName As String
    Dim slStr As String
    
    ilIndex = lbcNotAired.ListIndex
    If ilIndex >= 0 Then
        slDefaultName = lbcNotAired.List(ilIndex)
    Else
        slDefaultName = "Missed"
        slStr = ""
        'ReDim tgMRMnf(1 To 1) As MNF
        ReDim tgMRMnf(0 To 0) As MNF
        ilRet = gObtainMnfForType("M", slStr, tgMRMnf())
        For ilLoop = LBound(tgMRMnf) To UBound(tgMRMnf) - 1 Step 1
            If Trim$(tgMRMnf(ilLoop).sUnitType) = "Y" Then
                slDefaultName = Trim$(tgMRMnf(ilLoop).sName)
                Exit For
            End If
        Next ilLoop
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "M"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilfilter(1) = CHARFILTERNOT
    slFilter(1) = "A"
    ilOffSet(1) = gFieldOffset("Mnf", "MnfCodeStn")
    'ilRet = gIMoveListBox(PostLog, lbcNotAired, lbcNotAiredCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(LiveLog, lbcNotAired, tmMissedCode(), smMissedCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        'On Error GoTo mMissedPopErr
        'gCPErrorMsg ilRet, "mMissedPop (gIMoveListBox)", LiveLog
        'On Error GoTo 0
        imChgMode = True
        gFindMatch slDefaultName, 0, lbcNotAired
        If gLastFound(lbcNotAired) >= 0 Then
            lbcNotAired.ListIndex = gLastFound(lbcNotAired)
        Else
            lbcNotAired.ListIndex = -1
        End If
        imChgMode = False
    End If
    Exit Sub
mMissedPopErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
End Sub

Private Sub mAddMissedReason(slMissedReason As String)
    Dim ilRet As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String

    tmMnf.iCode = 0
    tmMnf.sType = "M"
    tmMnf.sName = slMissedReason
    tmMnf.sRPU = ""
    tmMnf.sUnitType = ""
    tmMnf.sSSComm = ""
    tmMnf.iMerge = 0
    tmMnf.iGroupNo = 0
    tmMnf.sCodeStn = ""
    tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
    tmMnf.iAutoCode = tmMnf.iCode
    ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
    Do
        tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
        tmMnf.iAutoCode = tmMnf.iCode
        gGetSyncDateTime slSyncDate, slSyncTime
        gPackDate slSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
        gPackTime slSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
        ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    smMissedCodeTag = ""
    mMissedPop

End Sub

Private Sub mMoveRow(llRowNo As Long)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilRet As Integer
    Dim slTime As String
    Dim ilMoveEvent As Integer
    Dim slMissedReason As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilBkQH As Integer
    Dim ilPriceLevel As Integer
    Dim slSchStatus As String
    Dim llSdfRecPos As Long
    Dim ilSdf As Integer
    Dim llAvailTime As Long
    Dim slStr As String
    Dim ilAvailIndex As Integer
    Dim ilFound As Integer
    Dim ilGameNo As Integer
    Dim slNowDate As String
    Dim llRowTest As Long
    Dim ilLastMatchingSpot As Integer
    Dim tlSdf As SDF

    ilMoveEvent = True
    If imWhichGrid = 0 Then
        If (llRowNo < grdLog.FixedRows) Or (llRowNo >= grdLog.Rows) Then
            Exit Sub
        End If
        grdLog.Redraw = False
        grdLogged.Redraw = False
        If grdLogged.TextMatrix(grdLogged.Rows - 1, TYPEINDEX) <> "" Then
            grdLogged.AddItem ""
        End If
        grdLogged.Row = grdLogged.Rows - 1
        llRow = grdLogged.Row
        grdLogged.Col = SCHDTIMEINDEX
        grdLogged.CellBackColor = LIGHTYELLOW
        grdLogged.Col = LENGTHINDEX
        grdLogged.CellBackColor = LIGHTYELLOW
        grdLogged.Col = TYPEINDEX
        grdLogged.CellBackColor = LIGHTYELLOW
        grdLogged.Col = ADVTPRODINDEX
        grdLogged.CellBackColor = LIGHTYELLOW
        For llCol = 0 To grdLogged.Cols - 1 Step 1
            grdLogged.Col = llCol
            grdLogged.CellAlignment = flexAlignLeftTop
        Next llCol
        For llCol = 0 To grdLogged.Cols - 1 Step 1
            grdLogged.Col = llCol
            If grdLog.TextMatrix(llRowNo, NOTAIREDINDEX) = "" Then
                grdLogged.CellForeColor = DARKGREEN 'vbGreen
            Else
                grdLogged.CellForeColor = vbRed
            End If
            grdLogged.TextMatrix(llRow, llCol) = grdLog.TextMatrix(llRowNo, llCol)
        Next llCol
        If (grdLog.TextMatrix(llRowNo, TYPEINDEX) = "Prg Start") Or (grdLog.TextMatrix(llRowNo, TYPEINDEX) = "Event Start") Then
            grdLogged.Col = COPYINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = CERTIYINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            'Do
            '    tmLlfSrchKey.lCode = tmLlf.lCode
            '    ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
            '    If ilRet = BTRV_ERR_NONE Then
            '        slTime = grdLog.TextMatrix(llRowNo, AIREDINDEX)
            '        gPackTime slTime, tmLlf.iStartTime(0), tmLlf.iStartTime(1)
            '        ilRet = btrUpdate(hmLlf, tmLlf, imLlfRecLen)
            '    Else
            '        Exit Do
            '    End If
            'Loop While ilRet = BTRV_ERR_CONFLICT
            tmLdf.lCode = 0
            If grdLog.TextMatrix(llRowNo, TYPEINDEX) = "Prg Start" Then
                tmLdf.sType = "P"
            Else
                tmLdf.sType = "G"
            End If
            tmLdf.sSubType = "S"
            slTime = grdLog.TextMatrix(llRowNo, AIREDINDEX)
            gPackTime slTime, tmLdf.iTime(0), tmLdf.iTime(1)
            tmLdf.lLlfCode = tmLlf.lCode
            tmLdf.iUrfCode = tgUrf(0).iCode
            tmLdf.sCurrentFeed = smCurrentFeed
            slNowDate = gNow()
            smActualStartDate = smEventDate 'slNowDate
            tmLdf.lEnteredDate = gDateValue(Format$(slNowDate, "m/d/yy"))
            tmLdf.lEnteredTime = gTimeToLong(Format$(slNowDate, "h:mm:ssAM/PM"), False)
            tmLdf.sUnused = ""
            ilRet = btrInsert(hmLdf, tmLdf, imLdfRecLen, INDEXKEY0)
            imGameStart = 1
            imPrgStartUrfCode = tgUrf(0).iCode
        ElseIf (grdLog.TextMatrix(llRowNo, TYPEINDEX) = "Prg Over") Or (grdLog.TextMatrix(llRowNo, TYPEINDEX) = "Event Over") Then
            grdLogged.Col = COPYINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = CERTIYINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            ''Do
            ''    tmLlfSrchKey.lCode = tmLlf.lCode
            ''    ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
            ''    If ilRet = BTRV_ERR_NONE Then
            ''        slTime = grdLog.TextMatrix(llRowNo, AIREDINDEX)
            ''        gPackTime slTime, tmLlf.iEndTime(0), tmLlf.iEndTime(1)
            ''        ilRet = btrUpdate(hmLlf, tmLlf, imLlfRecLen)
            ''    Else
            ''        Exit Do
            ''    End If
            ''Loop While ilRet = BTRV_ERR_CONFLICT
            'Do
            '    tmLDFSrchKey1.lLlfCode = tmLlf.lCode
            '    ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            '    Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlf.lCode)
            '        If (tmLdf.sType = "P") Or (tmLdf.sType = "G") Then
            '            slTime = grdLog.TextMatrix(llRowNo, AIREDINDEX)
            '            gPackTime slTime, tmLdf.iEndTime(0), tmLdf.iEndTime(1)
            '            ilRet = btrUpdate(hmLdf, tmLdf, imLdfRecLen)
            '            Exit Do
            '        End If
            '        ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
            '    Loop
            'Loop While ilRet = BTRV_ERR_CONFLICT
            tmLdf.lCode = 0
            If grdLog.TextMatrix(llRowNo, TYPEINDEX) = "Prg Over" Then
                tmLdf.sType = "P"
            Else
                tmLdf.sType = "G"
            End If
            tmLdf.sSubType = "E"
            slTime = grdLog.TextMatrix(llRowNo, AIREDINDEX)
            gPackTime slTime, tmLdf.iTime(0), tmLdf.iTime(1)
            tmLdf.lLlfCode = tmLlf.lCode
            tmLdf.iUrfCode = tgUrf(0).iCode
            tmLdf.sCurrentFeed = smCurrentFeed
            slNowDate = gNow()
            tmLdf.lEnteredDate = gDateValue(Format$(slNowDate, "m/d/yy"))
            tmLdf.lEnteredTime = gTimeToLong(Format$(slNowDate, "h:mm:ssAM/PM"), False)
            tmLdf.sUnused = ""
            ilRet = btrInsert(hmLdf, tmLdf, imLdfRecLen, INDEXKEY0)
            imGameStart = 2
            imPrgEndUrfCode = tgUrf(0).iCode
        Else
            If grdLog.TextMatrix(llRowNo, VEHICLETYPEINDEX) = "S" Then
                grdLogged.Col = AIREDINDEX
                grdLogged.CellBackColor = LIGHTYELLOW
                grdLogged.Col = NOTAIREDINDEX
                grdLogged.CellBackColor = LIGHTYELLOW
                grdLogged.Col = COPYINDEX
                grdLogged.CellBackColor = LIGHTYELLOW
                grdLogged.Col = CERTIYINDEX
                grdLogged.CellBackColor = CYAN 'DARKGREEN    'vbGreen
                For llCol = 0 To grdLogged.Cols - 1 Step 1
                    grdLogged.Col = llCol
                    grdLogged.CellForeColor = MIDDLEBLUE
                Next llCol
            Else
                grdLogged.Col = CERTIYINDEX
                grdLogged.CellBackColor = LIGHTYELLOW
            End If
            Do
                tmSdfSrchKey3.lCode = Val(grdLog.TextMatrix(llRowNo, SDFCODEINDEX))
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                If ilRet = BTRV_ERR_NONE Then
                    If grdLog.TextMatrix(llRowNo, VEHICLETYPEINDEX) <> "S" Then
                        If grdLog.TextMatrix(llRowNo, NOTAIREDINDEX) = "" Then
                            'Change time of spot
                            If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
                                If grdLog.TextMatrix(llRowNo, VEHICLETYPEINDEX) <> "S" Then
                                    tmSdf.sXCrossMidnight = mCrossMidnight(grdLog.TextMatrix(llRowNo, ORIGTIMEINDEX), grdLog.TextMatrix(llRowNo, AIREDINDEX))
                                    gPackTime grdLog.TextMatrix(llRowNo, AIREDINDEX), tmSdf.iTime(0), tmSdf.iTime(1)
                                End If
                                tmSdf.sAffChg = "Y"
                                tmSdf.iUrfCode = tgUrf(0).iCode
                                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                            End If
                        Else
                            'Missed Spot
                            tmChfSrchKey.lCode = tmSdf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                'Force reading of SSF
                                lmSsfDate = 0
                                ilGameNo = tmSdf.iGameNo
                                If ((tmChf.sType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmChf.sType = "Q") Or ((tmChf.sType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmChf.sType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (tmSdf.sSpotType = "X") Then
                                    tmDelOrMissedSdf(UBound(tmDelOrMissedSdf)) = tmSdf
                                    ReDim Preserve tmDelOrMissedSdf(0 To UBound(tmDelOrMissedSdf) + 1) As SDF
                                    ilRet = gChgSchSpot("D", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf, lmSsfDate, lmSsfRecPos, hmSxf, hmGsf, hmGhf)
                                    If Not ilRet Then
                                        ilMoveEvent = False
                                    End If
                                ElseIf (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
                                    tmDelOrMissedSdf(UBound(tmDelOrMissedSdf)) = tmSdf
                                    ReDim Preserve tmDelOrMissedSdf(0 To UBound(tmDelOrMissedSdf) + 1) As SDF
                                    ilRet = btrDelete(hmSdf)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        ilMoveEvent = False
                                    End If
                                Else
                                    tmDelOrMissedSdf(UBound(tmDelOrMissedSdf)) = tmSdf
                                    ReDim Preserve tmDelOrMissedSdf(0 To UBound(tmDelOrMissedSdf) + 1) As SDF
                                    ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf, lmSsfDate, lmSsfRecPos, hmSxf, hmGsf, hmGhf)
                                    If Not ilRet Then
                                        ilMoveEvent = False
                                    Else
                                        'Set Missed reason
                                        tmSdf.iMnfMissed = igDefaultMnfMissed
                                        slMissedReason = grdLog.TextMatrix(llRowNo, NOTAIREDINDEX)
                                        gFindMatch slMissedReason, 0, lbcNotAired
                                        If gLastFound(lbcNotAired) >= 0 Then
                                            slNameCode = tmMissedCode(gLastFound(lbcNotAired)).sKey
                                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                            tmSdf.iMnfMissed = Val(slCode)
                                        End If
                                        tmSdf.sAffChg = ""
                                        tmSdf.iUrfCode = tgUrf(0).iCode
                                        ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            ilMoveEvent = False
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        ilLastMatchingSpot = True
                        For llRowTest = grdLog.FixedRows To grdLog.Rows - 1 Step 1
                            If llRowTest <> llRowNo Then
                                If grdLog.TextMatrix(llRowNo, SDFCODEINDEX) = grdLog.TextMatrix(llRowTest, SDFCODEINDEX) Then
                                    ilLastMatchingSpot = False
                                    Exit For
                                End If
                            End If
                        Next llRowTest
                        If ilLastMatchingSpot Then
                            tmSdf.sAffChg = "Y"
                            tmSdf.iUrfCode = tgUrf(0).iCode
                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                        End If
                    End If
                Else
                    Exit Do
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
        If ilMoveEvent Then
            If grdLog.FixedRows = grdLog.Rows - 1 Then
                For llCol = 0 To grdLog.Cols - 1 Step 1
                    grdLog.TextMatrix(grdLog.FixedRows, llCol) = ""
                Next llCol
            Else
                grdLog.RemoveItem llRowNo
            End If
        Else
            grdLogged.RemoveItem grdLogged.Rows - 1
        End If
        grdLog.Redraw = True
        grdLogged.Redraw = True
        Do While Not grdLogged.RowIsVisible(grdLogged.Rows - 1)
            grdLogged.TopRow = grdLogged.TopRow + 1
        Loop
    Else
        If (llRowNo < grdLogged.FixedRows) Or (llRowNo >= grdLogged.Rows) Then
            Exit Sub
        End If
        ilMoveEvent = False
        grdLogged.Redraw = False
        grdLog.Redraw = False
        If (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Prg Start") Or (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Event Start") Then
            If (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Prg Start") Then
                ilRet = mFindLdf("P", "S")
            ElseIf (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Event Start") Then
                ilRet = mFindLdf("G", "S")
            Else
                ilRet = False
            End If
            If ilRet Then
                If Trim$(grdLogged.TextMatrix(llRowNo, AIREDINDEX)) <> "" Then
                    ''Do
                    ''    tmLlfSrchKey.lCode = tmLlf.lCode
                    ''    ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                    ''    If ilRet = BTRV_ERR_NONE Then
                    ''        slTime = grdLogged.TextMatrix(llRowNo, AIREDINDEX)
                    ''        gPackTime slTime, tmLlf.iStartTime(0), tmLlf.iStartTime(1)
                    ''        ilRet = btrUpdate(hmLlf, tmLlf, imLlfRecLen)
                    ''    Else
                    ''        Exit Do
                    ''    End If
                    ''Loop While ilRet = BTRV_ERR_CONFLICT
                    Do
                        'tmLDFSrchKey1.lLlfCode = tmLlf.lCode
                        'ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        'Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlf.lCode)
                        '    If (tmLdf.sType = "P") Or (tmLdf.sType = "G") Then
                        '        If tmLdf.sSubType = "S" Then
                        '            slTime = grdLogged.TextMatrix(llRowNo, AIREDINDEX)
                        '            gPackTime slTime, tmLdf.iTime(0), tmLdf.iTime(1)
                        '            ilRet = btrUpdate(hmLdf, tmLdf, imLdfRecLen)
                        '            Exit Do
                        '        End If
                        '    End If
                        '    ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                        'Loop
                        slTime = grdLogged.TextMatrix(llRowNo, AIREDINDEX)
                        gPackTime slTime, tmLdf.iTime(0), tmLdf.iTime(1)
                        ilRet = btrUpdate(hmLdf, tmLdf, imLdfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                Else
                    ''Do
                    ''    tmLlfSrchKey.lCode = tmLlf.lCode
                    ''    ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                    ''    If ilRet = BTRV_ERR_NONE Then
                    ''        slTime = ""
                    ''        gPackTime slTime, tmLlf.iStartTime(0), tmLlf.iStartTime(1)
                    ''        ilRet = btrUpdate(hmLlf, tmLlf, imLlfRecLen)
                    ''        ilMoveEvent = True
                    ''    Else
                    ''        Exit Do
                    ''    End If
                    ''Loop While ilRet = BTRV_ERR_CONFLICT
                    Do
                    '    tmLDFSrchKey1.lLlfCode = tmLlf.lCode
                    '    ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    '    Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlf.lCode)
                    '        If (tmLdf.sType = "P") Or (tmLdf.sType = "G") Then
                    '            If tmLdf.sSubType = "S" Then
                    '                ilRet = btrDelete(hmLdf)
                    '                Exit Do
                    '            End If
                    '        End If
                    '        ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    '    Loop
                        ilRet = btrDelete(hmLdf)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    ilMoveEvent = True
                    imGameStart = 0
                End If
            End If
        ElseIf (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Prg Over") Or (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Event Over") Then
            If (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Prg Over") Then
                ilRet = mFindLdf("P", "E")
            ElseIf (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Event Over") Then
                ilRet = mFindLdf("G", "E")
            Else
                ilRet = False
            End If
            If ilRet Then
                If Trim$(grdLogged.TextMatrix(llRowNo, AIREDINDEX)) <> "" Then
                    ''Do
                    ''    tmLlfSrchKey.lCode = tmLlf.lCode
                    ''    ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                    ''    If ilRet = BTRV_ERR_NONE Then
                    ''        slTime = grdLogged.TextMatrix(llRowNo, AIREDINDEX)
                    ''        gPackTime slTime, tmLlf.iEndTime(0), tmLlf.iEndTime(1)
                    ''        ilRet = btrUpdate(hmLlf, tmLlf, imLlfRecLen)
                    ''    Else
                    ''        Exit Do
                    ''    End If
                    ''Loop While ilRet = BTRV_ERR_CONFLICT
                    Do
                        'tmLDFSrchKey1.lLlfCode = tmLlf.lCode
                        'ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        'Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlf.lCode)
                        '    If (tmLdf.sType = "P") Or (tmLdf.sType = "G") Then
                        '        If tmLdf.sSubType = "E" Then
                        '            slTime = grdLogged.TextMatrix(llRowNo, AIREDINDEX)
                        '            gPackTime slTime, tmLdf.iTime(0), tmLdf.iTime(1)
                        '            ilRet = btrUpdate(hmLdf, tmLdf, imLdfRecLen)
                        '            Exit Do
                        '        End If
                        '    End If
                        '    ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                        'Loop
                        slTime = grdLogged.TextMatrix(llRowNo, AIREDINDEX)
                        gPackTime slTime, tmLdf.iTime(0), tmLdf.iTime(1)
                        ilRet = btrUpdate(hmLdf, tmLdf, imLdfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                Else
                    ''Do
                    ''    tmLlfSrchKey.lCode = tmLlf.lCode
                    ''    ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                    ''    If ilRet = BTRV_ERR_NONE Then
                    ''        slTime = ""
                    ''        gPackTime slTime, tmLlf.iEndTime(0), tmLlf.iEndTime(1)
                    ''        ilRet = btrUpdate(hmLlf, tmLlf, imLlfRecLen)
                    ''        ilMoveEvent = True
                    ''    Else
                    ''        Exit Do
                    ''    End If
                    ''Loop While ilRet = BTRV_ERR_CONFLICT
                    Do
                        'tmLDFSrchKey1.lLlfCode = tmLlf.lCode
                        'ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        'Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlf.lCode)
                        '    If (tmLdf.sType = "P") Or (tmLdf.sType = "G") Then
                        '        If tmLdf.sSubType = "E" Then
                        '            ilRet = btrDelete(hmLdf)
                        '            Exit Do
                        '        End If
                        '    End If
                        '    ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                        'Loop
                        ilRet = btrDelete(hmLdf)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    ilMoveEvent = True
                    imGameStart = 1
                End If
            End If
        Else
            If (grdLogged.TextMatrix(llRowNo, AIREDINDEX) = "") And (grdLogged.TextMatrix(llRowNo, NOTAIREDINDEX) = "") And (grdLogged.TextMatrix(llRowNo, VEHICLETYPEINDEX) <> "S") Then
                tmSdfSrchKey3.lCode = Val(grdLogged.TextMatrix(llRowNo, SDFCODEINDEX))
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                If ilRet <> BTRV_ERR_NONE Then
                    tmSdf.lCode = Val(grdLogged.TextMatrix(llRowNo, SDFCODEINDEX))
                    For ilSdf = 0 To UBound(tmDelOrMissedSdf) - 1 Step 1
                        If tmSdf.lCode = tmDelOrMissedSdf(ilSdf).lCode Then
                            tlSdf = tmDelOrMissedSdf(ilSdf)
                            tlSdf.sSchStatus = "M"
                            tlSdf.iUrfCode = tgUrf(0).iCode
                            ilRet = btrInsert(hmSdf, tlSdf, imSdfRecLen, INDEXKEY0)
                            Exit For
                        End If
                    Next ilSdf
                End If
            End If
            Do
                tmSdfSrchKey3.lCode = Val(grdLogged.TextMatrix(llRowNo, SDFCODEINDEX))
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                If ilRet <> BTRV_ERR_NONE Then
                    For ilSdf = 0 To UBound(tmDelOrMissedSdf) - 1 Step 1
                        tmSdf.lCode = Val(grdLogged.TextMatrix(llRowNo, SDFCODEINDEX))
                        If tmSdf.lCode = tmDelOrMissedSdf(ilSdf).lCode Then
                            tmSdf = tmDelOrMissedSdf(ilSdf)
                            tmSdf.sSchStatus = "M"
                            tmSdf.iUrfCode = tgUrf(0).iCode
                            ilRet = btrInsert(hmSdf, tmSdf, imSdfRecLen, INDEXKEY0)
                            Exit For
                        End If
                    Next ilSdf
                End If
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrGetPosition(hmSdf, llSdfRecPos)
                    If grdLogged.TextMatrix(llRowNo, VEHICLETYPEINDEX) <> "S" Then
                        If (grdLogged.TextMatrix(llRowNo, AIREDINDEX) = "") And (grdLogged.TextMatrix(llRowNo, NOTAIREDINDEX) = "") Then
                            If grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Added Spot" Then
                                Do
                                    tmSdfSrchKey3.lCode = tmSdf.lCode
                                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                                    ilRet = btrDelete(hmSdf)
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                grdLogged.RemoveItem llRowNo
                                grdLogged.Redraw = True
                                grdLog.Redraw = True
                                Exit Sub
                            End If
                            'Move spot back
                            slStr = grdLogged.TextMatrix(llRowNo, ORIGTIMEINDEX)
                            llAvailTime = gTimeToLong(slStr, False)
                            ilAvailIndex = mFindAvailIndex(tmSdf.iVefCode, llAvailTime)
                            ilFound = False
                            For ilSdf = 0 To UBound(tmDelOrMissedSdf) - 1 Step 1
                                If tmSdf.lCode = tmDelOrMissedSdf(ilSdf).lCode Then
                                    If grdLogged.TextMatrix(llRowNo, TYPEINDEX) <> "Added Spot" Then
                                        tlSdf = tmSdf
                                        ilRet = mReadChfClfRdfRec(tlSdf.lChfCode, tlSdf.iLineNo, tlSdf.lFsfCode)
                                        If ilRet Then
                                            slSchStatus = tmDelOrMissedSdf(ilSdf).sSchStatus
                                            ilBkQH = Val(grdLogged.TextMatrix(llRowNo, BKQHINDEX)) And RANKMASK
                                            ilPriceLevel = (Val(grdLogged.TextMatrix(llRowNo, BKQHINDEX)) And PRICELEVELMASK) / SHIFT11
                                            ilRet = gBookSpot(slSchStatus, hmSdf, tmSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf, lmSsfRecPos, ilAvailIndex, -1, tmChf, tmClf, tmRdf, imVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, ilPriceLevel, False, hmGsf)
                                            If ilRet Then
                                                ilMoveEvent = True
                                                tmSdf.sPtType = tlSdf.sPtType
                                                tmSdf.lCopyCode = tlSdf.lCopyCode
                                                tmSdf.iMnfMissed = 0
                                                tmSdf.sAffChg = ""
                                                tmSdf.sXCrossMidnight = "N"
                                                tmSdf.iUrfCode = tgUrf(0).iCode
                                                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                                ilFound = True
                                            End If
                                        Else
                                            ilMoveEvent = False
                                        End If
                                    Else
                                    End If
                                    Exit For
                                End If
                            Next ilSdf
                            If Not ilFound Then
                                ilMoveEvent = True
                                slStr = grdLogged.TextMatrix(llRowNo, ORIGTIMEINDEX)
                                gPackTime slStr, tmSdf.iTime(0), tmSdf.iTime(1)
                                tmSdf.iMnfMissed = 0
                                tmSdf.sAffChg = ""
                                tmSdf.sXCrossMidnight = "N"
                                tmSdf.iUrfCode = tgUrf(0).iCode
                                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                            End If
                        ElseIf grdLogged.TextMatrix(llRowNo, NOTAIREDINDEX) = "" Then
                            'Change time of spot
                            If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
                                ilMoveEvent = False
                                tmSdf.sXCrossMidnight = mCrossMidnight(grdLogged.TextMatrix(llRowNo, ORIGTIMEINDEX), grdLogged.TextMatrix(llRowNo, AIREDINDEX))
                                gPackTime grdLogged.TextMatrix(llRowNo, AIREDINDEX), tmSdf.iTime(0), tmSdf.iTime(1)
                                tmSdf.sAffChg = "Y"
                                tmSdf.iUrfCode = tgUrf(0).iCode
                                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                            Else
                                'Rebook spot
                                slStr = grdLogged.TextMatrix(llRowNo, ORIGTIMEINDEX)
                                llAvailTime = gTimeToLong(slStr, False)
                                ilAvailIndex = mFindAvailIndex(tmSdf.iVefCode, llAvailTime)
                                For ilSdf = 0 To UBound(tmDelOrMissedSdf) - 1 Step 1
                                    If tmSdf.lCode = tmDelOrMissedSdf(ilSdf).lCode Then
                                        tlSdf = tmSdf
                                        ilRet = mReadChfClfRdfRec(tlSdf.lChfCode, tlSdf.iLineNo, tlSdf.lFsfCode)
                                        If ilRet Then
                                            slSchStatus = tmDelOrMissedSdf(ilSdf).sSchStatus
                                            ilBkQH = Val(grdLogged.TextMatrix(llRowNo, BKQHINDEX)) And RANKMASK
                                            ilPriceLevel = (Val(grdLogged.TextMatrix(llRowNo, BKQHINDEX)) And PRICELEVELMASK) / SHIFT11
                                            ilRet = gBookSpot(slSchStatus, hmSdf, tmSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf, lmSsfRecPos, ilAvailIndex, -1, tmChf, tmClf, tmRdf, imVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, ilPriceLevel, False, hmGsf)
                                            tmSdfSrchKey3.lCode = tmSdf.lCode
                                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                                            tmSdf.sPtType = tlSdf.sPtType
                                            tmSdf.lCopyCode = tlSdf.lCopyCode
                                            gPackTime grdLogged.TextMatrix(llRowNo, AIREDINDEX), tmSdf.iTime(0), tmSdf.iTime(1)
                                            tmSdf.iMnfMissed = 0
                                            tmSdf.sAffChg = "Y"
                                            tmSdf.sXCrossMidnight = "N"
                                            tmSdf.iUrfCode = tgUrf(0).iCode
                                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                            grdLogged.Row = llRowNo
                                            For llCol = 0 To grdLogged.Cols - 1 Step 1
                                                grdLogged.Col = llCol
                                                grdLogged.CellForeColor = DARKGREEN 'vbGreen
                                            Next llCol
                                        Else
                                            ilMoveEvent = False
                                        End If
                                        Exit For
                                    End If
                                Next ilSdf
                            End If
                        Else
                            'Missed Spot
                            If tmSdf.sSchStatus <> "M" Then
                                tmChfSrchKey.lCode = tmSdf.lChfCode
                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    'Force reading of SSF
                                    lmSsfDate = 0
                                    ilGameNo = tmSdf.iGameNo
                                    If ((tmChf.sType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmChf.sType = "Q") Or ((tmChf.sType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmChf.sType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (tmSdf.sSpotType = "X") Then
                                        tmDelOrMissedSdf(UBound(tmDelOrMissedSdf)) = tmSdf
                                        ReDim Preserve tmDelOrMissedSdf(0 To UBound(tmDelOrMissedSdf) + 1) As SDF
                                        ilRet = gChgSchSpot("D", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf, lmSsfDate, lmSsfRecPos, hmSxf, hmGsf, hmGhf)
                                    ElseIf (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
                                        tmDelOrMissedSdf(UBound(tmDelOrMissedSdf)) = tmSdf
                                        ReDim Preserve tmDelOrMissedSdf(0 To UBound(tmDelOrMissedSdf) + 1) As SDF
                                        ilRet = btrDelete(hmSdf)
                                    Else
                                        tmDelOrMissedSdf(UBound(tmDelOrMissedSdf)) = tmSdf
                                        ReDim Preserve tmDelOrMissedSdf(0 To UBound(tmDelOrMissedSdf) + 1) As SDF
                                        ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf, lmSsfDate, lmSsfRecPos, hmSxf, hmGsf, hmGhf)
                                        If ilRet Then
                                            'Set Missed reason
                                            tmSdf.iMnfMissed = igDefaultMnfMissed
                                            slMissedReason = grdLogged.TextMatrix(llRowNo, NOTAIREDINDEX)
                                            gFindMatch slMissedReason, 0, lbcNotAired
                                            If gLastFound(lbcNotAired) >= 0 Then
                                                slNameCode = tmMissedCode(gLastFound(lbcNotAired)).sKey
                                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                                tmSdf.iMnfMissed = Val(slCode)
                                            End If
                                            tmSdf.sAffChg = ""
                                            tmSdf.iUrfCode = tgUrf(0).iCode
                                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                        End If
                                    End If
                                    grdLogged.Row = llRowNo
                                    For llCol = 0 To grdLogged.Cols - 1 Step 1
                                        grdLogged.Col = llCol
                                        grdLogged.CellForeColor = vbRed
                                    Next llCol
                                End If
                            Else
                                slMissedReason = grdLogged.TextMatrix(llRowNo, NOTAIREDINDEX)
                                gFindMatch slMissedReason, 0, lbcNotAired
                                If gLastFound(lbcNotAired) >= 0 Then
                                    slNameCode = tmMissedCode(gLastFound(lbcNotAired)).sKey
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    tmSdf.iMnfMissed = Val(slCode)
                                Else
                                    tmSdf.iMnfMissed = igDefaultMnfMissed
                                End If
                                tmSdf.sAffChg = ""
                                tmSdf.iUrfCode = tgUrf(0).iCode
                                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                            End If
                        End If
                    Else
                        ilMoveEvent = True
                        ilLastMatchingSpot = True
                        For llRowTest = grdLogged.FixedRows To grdLogged.Rows - 1 Step 1
                            If llRowTest <> llRowNo Then
                                If grdLogged.TextMatrix(llRowNo, SDFCODEINDEX) = grdLogged.TextMatrix(llRowTest, SDFCODEINDEX) Then
                                    ilLastMatchingSpot = False
                                    Exit For
                                End If
                            End If
                        Next llRowTest
                        If ilLastMatchingSpot Then
                            tmSdf.sAffChg = ""
                            tmSdf.iUrfCode = tgUrf(0).iCode
                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                        End If
                    End If
                Else
                    Exit Do
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
        If ilMoveEvent Then
            If grdLog.TextMatrix(grdLog.FixedRows, TYPEINDEX) <> "" Then
                grdLog.AddItem "", grdLog.FixedRows
            End If
            grdLog.Row = grdLog.FixedRows
            llRow = grdLog.FixedRows
            grdLog.Col = AIREDINDEX
            grdLog.CellBackColor = vbGreen
            grdLog.Col = NOTAIREDINDEX
            grdLog.CellBackColor = vbRed
            grdLog.Col = SCHDTIMEINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = LENGTHINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = TYPEINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            grdLog.Col = ADVTPRODINDEX
            grdLog.CellBackColor = LIGHTYELLOW
            If (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Prg Start") Or (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Event Start") Then
                grdLog.Col = CERTIYINDEX
                grdLog.CellBackColor = LIGHTYELLOW
                grdLog.Col = COPYINDEX
                grdLog.CellBackColor = LIGHTYELLOW
            ElseIf (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Prg Over") Or (grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Event Over") Then
                grdLog.Col = CERTIYINDEX
                grdLog.CellBackColor = LIGHTYELLOW
                grdLog.Col = COPYINDEX
                grdLog.CellBackColor = LIGHTYELLOW
            Else
                If grdLogged.TextMatrix(llRowNo, VEHICLETYPEINDEX) = "S" Then
                    grdLog.Col = AIREDINDEX
                    grdLog.CellBackColor = LIGHTYELLOW
                    grdLog.Col = NOTAIREDINDEX
                    grdLog.CellBackColor = LIGHTYELLOW
                    grdLog.Col = COPYINDEX
                    grdLog.CellBackColor = LIGHTYELLOW
                    grdLog.Col = CERTIYINDEX
                    grdLog.CellBackColor = CYAN 'DARKGREEN    'vbGreen
                    For llCol = 0 To grdLog.Cols - 1 Step 1
                        grdLog.Col = llCol
                        grdLog.CellForeColor = MIDDLEBLUE 'CYAN
                    Next llCol
                Else
                    grdLog.Col = CERTIYINDEX
                    grdLog.CellBackColor = LIGHTYELLOW
                End If
            End If
            For llCol = 0 To grdLog.Cols - 1 Step 1
                grdLog.Col = llCol
                grdLog.CellAlignment = flexAlignLeftTop
            Next llCol
            For llCol = 0 To grdLog.Cols - 1 Step 1
                grdLog.Col = llCol
                If grdLogged.TextMatrix(llRowNo, VEHICLETYPEINDEX) <> "S" Then
                    grdLog.CellForeColor = vbBlack
                End If
                If llCol = AIREDINDEX Then
                    grdLog.TextMatrix(llRow, llCol) = ""
                ElseIf llCol = NOTAIREDINDEX Then
                    grdLog.TextMatrix(llRow, llCol) = ""
                Else
                    grdLog.TextMatrix(llRow, llCol) = grdLogged.TextMatrix(llRowNo, llCol)
                End If
            Next llCol
            If grdLogged.FixedRows = grdLogged.Rows - 1 Then
                For llCol = 0 To grdLogged.Cols - 1 Step 1
                    grdLogged.TextMatrix(grdLogged.FixedRows, llCol) = ""
                Next llCol
            Else
                grdLogged.RemoveItem llRowNo
            End If
        End If
        grdLogged.Redraw = True
        grdLog.Redraw = True
        Do While Not grdLog.RowIsVisible(grdLog.Rows - 1)
            grdLog.TopRow = grdLog.TopRow + 1
        Loop
    End If
    If smVefType <> "L" Then
        If (imGameStart = 1) Or (imGameStart = 2) Then
            cmcAddSpot.Enabled = True
        Else
            cmcAddSpot.Enabled = False
        End If
    Else
        cmcAddSpot.Enabled = False
    End If
End Sub

Private Sub lacBanner_Click(Index As Integer)
    mSetShow True
End Sub

Private Sub lacHelp_Click()
    mSetShow True
    igHelpScreen = 1
    If pbcLog.Visible Then
        igHelpScreen = 2
    End If
    LLHelp.Show vbModal
End Sub

Private Sub lbcCopy_Click()
    gProcessLbcClick lbcCopy, edcDropdown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcCopy_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcNotAired_Click()
    gProcessLbcClick lbcNotAired, edcDropdown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcNotAired_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcBanner_Click()
    mSetShow True
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow True
    grdLog.Col = LIBSORTINDEX
    grdLog.Row = 0
    grdLogged.Col = BKQHINDEX
    grdLogged.Row = 0
    edcDropdown.Visible = False
    cmcDropdown.Visible = False
    lbcNotAired.Visible = False
    lbcCopy.Visible = False
    plcTme.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub pbcDateTab_GotFocus()
    tmcDelay.Enabled = False
    tmcDelay_Timer
End Sub

Private Sub pbcFocusStop_GotFocus()
    mSetFocus
End Sub

Private Sub pbcLog_GotFocus()
    mSetShow True
End Sub

Private Sub pbcSTab_GotFocus()
    Dim slStr As String

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        If lmEnableCol = AIREDINDEX Then
            slStr = edcDropdown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropdown.SetFocus
                    Exit Sub
                End If
            End If
        End If
        mSetShow True
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim slStr As String

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        If lmEnableCol = AIREDINDEX Then
            slStr = edcDropdown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropdown.SetFocus
                    Exit Sub
                End If
            End If
        End If
        mSetShow True
    End If
End Sub

Private Sub mAddSpotToGrid(grdGrid As MSHFlexGrid, ilWhichGrid As Integer, slAvailTime As String, slRunTime As String, Optional ilPosition As Integer = 1)
    Dim llRow As Long
    Dim ilAdf As Integer
    Dim slCart As String
    Dim slISCI As String
    Dim slProduct As String
    Dim ilCol As Integer
    Dim ilRet As Integer
    Dim slTime As String
    Dim ilMnf As Integer
    Dim slName As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLive As Integer
    Dim slSortTime As String
    Dim ilVef As Integer
    Dim slPosition As String
    Dim blAddAdvt As Boolean

    If grdGrid.TextMatrix(grdGrid.Rows - 1, TYPEINDEX) <> "" Then
        grdGrid.AddItem ""
    End If
    grdGrid.Row = grdGrid.Rows - 1
    llRow = grdGrid.Row
    For ilCol = 0 To grdGrid.Cols - 1 Step 1
        grdGrid.Col = ilCol
        grdGrid.CellAlignment = flexAlignLeftTop
    Next ilCol
    mObtainCopy slCart, slISCI, slProduct
    If ilWhichGrid = 0 Then
        grdGrid.Col = AIREDINDEX
        grdGrid.CellBackColor = vbGreen
        grdGrid.Col = NOTAIREDINDEX
        grdGrid.CellBackColor = vbRed
    Else
        If tmSdf.iMnfMissed > 0 Then
            For ilMnf = LBound(tmMissedCode) To UBound(tmMissedCode) - 1 Step 1
                slNameCode = tmMissedCode(ilMnf).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If tmSdf.iMnfMissed = Val(slCode) Then
                    slNameCode = tmMissedCode(ilMnf).sKey
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    grdGrid.TextMatrix(llRow, NOTAIREDINDEX) = slName
                    Exit For
                End If
            Next ilMnf
            For ilCol = 0 To grdGrid.Cols - 1 Step 1
                grdGrid.Col = ilCol
                grdGrid.ForeColor = vbRed
            Next ilCol
        Else
            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
            grdGrid.TextMatrix(llRow, AIREDINDEX) = slTime
            For ilCol = 0 To grdGrid.Cols - 1 Step 1
                grdGrid.Col = ilCol
                grdGrid.ForeColor = DARKGREEN
            Next ilCol
        End If
    End If
    grdGrid.Col = SCHDTIMEINDEX
    grdGrid.CellBackColor = LIGHTYELLOW
    grdGrid.TextMatrix(llRow, SCHDTIMEINDEX) = slRunTime 'slTime
    grdGrid.TextMatrix(llRow, ORIGTIMEINDEX) = slAvailTime
    grdGrid.Col = LENGTHINDEX
    grdGrid.CellBackColor = LIGHTYELLOW
    grdGrid.TextMatrix(llRow, LENGTHINDEX) = Trim$(Str$(tmSdf.iLen))
    grdGrid.Col = TYPEINDEX
    grdGrid.CellBackColor = LIGHTYELLOW
    grdGrid.TextMatrix(llRow, TYPEINDEX) = ""
    If tmAvail.ianfCode > 0 Then
        tmAnfSrchKey.iCode = tmAvail.ianfCode
        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            grdGrid.TextMatrix(llRow, TYPEINDEX) = Trim$(tmAnf.sName)
        End If
    Else
        If (slRunTime = "") And (slAvailTime = "") Then
            grdGrid.TextMatrix(llRow, TYPEINDEX) = "Added Spot"
        End If
    End If
    grdGrid.Col = ADVTPRODINDEX
    grdGrid.CellBackColor = LIGHTYELLOW
    ilAdf = gBinarySearchAdf(tmSdf.iAdfCode)
    If ilAdf <> -1 Then
        If slProduct = "" Then
            tmChfSrchKey.lCode = tmSdf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slProduct = Trim$(tmChf.sProduct)
            End If
        Else
        End If
        If Trim$(slProduct) = "" Then
            grdGrid.TextMatrix(llRow, ADVTPRODINDEX) = Trim$(tgCommAdf(ilAdf).sName)
        Else
            grdGrid.TextMatrix(llRow, ADVTPRODINDEX) = Trim$(tgCommAdf(ilAdf).sName) & "/" & slProduct
        End If
        blAddAdvt = True
        For ilAdf = 0 To UBound(igEventAdfCode) Step 1
            If igEventAdfCode(ilAdf) = tmSdf.iAdfCode Then
                blAddAdvt = False
                Exit For
            End If
        Next ilAdf
        If blAddAdvt Then
            igEventAdfCode(UBound(igEventAdfCode)) = tmSdf.iAdfCode
            ReDim Preserve igEventAdfCode(0 To UBound(igEventAdfCode) + 1) As Integer
        End If
    End If
    grdGrid.ColWordWrapOption(ADVTPRODINDEX) = True
    grdGrid.Col = COPYINDEX
    grdGrid.CellAlignment = flexAlignLeftTop
    If Trim$(slCart) <> "" Then
        grdGrid.TextMatrix(llRow, COPYINDEX) = Trim$(slCart)
    Else
        grdGrid.TextMatrix(llRow, COPYINDEX) = Trim$(slISCI)
    End If
    grdGrid.TextMatrix(llRow, ORIGCOPYINDEX) = grdGrid.TextMatrix(llRow, COPYINDEX)
    grdGrid.TextMatrix(llRow, SDFCODEINDEX) = tmSdf.lCode
    grdGrid.TextMatrix(llRow, POSITIONINDEX) = ilPosition
    slPosition = ilPosition
    Do While Len(slPosition) < 2
        slPosition = "0" & slPosition
    Loop
    ilVef = gBinarySearchVef(tmSdf.iVefCode)
    If ilVef <> -1 Then
        grdGrid.TextMatrix(llRow, VEHICLETYPEINDEX) = tgMVef(ilVef).sType
    Else
        grdGrid.TextMatrix(llRow, VEHICLETYPEINDEX) = ""
    End If
    If grdGrid.TextMatrix(llRow, VEHICLETYPEINDEX) = "S" Then
        If ilWhichGrid = 0 Then
            grdGrid.Col = CERTIYINDEX
            grdGrid.CellBackColor = CYAN    'DARKGREEN   'vbGreen
        End If
        grdGrid.Col = AIREDINDEX
        grdGrid.CellBackColor = LIGHTYELLOW
        grdGrid.Col = NOTAIREDINDEX
        grdGrid.CellBackColor = LIGHTYELLOW
        grdGrid.Col = COPYINDEX
        grdGrid.CellBackColor = LIGHTYELLOW
        For ilCol = 0 To grdGrid.Cols - 1 Step 1
            grdGrid.Col = ilCol
            grdGrid.CellForeColor = MIDDLEBLUE 'CYAN
        Next ilCol
    Else
        grdGrid.Col = CERTIYINDEX
        grdGrid.CellBackColor = LIGHTYELLOW
    End If
    slSortTime = Trim$(Str$(gTimeToLong(grdGrid.TextMatrix(llRow, SCHDTIMEINDEX), False)))
    Do While Len(slSortTime) < 6
        slSortTime = "0" & slSortTime
    Loop
    grdGrid.TextMatrix(llRow, SDFSORTINDEX) = slSortTime & slPosition
    grdGrid.TextMatrix(llRow, BKQHINDEX) = tmSpot.iRank
    ilLive = False
    If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
        If tmMcf.iCode <> tmCif.iMcfCode Then
            tmMcfSrchKey.iCode = tmCif.iMcfCode
            ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        End If
        If (tmCif.lCsfCode > 0) And (tmMcf.sScript = "Y") Then
            ilLive = True
        End If
    Else
        If (tmCif.lCsfCode > 0) Then
            ilLive = True
        End If
    End If
    If (ilLive) And (ilWhichGrid = 0) Then
        grdGrid.Col = SCHDTIMEINDEX
        grdGrid.CellBackColor = vbCyan
        grdGrid.Col = LENGTHINDEX
        grdGrid.CellBackColor = vbCyan
        grdGrid.Col = TYPEINDEX
        grdGrid.CellBackColor = vbCyan
        grdGrid.Col = ADVTPRODINDEX
        grdGrid.CellBackColor = vbCyan
    End If
End Sub

Private Function mGridEnable(grdGrid As MSHFlexGrid) As Integer
    Dim slStr As String
    Dim slTime As String
    If (grdGrid.Row < grdGrid.FixedRows) Or (grdGrid.Row >= grdGrid.Rows) Or (grdGrid.Col < grdGrid.FixedCols) Or (grdGrid.Col >= grdGrid.Cols - 1) Then
        mGridEnable = False
        Exit Function
    End If
    lmEnableRow = grdGrid.Row
    lmEnableCol = grdGrid.Col
    Select Case grdGrid.Col
        Case AIREDINDEX
            edcDropdown.MaxLength = 10
            slStr = grdGrid.Text
            If grdGrid.Text = "" Then
                slTime = Format$(gNow(), "h:mm:ssAM/PM")
                edcDropdown.Text = slTime
            Else
                edcDropdown.Text = grdGrid.Text
            End If
        Case NOTAIREDINDEX
            mMissedPop
            lbcNotAired.Height = gListBoxHeight(lbcNotAired.ListCount, 10)
            edcDropdown.MaxLength = 20
            imChgMode = True
            slStr = grdGrid.Text
            gFindMatch slStr, 0, lbcNotAired
            If gLastFound(lbcNotAired) >= 0 Then
                lbcNotAired.ListIndex = gLastFound(lbcNotAired)
                edcDropdown.Text = lbcNotAired.List(lbcNotAired.ListIndex)
            Else
                If lbcNotAired.ListCount > 0 Then
                    lbcNotAired.ListIndex = 0
                    edcDropdown.Text = lbcNotAired.List(lbcNotAired.ListIndex)
                Else
                    edcDropdown.Text = ""
                End If
            End If
            imChgMode = False
        Case COPYINDEX
            mCopyPop
            lbcCopy.Height = gListBoxHeight(lbcCopy.ListCount, 10)
            If (tgSpf.sUseCartNo <> "N") Then
                edcDropdown.MaxLength = 10
            Else
                edcDropdown.MaxLength = 20
            End If
            imChgMode = True
            slStr = grdGrid.Text
            gFindMatch slStr, 0, lbcCopy
            If gLastFound(lbcCopy) >= 0 Then
                lbcCopy.ListIndex = gLastFound(lbcCopy)
                edcDropdown.Text = lbcCopy.List(lbcCopy.ListIndex)
            Else
                If lbcCopy.ListCount > 0 Then
                    lbcCopy.ListIndex = 0
                    edcDropdown.Text = lbcCopy.List(lbcCopy.ListIndex)
                Else
                    edcDropdown.Text = ""
                End If
            End If
            imComboBoxIndex = lbcCopy.ListIndex
            imChgMode = False
    End Select
    mGridEnable = True
End Function

Private Function mGridSetFocus(grdGrid As MSHFlexGrid) As Integer
    Dim llColPos As Long
    Dim ilCol As Integer

    If (grdGrid.Row < grdGrid.FixedRows) Or (grdGrid.Row >= grdGrid.Rows) Or (grdGrid.Col < grdGrid.FixedCols) Or (grdGrid.Col >= grdGrid.Cols - 1) Then
        mGridSetFocus = False
        Exit Function
    End If
    imCtrlVisible = True
    If imAllowUndo Then
        cmcUndo.Enabled = True
    End If
    llColPos = 0
    For ilCol = 0 To grdGrid.Col - 1 Step 1
        llColPos = llColPos + grdGrid.ColWidth(ilCol)
    Next ilCol
    Select Case grdGrid.Col
        Case AIREDINDEX
            edcDropdown.Move pbcLog.Left + grdGrid.Left + llColPos + 45, pbcLog.Top + grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 45, grdGrid.ColWidth(grdGrid.Col) - 30, grdGrid.RowHeight(grdGrid.Row) - 15
            cmcDropdown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropdown.Width, edcDropdown.Height
            plcTme.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height
            If plcTme.Top + plcTme.Height > cmcSignOn.Top Then
                plcTme.Top = edcDropdown.Top - plcTme.Height
            End If
            edcDropdown.Visible = True
            cmcDropdown.Visible = True
            If edcDropdown.Text = "" Then
                plcTme.Visible = True
            End If
            edcDropdown.SetFocus
        Case NOTAIREDINDEX
            edcDropdown.Move pbcLog.Left + grdGrid.Left + llColPos + 45, pbcLog.Top + grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 45, lbcNotAired.Width - cmcDropdown.Width, grdGrid.RowHeight(grdGrid.Row) - 15
            cmcDropdown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropdown.Width, edcDropdown.Height
            lbcNotAired.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropdown.Width
            If lbcNotAired.Top + lbcNotAired.Height > cmcSignOn.Top Then
                lbcNotAired.Top = edcDropdown.Top - lbcNotAired.Height
            End If
            edcDropdown.Visible = True
            cmcDropdown.Visible = True
            edcDropdown.SetFocus
        Case COPYINDEX
            lbcCopy.Width = 2 * (grdGrid.Width - llColPos)
            edcDropdown.Move pbcLog.Left + grdGrid.Left + grdGrid.Width - lbcCopy.Width - GRIDSCROLLWIDTH, pbcLog.Top + grdGrid.Top + grdGrid.RowPos(grdGrid.Row) + 45, lbcCopy.Width - cmcDropdown.Width, grdGrid.RowHeight(grdGrid.Row) - 15
            cmcDropdown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropdown.Width, edcDropdown.Height
            lbcCopy.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropdown.Width
            If lbcCopy.Top + lbcCopy.Height > cmcSignOn.Top Then
                lbcCopy.Top = edcDropdown.Top - lbcCopy.Height
            End If
            edcDropdown.Visible = True
            cmcDropdown.Visible = True
            edcDropdown.SetFocus
    End Select
    mGridSetFocus = True
End Function

Private Function mGridSetShow(grdGrid As MSHFlexGrid, ilInMoveRow As Integer) As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilUpdateCopy As Integer
    Dim ilMoveRow As Integer

    ilUpdateCopy = False
    ilMoveRow = ilInMoveRow
    If (lmEnableRow >= grdGrid.FixedRows) And (lmEnableRow < grdGrid.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case AIREDINDEX
                cmcDropdown.Visible = False
                plcTme.Visible = False
                edcDropdown.Visible = False  'Set visibility
                slStr = edcDropdown.Text
                If slStr <> "" Then
                    If gValidTime(slStr) Then
                        If gTimeToLong(grdGrid.TextMatrix(lmEnableRow, lmEnableCol), False) <> gTimeToLong(slStr, False) Then
                            slStr = gFormatTime(slStr, "A", "1")
                            grdGrid.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                            grdGrid.TextMatrix(lmEnableRow, NOTAIREDINDEX) = ""
                            'Update Spot
                        End If
                    End If
                Else
                    If (imWhichGrid = 0) Or ((imWhichGrid = 1) And (imAllowUndo)) Then
                        If mUndoAllowed(grdGrid, lmEnableRow) Then
                            grdGrid.TextMatrix(lmEnableRow, lmEnableCol) = ""
                        Else
                            ilMoveRow = False
                        End If
                    Else
                        ilMoveRow = False
                    End If
                End If
            Case NOTAIREDINDEX
                edcDropdown.Visible = False
                cmcDropdown.Visible = False
                lbcNotAired.Visible = False
                slStr = Trim$(edcDropdown.Text)
                If slStr <> "" Then
                    ilRet = gOptionalLookAhead(edcDropdown, lbcNotAired, imBSMode, slStr)
                    If ilRet <> 0 Then
                        'Add new Missed reason
                        mAddMissedReason slStr
                    End If
                    grdGrid.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    grdGrid.TextMatrix(lmEnableRow, AIREDINDEX) = ""
                Else
                    If (imWhichGrid = 0) Or ((imWhichGrid = 1) And (imAllowUndo)) Then
                        grdGrid.TextMatrix(lmEnableRow, lmEnableCol) = ""
                    Else
                        ilMoveRow = False
                    End If
                End If
            Case COPYINDEX
                edcDropdown.Visible = False
                cmcDropdown.Visible = False
                lbcCopy.Visible = False
                If lbcCopy.ListIndex >= 0 Then
                    If grdGrid.TextMatrix(lmEnableRow, lmEnableCol) <> lbcCopy.List(lbcCopy.ListIndex) Then
                        grdGrid.TextMatrix(lmEnableRow, lmEnableCol) = lbcCopy.List(lbcCopy.ListIndex)
                        'Update spot
                        mUpdateCopy grdGrid
                    End If
                Else
                    grdGrid.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
        End Select
        If ilMoveRow Then
            If grdGrid.TextMatrix(lmEnableRow, VEHICLETYPEINDEX) = "S" Then
                mMoveRow lmEnableRow
            Else
                If (grdGrid.TextMatrix(lmEnableRow, AIREDINDEX) <> "") Or (grdGrid.TextMatrix(lmEnableRow, NOTAIREDINDEX) <> "") Then
                    'Update spots and maybe move spot
                    mMoveRow lmEnableRow
                Else
                    If imWhichGrid = 1 Then
                        'Undo operation
                        mMoveRow lmEnableRow
                    End If
                End If
            End If
        End If
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    cmcUndo.Enabled = False
    DoEvents
    mGridSetShow = True
End Function

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
                    Select Case lmEnableCol
                        Case AIREDINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropdown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropdown, slKey
                    End Select
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
'*      Procedure Name:mReadChfClfRdfRec               *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read records                   *
'*                                                     *
'*******************************************************
Private Function mReadChfClfRdfRec(llChfCode As Long, ilLineNo As Integer, llFsfCode As Long) As Integer
'
'   iRet = mReadChfClfRdfRec(llChfCode, ilLineNo)
'   Where:
'       llChfCode (I) - Contract code
'       ilLineNo(I)- Line number
'       iRet (O)- True if all records read,
'                 False if error in read
'
    Dim ilRet As Integer
    If llChfCode > 0 Then
        'If llChfCode <> tmChf.lCode Then
            tmChfSrchKey.lCode = llChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                mReadChfClfRdfRec = False
                Exit Function
            End If
        'End If
        'If (tmClf.lChfCode <> llChfCode) Or (tmClf.iLine <> ilLineNo) Then
            tmClfSrchKey.lChfCode = llChfCode
            tmClfSrchKey.iLine = ilLineNo
            tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
            tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iLine = ilLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))    'And (tmClf.sSchStatus = "A")
                ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        'Else
        '    ilRet = BTRV_ERR_NONE
        'End If
        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iLine = ilLineNo) Then
            If tmRdf.iCode <> tmClf.iRdfCode Then
                tmRdfSrchKey.iCode = tmClf.iRdfCode  ' Rate card program/time File Code
                ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    mReadChfClfRdfRec = False
                    Exit Function
                End If
            End If
            mReadChfClfRdfRec = True
        Else
            mReadChfClfRdfRec = False
        End If
    Else
        tmFSFSrchKey.lCode = llFsfCode
        ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        gMoveFeedToCntr tmFsf, tmRdf, tmChf, tmClf, tmFCff(), hmFnf, hmPrf
        tmCff = tmFCff(1)
        mReadChfClfRdfRec = True
    End If
    Exit Function
End Function


Private Function mFindAvailIndex(ilVefCode As Integer, llAvailTime As Long) As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRet As Integer
    Dim ilEvt As Integer
    Dim llTime As Long

    mFindAvailIndex = -1
    gPackDateLong lmEventDate, ilDate0, ilDate1   'lmNowDate, ilDate0, ilDate1
    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
    tmSsfSrchKey.iType = imGameNo 'slType-On Air
    tmSsfSrchKey.iVefCode = ilVefCode   'imVefCode
    tmSsfSrchKey.iDate(0) = ilDate0
    tmSsfSrchKey.iDate(1) = ilDate1
    tmSsfSrchKey.iStartTime(0) = 0
    tmSsfSrchKey.iStartTime(1) = 0
    ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = imGameNo) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1) Then
        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = imGameNo) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1)
            ilRet = btrGetPosition(hmSsf, lmSsfRecPos)
            For ilEvt = 1 To tmSsf.iCount Step 1
               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                    If llTime = llAvailTime Then
                        mFindAvailIndex = ilEvt
                        Exit Function
                    End If
                End If
            Next ilEvt
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If

End Function
Private Function mAddSpotFindAvailIndex(llAvailTime As Long, ilLen As Integer) As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRet As Integer
    Dim ilEvt As Integer
    Dim llTime As Long
    Dim ilAvailIndex As Integer
    Dim ilSpotLen As Integer
    Dim ilSpotUnits As Integer
    Dim ilAvailUnits As Integer
    Dim ilAvailSec As Integer

    mAddSpotFindAvailIndex = -1
    ilAvailIndex = -1
    gPackDateLong lmEventDate, ilDate0, ilDate1   'lmNowDate, ilDate0, ilDate1
    ilSpotLen = ilLen
    If (tgVpf(imVpfIndex).sSSellOut = "T") Then
        ilSpotUnits = ilSpotLen \ 30
        If ilSpotUnits <= 0 Then
            ilSpotUnits = 1
        End If
        ilSpotLen = 0
    Else
        ilSpotUnits = 1
    End If
    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
'Switch to Key2 in Dev
    tmSsfSrchKey.iType = imGameNo 'slType-On Air
    tmSsfSrchKey.iVefCode = imVefCode
    tmSsfSrchKey.iDate(0) = ilDate0
    tmSsfSrchKey.iDate(1) = ilDate1
    tmSsfSrchKey.iStartTime(0) = 0
    tmSsfSrchKey.iStartTime(1) = 0
    ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = imGameNo) And (tmSsf.iVefCode = imVefCode) And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1)
        ilRet = btrGetPosition(hmSsf, lmSsfRecPos)
        For ilEvt = 1 To tmSsf.iCount Step 1
           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                If ((llTime >= lmLibStartTime(0)) And (llTime <= lmLibEndTime(0))) Or ((llTime >= lmLibStartTime(1)) And (llTime <= lmLibEndTime(1))) Then
                    If ilAvailIndex = -1 Then
                        ilAvailIndex = ilEvt
                    End If
                    If llTime > llAvailTime Then
                        Do
                            ilRet = gGetByKeyForUpdateSSF(hmSsf, tmSsf)
                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
                            ilAvailUnits = (tmAvail.iAvInfo And &H1F) + ilSpotUnits
                            ilAvailSec = tmAvail.iLen + ilSpotLen
                            tmAvail.iAvInfo = (tmAvail.iAvInfo And (Not &H1F)) + ilAvailUnits
                            tmAvail.iLen = ilAvailSec
                            tmSsf.tPas(ADJSSFPASBZ + ilAvailIndex) = tmAvail
                            imSsfRecLen = igSSFBaseLen + tmSsf.iCount * Len(tmAvail)
                            ilRet = gSSFUpdate(hmSsf, tmSsf, imSsfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        mAddSpotFindAvailIndex = ilAvailIndex
                        Exit Function
                    Else
                        ilAvailIndex = ilEvt
                    End If
                Else
                    'If (llTime > lmLibEndTime(0)) Then
                    '    Exit Do
                    'End If
                End If
            End If
        Next ilEvt
        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
        ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If ilAvailIndex > 0 Then
        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
        ilRet = btrGetDirect(hmSsf, tmSsf, imSsfRecLen, lmSsfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        Do
            ilRet = gGetByKeyForUpdateSSF(hmSsf, tmSsf)
           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
            ilAvailUnits = (tmAvail.iAvInfo And &H1F) + ilSpotUnits
            ilAvailSec = tmAvail.iLen + ilSpotLen
            tmAvail.iAvInfo = (tmAvail.iAvInfo And (Not &H1F)) + ilAvailUnits
            tmAvail.iLen = ilAvailSec
            tmSsf.tPas(ADJSSFPASBZ + ilAvailIndex) = tmAvail
            imSsfRecLen = igSSFBaseLen + tmSsf.iCount * Len(tmAvail)
            ilRet = gSSFUpdate(hmSsf, tmSsf, imSsfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        mAddSpotFindAvailIndex = ilAvailIndex
        Exit Function
    End If

End Function
Private Sub mUpdateCopy(grdGrid As MSHFlexGrid)
    Dim slCopy As String
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String

    slCopy = grdGrid.TextMatrix(lmEnableRow, lmEnableCol)
    Do
        If imWhichGrid = 0 Then
            tmSdfSrchKey3.lCode = Val(grdLog.TextMatrix(lmEnableRow, SDFCODEINDEX))
        Else
            tmSdfSrchKey3.lCode = Val(grdLogged.TextMatrix(lmEnableRow, SDFCODEINDEX))
        End If
        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
        If ilRet = BTRV_ERR_NONE Then
            If (slCopy <> "") And (slCopy <> "[None]") Then
                mCopyPop
                gFindMatch slCopy, 1, lbcCopy
                If gLastFound(lbcCopy) >= 1 Then
                    slNameCode = tmCopyCode(gLastFound(lbcCopy) - 1).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tmSdf.sPtType = "1"
                    tmSdf.lCopyCode = Val(slCode)
                Else
                    tmSdf.sPtType = "0"
                    tmSdf.lCopyCode = 0        'Copy inventory code
                    tmSdf.iRotNo = 0
                End If
            Else
                tmSdf.sPtType = "0"
                tmSdf.lCopyCode = 0        'Copy inventory code
                tmSdf.iRotNo = 0
            End If
            tmSdf.sAffChg = "Y"
            tmSdf.iUrfCode = tgUrf(0).iCode
            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
End Sub

Private Function mGetCntrFromRot(ilLen As Integer) As Long
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim slType As String
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim llRecPos As Long
    Dim slDate As String
    Dim ilCrf As Integer
    Dim ilLineNo As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record


    mGetCntrFromRot = 0
    ReDim tmOldCrf(0 To 0) As CRF
    btrExtClear hmCrf   'Clear any previous extend operation
    ilExtLen = Len(tmCrf)  'Extract operation record size
    slType = "A"
    tmCrfSrchKey1.sRotType = slType
    tmCrfSrchKey1.iEtfCode = 0
    tmCrfSrchKey1.iEnfCode = 0
    tmCrfSrchKey1.iAdfCode = igAddAvtCode
    tmCrfSrchKey1.lChfCode = 0
    tmCrfSrchKey1.lFsfCode = 0
    tmCrfSrchKey1.iVefCode = 0
    tmCrfSrchKey1.iRotNo = 32000
    ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (tmCrf.iAdfCode = igAddAvtCode) And (ilRet <> BTRV_ERR_END_OF_FILE) Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmCrf, llNoRec, -1, "UC", "CRF", "") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Crf", "CrfAdfCode")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, igAddAvtCode, 2)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
        tlCharTypeBuff.sType = slType
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
        ilOffSet = 0
        ilRet = btrExtAddField(hmCrf, ilOffSet, ilExtLen)  'Extract start/end time, and days
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
        'ilRet = btrExtGetNextExt(hmClf)    'Extract record
        ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Function
            End If
            'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
            If ilRet = BTRV_ERR_REJECT_COUNT Then
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
            End If
            Do While ilRet = BTRV_ERR_NONE
                gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
                ''If (gDateValue(slDate) > lmNowDate) Or (rbcCRShow(1).Value) Then
                'If (gDateValue(slDate) > lmNowDate - 30) Then
                If (gDateValue(slDate) > lmEventDate - 30) Then
                    tmCnfSrchKey.lCrfCode = tmCrf.lCode
                    tmCnfSrchKey.iInstrNo = 0
                    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                        If tmCnf.lCifCode = lgAddCopyCode Then
                            ilLineNo = mChkCopyCntr(tmCrf.lChfCode, ilLen)
                            If ilLineNo > 0 Then
                                mGetCntrFromRot = tmCrf.lChfCode
                                btrExtClear hmCrf   'Clear any previous extend operation
                                Exit Function
                            Else
                                Exit Do
                            End If
                        End If
                        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                Else
                    tmOldCrf(UBound(tmOldCrf)) = tmCrf
                    ReDim Preserve tmOldCrf(0 To UBound(tmOldCrf) + 1) As CRF
                End If
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                End If
            Loop
            For ilCrf = 0 To UBound(tmOldCrf) - 1 Step 1
                tmCrf = tmOldCrf(ilCrf)
                tmCnfSrchKey.lCrfCode = tmCrf.lCode
                tmCnfSrchKey.iInstrNo = 0
                ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                    If tmCnf.lCifCode = lgAddCopyCode Then
                        ilLineNo = mChkCopyCntr(tmCrf.lChfCode, ilLen)
                        If ilLineNo > 0 Then
                            mGetCntrFromRot = tmCrf.lChfCode
                            btrExtClear hmCrf   'Clear any previous extend operation
                            Exit Function
                        Else
                            Exit Do
                        End If
                    End If
                    ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            Next ilCrf
        End If
        btrExtClear hmCrf   'Clear any previous extend operation
    End If
End Function

Private Function mChkCopyCntr(llChfCode As Long, ilLen As Integer) As Integer
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim ilVef As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilIncludeHistory As Integer

    mChkCopyCntr = 0
    ilIncludeHistory = False
    ilRet = gObtainChfClf(hmCHF, hmClf, llChfCode, ilIncludeHistory, tmChf, tmAddClf())
    If ilRet Then
        For ilClf = LBound(tmAddClf) To UBound(tmAddClf) - 1 Step 1
            If (tmAddClf(ilClf).ClfRec.sType <> "O") And (tmAddClf(ilClf).ClfRec.sType <> "A") And (tmAddClf(ilClf).ClfRec.sType <> "E") Then
                ilVef = gBinarySearchVef(tmAddClf(ilClf).ClfRec.iVefCode)
                If ilVef <> -1 Then
                    If (tgMVef(ilVef).sType <> "R") And (tgMVef(ilVef).sType <> "P") Then
                        gUnpackDateLong tmAddClf(ilClf).ClfRec.iStartDate(0), tmAddClf(ilClf).ClfRec.iStartDate(1), llStartDate
                        gUnpackDateLong tmAddClf(ilClf).ClfRec.iEndDate(0), tmAddClf(ilClf).ClfRec.iEndDate(1), llEndDate
                        If (llStartDate <= llEndDate) And (ilLen = tmAddClf(ilClf).ClfRec.iLen) Then
                            tmClf = tmAddClf(ilClf).ClfRec
                            mChkCopyCntr = tmAddClf(ilClf).ClfRec.iLine
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next ilClf
    End If

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeUnschSpot                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Create a Sdf records            *
'*                                                     *
'*                     Similar to code within          *
'*                     CntSchd.Bas                     *
'*                                                     *
'*******************************************************
Private Function mMakeUnschSpot(llChfCode As Long, ilLineNo As Integer, llFsfCode As Long, ilVefCode As Integer) As Integer
'
'   ilRet = mMakeUnschSpot(llChfCode, ilLineNo, lMisseddate, slPriceType, llSdfRecPos)
'   Where:
'       llChfCode(I)- Chf Code
'       ilLineNo(I)- Line number
'       slMissedDate(I)- Date to create spot for
'       ilExtraSpot(I)- True=Extra Bonus Spot
'       llSdfRecPos(O)- Sdf Record position
'
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilRet As Integer
    Dim slMissedDate As String

    tmSdf.lCode = 0
    tmSdf.iVefCode = tmClf.iVefCode
    tmSdf.lChfCode = llChfCode    'Contract code
    tmSdf.iLineNo = ilLineNo    'Line number
    tmSdf.lFsfCode = llFsfCode
    tmSdf.iAdfCode = tmChf.iAdfCode 'Advertiser code number
    'slMissedDate = Format$(lmNowDate, "m/d/yy")
    slMissedDate = Format$(lmEventDate, "m/d/yy")
    gPackDate slMissedDate, tmSdf.iDate(0), tmSdf.iDate(1)
    llDate = gDateValue(slMissedDate)
    ilDay = gWeekDayLong(llDate)
    'If (tmLnRdf.iLtfCode(0) <> 0) Or (tmLnRdf.iLtfCode(1) <> 0) Or (tmLnRdf.iLtfCode(2) <> 0) Then
    '    tmSdf.iTime(0) = 0
    '    tmSdf.iTime(1) = 0
    'Else    'Time buy- check if override times defined (if so, use them as bump times)
        If (tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0) Then
            tmSdf.iTime(0) = 0
            tmSdf.iTime(1) = 0
        Else
            tmSdf.iTime(0) = tmClf.iStartTime(0)
            tmSdf.iTime(1) = tmClf.iStartTime(1)
        End If
    'End If
    tmSdf.sSchStatus = "M"    'S=Scheduled, M=Missed,
                                'G=Makegood, A=on alternate log but not MG, B=on alternate Log and MG,
                                'C=Cancelled
    tmSdf.iMnfMissed = igDefaultMnfMissed   'Missed reason
    tmSdf.sTracer = " "   'M=Mouse move, N=On demand & mouse moved, C=Created in post log,
                            'N=N/A, D=on Demand & created in post log
    tmSdf.sAffChg = " "   'T=Time change, C=Copy change, B=Time and copy changed, blank=no change
    If (llFsfCode <= 0) Or (tmFsf.lCifCode <= 0) Then
        tmSdf.sPtType = "0"
        tmSdf.lCopyCode = 0        'Copy inventory code
        tmSdf.iRotNo = 0
    Else
        tmSdf.sPtType = "1"
        tmSdf.lCopyCode = tmFsf.lCifCode
    End If
    tmSdf.iRotNo = 0
    tmSdf.iLen = tmClf.iLen         'Spot length
    tmSdf.sPriceType = "+"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
    tmSdf.sSpotType = "X"   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo; X=Extra Spot
    tmSdf.sBill = "N"
    tmSdf.lSmfCode = 0
    tmSdf.iGameNo = imGameNo
    tmSdf.iUrfCode = tgUrf(0).iCode      'Last user who modified spot
    tmSdf.sXCrossMidnight = "N"
    ilRet = btrInsert(hmSdf, tmSdf, imSdfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        mMakeUnschSpot = False
        Exit Function
    End If
    mMakeUnschSpot = True
    Exit Function
End Function


Private Function mExitOk() As Integer
    Dim ilRet As Integer
    Dim slTime As String
    Dim llRow As Long
    Dim llCol As Long
    Dim ilAnyLogged As Integer
    Dim ilFeedChgd As Integer
    Dim slNowDate As String
    Dim ilVefCode As Integer
    Dim slEventTitle1 As String
    Dim slEventTitle2 As String

    mExitOk = True
    ilAnyLogged = True
    If (grdLogged.TextMatrix(grdLogged.FixedRows, TYPEINDEX) = "") Or (imSignOn = 2) Then
        ilAnyLogged = False
    Else
        For llRow = grdLog.FixedRows To grdLog.Rows - 1 Step 1
            If (grdLog.TextMatrix(llRow, TYPEINDEX) = "Prg Start") Or (grdLog.TextMatrix(llRow, TYPEINDEX) = "Event Start") Then
                ilAnyLogged = False
                Exit For
            End If
        Next llRow
    End If
    If ilAnyLogged Then
        Screen.MousePointer = vbHourglass
        'Check if day logged, if not ask if Ok to exit
        If grdLog.TextMatrix(grdLog.FixedRows, TYPEINDEX) <> "" Then
            For llRow = grdLog.FixedRows To grdLog.Rows - 1 Step 1
                If (grdLog.TextMatrix(llRow, TYPEINDEX) = "Prg Over") Or (grdLog.TextMatrix(llRow, TYPEINDEX) = "Event Over") Then
                    Screen.MousePointer = vbDefault
                    If imSSFType > 0 Then
                        ilRet = MsgBox("Is Event Over", vbYesNo + vbQuestion, "Status")
                    Else
                        ilRet = MsgBox("Is Program Over", vbYesNo + vbQuestion, "Status")
                    End If
                    If ilRet = vbNo Then
                        ilRet = MsgBox("Is Someone Else Taking Over Log Duties", vbYesNo + vbQuestion, "Task")
                        If ilRet = vbYes Then
                            'Do
                            '    tmLDFSrchKey1.lLlfCode = tmLlf.lCode
                            '    ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                            '    Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlf.lCode)
                            '        If (tmLdf.sType = "S") Then
                            '            slTime = Format$(gNow(), "h:mm:ssAM/PM")
                            '            gPackTime slTime, tmLdf.iEndTime(0), tmLdf.iEndTime(1)
                            '            ilRet = btrUpdate(hmLdf, tmLdf, imLdfRecLen)
                            '            Exit Do
                            '        End If
                            '        ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                            '    Loop
                            'Loop While ilRet = BTRV_ERR_CONFLICT
                            tmLdf.lCode = 0
                            tmLdf.sType = "S"
                            tmLdf.sSubType = "E"
                            slTime = Format$(gNow(), "h:mm:ssAM/PM")
                            gPackTime slTime, tmLdf.iTime(0), tmLdf.iTime(1)
                            tmLdf.lLlfCode = tmLlf.lCode
                            tmLdf.iUrfCode = tgUrf(0).iCode
                            tmLdf.sCurrentFeed = smCurrentFeed
                            slNowDate = gNow()
                            tmLdf.lEnteredDate = gDateValue(Format$(slNowDate, "m/d/yy"))
                            tmLdf.lEnteredTime = gTimeToLong(Format$(slNowDate, "h:mm:ssAM/PM"), False)
                            tmLdf.sUnused = ""
                            ilRet = btrInsert(hmLdf, tmLdf, imLdfRecLen, INDEXKEY0)
                            Do
                                tmLlfSrchKey.lCode = tmLlf.lCode
                                ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                                If ilRet = BTRV_ERR_NONE Then
                                    'slTime = Format$(gNow(), "h:mm:ssAM/PM")
                                    'gPackTime slTime, tmLlf.iSignOffTime(0), tmLlf.iSignOffTime(1)
                                    tmLlf.sSignOffStatus = "Y"
                                    ilRet = btrUpdate(hmLlf, tmLlf, imLlfRecLen)
                                Else
                                    Exit Do
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            Exit Function
                        Else
                            If imSSFType > 0 Then
                                MsgBox "End Event and Resolve All Spots Prior to Sign-Off", vbOKOnly + vbExclamation, "Sign-Off"
                            Else
                                MsgBox "End Program and Resolve All Spots Prior to Sign-Off", vbOKOnly + vbExclamation, "Sign-Off"
                            End If
                            mExitOk = False
                            Exit Function
                        End If
                    Else
                        grdLog.Row = llRow
                        grdLog.Col = AIREDINDEX
                        mEnableBox
                        mExitOk = False
                        Exit Function
                    End If
                End If
            Next llRow
            Screen.MousePointer = vbDefault
            MsgBox "Resolve All Spots Prior to Sign-Off", vbOKOnly + vbExclamation, "Sign-Off"
            mExitOk = False
            Exit Function
        Else
            'Do
            '    tmLDFSrchKey1.lLlfCode = tmLlf.lCode
            '    ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            '    Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlf.lCode)
            '        If (tmLdf.sType = "S") Then
            '            slTime = Format$(gNow(), "h:mm:ssAM/PM")
            '            gPackTime slTime, tmLdf.iEndTime(0), tmLdf.iEndTime(1)
            '            ilRet = btrUpdate(hmLdf, tmLdf, imLdfRecLen)
            '            Exit Do
            '        End If
            '        ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
            '    Loop
            'Loop While ilRet = BTRV_ERR_CONFLICT
            tmLdf.lCode = 0
            tmLdf.sType = "S"
            tmLdf.sSubType = "E"
            slTime = Format$(gNow(), "h:mm:ssAM/PM")
            gPackTime slTime, tmLdf.iTime(0), tmLdf.iTime(1)
            tmLdf.lLlfCode = tmLlf.lCode
            tmLdf.iUrfCode = tgUrf(0).iCode
            tmLdf.sCurrentFeed = smCurrentFeed
            slNowDate = gNow()
            tmLdf.lEnteredDate = gDateValue(Format$(slNowDate, "m/d/yy"))
            tmLdf.lEnteredTime = gTimeToLong(Format$(slNowDate, "h:mm:ssAM/PM"), False)
            tmLdf.sUnused = ""
            ilRet = btrInsert(hmLdf, tmLdf, imLdfRecLen, INDEXKEY0)
            Do
                tmLlfSrchKey.lCode = tmLlf.lCode
                ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                If ilRet = BTRV_ERR_NONE Then
                    'slTime = Format$(gNow(), "h:mm:ssAM/PM")
                    'gPackTime slTime, tmLlf.iSignOffTime(0), tmLlf.iSignOffTime(1)
                    tmLlf.sSignOffStatus = "Y"
                    tmLlf.sLogCompleted = "Y"
                    ilRet = btrUpdate(hmLlf, tmLlf, imLlfRecLen)

                Else
                    Exit Do
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet = BTRV_ERR_NONE Then
                'Update complete flag in LCF
                ilRet = mUpdateLCFAffPost(tmLlf, "C")
            End If
            llRow = grdLogged.Rows
            If llRow >= grdLogged.Rows Then
                grdLogged.AddItem ""
            End If
            grdLogged.Row = llRow
            grdLogged.Col = SCHDTIMEINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = LENGTHINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = TYPEINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            grdLogged.Col = ADVTPRODINDEX
            grdLogged.CellBackColor = LIGHTYELLOW
            For llCol = 0 To grdLogged.Cols - 1 Step 1
                grdLogged.Col = llCol
                grdLogged.CellAlignment = flexAlignLeftTop
            Next llCol
            grdLogged.TextMatrix(llRow, TYPEINDEX) = "Sign-Off"
            grdLogged.TextMatrix(llRow, AIREDINDEX) = slTime
            grdLogged.TextMatrix(llRow, ADVTPRODINDEX) = sgLLName
            grdLogged.Row = llRow
            For llCol = 0 To grdLogged.Cols - 1 Step 1
                grdLogged.Col = llCol
                grdLogged.CellForeColor = DARKGREEN 'vbGreen
            Next llCol
            ilRet = mPreemptRegularProg()
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    Else
        If imSignOn <> 2 Then
            ilFeedChgd = False
            If (Asc(tgSpf.sSportInfo) And USINGFEED) = USINGFEED Then
                ilVefCode = grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, VEHICLECODEINDEX)
                gGetEventTitles ilVefCode, slEventTitle1, slEventTitle2
                If grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMEFEEDINDEX) = slEventTitle1 Then  '"Visiting" Then
                    If smCurrentFeed <> "V" Then
                        ilFeedChgd = True
                    End If
                ElseIf grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMEFEEDINDEX) = slEventTitle2 Then  '"Home" Then
                    If smCurrentFeed <> "H" Then
                        ilFeedChgd = True
                    End If
                ElseIf grdVehicleSelection.TextMatrix(grdVehicleSelection.Row, GAMEFEEDINDEX) = "National" Then
                    If smCurrentFeed <> "N" Then
                        ilFeedChgd = True
                    End If
                End If
            End If
            If Not ilFeedChgd Then
                Screen.MousePointer = vbHourglass
                If (grdLogged.TextMatrix(grdLogged.FixedRows, TYPEINDEX) <> "") Then
                    'User only signed on, then signed off
                    tmLlfSrchKey.lCode = tmLlf.lCode
                    ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrDelete(hmLlf)
                        Do
                            tmLDFSrchKey1.lLlfCode = tmLlf.lCode
                            ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                            If (ilRet = BTRV_ERR_NONE) Then
                                ilRet = btrDelete(hmLdf)
                            Else
                                Exit Do
                            End If
                        Loop
                    End If
                End If
            End If
        End If
        imSignOn = 0
        mPopVehSel
        pbcVehicleSelection.Visible = True
        csiLiveLogDate.SetEnabled True
        Form_Resize
        pbcLog.Visible = False
        mExitOk = False
        Screen.MousePointer = vbDefault
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTeamPop                        *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Tema list box         *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mLanguagePop()
'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, lbcCtrl, sType)
'
'   mdemoPop
'   Where:
'
    Dim ilRet As Integer
    ilRet = gObtainMnfForType("L", smLangTag, tmLang())
    Exit Sub
End Sub '*******************************************************
'*                                                     *
'*      Procedure Name:mTeamPop                        *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Tema list box         *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mTeamPop()
'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, lbcCtrl, sType)
'
'   mdemoPop
'   Where:
'
    Dim ilRet As Integer
    ilRet = gObtainMnfForType("Z", smTeamTag, tmTeam())
    Exit Sub
End Sub

Private Sub mGetTeamAndLang(ilVefCode As Integer, ilGameNo As Integer, slVisitTeam As String, slHomeTeam As String, slLanguage As String, slFeedSource As String, slFullName As String)
    Dim ilRet As Integer
    Dim ilTeam As Integer
    Dim ilLang As Integer
    Dim llDate As Long
    Dim slEventTitle1 As String
    Dim slEventTitle2 As String

    'Get Names
    slVisitTeam = ""
    slHomeTeam = ""
    slLanguage = ""
    slFeedSource = ""
    slFullName = ""
    tmGsfSrchKey3.iVefCode = ilVefCode
    tmGsfSrchKey3.iGameNo = ilGameNo
    ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
    Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode) And (tmGsf.iGameNo = ilGameNo)
        gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
        'If (llDate = lmNowDate) Then
        If (llDate = lmEventDate) Then
            Exit Do
        End If
        ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
    Loop
    If (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode) And (tmGsf.iGameNo = ilGameNo) Then
        'For ilTeam = 1 To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
        For ilTeam = LBound(tmTeam) To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            If tmGsf.iVisitMnfCode = tmTeam(ilTeam).iCode Then
                slFullName = Trim$(tmTeam(ilTeam).sName)
                slVisitTeam = Trim$(Left$(tmTeam(ilTeam).sName, 4))
                If Trim$(tmTeam(ilTeam).sUnitType) <> "" Then
                    slVisitTeam = Trim$(tmTeam(ilTeam).sUnitType)
                End If
                Exit For
            End If
        Next ilTeam
        'For ilTeam = 1 To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
        For ilTeam = LBound(tmTeam) To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            If tmGsf.iHomeMnfCode = tmTeam(ilTeam).iCode Then
                slFullName = slFullName & "@" & Trim$(tmTeam(ilTeam).sName)
                slHomeTeam = Trim$(Left$(tmTeam(ilTeam).sName, 4))
                If Trim$(tmTeam(ilTeam).sUnitType) <> "" Then
                    slHomeTeam = Trim$(tmTeam(ilTeam).sUnitType)
                End If
                Exit For
            End If
        Next ilTeam
        'For ilLang = 1 To UBound(tmLang) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
        For ilLang = LBound(tmLang) To UBound(tmLang) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            If tmGsf.iLangMnfCode = tmLang(ilLang).iCode Then
                slLanguage = Trim$(tmLang(ilLang).sName)
                Exit For
            End If
        Next ilLang
        gGetEventTitles ilVefCode, slEventTitle1, slEventTitle2
        If tmGsf.sFeedSource = "V" Then
            slFeedSource = slEventTitle1    '"Visiting"
        ElseIf tmGsf.sFeedSource = "H" Then
            slFeedSource = slEventTitle2    '"Home"
        ElseIf tmGsf.sFeedSource = "N" Then
            slFeedSource = "National"
        End If
    End If

End Sub

Private Function mPreemptRegularProg() As Integer
    Dim ilRet As Integer
    Dim llChfCode As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llLastLogDate As Long
    Dim ilAvailLock As Integer
    Dim ilSpotLock As Integer
    Dim ilVpfIndex As Integer
    Dim slNewStartTime As String
    Dim slNewEndTime As String
    Dim llRowNo As Long
    Dim llDate As Long

    If imGameNo > 0 Then
        tmGsfSrchKey3.iVefCode = imVefCode
        tmGsfSrchKey3.iGameNo = imGameNo
        ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = imVefCode) And (tmGsf.iGameNo = imGameNo)
            gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
            'If (llDate = lmNowDate) Then
            If (llDate = lmEventDate) Then
                Exit Do
            End If
            ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        Loop
        If (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = imVefCode) And (tmGsf.iGameNo = imGameNo) Then
            If (tmGsf.sGameStatus = "F") And (tmGsf.iAirVefCode > 0) Then
                slNewStartTime = ""
                For llRowNo = grdLogged.FixedRows To grdLogged.Rows - 1 Step 1
                    If grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Event Start" Then
                        slNewStartTime = grdLogged.TextMatrix(llRowNo, AIREDINDEX)
                        Exit For
                    End If
                Next llRowNo
                If slNewStartTime = "" Then
                    mPreemptRegularProg = False
                    Exit Function
                End If
                For llRowNo = grdLogged.FixedRows To grdLogged.Rows - 1 Step 1
                    If grdLogged.TextMatrix(llRowNo, TYPEINDEX) = "Event Over" Then
                        slNewEndTime = grdLogged.TextMatrix(llRowNo, AIREDINDEX)
                        'If (gTimeToLong(smGameEndTime, True) < gTimeToLong(slNewEndTime, True)) Or ((gTimeToLong(smGameStartTime, False) <= gTimeToLong(smGameEndTime, True)) And (gTimeToLong(smGameStartTime, False) > gTimeToLong(slNewEndTime, True))) Then
                            llChfCode = -1

                            slStartDate = Format$(smActualStartDate, "m/d/yy")   'Format$(lmNowDate, "m/d/yy")
                            slEndDate = slStartDate
                            ilVpfIndex = gBinarySearchVpf(tmGsf.iAirVefCode)
                            If ilVpfIndex >= 0 Then
                                gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), llLastLogDate
                                ilAvailLock = 1
                                ilSpotLock = 0
                                'If gTimeToLong(smGameStartTime, False) <= gTimeToLong(smGameEndTime, True) Then
                                '    If gTimeToLong(smGameStartTime, False) <= gTimeToLong(slNewEndTime, True) Then
                                '        If gTimeToLong(slNewEndTime, True) > gTimeToLong(smGameEndTime, True) Then
                                '            ilRet = gUnschSpots(tmGsf.iAirVefCode, llChfCode, llLastLogDate, slStartDate, slEndDate, smGameEndTime, slNewEndTime, -1)
                                '            gSetLockStatus tmGsf.iAirVefCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, smGameEndTime, slNewEndTime
                                '        End If
                                '    Else
                                '        If gTimeToLong(smGameEndTime, True) < gTimeToLong("12M", True) Then
                                '            ilRet = gUnschSpots(tmGsf.iAirVefCode, llChfCode, llLastLogDate, slStartDate, slEndDate, smGameEndTime, "12M", -1)
                                '            gSetLockStatus tmGsf.iAirVefCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, smGameEndTime, "12M"
                                '        End If
                                '        ilRet = gUnschSpots(tmGsf.iAirVefCode, llChfCode, llLastLogDate, gIncOneDay(slStartDate), gIncOneDay(slEndDate), "12M", slNewEndTime, -1)
                                '        gSetLockStatus tmGsf.iAirVefCode, ilAvailLock, ilSpotLock, gIncOneDay(slStartDate), gIncOneDay(slEndDate), 0, "12M", slNewEndTime
                                '    End If
                                'Else
                                '    'Game original time was across midnight.  Preempt only if game ended later then originally scheduled
                                '    If gTimeToLong(smGameStartTime, False) > gTimeToLong(slNewEndTime, True) Then
                                '        If (gTimeToLong(slNewEndTime, True) > gTimeToLong(smGameEndTime, True)) Then
                                '            ilRet = gUnschSpots(tmGsf.iAirVefCode, llChfCode, llLastLogDate, gIncOneDay(slStartDate), gIncOneDay(slEndDate), smGameEndTime, slNewEndTime, -1)
                                '            gSetLockStatus tmGsf.iAirVefCode, ilAvailLock, ilSpotLock, gIncOneDay(slStartDate), gIncOneDay(slEndDate), 0, smGameEndTime, slNewEndTime
                                '        End If
                                '    End If
                                'End If
                                If gTimeToLong(slNewStartTime, False) <= gTimeToLong(slNewEndTime, True) Then
                                    ilRet = gUnschSpots(tmGsf.iAirVefCode, llChfCode, llLastLogDate, slStartDate, slEndDate, slNewStartTime, slNewEndTime, -1)
                                    gSetLockStatus tmGsf.iAirVefCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, slNewStartTime, slNewEndTime
                                ElseIf gTimeToLong(slNewStartTime, False) <= gTimeToLong(slNewEndTime, True) + 60 Then
                                    ilRet = gUnschSpots(tmGsf.iAirVefCode, llChfCode, llLastLogDate, slStartDate, slEndDate, slNewStartTime, slNewStartTime, -1)
                                    gSetLockStatus tmGsf.iAirVefCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, slNewStartTime, slNewStartTime
                                Else
                                    'Across midnight
                                    ilRet = gUnschSpots(tmGsf.iAirVefCode, llChfCode, llLastLogDate, slStartDate, slEndDate, slNewStartTime, "12M", -1)
                                    gSetLockStatus tmGsf.iAirVefCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, slNewStartTime, "12M"
                                    ilRet = gUnschSpots(tmGsf.iAirVefCode, llChfCode, llLastLogDate, gIncOneDay(slStartDate), gIncOneDay(slEndDate), "12M", slNewEndTime, -1)
                                    gSetLockStatus tmGsf.iAirVefCode, ilAvailLock, ilSpotLock, gIncOneDay(slStartDate), gIncOneDay(slEndDate), 0, "12M", slNewEndTime
                                End If
                                smGameEndTime = slNewEndTime
                            End If
                        'End If
                        Exit For
                    End If
                Next llRowNo
            End If
        End If
    End If
    mPreemptRegularProg = True
End Function

Private Function mGetGameStatus(ilVefCode As Integer, ilGameNo As Integer) As String
    Dim ilRet As Integer
    Dim llDate As Long

    mGetGameStatus = ""
    If ilGameNo > 0 Then
        tmGsfSrchKey3.iVefCode = ilVefCode
        tmGsfSrchKey3.iGameNo = ilGameNo
        ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode) And (tmGsf.iGameNo = ilGameNo)
            gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
            'If (llDate = lmNowDate) Then
            If (llDate = lmEventDate) Then
                Exit Do
            End If
            ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        Loop
        If (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode) And (tmGsf.iGameNo = ilGameNo) Then
            mGetGameStatus = tmGsf.sGameStatus
            '9/21/09: Set as if cancelled to aviod showing game
            If ((tgVpf(imVpfIndex).sGenLog = "A") And (tmGsf.sLiveLogMerge = "M")) Then
                mGetGameStatus = "C"
            End If
        End If
    End If
End Function

Private Sub tmcDelay_Timer()
    Dim slEventDate As String
    Dim slDate As String
    tmcDelay.Enabled = False
    On Error Resume Next
    slEventDate = csiLiveLogDate.Text
    If slEventDate = "" Then
        mClearVehSelGrid
        csiLiveLogDate.SetFocus
        Exit Sub
    End If
    If Not gIsDate(slEventDate) Then
        mClearVehSelGrid
        csiLiveLogDate.SetFocus
        Exit Sub
    End If
    smEventDate = slEventDate
    lmEventDate = gDateValue(slEventDate)
    If lmLastStdMnth <= 0 Then
        If (lmEventDate > lmNowDate) Then
            MsgBox "Date must be on or prior to " & smNowDate, vbOKOnly + vbInformation, "Date"
            mClearVehSelGrid
            slDate = Format$(gNow(), "m/d/yy")   'Get year
            csiLiveLogDate.Text = slDate
            csiLiveLogDate.SetFocus
            Exit Sub
        End If
    ElseIf (lmEventDate <= lmLastStdMnth) Or (lmEventDate > lmNowDate) Then
        Beep
        MsgBox "Date must be after " & smLastStdMnth & " and on or prior to " & smNowDate, vbOKOnly + vbInformation, "Date"
        mClearVehSelGrid
        slDate = Format$(gNow(), "m/d/yy")   'Get year
        csiLiveLogDate.Text = slDate
        csiLiveLogDate.SetFocus
        Exit Sub
    End If
    mPopVehSel
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload LiveLog
End Sub

Private Function mFindLdf(slType As String, slSubType As String) As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRet As Integer
    Dim tlLlf As LLF

    mFindLdf = False
    'gPackDateLong lmNowDate, ilDate0, ilDate1
    gPackDateLong lmEventDate, ilDate0, ilDate1
    tmLLFSrchKey1.iAirDate(0) = ilDate0
    tmLLFSrchKey1.iAirDate(1) = ilDate1
    tmLLFSrchKey1.iVefCode = imVefCode
    ilRet = btrGetEqual(hmLlf, tlLlf, imLlfRecLen, tmLLFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
    Do While (ilRet = BTRV_ERR_NONE) And (tlLlf.iVefCode = imVefCode) And (tlLlf.iAirDate(0) = ilDate0) And (tlLlf.iAirDate(1) = ilDate1)
        If (tlLlf.iGameNo = imGameNo) Then
            If (Asc(tgSpf.sSportInfo) And USINGFEED) = USINGFEED Then
                tmLDFSrchKey1.lLlfCode = tlLlf.lCode
                ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tlLlf.lCode)
                    If (tmLdf.sType = slType) And (tmLdf.sSubType = slSubType) Then
                        mFindLdf = True
                        Exit Function
                    End If
                    ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                Loop
            End If
        End If
        ilRet = btrGetNext(hmLlf, tlLlf, imLlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Function


Private Function mUpdateLCFAffPost(tlLlf As LLF, slAffPost As String) As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRet As Integer

    Do
        'gPackDateLong lmNowDate, ilDate0, ilDate1
        gPackDateLong lmEventDate, ilDate0, ilDate1
        tmLcfSrchKey.iType = tlLlf.iGameNo
        tmLcfSrchKey.sStatus = "C"
        tmLcfSrchKey.iVefCode = tlLlf.iVefCode
        tmLcfSrchKey.iLogDate(0) = ilDate0
        tmLcfSrchKey.iLogDate(1) = ilDate1
        tmLcfSrchKey.iSeqNo = 1
        ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet = BTRV_ERR_NONE Then
            tmLcf.sAffPost = slAffPost  'C or I
            ilRet = btrUpdate(hmLcf, tmLcf, imLcfRecLen)
        Else
            Exit Do
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    mUpdateLCFAffPost = ilRet
End Function

Private Function mUndoAllowed(grdGrid As MSHFlexGrid, llInMoveRow As Long) As Integer
    Dim llInRow As Long
    Dim llRow As Long

    llInRow = llInMoveRow
    mUndoAllowed = True
    If imWhichGrid = 1 Then
        If (grdGrid.TextMatrix(llInRow, TYPEINDEX) = "Prg Start") Or (grdGrid.TextMatrix(llInRow, TYPEINDEX) = "Event Start") Then
            For llRow = grdGrid.FixedRows To grdGrid.Rows - 1 Step 1
                If grdGrid.TextMatrix(llRow, TYPEINDEX) <> "" Then
                    If (grdGrid.TextMatrix(llRow, TYPEINDEX) <> "Prg Start") And (grdGrid.TextMatrix(llRow, TYPEINDEX) <> "Event Start") Then
                        If (grdGrid.TextMatrix(llRow, TYPEINDEX) <> "Sign-On") Then
                            mUndoAllowed = False
                            Exit Function
                        End If
                    End If
                End If
            Next llRow
        End If
    End If
End Function


Private Function mReplaced() As Integer
    Dim ilRet As Integer
    Dim tlLlf As LLF
    Dim slStr As String

    mReplaced = False
    tmLlfSrchKey.lCode = tmLlf.lCode
    ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
    If ilRet = BTRV_ERR_NONE Then
        If tmLlf.sReplaceStatus = "Y" Then
            mReplaced = True
            tmLlfSrchKey.lCode = tmLlf.lCode
            ilRet = btrGetEqual(hmLlf, tlLlf, imLlfRecLen, tmLlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
            If ilRet = BTRV_ERR_NONE Then
                slStr = "You have been replaced by " & Trim$(tlLlf.sName)
                If Trim$(tlLlf.sStudioID) <> "" Then
                    slStr = slStr & " in Studio " & Trim$(tlLlf.sStudioID)
                End If
                MsgBox slStr, vbOKOnly + vbInformation, "Replaced"
            End If
        End If
    End If
End Function
Private Function mCrossMidnight(slOrigTime As String, slAirTime As String) As String
    Dim slXMid As String
    
    slXMid = tmSdf.sXCrossMidnight
    '6/13/08: Moving a spot after 9pm to before 6a, make it XMid
    If gTimeToLong(slAirTime, False) >= gTimeToLong("6:00:00AM", False) Then
        slXMid = "N"
    Else
        If gTimeToLong(slOrigTime, False) >= gTimeToLong("12:00:00PM", False) Then
            slXMid = "Y"
        End If
    End If
    mCrossMidnight = slXMid
End Function

Private Function mMergeWithLog(ilVefCode As Integer) As Integer
    Dim ilVff As Integer
    
    mMergeWithLog = True
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If ilVefCode = tgVff(ilVff).iVefCode Then
            If tgVff(ilVff).sMergeTraffic = "S" Then
                mMergeWithLog = False
            End If
            Exit For
        End If
    Next ilVff

End Function

Private Sub mBuildVehicleArray()
    Dim ilVef As Integer
    Dim slStartDate As String
    Dim ilLoop As Integer
    Dim blFound As Boolean
    Dim ilRow As Integer
    Dim ilVpf As Integer
    
    ReDim imLLVefCode(0 To 0) As Integer
    ReDim imAirVefCode(0 To 0) As Integer
    ilVef = gBinarySearchVef(imVefCode)
    If ilVef <> -1 Then
        smVefType = tgMVef(ilVef).sType
        If smVefType <> "L" Then
            ReDim imLLVefCode(0 To 1) As Integer
            imLLVefCode(0) = imVefCode
        Else
            ilVpf = gBinarySearchVpf(imVefCode)
            If ilVpf <> -1 Then
                If tgVpf(ilVpf).sGMedium <> "P" Then
                    'slStartDate = Format(lmNowDate, "m/d/yy")
                    slStartDate = Format(lmEventDate, "m/d/yy")
                    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G")) And (tgMVef(ilVef).iVefCode = imVefCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                            imLLVefCode(UBound(imLLVefCode)) = tgMVef(ilVef).iCode
                            ReDim Preserve imLLVefCode(0 To UBound(imLLVefCode) + 1) As Integer
                        ElseIf (tgMVef(ilVef).sType = "A") And (tgMVef(ilVef).iVefCode = imVefCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                            gBuildLinkArray hmVLF, tgMVef(ilVef), slStartDate, igSVefCode()
                            blFound = False
                            For ilLoop = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                                For ilRow = 0 To UBound(imLLVefCode) - 1 Step 1
                                    If imLLVefCode(ilRow) = igSVefCode(ilLoop) Then
                                        blFound = True
                                    End If
                                Next ilRow
                                If Not blFound Then
                                    imLLVefCode(UBound(imLLVefCode)) = igSVefCode(ilLoop)
                                    ReDim Preserve imLLVefCode(0 To UBound(imLLVefCode) + 1) As Integer
                                End If
                            Next ilLoop
                            imAirVefCode(UBound(imAirVefCode)) = tgMVef(ilVef).iCode
                            ReDim Preserve imAirVefCode(0 To UBound(imAirVefCode) + 1) As Integer
                        End If
                    Next ilVef
                End If
            End If
        End If
    Else
        smVefType = ""
    End If
End Sub

Private Function mObtainVlfStartDate(ilVefCode As Integer, llDate As Long)
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim ilTerminated As Integer
    Dim ilRet As Integer
    
    gPackDateLong llDate, ilLogDate0, ilLogDate1
    ilDay = gWeekDayLong(llDate)
    If (ilDay >= 0) And (ilDay <= 4) Then
        slDay = "0"
    ElseIf ilDay = 5 Then
        slDay = "6"
    Else
        slDay = "7"
    End If
    imVlfDate0 = 0
    imVlfDate1 = 0
    tmVlfSrchKey0.iSellCode = ilVefCode
    tmVlfSrchKey0.iSellDay = Val(slDay)
    tmVlfSrchKey0.iEffDate(0) = ilLogDate0
    tmVlfSrchKey0.iEffDate(1) = ilLogDate1
    tmVlfSrchKey0.iSellTime(0) = 0
    tmVlfSrchKey0.iSellTime(1) = 6144    '24*256
    tmVlfSrchKey0.iSellPosNo = 32000
    tmVlfSrchKey0.iSellSeq = 32000
    ilRet = btrGetLessOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = ilVefCode) 'And (tmVlf.iSellDay = Val(slDay))
        If (tmVlf.iSellDay = Val(slDay)) Then
            ilTerminated = False
            If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                    ilTerminated = True
                End If
            End If
            If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                imVlfDate0 = tmVlf.iEffDate(0)
                imVlfDate1 = tmVlf.iEffDate(1)
                Exit Do
            End If
        End If
        ilRet = btrGetPrevious(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mObtainVlfStartDate = True
End Function

Private Function mTranslateSellToAir(ilVefCode As Integer, llDate As Long, ilSpotAdded As Integer) As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim ilVef As Integer
    Dim ilRet As Integer
    Dim slTime As String
    Dim slRunTime As String
    Dim tlSdf As SDF
    
    tlSdf = tmSdf
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        If (tgMVef(ilVef).sType = "S") And ((imVlfDate0 <> 0) Or (imVlfDate1 <> 0)) Then
            gPackDateLong llDate, ilLogDate0, ilLogDate1
            ilDay = gWeekDayLong(llDate)
            If (ilDay >= 0) And (ilDay <= 4) Then
                slDay = "0"
            ElseIf ilDay = 5 Then
                slDay = "6"
            Else
                slDay = "7"
            End If
            tmVlfSrchKey0.iSellCode = ilVefCode
            tmVlfSrchKey0.iSellDay = Val(slDay)
            tmVlfSrchKey0.iEffDate(0) = imVlfDate0
            tmVlfSrchKey0.iEffDate(1) = imVlfDate1
            tmVlfSrchKey0.iSellTime(0) = tmSdf.iTime(0)
            tmVlfSrchKey0.iSellTime(1) = tmSdf.iTime(1)
            tmVlfSrchKey0.iSellPosNo = 0
            tmVlfSrchKey0.iSellSeq = 1
            ilRet = btrGetGreaterOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = ilVefCode) And (tmVlf.iSellDay = Val(slDay)) And (tmVlf.iEffDate(0) = imVlfDate0) And (tmVlf.iEffDate(1) = imVlfDate1) And (tmVlf.iSellTime(0) = tmSdf.iTime(0)) And (tmVlf.iSellTime(1) = tmSdf.iTime(1))
                For ilVef = 0 To UBound(imAirVefCode) - 1 Step 1
                    If tmVlf.iAirCode = imAirVefCode(ilVef) Then
                        'Process spot
                        tmSdf.iTime(0) = tmVlf.iAirTime(0)
                        tmSdf.iTime(1) = tmVlf.iAirTime(1)
                        'Avail time
                        gUnpackTime tmVlf.iAirTime(0), tmVlf.iAirTime(1), "A", "1", slTime
                        slRunTime = slTime
                        If (Trim$(tmSdf.sAffChg) = "") Or (Trim$(tmSdf.sAffChg) = "N") Then
                            ilSpotAdded = True
                            mAddSpotToGrid grdLog, 0, slTime, slRunTime, tmVlf.iAirPosNo
                        Else
                            mAddSpotToGrid grdLogged, 1, slTime, slRunTime, tmVlf.iAirPosNo
                        End If
                        tmSdf = tlSdf
                        Exit For
                    End If
                Next ilVef
                ilRet = btrGetNext(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    End If
End Function

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
Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slStartIn As String
    Dim slCSIName As String

    sgCommandStr = Command$
    slStartIn = CurDir$
    slCSIName = ""
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    igShowVersionNo = 0
    If (InStr(1, slStartIn, "Prod", vbTextCompare) = 0) And (InStr(1, slStartIn, "Test", vbTextCompare) = 0) Then
        igShowVersionNo = 1
        If InStr(1, sgCommandStr, "Debug", vbTextCompare) > 0 Then
            igShowVersionNo = 2
        End If
    End If
    slCommand = sgCommandStr    'Command$
    lgCurrHRes = GetDeviceCaps(Traffic!pbcList.hdc, HORZRES)
    lgCurrVRes = GetDeviceCaps(Traffic!pbcList.hdc, VERTRES)
    lgCurrBPP = GetDeviceCaps(Traffic!pbcList.hdc, BITSPIXEL)
    mTestPervasive
    '4/2/11: Add setting of value
    lgUlfCode = 0
    'If (Trim$(sgCommandStr) = "") Or (Trim$(sgCommandStr) = "/UserInput") Or (Trim$(sgCommandStr) = "Debug") Then
    If InStr(1, sgCommandStr, "^", vbTextCompare) <= 0 Then
        Signon.Show vbModal
        If igExitTraffic Then
            imTerminate = True
            Exit Sub
        End If
        slStr = sgUserName
        sgCallAppName = "Traffic"
        '5/22/18: Password check moved below
        'If Not mCheckPW() Then
        '    imTerminate = True
        '    Exit Sub
        'End If
    Else
        igSportsSystem = 0
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        ilRet = gParseItem(slCommand, 3, "\", slStr)
        igRptCallType = Val(slStr)
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
        sgUrfStamp = "~" 'Clear time stamp incase same name
        sgUserName = Trim$(slStr)
        '6/20/09:  Jim requested that the Guide sign in be changed to CSI for internal Guide only
        If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
            slCSIName = "CSI"
            sgUserName = "Guide"
        End If
        gUrfRead Signon, sgUserName, True, tgUrf(), False  'Obtain user records
        If StrComp(slCSIName, "CSI", vbTextCompare) = 0 Then
            gExpandGuideAsUser tgUrf(0)
        End If
        mGetUlfCode
    End If
    'End If
    DoEvents
'    gInitStdAlone ReportList, slStr, igTestSystem
    gInitStdAlone
    mCheckForDate
    '5/22/18: Password check moved here after csi date set so traffic and live log check the same date
    If InStr(1, sgCommandStr, "^", vbTextCompare) <= 0 Then
        If Not mCheckPW() Then
            imTerminate = True
            Exit Sub
        End If
    End If
    '4/2/11: Add setting and call.  Note: The call in _Load will be ignored
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "LiveLog.Frm"
    If tgUrf(0).sLiveLogPostOnly <> "Y" Then
        If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Or (StrComp(slCSIName, "CSI", vbTextCompare) = 0) Then
            igSportsSystem = 2
        Else
            MsgBox "You are not authorized to use Live Log", vbOKOnly + vbExclamation, "Disallowed"
            imTerminate = True
        End If
    Else
        igSportsSystem = 2
    End If
End Sub

Private Sub mTestPervasive()
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim hlSpf As Integer
    Dim tlSpf As SPF

    gInitGlobalVar
    hlSpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    ilRecLen = Len(tlSpf)
    ilRet = btrGetFirst(hlSpf, tlSpf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    btrDestroy hlSpf
End Sub
Private Sub mCheckForDate()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slDate As String
    Dim slSetDate As String
    Dim ilRet As Integer
    
    ilPos = InStr(1, sgCommandStr, "/D:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3))
        Else
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gValidDate(slDate) Then
            slDate = gAdjYear(slDate)
            slSetDate = slDate
        End If
    End If
    If Trim$(slSetDate) = "" Then
        If (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0) Or (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) > 0) Then
            slSetDate = "12/15/2015"
            slDate = slSetDate
        End If
    End If
    If Trim$(slSetDate) <> "" Then
        'Dan M 9/20/10 problems with gGetCSIName("SYSDate") in v57 reports.exe... change to global variable
     '   ilRet = csiSetName("SYSDate", slDate)
        ilRet = gCsiSetName(slDate)
    End If
    
End Sub
'4/2/11: Add routine
Private Sub mGetUlfCode()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    ilPos = InStr(1, sgCommandStr, "/ULF:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5)))
        Else
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5, ilSpace - ilPos - 3)))
        End If
    End If
End Sub

Private Sub mCheckGG()
    Dim ilRet As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llNow As Long
    Dim llDate As Long
    Dim slStr As String
    Dim ilLoop As Integer
    
    On Error Resume Next
    
    'If imLastHourGGChecked = Hour(Now) Then
    '    Exit Sub
    'End If
    'imLastHourGGChecked = Hour(Now)
    igGGFlag = 1
    hmSaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    imSafRecLen = Len(tmSaf)
    ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    ilField1 = Asc(tmSaf.sName)
    slStr = Mid$(tmSaf.sName, 2, 5)
    llDate = Val(slStr)
    ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
    llNow = gDateValue(Format$(Now, "m/d/yy"))
    If (ilField1 = 0) And (ilField2 = 1) Then
        If llDate <= llNow Then
            ilField2 = 0
        End If
    End If
    If (ilField1 = 0) And (ilField2 = 0) Then
        igGGFlag = 0
    End If
    'gSetRptGGFlag tmSaf
    btrDestroy hmSaf
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mCheckPW                        *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Check User Password             *
'*                                                     *
'*******************************************************
Public Function mCheckPW() As Integer
    Dim ilSetPassword As Integer
    Dim llPWDate As Long
    Dim slNowDate As String
    Dim llNowDate As Long
    Dim slNowTime As String
    Dim ilRet As Integer
    Dim ilUrf As Integer
    Dim slStr As String
    Dim ilNewPWRequired As Integer

    'Check user password date- if over allowed time ask for new password
    ilSetPassword = False
    gUnpackDateLong tgUrf(0).iPasswordDate(0), tgUrf(0).iPasswordDate(1), llPWDate
    slNowDate = Format$(gNow(), "m/d/yy")
    llNowDate = gDateValue(slNowDate)
    slNowTime = Format$(gNow(), "h:m:s AM/PM")
    If tgUrf(0).iCode > 2 Then
        If igSportsSystem <> 3 Then
            ilNewPWRequired = False
            If (Asc(tgSpf.sUsingFeatures2) And STRONGPASSWORD) = STRONGPASSWORD Then
                slStr = Trim$(tgUrf(0).sPassword)
                If Not gStrongPassword(slStr) Then
                    ilNewPWRequired = True
                End If
            End If
            If (llPWDate + tgSpf.iGNoDaysPass < llNowDate) Or (Trim$(tgUrf(0).sPassword) = "") Or (ilNewPWRequired) Then
                'testing code so guide can login: work around until fix guide ability to change blank password
                If (StrComp(Trim(tgUrf(0).sPassword), "", vbTextCompare)) = 0 And tgUrf(0).iCode > 2 Then
                    CSINewPW.lacCurrentPassword.Visible = False
                    CSINewPW.edcCurrentPassword.Visible = False
                End If
                tgPWUrf = tgUrf(0)
                CSINewPW.Show vbModal
                If (igExitTraffic) Then
                    mCheckPW = False
                    Exit Function
                End If
                hmUrf = CBtrvTable(TWOHANDLES)
                ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    mCheckPW = False
                    Exit Function
                End If
                imUrfRecLen = Len(tmUrf)  'btrRecordLength(hlUrf)  'Get and save record length
                tmUrfSrchKey.iCode = tgUrf(0).iCode
                ilRet = btrGetEqual(hmUrf, tmUrf, imUrfRecLen, tmUrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    mCheckPW = False
                    Exit Function
                End If
                gUrfDecrypt tmUrf
                gPackDate slNowDate, tmUrf.iPasswordDate(0), tmUrf.iPasswordDate(1)
                tmUrf.sOldPassword(2) = tmUrf.sOldPassword(1)
                tmUrf.sOldPassword(1) = tmUrf.sOldPassword(0)
                tmUrf.sOldPassword(0) = tmUrf.sPassword
                tmUrf.sPassword = sgPWResult
                'gPackDate slNowDate, tmUrf.iSignOnDate(0), tmUrf.iSignOnDate(1)
                'gPackTime slNowTime, tmUrf.iSignOnTime(0), tmUrf.iSignOnTime(1)
                gUrfEncrypt tmUrf
                ilRet = btrUpdate(hmUrf, tmUrf, imUrfRecLen)
                btrDestroy hmUrf
               
            End If
        End If
    End If
    
    mCheckPW = True
    Exit Function
End Function

Private Sub mClearVehSelGrid()
    Dim llRow As Long
    Dim ilCol As Integer
    
    grdVehicleSelection.Redraw = False
    For llRow = grdVehicleSelection.FixedRows To grdVehicleSelection.Rows - 1 Step 1
        For ilCol = 0 To grdVehicleSelection.Cols - 1 Step 1
            grdVehicleSelection.TextMatrix(llRow, ilCol) = ""
        Next ilCol
    Next llRow
    grdVehicleSelection.Rows = grdVehicleSelection.FixedRows + 1
    mVehicleSelectionColumnNames
    grdVehicleSelection.Redraw = True

End Sub
