VERSION 5.00
Begin VB.Form UnSchd 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4950
   ClientLeft      =   -120
   ClientTop       =   2160
   ClientWidth     =   9030
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
   ScaleHeight     =   4950
   ScaleWidth      =   9030
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   6750
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Unschd.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   26
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
         TabIndex        =   22
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   23
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcType 
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1950
      ScaleHeight     =   285
      ScaleWidth      =   5235
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   5295
      Begin VB.OptionButton rbcUnschType 
         Caption         =   "Vehicle Schedule/Unschedule"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   60
         Value           =   -1  'True
         Width           =   2910
      End
      Begin VB.OptionButton rbcUnschType 
         Caption         =   "Contract Balancing"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   3135
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   60
         Width           =   1995
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
      Left            =   2880
      Picture         =   "Unschd.frx":2E1A
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3420
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1575
      ScaleHeight     =   210
      ScaleWidth      =   1740
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3225
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2445
      MaxLength       =   10
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2955
      Visible         =   0   'False
      Width           =   930
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
      Left            =   1380
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2355
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
         Picture         =   "Unschd.frx":2F14
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   20
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
            Picture         =   "Unschd.frx":3BD2
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3585
      Width           =   75
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   75
      ScaleHeight     =   165
      ScaleWidth      =   135
      TabIndex        =   27
      Top             =   3360
      Width           =   135
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   14
      Top             =   225
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
      Left            =   4950
      TabIndex        =   29
      Top             =   4590
      Width           =   1155
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
      Left            =   75
      ScaleHeight     =   240
      ScaleWidth      =   3495
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3495
   End
   Begin VB.PictureBox plcSelection 
      ForeColor       =   &H00000000&
      Height          =   1830
      Left            =   255
      ScaleHeight     =   1770
      ScaleWidth      =   8565
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   8625
      Begin VB.ListBox lbcSelection 
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
         Height          =   1710
         Index           =   3
         Left            =   4425
         MultiSelect     =   2  'Extended
         TabIndex        =   36
         Top             =   30
         Visible         =   0   'False
         Width           =   4125
      End
      Begin VB.ListBox lbcSelection 
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
         Height          =   1710
         Index           =   1
         Left            =   4425
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   30
         Width           =   4125
      End
      Begin VB.ListBox lbcSelection 
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
         Height          =   1710
         Index           =   0
         Left            =   30
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   30
         Width           =   4125
      End
      Begin VB.ListBox lbcSelection 
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
         Height          =   1710
         Index           =   2
         Left            =   30
         TabIndex        =   35
         Top             =   30
         Visible         =   0   'False
         Width           =   4125
      End
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Process"
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
      Left            =   3060
      TabIndex        =   28
      Top             =   4575
      Width           =   1170
   End
   Begin VB.PictureBox pbcDates 
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
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   2835
      Picture         =   "Unschd.frx":3EDC
      ScaleHeight     =   1065
      ScaleWidth      =   3525
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2595
      Width           =   3525
   End
   Begin VB.PictureBox plcDates 
      ForeColor       =   &H00000000&
      Height          =   1185
      Left            =   2775
      ScaleHeight     =   1125
      ScaleWidth      =   3585
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2535
      Width           =   3645
   End
   Begin VB.PictureBox plcLine 
      ForeColor       =   &H00000000&
      Height          =   1905
      Left            =   270
      ScaleHeight     =   1845
      ScaleWidth      =   8535
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2490
      Visible         =   0   'False
      Width           =   8595
      Begin VB.PictureBox pbcLbcLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
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
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   30
         ScaleHeight     =   1470
         ScaleWidth      =   8100
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   75
         Width           =   8100
      End
      Begin VB.CheckBox ckcCheckAll 
         Caption         =   "Check All Vehicles"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1350
         TabIndex        =   10
         Top             =   1620
         Width           =   1920
      End
      Begin VB.TextBox edcSDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   7425
         MaxLength       =   10
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1635
         Width           =   930
      End
      Begin VB.CommandButton cmcSDate 
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
         Left            =   8355
         Picture         =   "Unschd.frx":81CE
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1635
         Width           =   195
      End
      Begin VB.CheckBox ckcAll 
         Caption         =   "All Lines"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   105
         TabIndex        =   9
         Top             =   1620
         Width           =   1170
      End
      Begin VB.ListBox lbcLine 
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
         Height          =   1500
         Left            =   15
         MultiSelect     =   2  'Extended
         TabIndex        =   8
         Top             =   60
         Width           =   8415
      End
      Begin VB.Label lacSDate 
         Appearance      =   0  'Flat
         Caption         =   "Start Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6480
         TabIndex        =   11
         Top             =   1635
         Width           =   885
      End
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3615
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4245
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3900
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcLnCode 
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
      Left            =   -60
      Sorted          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   750
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   4470
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "UnSchd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Unschd.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: UnSchd.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Unschedule information screen code
'
'   Activated from contract by pressing shift and alt key while
'   mouse down on Mass Sch button (hold button- cmcHold)
'
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmCtrls(0 To 6)  As FIELDAREA
Dim imLBCtrls As Integer
Dim tmCntrCode0() As SORTCODE
Dim smCntrCodeTag0 As String
Dim tmCntrCode1() As SORTCODE
Dim smCntrCodeTag1 As String
Dim tmVehCode() As SORTCODE
Dim smVehCodeTag As String
Dim imProcessing As Integer 'Processing operation
Dim imBoxNo As Integer   'Current Media Box
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imChfSelectIndex As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imBypassFocus As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imLbcArrowSetting As Integer 'True = Processing arrow key- retain current list box visible
                                'False= Make list box invisible
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
'Dim smSave(1 To 4) As String  'Values saved (1=Start date; 2=End date; 3=Start Time; 4=End Time)
Dim smSave(0 To 4) As String  'Values saved (1=Start date; 2=End date; 3=Start Time; 4=End Time)
'Dim imSave(1 To 2) As Integer   'Index 1:Unschedule 0=No, 1=Yes; Index 2: Schedule missed 0=No, 1=Yes
Dim imSave(0 To 2) As Integer   'Index 1:Unschedule 0=No, 1=Yes; Index 2: Schedule missed 0=No, 1=Yes
Dim lmEarliestAllowedDate As Long   'Todays date + 1
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim hmVef As Integer 'Vehicle file handle
Dim tmVef As VEF        'VEF record image
Dim tmVefSrchKey As INTKEY0    'VEF key record image
Dim imVefRecLen As Integer        'VEF record length
Dim imVehCode As Integer        'vehicle code
'Virtual Vehicle
Dim hmVsf As Integer 'Virtual Vehicle file handle
Dim tmVsf As VSF        'VSF record image
Dim tmVsfSrchKey As LONGKEY0    'VSF key record image
Dim imVsfRecLen As Integer        'VSF record length
'Advertiser
Dim hmAdf As Integer 'Advertiser file handle
Dim tmAdf As ADF        'ADF record image
Dim tmAdfSrchKey As INTKEY0    'ADF key record image
Dim imAdfRecLen As Integer        'ADF record length
'Contract header
Dim tmHdSrchKey As LONGKEY0  'CHF key record image
Dim hmCHF As Integer        'CHF Handle
Dim imHdRecLen As Integer      'CHF record length
Dim lmChfRecPos As Long
'Contract Lines
Dim tmLnSrchKey As CLFKEY0  'CLF key record image
Dim hmClf As Integer        'CLF Handle
Dim tmClf As CLF
Dim imLnRecLen As Integer      'CLF record length
Dim tmAirSrchKey As CFFKEY0 'CFF key record image
Dim hmCff As Integer        'CFF Handle
Dim imAirRecLen As Integer     'CFF record length

Dim hmCgf As Integer
Dim tmCgf As CGF        'CGF record image
Dim tmCgfSrchKey1 As CGFKEY1    'CGF key record image
Dim imCgfRecLen As Integer        'CGF record length
Dim tmCgfCff() As CFF

' Rate Card Programs/Times File
Dim hmRdf As Integer        'Rate Card Programs/Times file handle
Dim tmRdf As RDF            'RDF record image
Dim tmRdfSrchKey As INTKEY0 'RDF key record image
Dim imRdfRecLen As Integer     'RDF record length
Dim imLineCode() As Integer
Dim hmLcf As Integer
'Spot record
Dim hmSdf As Integer        'Spot Detail
Dim tmSdf As SDF
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey0 As SDFKEY0
Dim tmSdfSrchKey1 As SDFKEY1
'M for N Tracer record
Dim hmMtf As Integer        'Spot Detail
Dim smPreemptPass As String
Dim imChg As Integer    'Field changed
Dim imUpdateAllowed As Integer    'User can update records
Dim imFirstActivate As Integer
'Dim imListField(1 To 5) As Integer  'One more then number of fields to display
Dim imListField(0 To 5) As Integer  'One more then number of fields to display

Const STARTDATEINDEX = 1    'Start date control/field
Const ENDDATEINDEX = 2      'End date control/field
Const STARTTIMEINDEX = 3    'Start time control/field
Const ENDTIMEINDEX = 4      'End time control/field
Const UNSCHDINDEX = 5         'Unlock/Lock control/field
Const SCHDINDEX = 6        'Avail/Spot control/field
Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value
    If imSetAll Then
        If lbcLine.ListCount <= 0 Then
            Exit Sub
        End If
        imAllClicked = True
        llRg = CLng(lbcLine.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcLine.hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
        pbcLbcLine_Paint
    End If
End Sub
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    If rbcUnschType(0).Value Then
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
        edcDropDown.SetFocus
    Else
        edcSDate.SelStart = 0
        edcSDate.SelLength = Len(edcSDate.Text)
        edcSDate.SetFocus
    End If
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    If rbcUnschType(0).Value Then
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
        edcDropDown.SetFocus
    Else
        edcSDate.SelStart = 0
        edcSDate.SelLength = Len(edcSDate.Text)
        edcSDate.SetFocus
    End If
End Sub
Private Sub cmcCancel_Click()
    If imProcessing Then
        igUnSchdCallSource = CALLCANCELLED
        imTerminate = True
        Exit Sub
    End If
    igUnSchdCallSource = CALLCANCELLED
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
    mSetShow imBoxNo
    imBoxNo = -1
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case STARTDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case ENDDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case STARTTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case ENDTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
    End Select
    If rbcUnschType(0).Value Then
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
        edcDropDown.SetFocus
    Else
        edcSDate.SelStart = 0
        edcSDate.SelLength = Len(edcSDate.Text)
        edcSDate.SetFocus
    End If
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcSDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcSDate.SelStart = 0
    edcSDate.SelLength = Len(edcSDate.Text)
    edcSDate.SetFocus
End Sub
Private Sub cmcUpdate_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilCount As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    cmcCancel.Caption = "&Cancel"
    imProcessing = True
    If rbcUnschType(0).Value Then
        If imSave(1) = 1 Then   'Unschedule
            If mUnSchd() = False Then
                If imTerminate Then
                    imProcessing = False
                    cmcCancel_Click
                    Exit Sub
                End If
                imProcessing = False
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
        If imSave(2) = 1 Then   'Schedule
            If mSchd() = False Then
                If imTerminate Then
                    imProcessing = False
                    cmcCancel_Click
                    Exit Sub
                End If
                imProcessing = False
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    Else
        ilRet = mUnSchd()
        If imTerminate Then
            imProcessing = False
            cmcCancel_Click
            Exit Sub
        End If
    End If
    For ilLoop = LBound(tmCtrls) To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
        tmCtrls(ilLoop).sShow = ""
    Next ilLoop
    For ilCount = LBound(smSave) To UBound(smSave) Step 1
        smSave(ilCount) = ""
    Next ilCount
    For ilCount = LBound(imSave) To UBound(imSave) Step 1
        imSave(ilCount) = -1
    Next ilCount
    imProcessing = False
    pbcDates.Cls
    cmcCancel.Caption = "&Done"
End Sub
Private Sub cmcUpdate_GotFocus()
    plcCalendar.Visible = False
    mSetShow imBoxNo
    imBoxNo = -1
    gCtrlGotFocus cmcUpdate
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Select Case imBoxNo
        Case STARTDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                Exit Sub
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case ENDDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                Exit Sub
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case STARTTIMEINDEX
        Case ENDTIMEINDEX
    End Select
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case STARTDATEINDEX
        Case ENDDATEINDEX
        Case STARTTIMEINDEX
        Case ENDTIMEINDEX
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
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imBoxNo
        Case STARTDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case ENDDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case STARTTIMEINDEX
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
        Case ENDTIMEINDEX
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
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case STARTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case ENDDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case STARTTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case ENDTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNo
            Case STARTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case ENDDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case STARTTIMEINDEX
            Case ENDTIMEINDEX
        End Select
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcSDate_Change()
    Dim slStr As String
    slStr = edcSDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcSDate_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcSDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcSDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcSDate.SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcSDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcSDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcSDate.Text = slDate
            End If
        End If
        edcSDate.SelStart = 0
        edcSDate.SelLength = Len(edcSDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcSDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcSDate.Text = slDate
            End If
        End If
        edcSDate.SelStart = 0
        edcSDate.SelLength = Len(edcSDate.Text)
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
    If (igWinStatus(SPOTSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BF"
    Else
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BT"
    End If
    gShowBranner imUpdateAllowed
    'This loop is required to prevent a timing problem- if calling
    'with sg----- = "", then loss GotFocus to first control
    'without this loop
'    For ilLoop = 1 To 100 Step 1
'        DoEvents
'    Next ilLoop
'    gShowBranner
    If plcType.Visible = True Then
        lbcSelection(0).Visible = False
        lbcSelection(0).Visible = True
    Else
        lbcSelection(2).Visible = False
        lbcSelection(2).Visible = True
    End If
    Me.KeyPreview = True
    UnSchd.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
        If plcType.Visible Then
            plcType.Visible = False
            plcType.Visible = True
        End If
        If plcSelection.Visible Then
            plcSelection.Visible = False
            plcSelection.Visible = True
        End If
        If plcLine.Visible Then
            plcLine.Visible = False
            plcLine.Visible = True
        End If
        If plcDates.Visible Then
            plcDates.Visible = False
            plcDates.Visible = True
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
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer

    On Error Resume Next
    
    gGetSchParameters

    Erase tmCntrCode0
    Erase tmCntrCode1
    Erase tmVehCode
    Erase smSave
    Erase imSave
    Erase imLineCode
    Erase tgClfSpot
    Erase tgCffSpot
    Erase tmCgfCff

    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
    ilRet = btrClose(hmCgf)
    btrDestroy hmCgf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    btrExtClear hmVsf   'Clear any previous extend operation
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    
    Set UnSchd = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcLine_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
        pbcLbcLine_Paint
    End If
End Sub
Private Sub lbcLine_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub lbcLine_Scroll()
    pbcLbcLine_Paint
End Sub

Private Sub lbcSelection_Click(Index As Integer)
    Dim ilNoSelected As Integer
    Dim ilLoop As Integer
    Screen.MousePointer = vbHourglass
    If rbcUnschType(1).Value Then
        lbcLine.Clear
        pbcLbcLine_Paint
        If Index = 2 Then
            mCntrPop 2
        Else
            imChfSelectIndex = lbcSelection(3).ListIndex
            mLinePop
        End If
    Else
        If Index = 0 Then
            ilNoSelected = 0
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                If lbcSelection(0).Selected(ilLoop) Then
                    ilNoSelected = ilNoSelected + 1
                End If
            Next ilLoop
            If ilNoSelected <= 1 Then
                lbcSelection(1).Visible = True
                mCntrPop 0
            Else
                lbcSelection(1).Visible = False
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub lbcSelection_GotFocus(Index As Integer)
    plcCalendar.Visible = False
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
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
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(UnSchd, lbcSelection(2), Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(UnSchd, lbcSelection(2), tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", UnSchd
        On Error GoTo 0
    End If
    Exit Sub
mAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    If rbcUnschType(0).Value Then
        slStr = edcDropDown.Text
    Else
        slStr = edcSDate.Text
    End If
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
'*      Procedure Name:mCntrPop                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan     *
'*                                                     *
'*            Comments: Populate contracts for selected*
'*                      vehicle                        *
'*                                                     *
'*******************************************************
Private Sub mCntrPop(ilIndex As Integer)
'
'   mCntrPop ilIndex
'   Where:
'       ilIndex(I)- 0=Vehicle selection; 2= Advertiser
'
    Dim ilRet As Integer            'return status
    Dim ilCurrent As Integer
    Dim ilVehCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilAAS As Integer
    Dim ilShow As Integer
    Dim ilAdfCode As Integer
    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilState As Integer
    If ilIndex = 0 Then
        'Populate vehicle list box
        lbcSelection(1).Clear
        'lbcCntrCode(0).Clear
        ReDim tmCntrCode0(0 To 0) As SORTCODE
        If lbcSelection(0).ListIndex < 0 Then
            Exit Sub
        End If
        slNameCode = tmVehCode(lbcSelection(0).ListIndex).sKey 'lbcVehCode.List(lbcSelection(0).ListIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)

        'ilCntrType = 1
        'ilCurrent = 2   'Current; 1=All; 2=Current plus CBS
        'ilFilter = -2
        ilVehCode = Val(slCode)  'All
        'ilShowAdvt = True
        'ilShowDates = False 'True
        'ilRet = gPopCntrBox(UnSchd, ilCntrType, ilFilter, ilCurrent, 0, ilVehCode, lbcSelection(1), lbcCntrCode(0), ilShowAdvt, ilShowDates, False, False)
        ilAAS = 3
        slCntrStatus = "HO" 'Order; Hold
        If tgSpf.sSchdRemnant = "Y" Then
            slCntrType = "CT"
        Else
            slCntrType = "C" 'Standard only
        End If
        If tgSpf.sSchdPromo = "Y" Then
            slCntrType = slCntrType & "M"
        End If
        If tgSpf.sSchdPSA = "Y" Then
            slCntrType = slCntrType & "S"
        End If
        ilCurrent = 2
        ilShow = 3
        ilState = 1
        'ilRet = gPopCntrForAASBox(UnSchd, ilAAS, ilVehCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcSelection(1), lbcCntrCode(0))
        ilRet = gPopCntrForAASBox(UnSchd, ilAAS, ilVehCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcSelection(1), tmCntrCode0(), smCntrCodeTag0)
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mCntrPopErr
            gCPErrorMsg ilRet, "mCntrPop (gPopCntrForAASBox)", UnSchd
            On Error GoTo 0
        End If
        ilRet = mGetMissedMG(ilVehCode)
    Else
        lbcSelection(3).Clear
        'lbcCntrCode(1).Clear
        ReDim tmCntrCode1(0 To 0) As SORTCODE
        If lbcSelection(2).ListIndex < 0 Then
            Exit Sub
        End If
        slNameCode = tgAdvertiser(lbcSelection(2).ListIndex).sKey  'Traffic!lbcAdvertiser.List(lbcSelection(2).ListIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)

        'ilCntrType = 1
        'ilCurrent = 2   'Current; 1=All; 2=Current plus CBS
        'ilFilter = Val(slCode)
        'ilVehCode = -1  'All
        'ilShowAdvt = False
        'ilShowDates = True 'True
        'ilRet = gPopCntrBox(UnSchd, ilCntrType, ilFilter, ilCurrent, 0, ilVehCode, lbcSelection(3), lbcCntrCode(1), ilShowAdvt, ilShowDates, False, False)
        slCntrStatus = "HO" 'Order; Hold
        If tgSpf.sSchdRemnant = "Y" Then
            slCntrType = "CT"
        Else
            slCntrType = "C" 'Standard only
        End If
        If tgSpf.sSchdPromo = "Y" Then
            slCntrType = slCntrType & "M"
        End If
        If tgSpf.sSchdPSA = "Y" Then
            slCntrType = slCntrType & "S"
        End If
        ilCurrent = 2
        ilShow = 0
        ilState = 1
        ilAdfCode = Val(slCode)
        'ilRet = gPopCntrForAASBox(UnSchd, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcSelection(3), lbcCntrCode(1))
        ilRet = gPopCntrForAASBox(UnSchd, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcSelection(3), tmCntrCode1(), smCntrCodeTag1)
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mCntrPopErr
            gCPErrorMsg ilRet, "mCntrPop (gPopCntrBox)", UnSchd
            On Error GoTo 0
        End If
        'For ilLoop = 0 To lbcCntrCode.ListCount - 1 Step 1
        '    slNameCode = lbcCntrCode.List(ilLoop)
        '    ilRet = gParseItem(slNameCode, 1, "\", slName)
        '    If ilRet <> CP_MSG_NONE Then
        '        Exit Sub
        '    End If
        '    ilRet = gParseItem(slName, 2, "|", slAdvtName)
        '    If ilRet <> CP_MSG_NONE Then
        '        Exit Sub
        '    End If
        '    ilRet = gParseItem(slName, 1, "|", slCode)
        '    If ilRet <> CP_MSG_NONE Then
        '        Exit Sub
        '    End If
        '    llCntrNo = 99999999 - CLng(slCode)
        '    slName = Trim$(Str$(llCntrNo)) & "/" & slAdvtName
        '    lbcSelection(0).AddItem Trim$(slName)  'Add ID to list box
        'Next ilLoop
    End If
    Exit Sub
mCntrPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If (ilBoxNo < LBound(tmCtrls)) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case STARTDATEINDEX 'Start date
            edcDropDown.Width = tmCtrls(STARTDATEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcDates, edcDropDown, tmCtrls(STARTDATEINDEX).fBoxX, tmCtrls(STARTDATEINDEX).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If smSave(1) = "" Then
                'If imVehCode > 0 Then
                '    slStr = gFindVehicleLatestDate(UnSchd, imVehCode)   'Find latest date Last Log Date or Now
                '    slStr = gObtainNextMonday(slStr)
                'Else
                    slStr = gObtainMondayFromToday()
                'End If
            Else
                slStr = smSave(1)
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
            End If
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ENDDATEINDEX 'Start date
            edcDropDown.Width = tmCtrls(ENDDATEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcDates, edcDropDown, tmCtrls(ENDDATEINDEX).fBoxX, tmCtrls(ENDDATEINDEX).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top + edcDropDown.Height
            If smSave(2) <> "" Then
                slStr = smSave(2)
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
            End If
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case STARTTIMEINDEX 'Start time
            edcDropDown.Width = tmCtrls(STARTTIMEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcDates, edcDropDown, tmCtrls(STARTTIMEINDEX).fBoxX, tmCtrls(STARTTIMEINDEX).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If smSave(3) = "" Then
                edcDropDown.Text = "12M"
            Else
                edcDropDown.Text = smSave(3)
            End If
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ENDTIMEINDEX 'Start time
            edcDropDown.Width = tmCtrls(ENDTIMEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcDates, edcDropDown, tmCtrls(ENDTIMEINDEX).fBoxX, tmCtrls(ENDTIMEINDEX).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If smSave(4) = "" Then
                edcDropDown.Text = "12M"
            Else
                edcDropDown.Text = smSave(4)
            End If
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case UNSCHDINDEX
            If (imSave(1) = -1) Then
                imSave(1) = 0 'No
                tmCtrls(ilBoxNo).iChg = True
            End If
            pbcType.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcDates, pbcType, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcType_Paint
            pbcType.Visible = True
            pbcType.SetFocus
        Case SCHDINDEX
            If (imSave(2) = -1) Then
                imSave(2) = 0 'No
                tmCtrls(ilBoxNo).iChg = True
            End If
            pbcType.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcDates, pbcType, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcType_Paint
            pbcType.Visible = True
            pbcType.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetMissedMG                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add Contracts for missed MG's  *
'*                      and Bonus                      *
'*                                                     *
'*******************************************************
Private Function mGetMissedMG(ilVehCode As Integer) As Integer
    Dim slDate As String
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim llChfCode As Long
    Dim llDate As Long
    Dim ilChf As Integer
    Dim slStr As String
    If lgMtfNoRecs > 0 Then
        tmSdfSrchKey1.iVefCode = ilVehCode
        slDate = Format$(gNow(), "m/d/yy")
        llStartDate = gDateValue(slDate)
        llEndDate = llStartDate + 60
        gPackDate slDate, ilDate0, ilDate1
        tmSdfSrchKey1.iDate(0) = ilDate0
        tmSdfSrchKey1.iDate(1) = ilDate1
        tmSdfSrchKey1.iTime(0) = 0
        tmSdfSrchKey1.iTime(1) = 0
        tmSdfSrchKey1.sSchStatus = "M"
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVehCode)
            ilFound = False
            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
            If (llDate > llEndDate) Then
                Exit Do
            End If
            If (tmSdf.sSchStatus = "M") And (llDate >= llStartDate) Then
                For ilChf = LBound(tmCntrCode0) To UBound(tmCntrCode0) - 1 Step 1
                    slNameCode = tmCntrCode0(ilChf).sKey 'lbcCntrCode(0).List(ilChf)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    llChfCode = Val(slCode)
                    If llChfCode = tmSdf.lChfCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilChf
                If Not ilFound Then
                    tmHdSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tgChfSpot, imHdRecLen, tmHdSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    tmAdfSrchKey.iCode = tgChfSpot.iAdfCode
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                        slStr = Trim$(str$(tgChfSpot.lCntrNo)) & " R" & Trim$(str$(tgChfSpot.iCntRevNo)) & "-" & Trim$(str$(tgChfSpot.iExtRevNo)) & " " & Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                    Else
                        slStr = Trim$(str$(tgChfSpot.lCntrNo)) & " R" & Trim$(str$(tgChfSpot.iCntRevNo)) & "-" & Trim$(str$(tgChfSpot.iExtRevNo)) & " " & Trim$(tmAdf.sName)
                    End If
                    lbcSelection(1).AddItem slStr
                    tmCntrCode0(UBound(tmCntrCode0)).sKey = slStr & "\" & Trim$(str$(tgChfSpot.lCode))
                    ReDim Preserve tmCntrCode0(0 To UBound(tmCntrCode0) + 1) As SORTCODE
                End If
            End If
            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Loop
    End If
End Function
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
    Dim ilCount As Integer      'general counter
    Dim slStr As String
    Dim tlCff As CFF    'Only required to obtain length
    ReDim tgClfSpot(0 To 0) As CLFLIST
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imLBCDCtrls = 1
    imFirstActivate = True
    'mParseCmmdLine
    gGetSchParameters
    'Remove Move; Compact; Preempt passes
    sgMovePass = "N"
    sgCompPass = "N"
    'Allow preempt 11/18/98 (required for rank to work when altered)
    smPreemptPass = sgPreemptPass
    'sgPreemptPass = "N"
    imTerminate = False
    imFirstFocus = True
    imBypassFocus = False
    imSettingValue = False
    imProcessing = False
    imSetAll = True
    imAllClicked = False
    imLbcArrowSetting = False
    imChfSelectIndex = -1
    imBoxNo = -1 'Initialize current Box to N/A
    imChg = False
    imChgMode = False
    imBSMode = False
    imCalType = 0   'Standard
    lmEarliestAllowedDate = gDateValue(Format$(gNow(), "m/d/yy")) + 1
    imListField(1) = 15
    imListField(2) = 6 * igAlignCharWidth
    imListField(3) = 25 * igAlignCharWidth
    imListField(4) = 45 * igAlignCharWidth
    imListField(5) = 95 * igAlignCharWidth
    mInitBox
    UnSchd.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone UnSchd
    'UnSchd.Show
    'imcHelp.Picture = Traffic!imcHelp.Picture
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", UnSchd
    On Error GoTo 0
    imVefRecLen = Len(tmVef)     'Get and save VEF record length
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", UnSchd
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)     'Get and save VSF record length
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", UnSchd
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)     'Get and save ADF record length
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", UnSchd
    On Error GoTo 0
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", UnSchd
    On Error GoTo 0
    imHdRecLen = Len(tgChfSpot)     'Get and save CHF record length
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", UnSchd
    On Error GoTo 0
    imLnRecLen = Len(tgClfSpot(0).ClfRec) 'btrRecordLength(hmClf)     'Get Clf size
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", UnSchd
    On Error GoTo 0
    imAirRecLen = Len(tlCff) 'btrRecordLength(hmCff)    'Get Cff size
    hmCgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", UnSchd
    On Error GoTo 0
    imCgfRecLen = Len(tmCgf)     'Get and save ADF record length
    hmRdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rdf.Btr)", UnSchd
    On Error GoTo 0
    imRdfRecLen = Len(tmRdf) 'btrRecordLength(hmRdf)    'Get Cff size
    hmSdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", UnSchd
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf) 'btrRecordLength(hmRdf)    'Get Cff size
    hmMtf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMtf, "", sgDBPath & "Mtf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        lgMtfNoRecs = btrRecords(hmMtf)
        btrDestroy hmMtf
    Else
        lgMtfNoRecs = 0
    End If
    'Initialize save arrays
    For ilCount = LBound(smSave) To UBound(smSave) Step 1
        smSave(ilCount) = ""
    Next ilCount
    For ilCount = LBound(imSave) To UBound(imSave) Step 1
        imSave(ilCount) = -1
    Next ilCount
    imVefRecLen = Len(tmVef)
    lbcSelection(0).Clear    'Initialize List boxes
    lbcSelection(1).Clear    'Initialize List boxes
    'lbcVehCode.Clear
    ReDim tmVehCode(0 To 0) As SORTCODE
    lbcSelection(2).Clear
    lbcSelection(3).Clear
    'lbcCntrCode(0).Clear
    'lbcCntrCode(1).Clear
    ReDim tmCntrCode0(0 To 0) As SORTCODE
    ReDim tmCntrCode1(0 To 0) As SORTCODE
    mVehPop           'Populate vehicle list boxes
    If imTerminate Then
        Exit Sub
    End If
    mAdvtPop           'Populate vehicle list boxes
    If imTerminate Then
        Exit Sub
    End If
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    edcSDate.Text = slStr
    If plcType.Visible = False Then
        rbcUnschType(1).Value = True
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
    Dim ilLoop As Integer
    flTextHeight = pbcDates.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcSelection.Move 255, 480 ', lbcSelection(2).Width + fgPanelAdj, lbcSelection(2).Height + fgPanelAdj
    lbcSelection(0).Move fgBevelX - 30, fgBevelY - 15
    lbcSelection(1).Move lbcSelection(0).Left + plcSelection.Width - 2 * fgBevelX - lbcSelection(1).Width, lbcSelection(0).Top
    lbcSelection(2).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width
    lbcSelection(3).Move lbcSelection(1).Left, lbcSelection(1).Top, lbcSelection(1).Width
    lbcLine.Move fgBevelX, fgBevelY, plcSelection.Width - 2 * fgBevelX
    plcLine.Move plcSelection.Left, plcSelection.Top + plcSelection.Height + 90, lbcLine.Width + 2 * fgPanelAdj
    plcDates.Move (UnSchd.Width - (pbcDates.Width + fgPanelAdj)) \ 2, plcLine.Top, pbcDates.Width + fgPanelAdj, pbcDates.Height + fgPanelAdj
    pbcDates.Move plcDates.Left + fgBevelX, plcDates.Top + fgBevelY
    pbcLbcLine.Move lbcLine.Left + 15, lbcLine.Top + 15, lbcLine.Width - 315
    'Unlock/Lock
    gSetCtrl tmCtrls(STARTDATEINDEX), 30, 30, 1725, fgBoxStH
    'Avail/Spot
    gSetCtrl tmCtrls(ENDDATEINDEX), 1770, tmCtrls(STARTDATEINDEX).fBoxY, 1725, fgBoxStH
    'Start date
    gSetCtrl tmCtrls(STARTTIMEINDEX), tmCtrls(STARTDATEINDEX).fBoxX, tmCtrls(STARTDATEINDEX).fBoxY + fgStDeltaY, 1725, fgBoxStH
    'End date
    gSetCtrl tmCtrls(ENDTIMEINDEX), tmCtrls(ENDDATEINDEX).fBoxX, tmCtrls(STARTTIMEINDEX).fBoxY, 1725, fgBoxStH
    'Start Time
    gSetCtrl tmCtrls(UNSCHDINDEX), tmCtrls(STARTDATEINDEX).fBoxX, tmCtrls(STARTTIMEINDEX).fBoxY + fgStDeltaY, 1725, fgBoxStH
    'End date
    gSetCtrl tmCtrls(SCHDINDEX), tmCtrls(ENDDATEINDEX).fBoxX, tmCtrls(UNSCHDINDEX).fBoxY, 1725, fgBoxStH
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub
Private Sub mLinePop()
    Dim ilClf As Integer
    Dim slStr As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slProgTime As String
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilCff As Integer
    Dim ilVef As Integer
    Dim slGDate As String
    
    lbcLine.Clear
    pbcLbcLine_Paint
    ReDim tgCffSpot(0 To 0) As CFFLIST
    ReDim tgClfSpot(0 To 0) As CLFLIST
    Screen.MousePointer = vbHourglass
    If mReadChfRec() Then
        If mReadClfRec() Then
            For ilClf = 0 To UBound(tgClfSpot) - 1 Step 1
                If mReadCffRec(ilClf) Then
                    'Build line that is to be shown
                    tmRdfSrchKey.iCode = tgClfSpot(ilClf).ClfRec.iRdfCode  ' Rate card program/time File Code
                    ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        slProgTime = Trim$(tmRdf.sName)
                    Else
                        slProgTime = "Rate/Program missing"
                    End If
                    '5/17/11
                    ilVef = gBinarySearchVef(tgClfSpot(ilClf).ClfRec.iVefCode)
                    'If (tgClfSpot(ilClf).iFirstCff <> -1) Then
                    If (tgClfSpot(ilClf).iFirstCff <> -1) And (ilVef <> -1) Then
                        slStartDate = ""
                        slEndDate = ""
                        ilCff = tgClfSpot(ilClf).iFirstCff
                        gUnpackDate tgCffSpot(ilCff).CffRec.iStartDate(0), tgCffSpot(ilCff).CffRec.iStartDate(1), slStartDate
                        gUnpackDate tgCffSpot(ilCff).CffRec.iEndDate(0), tgCffSpot(ilCff).CffRec.iEndDate(1), slEndDate
                        Do
                            If (tgCffSpot(ilCff).iStatus = 0) Or (tgCffSpot(ilCff).iStatus = 1) Then
                                If tgMVef(ilVef).sType <> "G" Then
                                    gUnpackDate tgCffSpot(ilCff).CffRec.iEndDate(0), tgCffSpot(ilCff).CffRec.iEndDate(1), slEndDate
                                Else
                                    gUnpackDate tgCffSpot(ilCff).CffRec.iStartDate(0), tgCffSpot(ilCff).CffRec.iStartDate(1), slGDate
                                    If gDateValue(slGDate) < gDateValue(slStartDate) Then
                                        slStartDate = slGDate
                                    End If
                                    If gDateValue(slGDate) > gDateValue(slEndDate) Then
                                        slEndDate = slGDate
                                    End If
                                End If
                            End If
                            ilCff = tgCffSpot(ilCff).iNextCff
                        Loop While ilCff <> -1
                        If gDateValue(slEndDate) < gDateValue(slStartDate) Then
                            slDate = "Canceled before Start"
                        Else
                            slDate = slStartDate & "-" & slEndDate
                        End If
                    End If
                    slStr = Trim$(str$(tgClfSpot(ilClf).ClfRec.iLine)) & "|" & slProgTime & "|" & slDate & "||\" & Trim$(str$(ilClf))
                    lbcLine.AddItem slStr
                End If
            Next ilClf
        End If
    End If
    pbcLbcLine_Paint
    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadCffRec                     *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadCffRec(ilClfIndex As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilFirst                                                                               *
'******************************************************************************************

'
'   iRet = mReadCffRec(ilClfIndex)
'   Where:
'       ilClfIndex (I) - CLF index (starting at 0)
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String
    Dim ilUpperBound As Integer
    Dim tlCff As CFF
    Dim tlCffExt As CFFEXT    'Flight extract record
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilVef As Integer

    ilUpperBound = UBound(tgCffSpot)
    ilVef = gBinarySearchVef(tgClfSpot(ilClfIndex).ClfRec.iVefCode)
    lbcLnCode.Clear
    If ilVef = -1 Then
        mReadCffRec = False
        Exit Function
    End If
    btrExtClear hmCff   'Clear any previous extend operation
    ilExtLen = Len(tlCffExt)  'Extract operation record size
    tmAirSrchKey.lChfCode = tgChfSpot.lCode
    tmAirSrchKey.iClfLine = tgClfSpot(ilClfIndex).ClfRec.iLine
    tmAirSrchKey.iCntRevNo = tgClfSpot(ilClfIndex).ClfRec.iCntRevNo
    tmAirSrchKey.iPropVer = tgClfSpot(ilClfIndex).ClfRec.iPropVer
    tmAirSrchKey.iStartDate(0) = 0
    tmAirSrchKey.iStartDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmCff, tlCff, imAirRecLen, tmAirSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (tlCff.lChfCode = tgChfSpot.lCode) And (tlCff.iClfLine = tgClfSpot(ilClfIndex).ClfRec.iLine) And (ilVef <> -1) Then
        'If (tlCff.iClfVersion = tgClfSpot(ilClfIndex).ClfRec.iVersion) And (tlCff.sDelete <> "Y") Then
        '    gUnpackDateForSort tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
        '    ilRet = btrGetPosition(hmCff, llRecPos)
        '    slStr = slStr & "\" & Trim$(Str$(llRecPos))
        '    lbcLnCode.AddItem slStr    'Add ID (retain matching sorted order) and Code number to list box
        'End If
        If tgMVef(ilVef).sType <> "G" Then
            llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
            Call btrExtSetBounds(hmCff, llNoRec, -1, "UC", "CFFEXTPK", CFFEXTPK) 'Set extract limits (all records)
            ilOffSet = gFieldOffset("Cff", "CffChfCode")
            ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgChfSpot.lCode, 4)
            On Error GoTo mReadCffRecErr
            gBtrvErrorMsg ilRet, "mReadCffRec (btrExtAddLogicConst):" & "Cff.Btr", UnSchd
            On Error GoTo 0
            ilOffSet = gFieldOffset("Cff", "CffClfLine")
            ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgClfSpot(ilClfIndex).ClfRec.iLine, 2)
            On Error GoTo mReadCffRecErr
            gBtrvErrorMsg ilRet, "mReadCffRec (btrExtAddLogicConst):" & "Cff.Btr", UnSchd
            On Error GoTo 0
            ilOffSet = gFieldOffset("Cff", "CffCntRevNo")
            ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgClfSpot(ilClfIndex).ClfRec.iCntRevNo, 2)
            On Error GoTo mReadCffRecErr
            gBtrvErrorMsg ilRet, "mReadCffRec (btrExtAddLogicConst):" & "Cff.Btr", UnSchd
            On Error GoTo 0
            ilOffSet = gFieldOffset("Cff", "CffPropVer")
            ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgClfSpot(ilClfIndex).ClfRec.iPropVer, 2)
            On Error GoTo mReadCffRecErr
            gBtrvErrorMsg ilRet, "mReadCffRec (btrExtAddLogicConst):" & "Cff.Btr", UnSchd
            On Error GoTo 0
            ilOffSet = gFieldOffset("Cff", "CffDelete")
            ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, ByVal "Y", 1)
            On Error GoTo mReadCffRecErr
            gBtrvErrorMsg ilRet, "mReadCffRec (btrExtAddLogicConst):" & "Cff.Btr", UnSchd
            On Error GoTo 0
            ilOffSet = gFieldOffset("Cff", "CffStartDate")
            ilRet = btrExtAddField(hmCff, ilOffSet, ilExtLen)  'Extract start date
            On Error GoTo mReadCffRecErr
            gBtrvErrorMsg ilRet, "mReadCffRec (btrExtAddField):" & "Cff.Btr", UnSchd
            On Error GoTo 0
            'ilRet = btrExtGetNextExt(hmCff)    'Extract record
            ilRet = btrExtGetNext(hmCff, tlCffExt, ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo mReadCffRecErr
                gBtrvErrorMsg ilRet, "mReadCffRec (btrExtGetNextExt):" & "Cff.Btr", UnSchd
                On Error GoTo 0
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmCff, tlCffExt, ilExtLen, llRecPos)
                End If
                Do While ilRet = BTRV_ERR_NONE
                    gUnpackDateForSort tlCffExt.iStartDate(0), tlCffExt.iStartDate(1), slStr
                    slStr = slStr & "\" & Trim$(str$(llRecPos))
                    lbcLnCode.AddItem slStr    'Add ID (retain matching sorted order) and Code number to list box
                    ilRet = btrExtGetNext(hmCff, tlCffExt, ilExtLen, llRecPos)
                    If ilRet = BTRV_ERR_REJECT_COUNT Then
                        ilRet = btrExtGetNext(hmCff, tlCffExt, ilExtLen, llRecPos)
                    End If
                Loop
                btrExtClear hmCff   'Clear any previous extend operation
                For ilLoop = 0 To lbcLnCode.ListCount - 1 Step 1
                    slNameCode = lbcLnCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    On Error GoTo mReadCffRecErr
                    gCPErrorMsg ilRet, "mReadCffRec (gParseItem field 2: lbcPrg)", UnSchd
                    On Error GoTo 0
                    slCode = Trim$(slCode)
                    llRecPos = CLng(slCode)
                    ilRet = btrGetDirect(hmCff, tgCffSpot(ilUpperBound), imAirRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    On Error GoTo mReadCffRecErr
                    gBtrvErrorMsg ilRet, "ReadCLFRec (btrGetDirect):" & "Cff.Btr", UnSchd
                    On Error GoTo 0
                    If tgClfSpot(ilClfIndex).iFirstCff = -1 Then
                        tgClfSpot(ilClfIndex).iFirstCff = ilUpperBound
                    Else
                        tgCffSpot(ilUpperBound - 1).iNextCff = ilUpperBound
                    End If
                    tgCffSpot(ilUpperBound).iNextCff = -1
                    tgCffSpot(ilUpperBound).lRecPos = llRecPos
                    tgCffSpot(ilUpperBound).iStatus = 1 'Old and retain
                    ilUpperBound = ilUpperBound + 1
                    ReDim Preserve tgCffSpot(0 To ilUpperBound)
                    tgCffSpot(ilUpperBound).iStatus = -1 'Not Used
                    tgCffSpot(ilUpperBound).iNextCff = -1
                    tgCffSpot(ilUpperBound).lRecPos = 0
                Next ilLoop
            End If
        Else
            tmCgfSrchKey1.lClfCode = tgClfSpot(ilClfIndex).ClfRec.lCode
            ilRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tgClfSpot(ilClfIndex).ClfRec.lCode = tmCgf.lClfCode)
                gCgfToCff tgClfSpot(ilClfIndex).ClfRec, tmCgf, tmCgfCff()
                LSet tgCffSpot(ilUpperBound) = tmCgfCff(0)  'tmCgfCff(1)
                tgCffSpot(ilUpperBound).iGameNo = tmCgf.iGameNo
                If tgClfSpot(ilClfIndex).iFirstCff = -1 Then
                    tgClfSpot(ilClfIndex).iFirstCff = ilUpperBound
                Else
                    tgCffSpot(ilUpperBound - 1).iNextCff = ilUpperBound
                End If
                tgCffSpot(ilUpperBound).iNextCff = -1
                tgCffSpot(ilUpperBound).lRecPos = llRecPos
                tgCffSpot(ilUpperBound).iStatus = 1 'Old and retain
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tgCffSpot(0 To ilUpperBound)
                tgCffSpot(ilUpperBound).iStatus = -1 'Not Used
                tgCffSpot(ilUpperBound).iNextCff = -1
                tgCffSpot(ilUpperBound).lRecPos = 0
                ilRet = btrGetNext(hmCgf, tmCgf, imCgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            Erase tmCgfCff
        End If
    End If
    mReadCffRec = True
    Exit Function
mReadCffRecErr:
    On Error GoTo 0
    mReadCffRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadChfRec                     *
'*                                                     *
'*             Created:7/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadChfRec() As Integer
'
'   iRet = mReadChfRec()
'   Where:
'       imChfSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    If imChfSelectIndex < 0 Then
        mReadChfRec = False
        Exit Function
    End If
    slNameCode = tmCntrCode1(imChfSelectIndex).sKey  'lbcCntrCode(1).List(imChfSelectIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadChfRecErr
    gCPErrorMsg ilRet, "mReadRecErr (gParseItem field 2)", UnSchd
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmHdSrchKey.lCode = CLng(slCode)
    ilRet = btrGetEqual(hmCHF, tgChfSpot, imHdRecLen, tmHdSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mReadChfRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetEqual: Contract)", UnSchd
    On Error GoTo 0
    ilRet = btrGetPosition(hmCHF, lmChfRecPos)
    On Error GoTo mReadChfRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetPosition: Contract)", UnSchd
    On Error GoTo 0
    mReadChfRec = True
    ReDim tgClfSpot(0 To 0) As CLFLIST
    tgClfSpot(0).iStatus = -1 'Not Used
    tgClfSpot(0).lRecPos = 0
    tgClfSpot(0).iFirstCff = -1
    ReDim tgCffSpot(0 To 0) As CFFLIST
    tgCffSpot(0).iStatus = -1 'Not Used
    tgCffSpot(0).lRecPos = 0
    tgCffSpot(0).iNextCff = -1
    Exit Function
mReadChfRecErr:
    On Error GoTo 0
    mReadChfRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadClfRec                     *
'*                                                     *
'*             Created:8/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadClfRec() As Integer
'
'   iRet = mReadClfRec
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpperBound As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim tlClf As CLF
    Dim tlClfExt As CLFEXT    'Contract line extract record
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim slLine As String
    Dim slVersion As String
    Dim ilAddLine As Integer
    Dim ilOffSet As Integer

    ReDim tgClfSpot(0 To 0) As CLFLIST
    ilUpperBound = UBound(tgClfSpot)
    tgClfSpot(ilUpperBound).iStatus = -1 'Not Used
    tgClfSpot(ilUpperBound).lRecPos = 0
    tgClfSpot(ilUpperBound).iFirstCff = -1
    lbcLnCode.Clear
    btrExtClear hmClf   'Clear any previous extend operation
    ilExtLen = Len(tlClfExt)  'Extract operation record size
    tmLnSrchKey.lChfCode = tgChfSpot.lCode
    tmLnSrchKey.iLine = 0
    tmLnSrchKey.iCntRevNo = 32000 ' 0 show latest version
    tmLnSrchKey.iPropVer = 32000 ' 0 show latest version
    ilRet = btrGetGreaterOrEqual(hmClf, tlClf, imLnRecLen, tmLnSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (tlClf.lChfCode = tgChfSpot.lCode) And (tlClf.sDelete <> "Y") Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmClf, llNoRec, -1, "UC", "CLFEXTPK", CLFEXTPK) 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Clf", "ClfChfCode")
        ilRet = btrExtAddLogicConst(hmClf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgChfSpot.lCode, 4)
        On Error GoTo mReadClfRecErr
        gBtrvErrorMsg ilRet, "mReadClfRec (btrExtAddLogicConst):" & "Clf.Btr", UnSchd
        On Error GoTo 0
        ilOffSet = gFieldOffset("Clf", "ClfDelete")
        ilRet = btrExtAddLogicConst(hmClf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, ByVal "Y", 1)
        On Error GoTo mReadClfRecErr
        gBtrvErrorMsg ilRet, "mReadClfRec (btrExtAddLogicConst):" & "Clf.Btr", UnSchd
        On Error GoTo 0
        ilOffSet = gFieldOffset("Clf", "ClfChfCode")
        ilRet = btrExtAddField(hmClf, ilOffSet, ilExtLen - 3) 'Extract start/end time, and days
        On Error GoTo mReadClfRecErr
        gBtrvErrorMsg ilRet, "mReadCLFRec (btrExtAddField):" & "Clf.Btr", UnSchd
        On Error GoTo 0
        ilOffSet = gFieldOffset("Clf", "ClfSchStatus")
        ilRet = btrExtAddField(hmClf, ilOffSet, 1) 'Extract start/end time, and days
        On Error GoTo mReadClfRecErr
        gBtrvErrorMsg ilRet, "mReadCLFRec (btrExtAddField):" & "Clf.Btr", UnSchd
        On Error GoTo 0
        ilOffSet = gFieldOffset("Clf", "ClfPropVer")
        ilRet = btrExtAddField(hmClf, ilOffSet, 2) 'Extract start/end time, and days
        On Error GoTo mReadClfRecErr
        gBtrvErrorMsg ilRet, "mReadCLFRec (btrExtAddField):" & "Clf.Btr", UnSchd
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmClf)    'Extract record
        ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mReadClfRecErr
            gBtrvErrorMsg ilRet, "mReadClfRec (btrExtGetNextExt):" & "Clf.Btr", UnSchd
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
            If ilRet = BTRV_ERR_REJECT_COUNT Then
                ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
            End If
            Do While ilRet = BTRV_ERR_NONE
                'Only show the latest line
                ilAddLine = True
                For ilLoop = 0 To lbcLnCode.ListCount - 1 Step 1
                    slNameCode = lbcLnCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 1, "\", slLine)
                    ilRet = gParseItem(slNameCode, 2, "\", slVersion)
                    If tlClfExt.iLine = Val(slCode) Then
                        If tlClfExt.iCntRevNo > Val(slVersion) Then
                            lbcLnCode.RemoveItem ilLoop
                        Else
                            ilAddLine = False
                        End If
                        Exit For
                    End If
                Next ilLoop
                If ilAddLine Then
                    slStr = Trim$(str$(tlClfExt.iLine))
                    Do While Len(slStr) < 4
                        slStr = "0" & slStr
                    Loop
                    slStr = slStr & "\" & Trim$(str$(tlClfExt.iCntRevNo))
                    slStr = slStr & "\" & Trim$(str$(llRecPos))
                    lbcLnCode.AddItem slStr
                End If
                ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
                End If
            Loop
            btrExtClear hmClf   'Clear any previous extend operation
            For ilLoop = lbcLnCode.ListCount - 1 To 0 Step -1
                slNameCode = lbcLnCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 3, "\", slCode)
                On Error GoTo mReadClfRecErr
                gCPErrorMsg ilRet, "mReadClfRec (gParseItem field 2: lbcPrg)", UnSchd
                On Error GoTo 0
                slCode = Trim$(slCode)
                llRecPos = CLng(slCode)
                ilRet = btrGetDirect(hmClf, tmClf, imLnRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If (Trim$(tmClf.sSchStatus) <> "F") Or (tmClf.sType = "O") Or (tmClf.sType = "A") Then
                    lbcLnCode.RemoveItem ilLoop
                End If
            Next ilLoop
            For ilLoop = 0 To lbcLnCode.ListCount - 1 Step 1
                slNameCode = lbcLnCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 3, "\", slCode)
                On Error GoTo mReadClfRecErr
                gCPErrorMsg ilRet, "mReadClfRec (gParseItem field 2: lbcPrg)", UnSchd
                On Error GoTo 0
                slCode = Trim$(slCode)
                llRecPos = CLng(slCode)
                ilRet = btrGetDirect(hmClf, tgClfSpot(ilUpperBound).ClfRec, imLnRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                On Error GoTo mReadClfRecErr
                gBtrvErrorMsg ilRet, "ReadClfRec (btrGetDirect):" & "Clf.Btr", UnSchd
                On Error GoTo 0
                tgClfSpot(ilUpperBound).iFirstCff = -1
                tgClfSpot(ilUpperBound).lRecPos = llRecPos
                tgClfSpot(ilUpperBound).iStatus = 1 'Old line
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tgClfSpot(0 To ilUpperBound) As CLFLIST
                tgClfSpot(ilUpperBound).iStatus = -1 'Not Used
                tgClfSpot(ilUpperBound).iFirstCff = -1
                tgClfSpot(ilUpperBound).lRecPos = 0
            Next ilLoop
        End If
    End If
    mReadClfRec = True
    Exit Function
mReadClfRecErr:
    On Error GoTo 0
    mReadClfRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSchd                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Schedule spots missed          *
'*                                                     *
'*******************************************************
Private Function mSchd() As Integer
    Dim ilVehCode As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slDate As String
    Dim llDate As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVpfIndex As Integer
    Dim ilRet As Integer
    Dim ilChf As Integer
    Dim llChfCode As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slTime As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilFound As Integer
    Dim llSpotTime As Long
    Dim ilVeh As Integer
    Dim ilVef As Integer
    Dim ilNoSelected As Integer
    Dim ilLineNo As Integer
    If rbcUnschType(0).Value Then   'Schedule/Unschedule
        If mTestSaveFields() = NO Then
            mSchd = False
            Exit Function
        End If
        If lbcSelection(0).ListIndex < 0 Then
            mSchd = False
            Exit Function
        End If
        If lbcSelection(1).Visible Then
            ilNoSelected = 0
            For ilChf = 0 To lbcSelection(1).ListCount - 1 Step 1
                If lbcSelection(1).Selected(ilChf) Then
                    ilNoSelected = ilNoSelected + 1
                End If
            Next ilChf
            If ilNoSelected < 1 Then
                mSchd = False
                Exit Function
            End If
        End If
        gObtainMissedReasonCode
        For ilVeh = 0 To lbcSelection(0).ListCount - 1 Step 1
            If lbcSelection(0).Selected(ilVeh) Then
                DoEvents
                If imTerminate Then
                    mSchd = False
                    Exit Function
                End If
                slNameCode = tmVehCode(ilVeh).sKey 'lbcVehCode.List(ilVeh)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                imVehCode = Val(slCode)
                Screen.MousePointer = vbHourglass
                ilVehCode = imVehCode
                ilVpfIndex = gVpfFind(UnSchd, ilVehCode)
                slStartDate = smSave(1)
                If smSave(2) <> "" Then
                    slEndDate = smSave(2)
                Else
                    slEndDate = "12/31/2040"
                    ilVef = gBinarySearchVef(imVehCode)
                    If ilVef <> -1 Then
                        llDate = gGetLatestLCFDate(hmLcf, "C", imVehCode)
                        If llDate > 0 Then
                            slEndDate = Format$(llDate, "m/d/yy")
                        End If
                    End If
                End If
                slStartTime = smSave(3)
                slEndTime = smSave(4)
                llStartDate = gDateValue(slStartDate)
                llEndDate = gDateValue(slEndDate)
                llStartTime = CLng(gTimeToCurrency(slStartTime, False))
                llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1
                If lbcSelection(1).Visible Then
                    For ilChf = 0 To lbcSelection(1).ListCount - 1 Step 1
                        If lbcSelection(1).Selected(ilChf) Then
                            DoEvents
                            If imTerminate Then
                                mSchd = False
                                Exit Function
                            End If
                            slNameCode = tmCntrCode0(ilChf).sKey 'lbcCntrCode(0).List(ilChf)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            llChfCode = Val(slCode)
                            'ReDim lgReschSdfCode(1 To 1) As Long
                            ReDim lgReschSdfCode(0 To 0) As Long
                            ilLineNo = 0
                            Do
                                DoEvents
                                If imTerminate Then
                                    mSchd = False
                                    Exit Function
                                End If
                                ilFound = False
                                tmSdfSrchKey0.iVefCode = ilVehCode
                                tmSdfSrchKey0.lChfCode = llChfCode
                                tmSdfSrchKey0.iLineNo = ilLineNo
                                tmSdfSrchKey0.lFsfCode = 0
                                slDate = Format$(llStartDate, "m/d/yy")
                                gPackDate slDate, ilDate0, ilDate1
                                tmSdfSrchKey0.iDate(0) = ilDate0
                                tmSdfSrchKey0.iDate(1) = ilDate1
                                tmSdfSrchKey0.sSchStatus = ""
                                tmSdfSrchKey0.iTime(0) = 0
                                tmSdfSrchKey0.iTime(1) = 0
                                ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVehCode) And (tmSdf.lChfCode = llChfCode)
                                    ilFound = True
                                    ilLineNo = tmSdf.iLineNo
                                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                                    If (llDate > llEndDate) Then
                                        Exit Do
                                    End If
                                    If (tmSdf.sSchStatus = "M") And (llDate >= llStartDate) Then
                                        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                                        llSpotTime = CLng(gTimeToCurrency(slTime, False))
                                        If (llSpotTime >= llStartTime) And (llSpotTime <= llEndTime) Then
                                            'ilRet = btrGetPosition(hmSdf, lgReschRecPos(UBound(lgReschRecPos)))
                                            lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                                            'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                                            ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                Loop
                                If (tmSdf.lChfCode <> llChfCode) Then
                                    Exit Do
                                End If
                                ilLineNo = ilLineNo + 1
                            Loop While ilFound
                            DoEvents
                            If imTerminate Then
                                mSchd = False
                                Exit Function
                            End If
                            If gOpenSchFiles() Then
                                'sgPreemptPass = smPreemptPass
                                If imSave(1) = 1 Then
                                    If (llStartDate + 6 <= llEndDate) And (llStartTime = 0) And (llEndTime >= 86399) Then
                                        igUsePreferred = True
                                    End If
                                End If
                                ilRet = gReSchSpots(True, 0, "YYYYYYY", 0, 86400)
                                igUsePreferred = False
                                'sgPreemptPass = "N"
                                gCloseSchFiles
                                If Not ilRet Then
                                    Screen.MousePointer = vbDefault
                                    ilRet = MsgBox("Scheduling Not Completed, Try Later", vbOKOnly + vbExclamation, "Rectify")
                                    mSchd = False
                                    imTerminate = True
                                    Exit Function
                                End If
                            End If
                        End If
                    Next ilChf
                Else
                    'ReDim lgReschSdfCode(1 To 1) As Long
                    ReDim lgReschSdfCode(0 To 0) As Long
                    tmSdfSrchKey1.iVefCode = ilVehCode
                    slDate = Format$(llStartDate, "m/d/yy")
                    gPackDate slDate, ilDate0, ilDate1
                    tmSdfSrchKey1.iDate(0) = ilDate0
                    tmSdfSrchKey1.iDate(1) = ilDate1
                    tmSdfSrchKey1.iTime(0) = 0
                    tmSdfSrchKey1.iTime(1) = 0
                    tmSdfSrchKey1.sSchStatus = "M"
                    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVehCode)
                        DoEvents
                        If imTerminate Then
                            mSchd = False
                            Exit Function
                        End If
                        ilFound = False
                        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                        If (llDate > llEndDate) Then
                            Exit Do
                        End If
                        If (tmSdf.sSchStatus = "M") And (llDate >= llStartDate) Then
                            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                            llSpotTime = CLng(gTimeToCurrency(slTime, False))
                            If (llSpotTime >= llStartTime) And (llSpotTime <= llEndTime) Then
                                'ilRet = btrGetPosition(hmSdf, lgReschRecPos(UBound(lgReschRecPos)))
                                lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                                'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                                ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                            End If
                        End If
                        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    Loop
                    DoEvents
                    If imTerminate Then
                        mSchd = False
                        Exit Function
                    End If
                    If gOpenSchFiles() Then
                        'sgPreemptPass = smPreemptPass
                        If imSave(1) = 1 Then
                            If (llStartDate + 6 <= llEndDate) And (llStartTime = 0) And (llEndTime >= 86399) Then
                                igUsePreferred = True
                            End If
                        End If
                        ilRet = gReSchSpots(True, 0, "YYYYYYY", 0, 86400)
                        'sgPreemptPass = "N"
                        igUsePreferred = False
                        gCloseSchFiles
                        If Not ilRet Then
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Scheduling Not Completed, Try Later", vbOKOnly + vbExclamation, "Rectify")
                            mSchd = False
                            imTerminate = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next ilVeh
        Screen.MousePointer = vbDefault
    Else
    End If
    mSchd = True
End Function
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
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < LBound(tmCtrls)) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case STARTDATEINDEX 'Start Date
            plcCalendar.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smSave(1) = slStr
            slStr = gFormatDate(slStr)
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
        Case ENDDATEINDEX 'End Date
            plcCalendar.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smSave(2) = slStr
            slStr = gFormatDate(slStr)
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
        Case STARTTIMEINDEX
            cmcDropDown.Visible = False
            plcTme.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smSave(3) = slStr
            If slStr <> "" Then
                slStr = gFormatTime(slStr, "A", "1")
            End If
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
        Case ENDTIMEINDEX
            cmcDropDown.Visible = False
            plcTme.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smSave(4) = slStr
            If slStr <> "" Then
                slStr = gFormatTime(slStr, "A", "1")
            End If
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
        Case UNSCHDINDEX
            pbcType.Visible = False  'Set visibility
            If imSave(1) = 0 Then
                slStr = "No"
            ElseIf imSave(1) = 1 Then
                slStr = "Yes"
            Else
                slStr = ""
            End If
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
        Case SCHDINDEX
            pbcType.Visible = False  'Set visibility
            If imSave(2) = 0 Then
                slStr = "No"
            ElseIf imSave(2) = 1 Then
                slStr = "Yes"
            Else
                slStr = ""
            End If
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
    End Select
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
    Unload UnSchd
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
Private Function mTestSaveFields() As Integer
'
'   iRet = mTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    If smSave(1) = "" Then
        ilRes = MsgBox("Start date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = STARTDATEINDEX
        mTestSaveFields = NO
        Exit Function
    Else
        If Not gValidDate(smSave(1)) Then
            ilRes = MsgBox("Start date must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = STARTDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If gDateValue(smSave(1)) < lmEarliestAllowedDate Then
        ilRes = MsgBox("Start date must be after today's date", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = STARTDATEINDEX
        mTestSaveFields = NO
        Exit Function
   End If
    If (smSave(2) = "") Then
        'ilRes = MsgBox("End date must be specified", vbOkOnly + vbExclamation, "Incomplete")
        'imBoxNo = ENDDATEINDEX
        'mTestSaveFields = NO
        'Exit Function
    Else
        If Not gValidDate(smSave(2)) Then
            ilRes = MsgBox("End date must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = ENDDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
        If smSave(1) <> "" Then
            If gDateValue(smSave(1)) > gDateValue(smSave(2)) Then
                ilRes = MsgBox("End date must be after start date", vbOKOnly + vbExclamation, "Incomplete")
                imBoxNo = ENDDATEINDEX
                mTestSaveFields = NO
                Exit Function
            End If
        End If
    End If
    If smSave(3) = "" Then
        ilRes = MsgBox("Start time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = STARTTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    Else
        If Not gValidTime(smSave(3)) Then
            ilRes = MsgBox("Time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = STARTTIMEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If smSave(4) = "" Then
        ilRes = MsgBox("End time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = ENDTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    Else
        If Not gValidTime(smSave(4)) Then
            ilRes = MsgBox("End time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = ENDTIMEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If gTimeToCurrency(smSave(3), False) > gTimeToCurrency(smSave(4), True) Then
        ilRes = MsgBox("End time must be after start time", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = ENDTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    mTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mUnSchd                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Sch/Unsch vehicles or balance  *
'*                      contract                       *
'*                                                     *
'*******************************************************
Private Function mUnSchd() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVef                                                                                 *
'******************************************************************************************

    Dim ilVehCode As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slDate As String
    Dim llDate As Long
    Dim llLastLogDate As Long
    Dim llEarliestAllowedDate As Long
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilAllSel As Integer
    Dim ilChf As Integer
    Dim ilClf As Integer
    Dim ilPos As Integer
    Dim slRight As String
    Dim slLeft As String
    Dim llChfCode As Long
    Dim ilVpfIndex As Integer
    Dim ilTotalAdded As Integer
    Dim ilTotalDeleted As Integer
    Dim ilCheckAllVeh As Integer    'Check other vehicles for line spots
    Dim ilCheckType As Integer      'Check contract type
    Dim ilVeh As Integer
    Dim ilVsf As Integer
    Dim ilNoSelected As Integer
    ReDim imLineCode(0 To 0) As Integer
    If rbcUnschType(0).Value Then   'Schedule/Unschedule
        If mTestSaveFields() = NO Then
            mUnSchd = False
            Exit Function
        End If
        If lbcSelection(1).Visible Then
            ilNoSelected = 0
            For ilChf = 0 To lbcSelection(1).ListCount - 1 Step 1
                If lbcSelection(1).Selected(ilChf) Then
                    ilNoSelected = ilNoSelected + 1
                End If
            Next ilChf
            If ilNoSelected < 1 Then
                mUnSchd = False
                Exit Function
            End If
        End If
        For ilVeh = 0 To lbcSelection(0).ListCount - 1 Step 1
            If lbcSelection(0).Selected(ilVeh) Then
                DoEvents
                If imTerminate Then
                    mUnSchd = False
                    Exit Function
                End If
                slNameCode = tmVehCode(ilVeh).sKey 'lbcVehCode.List(ilVeh)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                imVehCode = Val(slCode)
                Screen.MousePointer = vbHourglass
                ilVehCode = imVehCode
                ilVpfIndex = gVpfFind(UnSchd, ilVehCode)
                gUnpackDateLong tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), llLastLogDate
                slStartDate = smSave(1)
                If smSave(2) <> "" Then
                    slEndDate = smSave(2)
                Else
                    slEndDate = "12/31/2040"
                    llDate = gGetLatestLCFDate(hmLcf, "C", imVehCode)
                    If llDate > 0 Then
                        slEndDate = Format$(llDate, "m/d/yy")
                    End If
                End If
                slStartTime = smSave(3)
                slEndTime = smSave(4)
                ilAllSel = True
                If lbcSelection(1).Visible Then
                    For ilChf = 0 To lbcSelection(1).ListCount - 1 Step 1
                        If Not lbcSelection(1).Selected(ilChf) Then
                            ilAllSel = False
                            Exit For
                        End If
                    Next ilChf
                End If
                If ilAllSel Then
                    DoEvents
                    If imTerminate Then
                        mUnSchd = False
                        Exit Function
                    End If
                    llChfCode = -1
                    ilRet = gUnschSpots(ilVehCode, llChfCode, llLastLogDate, slStartDate, slEndDate, slStartTime, slEndTime, -1)
                    If Not ilRet Then
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Unschedule Not Completed, Try Later", vbOKOnly + vbExclamation, "Rectify")
                        mUnSchd = False
                        imTerminate = True
                        Exit Function
                    End If
                Else
                    For ilChf = 0 To lbcSelection(1).ListCount - 1 Step 1
                        If lbcSelection(1).Selected(ilChf) Then
                            DoEvents
                            If imTerminate Then
                                mUnSchd = False
                                Exit Function
                            End If
                            slNameCode = tmCntrCode0(ilChf).sKey 'lbcCntrCode(0).List(ilChf)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            llChfCode = Val(slCode)
                            ilRet = gUnschSpots(ilVehCode, llChfCode, llLastLogDate, slStartDate, slEndDate, slStartTime, slEndTime, -1)
                            If Not ilRet Then
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("Unschedule Not Completed, Try Later", vbOKOnly + vbExclamation, "Rectify")
                                mUnSchd = False
                                imTerminate = True
                                Exit Function
                            End If
                        End If
                    Next ilChf
                End If
            End If
        Next ilVeh
    Else    'Contract balancing
        If imChfSelectIndex < 0 Then
            mUnSchd = False
            Exit Function
        End If
        If ckcCheckAll.Value = vbChecked Then
            ilCheckAllVeh = True
        Else
            ilCheckAllVeh = False
        End If
        ilCheckType = True
        slStartDate = edcSDate.Text
        If slStartDate <> "" Then
            If Not gValidDate(slStartDate) Then
                edcSDate.SetFocus
                Exit Function
            End If
        Else
            slStartDate = Format$(gNow(), "m/d/yy")
        End If
        'Build lines to be checked
        Screen.MousePointer = vbHourglass
        ilRet = gOpenSchFiles()
        'ReDim lgUnschSdfCode(1 To 1) As Long
        ReDim lgUnschSdfCode(0 To 0) As Long
        For ilLoop = 0 To lbcLine.ListCount - 1 Step 1
            If lbcLine.Selected(ilLoop) Then
                DoEvents
                If imTerminate Then
                    gCloseSchFiles
                    mUnSchd = False
                    Exit Function
                End If
                slNameCode = lbcLine.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilClf = Val(slCode)
                imLineCode(UBound(imLineCode)) = ilClf
                ReDim Preserve imLineCode(UBound(imLineCode) + 1) As Integer
                'tgChfSpot set within mLinePop with a call to mReadChfRec
                'tgChfSpot required when ilCheckType = True
                ilRet = gCntrSchdSpotChkUnSchd(ilClf, ilCheckAllVeh, ilCheckType, slStartDate, ilTotalAdded, ilTotalDeleted)
                If Not ilRet Then
                    gCloseSchFiles
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Unschedule Not Completed, Try Later", vbOKOnly + vbExclamation, "Rectify")
                    mUnSchd = False
                    imTerminate = True
                    Exit Function
                End If
                ilPos = InStr(slNameCode, "|\")
                If ilPos > 0 Then
                    slRight = Mid$(slNameCode, ilPos)
                    ilPos = ilPos - 1
                    Do While ilPos > 0
                        If Mid$(slNameCode, ilPos, 1) = "|" Then
                            slLeft = Left$(slNameCode, ilPos)
                            slNameCode = slLeft & "Add:" & str$(ilTotalAdded) & " Del:" & str$(ilTotalDeleted) & slRight
                            lbcLine.RemoveItem ilLoop
                            lbcLine.AddItem slNameCode, ilLoop
                            Exit Do
                        End If
                        ilPos = ilPos - 1
                    Loop
                    pbcLbcLine_Paint
                End If
            End If
        Next ilLoop
        DoEvents
        If imTerminate Then
            gCloseSchFiles
            mUnSchd = False
            Exit Function
        End If
        'Reschedule any missed spots for lines
        'ReDim lgReschSdfCode(1 To 1) As Long
        ReDim lgReschSdfCode(0 To 0) As Long
        For ilLoop = 0 To UBound(imLineCode) - 1 Step 1
            ilClf = imLineCode(ilLoop)
            tmVefSrchKey.iCode = tgClfSpot(ilClf).ClfRec.iVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                Exit For
            End If
            If tmVef.sType <> "V" Then
                For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                    tmVsf.iFSCode(ilVsf) = 0
                Next ilVsf
                tmVsf.iFSCode(LBound(tmVsf.iFSCode)) = tgClfSpot(ilClf).ClfRec.iVefCode
            Else
                tmVsfSrchKey.lCode = tmVef.lVsfCode
                ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit For
                End If
                For ilVsf = UBound(tmVsf.iFSCode) To LBound(tmVsf.iFSCode) + 1 Step -1
                    tmVsf.iFSCode(ilVsf) = tmVsf.iFSCode(ilVsf - 1)
                Next ilVsf
                tmVsf.iFSCode(LBound(tmVsf.iFSCode)) = tgClfSpot(ilClf).ClfRec.iVefCode
            End If
            For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If tmVsf.iFSCode(ilVsf) > 0 Then
                    ilVehCode = tmVsf.iFSCode(ilVsf)
                    ilVpfIndex = -1
                    'For ilVeh = 0 To UBound(tgVpf) Step 1
                    '    If ilVehCode = tgVpf(ilVeh).iVefKCode Then
                        ilVeh = gBinarySearchVpfPlus(ilVehCode)
                        If ilVeh <> -1 Then
                            ilVpfIndex = ilVeh
                    '        Exit For
                        End If
                    'Next ilVeh
                    If ilVpfIndex = -1 Then
                        gCloseSchFiles
                        Screen.MousePointer = vbDefault
                        mUnSchd = True
                        Exit Function
                    End If
                    slDate = Format$(gNow(), "m/d/yy")
                    llEarliestAllowedDate = gDateValue(slDate) + 1
                    If (tgVpf(ilVpfIndex).iLLD(0) <> 0) Or (tgVpf(ilVpfIndex).iLLD(1) <> 0) Then
                        gUnpackDate tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), slDate
                        If slDate <> "" Then
                            llLastLogDate = gDateValue(slDate)
                        Else
                            llLastLogDate = 0
                        End If
                        If gDateValue(slDate) >= llEarliestAllowedDate Then
                            llEarliestAllowedDate = gDateValue(slDate) + 1
                        End If
                    Else
                        llLastLogDate = 0
                    End If
                    If gDateValue(slStartDate) > llEarliestAllowedDate Then
                        llEarliestAllowedDate = gDateValue(slStartDate)
                    End If
                    tmSdfSrchKey0.iVefCode = ilVehCode
                    tmSdfSrchKey0.lChfCode = tgClfSpot(ilClf).ClfRec.lChfCode
                    tmSdfSrchKey0.iLineNo = tgClfSpot(ilClf).ClfRec.iLine
                    tmSdfSrchKey0.lFsfCode = 0
                    slDate = Format$(llEarliestAllowedDate, "m/d/yy")
                    gPackDate slDate, ilDate0, ilDate1
                    tmSdfSrchKey0.iDate(0) = ilDate0
                    tmSdfSrchKey0.iDate(1) = ilDate1
                    tmSdfSrchKey0.sSchStatus = ""
                    tmSdfSrchKey0.iTime(0) = 0
                    tmSdfSrchKey0.iTime(1) = 0
                    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVehCode) And (tmSdf.lChfCode = tgClfSpot(ilClf).ClfRec.lChfCode) And (tmSdf.iLineNo = tgClfSpot(ilClf).ClfRec.iLine)
                        DoEvents
                        If imTerminate Then
                            gCloseSchFiles
                            mUnSchd = False
                            Exit Function
                        End If
                        If (tmSdf.sSchStatus = "M") Then
                            'ilRet = btrGetPosition(hmSdf, lgReschRecPos(UBound(lgReschRecPos)))
                            lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                            'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                            ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                        End If
                        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    Loop
                End If
            Next ilVsf
        Next ilLoop
        DoEvents
        If imTerminate Then
            gCloseSchFiles
            mUnSchd = False
            Exit Function
        End If
        sgPreemptPass = "N"
        ilRet = gReSchSpots(True, 0, "YYYYYYY", 0, 86400)
        sgPreemptPass = smPreemptPass
        gCloseSchFiles
        If Not ilRet Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Unschedule Not Completed, Try Later", vbOKOnly + vbExclamation, "Rectify")
            mUnSchd = False
            imTerminate = True
            Exit Function
        End If
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
    mUnSchd = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan     *
'*                                                     *
'*            Comments: Populate Vehicle and time zone *
'*                      list boxes                     *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer            'return status
    Dim llFilter As Long    'btrieve filter
    'Populate vehicle list box
    llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH ' Selling conventional vehicles
    'ilRet = gPopUserVehicleBox(UnSchd, ilFilter, lbcSelection(0), lbcVehCode)
    ilRet = gPopUserVehicleBox(UnSchd, llFilter, lbcSelection(0), tmVehCode(), smVehCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopUserVehicleBox)", UnSchd
        On Error GoTo 0
    End If
    'ilCntrType = 1
    'ilCurrent = 2   'Current; 1=All; 2=Current plus CBS
    'ilFilter = -2
    'ilVehCode = -1  'All
    'ilShowAdvt = True
    'ilShowDates = False 'True
    'ilRet = gPopCntrBox(UnSchd, ilCntrType, ilFilter, ilCurrent, 0, ilVehCode, lbcSelection(2), lbcCntrCode(1), ilShowAdvt, ilShowDates, False, False)
    'If ilRet <> CP_MSG_NOPOPREQ Then
    '    On Error GoTo mPopulateErr
    '    gCPErrorMsg ilRet, "mPopulate (gPopCntrBox)", UnSchd
    '    On Error GoTo 0
    'End If

    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
                If rbcUnschType(0).Value Then
                    edcDropDown.Text = Format$(llDate, "m/d/yy")
                    edcDropDown.SelStart = 0
                    edcDropDown.SelLength = Len(edcDropDown.Text)
                    imBypassFocus = True
                    edcDropDown.SetFocus
                    Exit Sub
                Else
                    edcSDate.Text = Format$(llDate, "m/d/yy")
                    edcSDate.SelStart = 0
                    edcSDate.SelLength = Len(edcSDate.Text)
                    imBypassFocus = True
                    edcSDate.SetFocus
                    Exit Sub
                End If
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    If rbcUnschType(0).Value Then
        edcDropDown.SetFocus
    Else
        edcSDate.SetFocus
    End If
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = LBound(tmCtrls) To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
            If (Y >= (tmCtrls(ilBox).fBoxY)) And (Y <= (tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcDates_Paint()
    Dim ilBox As Integer

    For ilBox = LBound(tmCtrls) To UBound(tmCtrls) Step 1
        'gPaintArea pbcDates, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
        pbcDates.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcDates.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY '- 30 '+ fgBoxInsetY
        pbcDates.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub

Private Sub pbcLbcLine_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLineEnd As Integer
    Dim ilField As Integer
    Dim slFields(0 To 3) As String  'Number of fields to display
    Dim llFgColor As Long
    Dim llWidth As Long
    Dim ilFieldIndex As Integer

    ilLineEnd = lbcLine.TopIndex + lbcLine.Height \ fgListHtArial825
    If ilLineEnd > lbcLine.ListCount Then
        ilLineEnd = lbcLine.ListCount
    End If
    If lbcLine.ListCount <= lbcLine.Height \ fgListHtArial825 Then
        llWidth = lbcLine.Width - 30
    Else
        llWidth = lbcLine.Width - igScrollBarWidth - 30
    End If
    pbcLbcLine.Width = llWidth
    pbcLbcLine.Cls
    llFgColor = pbcLbcLine.ForeColor
    For ilLoop = lbcLine.TopIndex To ilLineEnd - 1 Step 1
        pbcLbcLine.ForeColor = llFgColor
        If lbcLine.MultiSelect = 0 Then
            If lbcLine.ListIndex = ilLoop Then
                gPaintArea pbcLbcLine, CSng(0), CSng((ilLoop - lbcLine.TopIndex) * fgListHtArial825), CSng(pbcLbcLine.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLine.ForeColor = vbWhite
            End If
        Else
            If lbcLine.Selected(ilLoop) Then
                gPaintArea pbcLbcLine, CSng(0), CSng((ilLoop - lbcLine.TopIndex) * fgListHtArial825), CSng(pbcLbcLine.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLine.ForeColor = vbWhite
            End If
        End If
        slStr = lbcLine.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = 1 To 4 Step 1
            pbcLbcLine.CurrentX = imListField(ilField)
            pbcLbcLine.CurrentY = (ilLoop - lbcLine.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcLine, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcLine.Print slStr
        Next ilField
        pbcLbcLine.ForeColor = llFgColor
    Next ilLoop

End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            ilBox = 1
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case STARTDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            End If
            mSetShow imBoxNo
            imBoxNo = -1
            If lbcSelection(0).Visible Then
                lbcSelection(0).SetFocus
            Else
                lbcSelection(2).SetFocus
            End If
            Exit Sub
        Case ENDDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            End If
            ilBox = imBoxNo - 1
        Case STARTTIMEINDEX 'Time (first control within header)
            slStr = edcDropDown.Text
            If Not gValidTime(slStr) Then
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo - 1
        Case ENDTIMEINDEX 'Time (first control within header)
            slStr = edcDropDown.Text
            If Not gValidTime(slStr) Then
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
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            ilBox = SCHDINDEX
            Exit Sub
        Case STARTDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            End If
            ilBox = imBoxNo + 1
        Case ENDDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            End If
            ilBox = imBoxNo + 1
        Case STARTTIMEINDEX
            slStr = edcDropDown.Text
            If Not gValidTime(slStr) Then
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo + 1
        Case ENDTIMEINDEX
            slStr = edcDropDown.Text
            If Not gValidTime(slStr) Then
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo + 1
        Case SCHDINDEX
            mSetShow imBoxNo
            imBoxNo = -1
            cmcUpdate.SetFocus
            Exit Sub
        Case Else 'Last control within header
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
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
                        Case STARTTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                        Case ENDTIMEINDEX
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
Private Sub pbcType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcType_KeyPress(KeyAscii As Integer)
    If imBoxNo = UNSCHDINDEX Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            imSave(1) = 1
            tmCtrls(imBoxNo).iChg = True
            pbcType_Paint
        End If
        If (KeyAscii = Asc("N")) Or (KeyAscii = Asc("n")) Then
            imSave(1) = 0
            tmCtrls(imBoxNo).iChg = True
            pbcType_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSave(1) = 0 Then
                imSave(1) = 1
            Else
                imSave(1) = 0
            End If
            tmCtrls(imBoxNo).iChg = True
            pbcType_Paint
        End If
    ElseIf imBoxNo = SCHDINDEX Then
        If (KeyAscii = Asc("N")) Or (KeyAscii = Asc("n")) Then
            imSave(2) = 0
            tmCtrls(imBoxNo).iChg = True
            pbcType_Paint
        End If
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            imSave(2) = 1
            tmCtrls(imBoxNo).iChg = True
            pbcType_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSave(2) = 0 Then
                imSave(2) = 1
            Else
                imSave(2) = 0
            End If
            pbcType_Paint
            tmCtrls(imBoxNo).iChg = True
        End If
    End If
End Sub
Private Sub pbcType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = UNSCHDINDEX Then
        If imSave(1) = 0 Then
            imSave(1) = 1
        Else
            imSave(1) = 0
        End If
        tmCtrls(imBoxNo).iChg = True
        pbcType_Paint
    ElseIf imBoxNo = SCHDINDEX Then
        If imSave(2) = 0 Then
            imSave(2) = 1
        Else
            imSave(2) = 0
        End If
        tmCtrls(imBoxNo).iChg = True
        pbcType_Paint
    End If
End Sub
Private Sub pbcType_Paint()
    pbcType.Cls
    pbcType.CurrentX = fgBoxInsetX
    pbcType.CurrentY = -15 'fgBoxInsetY
    If imBoxNo = UNSCHDINDEX Then
        If imSave(1) = 0 Then
            pbcType.Print "No"
        ElseIf imSave(1) = 1 Then
            pbcType.Print "Yes"
        End If
    ElseIf imBoxNo = SCHDINDEX Then
        If imSave(2) = 0 Then
            pbcType.Print "No"
        ElseIf imSave(2) = 1 Then
            pbcType.Print "Yes"
        End If
    End If
End Sub
Private Sub plcDates_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcUnschType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcUnschType(Index).Value
    'End of coded added
    Dim ilNoSelected As Integer
    Dim ilLoop As Integer
    If Value Then
        If Index = 0 Then   'Vehicle
            pbcSTab.Visible = True
            pbcTab.Visible = True
            lbcSelection(3).Visible = False
            lbcSelection(2).Visible = False
            ilNoSelected = 0
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                If lbcSelection(0).Selected(ilLoop) Then
                    ilNoSelected = ilNoSelected + 1
                End If
            Next ilLoop
            If ilNoSelected <= 1 Then
                lbcSelection(1).Visible = True
            Else
                lbcSelection(1).Visible = False
            End If
            lbcSelection(0).Visible = True
            plcLine.Visible = False
            plcDates.Visible = True
            pbcDates.Visible = True
        Else
            pbcSTab.Visible = False
            pbcTab.Visible = False
            lbcSelection(0).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = True
            lbcSelection(3).Visible = True
            plcDates.Visible = False
            pbcDates.Visible = False
            plcLine.Visible = True
            plcCalendar.Move plcLine.Left + plcLine.Width - plcCalendar.Width - fgBevelX, plcLine.Top + edcSDate.Top - plcCalendar.Height
        End If
    End If
End Sub
Private Sub rbcUnschType_GotFocus(Index As Integer)
    plcCalendar.Visible = False
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    If plcType.Visible = True Then
        plcScreen.Print "Rectify"
    Else
        plcScreen.Print "Rectify: Contract Balancing"
    End If
End Sub
