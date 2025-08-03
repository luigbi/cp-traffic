VERSION 5.00
Begin VB.Form Purge 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5955
   ClientLeft      =   330
   ClientTop       =   1080
   ClientWidth     =   10485
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
   ScaleHeight     =   5955
   ScaleWidth      =   10485
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
      Left            =   3780
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Purge.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   27
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
            TabIndex        =   28
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
         TabIndex        =   24
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
         TabIndex        =   26
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
         Left            =   330
         TabIndex        =   25
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8760
      Top             =   4995
   End
   Begin VB.PictureBox plcInvInfo 
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
      Height          =   1020
      Left            =   645
      ScaleHeight     =   990
      ScaleWidth      =   4065
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   825
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label lacInvInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Rotation Dates:"
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
         Left            =   105
         TabIndex        =   34
         Top             =   45
         Width           =   3840
      End
      Begin VB.Label lacInvInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Short Title:"
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
         Left            =   105
         TabIndex        =   37
         Top             =   270
         Width           =   3765
      End
      Begin VB.Label lacInvInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "ISCI:"
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
         Left            =   105
         TabIndex        =   36
         Top             =   510
         Width           =   3870
      End
      Begin VB.Label lacInvInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Creative Title:"
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
         Index           =   3
         Left            =   105
         TabIndex        =   35
         Top             =   735
         Width           =   3885
      End
   End
   Begin VB.Timer tmcClick 
      Interval        =   2000
      Left            =   7935
      Top             =   5040
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
      Left            =   8640
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4935
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
      Left            =   8355
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4935
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
      Left            =   8610
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5220
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
      Height          =   165
      Left            =   45
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5385
      Width           =   75
   End
   Begin VB.PictureBox plcSort 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   360
      ScaleHeight     =   420
      ScaleWidth      =   3840
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4950
      Width           =   3840
      Begin VB.OptionButton rbcSort 
         Caption         =   "Cart #"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   1575
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   195
         Width           =   1350
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "ISCI"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   825
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   195
         Width           =   825
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Advertiser"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   2580
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Width           =   1185
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Rotation End Date"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   825
         TabIndex        =   18
         Top             =   0
         Value           =   -1  'True
         Width           =   1800
      End
   End
   Begin VB.PictureBox pbcDnMove 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5355
      Picture         =   "Purge.frx":2E1A
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1935
      Width           =   180
   End
   Begin VB.PictureBox pbcUpMove 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4785
      Picture         =   "Purge.frx":2EF4
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2985
      Width           =   180
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   3045
      TabIndex        =   20
      Top             =   5595
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4665
      TabIndex        =   21
      Top             =   5595
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   6255
      TabIndex        =   22
      Top             =   5595
      Width           =   1050
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   1095
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox plcPurgeDisposition 
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
      ForeColor       =   &H00000000&
      Height          =   4200
      Left            =   5940
      ScaleHeight     =   4200
      ScaleWidth      =   4260
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   705
      Width           =   4260
      Begin VB.ListBox lbcPurgeDisposition 
         Appearance      =   0  'Flat
         Height          =   3810
         Left            =   60
         MultiSelect     =   2  'Extended
         TabIndex        =   16
         Top             =   285
         Width           =   4140
      End
   End
   Begin VB.PictureBox plcAskDisposition 
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
      ForeColor       =   &H00000000&
      Height          =   4200
      Left            =   165
      ScaleHeight     =   4200
      ScaleWidth      =   4260
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   705
      Width           =   4260
      Begin VB.ListBox lbcAskDisposition 
         Appearance      =   0  'Flat
         Height          =   3810
         Left            =   60
         MultiSelect     =   2  'Extended
         TabIndex        =   10
         Top             =   285
         Width           =   4140
      End
   End
   Begin VB.PictureBox plcCopyInv 
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   180
      ScaleHeight     =   360
      ScaleWidth      =   10035
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   10095
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
         Left            =   60
         TabIndex        =   2
         Top             =   30
         Width           =   1575
      End
      Begin VB.TextBox edcNowMany 
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
         Left            =   5835
         MaxLength       =   20
         TabIndex        =   7
         Top             =   75
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmcRotEndDate 
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
         Left            =   4620
         Picture         =   "Purge.frx":2FCE
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   75
         Width           =   195
      End
      Begin VB.TextBox edcRotEndDate 
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
         Left            =   3450
         MaxLength       =   20
         TabIndex        =   4
         Top             =   75
         Width           =   1170
      End
      Begin VB.ComboBox cbcAdvt 
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
         Left            =   5520
         TabIndex        =   8
         Top             =   45
         Width           =   4485
      End
      Begin VB.Label lacNowMany 
         Appearance      =   0  'Flat
         Caption         =   "How Many"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4935
         TabIndex        =   6
         Top             =   75
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lacRotEndDate 
         Appearance      =   0  'Flat
         Caption         =   "Rotation End Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1950
         TabIndex        =   3
         Top             =   75
         Width           =   1470
      End
   End
   Begin VB.CommandButton cmcMoveToAsk 
      Appearance      =   0  'Flat
      Caption         =   "    Mo&ve"
      Height          =   300
      Left            =   4665
      TabIndex        =   13
      Top             =   2925
      Width           =   945
   End
   Begin VB.CommandButton cmcMoveToPurge 
      Appearance      =   0  'Flat
      Caption         =   "M&ove   "
      Height          =   300
      Left            =   4665
      TabIndex        =   11
      Top             =   1875
      Width           =   945
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   5475
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "Purge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Purge.frm on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Purge.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Copy Purge screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmCif As CIF        'Cif record image
Dim tmCifSrchKey0 As LONGKEY0    'Cif key record image
Dim tmCifSrchKey1 As CIFKEY1    'Cif key record image
Dim hmCif As Integer    'Copy inventory file handle
Dim imCifRecLen As Integer        'Cif record length
Dim tmSortCif() As SORTCIF
'Copy product/ISCI file
Dim hmCpf As Integer 'Copy product/ISCI file handle
Dim tmCpf As CPF        'CPF record image
Dim tmCpfSrchKey As LONGKEY0    'CPF key record image
Dim imCpfRecLen As Integer        'CPF record length
'Advertiser file
Dim hmAdf As Integer 'Advertiser file handle
Dim tmAdf As ADF        'ADF record image
Dim tmAdfSrchKey As INTKEY0    'ADF key record image
Dim imAdfRecLen As Integer        'ADF record length
Dim tmMcf As MCF        'Mcf record image
Dim hmMcf As Integer    'Media code file handle
Dim tmMcfSrchKey As INTKEY0    'MCF key record image
Dim imMcfRecLen As Integer        'MCF record length
Dim imMcfIndex As Integer
Dim tmAdvtCode() As SORTCODE
Dim smAdvtCodeTag As String
Dim tmMediaCode() As SORTCODE
Dim smMediaCodeTag As String
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imChgModeMedia As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstFocus As Integer
Dim imFirstFocusMedia As Integer
Dim imAdvtIndex As Integer
Dim imComboBoxIndexAdvt As Integer
Dim imComboBoxIndexMedia As Integer
Dim imSelectDelay As Integer    'True=cbcSelect change mode
Dim imButton As Integer 'Value 1= Left button; 2=Right button; 4=Middle button
Dim imButtonIndex As Integer
Dim imIgnoreRightMove As Integer
Dim lmNowDate As Long
Dim imPurgedDate(0 To 1) As Integer
Dim lmErrorDate As Long
Dim imUpdateAllowed As Integer
'Drag
Dim imLbcHeight As Integer
Dim imDragIndexSrce As Integer  '
Dim imDragIndexDest As Integer  '
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragButton As Integer 'Value 1= Left button; 2=Right button; 4=Middle button
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer  'Shift state when mouse down event occurrs
Dim imDragSrce As Integer 'Values defined below
Dim imDragDest As Integer 'Values defined below
Dim imDragScroll As Integer 'Object to be scrolled (same values as below)
Const DRAGASK = 1
Const DRAGPURGE = 2
'Calendar info
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar

Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim bmMediaPopped() As Boolean
Dim rmRecSet As ADODB.Recordset
Dim lmCounter As Long

Private Sub cbcAdvt_Change()
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        If cbcAdvt.Text <> "" Then
            gManLookAhead cbcAdvt, imBSMode, imComboBoxIndexAdvt
            tmcClick.Enabled = False
            imSelectDelay = True
            tmcClick.Interval = 2000    '2 seconds
            tmcClick.Enabled = True
        Else
            tmcClick.Enabled = False
            imSelectDelay = True
            tmcClick.Interval = 2000    '2 seconds
            tmcClick.Enabled = True
        End If
    End If
    Exit Sub
End Sub
Private Sub cbcAdvt_Click()
    cbcAdvt_Change
End Sub
Private Sub cbcAdvt_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cbcAdvt_DropDown()
    tmcClick.Enabled = False
    imSelectDelay = False
End Sub
Private Sub cbcAdvt_GotFocus()
    plcCalendar.Visible = False
    imComboBoxIndexAdvt = imAdvtIndex
End Sub
Private Sub cbcAdvt_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcAdvt_KeyPress(KeyAscii As Integer)
    tmcClick.Enabled = False
    imSelectDelay = False
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcAdvt.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcAdvt_LostFocus()
    If imSelectDelay Then
        tmcClick.Enabled = False
        imSelectDelay = False
        mCbcAdvtChange
    End If
End Sub
Private Sub cbcMedia_Change()
    If imChgModeMedia = False Then
        imChgModeMedia = True
        If cbcMedia.Text <> "" Then
            gManLookAhead cbcMedia, imBSMode, imComboBoxIndexMedia
        End If
        imMcfIndex = cbcMedia.ListIndex
        mInvPop
        imChgModeMedia = False
    End If
End Sub
Private Sub cbcMedia_Click()
    cbcMedia_Change
End Sub
Private Sub cbcMedia_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cbcMedia_GotFocus()
    Dim ilIndex As Integer
    If imTerminate Then
        Exit Sub
    End If
    If imFirstFocusMedia Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocusMedia = False
        If imFirstFocus Then
            imFirstFocus = False
        End If
        ilIndex = 0
        cbcMedia.ListIndex = ilIndex
        If cbcMedia.ListCount = 1 Then
            DoEvents
            edcRotEndDate.SetFocus
            Exit Sub
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
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcRotEndDate.SelStart = 0
    edcRotEndDate.SelLength = Len(edcRotEndDate.Text)
    edcRotEndDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcRotEndDate.SelStart = 0
    edcRotEndDate.SelLength = Len(edcRotEndDate.Text)
    edcRotEndDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If mSaveRecChg(True) = False Then
        If Not imTerminate Then
            pbcClickFocus.SetFocus
            Exit Sub
        Else
            cmcCancel_Click
            Exit Sub
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cmcDone_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcMoveToAsk_Click()
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim slName As String
    Dim slStr As String
    Dim ilIndex As Integer
    For ilIndex = lbcPurgeDisposition.ListCount - 1 To 0 Step -1
        If lbcPurgeDisposition.Selected(ilIndex) Then
            ilCount = 0
            slName = lbcPurgeDisposition.List(ilIndex)
            For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
                If tmSortCif(ilLoop).tCif.sPurged <> "P" Then
                    ilCount = ilCount + 1
                Else
                    If rbcSort(2).Value Then            'sort by isci
                        slStr = Trim$(tmSortCif(ilLoop).sISCI) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
                    Else
                        slStr = Trim$(tmSortCif(ilLoop).sInvName) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
                    End If
                    If slStr = slName Then
                        tmSortCif(ilLoop).iChg = True
                        tmSortCif(ilLoop).tCif.sPurged = "A"         'Change status from Active to Purged
                        tmSortCif(ilLoop).tCif.sCartDisp = "A"         'Change status from Active to Purged
                        lbcPurgeDisposition.RemoveItem ilIndex
                        lbcAskDisposition.AddItem slName, ilCount
                        Exit For
                    End If
                End If
            Next ilLoop
        End If
    Next ilIndex
End Sub
Private Sub cmcMoveToAsk_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcMoveToPurge_Click()
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim slName As String
    Dim slStr As String
    Dim ilIndex As Integer
    For ilIndex = lbcAskDisposition.ListCount - 1 To 0 Step -1
        If lbcAskDisposition.Selected(ilIndex) Then
            ilCount = 0
            slName = lbcAskDisposition.List(ilIndex)
            For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
                If tmSortCif(ilLoop).tCif.sPurged = "P" Then
                    ilCount = ilCount + 1
                Else
                    If rbcSort(2).Value Then            'sort by isci
                        slStr = Trim$(tmSortCif(ilLoop).sISCI) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
                    Else
                        slStr = Trim$(tmSortCif(ilLoop).sInvName) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
                    End If
                    If slStr = slName Then
                        tmSortCif(ilLoop).iChg = True
                        tmSortCif(ilLoop).tCif.sPurged = "P"         'Change status from Active to Purged
                        tmSortCif(ilLoop).tCif.sCartDisp = "P"         'Change status from Active to Purged
                        lbcAskDisposition.RemoveItem ilIndex
                        lbcPurgeDisposition.AddItem slName, ilCount
                        Exit For
                    End If
                End If
            Next ilLoop
        End If
    Next ilIndex
End Sub
Private Sub cmcMoveToPurge_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcRotEndDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcRotEndDate.SelStart = 0
    edcRotEndDate.SelLength = Len(edcRotEndDate.Text)
    edcRotEndDate.SetFocus
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
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    mInvPop
    Screen.MousePointer = vbDefault  'Wait
End Sub
Private Sub cmcUpdate_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cmcUpdate_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcNowMany_Change()
    mInvPop
End Sub
Private Sub edcNowMany_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub edcNowMany_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
    plcCalendar.Visible = False
End Sub
Private Sub edcNowMany_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRotEndDate_Change()
    Dim slStr As String
    slStr = edcRotEndDate.Text
    If Not gValidDate(slStr) Then
        Exit Sub
    End If
    If gDateValue(slStr) >= lmNowDate Then
        Beep
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    mInvPop
End Sub
Private Sub edcRotEndDate_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub edcRotEndDate_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcRotEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcRotEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If ActiveControl.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub edcRotEndDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcRotEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcRotEndDate.Text = slDate
            End If
        End If
        edcRotEndDate.SelStart = 0
        edcRotEndDate.SelLength = Len(edcRotEndDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcRotEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Left arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcRotEndDate.Text = slDate
            End If
        End If
        edcRotEndDate.SelStart = 0
        edcRotEndDate.SelLength = Len(edcRotEndDate.Text)
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
    If (igWinStatus(COPYJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    Purge.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub Form_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
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
    
    Erase tmMediaCode
    Erase tmAdvtCode
    Erase tmSortCif
    btrExtClear hmMcf   'Clear any previous extend operation
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    btrExtClear hmCpf   'Clear any previous extend operation
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    btrExtClear hmCif   'Clear any previous extend operation
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf

    Set Purge = Nothing   'Remove data segment
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lacNowMany_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub lacRotEndDate_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub lbcAskDisposition_DragDrop(Source As control, X As Single, Y As Single)
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim slName As String
    Dim slStr As String

    If imDragDest = -1 Then
        mClearDrag
        Exit Sub
    End If
    Select Case imDragSrce
        Case DRAGPURGE
            'Move item from Purge to Ask
            ilCount = 0
            'Determine insert point by counting number of P until
            'moved item found
            slName = lbcPurgeDisposition.List(imDragIndexSrce)
            For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
                If tmSortCif(ilLoop).tCif.sPurged <> "P" Then
                    ilCount = ilCount + 1
                Else
                    If rbcSort(2).Value Then            'sort by ISCI
                        slStr = Trim$(tmSortCif(ilLoop).sISCI) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
                    Else
                        slStr = Trim$(tmSortCif(ilLoop).sInvName) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
                    End If
                    If slStr = slName Then
                        tmSortCif(ilLoop).iChg = True
                        tmSortCif(ilLoop).tCif.sPurged = "A"         'Change status from Active to Purged
                        tmSortCif(ilLoop).tCif.sCartDisp = "A"         'Change status from Active to Purged
                        lbcPurgeDisposition.RemoveItem imDragIndexSrce
                        lbcAskDisposition.AddItem slName, ilCount
                        Exit For
                    End If
                End If
            Next ilLoop
            mClearDrag
        Case DRAGASK
            mClearDrag
    End Select
End Sub
Private Sub lbcAskDisposition_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    imDragDest = -1
    If imDragSrce = DRAGPURGE Then
        If State = vbLeave Then
            lbcPurgeDisposition.DragIcon = IconTraf!imcIconDrag.DragIcon
            Exit Sub
        End If
        lbcPurgeDisposition.DragIcon = IconTraf!imcIconInsert.DragIcon
        imDragDest = DRAGASK
    ElseIf imDragSrce = DRAGASK Then
        lbcAskDisposition.DragIcon = IconTraf!imcIconDrag.DragIcon
    End If
End Sub
Private Sub lbcAskDisposition_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub lbcAskDisposition_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imButton = Button
    If Button = 2 Then  'Right Mouse
        imButtonIndex = (Y \ imLbcHeight) + lbcAskDisposition.TopIndex
        If imButtonIndex <= lbcAskDisposition.ListCount - 1 Then
            If imButtonIndex >= 0 Then
                mShowInvInfo DRAGASK, Y
            End If
        End If
        Exit Sub
    End If
    fmDragX = X
    fmDragY = Y
    imDragButton = Button
    imDragType = 0
    imDragShift = Shift
    imDragSrce = DRAGASK
    imDragIndexDest = -1
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub lbcAskDisposition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        If (Y < 0) Or (Y > lbcAskDisposition.height) Then
            imButtonIndex = 0
            plcInvInfo.Visible = False
            Exit Sub
        End If
        If (X < 0) Or (X > lbcAskDisposition.Width) Then
            imButtonIndex = 0
            plcInvInfo.Visible = False
            Exit Sub
        End If
        If imButtonIndex <> (Y \ imLbcHeight) + lbcAskDisposition.TopIndex Then
            imIgnoreRightMove = True
            imButtonIndex = (Y \ imLbcHeight) + lbcAskDisposition.TopIndex
            If (imButtonIndex >= 0) And (imButtonIndex <= lbcAskDisposition.ListCount - 1) Then
                mShowInvInfo DRAGASK, Y
            Else
                plcInvInfo.Visible = False
            End If
            imIgnoreRightMove = False
        End If
    End If
End Sub
Private Sub lbcAskDisposition_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If Button = 2 Then
        imButtonIndex = -1
        plcInvInfo.Visible = False
    End If
End Sub
Private Sub lbcPurgeDisposition_DragDrop(Source As control, X As Single, Y As Single)
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim slName As String
    Dim slStr As String
    If imDragDest = -1 Then
        mClearDrag
        Exit Sub
    End If
    Select Case imDragSrce
        Case DRAGASK
            'Move item from Ask to Purge
            ilCount = 0
            'Determine insert point by counting number of P until
            'moved item found
            slName = lbcAskDisposition.List(imDragIndexSrce)
            For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
                If tmSortCif(ilLoop).tCif.sPurged = "P" Then
                    ilCount = ilCount + 1
                Else
                    If rbcSort(2).Value Then            'sort by isci
                        slStr = Trim$(tmSortCif(ilLoop).sISCI) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
                    Else
                        slStr = Trim$(tmSortCif(ilLoop).sInvName) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
                    End If
                    If slStr = slName Then
                        tmSortCif(ilLoop).iChg = True
                        tmSortCif(ilLoop).tCif.sPurged = "P"         'Change status from Active to Purged
                        tmSortCif(ilLoop).tCif.sCartDisp = "P"         'Change status from Active to Purged
                        lbcAskDisposition.RemoveItem imDragIndexSrce
                        lbcPurgeDisposition.AddItem slName, ilCount
                        Exit For
                    End If
                End If
            Next ilLoop
            mClearDrag
        Case DRAGPURGE
            mClearDrag
    End Select
End Sub
Private Sub lbcPurgeDisposition_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    imDragDest = -1
    If imDragSrce = DRAGASK Then
        If State = vbLeave Then
            lbcAskDisposition.DragIcon = IconTraf!imcIconDrag.DragIcon
            Exit Sub
        End If
        lbcAskDisposition.DragIcon = IconTraf!imcIconInsert.DragIcon
        imDragDest = DRAGPURGE
    ElseIf imDragSrce = DRAGPURGE Then
        lbcPurgeDisposition.DragIcon = IconTraf!imcIconDrag.DragIcon
    End If
End Sub
Private Sub lbcPurgeDisposition_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub lbcPurgeDisposition_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imButton = Button
    If Button = 2 Then  'Right Mouse
        imButtonIndex = Y \ imLbcHeight + lbcPurgeDisposition.TopIndex
        If imButtonIndex <= lbcPurgeDisposition.ListCount - 1 Then
            If imButtonIndex >= 0 Then
                mShowInvInfo DRAGPURGE, Y
            End If
        End If
        Exit Sub
    End If
    fmDragX = X
    fmDragY = Y
    imDragButton = Button
    imDragType = 0
    imDragShift = Shift
    imDragSrce = DRAGPURGE
    imDragIndexDest = -1
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub lbcPurgeDisposition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        If (Y < 0) Or (Y > lbcPurgeDisposition.height) Then
            imButtonIndex = 0
            plcInvInfo.Visible = False
            Exit Sub
        End If
        If (X < 0) Or (X > lbcPurgeDisposition.Width) Then
            imButtonIndex = 0
            plcInvInfo.Visible = False
            Exit Sub
        End If
        If imButtonIndex <> (Y \ imLbcHeight) + lbcPurgeDisposition.TopIndex Then
            imIgnoreRightMove = True
            imButtonIndex = (Y \ imLbcHeight) + lbcPurgeDisposition.TopIndex
            If (imButtonIndex >= 0) And (imButtonIndex <= lbcPurgeDisposition.ListCount - 1) Then
                mShowInvInfo DRAGPURGE, Y
            Else
                plcInvInfo.Visible = False
            End If
            imIgnoreRightMove = False
        End If
    End If
End Sub
Private Sub lbcPurgeDisposition_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If Button = 2 Then
        imButtonIndex = -1
        plcInvInfo.Visible = False
    End If
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
    'ilRet = gPopAdvtBox(Purge, cbcAdvt, lbcAdvtCode)
    ilRet = gPopAdvtBox(Purge, cbcAdvt, tmAdvtCode(), smAdvtCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", Purge
        On Error GoTo 0
        cbcAdvt.AddItem "[N/A]", 0  'Force as first item on list
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
    slStr = edcRotEndDate.Text
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
'*      Procedure Name:mCbcAdvtChange                  *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process advertiser change      *
'*                                                     *
'*******************************************************
Private Sub mCbcAdvtChange()

    If imChgMode = False Then
        imChgMode = True
        If cbcAdvt.Text <> "" Then
            gManLookAhead cbcAdvt, imBSMode, imComboBoxIndexAdvt
            imAdvtIndex = cbcAdvt.ListIndex
            mInvPop
        Else
            ReDim tmSortCif(0 To 0) As SORTCIF
            lbcAskDisposition.Clear
            lbcPurgeDisposition.Clear
        End If
        imChgMode = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearDrag                      *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear drag when drop on illegal*
'*                      object                         *
'*                                                     *
'*******************************************************
Private Sub mClearDrag()
    imDragIndexSrce = -1
    imDragSrce = -1
    imDragScroll = -1
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
    Dim slDate As String
    Dim ilRet As Integer
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    'mParseCmmdLine
    Purge.height = cmcDone.Top + 5 * cmcDone.height / 3
    gCenterStdAlone Purge
    'Purge.Show
    Screen.MousePointer = vbHourglass
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    imFirstFocus = True
    imFirstFocusMedia = True
    imSelectDelay = False
    imIgnoreRightMove = False
    imChgModeMedia = False
    imChgMode = False
    imBSMode = False
    imButton = 0
    imMcfIndex = -1
    imAdvtIndex = -1
    imLbcHeight = fgListHtArial825
    hmAdf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", Purge
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmCif = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", Purge
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmCpf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", Purge
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmMcf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mcf.Btr)", Purge
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)
    imCalType = 0   'Standard
    slDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slDate)
    gPackDate slDate, imPurgedDate(0), imPurgedDate(1)
    lmErrorDate = 0
    slDate = Format$(gDateValue(gNow()) - 30, "m/d/yy")
    edcRotEndDate.Text = slDate
    cbcMedia.Clear  'Force list to be populated
    mMediaPop
    If imTerminate Then
        Exit Sub
    End If
    mAdvtPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mInitBox
    'Purge.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    'gCenterModalForm Purge
    Purge.Refresh
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
'*             Created:7/19/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInvPop                         *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection Name    *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mInvPop()
    Dim ilRet As Integer
    Dim ilAdvtCode As Integer
    Dim ilMediaCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim llEndDate As Long
    Dim slDate As String
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim slName As String
    Dim slAdvtName As String
    Dim slISCI As String
    Dim slStr As String
    Dim dlDate As Date

    Screen.MousePointer = vbHourglass
    ReDim tmSortCif(0 To 0) As SORTCIF
    ilUpper = 0
    lbcAskDisposition.Clear
    lbcPurgeDisposition.Clear
    If imMcfIndex < 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If ((tgSpf.sUseCartNo = "N") Or (tgSpf.sUseCartNo = "B")) Then
        If imMcfIndex = 0 Then
            ilMediaCode = 0
        Else
            slNameCode = tmMediaCode(imMcfIndex - 1).sKey 'lbcMediaCode.List(imMcfIndex)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilMediaCode = Val(slCode)
        End If
    Else
        slNameCode = tmMediaCode(imMcfIndex).sKey  'lbcMediaCode.List(imMcfIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilMediaCode = Val(slCode)
    End If
    slDate = edcRotEndDate.Text
    If Not gValidDate(slDate) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    llEndDate = gDateValue(slDate)
    If imAdvtIndex < 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If llEndDate >= lmNowDate Then
        Screen.MousePointer = vbDefault
        If lmErrorDate <> llEndDate Then
            MsgBox "Date must be prior to Today's Date", vbInformation, "Error"
        End If
        lmErrorDate = llEndDate
        Exit Sub
    End If
    If imAdvtIndex > 0 Then
        slNameCode = tmAdvtCode(imAdvtIndex - 1).sKey  'lbcAdvtCode.List(imAdvtIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilAdvtCode = Val(slCode)
    Else
        ilAdvtCode = 0
    End If
    
    ' ilAdvtCode = Test()
    
    'slStr = Format$(gNow(), "m/d/yy")
    'slStr = gObtainStartStd(slStr)
    slStr = slDate
    
    ' TTP 10807 - JD 08-15-23
    ' Run the same process as the utility "SetCopyDates" before the purge process to correct the
    ' blank dates.
    If Not bmMediaPopped(imMcfIndex) Then   ' Only do this one time for each media type selected.
        If Not gSetCopyDates(slStr, cbcMedia.Text, ilMediaCode) Then
            gMsgBox "Set Copy Dates failed. Purge cannot be completed."
            Exit Sub
        End If
        bmMediaPopped(imMcfIndex) = True
    End If
    
    tmCifSrchKey1.iMcfCode = ilMediaCode
    tmCifSrchKey1.sName = ""
    tmCifSrchKey1.sCut = ""
    ilRet = btrGetGreaterOrEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    Do While (ilRet = BTRV_ERR_NONE) And (tmCif.iMcfCode = ilMediaCode)
        gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slDate
        If slDate = "" Then
        
            gUnpackDate tmCif.iDateEntrd(0), tmCif.iDateEntrd(1), slDate
            
'            ' TTP 10807 - JD 08-15-23
'            ' One more check was required when this condition occurs. If any rows are returned
'            ' in the new function RotDatesDetected then do not use this record.
'            If tmCif.sPurged = "A" Then
'                lmCounter = lmCounter + 1
'                If RotDatesDetected(tmCif.lCode, slStr, ilMediaCode) Then
'                    gLogMsg "mInvPop: Rotation date not set properly ", "SetCopyInventoryDates.txt", False
'                Else
'                    gUnpackDate tmCif.iDateEntrd(0), tmCif.iDateEntrd(1), slDate
'                End If
'            End If
        End If
        If slDate <> "" Then
            If (gDateValue(slDate) <= llEndDate) And (tmCif.sPurged = "A") Then
                If (ilAdvtCode = 0) Or (ilAdvtCode = tmCif.iAdfCode) Then
                    If tmAdf.iCode <> tmCif.iAdfCode Then
                        tmAdfSrchKey.iCode = tmCif.iAdfCode
                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point
                    End If
                    gUnpackDateForSort tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slDate
                    If (slDate = "") Or (slDate = "00000") Then
                        gUnpackDateForSort tmCif.iDateEntrd(0), tmCif.iDateEntrd(1), slDate
                    End If
                    If tmCif.iMcfCode = 0 Then
                        tmCpfSrchKey.lCode = tmCif.lcpfCode
                        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            slName = Trim$(tmCpf.sISCI)
                            slISCI = Trim$(tmCpf.sISCI)          '3-19-20
                        End If
                    Else
                        If tmMcf.iCode <> tmCif.iMcfCode Then
                            tmMcfSrchKey.iCode = tmCif.iMcfCode
                            ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point
                        End If
                        slName = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                        If (Len(Trim$(tmCif.sCut)) <> 0) Then
                            slName = slName & "-" & tmCif.sCut
                        End If
                        
                        'using carts, but just in case sorting by isci
                        slISCI = ""
                        If tmCif.lcpfCode > 0 Then
                            tmCpfSrchKey.lCode = tmCif.lcpfCode
                            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                slISCI = Trim$(tmCpf.sISCI)          '3-19-20
                            End If
                        End If
                        
                    End If
                    If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                        slAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                    Else
                        slAdvtName = Trim$(tmAdf.sName)
                    End If
                    If rbcSort(0).Value Then
                        tmSortCif(ilUpper).sKey = slDate & "|" & slAdvtName & "|" & slName
                    ElseIf rbcSort(1).Value Then
                        tmSortCif(ilUpper).sKey = slAdvtName & "|" & slDate & "|" & slName
                    ElseIf rbcSort(2).Value Then        'sort by ISCI
                        tmSortCif(ilUpper).sKey = slISCI & "|" & slAdvtName & "|" & slDate
                    Else                                'sort by Cart #
                        tmSortCif(ilUpper).sKey = slName & "|" & slAdvtName & "|" & slDate
                    End If
                    tmSortCif(ilUpper).sInvName = slName
                    tmSortCif(ilUpper).sAdvtName = slAdvtName
                    tmSortCif(ilUpper).sDate = slDate
                    tmSortCif(ilUpper).sISCI = slISCI
                    If tmCif.sCartDisp = "P" Then
                        tmSortCif(ilUpper).iChg = True
                        tmCif.sPurged = "P"         'Change status from Active to Purged
                    Else
                        tmSortCif(ilUpper).iChg = False
                    End If
                    tmSortCif(ilUpper).tCif = tmCif
                    ilUpper = ilUpper + 1
                    ReDim Preserve tmSortCif(0 To ilUpper) As SORTCIF
                End If
            End If
        End If
        ilRet = btrGetNext(hmCif, tmCif, imCifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilUpper = UBound(tmSortCif)
    If ilUpper > 0 Then
        ArraySortTyp fnAV(tmSortCif(), 0), ilUpper, 0, LenB(tmSortCif(0)), 0, LenB(tmSortCif(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
        If rbcSort(2).Value Then            'sort by isci
            slName = Trim$(tmSortCif(ilLoop).sISCI) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
        Else
            slName = Trim$(tmSortCif(ilLoop).sInvName) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
        End If
        If tmSortCif(ilLoop).tCif.sCartDisp = "P" Then
            lbcPurgeDisposition.AddItem slName
        Else
            lbcAskDisposition.AddItem slName
        End If
    Next ilLoop
    Screen.MousePointer = vbDefault
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
    ilIndex = cbcMedia.ListIndex
    If ilIndex >= 0 Then
        slName = cbcMedia.List(ilIndex)
    End If
    ilFilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffSet(0) = 0
    'ilRet = gIMoveListBox(Purge, cbcMedia, lbcMediaCode, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(Purge, cbcMedia, tmMediaCode(), smMediaCodeTag, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mMediaPopErr
        gCPErrorMsg ilRet, "mMediaPop (gIMoveListBox)", Purge
        On Error GoTo 0
'        cbcMedia.AddItem "[New]", 0  'Force as first item on list
        If (tgSpf.sUseCartNo = "N") Or (tgSpf.sUseCartNo = "B") Then
            cbcMedia.AddItem "[ISCI Only]", 0
        End If
        If (tgSpf.sUseCartNo = "N") Then        'no carts used
            rbcSort(3).Visible = False          'disable cart sort
        End If
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
    
    ReDim bmMediaPopped(0 To cbcMedia.ListCount)
    For ilIndex = 0 To cbcMedia.ListCount
        bmMediaPopped(ilIndex) = False
    Next
    
    Exit Sub
mMediaPopErr:
    On Error GoTo 0
    imTerminate = True
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
    Dim ilRet As Integer
    Dim ilErrRet As Integer
    Dim slMsg As String
    Dim slMsg2 As String
    Dim slStartDate As String
    Dim slEndDate As String
    
    Screen.MousePointer = vbHourglass  'Wait
    
    slMsg = "mSaveRec (btrUpdate: Cif)"
    For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
        If tmSortCif(ilLoop).iChg Then
            Do
                tmCifSrchKey0.lCode = tmSortCif(ilLoop).tCif.lCode
                ilErrRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilErrRet, slMsg, Purge
                On Error GoTo 0
                If tmSortCif(ilLoop).tCif.sPurged = "P" Then
                    tmSortCif(ilLoop).tCif.iPurgeDate(0) = imPurgedDate(0)
                    tmSortCif(ilLoop).tCif.iPurgeDate(1) = imPurgedDate(1)
                
                    ' JD 9/8/23 TTP 10809 Add this to the log file.
                    gUnpackDate tmSortCif(ilLoop).tCif.iRotStartDate(0), tmSortCif(ilLoop).tCif.iRotStartDate(1), slStartDate
                    gUnpackDate tmSortCif(ilLoop).tCif.iRotEndDate(0), tmSortCif(ilLoop).tCif.iRotEndDate(1), slEndDate
                    
                    slMsg2 = "Purging " & _
                    "Mcf Code:" & tmSortCif(ilLoop).tCif.iMcfCode & ", " & _
                    "Adf Code:" & Trim$(tmSortCif(ilLoop).sAdvtName) & ", " & _
                    "ISCI:" & Trim$(tmSortCif(ilLoop).sISCI) & ", " & _
                    "Inv Name:" & Trim$(tmSortCif(ilLoop).sInvName) & ", " & _
                    "Rot Start Date:" & slStartDate & ", " & _
                    "Rot End Date:" & slEndDate

                    gUnpackDate imPurgedDate(0), imPurgedDate(1), slStartDate
                    slMsg2 = slMsg2 & ", Purge Date:" & slStartDate
                    
                    gLogMsg slMsg2, "SetCopyInventoryDates.txt", False
                Else
                    tmSortCif(ilLoop).tCif.iPurgeDate(0) = 0
                    tmSortCif(ilLoop).tCif.iPurgeDate(1) = 0
                End If
                ilRet = btrUpdate(hmCif, tmSortCif(ilLoop).tCif, imCifRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Purge
            On Error GoTo 0
        End If
    Next ilLoop
    Screen.MousePointer = vbDefault  'Wait
    mSaveRec = True
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
    Dim ilLoop As Integer
    ilAltered = False
    For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
        If tmSortCif(ilLoop).iChg Then
            ilAltered = True
        End If
    Next ilLoop
    If ilAltered = True Then
        If ilAsk Then
            slMess = "Update inventory"
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
'*      Procedure Name:mShowInvInfo                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show inventory information     *
'*                                                     *
'*******************************************************
Private Sub mShowInvInfo(ilLbc As Integer, flY As Single)
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim slName As String
    Dim ilLoop As Integer
    Dim ilButtonIndex As Integer
    Dim slDate As String
    ilButtonIndex = imButtonIndex

    If ilLbc = DRAGASK Then
        If (imButtonIndex < 0) And (imButtonIndex > lbcAskDisposition.ListCount - 1) Then
            imButtonIndex = -1
            plcInvInfo.Visible = False
            Exit Sub
        End If
        slName = lbcAskDisposition.List(ilButtonIndex)
    Else
        If (imButtonIndex < 0) And (imButtonIndex > lbcPurgeDisposition.ListCount - 1) Then
            imButtonIndex = -1
            plcInvInfo.Visible = False
            Exit Sub
        End If
       slName = lbcPurgeDisposition.List(ilButtonIndex)
    End If
    For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
        slStr = Trim$(tmSortCif(ilLoop).sInvName) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
        If slStr = slName Then
            gUnpackDate tmSortCif(ilLoop).tCif.iRotStartDate(0), tmSortCif(ilLoop).tCif.iRotStartDate(1), slStartDate
            gUnpackDate tmSortCif(ilLoop).tCif.iRotEndDate(0), tmSortCif(ilLoop).tCif.iRotEndDate(1), slEndDate
            If slStartDate <> "" Then
                lacInvInfo(0).Caption = Trim$(tmSortCif(ilLoop).sInvName) & " " & "Rotation Date: " & slStartDate & " - " & slEndDate
            Else
                gUnpackDate tmCif.iDateEntrd(0), tmCif.iDateEntrd(1), slDate
                lacInvInfo(0).Caption = Trim$(tmSortCif(ilLoop).sInvName) & " " & "Rotation: None; Enter: " & slDate
            End If
            If tmSortCif(ilLoop).tCif.lcpfCode > 0 Then
                tmCpfSrchKey.lCode = tmSortCif(ilLoop).tCif.lcpfCode
                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    lacInvInfo(1).Caption = "Product: " & Trim$(tmCpf.sName)
                    lacInvInfo(2).Caption = "ISCI Code: " & Trim$(tmCpf.sISCI)
                    lacInvInfo(3).Caption = "Creative Title: " & Trim$(tmCpf.sCreative)
                Else
                    lacInvInfo(1).Caption = ""
                    lacInvInfo(2).Caption = ""
                    lacInvInfo(3).Caption = ""
                End If
            Else
                lacInvInfo(1).Caption = ""
                lacInvInfo(2).Caption = ""
                lacInvInfo(3).Caption = ""
            End If
        End If
    Next ilLoop
    If ilLbc = DRAGASK Then
        plcInvInfo.Move plcAskDisposition.Left, plcAskDisposition.Top + lbcAskDisposition.Top + lbcAskDisposition.height '+ flY + 120
    Else
        plcInvInfo.Move plcPurgeDisposition.Left + plcPurgeDisposition.Width - plcInvInfo.Width, plcPurgeDisposition.Top + lbcPurgeDisposition.Top + lbcPurgeDisposition.height ' + flY + 120
    End If
    DoEvents
    If ilLbc = DRAGASK Then
        If (imButtonIndex < 0) And (imButtonIndex > lbcAskDisposition.ListCount - 1) Then
            imButtonIndex = -1
            plcInvInfo.Visible = False
            Exit Sub
        End If
    Else
        If (imButtonIndex < 0) And (imButtonIndex > lbcPurgeDisposition.ListCount - 1) Then
            imButtonIndex = -1
            plcInvInfo.Visible = False
            Exit Sub
        End If
    End If
    plcInvInfo.Visible = True
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
    Unload Purge
    igManUnload = NO
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
                edcRotEndDate.Text = Format$(llDate, "m/d/yy")
                edcRotEndDate.SelStart = 0
                edcRotEndDate.SelLength = Len(edcRotEndDate.Text)
                edcRotEndDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcRotEndDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcAskDisposition_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcAskDisposition_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub plcCopyInv_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcCopyInv_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub plcPurgeDisposition_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcPurgeDisposition_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub plcSort_DragDrop(Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub rbcSort_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSort(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim slName As String
    If Value Then
        If Index = 0 Then               'rot end date, sort date/advt name/cart
            For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
                tmSortCif(ilLoop).sKey = tmSortCif(ilLoop).sDate & "|" & tmSortCif(ilLoop).sAdvtName & "|" & tmSortCif(ilLoop).sInvName
            Next ilLoop
        ElseIf Index = 1 Then       'advt, sort advt name/date/cart
            For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
                tmSortCif(ilLoop).sKey = tmSortCif(ilLoop).sAdvtName & "|" & tmSortCif(ilLoop).sDate & "|" & tmSortCif(ilLoop).sInvName
            Next ilLoop
        ElseIf Index = 2 Then           'ISCI, sort isci/advt name/date
            For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
                tmSortCif(ilLoop).sKey = tmSortCif(ilLoop).sISCI & "|" & tmSortCif(ilLoop).sAdvtName & "|" & tmSortCif(ilLoop).sDate
            Next ilLoop
        Else                        'cart #, sort cart/advt name/date
            For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
                tmSortCif(ilLoop).sKey = tmSortCif(ilLoop).sInvName & "|" & tmSortCif(ilLoop).sAdvtName & "|" & tmSortCif(ilLoop).sDate
            Next ilLoop
        End If
        lbcAskDisposition.Clear
        lbcPurgeDisposition.Clear
        ilUpper = UBound(tmSortCif)
        If ilUpper > 0 Then
            ArraySortTyp fnAV(tmSortCif(), 0), ilUpper, 0, LenB(tmSortCif(0)), 0, LenB(tmSortCif(0).sKey), 0
        End If
        For ilLoop = 0 To UBound(tmSortCif) - 1 Step 1
            If rbcSort(2).Value Then                'sort the ISCI only if sorting by ISCI
                slName = Trim$(tmSortCif(ilLoop).sISCI) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
            Else
                slName = Trim$(tmSortCif(ilLoop).sInvName) & " " & Trim$(tmSortCif(ilLoop).sAdvtName)
            End If
            If tmSortCif(ilLoop).tCif.sCartDisp = "P" Then
                lbcPurgeDisposition.AddItem slName
            Else
                lbcAskDisposition.AddItem slName
            End If
        Next ilLoop
    End If
End Sub
Private Sub rbcSort_DragDrop(Index As Integer, Source As control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If imSelectDelay Then
        imSelectDelay = False
        mCbcAdvtChange
    Else
    End If
End Sub
Private Sub tmcDrag_Timer()
    Dim ilListIndex As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            If imDragButton <> 1 Then
                Exit Sub
            End If
            Select Case imDragSrce
                Case DRAGASK
                    ilListIndex = (fmDragY \ imLbcHeight) + lbcAskDisposition.TopIndex
                    If (ilListIndex >= 0) And (ilListIndex <= lbcAskDisposition.ListCount - 1) Then
                        lbcAskDisposition.DragIcon = IconTraf!imcIconDrag.DragIcon
                        imDragIndexSrce = ilListIndex
                        lbcAskDisposition.Drag vbBeginDrag
                    Else
                        lbcAskDisposition.ListIndex = -1
                    End If
                Case DRAGPURGE
                    ilListIndex = (fmDragY \ imLbcHeight) + lbcPurgeDisposition.TopIndex
                    If (ilListIndex >= 0) And (ilListIndex <= lbcPurgeDisposition.ListCount - 1) Then
                        lbcPurgeDisposition.DragIcon = IconTraf!imcIconDrag.DragIcon
                        imDragIndexSrce = ilListIndex
                        lbcPurgeDisposition.Drag vbBeginDrag
                    Else
                        lbcPurgeDisposition.ListIndex = -1
                    End If
            End Select
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub plcSort_Paint()
    plcSort.CurrentX = 0
    plcSort.CurrentY = 0
    plcSort.Print "Sort by"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Copy Purge"
End Sub
Private Sub plcPurgeDisposition_Paint()
    plcPurgeDisposition.CurrentX = 0
    plcPurgeDisposition.CurrentY = 0
    plcPurgeDisposition.Print "To Be Purged"
End Sub
Private Sub plcAskDisposition_Paint()
    plcAskDisposition.CurrentX = 0
    plcAskDisposition.CurrentY = 0
    plcAskDisposition.Print "To Be Saved"
End Sub

' TTP 10807 - JD 08-15-23
'
Public Function gSetCopyDates(slDate As String, slMediaCode As String, ilMcfCode As Integer)
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String
    Dim ilUpper As Integer
    Dim hlMcf As Integer        'MCF Handle
    Dim ilMcfRecLen As Integer  'MCF record length
    Dim tlMcfSrchKey As INTKEY0 'MCF key record image
    Dim tlMcf As MCF            'MCF record image
    'Dim ilMcfCode As Integer
    Dim llChgCount As Long
    Dim llTotalCount As Long
    Dim hlCrf As Integer
    Dim hlCif As Integer
    Dim hlCnf As Integer
    Dim hlCuf As Integer

    gSetCopyDates = True
    
    On Error GoTo Err1
    hlCrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    
    hlCif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    
    hlCnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    
    hlCuf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlCuf, "", sgDBPath & "Cuf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    
    'hlMcf = CBtrvTable(ONEHANDLE)
    'ilRet = btrOpen(hlMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    
    On Error GoTo Err2

    ReDim tlCifInfo(0 To 0) As CIFINFO
    tlCifInfo(0).lFirst = -1
    ReDim tlCrfDates(0 To 0) As CRFDATES
    tlCrfDates(0).lNext = -1
    llChgCount = 0
    llTotalCount = 0
    
'    ilMcfCode = -1
'    ilMcfRecLen = Len(tlMcf)
'    ilRet = btrGetFirst(hlMcf, tlMcf, ilMcfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'    Do While (ilRet = BTRV_ERR_NONE)
'        If Trim$(tlMcf.sName) = Trim$(slMediaCode) Then
'            ilMcfCode = tlMcf.iCode
'            Exit Do
'        End If
'        ilRet = btrGetNext(hlMcf, tlMcf, ilMcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'    Loop
'    If ilMcfCode = -1 Then
'        gLogMsg "gSetCopyDates: Error: Unable to find media code for " & slMediaCode, "SetCopyInventoryDates.txt", False
'        GoTo Err2
'    End If
    
    gLogMsg "Gathering Copy Inventory for Media Code " & slMediaCode & ": Started", "SetCopyInventoryDates.txt", False

    gBuildCifArray hlCif, ilMcfCode, tlCifInfo()
    
    gLogMsg "Gathering Copy Inventory: Completed", "SetCopyInventoryDates.txt", False
    gLogMsg "Gathering Rotations active as of " & slDate & ": Started", "SetCopyInventoryDates.txt", False
    
    gBuildRotDates hlCrf, hlCnf, slDate, tlCifInfo(), tlCrfDates()
    
    gLogMsg "Gathering Rotations: Completed", "SetCopyInventoryDates.txt", False
    gLogMsg "Create Copy Usage: Started", "SetCopyInventoryDates.txt", False
    
    gCreateCufDates hlCif, hlCuf, tlCifInfo(), tlCrfDates()
    
    gLogMsg "Create Copy Usage: Completed", "SetCopyInventoryDates.txt", False
    gLogMsg "Updating Copy Inventory: Started", "SetCopyInventoryDates.txt", False
    
    gUpdateCifDates hlCif, tlCifInfo(), tlCrfDates(), llChgCount, llTotalCount, True

    gLogMsg "Updating Copy Inventory: Completed", "SetCopyInventoryDates.txt", False
    gLogMsg "Updated " & llChgCount & " Copy Inventory dates, out of " & llTotalCount & " pieces of Copy", "SetCopyInventoryDates.txt", False
    
Finally:
    On Error GoTo 0
    ilRet = btrClose(hlCrf)
    btrDestroy hlCrf
    ilRet = btrClose(hlCif)
    btrDestroy hlCif
    ilRet = btrClose(hlCnf)
    btrDestroy hlCnf
    ilRet = btrClose(hlCuf)
    btrDestroy hlCuf
    ilRet = btrClose(hlMcf)
    btrDestroy hlMcf
    Exit Function
    
Err1:
    gLogMsg "gSetCopyDates: Unable to open tables", "SetCopyInventoryDates.txt", False
    gSetCopyDates = False
    Exit Function
    
Err2:
    gSetCopyDates = False
    gLogMsg "gSetCopyDates: did not complete properly ", "SetCopyInventoryDates.txt", False
    GoTo Finally
   
End Function

' TTP 10807 - JD 08-15-23
' New function to detect if there is actually a problem when the cifDate is blank.
Private Function RotDatesDetected(cifCode As Long, slDate As String, iMediaCode As Integer)
    Dim slSQLQuery As String
    
    RotDatesDetected = False
    On Error GoTo Err1
    slSQLQuery = "Select count(1) as Total "
    slSQLQuery = slSQLQuery & "from CIF_Copy_Inventory  "
    slSQLQuery = slSQLQuery & "left outer join CNF_Copy_Instruction on cifCode = cnfcifCode "
    slSQLQuery = slSQLQuery & "left outer join CRF_Copy_Rot_Header on cnfcrfcode = crfcode "
    slSQLQuery = slSQLQuery & "Where CIF_Copy_Inventory.cifCode = " & cifCode & " "
    slSQLQuery = slSQLQuery & "And crfEndDate >= '" & Format(slDate, sgSQLDateForm) & "' "
    slSQLQuery = slSQLQuery & "And cifmcfCode = " & iMediaCode & " "
    slSQLQuery = slSQLQuery & "And cifRotEndDate is null  "
    slSQLQuery = slSQLQuery & "And crfState <> 'D' "
    
    Set rmRecSet = gSQLSelectCall(slSQLQuery)
    If rmRecSet!Total > 0 Then
        RotDatesDetected = True
    End If
    Exit Function
    
Err1:
    gLogMsg "RotDatesDetected: failed ", "SetCopyInventoryDates.txt", False
End Function

'Private Function Test()
'    Dim ilRet As Integer
'    Dim hlRlf As Integer
'    Dim lmLock1RecCode As Integer
'
'    hlRlf = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hlRlf, "", sgDBPath & "Rlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'
'    lmLock1RecCode = gCreateLockRec(hlRlf, "Y", "S", 1, False, "")
'    If lmLock1RecCode = 0 Then
'        ilRet = MsgBox("Unable to save copy at this time. Try again in 10 seconds.", vbOKOnly + vbInformation, "Block")
'        Exit Function
'    End If
'
'    If lmLock1RecCode <> 0 Then
'        ilRet = gDeleteLockRec_ByType(hlRlf, "Y", 1)
'    End If
'    ilRet = btrClose(hlRlf)
'    btrDestroy hlRlf
'
'End Function
