VERSION 5.00
Begin VB.Form Links 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5265
   ClientLeft      =   1155
   ClientTop       =   1305
   ClientWidth     =   7320
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
   ScaleHeight     =   5265
   ScaleWidth      =   7320
   Begin VB.ListBox lbcAiring 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   135
      MultiSelect     =   2  'Extended
      TabIndex        =   36
      Top             =   225
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lbcSelling 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   180
      MultiSelect     =   2  'Extended
      TabIndex        =   35
      Top             =   405
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.ListBox lbcDEName 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   345
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   1035
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Links.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
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
            TabIndex        =   21
            Top             =   405
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
         TabIndex        =   17
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
         TabIndex        =   19
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
         Left            =   315
         TabIndex        =   18
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Height          =   285
      Left            =   5475
      TabIndex        =   29
      Top             =   4905
      Width           =   1140
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   780
      TabIndex        =   22
      Top             =   4905
      Width           =   1140
   End
   Begin VB.CommandButton cmcLinksDef 
      Appearance      =   0  'Flat
      Caption         =   "Define &Links"
      Height          =   285
      Left            =   2190
      TabIndex        =   23
      Top             =   4905
      Width           =   3000
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   615
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox plcLinkDates 
      ForeColor       =   &H00000000&
      Height          =   1185
      Left            =   1260
      ScaleHeight     =   1125
      ScaleWidth      =   4725
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   210
      Width           =   4785
      Begin VB.CommandButton cmcEndDate 
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
         Left            =   4305
         Picture         =   "Links.frx":2E1A
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   780
         Width           =   195
      End
      Begin VB.TextBox edcEndDate 
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
         Left            =   3270
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   795
         Width           =   1020
      End
      Begin VB.TextBox edcStartDate 
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
         Left            =   1005
         MaxLength       =   10
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   780
         Width           =   1020
      End
      Begin VB.CommandButton cmcStartDate 
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
         Left            =   2040
         Picture         =   "Links.frx":2F14
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   780
         Width           =   195
      End
      Begin VB.PictureBox plcLinks 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   4545
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   405
         Width           =   4545
         Begin VB.OptionButton rbcLinks 
            Caption         =   "Engineering"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   3150
            TabIndex        =   9
            Top             =   0
            Width           =   1350
         End
         Begin VB.OptionButton rbcLinks 
            Caption         =   "Delivery"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2175
            TabIndex        =   8
            Top             =   0
            Width           =   1005
         End
         Begin VB.OptionButton rbcLinks 
            Caption         =   "Selling to Airing"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   555
            TabIndex        =   7
            Top             =   -15
            Width           =   1635
         End
      End
      Begin VB.PictureBox plcDay 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   2175
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   75
         Width           =   2175
         Begin VB.OptionButton rbcDay 
            Caption         =   "Su"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   1545
            TabIndex        =   5
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton rbcDay 
            Caption         =   "Sa"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   990
            TabIndex        =   4
            Top             =   0
            Width           =   540
         End
         Begin VB.OptionButton rbcDay 
            Caption         =   "M-F"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   0
            Width           =   600
         End
      End
      Begin VB.Label lacEndDate 
         Appearance      =   0  'Flat
         Caption         =   "End Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2460
         TabIndex        =   13
         Top             =   765
         Width           =   765
      End
      Begin VB.Label lacStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Start Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   765
         Width           =   825
      End
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
      Left            =   4725
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4305
      Visible         =   0   'False
      Width           =   525
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
      Left            =   5640
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4755
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcFeedCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6075
      Sorted          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   1140
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
      Left            =   5610
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4770
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox plcACVehicle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   225
      ScaleHeight     =   3225
      ScaleWidth      =   6840
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   6900
      Begin VB.ListBox lbcACVeh 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   150
         TabIndex        =   31
         Top             =   405
         Width           =   6600
      End
   End
   Begin VB.PictureBox plcSAVehicle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   240
      ScaleHeight     =   3225
      ScaleWidth      =   6840
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   6900
      Begin VB.ListBox lbcSAVeh 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   150
         TabIndex        =   34
         Top             =   405
         Width           =   6600
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   4845
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Status"
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
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Links"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Links.frm on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'**********************************************************
'                LINKS MODULE DEFINITIONS
'
'   Created : 4/17/94       By : D. Hannifan
'   Modified :              By :
'
'**********************************************************
Option Explicit
Option Compare Text
'Btrieve file variables
Dim hmLcf As Integer            'Log calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim tmLcfSrchKey As LCFKEY0     'LCF key record image
Dim imLcfRecLen As Integer         'LCF record length
Dim hmVLF As Integer            'Vehicle link file handle
Dim tmVlf As VLF                'VLF record image
Dim imVlfRecLen As Integer        'VLF record length
Dim tmVlfSrchKey0 As VLFKEY0
Dim tmVlfSrchKey1 As VLFKEY1
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer         'VEF record length
Dim hmDlf As Integer            'Delivery Vehicle link file handle
Dim hmEgf As Integer            'Delivery Vehicle link file handle
Dim tmDlf As DLF                'DLF record image
Dim tmDlfSrchKey As DLFKEY0            'DLF record image
Dim imDlfRecLen As Integer        'VLF record length
'Module Status Flags
Dim imFirstActivate As Integer
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim imBypassFocus As Integer    'Bypass gotfocus
Dim imLcfPending As Integer     'LCF Status: Pending = 1, Current = 0
Dim imVlfPending As Integer     'VLF Status: Pending = 1, Current = 0
Dim imDateSetFlag As Integer    'Valid date found for pending LCF or VLF = True
Dim imDateCode As Integer       'Date code for search 0=M-F, 5=Sa, 6=Su
Dim smVLFPendingDate As String
Dim imIgnoreClick As Integer    'Used for clearing selection between lbcACVeh
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer    'Calendar type
Dim imDateIndex As Integer  '0=Start Date; 1=End Date
Dim imFirstTime As Integer
Dim imShowHelpMsg As Integer    'True=Show help message; False=Ignore help message system
Dim imGroupNo() As Integer
Dim bmFirstCallToVpfFind As Boolean
'*******************************************************
'
'       Procedure Name : cmcCalDn_Click
'
'       Created : ?             By : D. Levine
'       Modified :              By :
'
'       Comments : Reset calendar month
'
'*******************************************************
'
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    If imDateIndex = 0 Then
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
        edcStartDate.SetFocus
    Else
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
        edcEndDate.SetFocus
    End If
End Sub
'*******************************************************
'
'       Procedure Name : cmcCalUp_Click
'
'       Created : ?             By : D. Levine
'       Modified :              By :
'
'       Comments : Reset calendar month
'
'*******************************************************
'
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    If imDateIndex = 0 Then
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
        edcStartDate.SetFocus
    Else
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
        edcEndDate.SetFocus
    End If
End Sub
Private Sub cmcDone_Click()
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcEndDate_Click()
    If imDateIndex = 0 Then
        plcCalendar.Visible = True
    Else
        plcCalendar.Visible = Not plcCalendar.Visible
    End If
    imDateIndex = 1
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
End Sub
Private Sub cmcEndDate_GotFocus()
    Dim slStr As String
    gCtrlGotFocus ActiveControl
    'imDateIndex = 1
    plcCalendar.Move plcLinkDates.Left + cmcEndDate.Left + cmcEndDate.Width - plcCalendar.Width, plcLinkDates.Top + edcEndDate.Top + edcEndDate.Height
    slStr = edcEndDate.Text
    If Not gValidDate(slStr) Then
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
'*******************************************************
'
'       Procedure Name : cmcLinksDef_Click
'
'       Created : ?             By : D. Levine
'       Modified : 4/17/94      By : D. Hannifan
'
'       Comments : Reset calendar month
'
'*******************************************************
'
Private Sub cmcLinksDef_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilDay                         ilPending                                               *
'******************************************************************************************

    Dim llSellDate As Long
    Dim llAirDate As Long
    Dim llEffDate As Long
    Dim llEndDate As Long
    Dim slDate As String
    Dim llDate As Long
    Dim ilLoop As Integer
    Dim ilVeh As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilVefCode As Integer
    Dim llTestDate As Long
    Dim sLCP As String
    Dim ilVpfIndex As Integer
    Dim slDay As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slName As String
    Dim ilGroupNo As Integer
    slStartDate = Trim$(edcStartDate.Text)
    If Not gValidDate(slStartDate) Then
        Beep
        edcStartDate.SetFocus
        Exit Sub
    End If
    llDate = gDateValue(slStartDate)
    If rbcDay(0).Value Then
        If rbcLinks(0).Value Then
            If (gWeekDayLong(llDate) <> 0) Then
                Beep
                edcStartDate.SetFocus
                Exit Sub
            End If
        Else
            If gWeekDayLong(llDate) > 4 Then
                Beep
                edcStartDate.SetFocus
                Exit Sub
            End If
        End If
    ElseIf rbcDay(1).Value Then
        If rbcLinks(0).Value Then
            If (gWeekDayLong(llDate) <> 5) Then
                Beep
                edcStartDate.SetFocus
                Exit Sub
            End If
        Else
            If gWeekDayLong(llDate) <> 5 Then
                Beep
                edcStartDate.SetFocus
                Exit Sub
            End If
        End If
    Else
        If rbcLinks(0).Value Then
            If (gWeekDayLong(llDate) <> 6) Then
                Beep
                edcStartDate.SetFocus
                Exit Sub
            End If
        Else
            If gWeekDayLong(llDate) <> 6 Then
                Beep
                edcStartDate.SetFocus
                Exit Sub
            End If
        End If
    End If
    slEndDate = Trim$(edcEndDate.Text)
    If (slEndDate <> "") And (Not rbcLinks(0).Value) Then
        If Not gValidDate(slEndDate) Then
            Beep
            edcEndDate.SetFocus
            Exit Sub
        End If
        If gDateValue(slEndDate) < gDateValue(slStartDate) Then
            MsgBox "End Date Must be on or after Start Date", vbExclamation, "Links"
            edcEndDate.SetFocus
            Exit Sub
        End If
        llDate = gDateValue(slEndDate)
        If rbcDay(0).Value Then
            If rbcLinks(0).Value Then
                If (gWeekDayLong(llDate) <> 6) Then
                    Beep
                    edcEndDate.SetFocus
                    Exit Sub
                End If
            Else
                If gWeekDayLong(llDate) > 4 Then
                    Beep
                    edcEndDate.SetFocus
                    Exit Sub
                End If
            End If
        ElseIf rbcDay(1).Value Then
            If rbcLinks(0).Value Then
                If (gWeekDayLong(llDate) <> 4) Then
                    Beep
                    edcEndDate.SetFocus
                    Exit Sub
                End If
            Else
                If gWeekDayLong(llDate) <> 5 Then
                    Beep
                    edcEndDate.SetFocus
                    Exit Sub
                End If
            End If
        Else
            If rbcLinks(0).Value Then
                If (gWeekDayLong(llDate) <> 5) Then
                    Beep
                    edcEndDate.SetFocus
                    Exit Sub
                End If
            Else
                If gWeekDayLong(llDate) <> 6 Then
                    Beep
                    edcEndDate.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    If Not rbcLinks(0).Value Then
        ilVeh = lbcACVeh.ListIndex
        slName = lbcACVeh.List(ilVeh)
        ilPos = InStr(slName, ": ")
        If ilPos > 0 Then
            ilPos = InStr(slName, "-TFN")
            If ilPos > 0 Then
                ilSpace = ilPos - 6
                Do While Mid$(slName, ilSpace, 1) <> " "
                    ilSpace = ilSpace - 1
                    If ilSpace <= 0 Then
                        Exit Do
                    End If
                Loop
                If ilSpace > 0 Then
                    slDate = Mid$(slName, ilSpace + 1, ilPos - ilSpace - 1)
                    If gDateValue(slStartDate) < gDateValue(slDate) Then
                        ilRet = MsgBox("This will remove links defined from " & slDate, vbOKCancel + vbExclamation, "Links")
                        If ilRet = vbCancel Then
                            cmcDone.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Else
        If lbcSAVeh.ListIndex < 0 Then
            ilRet = MsgBox("Select a Group to Define Links", vbOKOnly + vbExclamation, "Links")
            lbcSAVeh.SetFocus
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    If rbcLinks(0).Value Then   'Define links
        'If Not gWinRoom(igNoExeWinRes(LINKSDEFEXE)) Then
        '    Screen.MousePointer = vbDefault
        '    Exit Sub
        'End If
        slDate = Trim$(edcStartDate.Text)   'Store Effective Date
        llEffDate = gDateValue(slDate)
        If Trim$(edcEndDate.Text) <> "" Then
            slDate = Trim$(edcEndDate.Text)   'Store Effective Date
            llEndDate = gDateValue(slDate)
        Else
            llEndDate = 0
        End If
        If imLcfPending Then 'Set LinksDef module Status to Pending
            lacStatus.Caption = "P"
            sLCP = "P"
        Else
            lacStatus.Caption = "C"  'Current
            sLCP = "C"
        End If
        If imVlfPending Then 'Set LinksDef module Status to Pending
            lacStatus.Caption = lacStatus.Caption & "\" & "P"
        Else
            lacStatus.Caption = lacStatus.Caption & "\" & "C"  'Current
        End If
        'Select all vehicle combinations
        sLCP = "C"
        llSellDate = 0
        ilGroupNo = imGroupNo(lbcSAVeh.ListIndex)
        For ilLoop = 0 To lbcSelling.ListCount - 1 Step 1
            lbcSelling.Selected(ilLoop) = False
            slNameCode = tgUserVehicle(ilLoop).sKey 'lbcVehName.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            If bmFirstCallToVpfFind Then
                ilVpfIndex = gVpfFind(Links, ilVefCode)
                bmFirstCallToVpfFind = False
            Else
                ilVpfIndex = gVpfFindIndex(ilVefCode)
            End If
            If tgVpf(ilVpfIndex).iSAGroupNo = ilGroupNo Then
                lbcSelling.Selected(ilLoop) = True
            End If
        Next ilLoop
        For ilLoop = 0 To lbcAiring.ListCount - 1 Step 1
            lbcAiring.Selected(ilLoop) = False
            slNameCode = tgVehicle(ilLoop).sKey 'lbcVehMName.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            If bmFirstCallToVpfFind Then
                ilVpfIndex = gVpfFind(Links, ilVefCode)
                bmFirstCallToVpfFind = False
            Else
                ilVpfIndex = gVpfFindIndex(ilVefCode)
            End If
            If tgVpf(ilVpfIndex).iSAGroupNo = ilGroupNo Then
                lbcAiring.Selected(ilLoop) = True
            End If
        Next ilLoop
'Remove check LCF pending date as they might need to alter Links not related to pending dates
'        'Find the earliest of the latest date in the selling vehicles
'        For ilLoop = 0 To lbcSelling.ListCount - 1 Step 1
'            If lbcSelling.Selected(ilLoop) Then 'Selected vehicle found
'                slNameCode = tgUserVehicle(ilLoop).sKey 'lbcVehName.List(ilLoop)
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                ilVefCode = Val(slCode)
'                llTestDate = gGetLatestLCFDate(hmLcf, "O", sLCP, ilVefCode)
'                If llTestDate <= 0 Then
'                    If sLCP = "C" Then
'                        Screen.MousePointer = vbDefault
'                        ilRet = gParseItem(slNameCode, 1, "\", slCode)
'                        ilRet = gParseItem(slCode, 3, "|", slCode)
'                        MsgBox slCode & " must have a calendar day defined prior to creating links", vbExclamation, "Links"
'                        'lbcSelling.Enabled = True
'                        'lbcSelling.SetFocus
'                        Exit Sub
'                    Else
'                        llTestDate = gGetLatestLCFDate(hmLcf, "O", "C", ilVefCode)
'                        If llTestDate <= 0 Then
'                            Screen.MousePointer = vbDefault
'                            ilRet = gParseItem(slNameCode, 1, "\", slCode)
'                            ilRet = gParseItem(slCode, 3, "|", slCode)
'                            MsgBox slCode & " must have a calendar day defined prior to creating links", vbExclamation, "Links"
'                            'lbcSelling.Enabled = True
'                            'lbcSelling.SetFocus
'                            Exit Sub
'                        End If
'                    End If
'                End If
'                'ilDay = gWeekDayLong(llTestDate)
'                'If rbcDay(0).Value Then 'Mo-Fr
'                '    llTestDate = llTestDate - ilDay 'Back up to monday
'                'ElseIf rbcDay(1).Value Then 'Sat
'                '    If ilDay = 6 Then
'                '        llTestDate = llTestDate - 1
'                '    ElseIf ilDay <= 4 Then
'                '        llTestDate = llTestDate - ilDay - 2
'                '    End If
'                'Else    'Sunday
'                '    If ilDay <= 5 Then
'                '        llTestDate = llTestDate - ilDay - 1
'                '    End If
'                'End If
'                ''If llSellDate = 0 Then
'                ''    llSellDate = llTestDate
'                ''ElseIf llTestDate < llSellDate Then
'                ''    llSellDate = llTestDate
'                ''End If
'                llTestDate = mPendingDate(ilVefCode)
'                If (llSellDate = 0) And (llTestDate > 0) Then
'                    llSellDate = llTestDate
'                ElseIf llTestDate > 0 Then
'                    If llTestDate < llSellDate Then
'                        llSellDate = llTestDate
'                    End If
'                End If
'            End If
'        Next ilLoop
        'Find the latest of the latest airing vehicle date
        'Effective date has to be equal or after latest date
'        llAirDate = 0
'        For ilLoop = 0 To lbcAiring.ListCount - 1 Step 1
'            If lbcAiring.Selected(ilLoop) Then 'Selected vehicle found
'                slNameCode = tgVehicle(ilLoop).sKey 'lbcVehMName.List(ilLoop)
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                ilVefCode = Val(slCode)
'                llTestDate = gGetLatestLCFDate(hmLcf, "O", sLCP, ilVefCode)
'                If llTestDate <= 0 Then
'                    If sLCP = "C" Then
'                        Screen.MousePointer = vbDefault
'                        ilRet = gParseItem(slNameCode, 1, "\", slCode)
'                        ilRet = gParseItem(slCode, 3, "|", slCode)
'                        MsgBox slCode & " must have a calendar day defined prior to creating links", vbExclamation, "Links"
'                        'lbcAiring.Enabled = True
'                        'lbcAiring.SetFocus
'                        Exit Sub
'                    Else
'                        llTestDate = gGetLatestLCFDate(hmLcf, "O", "C", ilVefCode)
'                        If llTestDate <= 0 Then
'                            Screen.MousePointer = vbDefault
'                            ilRet = gParseItem(slNameCode, 1, "\", slCode)
'                            ilRet = gParseItem(slCode, 3, "|", slCode)
'                            MsgBox slCode & " must have a calendar day defined prior to creating links", vbExclamation, "Links"
'                            'lbcAiring.Enabled = True
'                            'lbcAiring.SetFocus
'                            Exit Sub
'                        End If
'                    End If
'                End If
'                'ilDay = gWeekDayLong(llTestDate)
'                'If rbcDay(0).Value Then 'Mo-Fr
'                '    llTestDate = llTestDate - ilDay 'Back up to monday
'                'ElseIf rbcDay(1).Value Then 'Sat
'                '    If ilDay = 6 Then
'                 '       llTestDate = llTestDate - 1
'                '    ElseIf ilDay <= 4 Then
'                 '       llTestDate = llTestDate + 5 - ilDay
'                '    End If
'                'Else    'Sunday
'                '    If ilDay <= 5 Then
'                '        llTestDate = llTestDate + 6 - ilDay
'                '    End If
'                'End If
'                ''If llAirDate = 0 Then
'                ''    llAirDate = llTestDate
'                ''ElseIf llTestDate > llAirDate Then
'                ''    llAirDate = llTestDate
'                ''End If
'                llTestDate = mPendingDate(ilVefCode)
'                If (llAirDate = 0) And (llTestDate > 0) Then
'                    llAirDate = llTestDate
'                ElseIf llTestDate > 0 Then
'                    If llTestDate < llAirDate Then
'                        llAirDate = llTestDate
'                    End If
'                End If
'            End If
'        Next ilLoop
'        'If llEffDate > llSellDate Then
'        '    Screen.MousePointer = vbDefault
'        '    slDate = Format$(llSellDate + 1, "m/d/yy")
'        '    MsgBox "Date must be prior to " & slDate, vbExclamation, "Links"
'        '    edcStartDate.Enabled = True
'        '    edcStartDate.SetFocus
'        '    Exit Sub
'        'End If
'
'This is being removed as defining end date removes this restriction.
'Also, since the avail structure is obtained from the start date
'      if programming not defined, then it will show empty in the link screen.
'        If (llEffDate < llSellDate) And (llSellDate > 0) Then
'            Screen.MousePointer = vbDefault
'            slDate = Format$(llSellDate - 1, "m/d/yy")
'            MsgBox "Date must be after " & slDate, vbExclamation, "Links"
'            edcStartDate.Enabled = True
'            edcStartDate.SetFocus
'            Exit Sub
'        End If
'        If (llEffDate < llAirDate) And (llAirDate > 0) Then
'            Screen.MousePointer = vbDefault
'            slDate = Format$(llAirDate - 1, "m/d/yy")
'            MsgBox "Date must be after " & slDate, vbExclamation, "Links"
'            edcStartDate.Enabled = True
'            edcStartDate.SetFocus
'            Exit Sub
'        End If
'        If llEndDate = 0 Then
'            'Find the earliest of the latest date in the selling vehicles
'            For ilLoop = 0 To lbcSelling.ListCount - 1 Step 1
'                If lbcSelling.Selected(ilLoop) Then 'Selected vehicle found
'                    slNameCode = tgUserVehicle(ilLoop).sKey 'lbcVehName.List(ilLoop)
'                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                    ilVefCode = Val(slCode)
'                    If rbcDay(0).Value Then 'Mo-Fr
'                        ilDay = 0
'                    ElseIf rbcDay(1).Value Then 'Sat
'                        ilDay = 6
'                    Else    'Sunday
'                        ilDay = 7
'                    End If
'                    llTestDate = gGetLatestVLFDate(hmVlf, "S", ilVefCode, ilDay, ilPending)
'                    If llTestDate > 0 Then
'                        ilDay = gWeekDayLong(llTestDate)
'                        If rbcDay(0).Value Then 'Mo-Fr
'                            llTestDate = llTestDate - ilDay 'Back up to monday
'                        ElseIf rbcDay(1).Value Then 'Sat
'                            If ilDay = 6 Then
'                                llTestDate = llTestDate - 1
'                            ElseIf ilDay <= 4 Then
'                                llTestDate = llTestDate - ilDay - 2
'                            End If
'                        Else    'Sunday
'                            If ilDay <= 5 Then
'                                llTestDate = llTestDate - ilDay - 1
'                            End If
'                        End If
'                        If ilPending Then
'                            If llTestDate <> llEffDate Then
'                                Screen.MousePointer = vbDefault
'                                slDate = Format$(llTestDate, "m/d/yy")
'                                MsgBox "Date must be " & slDate, vbExclamation, "Links"
'                                edcStartDate.Enabled = True
'                                edcStartDate.SetFocus
'                                Exit Sub
'                            End If
'                        Else
'                            If llTestDate > llEffDate Then
'                                Screen.MousePointer = vbDefault
'                                slDate = Format$(llTestDate - 1, "m/d/yy")
'                                MsgBox "Date must be after " & slDate, vbExclamation, "Links"
'                                edcStartDate.Enabled = True
'                                edcStartDate.SetFocus
'                                Exit Sub
'                            End If
'                        End If
'                    End If
'                End If
'            Next ilLoop
'            'Find the latest of the latest airing vehicle date
'            'Effective date has to be equal or after latest date
'            For ilLoop = 0 To lbcAiring.ListCount - 1 Step 1
'                If lbcAiring.Selected(ilLoop) Then 'Selected vehicle found
'                    slNameCode = tgVehicle(ilLoop).sKey 'lbcVehMName.List(ilLoop)
'                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                    ilVefCode = Val(slCode)
'                    If rbcDay(0).Value Then 'Mo-Fr
'                        ilDay = 0
'                    ElseIf rbcDay(1).Value Then 'Sat
'                        ilDay = 6
'                    Else    'Sunday
'                        ilDay = 7
'                    End If
'                    llTestDate = gGetLatestVLFDate(hmVlf, "A", ilVefCode, ilDay, ilPending)
'                    If llTestDate > 0 Then
'                        ilDay = gWeekDayLong(llTestDate)
'                        If rbcDay(0).Value Then 'Mo-Fr
'                            llTestDate = llTestDate - ilDay 'Back up to monday
'                        ElseIf rbcDay(1).Value Then 'Sat
'                            If ilDay = 6 Then
'                                llTestDate = llTestDate - 1
'                            ElseIf ilDay <= 4 Then
'                                llTestDate = llTestDate - ilDay - 2
'                            End If
'                        Else    'Sunday
'                            If ilDay <= 5 Then
'                                llTestDate = llTestDate - ilDay - 1
'                            End If
'                        End If
'                        If ilPending Then
'                            If llTestDate <> llEffDate Then
'                                Screen.MousePointer = vbDefault
'                                slDate = Format$(llTestDate, "m/d/yy")
'                                MsgBox "Date must be " & slDate, vbExclamation, "Links"
'                                edcStartDate.Enabled = True
'                                edcStartDate.SetFocus
'                                Exit Sub
'                            End If
'                        Else
'                            If llTestDate > llEffDate Then
'                                Screen.MousePointer = vbDefault
'                                slDate = Format$(llTestDate - 1, "m/d/yy")
'                                MsgBox "Date must be after " & slDate, vbExclamation, "Links"
'                                edcStartDate.Enabled = True
'                                edcStartDate.SetFocus
'                                Exit Sub
'                            End If
'                        End If
'                    End If
'                End If
'            Next ilLoop
'        End If
        mGetVLFPendingDate
        If smVLFPendingDate <> "" Then
            If gDateValue(slStartDate) <> gDateValue(smVLFPendingDate) Then
                Screen.MousePointer = vbDefault
                MsgBox "Start Date Date must be " & smVLFPendingDate, vbExclamation, "Links"
                edcStartDate.Text = smVLFPendingDate
                edcStartDate.Enabled = True
                edcStartDate.SetFocus
                Exit Sub
            End If
        End If
        mSetCommands
        Screen.MousePointer = vbDefault
        'mCloseFiles
        'Screen.MousePointer = vbHourGlass  'Wait
        On Error Resume Next
        LinksDef.Show vbModal
        On Error GoTo 0
        'Screen.MousePointer = vbDefault    'Default
    Else    'Delivery or Engineering
        'If Not gWinRoom(igNoExeWinRes(LINKDLVYEXE)) Then
        '    Screen.MousePointer = vbDefault
        '    Exit Sub
        'End If

        'Date must be after latest terminate date
        slDate = Trim$(edcStartDate.Text)   'Store Effective Date
        llEffDate = gDateValue(slDate)
        If rbcDay(1).Value Then
            slDay = "6"
        ElseIf rbcDay(2).Value Then
            slDay = "7"
        Else
            slDay = "0"
        End If
        llAirDate = -1
        ilVeh = lbcACVeh.ListIndex
        slNameCode = tgVehCombo(ilVeh).sKey    'lbcACVehCode.List(ilVeh)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        tmDlfSrchKey.iVefCode = ilVefCode
        tmDlfSrchKey.sAirDay = slDay
        tmDlfSrchKey.iStartDate(0) = 257  'Year 1/1/1900
        tmDlfSrchKey.iStartDate(1) = 2100
        tmDlfSrchKey.iAirTime(0) = 0
        tmDlfSrchKey.iAirTime(1) = 0
        ilRet = btrGetLessOrEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        'Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVefCode) And (tmDlf.sAirDay = slDay)
        '    ilRet = btrGetPrevious(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        'Loop
        If (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVefCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iMnfFeed > 0) Then
            gUnpackDate tmDlf.iStartDate(0), tmDlf.iStartDate(1), slDate
            llTestDate = gDateValue(slDate)
            If llAirDate = -1 Then
                llAirDate = llTestDate
            Else
                If llTestDate < llAirDate Then
                    llAirDate = llTestDate
                End If
            End If
        End If
        llTestDate = mPendingDate(ilVefCode)
        If llTestDate > 0 Then
            If llTestDate < llAirDate Then
                llAirDate = llTestDate
            End If
        End If
        'If llAirDate <> -1 Then
        '    If llEffDate < llAirDate Then
        '        Screen.MousePointer = vbDefault
        '        slDate = Format$(llAirDate - 1, "m/d/yy")
        '        MsgBox "Date must be after " & slDate, vbExclamation, "Links"
        '        edcStartDate.Enabled = True
        '        edcStartDate.SetFocus
        '        Exit Sub
        '    End If
        'End If
        Screen.MousePointer = vbDefault
        LinkDlvy.Show vbModal
        'Reset dates
        mAddDEDates
        mSetCommands
    End If
End Sub
Private Sub cmcLinksDef_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcReport_Click()
    Dim slStr As String
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = PROGRAMMINGJOB
    If rbcLinks(0).Value Then
        igRptType = 0   'Selling to airing
    ElseIf rbcLinks(1).Value Then
        igRptType = 1   'Delivery
    Else
        igRptType = 2   'Engineering
    End If
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Links^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
        Else
            slStr = "Links^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Links^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    Else
    '        slStr = "Links^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'Links.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'Links.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    sgCommandStr = slStr
    RptList.Show vbModal
    ''Screen.MousePointer = vbDefault    'Default
End Sub
'*******************************************************
'
'       Procedure Name : cmcStartDate_Click
'
'       Created : ?             By : D. Levine
'       Modified :              By :
'
'       Comments : Reset startdate via call to edcStartdate
'
'*******************************************************
'
Private Sub cmcStartDate_Click()
    If imDateIndex = 1 Then
        plcCalendar.Visible = True
    Else
        plcCalendar.Visible = Not plcCalendar.Visible
    End If
    imDateIndex = 0
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcStartDate_GotFocus()
    Dim slStr As String
    gCtrlGotFocus ActiveControl
    'imDateIndex = 0
    plcCalendar.Move plcLinkDates.Left + edcStartDate.Left, plcLinkDates.Top + edcStartDate.Top + edcStartDate.Height
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcEndDate_Change()
    Dim slStr As String
    slStr = edcEndDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcEndDate_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
        If imDateIndex = 0 Then
            plcCalendar.Visible = False
        End If
        gCtrlGotFocus ActiveControl
        imDateIndex = 1
        plcCalendar.Move plcLinkDates.Left + cmcEndDate.Left + cmcEndDate.Width - plcCalendar.Width, plcLinkDates.Top + edcEndDate.Top + edcEndDate.Height
    End If
    imBypassFocus = False
End Sub
Private Sub edcEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcEndDate.SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcEndDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDate.Text = slDate
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            If rbcLinks(0).Value Then
                Beep
            Else
                slDate = edcEndDate.Text
                If gValidDate(slDate) Then
                    If KeyCode = KEYLEFT Then 'Up arrow
                        slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                    Else
                        slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                    End If
                    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                    edcEndDate.Text = slDate
                End If
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
'*******************************************************
'
'       Procedure Name : edcStartDate_Change
'
'       Created : ?             By : D. Levine
'       Modified :              By :
'
'       Comments : Reset startdate
'
'*******************************************************
'
Private Sub edcStartDate_Change()
    Dim slStr As String
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    If rbcLinks(0).Value Then
        Screen.MousePointer = vbHourglass
        mGetVlfEndDate
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub edcStartDate_GotFocus()
    If Not imBypassFocus Then
        If imDateIndex = 1 Then
            plcCalendar.Visible = False
        End If
        gCtrlGotFocus ActiveControl
        imDateIndex = 0
        plcCalendar.Move plcLinkDates.Left + edcStartDate.Left, plcLinkDates.Top + edcStartDate.Top + edcStartDate.Height
    End If
    imBypassFocus = False
End Sub
Private Sub edcStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcStartDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcStartDate.SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcStartDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            If rbcLinks(0).Value Then
                Beep
            Else
                slDate = edcStartDate.Text
                If gValidDate(slDate) Then
                    If KeyCode = KEYLEFT Then 'Up arrow
                        slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                    Else
                        slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                    End If
                    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                    edcStartDate.Text = slDate
                End If
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner
    Me.KeyPreview = True
    Links.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
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
        mTerminate
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    
    Erase imGroupNo
    ilRet = btrClose(hmEgf)
    btrDestroy hmEgf
    ilRet = btrClose(hmDlf)
    btrDestroy hmDlf
    ilRet = btrClose(hmVLF)
    btrDestroy hmVLF
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef

    Set Links = Nothing
    End
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcACVeh_Click()
    If Not imIgnoreClick Then
        imIgnoreClick = True
        mSetCommands
        imIgnoreClick = False
    End If
End Sub
Private Sub lbcACVeh_GotFocus()
    plcCalendar.Visible = False
    mSetCommands
End Sub
Private Sub lbcAiring_Click()
    'mSetCommands
End Sub
Private Sub lbcAiring_GotFocus()
    'plcCalendar.Visible = False
    'mSetCommands
End Sub
Private Sub lbcSAVeh_Click()
    Screen.MousePointer = vbHourglass
    mGetVlfEndDate
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub
Private Sub lbcSAVeh_GotFocus()
    plcCalendar.Visible = False
    mSetCommands
End Sub
Private Sub lbcSelling_Click()
    'mSetCommands
End Sub
Private Sub lbcSelling_GotFocus()
    plcCalendar.Visible = False
    mSetCommands

End Sub
'*******************************************************
'
'       Procedure Name : mAddDEDates
'
'       Created : 4/17/94       By : D. Hannifan
'       Modified :              By :
'
'       Comments : Add Delivery or Engineering Dates
'
'*******************************************************
Private Sub mAddDEDates()
    Dim ilVeh As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilVefCode As Integer
    Dim slNowDate As String
    Dim slDay As String
    Dim ilNowDate0 As Integer
    Dim ilNowDate1 As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilFound As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    If rbcLinks(0).Value Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    slNowDate = Format$(Now, "m/d/yy")
    If rbcDay(0).Value Then
        slDay = "0"
    ElseIf rbcDay(1).Value Then
        slDay = "6"
    Else
        slDay = "7"
    End If
    lbcACVeh.Clear
    For ilVeh = 0 To UBound(tgVehCombo) - 1 Step 1  'lbcACVehCode.ListCount - 1 Step 1
        ilFound = False
        slNameCode = tgVehCombo(ilVeh).sKey    'lbcACVehCode.List(ilVeh)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        tmDlfSrchKey.iVefCode = ilVefCode
        tmDlfSrchKey.sAirDay = slDay
        gPackDate slNowDate, ilNowDate0, ilNowDate1
        tmDlfSrchKey.iStartDate(0) = ilNowDate0
        tmDlfSrchKey.iStartDate(1) = ilNowDate1
        tmDlfSrchKey.iAirTime(0) = 0
        tmDlfSrchKey.iAirTime(1) = 25 * 256 'Hours
        If rbcLinks(1).Value Then
            ilRet = btrGetLessOrEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        Else
            ilRet = btrGetLessOrEqual(hmEgf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        End If
        Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVefCode) And (tmDlf.sAirDay = slDay)
            gUnpackDate tmDlf.iStartDate(0), tmDlf.iStartDate(1), slStartDate
            If (tmDlf.iTermDate(0) <> 0) Or (tmDlf.iTermDate(1) <> 0) Then
                gUnpackDate tmDlf.iTermDate(0), tmDlf.iTermDate(1), slEndDate
            Else
                slEndDate = "TFN"
            End If
            If Not ilFound Then
                slName = slName & ": " & slStartDate & "-" & slEndDate
            Else
                slName = slName & "; " & slStartDate & "-" & slEndDate
            End If
            ilFound = True
            If slEndDate = "TFN" Then
                Exit Do
            End If
            tmDlfSrchKey.iVefCode = ilVefCode
            tmDlfSrchKey.sAirDay = slDay
            slStartDate = Format$(gDateValue(slEndDate) + 1, "m/d/yy")
            gPackDate slStartDate, ilDate0, ilDate1
            tmDlfSrchKey.iStartDate(0) = ilDate0
            tmDlfSrchKey.iStartDate(1) = ilDate1
            tmDlfSrchKey.iAirTime(0) = 0
            tmDlfSrchKey.iAirTime(1) = 0
            If rbcLinks(1).Value Then
                ilRet = btrGetGreaterOrEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            Else
                ilRet = btrGetGreaterOrEqual(hmEgf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            End If
        Loop
        If Not ilFound Then
            tmDlfSrchKey.iVefCode = ilVefCode
            tmDlfSrchKey.sAirDay = slDay
            gPackDate slNowDate, ilNowDate0, ilNowDate1
            tmDlfSrchKey.iStartDate(0) = ilNowDate0
            tmDlfSrchKey.iStartDate(1) = ilNowDate1
            tmDlfSrchKey.iAirTime(0) = 0
            tmDlfSrchKey.iAirTime(1) = 0
            If rbcLinks(1).Value Then
                ilRet = btrGetGreaterOrEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            Else
                ilRet = btrGetGreaterOrEqual(hmEgf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            End If
            Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVefCode) And (tmDlf.sAirDay = slDay)
                gUnpackDate tmDlf.iStartDate(0), tmDlf.iStartDate(1), slStartDate
                If (tmDlf.iTermDate(0) <> 0) Or (tmDlf.iTermDate(1) <> 0) Then
                    gUnpackDate tmDlf.iTermDate(0), tmDlf.iTermDate(1), slEndDate
                Else
                    slEndDate = "TFN"
                End If
                If Not ilFound Then
                    slName = slName & ": " & slStartDate & "-" & slEndDate
                Else
                    slName = slName & "; " & slStartDate & "-" & slEndDate
                End If
                ilFound = True
                If slEndDate = "TFN" Then
                    Exit Do
                End If
                tmDlfSrchKey.iVefCode = ilVefCode
                tmDlfSrchKey.sAirDay = slDay
                slStartDate = Format$(gDateValue(slEndDate) + 1, "m/d/yy")
                gPackDate slStartDate, ilDate0, ilDate1
                tmDlfSrchKey.iStartDate(0) = ilDate0
                tmDlfSrchKey.iStartDate(1) = ilDate1
                tmDlfSrchKey.iAirTime(0) = 0
                tmDlfSrchKey.iAirTime(1) = 0
                If rbcLinks(1).Value Then
                    ilRet = btrGetGreaterOrEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                Else
                    ilRet = btrGetGreaterOrEqual(hmEgf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                End If
            Loop
        End If
        lbcACVeh.AddItem slName
    Next ilVeh
    Screen.MousePointer = vbDefault
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
    If imDateIndex = 0 Then
        slStr = edcStartDate.Text
    Else
        slStr = edcEndDate.Text
    End If
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
'************************************************************
'          Procedure Name : mGetPending
'
'    Created : 4/17/94      By : D. Hannifan
'    Modified :             By :
'
'    Comments: Scan LCF and VLF for Pending and
'              an effective start date
'
'************************************************************
'
Private Sub mGetPending()
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim llDate As Long
    Dim slDate As String    'Effective date
    Dim ilRet As Integer    'Return from btrieve call
    Dim llNoRec As Long     'Number of LCF records
    Dim ilDay As Integer    'Day code 0-4=M-F, 6=Sa, 7=Su
    Dim ilPass As Integer
    Dim ilEndCount As Integer
    Dim ilFound As Integer
    Dim slPD As String
    Dim llLoop As Long
    On Error GoTo mGetPendingErr

    imLcfPending = False       'initialize pending to no pending
    imVlfPending = False
    llDate = 0
    ilEndCount = lbcSelling.ListCount - 1

    For ilPass = 0 To 3 Step 1
        If (ilPass = 0) Or (ilPass = 2) Then
            ilEndCount = lbcSelling.ListCount - 1
        ElseIf (ilPass = 1) Or (ilPass = 3) Then
            ilEndCount = lbcAiring.ListCount - 1
        End If
        For ilLoop = 0 To ilEndCount Step 1
            If (ilPass = 0) Or (ilPass = 2) Then
                slNameCode = tgUserVehicle(ilLoop).sKey 'lbcVehName.List(ilLoop)
            ElseIf (ilPass = 1) Or (ilPass = 3) Then
                slNameCode = tgVehicle(ilLoop).sKey 'lbcVehMName.List(ilLoop)
            End If
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilCode = Val(slCode)
            tmLcfSrchKey.iType = 0
            If (ilPass = 0) Or (ilPass = 1) Then
                slPD = "P"
            Else
                slPD = "D"
            End If
            tmLcfSrchKey.sStatus = slPD
            tmLcfSrchKey.iVefCode = ilCode
            tmLcfSrchKey.iLogDate(0) = 257  'Year 1/1/1900
            tmLcfSrchKey.iLogDate(1) = 1900
            tmLcfSrchKey.iSeqNo = 1
            ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.sStatus = slPD) And (tmLcf.iVefCode = ilCode) And (tmLcf.iType = 0)
                gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slDate
                If rbcDay(0).Value Then
                    If gWeekDayStr(slDate) = 0 Then
                        If llDate = 0 Then
                            llDate = gDateValue(slDate)
                        Else
                            If gDateValue(slDate) < llDate Then
                                llDate = gDateValue(slDate)
                            End If
                        End If
                        Exit Do
                    End If
                ElseIf rbcDay(1).Value Then
                    If gWeekDayStr(slDate) = 5 Then
                        If llDate = 0 Then
                            llDate = gDateValue(slDate)
                        Else
                            If gDateValue(slDate) < llDate Then
                                llDate = gDateValue(slDate)
                            End If
                        End If
                        Exit Do
                    End If
                Else
                    If gWeekDayStr(slDate) = 6 Then
                        If llDate = 0 Then
                            llDate = gDateValue(slDate)
                        Else
                            If gDateValue(slDate) < llDate Then
                                llDate = gDateValue(slDate)
                            End If
                        End If
                        Exit Do
                    End If
                End If
                ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next ilLoop
    Next ilPass
    If llDate > 0 Then
        slDate = Format$(llDate, "m/d/yy")
        edcStartDate.Text = slDate 'Set default date
        imLcfPending = True
        imDateSetFlag = True
        'Exit Sub- set test for pending Vlf
    End If
    ilFound = gChrRefExist(Links, "P", "Vlf.Btr", "VLFSTATUS")
    If Not ilFound Then
        Exit Sub
    End If
           ' If no pending in LCF : scan VLF for pending
'    If imLcfPending = False Then
        llLoop = 1
        llNoRec = btrRecords(hmVLF) 'Obtain number of records
        If (llNoRec > 0) Then
            On Error GoTo 0
            ilRet = btrGetFirst(hmVLF, tmVlf, imVlfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY) 'Get first record
            On Error GoTo mGetPendingErr
            gBtrvErrorMsg ilRet, "mGetPending (btrGetFirst: Vlf.Btr)", Links
            If (tmVlf.iEffDate(1) > 0) Then
                gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), slDate
                ilDay = gWeekDayStr(slDate)
                llLoop = 2
            Else
                llLoop = 2
                GoTo lGetNextVlf
            End If
        Else
            Exit Sub
        End If
        If (ilDay = imDateCode) And (tmVlf.sStatus = "P") Then
            edcStartDate.Text = slDate
            imVlfPending = True
            imDateSetFlag = True
            Exit Sub
        End If
lGetNextVlf:
        Do While (llLoop <= llNoRec)
            On Error GoTo 0
            ilRet = btrGetNext(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo mGetPendingErr
            gBtrvErrorMsg ilRet, "mGetPending (btrGetNext: Vlf.Btr)", Links
            If (tmVlf.iEffDate(1) > 0) Then
                gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), slDate
                ilDay = gWeekDayStr(slDate)
            Else
                llLoop = llLoop + 1
                GoTo lGetNextVlf
            End If
            If (ilDay = imDateCode) And (tmVlf.sStatus = "P") Then
                edcStartDate.Text = slDate
                imVlfPending = True
                imDateSetFlag = True
                Exit Sub
            End If
            llLoop = llLoop + 1
        Loop
'    End If
    Exit Sub
mGetPendingErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()

    Dim ilRet As Integer   'Return from btrieve calls

    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imFirstActivate = True
    imTerminate = False
    bmFirstCallToVpfFind = True
    Screen.MousePointer = vbHourglass
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    Links.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone Links
    'Links.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture


    Screen.MousePointer = vbHourglass
    imBypassFocus = False
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    imCalType = 0   'Standard
    imLcfPending = False
    imVlfPending = False
    imIgnoreClick = False
    'Open btrieve files
    imLcfRecLen = Len(tmLcf)  'Get and save LCF record length
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", Links
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)  'Get and save Vlf record length
    hmVLF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vlf.Btr)", Links
    On Error GoTo 0
    imVefRecLen = Len(tmVef)  'Get and save Vef record length
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", Links
    On Error GoTo 0
    imDlfRecLen = Len(tmDlf)  'Get and save DlF record length
    hmDlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDlf, "", sgDBPath & "Dlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dlf.Btr)", Links
    On Error GoTo 0
    hmEgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmEgf, "", sgDBPath & "Egf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Egf.Btr)", Links
    On Error GoTo 0

    mVehPop 2, lbcDEName, tgVehCombo(), sgVehComboTag  'lbcACVehCode
    If imTerminate Then
        Exit Sub
    End If
    If tgSpf.sSystemType = "R" Then
        plcLinks.Visible = False
    End If
    If (tgSpf.sSDelNet <> "Y") Then
        rbcLinks(1).Enabled = False
    End If
    If rbcLinks(0).Value Then
        rbcLinks_Click 0
    Else
        rbcLinks(0).Value = True
    End If
    If rbcDay(0).Value Then
        rbcDay_Click 0
    Else
        rbcDay(0).Value = True  'This will cause mInitDay to be called
    End If
    Links.Refresh
    Screen.MousePointer = vbDefault
    If imTerminate Then
        Exit Sub
    End If
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
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
    Dim ilLoop As Integer
    On Error GoTo mInitBoxErr
    plcCalendar.Move plcLinkDates.Left + edcStartDate.Left, plcLinkDates.Top + edcStartDate.Top + edcStartDate.Height
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    Exit Sub
mInitBoxErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitDay                        *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInitDay()

    Dim ilRet As Integer   'Return from btrieve calls
    Dim slDate As String
    Dim llDate As Long
    Dim ilVeh As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilGroupNo As Integer
    Dim ilFound As Integer
    Dim ilVef As Integer
    Dim ilLp As Integer
    Dim slStr As String
    Dim ilVpfIndex As Integer
    Dim ilSel As Integer
    Dim ilLoop As Integer
    Dim ilMaxGroupNo As Integer
    Dim llMaxWidth As Long
    Dim llValue As Long
    Dim llRg As Long
    Dim llRet As Long
    Dim blFirstCallToVpfFind As Boolean

    Screen.MousePointer = vbHourglass
    'lbcVehName.Clear
    'lbcVehName.Tag = ""
    ReDim tgUserVehicle(0 To 0) As SORTCODE
    sgUserVehicleTag = ""
    lbcSelling.Clear
    lbcSelling.Tag = ""
    mVehPop 0, lbcSelling, tgUserVehicle(), sgUserVehicleTag    'lbcVehName   'Selling
    If imTerminate Then
        Exit Sub
    End If
    'lbcVehMName.Clear
    'lbcVehMName.Tag = ""
    ReDim tgVehicle(0 To 0) As SORTCODE
    sgVehicleTag = ""
    lbcAiring.Clear
    lbcAiring.Tag = ""
    mVehPop 1, lbcAiring, tgVehicle(), sgVehicleTag 'lbcVehMName   'Airing
    If imTerminate Then
        Exit Sub
    End If

    lbcSelling.ListIndex = -1    'Clear lists and reset controls & files
    lbcAiring.ListIndex = -1

    edcStartDate.Text = ""
    edcEndDate.Text = ""
    imLcfPending = False
    imVlfPending = False
    imDateSetFlag = False
    plcCalendar.Visible = False

    If rbcDay(0).Value <> False Then     'Set Date code
        imDateCode = 0
    ElseIf rbcDay(1).Value <> False Then
        imDateCode = 5
    ElseIf rbcDay(2).Value <> False Then
        imDateCode = 6
    End If
    'Determine if LCF or VLF pending exist-and if so set date, if not use tomorrow
    'edcStartDate.Text = Format(DateAdd("d", 1, Now), "mm/dd/yy")

    mGetPending       ' Check for pending vehicles
    If imTerminate Then
        Exit Sub
    End If
    mSetCommands      ' Set Controls
    If imTerminate Then
        Exit Sub
    End If

    If Not imLcfPending And Not imVlfPending Then  'Neither LCF or VLF = Pending
        If rbcDay(0).Value Then
           slDate = gObtainMondayFromToday()
        ElseIf rbcDay(1).Value Then
           slDate = gObtainMondayFromToday()
           llDate = gDateValue(slDate) + 5
           slDate = Format$(llDate, "m/d/yy")
        Else
           slDate = gObtainMondayFromToday()
           llDate = gDateValue(slDate) + 6
           slDate = Format$(llDate, "m/d/yy")
        End If
        edcStartDate.Text = slDate
    End If
    Screen.MousePointer = vbHourglass
    DoEvents

    mSetSelected  ' Set pending vehicles in lists to selected
    If imTerminate Then
        Exit Sub
    End If
    ReDim imGroupNo(0 To 0) As Integer
    lbcSAVeh.Clear
    ilMaxGroupNo = 0
    For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If bmFirstCallToVpfFind Then
            ilVpfIndex = gVpfFind(Links, tgMVef(ilVeh).iCode)
            bmFirstCallToVpfFind = False
        Else
            ilVpfIndex = gVpfFindIndex(tgMVef(ilVeh).iCode)
        End If
        If tgVpf(ilVpfIndex).iSAGroupNo > 0 Then
            If tgVpf(ilVpfIndex).iSAGroupNo > ilMaxGroupNo Then
                ilMaxGroupNo = tgVpf(ilVpfIndex).iSAGroupNo
            End If
        End If
    Next ilVeh
    blFirstCallToVpfFind = True
    llMaxWidth = 0
    For ilLoop = 1 To ilMaxGroupNo Step 1
        ilSel = False
        ilGroupNo = ilLoop
        slStr = "Group #" & Str$(ilGroupNo) & ":"
        ilFound = False
        For ilLp = 0 To lbcSelling.ListCount - 1 Step 1
            slNameCode = tgUserVehicle(ilLp).sKey 'lbcVehName.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '    If ilVefCode = tgMVef(ilVef).iCode Then
                ilVef = gBinarySearchVef(ilVefCode)
                If ilVef <> -1 Then
                    If bmFirstCallToVpfFind Then
                        ilVpfIndex = gVpfFind(Links, ilVefCode)
                        bmFirstCallToVpfFind = False
                    Else
                        ilVpfIndex = gVpfFindIndex(ilVefCode)
                    End If
                    If tgVpf(ilVpfIndex).iSAGroupNo = ilGroupNo Then
                        'If lbcSelling.Selected(ilLp) Then 'Selected vehicle found
                        '    ilSel = True
                        'End If
                        If ilFound Then
                            '3/26/10: Limit length of vehicle names to 1000 characters to avoid overflow error
                            'slStr = slStr & "; " & Trim$(tgMVef(ilVef).sName)
                            If Len(slStr) < 1000 Then
                                slStr = slStr & "; " & Trim$(tgMVef(ilVef).sName)
                                If Len(slStr) > 1000 Then
                                    slStr = slStr & " ....."
                                End If
                            End If
                        Else
                            slStr = slStr & " " & Trim$(tgMVef(ilVef).sName)
                            ilFound = True
                        End If
                    End If
            '        Exit For
                End If
            'Next ilVef
        Next ilLp
        For ilLp = 0 To lbcAiring.ListCount - 1 Step 1
            slNameCode = tgVehicle(ilLp).sKey 'lbcVehName.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '    If ilVefCode = tgMVef(ilVef).iCode Then
                ilVef = gBinarySearchVef(ilVefCode)
                If ilVef <> -1 Then
                    If bmFirstCallToVpfFind Then
                        ilVpfIndex = gVpfFind(Links, ilVefCode)
                        bmFirstCallToVpfFind = False
                    Else
                        ilVpfIndex = gVpfFindIndex(ilVefCode)
                    End If
                    If tgVpf(ilVpfIndex).iSAGroupNo = ilGroupNo Then
                        'If lbcAiring.Selected(ilLp) Then 'Selected vehicle found
                        '    ilSel = True
                        'End If
                        If ilFound Then
                            '3/26/10: Limit length of vehicle names to 1000 characters to avoid overflow error
                            'slStr = slStr & "; " & Trim$(tgMVef(ilVef).sName)
                            If Len(slStr) < 1000 Then
                                slStr = slStr & "; " & Trim$(tgMVef(ilVef).sName)
                                If Len(slStr) > 1000 Then
                                    slStr = slStr & " ....."
                                End If
                            End If
                        Else
                            slStr = slStr & " " & Trim$(tgMVef(ilVef).sName)
                            ilFound = True
                        End If
                    End If
            '        Exit For
                End If
            'Next ilVef
        Next ilLp
        If ilFound Then
            imGroupNo(UBound(imGroupNo)) = ilGroupNo
            ReDim Preserve imGroupNo(0 To UBound(imGroupNo) + 1) As Integer
            lbcSAVeh.AddItem slStr
            If ilSel Then
                lbcSAVeh.Selected(lbcSAVeh.ListCount - 1) = True
            End If
            If (Traffic.pbcArial.TextWidth(slStr)) > llMaxWidth Then
                llMaxWidth = (Traffic.pbcArial.TextWidth(slStr))
            End If
        End If
    Next ilLoop
    If llMaxWidth > lbcSAVeh.Width Then
        llValue = llMaxWidth / 15
        '3/26/10:  Add extra room
        llValue = llValue + llValue / 4
        llRg = 0
        llRet = SendMessageByNum(lbcSAVeh.hWnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    mSetCommands      'Disable LinksDef if No Vehicles are selected
    If imTerminate Then
        Exit Sub
    End If

    If Not imDateSetFlag Then   ' Current mode : use a default date depending on day code
        If rbcDay(0).Value Then
           slDate = gObtainMondayFromToday()
        ElseIf rbcDay(1).Value Then
           slDate = gObtainMondayFromToday()
           llDate = gDateValue(slDate) + 5
           slDate = Format$(llDate, "m/d/yy")
        Else
           slDate = gObtainMondayFromToday()
           llDate = gDateValue(slDate) + 6
           slDate = Format$(llDate, "m/d/yy")
        End If
        edcStartDate.Text = slDate
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
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
'    Dim slCommand As String
'    Dim slStr As String
'    Dim ilRet As Integer
'    Dim slTestSystem As String
'    Dim ilTestSystem As Integer
'    Dim slHelpSystem As String
'    slCommand = sgCommandStr    'Command$
'    'If StrComp(slCommand, "Debug", 1) = 0 Then
'    '    igStdAloneMode = True 'Switch from/to stand alone mode
'    '    sgCallAppName = ""
'    '    slStr = "Guide"
'    '    ilTestSystem = False
'    '    imShowHelpMsg = False
'    'Else
'    '    igStdAloneMode = False  'Switch from/to stand alone mode
'        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
'        If Trim$(slStr) = "" Then
'            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
'            'End
'            imTerminate = True
'            Exit Sub
'        End If
'        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
'        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
'        If StrComp(slTestSystem, "Test", 1) = 0 Then
'            ilTestSystem = True
'        Else
'            ilTestSystem = False
'        End If
'        imShowHelpMsg = True
'        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
'        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
'            imShowHelpMsg = False
'        End If
'        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
'    'End If
'    'gInitStdAlone Links, slStr, ilTestSystem
'    'igShowHelpMsg = imShowHelpMsg
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slStartIn As String
    Dim slCSIName As String
    Dim slDate As String
    Dim slMonth As String
    Dim slYear As String
    Dim llValue As Long
    Dim ilValue As Integer

    
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
    Else
        igSportsSystem = 0
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        'ilRet = gParseItem(slCommand, 3, "\", slStr)
        'igRptCallType = Val(slStr)
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
        sgUrfStamp = "~" 'Clear time stamp incase same name
        sgUserName = Trim$(slStr)
        '6/20/09:  Jim requested that the Guide sign in be changed to CSI for internal Guide only
        If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
            slDate = Format$(Now(), "m/d/yy")
            slMonth = Month(slDate)
            slYear = Year(slDate)
            llValue = Val(slMonth) * Val(slYear)
            ilValue = Int(10000 * Rnd(-llValue) + 1)
            llValue = ilValue
            ilValue = Int(10000 * Rnd(-llValue) + 1)
            slStr = Trim$(Str$(ilValue))
            Do While Len(slStr) < 4
                slStr = "0" & slStr
            Loop
            sgSpecialPassword = slStr
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
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "UserOpt.Frm"
    If igWinStatus(PROGRAMMINGJOB) = 0 Then
        imTerminate = True
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPendingDate                    *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Determine pending for vehicle  *
'*                                                     *
'*******************************************************
Private Function mPendingDate(ilVefCode As Integer) As Long
    Dim ilPass As Integer
    Dim llDate As Long
    Dim slPD As String
    Dim slDate As String
    Dim ilRet As Integer
    llDate = 0
    For ilPass = 1 To 2
        tmLcfSrchKey.iType = 0
        If (ilPass = 1) Then
            slPD = "P"
        Else
            slPD = "D"
        End If
        tmLcfSrchKey.sStatus = slPD
        tmLcfSrchKey.iVefCode = ilVefCode
        tmLcfSrchKey.iLogDate(0) = 257  'Year 1/1/1900
        tmLcfSrchKey.iLogDate(1) = 1900
        tmLcfSrchKey.iSeqNo = 1
        ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.sStatus = slPD) And (tmLcf.iVefCode = ilVefCode) And (tmLcf.iType = 0)
            gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slDate
            If rbcDay(0).Value Then
                If gWeekDayStr(slDate) = 0 Then
                    If llDate = 0 Then
                        llDate = gDateValue(slDate)
                    Else
                        If gDateValue(slDate) < llDate Then
                            llDate = gDateValue(slDate)
                        End If
                    End If
                    Exit Do
                End If
            ElseIf rbcDay(1).Value Then
                If gWeekDayStr(slDate) = 5 Then
                    If llDate = 0 Then
                        llDate = gDateValue(slDate)
                    Else
                        If gDateValue(slDate) < llDate Then
                            llDate = gDateValue(slDate)
                        End If
                    End If
                    Exit Do
                End If
            Else
                If gWeekDayStr(slDate) = 6 Then
                    If llDate = 0 Then
                        llDate = gDateValue(slDate)
                    Else
                        If gDateValue(slDate) < llDate Then
                            llDate = gDateValue(slDate)
                        End If
                    End If
                    Exit Do
                End If
            End If
            ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next ilPass
    mPendingDate = llDate
End Function
'************************************************************
'          Procedure Name : mSetCommands
'
'    Created : 4/17/94      By : D. Hannifan
'    Modified :             By :
'
'    Comments:  Set Control properties
'
'
'************************************************************
'
Private Sub mSetCommands()

    On Error GoTo mSetCommandsErr
    If rbcLinks(0).Value Then
        'If (lbcSelling.SelectCount = 0) And (lbcAiring.SelectCount = 0) Then
        If lbcSAVeh.ListIndex < 0 Then
           cmcLinksDef.Enabled = False    'No links possible
        Else
           cmcLinksDef.Enabled = True
        End If
    Else    'Delivery and Engineering
        If (lbcACVeh.SelCount = 0) Then
           cmcLinksDef.Enabled = False    'No links possible
        Else
           cmcLinksDef.Enabled = True
        End If
    End If
    If imDateSetFlag Then   ' Pending mode
        'lacStartDate.Enabled = False
        'edcStartDate.Enabled = False
        'cmcStartDate.Enabled = False
'        lacStartDate.Visible = False
'        edcStartdate.Visible = False
'        cmcStartDate.Visible = False
    Else                     ' Current Mode
        'lacStartDate.Enabled = True
        'edcStartDate.Enabled = True
        'cmcStartDate.Enabled = True
        'lacStartDate.Visible = True
        'edcStartDate.Visible = True
        'cmcStartDate.Visible = True
    End If
    Exit Sub
mSetCommandsErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'************************************************************
'          Procedure Name : mSetSelected
'
'    Created : 4/17/94      By : D. Hannifan
'    Modified :             By :
'
'    Comments: Scan LCF for Pending and
'              set pending vehicles to selcted in list boxes
'
'************************************************************
'
Private Sub mSetSelected()
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilVefCode As Integer
    Dim llDate As Long
    For ilLoop = 0 To lbcSelling.ListCount - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey 'lbcVehName.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        llDate = gGetEarliestLCFDate(hmLcf, "P", ilVefCode)
        If llDate > 0 Then
            lbcSelling.Selected(ilLoop) = True ' Match Located
        Else
            llDate = gGetEarliestLCFDate(hmLcf, "D", ilVefCode)
            If llDate > 0 Then
                lbcSelling.Selected(ilLoop) = True ' Match Located
            End If
        End If
    Next ilLoop
    For ilLoop = 0 To lbcAiring.ListCount - 1 Step 1
        slNameCode = tgVehicle(ilLoop).sKey 'lbcVehMName.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        llDate = gGetEarliestLCFDate(hmLcf, "P", ilVefCode)
        If llDate > 0 Then
            lbcAiring.Selected(ilLoop) = True ' Match Located
        Else
            llDate = gGetEarliestLCFDate(hmLcf, "D", ilVefCode)
            If llDate > 0 Then
                lbcAiring.Selected(ilLoop) = True ' Match Located
            End If
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: terminate Links                *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    'Unload IconTraf
    igManUnload = YES
    Unload Links
    igManUnload = NO
End Sub
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
Private Sub mVehPop(ilType, lbcVeh As Control, tlSortCode() As SORTCODE, slSortCodeTag As String)

    Dim ilRet As Integer
    Dim llFilter As Long
    Dim llDate As Long
    Dim ilVeh As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer

    Screen.MousePointer = vbHourglass
    DoEvents
    If ilType = 0 Then       'Selling
        llFilter = VEHSELLING + ACTIVEVEH '5
    ElseIf ilType = 1 Then   'Airing
        llFilter = VEHAIRING + ACTIVEVEH '1
    ElseIf ilType = 2 Then   'Airing & Conventional with feed
        llFilter = VEHCONV_W_FEED + VEHEXCLUDESPORT + VEHAIRING + ACTIVEVEH '6
    End If
    'ilRet = gPopUserVehicleBox(Links, ilFilter, lbcVeh, lbcMVeh)
    ilRet = gPopUserVehicleBox(Links, llFilter, lbcVeh, tlSortCode(), slSortCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", Links
        On Error GoTo 0
        'Remove any vehicle which have pending events when defining delivery
        If ilType = 2 Then
            For ilVeh = UBound(tlSortCode) - 1 To 0 Step -1  'lbcMVeh.ListCount - 1 To 0 Step -1
                slNameCode = tlSortCode(ilVeh).sKey  'lbcMVeh.List(ilVeh)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                llDate = gGetEarliestLCFDate(hmLcf, "P", ilVefCode)
                If llDate > 0 Then
                    'lbcMVeh.RemoveItem ilVeh
                    gRemoveItemFromSortCode ilVeh, tlSortCode()
                    lbcVeh.RemoveItem ilVeh
                Else
                    llDate = gGetEarliestLCFDate(hmLcf, "D", ilVefCode)
                    If llDate > 0 Then
                        'lbcMVeh.RemoveItem ilVeh
                        gRemoveItemFromSortCode ilVeh, tlSortCode()
                        lbcVeh.RemoveItem ilVeh
                    End If
                End If
            Next ilVeh
        End If
    End If
    Exit Sub
mVehPopErr:
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
        slDay = Trim$(Str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                If rbcDay(0).Value Then
                    If rbcLinks(0).Value Then
                        If ((gWeekDayLong(llDate) <> 0) And (imDateIndex = 0)) Then
                            Beep
                            edcStartDate.SetFocus
                            Exit Sub
                        ElseIf ((gWeekDayLong(llDate) <> 6) And (imDateIndex = 1)) Then
                            Beep
                            edcEndDate.SetFocus
                            Exit Sub
                        End If
                    Else
                        If gWeekDayLong(llDate) > 4 Then
                            Beep
                            edcStartDate.SetFocus
                            Exit Sub
                        End If
                    End If
                ElseIf rbcDay(1).Value Then
                    If rbcLinks(0).Value Then
                        If ((gWeekDayLong(llDate) <> 5) And (imDateIndex = 0)) Then
                            Beep
                            edcStartDate.SetFocus
                            Exit Sub
                        ElseIf ((gWeekDayLong(llDate) <> 4) And (imDateIndex = 1)) Then
                            Beep
                            edcEndDate.SetFocus
                            Exit Sub
                        End If
                    Else
                        If gWeekDayLong(llDate) <> 5 Then
                            Beep
                            If imDateIndex = 0 Then
                                edcStartDate.SetFocus
                            Else
                                edcEndDate.SetFocus
                            End If
                            Exit Sub
                        End If
                    End If
                Else
                    If rbcLinks(0).Value Then
                        If ((gWeekDayLong(llDate) <> 6) And (imDateIndex = 0)) Then
                            Beep
                            edcStartDate.SetFocus
                            Exit Sub
                        ElseIf ((gWeekDayLong(llDate) <> 5) And (imDateIndex = 1)) Then
                            Beep
                            edcEndDate.SetFocus
                            Exit Sub
                        End If
                    Else
                        If gWeekDayLong(llDate) <> 6 Then
                            Beep
                            If imDateIndex = 0 Then
                                edcStartDate.SetFocus
                            Else
                                edcEndDate.SetFocus
                            End If
                            Exit Sub
                        End If
                    End If
                End If
                If imDateIndex = 0 Then
                    edcStartDate.Text = Format$(llDate, "m/d/yy")
                    edcStartDate.SelStart = 0
                    edcStartDate.SelLength = Len(edcStartDate.Text)
                    imBypassFocus = True
                    edcStartDate.SetFocus
                Else
                    edcEndDate.Text = Format$(llDate, "m/d/yy")
                    edcEndDate.SelStart = 0
                    edcEndDate.SelLength = Len(edcEndDate.Text)
                    imBypassFocus = True
                    edcEndDate.SetFocus
                End If
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    If imDateIndex = 0 Then
        edcStartDate.SetFocus
    Else
        edcEndDate.SetFocus
    End If
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(Str$(imCalMonth)) & "/15/" & Trim$(Str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
'*******************************************************
'
'       Procedure Name : rbcLinks_Click
'
'       Created : ?             By : D. Levine
'       Modified :4/17/94       By : D.Hannifan
'
'       Comments : Call reinitialize Links
'
'*******************************************************
'
Private Sub rbcDay_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcDay(Index).Value
    'End of coded added
    If Value Then
        mInitDay
        If ((rbcLinks(1).Enabled = False) And (rbcLinks(1).Value)) Or ((rbcLinks(2).Enabled = False) And (rbcLinks(2).Value)) Then
            rbcLinks(0).Value = True
        End If
        mAddDEDates
        mSetCommands
    End If
End Sub
Private Sub rbcDay_GotFocus(Index As Integer)
    If imFirstTime Then
        imFirstTime = False
    End If
    plcCalendar.Visible = False
End Sub
'*******************************************************
'
'       Procedure Name : rbcLinks_Click
'
'       Created : ?             By : D. Levine
'       Modified : 4/17/94      By : D. Hannifan
'
'       Comments : Populate list boxes in Links
'
'*******************************************************
'
Private Sub rbcLinks_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcLinks(Index).Value
    'End of coded added
    If Value Then
        Select Case Index
            Case 0  ' Selling to Airing
                plcACVehicle.Visible = False
                lacEndDate.Visible = True
                edcEndDate.Visible = True
                edcEndDate.Enabled = False
                cmcEndDate.Visible = False
                'plcSelling.Visible = True
                'plcAiring.Visible = True
                plcSAVehicle.Visible = True
                cmcLinksDef.Caption = "Define Selling to Airing &Links"
            Case 1  'Delivery
                'plcSelling.Visible = False
                'plcAiring.Visible = False
                plcSAVehicle.Visible = False
                lacEndDate.Visible = True
                edcEndDate.Visible = True
                cmcEndDate.Visible = True
                edcEndDate.Enabled = True
                plcACVehicle.Visible = True
                cmcLinksDef.Caption = "Define Delivery &Links"
                mAddDEDates
            Case 2  'Engineering
                'plcSelling.Visible = False
                'plcAiring.Visible = False
                plcSAVehicle.Visible = False
                lacEndDate.Visible = True
                edcEndDate.Visible = True
                cmcEndDate.Visible = True
                edcEndDate.Enabled = True
                plcACVehicle.Visible = True
                cmcLinksDef.Caption = "Define Engineering &Links"
                mAddDEDates
        End Select
    End If
    mSetCommands
End Sub
Private Sub rbcLinks_GotFocus(Index As Integer)
    plcCalendar.Visible = False
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print " Links"
End Sub
Private Sub plcLinks_Paint()
    plcLinks.CurrentX = 0
    plcLinks.CurrentY = 0
    plcLinks.Print "Links"
End Sub
Private Sub plcDay_Paint()
    plcDay.CurrentX = 0
    plcDay.CurrentY = 0
    plcDay.Print "Day"
End Sub
Private Sub plcACVehicle_Paint()
    plcACVehicle.CurrentX = 0
    plcACVehicle.CurrentY = 0
    plcACVehicle.Print "Vehicles"
End Sub
Private Sub plcSAVehicle_Paint()
    plcSAVehicle.CurrentX = 0
    plcSAVehicle.CurrentY = 0
    plcSAVehicle.Print "Vehicles"
End Sub

Private Sub mGetVlfEndDate()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPass                                                                                *
'******************************************************************************************

    Dim ilDay As Integer
    Dim slDate As String
    Dim ilEffDate0 As Integer
    Dim ilEffDate1 As Integer
    Dim ilTerminated As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStartDate As String
    Dim llDate As Long
    Dim slType As String
    Dim ilLoop As Integer
    Dim ilGroupNo As Integer
    Dim ilFound As Integer
    Dim ilUpper As Integer
    Dim ilVef As Integer
    Dim llTestDate As Long
    Dim llEndDate As Long
    ReDim ilVefCode(0 To 0) As Integer
    Dim ilVpfIndex As Integer

    If Not rbcLinks(0).Value Then
        Exit Sub
    End If
    If lbcSAVeh.ListIndex < 0 Then
        Exit Sub
    End If
    slDate = Trim$(edcStartDate.Text)
    If slDate = "" Then
        Exit Sub
    End If
    If Not gValidDate(slDate) Then
        Exit Sub
    End If
    ilFound = False
    ilUpper = 0
    llEndDate = -1
    ilGroupNo = imGroupNo(lbcSAVeh.ListIndex)
    For ilLoop = 0 To lbcSelling.ListCount - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey 'lbcVehName.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode(ilUpper) = Val(slCode)
        If bmFirstCallToVpfFind Then
            ilVpfIndex = gVpfFind(Links, ilVefCode(ilUpper))
            bmFirstCallToVpfFind = False
        Else
            ilVpfIndex = gVpfFindIndex(ilVefCode(ilUpper))
        End If
        If tgVpf(ilVpfIndex).iSAGroupNo = ilGroupNo Then
            slType = "S"
            ilFound = True
            ilUpper = ilUpper + 1
            ReDim Preserve ilVefCode(0 To ilUpper) As Integer
            'Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        For ilLoop = 0 To lbcAiring.ListCount - 1 Step 1
            slNameCode = tgVehicle(ilLoop).sKey 'lbcVehMName.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode(ilUpper) = Val(slCode)
            If bmFirstCallToVpfFind Then
                ilVpfIndex = gVpfFind(Links, ilVefCode(ilUpper))
                bmFirstCallToVpfFind = False
            Else
                ilVpfIndex = gVpfFindIndex(ilVefCode(ilUpper))
            End If
            If tgVpf(ilVpfIndex).iSAGroupNo = ilGroupNo Then
                slType = "A"
                ilFound = True
                ilUpper = ilUpper + 1
                ReDim Preserve ilVefCode(0 To ilUpper) As Integer
                'Exit For
            End If
        Next ilLoop
    End If
    slStartDate = edcStartDate.Text
    llDate = gDateValue(slStartDate)
    gPackDate gIncOneDay(slStartDate), ilEffDate0, ilEffDate1
    ilDay = gWeekDayLong(llDate)
    If rbcDay(2).Value Then
        ilDay = 7
    ElseIf rbcDay(1).Value Then   'Saturady
        ilDay = 6
    Else
        ilDay = 0
    End If
    If slType = "S" Then
        For ilVef = 0 To UBound(ilVefCode) - 1 Step 1
            tmVlfSrchKey0.iSellCode = ilVefCode(ilVef)
            tmVlfSrchKey0.iSellDay = ilDay
            tmVlfSrchKey0.iEffDate(0) = ilEffDate0
            tmVlfSrchKey0.iEffDate(1) = ilEffDate1
            tmVlfSrchKey0.iSellTime(0) = 0
            tmVlfSrchKey0.iSellTime(1) = 0
            tmVlfSrchKey0.iSellPosNo = 0
            ilRet = btrGetGreaterOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = ilVefCode(ilVef)) And (tmVlf.iSellDay = ilDay)
                ilTerminated = False
                'Check for CBS
                If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                    If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                        ilTerminated = True
                    End If
                End If
                If (Not ilTerminated) Then
                    'gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), slDate
                    'edcEndDate.Text = gDecOneDay(slDate)
                    'Exit Sub
                    gUnpackDateLong tmVlf.iEffDate(0), tmVlf.iEffDate(1), llTestDate
                    If llEndDate = -1 Then
                        llEndDate = llTestDate
                    ElseIf llTestDate < llEndDate Then
                        llEndDate = llTestDate
                    End If
                End If
                ilRet = btrGetNext(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next ilVef
        If llEndDate <> -1 Then
            edcEndDate.Text = Format$(llEndDate - 1, "m/d/yy")
            Exit Sub
        End If
    Else
        For ilVef = 0 To UBound(ilVefCode) - 1 Step 1
            tmVlfSrchKey1.iAirCode = ilVefCode(ilVef)
            tmVlfSrchKey1.iAirDay = ilDay
            tmVlfSrchKey1.iEffDate(0) = ilEffDate0
            tmVlfSrchKey1.iEffDate(1) = ilEffDate1
            ilEffDate0 = 0
            ilEffDate1 = 0
            tmVlfSrchKey1.iAirTime(0) = 0
            tmVlfSrchKey1.iAirTime(1) = 0
            tmVlfSrchKey1.iAirPosNo = 0
            ilRet = btrGetGreaterOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVefCode(ilVef)) And (tmVlf.iAirDay = ilDay)
                ilTerminated = False
                'Check for CBS
                If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                    If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                        ilTerminated = True
                    End If
                End If
                If (Not ilTerminated) Then
                    'gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), slDate
                    'edcEndDate.Text = gDecOneDay(slDate)
                    'Exit Sub
                    gUnpackDateLong tmVlf.iEffDate(0), tmVlf.iEffDate(1), llTestDate
                    If llEndDate = -1 Then
                        llEndDate = llTestDate
                    ElseIf llTestDate < llEndDate Then
                        llEndDate = llTestDate
                    End If
                End If
                ilRet = btrGetNext(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next ilVef
        If llEndDate <> -1 Then
            edcEndDate.Text = Format$(llEndDate - 1, "m/d/yy")
            Exit Sub
        End If
    End If
    edcEndDate.Text = ""
End Sub

Private Sub mGetVLFPendingDate()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPass                        slDate                        slStartDate               *
'*                                                                                        *
'******************************************************************************************

    Dim ilDay As Integer
    Dim ilEffDate0 As Integer
    Dim ilEffDate1 As Integer
    Dim ilTerminated As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim llDate As Long
    Dim slType As String
    Dim ilLoop As Integer
    Dim ilGroupNo As Integer
    Dim ilFound As Integer
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim slNowDate As String

    If Not rbcLinks(0).Value Then
        Exit Sub
    End If
    If lbcSAVeh.ListIndex < 0 Then
        Exit Sub
    End If
    slNowDate = Format(gNow(), "m/d/yy")
    ilFound = False
    ilGroupNo = imGroupNo(lbcSAVeh.ListIndex)
    For ilLoop = 0 To lbcSelling.ListCount - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey 'lbcVehName.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        If bmFirstCallToVpfFind Then
            ilVpfIndex = gVpfFind(Links, ilVefCode)
            bmFirstCallToVpfFind = False
        Else
            ilVpfIndex = gVpfFindIndex(ilVefCode)
        End If
        If tgVpf(ilVpfIndex).iSAGroupNo = ilGroupNo Then
            slType = "S"
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        For ilLoop = 0 To lbcAiring.ListCount - 1 Step 1
            slNameCode = tgVehicle(ilLoop).sKey 'lbcVehMName.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            If bmFirstCallToVpfFind Then
                ilVpfIndex = gVpfFind(Links, ilVefCode)
                bmFirstCallToVpfFind = False
            Else
                ilVpfIndex = gVpfFindIndex(ilVefCode)
            End If
            If tgVpf(ilVpfIndex).iSAGroupNo = ilGroupNo Then
                slType = "A"
                ilFound = True
                Exit For
            End If
        Next ilLoop
    End If
    gPackDate slNowDate, ilEffDate0, ilEffDate1
    ilDay = gWeekDayLong(llDate)
    If rbcDay(2).Value Then
        ilDay = 7
    ElseIf rbcDay(1).Value Then   'Saturady
        ilDay = 6
    Else
        ilDay = 0
    End If
    If slType = "S" Then
        tmVlfSrchKey0.iSellCode = ilVefCode
        tmVlfSrchKey0.iSellDay = ilDay
        tmVlfSrchKey0.iEffDate(0) = ilEffDate0
        tmVlfSrchKey0.iEffDate(1) = ilEffDate1
        tmVlfSrchKey0.iSellTime(0) = 0
        tmVlfSrchKey0.iSellTime(1) = 0
        tmVlfSrchKey0.iSellPosNo = 0
        ilRet = btrGetGreaterOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = ilVefCode) And (tmVlf.iSellDay = ilDay)
            ilTerminated = False
            'Check for CBS
            If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                    ilTerminated = True
                End If
            End If
            If (Not ilTerminated) And (tmVlf.sStatus = "P") Then
                gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), smVLFPendingDate
                Exit Sub
            End If
            ilRet = btrGetNext(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Else
        tmVlfSrchKey1.iAirCode = ilVefCode
        tmVlfSrchKey1.iAirDay = ilDay
        tmVlfSrchKey1.iEffDate(0) = ilEffDate0
        tmVlfSrchKey1.iEffDate(1) = ilEffDate1
        ilEffDate0 = 0
        ilEffDate1 = 0
        tmVlfSrchKey1.iAirTime(0) = 0
        tmVlfSrchKey1.iAirTime(1) = 0
        tmVlfSrchKey1.iAirPosNo = 0
        ilRet = btrGetGreaterOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVefCode) And (tmVlf.iAirDay = ilDay)
            ilTerminated = False
            'Check for CBS
            If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                    ilTerminated = True
                End If
            End If
            If (Not ilTerminated) And (tmVlf.sStatus = "P") Then
                gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), smVLFPendingDate
                Exit Sub
            End If
            ilRet = btrGetNext(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    smVLFPendingDate = ""
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
            slSetDate = "12/15/1999"
            slDate = slSetDate
        End If
    End If
    If Trim$(slSetDate) <> "" Then
        'Dan M 9/20/10 problems with gGetCSIName("SYSDate") in v57 reports.exe... change to global variable
     '   ilRet = csiSetName("SYSDate", slDate)
        ilRet = gCsiSetName(slDate)
    End If
End Sub
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

