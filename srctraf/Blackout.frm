VERSION 5.00
Begin VB.Form Blackout 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5865
   ClientLeft      =   195
   ClientTop       =   1710
   ClientWidth     =   9390
   ClipControls    =   0   'False
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
   ScaleHeight     =   5865
   ScaleWidth      =   9390
   Begin VB.PictureBox pbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   1710
      Picture         =   "Blackout.frx":0000
      ScaleHeight     =   540
      ScaleWidth      =   3285
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   375
      Visible         =   0   'False
      Width           =   3315
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
      Picture         =   "Blackout.frx":6342
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox plcSelect 
      Height          =   330
      Left            =   6090
      ScaleHeight     =   270
      ScaleWidth      =   3090
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   3150
      Begin VB.OptionButton rbcSR 
         Caption         =   "Replacement"
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
         Left            =   1545
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   45
         Width           =   1440
      End
      Begin VB.OptionButton rbcSR 
         Caption         =   "Suppression"
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
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Value           =   -1  'True
         Width           =   1410
      End
   End
   Begin VB.ListBox lbcSRCntr 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1980
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4605
      Visible         =   0   'False
      Width           =   5835
   End
   Begin VB.ListBox lbcDays 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6750
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2655
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.ListBox lbcLen 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6600
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lbcRCntr 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2040
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   5835
   End
   Begin VB.ListBox lbcSCntr 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1965
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4395
      Visible         =   0   'False
      Width           =   5835
   End
   Begin VB.ListBox lbcRAdvt 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1260
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3105
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.ListBox lbcComp 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   1
      Left            =   1095
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3825
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.ListBox lbcComp 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   1305
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3510
      Visible         =   0   'False
      Width           =   2550
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
      Left            =   4425
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2565
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
         Left            =   1620
         TabIndex        =   30
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
         Left            =   30
         TabIndex        =   28
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
         Left            =   30
         Picture         =   "Blackout.frx":664C
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   31
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
            TabIndex        =   32
            Top             =   390
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
         TabIndex        =   29
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox pbcDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6075
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1425
      Visible         =   0   'False
      Width           =   210
      Begin VB.CheckBox ckcDay 
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
         Left            =   15
         TabIndex        =   36
         Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   6540
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   30
         Picture         =   "Blackout.frx":9466
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   30
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
            Picture         =   "Blackout.frx":A124
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.ListBox lbcCart 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2520
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1905
      Visible         =   0   'False
      Width           =   5070
   End
   Begin VB.ListBox lbcShtTitle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3165
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2325
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcSAdvt 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   375
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2835
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
      Enabled         =   0   'False
      Height          =   285
      HelpContextID   =   5
      Left            =   6120
      TabIndex        =   42
      Top             =   5385
      Width           =   945
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      HelpContextID   =   3
      Left            =   4830
      TabIndex        =   41
      Top             =   5385
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      HelpContextID   =   2
      Left            =   3585
      TabIndex        =   40
      Top             =   5385
      Width           =   945
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      HelpContextID   =   1
      Left            =   2280
      TabIndex        =   39
      Top             =   5385
      Width           =   945
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
      Left            =   1560
      Picture         =   "Blackout.frx":A42E
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2145
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
      Left            =   255
      MaxLength       =   20
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1575
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2430
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.PictureBox pbcClickFocus 
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
      Height          =   105
      Left            =   690
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5775
      Width           =   105
   End
   Begin VB.PictureBox pbcTab 
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
      Height          =   135
      Left            =   15
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   37
      Top             =   5475
      Width           =   105
   End
   Begin VB.PictureBox pbcSTab 
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
      Height          =   135
      Left            =   45
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   4
      Top             =   405
      Width           =   105
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   45
      ScaleHeight     =   270
      ScaleWidth      =   1305
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   1305
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3735
      Top             =   210
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3255
      Top             =   210
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
      Left            =   8130
      TabIndex        =   46
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
      Left            =   8130
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4590
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
      Left            =   8100
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5310
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4380
      Top             =   180
   End
   Begin VB.PictureBox pbcReplacement 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4275
      Index           =   1
      Left            =   585
      Picture         =   "Blackout.frx":A528
      ScaleHeight     =   4275
      ScaleWidth      =   8730
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   435
      Visible         =   0   'False
      Width           =   8730
      Begin VB.Label lacRFrame 
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
         Index           =   1
         Left            =   -30
         TabIndex        =   47
         Top             =   675
         Visible         =   0   'False
         Width           =   8760
      End
   End
   Begin VB.PictureBox pbcReplacement 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4125
      Index           =   0
      Left            =   375
      Picture         =   "Blackout.frx":335DA
      ScaleHeight     =   4125
      ScaleWidth      =   8730
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   525
      Visible         =   0   'False
      Width           =   8730
      Begin VB.Label lacRFrame 
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
         Index           =   0
         Left            =   -30
         TabIndex        =   10
         Top             =   675
         Visible         =   0   'False
         Width           =   8760
      End
   End
   Begin VB.PictureBox pbcSuppression 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4275
      Index           =   1
      Left            =   360
      Picture         =   "Blackout.frx":5AFBC
      ScaleHeight     =   4275
      ScaleWidth      =   8730
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   465
      Visible         =   0   'False
      Width           =   8730
      Begin VB.Label lacSFrame 
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
         Index           =   1
         Left            =   15
         TabIndex        =   48
         Top             =   735
         Visible         =   0   'False
         Width           =   8760
      End
   End
   Begin VB.PictureBox pbcSuppression 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4125
      Index           =   0
      Left            =   225
      Picture         =   "Blackout.frx":D554E
      ScaleHeight     =   4125
      ScaleWidth      =   8730
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   615
      Visible         =   0   'False
      Width           =   8730
      Begin VB.Label lacSFrame 
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
         Index           =   0
         Left            =   15
         TabIndex        =   11
         Top             =   645
         Visible         =   0   'False
         Width           =   8760
      End
   End
   Begin VB.PictureBox plcBlackout 
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
      Height          =   4275
      Left            =   180
      ScaleHeight     =   4215
      ScaleWidth      =   9045
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   525
      Width           =   9105
      Begin VB.VScrollBar vbcSR 
         Height          =   4125
         LargeChange     =   19
         Left            =   8775
         Min             =   1
         TabIndex        =   38
         Top             =   75
         Value           =   1
         Width           =   240
      End
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   270
      Picture         =   "Blackout.frx":FCF30
      Top             =   315
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8805
      Picture         =   "Blackout.frx":FD23A
      Top             =   5190
      Width           =   480
   End
End
Attribute VB_Name = "Blackout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Blackout.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mPaintTitle                                                                           *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Blackout.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Blackout input screen code
Option Explicit
Option Compare Text

'
'smFromLog = Y (from Log) or N (From NY Export):  Test for Y
'            If Y, then imSRIndex = 1
'            If N, then imSRIndex = 0
'
'igView = 0 = Suppress selected
'igView = 1 = Replace selected
'

Dim imFirstActivate As Integer
Dim smFromLog As String * 1
Dim smSplitFill As String * 1   'Y or N
'Suppression
Dim tmSCtrls()  As FIELDAREA
Dim imLBSCtrls As Integer
Dim imSBoxNo As Integer   'Current Suppression Box
Dim imSRowNo As Integer
'Replacement
Dim tmRCtrls()  As FIELDAREA
Dim imLBRCtrls  As Integer
Dim imLBCtrls As Integer
Dim imRBoxNo As Integer   'Current Suppression Box
Dim imRRowNo As Integer
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
'Blackout
Dim hmBof As Integer    'Blackout file handle
Dim imBofRecLen As Integer        'Bof record length
Dim tmBof As BOF
'Prduct code file
Dim hmPrf As Integer 'Product file handle
Dim tmPrf As PRF        'PRF record image
Dim imPrfRecLen As Integer
Dim tmPrfSrchKey As LONGKEY0    'PRF key record image
'Short Title
Dim hmSif As Integer    'Short Title file handle
Dim tmSifSrchKey As LONGKEY0    'Bof key record image
Dim imSifRecLen As Integer        'Bof record length
Dim tmSif As SIF
'Inventory
Dim hmCif As Integer    'Inventory file handle
Dim tmCifSrchKey As LONGKEY0    'Cif key record image
Dim imCifRecLen As Integer        'Cif record length
Dim tmCif As CIF
'Copy Product
Dim hmCpf As Integer    'Copy Product file handle
Dim tmCpfSrchKey As LONGKEY0    'Cpf key record image
Dim imCpfRecLen As Integer        'Cpf record length
Dim tmCpf As CPF
'Rotation
Dim tmCrf As CRF            'CRF record image
Dim tmCrfSrchKey1 As CRFKEY1  'CRF key record image
Dim hmCrf As Integer        'CRF Handle
Dim imCrfRecLen As Integer      'CRF record length
'Instruction
Dim tmCnf As CNF            'CNF record image
Dim tmCnfSrchKey As CNFKEY0  'CNF key record image
Dim hmCnf As Integer        'CNF Handle
Dim imCnfRecLen As Integer      'CNF record length
'Media
Dim tmMcf As MCF            'CNF record image
Dim tmMcfSrchKey As INTKEY0  'CNF key record image
Dim hmMcf As Integer        'CNF Handle
Dim imMcfRecLen As Integer      'CNF record length
'Contract
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey As LONGKEY0  'CHF key record image
Dim hmCHF As Integer        'CHF Handle
Dim imCHFRecLen As Integer      'CHF record length
'Vehicle combo
Dim tmVsf As VSF            'VSF record image
Dim tmVsfSrchKey As LONGKEY0  'VSF key record image
Dim hmVsf As Integer        'VSF Handle
Dim imVsfRecLen As Integer      'VSF record length
Dim tmAdf As ADF    'Only used for Abbreivation
'Dim tmRec As LPOPREC
Dim tmCartCode() As SORTCODE
Dim smCartCodeTag As String
Dim smSShow() As String  'Values shown in Blackout area
Dim smSSave() As String  'Values saved (1=Advertiser Name; 2=Short Title or Product Or Replace Copy on Suppress; 3=Vehicle; 4=Start Date;
                         '5=End Date; 6=Start Time; 7=End Time;
                         '8=Replace Advertiser; 9=Length; 10=Days; 11=Suppress Cntr No; 12=Replace Cntr No
Dim imSSave() As Integer  'Values saved (1-7=Day)
Dim lmSSave() As Long       'Values Saved (1=Suppress Cntr Code; 2=Replace Cntr Code; 3=CifCode; 4=SifCode)
Dim smRShow() As String  'Values shown in Blackout area
Dim smRSave() As String  'Values saved (1=Advertiser; 2=Cart #; 3=Short Title or Product; 4=Product Protection 1;
                         '5=Product Protection 2; 6=Start Date; 7=End Date; 8=Start Time; 9=End Time;
                         '10=Days; 11=Cntr No)
Dim imRSave() As Integer  'Values saved (1-7 = Day)
Dim lmRSave() As Long     'Values saved (1=CifCode; 2=SifCode; 3=Cntr Code)
Dim imSChg As Integer  'True=value changed; False=No changes
Dim imRChg As Integer
'Vehicle
Dim imFirstTime As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSRIndex As Integer   '0=Blackout on Log = No; 1=Blackout on Log = Yes
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imSettingValue As Integer   'True=Don't enable any box with change
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imBypassFocus As Integer
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragShift As Integer
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imUpdateAllowed As Integer
Dim imShowHelpMsg As Integer    'True=Show help messages; False=Ignore help messages
Dim imIgnoreRightMove As Integer
Dim imButtonIndex As Integer
Dim lmNowDate As Long
Dim smNowDate As String
Dim tmShtTitleCode() As SORTCODE
Dim smShtTitleCodeTag As String

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

'Dim igShowHelpMsg As Integer    'True=Show help messages; False=Ignore help message system
Dim SADVTINDEX As Integer         'Blackout control/field
Dim SCNTRINDEX As Integer
Dim SRADVTINDEX As Integer         'Blackout control/field
Dim SRCNTRINDEX As Integer
Dim SRCARTINDEX As Integer
Dim SSHORTTITLEINDEX  As Integer
Dim SVEHINDEX  As Integer
Dim SLENINDEX As Integer
Dim SSTARTDATEINDEX  As Integer
Dim SENDDATEINDEX  As Integer
Dim SDAYINDEX  As Integer
Dim SSTARTTIMEINDEX  As Integer
Dim SENDTIMEINDEX  As Integer
Dim RADVTINDEX As Integer          'Blackout control/field
Dim RCNTRINDEX As Integer
Dim RCARTINDEX As Integer
Dim RSHORTTITLEINDEX As Integer     'Set in mInit
Dim RPPINDEX As Integer             'Set in mInit
Dim RVEHINDEX As Integer
Dim RSTARTDATEINDEX As Integer
Dim RENDDATEINDEX As Integer
Dim RDAYINDEX As Integer
Dim RSTARTTIMEINDEX As Integer
Dim RENDTIMEINDEX As Integer
Private Sub ckcDay_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    imSRowNo = -1
    mRSetShow imRBoxNo
    imRBoxNo = -1
    imRRowNo = -1
    pbcArrow.Visible = False
    lacSFrame(imSRIndex).Visible = False
    lacRFrame(imSRIndex).Visible = False
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
        If (imSBoxNo > 0) And (igView = 0) Then
            mSEnableBox imSBoxNo
        End If
        If (imRBoxNo > 0) And (igView = 1) Then
            mREnableBox imRBoxNo
        End If
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    imSRowNo = -1
    mRSetShow imRBoxNo
    imRBoxNo = -1
    imRRowNo = -1
    pbcArrow.Visible = False
    lacSFrame(imSRIndex).Visible = False
    lacRFrame(imSRIndex).Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_Click()
    If igView = 0 Then
        If smFromLog = "Y" Then
            Select Case imSBoxNo
                Case SADVTINDEX 'Advertiser
                    lbcSAdvt.Visible = Not lbcSAdvt.Visible
                Case SCNTRINDEX 'Advertiser
                    lbcSCntr.Visible = Not lbcSCntr.Visible
                Case SRADVTINDEX 'Advertiser
                    lbcSAdvt.Visible = Not lbcSAdvt.Visible
                Case SRCNTRINDEX 'Advertiser
                    lbcSRCntr.Visible = Not lbcSRCntr.Visible
                Case SRCARTINDEX 'Advertiser
                    lbcCart.Visible = Not lbcCart.Visible
                Case SVEHINDEX
                    lbcVehicle.Visible = Not lbcVehicle.Visible
                Case SLENINDEX
                    lbcLen.Visible = Not lbcLen.Visible
                Case SSTARTTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
                Case SSTARTDATEINDEX
                    plcCalendar.Visible = Not plcCalendar.Visible
                Case SENDDATEINDEX
                    plcCalendar.Visible = Not plcCalendar.Visible
                Case SDAYINDEX
                    lbcDays.Visible = Not lbcDays.Visible
                Case SSTARTTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
                Case SENDTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
            End Select
        Else
            Select Case imSBoxNo
                Case SADVTINDEX 'Advertiser
                    lbcSAdvt.Visible = Not lbcSAdvt.Visible
                Case SSHORTTITLEINDEX 'Advertiser
                    lbcShtTitle.Visible = Not lbcShtTitle.Visible
                Case SVEHINDEX
                    lbcVehicle.Visible = Not lbcVehicle.Visible
                Case SSTARTTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
                Case SSTARTDATEINDEX
                    plcCalendar.Visible = Not plcCalendar.Visible
                Case SENDDATEINDEX
                    plcCalendar.Visible = Not plcCalendar.Visible
                Case SSTARTTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
                Case SENDTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
            End Select
        End If
    Else
        If smFromLog = "Y" Then
            Select Case imRBoxNo
                Case RADVTINDEX 'Advertiser
                    lbcRAdvt.Visible = Not lbcRAdvt.Visible
                Case RCNTRINDEX 'Advertiser
                    lbcRCntr.Visible = Not lbcRCntr.Visible
                Case RCARTINDEX 'Advertiser
                    lbcCart.Visible = Not lbcCart.Visible
                Case RPPINDEX 'Advertiser
                    lbcComp(0).Visible = Not lbcComp(0).Visible
                Case RPPINDEX + 1 'Advertiser
                    lbcComp(1).Visible = Not lbcComp(1).Visible
                Case RVEHINDEX 'Advertiser
                    lbcVehicle.Visible = Not lbcVehicle.Visible
                Case RSTARTTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
                Case RSTARTDATEINDEX
                    plcCalendar.Visible = Not plcCalendar.Visible
                Case RENDDATEINDEX
                    plcCalendar.Visible = Not plcCalendar.Visible
                Case RDAYINDEX
                    lbcDays.Visible = Not lbcDays.Visible
                Case RSTARTTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
                Case RENDTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
            End Select
        Else
            Select Case imRBoxNo
                Case RADVTINDEX 'Advertiser
                    lbcRAdvt.Visible = Not lbcRAdvt.Visible
                Case RCARTINDEX 'Advertiser
                    lbcCart.Visible = Not lbcCart.Visible
                Case RSHORTTITLEINDEX 'Advertiser
                Case RPPINDEX 'Advertiser
                    lbcComp(0).Visible = Not lbcComp(0).Visible
                Case RPPINDEX + 1 'Advertiser
                    lbcComp(1).Visible = Not lbcComp(1).Visible
                Case RSTARTTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
                Case RSTARTDATEINDEX
                    plcCalendar.Visible = Not plcCalendar.Visible
                Case RENDDATEINDEX
                    plcCalendar.Visible = Not plcCalendar.Visible
                Case RSTARTTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
                Case RENDTIMEINDEX
                    plcTme.Visible = Not plcTme.Visible
            End Select
        End If
    End If
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcUndo_Click()
    Screen.MousePointer = vbHourglass
    If igView = 0 Then
        pbcSuppression(imSRIndex).Cls
        If smFromLog = "Y" Then
            If Not gReadBofRec(1, hmBof, hmCif, hmPrf, hmSif, hmCHF, "S", smNowDate, 0) Then
                GoTo cmcUndoErr
                Exit Sub
            End If
        Else
            If Not gReadBofRec(0, hmBof, hmCif, hmPrf, hmSif, hmCHF, "S", smNowDate, 0) Then
                GoTo cmcUndoErr
                Exit Sub
            End If
        End If
        mMoveRecToCtrl "S"
        pbcSuppression_Paint imSRIndex
        imSChg = False
        imSBoxNo = -1
        imSRowNo = -1
    Else
        pbcReplacement(imSRIndex).Cls
        If smFromLog = "Y" Then
            If Not gReadBofRec(1, hmBof, hmCif, hmPrf, hmSif, hmCHF, "R", smNowDate, 0) Then
                GoTo cmcUndoErr
                Exit Sub
            End If
        Else
            If Not gReadBofRec(0, hmBof, hmCif, hmPrf, hmSif, hmCHF, "R", smNowDate, 0) Then
                GoTo cmcUndoErr
                Exit Sub
            End If
        End If
        mMoveRecToCtrl "R"
        pbcReplacement_Paint imSRIndex
        imRChg = False
        imRBoxNo = -1
        imRRowNo = -1
    End If
    'mInitBlackoutCtrls
    mSetCommands
    Screen.MousePointer = vbDefault
    Exit Sub
cmcUndoErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Sub cmcUndo_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    imSRowNo = -1
    mRSetShow imRBoxNo
    imRBoxNo = -1
    imRRowNo = -1
    pbcArrow.Visible = False
    lacSFrame(imSRIndex).Visible = False
    lacRFrame(imSRIndex).Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
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
        If (imSBoxNo > 0) And (igView = 0) Then
            mSEnableBox imSBoxNo
        End If
        If (imRBoxNo > 0) And (igView = 1) Then
            mREnableBox imRBoxNo
        End If
        Exit Sub
    End If
    imSBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
    imRBoxNo = -1
    imSChg = False
    imRChg = False
    'ReDim tgBofDel(1 To 1) As BOFREC
    ReDim tgBofDel(0 To 0) As BOFREC
    mSetCommands
    'pbcSTab.SetFocus
End Sub
Private Sub cmcUpdate_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    imSRowNo = -1
    mRSetShow imRBoxNo
    imRBoxNo = -1
    imRRowNo = -1
    pbcArrow.Visible = False
    lacSFrame(imSRIndex).Visible = False
    lacRFrame(imSRIndex).Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUpdate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    If igView = 0 Then
        If smFromLog = "Y" Then
            Select Case imSBoxNo
                Case SADVTINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcSAdvt, imBSMode, imComboBoxIndex
                Case SCNTRINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcSCntr, imBSMode, imComboBoxIndex
                Case SRADVTINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcSAdvt, imBSMode, imComboBoxIndex
                Case SRCNTRINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcSRCntr, imBSMode, imComboBoxIndex
                Case SVEHINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
                Case SLENINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcLen, imBSMode, imComboBoxIndex
                Case SSTARTDATEINDEX
                    slStr = edcDropDown.Text
                    If Not gValidDate(slStr) Then
                        lacDate.Visible = False
                        Exit Sub
                    End If
                    lacDate.Visible = True
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint   'mBoxCalDate called within paint
                Case SENDDATEINDEX
                    slStr = edcDropDown.Text
                    If Not gValidDate(slStr) Then
                        lacDate.Visible = False
                        Exit Sub
                    End If
                    lacDate.Visible = True
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint   'mBoxCalDate called within paint
                Case SDAYINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcDays, imBSMode, imComboBoxIndex
                Case SSTARTTIMEINDEX
                Case SENDTIMEINDEX
            End Select
        Else
            Select Case imSBoxNo
                Case SADVTINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcSAdvt, imBSMode, imComboBoxIndex
                Case SSHORTTITLEINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcShtTitle, imBSMode, imComboBoxIndex
                Case SVEHINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
                Case SSTARTDATEINDEX
                    slStr = edcDropDown.Text
                    If Not gValidDate(slStr) Then
                        lacDate.Visible = False
                        Exit Sub
                    End If
                    lacDate.Visible = True
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint   'mBoxCalDate called within paint
                Case SENDDATEINDEX
                    slStr = edcDropDown.Text
                    If Not gValidDate(slStr) Then
                        lacDate.Visible = False
                        Exit Sub
                    End If
                    lacDate.Visible = True
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint   'mBoxCalDate called within paint
                Case SSTARTTIMEINDEX
                Case SENDTIMEINDEX
            End Select
        End If
    Else
        If smFromLog = "Y" Then
            Select Case imRBoxNo
                Case RADVTINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcRAdvt, imBSMode, imComboBoxIndex
                Case RCNTRINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcRCntr, imBSMode, imComboBoxIndex
                Case RCARTINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcCart, imBSMode, imComboBoxIndex
                Case RPPINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcComp(0), imBSMode, imComboBoxIndex
                Case RPPINDEX + 1
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcComp(1), imBSMode, imComboBoxIndex
                Case RVEHINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
                Case RSTARTDATEINDEX
                    slStr = edcDropDown.Text
                    If Not gValidDate(slStr) Then
                        lacDate.Visible = False
                        Exit Sub
                    End If
                    lacDate.Visible = True
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint   'mBoxCalDate called within paint
                Case RENDDATEINDEX
                    slStr = edcDropDown.Text
                    If Not gValidDate(slStr) Then
                        lacDate.Visible = False
                        Exit Sub
                    End If
                    lacDate.Visible = True
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint   'mBoxCalDate called within paint
                Case RDAYINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcDays, imBSMode, imComboBoxIndex
                Case RSTARTTIMEINDEX
                Case RENDTIMEINDEX
            End Select
        Else
            Select Case imRBoxNo
                Case RADVTINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcRAdvt, imBSMode, imComboBoxIndex
                Case RCARTINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcCart, imBSMode, imComboBoxIndex
                Case RSHORTTITLEINDEX
                    'imLbcArrowSetting = True
                    'gMatchLookAhead edcDropDown, lbcShtTitle, imBSMode, imComboBoxIndex
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
                Case RPPINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcComp(0), imBSMode, imComboBoxIndex
                Case RPPINDEX + 1
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcComp(1), imBSMode, imComboBoxIndex
                Case RSTARTDATEINDEX
                    slStr = edcDropDown.Text
                    If Not gValidDate(slStr) Then
                        lacDate.Visible = False
                        Exit Sub
                    End If
                    lacDate.Visible = True
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint   'mBoxCalDate called within paint
                Case RENDDATEINDEX
                    slStr = edcDropDown.Text
                    If Not gValidDate(slStr) Then
                        lacDate.Visible = False
                        Exit Sub
                    End If
                    lacDate.Visible = True
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint   'mBoxCalDate called within paint
                Case RSTARTTIMEINDEX
                Case RENDTIMEINDEX
            End Select
        End If
    End If
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcDropDown_GotFocus()
    If igView = 0 Then
        If smFromLog = "Y" Then
            Select Case imSBoxNo
                Case SADVTINDEX
                    imComboBoxIndex = lbcSAdvt.ListIndex
                Case SCNTRINDEX
                    imComboBoxIndex = lbcSCntr.ListIndex
                Case SRADVTINDEX
                    imComboBoxIndex = lbcSAdvt.ListIndex
                Case SRCNTRINDEX
                    imComboBoxIndex = lbcSRCntr.ListIndex
                Case SVEHINDEX
                    imComboBoxIndex = lbcVehicle.ListIndex
                Case SLENINDEX
                    imComboBoxIndex = lbcLen.ListIndex
                Case SSTARTDATEINDEX
                Case SENDDATEINDEX
                Case SDAYINDEX
                    imComboBoxIndex = lbcDays.ListIndex
                Case SSTARTTIMEINDEX
                Case SENDTIMEINDEX
            End Select
        Else
            Select Case imSBoxNo
                Case SADVTINDEX
                    imComboBoxIndex = lbcSAdvt.ListIndex
                Case SSHORTTITLEINDEX
                    imComboBoxIndex = lbcShtTitle.ListIndex
                Case SVEHINDEX
                    imComboBoxIndex = lbcVehicle.ListIndex
                Case SSTARTDATEINDEX
                Case SENDDATEINDEX
                Case SSTARTTIMEINDEX
                Case SENDTIMEINDEX
            End Select
        End If
    Else
        If smFromLog = "Y" Then
            Select Case imRBoxNo
                Case RADVTINDEX
                    imComboBoxIndex = lbcRAdvt.ListIndex
                Case RCNTRINDEX
                    imComboBoxIndex = lbcRCntr.ListIndex
                Case RCARTINDEX
                    imComboBoxIndex = lbcCart.ListIndex
                Case RPPINDEX
                    imComboBoxIndex = lbcComp(0).ListIndex
                Case RPPINDEX + 1
                    imComboBoxIndex = lbcComp(1).ListIndex
                Case RVEHINDEX
                    imComboBoxIndex = lbcVehicle.ListIndex
                Case RSTARTDATEINDEX
                Case RENDDATEINDEX
                Case RDAYINDEX
                    imComboBoxIndex = lbcDays.ListIndex
                Case RSTARTTIMEINDEX
                Case RENDTIMEINDEX
            End Select
        Else
            Select Case imRBoxNo
                Case RADVTINDEX
                    imComboBoxIndex = lbcRAdvt.ListIndex
                Case RCARTINDEX
                    imComboBoxIndex = lbcCart.ListIndex
                Case RSHORTTITLEINDEX
                    'imComboBoxIndex = lbcShtTitle.ListIndex
                    imComboBoxIndex = lbcVehicle.ListIndex
                Case RPPINDEX
                    imComboBoxIndex = lbcComp(0).ListIndex
                Case RPPINDEX + 1
                    imComboBoxIndex = lbcComp(1).ListIndex
                Case RSTARTDATEINDEX
                Case RENDDATEINDEX
                Case RSTARTTIMEINDEX
                Case RENDTIMEINDEX
            End Select
        End If
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
    Dim ilKey As Integer

    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    If igView = 0 Then
        If smFromLog = "Y" Then
            Select Case imSBoxNo
                Case SADVTINDEX
                    'ilKey = KeyAscii
                    'If Not gCheckKeyAscii(ilKey) Then
                    '    KeyAscii = 0
                    '    Exit Sub
                    'End If
                Case SCNTRINDEX
                    'ilKey = KeyAscii
                    'If Not gCheckKeyAscii(ilKey) Then
                    '    KeyAscii = 0
                    '    Exit Sub
                    'End If
                Case SRADVTINDEX
                    'ilKey = KeyAscii
                    'If Not gCheckKeyAscii(ilKey) Then
                    '    KeyAscii = 0
                    '    Exit Sub
                    'End If
                Case SRCNTRINDEX
                    'ilKey = KeyAscii
                    'If Not gCheckKeyAscii(ilKey) Then
                    '    KeyAscii = 0
                    '    Exit Sub
                    'End If
                Case SVEHINDEX
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case SLENINDEX
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case SSTARTDATEINDEX
                    'Filter characters (allow only BackSpace, numbers 0 thru 9
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case SENDDATEINDEX
                    'Disallow TFN for alternate
                    If (Len(edcDropDown.Text) = edcDropDown.SelLength) Then
                        If (KeyAscii = Asc("T")) Or (KeyAscii = Asc("t")) Then
                            edcDropDown.Text = "TFN"
                            edcDropDown.SelStart = 0
                            edcDropDown.SelLength = 3
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
                Case SDAYINDEX
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case SSTARTTIMEINDEX
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
                Case SENDTIMEINDEX
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
        Else
            Select Case imSBoxNo
                Case SADVTINDEX
                    'ilKey = KeyAscii
                    'If Not gCheckKeyAscii(ilKey) Then
                    '    KeyAscii = 0
                    '    Exit Sub
                    'End If
                Case SSHORTTITLEINDEX
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case SVEHINDEX
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case SSTARTDATEINDEX
                    'Filter characters (allow only BackSpace, numbers 0 thru 9
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case SENDDATEINDEX
                    'Disallow TFN for alternate
                    If (Len(edcDropDown.Text) = edcDropDown.SelLength) Then
                        If (KeyAscii = Asc("T")) Or (KeyAscii = Asc("t")) Then
                            edcDropDown.Text = "TFN"
                            edcDropDown.SelStart = 0
                            edcDropDown.SelLength = 3
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
                Case SSTARTTIMEINDEX
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
                Case SENDTIMEINDEX
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
        End If
    Else
        If smFromLog = "Y" Then
            Select Case imRBoxNo
                Case RADVTINDEX
                    'ilKey = KeyAscii
                    'If Not gCheckKeyAscii(ilKey) Then
                    '    KeyAscii = 0
                    '    Exit Sub
                    'End If
                Case RCNTRINDEX
                    'ilKey = KeyAscii
                    'If Not gCheckKeyAscii(ilKey) Then
                    '    KeyAscii = 0
                    '    Exit Sub
                    'End If
                Case RCARTINDEX
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RPPINDEX
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RPPINDEX + 1
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RVEHINDEX
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RSTARTDATEINDEX
                    'Filter characters (allow only BackSpace, numbers 0 thru 9
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RENDDATEINDEX
                    'Filter characters (allow only BackSpace, numbers 0 thru 9
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RDAYINDEX
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RSTARTTIMEINDEX
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
                Case RENDTIMEINDEX
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
        Else
            Select Case imRBoxNo
                Case RADVTINDEX
                    'ilKey = KeyAscii
                    'If Not gCheckKeyAscii(ilKey) Then
                    '    KeyAscii = 0
                    '    Exit Sub
                    'End If
                Case RCARTINDEX
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RSHORTTITLEINDEX
                Case RPPINDEX
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RPPINDEX + 1
                    ilKey = KeyAscii
                    If Not gCheckKeyAscii(ilKey) Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RSTARTDATEINDEX
                    'Filter characters (allow only BackSpace, numbers 0 thru 9
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RENDDATEINDEX
                    'Filter characters (allow only BackSpace, numbers 0 thru 9
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case RSTARTTIMEINDEX
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
                Case RENDTIMEINDEX
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
        End If
    End If

    'ilKey = KeyAscii
    'If Not gCheckKeyAscii(ilKey) Then
    '    KeyAscii = 0
    '    Exit Sub
    'End If
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If igView = 0 Then
        If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
            If smFromLog = "Y" Then
                Select Case imSBoxNo
                    Case SADVTINDEX
                        gProcessArrowKey Shift, KeyCode, lbcSAdvt, imLbcArrowSetting
                    Case SCNTRINDEX
                        gProcessArrowKey Shift, KeyCode, lbcSCntr, imLbcArrowSetting
                    Case SRADVTINDEX
                        gProcessArrowKey Shift, KeyCode, lbcSAdvt, imLbcArrowSetting
                    Case SRCNTRINDEX
                        gProcessArrowKey Shift, KeyCode, lbcSRCntr, imLbcArrowSetting
                    Case SVEHINDEX
                        gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
                    Case SLENINDEX
                        gProcessArrowKey Shift, KeyCode, lbcLen, imLbcArrowSetting
                    Case SSTARTDATEINDEX
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
                    Case SENDDATEINDEX
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
                    Case SDAYINDEX
                        gProcessArrowKey Shift, KeyCode, lbcDays, imLbcArrowSetting
                    Case SSTARTTIMEINDEX
                        If (Shift And vbAltMask) > 0 Then
                            plcTme.Visible = Not plcTme.Visible
                        End If
                    Case SENDTIMEINDEX
                        If (Shift And vbAltMask) > 0 Then
                            plcTme.Visible = Not plcTme.Visible
                        End If
                End Select
            Else
                Select Case imSBoxNo
                    Case SADVTINDEX
                        gProcessArrowKey Shift, KeyCode, lbcSAdvt, imLbcArrowSetting
                    Case SSHORTTITLEINDEX
                        gProcessArrowKey Shift, KeyCode, lbcShtTitle, imLbcArrowSetting
                    Case SVEHINDEX
                        gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
                    Case SSTARTDATEINDEX
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
                    Case SENDDATEINDEX
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
                    Case SSTARTTIMEINDEX
                        If (Shift And vbAltMask) > 0 Then
                            plcTme.Visible = Not plcTme.Visible
                        End If
                    Case SENDTIMEINDEX
                        If (Shift And vbAltMask) > 0 Then
                            plcTme.Visible = Not plcTme.Visible
                        End If
                End Select
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
        End If
        If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
            Select Case imSBoxNo
                Case SSTARTDATEINDEX
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
                Case SENDDATEINDEX
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
                Case SSTARTTIMEINDEX
                Case SENDTIMEINDEX
            End Select
        End If
    Else
        If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
            If smFromLog = "Y" Then
                Select Case imRBoxNo
                    Case RADVTINDEX
                        gProcessArrowKey Shift, KeyCode, lbcRAdvt, imLbcArrowSetting
                    Case RCNTRINDEX
                        gProcessArrowKey Shift, KeyCode, lbcRCntr, imLbcArrowSetting
                    Case RCARTINDEX
                        gProcessArrowKey Shift, KeyCode, lbcCart, imLbcArrowSetting
                    Case RVEHINDEX
                        gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
                    Case RPPINDEX
                        gProcessArrowKey Shift, KeyCode, lbcComp(0), imLbcArrowSetting
                    Case RPPINDEX + 1
                        gProcessArrowKey Shift, KeyCode, lbcComp(1), imLbcArrowSetting
                    Case RSTARTDATEINDEX
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
                    Case RENDDATEINDEX
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
                    Case RDAYINDEX
                        gProcessArrowKey Shift, KeyCode, lbcDays, imLbcArrowSetting
                    Case RSTARTTIMEINDEX
                        If (Shift And vbAltMask) > 0 Then
                            plcTme.Visible = Not plcTme.Visible
                        End If
                    Case RENDTIMEINDEX
                        If (Shift And vbAltMask) > 0 Then
                            plcTme.Visible = Not plcTme.Visible
                        End If
                End Select
            Else
                Select Case imRBoxNo
                    Case RADVTINDEX
                        gProcessArrowKey Shift, KeyCode, lbcRAdvt, imLbcArrowSetting
                    Case RCARTINDEX
                        gProcessArrowKey Shift, KeyCode, lbcCart, imLbcArrowSetting
                    Case RSHORTTITLEINDEX
                        'gProcessArrowKey Shift, KeyCode, lbcShtTitle, imLbcArrowSetting
                        gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
                    Case RPPINDEX
                        gProcessArrowKey Shift, KeyCode, lbcComp(0), imLbcArrowSetting
                    Case RPPINDEX + 1
                        gProcessArrowKey Shift, KeyCode, lbcComp(1), imLbcArrowSetting
                    Case RSTARTDATEINDEX
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
                    Case RENDDATEINDEX
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
                    Case RSTARTTIMEINDEX
                        If (Shift And vbAltMask) > 0 Then
                            plcTme.Visible = Not plcTme.Visible
                        End If
                    Case RENDTIMEINDEX
                        If (Shift And vbAltMask) > 0 Then
                            plcTme.Visible = Not plcTme.Visible
                        End If
                End Select
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
        End If
        If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
            Select Case imRBoxNo
                Case RSTARTDATEINDEX
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
                Case RENDDATEINDEX
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
                Case RSTARTTIMEINDEX
                Case RENDTIMEINDEX
            End Select
        End If
    End If
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        'If igView = 0 Then
        '    Select Case imSBoxNo
        '        Case SADVTINDEX
        '            If imTabDirection = -1 Then  'Right To Left
        '                pbcSTab.SetFocus
        '            Else
        '                pbcTab.SetFocus
        '            End If
        '            Exit Sub
        '        Case SSHORTTITLEINDEX
        '            If imTabDirection = -1 Then  'Right To Left
        '                pbcSTab.SetFocus
        '            Else
        '                pbcTab.SetFocus
        '            End If
        '            Exit Sub
        '        Case SVEHINDEX
        '            If imTabDirection = -1 Then  'Right To Left
        '                pbcSTab.SetFocus
        '            Else
        '                pbcTab.SetFocus
        '            End If
        '            Exit Sub
        '    End Select
        'Else
        '    Select Case imSBoxNo
        '        Case RADVTINDEX
        '            If imTabDirection = -1 Then  'Right To Left
        '                pbcSTab.SetFocus
        '            Else
        '                pbcTab.SetFocus
        '            End If
        '            Exit Sub
        '        Case RSHORTTITLEINDEX
        '            If imTabDirection = -1 Then  'Right To Left
        '                pbcSTab.SetFocus
        '            Else
        '                pbcTab.SetFocus
        '            End If
        '            Exit Sub
        '        Case RCARTINDEX
        '            If imTabDirection = -1 Then  'Right To Left
        '                pbcSTab.SetFocus
        '            Else
        '                pbcTab.SetFocus
        '            End If
        '            Exit Sub
        '        Case RPPINDEX
        '            If imTabDirection = -1 Then  'Right To Left
        '                pbcSTab.SetFocus
        '            Else
        '                pbcTab.SetFocus
        '            End If
        '            Exit Sub
        '        Case RPPINDEX + 1
        '            If imTabDirection = -1 Then  'Right To Left
        '                pbcSTab.SetFocus
        '            Else
        '                pbcTab.SetFocus
        '            End If
        '            Exit Sub
        '    End Select
        'End If
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
        Me.KeyPreview = True
        Exit Sub
    End If
    If smSplitFill = "Y" Then
        cmcCancel.SetFocus
    End If
    imFirstActivate = False
    'If (igWinStatus(BUDGETSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        'pbcSuppression.Enabled = False
        'pbcReplacement.Enabled = False
        'pbcSTab.Enabled = False
        'pbcTab.Enabled = False
        'imUpdateAllowed = False
    'Else
        pbcSuppression(imSRIndex).Enabled = True
        pbcReplacement(imSRIndex).Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    'End If
    Me.KeyPreview = True
    Blackout.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcTme.Visible = False
        plcSelect.Enabled = False
        gFunctionKeyBranch KeyCode
        If imSBoxNo > 0 Then
            mSEnableBox imSBoxNo
        ElseIf imRBoxNo > 0 Then
            mREnableBox imRBoxNo
        End If
        plcSelect.Visible = False
        plcSelect.Visible = True
        vbcSR.Visible = False
        vbcSR.Visible = True
        'If rbcSR(0).Value Then
        '    rbcSR(1).Visible = False
        '    rbcSR(0).Visible = False
        '    vbcSR.Visible = False
        '    rbcSR(0).Visible = True
        '    rbcSR(1).Visible = True
        '    vbcSR.Visible = True
        'Else
        '    rbcSR(0).Visible = False
        '    rbcSR(1).Visible = False
        '    vbcSR.Visible = False
        '    rbcSR(1).Visible = True
        '    rbcSR(0).Visible = True
        '    vbcSR.Visible = True
        'End If
        plcSelect.Enabled = True
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
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100) / Me.height
        Me.height = (lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100
    End If
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
    Dim ilRet As Integer

    On Error Resume Next
    
    If Not igManUnload Then
        mSSetShow imSBoxNo
        imSBoxNo = -1
        mRSetShow imRBoxNo
        imRBoxNo = -1
        pbcArrow.Visible = False
        lacSFrame(imSRIndex).Visible = False
        lacRFrame(imSRIndex).Visible = False
        If mSaveRecChg(True) = False Then
            If Not imTerminate Then
                If (imSBoxNo <> -1) And (igView = 0) Then
                    mSEnableBox imSBoxNo
                End If
                If (imRBoxNo <> -1) And (igView = 1) Then
                    mREnableBox imRBoxNo
                End If
                Cancel = 1
                igStopCancel = True
                Exit Sub
            End If
        End If
    End If
    
    btrExtClear hmVsf   'Clear any previous extend operation
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    btrExtClear hmCHF   'Clear any previous extend operation
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    btrExtClear hmMcf   'Clear any previous extend operation
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    btrExtClear hmCnf   'Clear any previous extend operation
    ilRet = btrClose(hmCnf)
    btrDestroy hmCnf
    btrExtClear hmCrf   'Clear any previous extend operation
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    btrExtClear hmCpf   'Clear any previous extend operation
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    btrExtClear hmCif   'Clear any previous extend operation
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    btrExtClear hmSif   'Clear any previous extend operation
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    btrExtClear hmPrf   'Clear any previous extend operation
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    btrExtClear hmBof   'Clear any previous extend operation
    ilRet = btrClose(hmBof)
    btrDestroy hmBof
    Erase tmSCtrls
    Erase tmRCtrls
    Erase smSShow
    Erase smSSave
    Erase imSSave
    Erase lmSSave
    Erase smRShow
    Erase smRSave
    Erase imRSave
    Erase lmRSave
    Erase tmCartCode
    Erase tgSBofRec
    Erase tgRBofRec
    Erase tgBofDel
    Erase tgManager
    Erase tgPlanner
    Erase tgCntrCode
    Erase tmShtTitleCode

    Set Blackout = Nothing   'Remove data segment
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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

Private Sub imcTrash_Click()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilUpperBound As Integer
    Dim ilRowNo As Integer
    Dim ilBof As Integer
    If igView = 0 Then
        If (imSRowNo < vbcSR.Value) Or (imSRowNo > vbcSR.Value + vbcSR.LargeChange) Then
            Exit Sub
        End If
        ilRowNo = imSRowNo
        mSSetShow imSBoxNo
        imSBoxNo = -1
        imSRowNo = -1
        pbcArrow.Visible = False
        lacSFrame(imSRIndex).Visible = False
        gCtrlGotFocus ActiveControl
        ilUpperBound = UBound(smSSave, 2)
        ilBof = ilRowNo
        If ilBof = ilUpperBound Then
            mInitNew ilBof
        Else
            If ilBof > 0 Then
                If tgSBofRec(ilBof).iStatus = 1 Then
                    tgBofDel(UBound(tgBofDel)).tBof = tgSBofRec(ilBof).tBof
                    tgBofDel(UBound(tgBofDel)).iStatus = tgSBofRec(ilBof).iStatus
                    tgBofDel(UBound(tgBofDel)).lRecPos = tgSBofRec(ilBof).lRecPos
                    'ReDim Preserve tgBofDel(1 To UBound(tgBofDel) + 1) As BOFREC
                    ReDim Preserve tgBofDel(0 To UBound(tgBofDel) + 1) As BOFREC
                End If
                'Remove record from tgRjf1Rec- Leave tgPjf2Rec
                For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
                    tgSBofRec(ilLoop) = tgSBofRec(ilLoop + 1)
                Next ilLoop
                'ReDim Preserve tgSBofRec(1 To UBound(tgSBofRec) - 1) As BOFREC
                ReDim Preserve tgSBofRec(0 To UBound(tgSBofRec) - 1) As BOFREC
            End If
            For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
                For ilIndex = 1 To UBound(smSSave, 1) Step 1
                    smSSave(ilIndex, ilLoop) = smSSave(ilIndex, ilLoop + 1)
                Next ilIndex
                For ilIndex = 1 To UBound(imSSave, 1) Step 1
                    imSSave(ilIndex, ilLoop) = imSSave(ilIndex, ilLoop + 1)
                Next ilIndex
                For ilIndex = 1 To UBound(lmSSave, 1) Step 1
                    lmSSave(ilIndex, ilLoop) = lmSSave(ilIndex, ilLoop + 1)
                Next ilIndex
                For ilIndex = 1 To UBound(smSShow, 1) Step 1
                    smSShow(ilIndex, ilLoop) = smSShow(ilIndex, ilLoop + 1)
                Next ilIndex
            Next ilLoop
            ilUpperBound = UBound(smSSave, 2)
            'ReDim Preserve smSShow(1 To 14, 1 To ilUpperBound - 1) As String 'Values shown in program area
            'ReDim Preserve smSSave(1 To 12, 1 To ilUpperBound - 1) As String    'Values saved (program name) in program area
            'ReDim Preserve imSSave(1 To 7, 1 To ilUpperBound - 1) As Integer 'Values saved (program name) in program area
            'ReDim Preserve lmSSave(1 To 4, 1 To ilUpperBound - 1) As Long 'Values saved (program name) in program area
            ReDim Preserve smSShow(0 To 14, 0 To ilUpperBound - 1) As String 'Values shown in program area
            ReDim Preserve smSSave(0 To 12, 0 To ilUpperBound - 1) As String    'Values saved (program name) in program area
            ReDim Preserve imSSave(0 To 7, 0 To ilUpperBound - 1) As Integer 'Values saved (program name) in program area
            ReDim Preserve lmSSave(0 To 4, 0 To ilUpperBound - 1) As Long 'Values saved (program name) in program area
            
            imSChg = True
        End If
        mSetCommands
        lacSFrame(imSRIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
        mSetMinMax
        'pbcSuppression.Cls
        'pbcSuppression_Paint imSRIndex
    Else
        If (imRRowNo < vbcSR.Value) Or (imRRowNo > vbcSR.Value + vbcSR.LargeChange) Then
            Exit Sub
        End If
        ilRowNo = imRRowNo
        mRSetShow imRBoxNo
        imRBoxNo = -1
        imRRowNo = -1
        pbcArrow.Visible = False
        lacRFrame(imSRIndex).Visible = False
        gCtrlGotFocus ActiveControl
        ilUpperBound = UBound(smRSave, 2)
        ilBof = ilRowNo
        If ilBof = ilUpperBound Then
            mInitNew ilBof
        Else
            If ilBof > 0 Then
                If tgRBofRec(ilBof).iStatus = 1 Then
                    tgBofDel(UBound(tgBofDel)).tBof = tgRBofRec(ilBof).tBof
                    tgBofDel(UBound(tgBofDel)).iStatus = tgRBofRec(ilBof).iStatus
                    tgBofDel(UBound(tgBofDel)).lRecPos = tgRBofRec(ilBof).lRecPos
                    'ReDim Preserve tgBofDel(1 To UBound(tgBofDel) + 1) As BOFREC
                    ReDim Preserve tgBofDel(0 To UBound(tgBofDel) + 1) As BOFREC
                End If
                'Remove record from tgRjf1Rec- Leave tgPjf2Rec
                For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
                    tgRBofRec(ilLoop) = tgRBofRec(ilLoop + 1)
                Next ilLoop
                'ReDim Preserve tgRBofRec(1 To UBound(tgRBofRec) - 1) As BOFREC
                ReDim Preserve tgRBofRec(0 To UBound(tgRBofRec) - 1) As BOFREC
            End If
            For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
                For ilIndex = 1 To UBound(smRSave, 1) Step 1
                    smRSave(ilIndex, ilLoop) = smRSave(ilIndex, ilLoop + 1)
                Next ilIndex
                For ilIndex = 1 To UBound(imRSave, 1) Step 1
                    imRSave(ilIndex, ilLoop) = imRSave(ilIndex, ilLoop + 1)
                Next ilIndex
                For ilIndex = 1 To UBound(lmRSave, 1) Step 1
                    lmRSave(ilIndex, ilLoop) = lmRSave(ilIndex, ilLoop + 1)
                Next ilIndex
                For ilIndex = 1 To UBound(smRShow, 1) Step 1
                    smRShow(ilIndex, ilLoop) = smRShow(ilIndex, ilLoop + 1)
                Next ilIndex
            Next ilLoop
            ilUpperBound = UBound(smRSave, 2)
'            ReDim Preserve smRShow(1 To 16, 1 To ilUpperBound - 1) As String 'Values shown in program area
'            ReDim Preserve smRSave(1 To 11, 1 To ilUpperBound - 1) As String     'Values saved (program name) in program area
'            ReDim Preserve imRSave(1 To 7, 1 To ilUpperBound - 1) As Integer 'Values saved (program name) in program area
'            ReDim Preserve lmRSave(1 To 3, 1 To ilUpperBound - 1) As Long 'Values saved (program name) in program area
            ReDim Preserve smRShow(0 To 16, 0 To ilUpperBound - 1) As String 'Values shown in program area
            ReDim Preserve smRSave(0 To 11, 0 To ilUpperBound - 1) As String     'Values saved (program name) in program area
            ReDim Preserve imRSave(0 To 7, 0 To ilUpperBound - 1) As Integer 'Values saved (program name) in program area
            ReDim Preserve lmRSave(0 To 3, 0 To ilUpperBound - 1) As Long 'Values saved (program name) in program area

            imRChg = True
        End If
        mSetCommands
        lacRFrame(imSRIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
        mSetMinMax
        'pbcReplacement.Cls
        'pbcReplacement_Paint
    End If
End Sub
Private Sub imcTrash_DragDrop(Source As control, X As Single, Y As Single)
'    lacRCFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        If igView = 0 Then
            lacSFrame(imSRIndex).DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        Else
            lacRFrame(imSRIndex).DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        End If
'        lacEvtFrame.DragIcon = IconTraf!imcIconTrash.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        If igView = 0 Then
            lacSFrame(imSRIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        Else
            lacRFrame(imSRIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        End If
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub lbcCart_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcCart, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcCart_DblClick()
    'tmcClick.Enabled = False
    'imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcCart_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcCart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcCart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcCart, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcComp_Click(Index As Integer)
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcComp(Index), edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcComp_DblClick(Index As Integer)
    'tmcClick.Enabled = False
    'imDoubleClickName = True    'Double click event is followed by a mouse up event
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
Private Sub lbcDays_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcDays, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcDays_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcDays_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcLen_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcLen, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcLen_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcLen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcRAdvt_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcRAdvt, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcRAdvt_DblClick()
    'tmcClick.Enabled = False
    'imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcRAdvt_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcRAdvt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcRAdvt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcRAdvt, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcRCntr_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcRCntr, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcRCntr_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcRCntr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcSAdvt_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcSAdvt, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcSAdvt_DblClick()
    'tmcClick.Enabled = False
    'imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcSAdvt_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcSAdvt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcSAdvt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcSAdvt, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcSCntr_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcSCntr, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcSCntr_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcSCntr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcShtTitle_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcShtTitle, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcShtTitle_DblClick()
    'tmcClick.Enabled = False
    'imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcShtTitle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcShtTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcShtTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcShtTitle, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcSRCntr_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcSRCntr, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcSRCntr_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcSRCntr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcVehicle_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        imLbcArrowSetting = False
        gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcVehicle_DblClick()
    'tmcClick.Enabled = False
    'imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcVehicle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcVehicle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
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
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilIndex As Integer
    If rbcSR(0).Value Then
        ilIndex = lbcSAdvt.ListIndex
        If ilIndex >= 0 Then
            slName = lbcSAdvt.List(ilIndex)
        End If
    Else
        ilIndex = lbcRAdvt.ListIndex
        If ilIndex >= 0 Then
            slName = lbcRAdvt.List(ilIndex)
        End If
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(Copy, cbcAdvt, Traffic!lbcAdvt)
    ilRet = gPopAdvtBox(Blackout, lbcRAdvt, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", Blackout
        On Error GoTo 0
        For ilLoop = 0 To lbcRAdvt.ListCount - 1 Step 1
            lbcSAdvt.AddItem lbcRAdvt.List(ilLoop), ilLoop
        Next ilLoop
        If smFromLog = "Y" Then
            lbcSAdvt.AddItem "[None]", 0  'Force as first item on list
        End If
'        cbcAdvt.AddItem "[New]", 0  'Force as first item on list
        If rbcSR(0).Value Then
            If ilIndex >= 0 Then
                gFindMatch slName, 0, lbcSAdvt
                If gLastFound(lbcSAdvt) >= 0 Then
                    lbcSAdvt.ListIndex = gLastFound(lbcSAdvt)
                Else
                    lbcSAdvt.ListIndex = -1
                End If
            Else
                lbcSAdvt.ListIndex = ilIndex
            End If
        Else
            If ilIndex >= 0 Then
                gFindMatch slName, 0, lbcRAdvt
                If gLastFound(lbcRAdvt) >= 0 Then
                    lbcRAdvt.ListIndex = gLastFound(lbcRAdvt)
                Else
                    lbcRAdvt.ListIndex = -1
                End If
            Else
                lbcRAdvt.ListIndex = ilIndex
            End If
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
    slStr = edcDropDown.Text
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
'*      Procedure Name:mCartPop                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection Name    *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mCartPop(ilRowNo As Integer)
    Dim ilRet As Integer
    Dim ilISCIProd As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilAdfCode As Integer
    Dim llCifCode As Long
    Dim ilAdfIndex As Integer
    Dim llSifCode As Long
    Dim slShtTitle As String
    Dim ilCif As Integer
    Dim slSDate As String
    Dim slEDate As String
    Screen.MousePointer = vbHourglass
    If tgSpf.sUseCartNo <> "N" Then
        ilISCIProd = 1
    Else
        ilISCIProd = 4
    End If
    ilISCIProd = ilISCIProd Or &H100    '&H100=Include Length
    'If ilRowNo <= 0 Then
    If ilRowNo < 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If igView = 0 Then
        If smFromLog = "Y" Then
            slStr = Trim$(smSSave(8, ilRowNo))
            gFindMatch slStr, 0, lbcSAdvt
            If gLastFound(lbcSAdvt) <= 0 Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            ilAdfIndex = gLastFound(lbcSAdvt) - 1
        Else
            slStr = Trim$(smSSave(1, ilRowNo))
            gFindMatch slStr, 0, lbcSAdvt
            If gLastFound(lbcSAdvt) < 0 Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            ilAdfIndex = gLastFound(lbcSAdvt)
        End If
    Else
        slStr = Trim$(smRSave(1, ilRowNo))
        gFindMatch slStr, 0, lbcRAdvt
        If gLastFound(lbcRAdvt) < 0 Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        ilAdfIndex = gLastFound(lbcRAdvt)
    End If
    slNameCode = tgAdvertiser(ilAdfIndex).sKey    'Traffic!lbcAdvt.List(imAdvtIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilAdfCode = Val(slCode)   '
    'ilRet = gPopCopyForAdvtBox(Copy, ilAdvtCode, ilISCIProd, 0, lbcActive, lbcCartCode)
    ilRet = gPopCopyForAdvtBox(Blackout, ilAdfCode, ilISCIProd, 0, lbcCart, tmCartCode(), smCartCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mCartPopErr
        gCPErrorMsg ilRet, "mInvPop (gPopCopyForAdvtBox: Copy)", Blackout
        On Error GoTo 0
        If lbcCart.ListCount = 0 Then
            '9/5/06: Disallow [None] if from Log and SplitFill
            If smSplitFill <> "Y" Then
                lbcCart.AddItem "[None]"
            End If
        Else
            'Add Short Title to Cart Description
            For ilLoop = UBound(tmCartCode) - 1 To 0 Step -1
                'Add code to tmCartCode and Short Title to lbcCart
                slNameCode = tmCartCode(ilLoop).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                llCifCode = Val(slCode)
                tmCifSrchKey.lCode = llCifCode
                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = mGetShtTitle(ilAdfIndex, llCifCode, llSifCode, slShtTitle)
                Else
                    llSifCode = 0
                    slShtTitle = ""
                    ilRet = False
                End If
                'UseProdSptSrc = P +> Short Title; A= Advertiser/prod)
                If smFromLog = "Y" Then
                    gUnpackDate tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), slSDate
                    gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slEDate
                    If (tgSpf.sUseProdSptScr <> "P") Then
                        slShtTitle = ""
                        llSifCode = 0
                    End If
                    tmCartCode(ilLoop).sKey = Trim$(tmCartCode(ilLoop).sKey) & "\" & slShtTitle & "\" & Trim$(str$(llSifCode)) & "\" & slSDate & "\" & slEDate
                    lbcCart.List(ilLoop) = lbcCart.List(ilLoop) & slShtTitle & " " & slSDate & " " & slEDate
                Else
                    If (ilRet) And ((llSifCode <> 0) Or (tgSpf.sUseProdSptScr <> "P")) Then ' And (Trim$(slShtTitle) <> "")) Then
                        gUnpackDate tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), slSDate
                        gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slEDate
                        tmCartCode(ilLoop).sKey = Trim$(tmCartCode(ilLoop).sKey) & "\" & slShtTitle & "\" & Trim$(str$(llSifCode)) & "\" & slSDate & "\" & slEDate
                        lbcCart.List(ilLoop) = lbcCart.List(ilLoop) & " " & slShtTitle & " " & slSDate & " " & slEDate
                    Else
                        For ilCif = ilLoop To UBound(tmCartCode) - 1 Step 1
                            tmCartCode(ilCif) = tmCartCode(ilCif + 1)
                        Next ilCif
                        ReDim Preserve tmCartCode(0 To UBound(tmCartCode) - 1) As SORTCODE
                        lbcCart.RemoveItem ilLoop
                    End If
                End If
            Next ilLoop
            '9/5/06: Disallow [None] if from Log and SplitFill
            If (lbcCart.ListCount = 0) Or (imSRIndex = 1) Then
                If smSplitFill <> "Y" Then
                    lbcCart.AddItem "[None]", 0
                End If
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    'Replace Blank delimiter with "~"
    'If tgSpf.sUseCartNo <> "N" Then
    '    For ilLoop = 0 To UBound(tmCartCode) - 1 Step 1'lbcCartCode.ListCount - 1 Step 1
    '        slNameCode = tmCartCode(ilLoop).sKey  'lbcCartCode.List(ilLoop)
    '        ilFirstPos = InStr(slNameCode, " ")
    '        ilLastPos = 0
    '        ilPos = ilFirstPos + 1
    '        ilPos = InStr(ilPos, slNameCode, " ")
    '        'If rbcShow(0).Value Then
    '            ilLastPos = ilPos
    '        'Else
    '        '    Do While ilPos > 0
    '        '        ilLastPos = ilPos
    '        '        ilPos = ilPos + 1
    '        '        ilPos = InStr(ilPos, slNameCode, " ")
    '        '    Loop
    '        'End If
    '        If ilFirstPos > 0 Then
     '           Mid$(slNameCode, ilFirstPos, 1) = "~"
    '        End If
    '        If ilLastPos > 0 Then
    '            Mid$(slNameCode, ilLastPos, 1) = "~"
    '        End If
    '        'lbcCartCode.List(ilLoop) = slNameCode
    '        tmCartCode(ilLoop).sKey = slNameCode
    '    Next ilLoop
    '    For ilLoop = 0 To lbcCart.ListCount - 1 Step 1
    '        slName = lbcCart.List(ilLoop)
    '        ilFirstPos = InStr(slName, " ")
    '        ilLastPos = 0
    '        ilPos = ilFirstPos + 1
    '        ilPos = InStr(ilPos, slName, " ")
    '        'If rbcShow(0).Value Then
    '            ilLastPos = ilPos
    '        'Else
    '        '    Do While ilPos > 0
    '        '        ilLastPos = ilPos
    '        '        ilPos = ilPos + 1
    '        '        ilPos = InStr(ilPos, slName, " ")
    '        '    Loop
    '        'End If
    '        If ilFirstPos > 0 Then
    '            Mid$(slName, ilFirstPos, 1) = "~"
    '        End If
    '        If ilLastPos > 0 Then
    '            Mid$(slName, ilLastPos, 1) = "~"
    '        End If
    '        lbcCart.List(ilLoop) = slName
    '    Next ilLoop
    'End If
    Exit Sub
mCartPopErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

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
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    ReDim slComp(0 To 1) As String      'Competitive name, saved to determine if changed
    ReDim ilComp(0 To 1) As Integer      'Competitive name, saved to determine if changed
    If smFromLog = "Y" Then
        If smSplitFill = "Y" Then
            lbcComp(0).Clear
            lbcComp(0).AddItem "[None]", 0
            lbcComp(1).Clear
            lbcComp(1).AddItem "[None]", 0
            Exit Sub
        End If
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
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
    'ilRet = gIMoveListBox(Advt, lbcComp(0), lbcCompCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    ilRet = gIMoveListBox(Blackout, lbcComp(0), tgCompCode(), sgCompCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mCompPopErr
        gCPErrorMsg ilRet, "mCompPop (gIMoveListBox)", Blackout
        On Error GoTo 0
        lbcComp(0).AddItem "[None]", 0
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
'*      Procedure Name:mDaysPop                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection Name    *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mDaysPop()
    lbcDays.Clear
    lbcDays.AddItem "Mo"
    lbcDays.AddItem "Tu"
    lbcDays.AddItem "We"
    lbcDays.AddItem "Th"
    lbcDays.AddItem "Fr"
    lbcDays.AddItem "Sa"
    lbcDays.AddItem "Su"
    lbcDays.AddItem "Sa-Su"
    lbcDays.AddItem "Mo-Su"
    lbcDays.AddItem "Mo-Sa"
    lbcDays.AddItem "Mo-Fr"
    lbcDays.AddItem "Mo-Th"
    lbcDays.AddItem "Mo-We"
    lbcDays.AddItem "Mo-Tu"
    lbcDays.AddItem "Tu-Su"
    lbcDays.AddItem "Tu-Sa"
    lbcDays.AddItem "Tu-Fr"
    lbcDays.AddItem "Tu-Th"
    lbcDays.AddItem "Tu-We"
    lbcDays.AddItem "We-Su"
    lbcDays.AddItem "We-Sa"
    lbcDays.AddItem "We-Fr"
    lbcDays.AddItem "We-Th"
    lbcDays.AddItem "Th-Su"
    lbcDays.AddItem "Th-Sa"
    lbcDays.AddItem "Th-Fr"
    lbcDays.AddItem "Fr-Su"
    lbcDays.AddItem "Fr-Sa"
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetCntrPP                      *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mGetCntrPP(ilRowNo As Integer)
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    smRSave(4, ilRowNo) = ""
    smRSave(5, ilRowNo) = ""
    If lmRSave(3, ilRowNo) > 0 Then
        tmChfSrchKey.lCode = lmRSave(3, ilRowNo)
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mGetCntrPPErr
        gBtrvErrorMsg ilRet, "mGetCntrPP (btrGetEqual):" & "Chf.Btr", Blackout
        On Error GoTo 0
        If smSplitFill <> "Y" Then
            If tmChf.iMnfComp(0) > 0 Then
                For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                    slNameCode = tgCompCode(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) = tmChf.iMnfComp(0) Then
                        ilRet = gParseItem(slNameCode, 1, "\", smRSave(4, ilRowNo))
                        Exit For
                    End If
                Next ilLoop
            End If
            If tmChf.iMnfComp(1) > 0 Then
                For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                    slNameCode = tgCompCode(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) = tmChf.iMnfComp(1) Then
                        ilRet = gParseItem(slNameCode, 1, "\", smRSave(5, ilRowNo))
                        Exit For
                    End If
                Next ilLoop
            End If
        End If
    End If
    gSetShow pbcReplacement(imSRIndex), smRSave(4, ilRowNo), tmRCtrls(RPPINDEX)
    smRShow(RPPINDEX, ilRowNo) = tmRCtrls(RPPINDEX).sShow
    gSetShow pbcReplacement(imSRIndex), smRSave(5, ilRowNo), tmRCtrls(RPPINDEX + 1)
    smRShow(RPPINDEX + 1, ilRowNo) = tmRCtrls(RPPINDEX + 1).sShow
    Exit Sub
mGetCntrPPErr:
    On Error GoTo 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetShtTitle                    *
'*                                                     *
'*             Created:8/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain rotation specifications *
'*                                                     *
'*******************************************************
Private Function mGetShtTitle(ilAdfIndex As Integer, llCifCode As Long, llSifCode As Long, slShtTitle As String) As Integer
'
'   iRet = mGetShtTitle(ilRowNo)
'   Where:
'
    Dim ilRet As Integer    'Return status
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilAdfCode As Integer
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim slDate As String
    Dim ilFound As Integer
    Dim ilVsf As Integer
    slShtTitle = ""
    llSifCode = 0
    mGetShtTitle = False
    If (rbcSR(0).Value) And (smFromLog = "Y") Then
        If ilAdfIndex <= 0 Then
            Exit Function
        End If
        slNameCode = tgAdvertiser(ilAdfIndex - 1).sKey  'Traffic!lbcAdvt.List(imAdvtIndex)
    Else
        If ilAdfIndex < 0 Then
            Exit Function
        End If
        slNameCode = tgAdvertiser(ilAdfIndex).sKey    'Traffic!lbcAdvt.List(imAdvtIndex)
    End If

    ilRet = gParseItem(slNameCode, 1, "\", slName)
    tmAdf.sAbbr = Trim$(slName)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilAdfCode = Val(slCode)   '
    btrExtClear hmCrf   'Clear any previous extend operation
    ilExtLen = Len(tmCrf)  'Extract operation record size
    tmCrfSrchKey1.sRotType = "A"
    tmCrfSrchKey1.iEtfCode = 0
    tmCrfSrchKey1.iEnfCode = 0
    tmCrfSrchKey1.iAdfCode = ilAdfCode
    tmCrfSrchKey1.lChfCode = 0
    tmCrfSrchKey1.lFsfCode = 0
    tmCrfSrchKey1.iVefCode = 0
    tmCrfSrchKey1.iRotNo = 32000
    ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (tmCrf.iAdfCode = ilAdfCode) Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmCrf, llNoRec, -1, "UC", "CRF", "") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Crf", "CrfAdfCode")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, ilAdfCode, 2)
        On Error GoTo mGetShtTitleErr
        gBtrvErrorMsg ilRet, "mGetShtTitle (btrExtAddLogicConst):" & "Crf.Btr", Blackout
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "M", 1)
        On Error GoTo mGetShtTitleErr
        gBtrvErrorMsg ilRet, "mGetShtTitle (btrExtAddLogicConst):" & "Crf.Btr", Blackout
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "S", 1)
        On Error GoTo mGetShtTitleErr
        gBtrvErrorMsg ilRet, "mGetShtTitle (btrExtAddLogicConst):" & "Crf.Btr", Blackout
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "R", 1)
        On Error GoTo mGetShtTitleErr
        gBtrvErrorMsg ilRet, "mGetShtTitle (btrExtAddLogicConst):" & "Crf.Btr", Blackout
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, ByVal "X", 1)
        On Error GoTo mGetShtTitleErr
        gBtrvErrorMsg ilRet, "mGetShtTitle (btrExtAddLogicConst):" & "Crf.Btr", Blackout
        On Error GoTo 0
        ilOffSet = 0
        ilRet = btrExtAddField(hmCrf, ilOffSet, ilExtLen)  'Extract start/end time, and days
        On Error GoTo mGetShtTitleErr
        gBtrvErrorMsg ilRet, "mGetShtTitle (btrExtAddField):" & "Crf.Btr", Blackout
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmClf)    'Extract record
        ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mGetShtTitleErr
            gBtrvErrorMsg ilRet, "mGetShtTitle (btrExtGetNextExt):" & "Crf.Btr", Blackout
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
            If ilRet = BTRV_ERR_REJECT_COUNT Then
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
            End If
            Do While ilRet = BTRV_ERR_NONE
                'Test if Cif used
                ilFound = False
                gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
                If (gDateValue(slDate) > lmNowDate) Then
                    tmCnfSrchKey.lCrfCode = tmCrf.lCode
                    tmCnfSrchKey.iInstrNo = 0
                    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                        If tmCnf.lCifCode = llCifCode Then
                            ilFound = True
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If ilFound Then
                        If tmChf.lCode <> tmCrf.lChfCode Then
                            tmChfSrchKey.lCode = tmCrf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo mGetShtTitleErr
                            gBtrvErrorMsg ilRet, "mGetShtTitle (btrGetEqual):" & "Chf.Btr", Blackout
                            On Error GoTo 0
                        End If
                        mGetShtTitle = True
                        'Code Taken from gGetShortTitle
                        llSifCode = 0
                        If tgSpf.sUseProdSptScr = "P" Then
                            If tmChf.lVefCode < 0 Then
                                tmVsfSrchKey.lCode = -tmChf.lVefCode
                                ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                Do While ilRet = BTRV_ERR_NONE
                                    'ilCrfVefCode = gGetCrfVefCode(hlClf, tlSdf)
                                    For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                                        If tmVsf.iFSCode(ilVsf) = tmCrf.iVefCode Then
                                            If tmVsf.lFSComm(ilVsf) > 0 Then
                                                llSifCode = tmVsf.lFSComm(ilVsf)
                                            End If
                                            Exit For
                                        End If
                                    Next ilVsf
                                    If llSifCode <> 0 Then
                                        Exit Do
                                    End If
                                    If tmVsf.lLkVsfCode <= 0 Then
                                        Exit Do
                                    End If
                                    tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
                                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            End If
                            If llSifCode = 0 Then
                                llSifCode = tmChf.lSifCode
                            End If
                            slShtTitle = gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf)
                        Else
                            slShtTitle = Trim$(tmChf.sProduct)
                            'mGetShtTitle = gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf)
                            tmCifSrchKey.lCode = llCifCode
                            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                If tmCif.lcpfCode > 0 Then
                                    tmCpfSrchKey.lCode = tmCif.lcpfCode
                                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If (ilRet = BTRV_ERR_NONE) And (Trim$(tmCpf.sName) <> "") Then
                                        slShtTitle = Trim$(tmCpf.sName)
                                        llSifCode = tmCif.lcpfCode
                                    End If
                                End If
                            End If
                        End If
                        If llSifCode > 0 Then
                            btrExtClear hmCrf   'Clear any previous extend operation
                            Exit Function
                        End If
                    End If
                End If
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                End If
            Loop
        End If
        btrExtClear hmCrf   'Clear any previous extend operation
    End If
    Exit Function
mGetShtTitleErr:
    On Error GoTo 0
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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

    Screen.MousePointer = vbHourglass
    imLBRCtrls = 1
    imLBSCtrls = 1
    imLBCDCtrls = 1
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.height = 165
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    imFirstActivate = True
    imFirstTime = True
    imTerminate = False
    imPopReqd = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    If smFromLog = "Y" Then
        ReDim tmSCtrls(0 To 12) As FIELDAREA
        imSRIndex = 1
        SADVTINDEX = 1          'Blackout control/field
        SCNTRINDEX = 2
        SRADVTINDEX = 3          'Blackout control/field
        SRCNTRINDEX = 4
        SRCARTINDEX = 5
        SVEHINDEX = 6
        SLENINDEX = 7
        SSTARTDATEINDEX = 8
        SENDDATEINDEX = 9
        SDAYINDEX = 10
        SSTARTTIMEINDEX = 11
        SENDTIMEINDEX = 12
        SSHORTTITLEINDEX = -100
        ReDim tmRCtrls(0 To 11) As FIELDAREA
        RADVTINDEX = 1
        RCNTRINDEX = 2
        RCARTINDEX = 3
        RPPINDEX = 4
        RVEHINDEX = 6
        RSTARTDATEINDEX = 7
        RENDDATEINDEX = 8
        RDAYINDEX = 9
        RSTARTTIMEINDEX = 10
        RENDTIMEINDEX = 11
        RSHORTTITLEINDEX = -100
    Else
        ReDim tmSCtrls(0 To 14) As FIELDAREA
        imSRIndex = 0
        SADVTINDEX = 1          'Blackout control/field
        SSHORTTITLEINDEX = 2
        SVEHINDEX = 3
        SSTARTDATEINDEX = 4
        SENDDATEINDEX = 5
        SDAYINDEX = 6
        SSTARTTIMEINDEX = 13
        SENDTIMEINDEX = 14
        SCNTRINDEX = -100
        SRADVTINDEX = -101          'Blackout control/field
        SRCNTRINDEX = -102
        SRCARTINDEX = -103
        SLENINDEX = -104
        ReDim tmRCtrls(0 To 16) As FIELDAREA
        RADVTINDEX = 1
        RCARTINDEX = 2
        RSHORTTITLEINDEX = 3
        RPPINDEX = 4
        RSTARTDATEINDEX = 6
        RENDDATEINDEX = 7
        RDAYINDEX = 8
        RSTARTTIMEINDEX = 15
        RENDTIMEINDEX = 16
        RCNTRINDEX = -100
        RVEHINDEX = -101
    End If
    pbcSuppression(imSRIndex).Visible = True
    'Blackout.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    'gCenterStdAlone Blackout
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    imSBoxNo = -1 'Initialize current Box to N/A
    imSRowNo = -1 'Initialize current Box to N/A
    imRBoxNo = -1 'Initialize current Box to N/A
    imRRowNo = -1 'Initialize current Box to N/A
    imCalType = 0   'Standard
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imDoubleClickName = False
    imLbcMouseDown = False
    imBypassFocus = False
    imChgMode = False
    imBSMode = False
    imSChg = False
    imRChg = False
    igView = 0
    imButtonIndex = -1
    smNowDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(smNowDate)
    imIgnoreRightMove = False
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    imSettingValue = False
    imLbcArrowSetting = False
    hmBof = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmBof, "", sgDBPath & "Bof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Bof.Btr)", Blackout
    On Error GoTo 0
    'ReDim tgSBofRec(1 To 1) As BOFREC
    ReDim tgSBofRec(0 To 0) As BOFREC
    'ReDim tgRBofRec(1 To 1) As BOFREC
    ReDim tgRBofRec(0 To 0) As BOFREC
    'ReDim tgBofDel(1 To 1) As BOFREC
    ReDim tgBofDel(0 To 0) As BOFREC
    
'    ReDim smSShow(1 To 14, 1 To 1) As String 'Values shown in program area
'    ReDim smSSave(1 To 12, 1 To 1) As String    'Values saved (program name) in program area
'    ReDim imSSave(1 To 7, 1 To 1) As Integer 'Values saved (program name) in program area
'    ReDim lmSSave(1 To 4, 1 To 1) As Long 'Values saved (program name) in program area
'    ReDim smRShow(1 To 16, 1 To 1) As String 'Values shown in program area
'    ReDim smRSave(1 To 11, 1 To 1) As String    'Values saved (program name) in program area
'    ReDim imRSave(1 To 7, 1 To 1) As Integer 'Values saved (program name) in program area
'    ReDim lmRSave(1 To 3, 1 To 1) As Long 'Values saved (program name) in program area
    ReDim smSShow(0 To 14, 0 To 0) As String 'Values shown in program area
    ReDim smSSave(0 To 12, 0 To 0) As String    'Values saved (program name) in program area
    ReDim imSSave(0 To 7, 0 To 0) As Integer 'Values saved (program name) in program area
    ReDim lmSSave(0 To 4, 0 To 0) As Long 'Values saved (program name) in program area
    ReDim smRShow(0 To 16, 0 To 0) As String 'Values shown in program area
    ReDim smRSave(0 To 11, 0 To 0) As String    'Values saved (program name) in program area
    ReDim imRSave(0 To 7, 0 To 0) As Integer 'Values saved (program name) in program area
    ReDim lmRSave(0 To 3, 0 To 0) As Long 'Values saved (program name) in program area


    imBofRecLen = Len(tmBof)
    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "PRF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: PRF.Btr)", Blackout
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)  'Get and save CSF record length
    hmSif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sif.Btr)", Blackout
    On Error GoTo 0
    imSifRecLen = Len(tmSif)
    hmCif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", Blackout
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmCpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", Blackout
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmCrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Crf.Btr)", Blackout
    On Error GoTo 0
    imCrfRecLen = Len(tmCrf)
    hmCnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cnf.Btr)", Blackout
    On Error GoTo 0
    imCnfRecLen = Len(tmCnf)
    hmMcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mcf.Btr)", Blackout
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", Blackout
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", Blackout
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    lbcRAdvt.Clear 'Force population
    lbcSAdvt.Clear 'Force population
    sgAdvertiserTag = ""
    mAdvtPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lbcVehicle.Clear 'Force population
    sgUserVehicleTag = ""
    mVehPop  'Create tmUserVeh
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    mCompPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    mDaysPop
    Screen.MousePointer = vbHourglass
    mInitBox
    gCenterStdAlone Blackout
    Screen.MousePointer = vbHourglass
    If smFromLog = "Y" Then
        If smSplitFill = "Y" Then
            ilRet = gReadBofRec(1, hmBof, hmCif, hmPrf, hmSif, hmCHF, "R", smNowDate, 0)
        Else
            ilRet = gReadBofRec(1, hmBof, hmCif, hmPrf, hmSif, hmCHF, "B", smNowDate, 0)
        End If
    Else
        ilRet = gReadBofRec(0, hmBof, hmCif, hmPrf, hmSif, hmCHF, "B", smNowDate, 0)
    End If
    If smSplitFill = "Y" Then
        rbcSR(1).Value = True
        plcSelect.Visible = False
    End If
    tmcStart.Enabled = True
    'If Not imTerminate Then
    '    mMoveRecToCtrl "B"
    '    'Add clear since auto select removed
    '    'mClearCtrlFields
    '    'Remove auto select until getting data is faster
    '    'If cbcSelect.ListCount <= 1 Then
    '    '    cbcSelect.ListIndex = 0 'This will generate a select_change event
    '    'Else
    '    '    cbcSelect.ListIndex = 1
    '    'End If
    '    'mSetCommands
    'End If
    DoEvents
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
    Dim llRMax As Long
    Dim llSMax As Long
    Dim llMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long

    flTextHeight = pbcSuppression(imSRIndex).TextHeight("1") - 35
    plcSelect.Move 5910, 60
    'Position panel and picture areas with panel
    'plcBlackout.Move 180, 525, pbcSuppression(imSRIndex).Width + vbcSR.Width + fgPanelAdj, pbcSuppression(imSRIndex).Height + fgPanelAdj
    'pbcSuppression(imSRIndex).Move plcBlackout.Left + fgBevelX, plcBlackout.Top + fgBevelY
    ''pbcReplacement(imSRIndex).Move pbcSuppression(imSRIndex).Left, pbcSuppression(imSRIndex).Top
    ''vbcSR.Move pbcSuppression(imSRIndex).Width + 60, 75, vbcSR.Width, pbcSuppression(imSRIndex).Height - 15
    'vbcSR.Move pbcSuppression(imSRIndex).Width, fgBevelY - 15, vbcSR.Width, pbcSuppression(imSRIndex).Height - 15
    pbcArrow.Move plcBlackout.Left - pbcArrow.Width - 15
    'Suppression
    If smFromLog = "Y" Then
        'Advertiser Name
        gSetCtrl tmSCtrls(SADVTINDEX), 30, 375, 990, fgBoxGridH
        'Contract
        gSetCtrl tmSCtrls(SCNTRINDEX), 1035, tmSCtrls(SADVTINDEX).fBoxY, 600, fgBoxGridH
        'Replace Advertiser Name
        gSetCtrl tmSCtrls(SRADVTINDEX), 1650, tmSCtrls(SADVTINDEX).fBoxY, 990, fgBoxGridH
        'Replace Contract
        gSetCtrl tmSCtrls(SRCNTRINDEX), 2655, tmSCtrls(SADVTINDEX).fBoxY, 600, fgBoxGridH
        'Replace Copy
        gSetCtrl tmSCtrls(SRCARTINDEX), 3270, tmSCtrls(SADVTINDEX).fBoxY, 495, fgBoxGridH
        'Vehicle
        gSetCtrl tmSCtrls(SVEHINDEX), 3780, tmSCtrls(SADVTINDEX).fBoxY, 1155, fgBoxGridH
        'Length
        gSetCtrl tmSCtrls(SLENINDEX), 4950, tmSCtrls(SADVTINDEX).fBoxY, 330, fgBoxGridH
        'Start Date
        gSetCtrl tmSCtrls(SSTARTDATEINDEX), 5295, tmSCtrls(SADVTINDEX).fBoxY, 690, fgBoxGridH
        'End Date
        gSetCtrl tmSCtrls(SENDDATEINDEX), 6000, tmSCtrls(SADVTINDEX).fBoxY, 690, fgBoxGridH
        'Days
        gSetCtrl tmSCtrls(SDAYINDEX), 6705, tmSCtrls(SADVTINDEX).fBoxY, 600, fgBoxGridH
        'Start Time
        gSetCtrl tmSCtrls(SSTARTTIMEINDEX), 7320, tmSCtrls(SADVTINDEX).fBoxY, 690, fgBoxGridH
        'End Time
        gSetCtrl tmSCtrls(SENDTIMEINDEX), 8025, tmSCtrls(SADVTINDEX).fBoxY, 690, fgBoxGridH
    Else
        'Advertiser Name
        gSetCtrl tmSCtrls(SADVTINDEX), 30, 225, 1380, fgBoxGridH
        'Short Title
        gSetCtrl tmSCtrls(SSHORTTITLEINDEX), 1425, tmSCtrls(SADVTINDEX).fBoxY, 1380, fgBoxGridH
        'Vehicle
        gSetCtrl tmSCtrls(SVEHINDEX), 2820, tmSCtrls(SADVTINDEX).fBoxY, 1380, fgBoxGridH
        'Start Date
        gSetCtrl tmSCtrls(SSTARTDATEINDEX), 4215, tmSCtrls(SADVTINDEX).fBoxY, 720, fgBoxGridH
        'End Date
        gSetCtrl tmSCtrls(SENDDATEINDEX), 4950, tmSCtrls(SADVTINDEX).fBoxY, 720, fgBoxGridH
        'Days of the week
        For ilLoop = 0 To 6 Step 1
            gSetCtrl tmSCtrls(SDAYINDEX + ilLoop), 5685 + 225 * (ilLoop), tmSCtrls(SADVTINDEX).fBoxY, 210, fgBoxGridH
        Next ilLoop
        'Start Time
        gSetCtrl tmSCtrls(SSTARTTIMEINDEX), 7260, tmSCtrls(SADVTINDEX).fBoxY, 720, fgBoxGridH
        'End Time
        gSetCtrl tmSCtrls(SENDTIMEINDEX), 7995, tmSCtrls(SADVTINDEX).fBoxY, 720, fgBoxGridH
    End If
    'Replacement
    'Advertiser Name
    If smFromLog = "Y" Then
        gSetCtrl tmRCtrls(RADVTINDEX), 30, 375, 1185, fgBoxGridH
        'Contract
        gSetCtrl tmRCtrls(RCNTRINDEX), 1230, tmRCtrls(RADVTINDEX).fBoxY, 825, fgBoxGridH
        'Cart
        gSetCtrl tmRCtrls(RCARTINDEX), 2070, tmRCtrls(RADVTINDEX).fBoxY, 510, fgBoxGridH
        'Product Protection 1 (Competitive)
        gSetCtrl tmRCtrls(RPPINDEX), 2595, tmRCtrls(RADVTINDEX).fBoxY, 615, fgBoxGridH
        'Product Protection 2 (Competitive)
        gSetCtrl tmRCtrls(RPPINDEX + 1), 3225, tmRCtrls(RADVTINDEX).fBoxY, 615, fgBoxGridH
        'Vehicle
        gSetCtrl tmRCtrls(RVEHINDEX), 3855, tmRCtrls(RADVTINDEX).fBoxY, 1185, fgBoxGridH
        'Start Date
        gSetCtrl tmRCtrls(RSTARTDATEINDEX), 5055, tmRCtrls(RADVTINDEX).fBoxY, 720, fgBoxGridH
        'End Date
        gSetCtrl tmRCtrls(RENDDATEINDEX), 5790, tmRCtrls(RADVTINDEX).fBoxY, 720, fgBoxGridH
        'Days
        gSetCtrl tmRCtrls(RDAYINDEX), 6525, tmSCtrls(RADVTINDEX).fBoxY, 720, fgBoxGridH
        'Start Time
        gSetCtrl tmRCtrls(RSTARTTIMEINDEX), 7260, tmSCtrls(RADVTINDEX).fBoxY, 720, fgBoxGridH
        'End Time
        gSetCtrl tmRCtrls(RENDTIMEINDEX), 7995, tmSCtrls(RADVTINDEX).fBoxY, 720, fgBoxGridH
    Else
        gSetCtrl tmRCtrls(RADVTINDEX), 30, 225, 1185, fgBoxGridH
        gSetCtrl tmRCtrls(RCARTINDEX), 1230, tmRCtrls(RADVTINDEX).fBoxY, 510, fgBoxGridH
        'Short Title or Product
        gSetCtrl tmRCtrls(RSHORTTITLEINDEX), 1755, tmRCtrls(RADVTINDEX).fBoxY, 1185, fgBoxGridH
        'Product Protection 1 (Competitive)
        gSetCtrl tmRCtrls(RPPINDEX), 2955, tmRCtrls(RADVTINDEX).fBoxY, 615, fgBoxGridH
        'Product Protection 2 (Competitive)
        gSetCtrl tmRCtrls(RPPINDEX + 1), 3585, tmRCtrls(RADVTINDEX).fBoxY, 615, fgBoxGridH
        'Start Date
        gSetCtrl tmRCtrls(RSTARTDATEINDEX), 4215, tmRCtrls(RADVTINDEX).fBoxY, 720, fgBoxGridH
        'End Date
        gSetCtrl tmRCtrls(RENDDATEINDEX), 4950, tmRCtrls(RADVTINDEX).fBoxY, 720, fgBoxGridH
        'Days of the week
        For ilLoop = 0 To 6 Step 1
            gSetCtrl tmRCtrls(RDAYINDEX + ilLoop), 5685 + 225 * (ilLoop), tmRCtrls(RADVTINDEX).fBoxY, 210, fgBoxGridH
        Next ilLoop
        'Start Time
        gSetCtrl tmRCtrls(RSTARTTIMEINDEX), 7260, tmSCtrls(RADVTINDEX).fBoxY, 720, fgBoxGridH
        'End Time
        gSetCtrl tmRCtrls(RENDTIMEINDEX), 7995, tmSCtrls(RADVTINDEX).fBoxY, 720, fgBoxGridH
    End If

    llSMax = 0
    For ilLoop = imLBSCtrls To UBound(tmSCtrls) Step 1
        tmSCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmSCtrls(ilLoop).fBoxW)
        Do While (tmSCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmSCtrls(ilLoop).fBoxW = tmSCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmSCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmSCtrls(ilLoop).fBoxX)
            Do While (tmSCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmSCtrls(ilLoop).fBoxX = tmSCtrls(ilLoop).fBoxX + 1
            Loop
            If tmSCtrls(ilLoop).fBoxX > 90 Then
                Do
                    If tmSCtrls(ilLoop - 1).fBoxX + tmSCtrls(ilLoop - 1).fBoxW + 15 < tmSCtrls(ilLoop).fBoxX Then
                        tmSCtrls(ilLoop - 1).fBoxW = tmSCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmSCtrls(ilLoop - 1).fBoxX + tmSCtrls(ilLoop - 1).fBoxW + 15 > tmSCtrls(ilLoop).fBoxX Then
                        tmSCtrls(ilLoop - 1).fBoxW = tmSCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmSCtrls(ilLoop).fBoxX + tmSCtrls(ilLoop).fBoxW + 15 > llSMax Then
            llSMax = tmSCtrls(ilLoop).fBoxX + tmSCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    llRMax = 0
    For ilLoop = imLBRCtrls To UBound(tmRCtrls) Step 1
        tmRCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmRCtrls(ilLoop).fBoxW)
        Do While (tmRCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmRCtrls(ilLoop).fBoxW = tmRCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmRCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmRCtrls(ilLoop).fBoxX)
            Do While (tmRCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmRCtrls(ilLoop).fBoxX = tmRCtrls(ilLoop).fBoxX + 1
            Loop
            If tmRCtrls(ilLoop).fBoxX > 90 Then
                Do
                    If tmRCtrls(ilLoop - 1).fBoxX + tmRCtrls(ilLoop - 1).fBoxW + 15 < tmRCtrls(ilLoop).fBoxX Then
                        tmRCtrls(ilLoop - 1).fBoxW = tmRCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmRCtrls(ilLoop - 1).fBoxX + tmRCtrls(ilLoop - 1).fBoxW + 15 > tmRCtrls(ilLoop).fBoxX Then
                        tmRCtrls(ilLoop - 1).fBoxW = tmRCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmRCtrls(ilLoop).fBoxX + tmRCtrls(ilLoop).fBoxW + 15 > llRMax Then
            llRMax = tmRCtrls(ilLoop).fBoxX + tmRCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    If llSMax < llRMax Then
        tmRCtrls(UBound(tmRCtrls)).fBoxW = tmRCtrls(UBound(tmRCtrls)).fBoxW - (llRMax - llSMax)
        llMax = llSMax
    ElseIf llSMax > llRMax Then
        tmSCtrls(UBound(tmSCtrls)).fBoxW = tmSCtrls(UBound(tmSCtrls)).fBoxW - (llSMax - llRMax)
        llMax = llRMax
    Else
        llMax = llSMax
    End If
    pbcSuppression(imSRIndex).Picture = LoadPicture("")
    pbcSuppression(imSRIndex).Width = llMax
    plcBlackout.Width = llMax + vbcSR.Width + 2 * fgBevelX + 15
    lacSFrame(imSRIndex).Width = llMax - 15
    pbcReplacement(imSRIndex).Picture = LoadPicture("")
    pbcReplacement(imSRIndex).Width = pbcSuppression(imSRIndex).Width
    lacRFrame(imSRIndex).Width = lacSFrame(imSRIndex).Width
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    cmcDone.Left = (Blackout.Width - 4 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
    cmcUpdate.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    cmcUndo.Left = cmcUpdate.Left + cmcUpdate.Width + ilSpaceBetweenButtons
    cmcDone.Top = Blackout.height - (3 * cmcDone.height) / 2
    cmcCancel.Top = cmcDone.Top
    cmcUpdate.Top = cmcDone.Top
    cmcUndo.Top = cmcDone.Top
    imcTrash.Top = cmcDone.Top - imcTrash.height / 2
    imcTrash.Left = Blackout.Width - (3 * imcTrash.Width) / 2
    llAdjTop = imcTrash.Top - plcBlackout.Top - 120 - tmRCtrls(1).fBoxY
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    If smFromLog <> "Y" Then
        llAdjTop = llAdjTop + 30
    End If
    Do While plcBlackout.Top + llAdjTop + 2 * fgBevelY + 240 < imcTrash.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    plcBlackout.height = llAdjTop + 2 * fgBevelY
    pbcSuppression(imSRIndex).Left = plcBlackout.Left + fgBevelX
    pbcSuppression(imSRIndex).Top = plcBlackout.Top + fgBevelY
    pbcSuppression(imSRIndex).height = plcBlackout.height - 2 * fgBevelY
    vbcSR.Left = plcBlackout.Width - vbcSR.Width - fgBevelX - 15
    vbcSR.Top = fgBevelY - 30
    vbcSR.height = pbcSuppression(imSRIndex).height
    pbcReplacement(imSRIndex).Left = pbcSuppression(imSRIndex).Left
    pbcReplacement(imSRIndex).Top = pbcSuppression(imSRIndex).Top
    pbcReplacement(imSRIndex).height = pbcSuppression(imSRIndex).height
    plcSelect.Left = plcBlackout.Left + plcBlackout.Width - plcSelect.Width
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

    If igView = 0 Then
        For ilLoop = LBound(smSSave, 1) To UBound(smSSave, 1) Step 1
            smSSave(ilLoop, ilRowNo) = ""
        Next ilLoop
        For ilLoop = LBound(imSSave, 1) To UBound(imSSave, 1) Step 1
            imSSave(ilLoop, ilRowNo) = -1
        Next ilLoop
        For ilLoop = LBound(lmSSave, 1) To UBound(lmSSave, 1) Step 1
            lmSSave(ilLoop, ilRowNo) = -1
        Next ilLoop
        For ilLoop = LBound(smSShow, 1) To UBound(smSShow, 1) Step 1
            smSShow(ilLoop, ilRowNo) = ""
        Next ilLoop
        tgSBofRec(ilRowNo).iStatus = 0
        tgSBofRec(ilRowNo).lRecPos = 0
    Else
        For ilLoop = LBound(smRSave, 1) To UBound(smRSave, 1) Step 1
            smRSave(ilLoop, ilRowNo) = ""
        Next ilLoop
        For ilLoop = LBound(imRSave, 1) To UBound(imRSave, 1) Step 1
            imRSave(ilLoop, ilRowNo) = -1
        Next ilLoop
        For ilLoop = LBound(lmRSave, 1) To UBound(lmRSave, 1) Step 1
            lmRSave(ilLoop, ilRowNo) = -1
        Next ilLoop
        For ilLoop = LBound(smRShow, 1) To UBound(smRShow, 1) Step 1
            smRShow(ilLoop, ilRowNo) = ""
        Next ilLoop
        tgRBofRec(ilRowNo).iStatus = 0
        tgRBofRec(ilRowNo).lRecPos = 0
    End If
End Sub
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
Private Sub mLenPop(ilRowNo As Integer)
    'Dim ilRet As Integer
    Dim ilVpfIndex As Integer
    Dim ilVefCode As Integer
    Dim ilLoop As Integer
    Dim ilLen As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer

    ilVpfIndex = -1
    If smSSave(3, ilRowNo) <> "" Then
        If (smFromLog = "Y") Then
            gFindMatch Trim$(smSSave(3, ilRowNo)), 1, lbcVehicle
            If gLastFound(lbcVehicle) > 0 Then
                slNameCode = tgUserVehicle(gLastFound(lbcVehicle)).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilVpfIndex = gVpfFind(Blackout, ilVefCode)
            End If
        Else
            gFindMatch Trim$(smSSave(3, ilRowNo)), 0, lbcVehicle
            If gLastFound(lbcVehicle) >= 0 Then
                slNameCode = tgUserVehicle(gLastFound(lbcVehicle)).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilVpfIndex = gVpfFind(Blackout, ilVefCode)
            End If
        End If
    End If
    lbcLen.Clear
    If ilVpfIndex >= 0 Then
        For ilLoop = LBound(tgVpf(ilVpfIndex).iSLen) To UBound(tgVpf(ilVpfIndex).iSLen) Step 1
            If tgVpf(ilVpfIndex).iSLen(ilLoop) <> 0 Then
                lbcLen.AddItem Trim$(str$(tgVpf(ilVpfIndex).iSLen(ilLoop)))
            End If
        Next ilLoop
    Else
        For ilLoop = LBound(tgVpf) To UBound(tgVpf) Step 1
            For ilLen = LBound(tgVpf(ilLoop).iSLen) To UBound(tgVpf(ilLoop).iSLen) Step 1
                If tgVpf(ilLoop).iSLen(ilLen) <> 0 Then
                    gFindMatch Trim$(str$(tgVpf(ilLoop).iSLen(ilLen))), 0, lbcLen
                    If gLastFound(lbcLen) < 0 Then
                        lbcLen.AddItem Trim$(str$(tgVpf(ilLoop).iSLen(ilLen)))
                    End If
                End If
            Next ilLen
        Next ilLoop
    End If
    lbcLen.AddItem "[All]", 0
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
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim ilDay As Integer
    Dim ilFDay As Integer
    Dim ilTDay As Integer
    Dim ilView As Integer
    Dim ilPos As Integer
    ilView = igView
    igView = 0
    For ilRowNo = LBound(tgSBofRec) To UBound(tgSBofRec) - 1 Step 1
        'Set Advertiser
        If (smFromLog = "Y") Then
            gFindMatch Trim$(smSSave(1, ilRowNo)), 0, lbcSAdvt
            If gLastFound(lbcSAdvt) > 0 Then
                slNameCode = tgAdvertiser(gLastFound(lbcSAdvt) - 1).sKey  'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tgSBofRec(ilRowNo).tBof.iAdfCode = Val(slCode)
                tgSBofRec(ilRowNo).tBof.lSChfCode = lmSSave(1, ilRowNo)
                tgSBofRec(ilRowNo).sAdfName = smSSave(1, ilRowNo)
                tgSBofRec(ilRowNo).lSCntrNo = Val(smSSave(11, ilRowNo))
            Else
                tgSBofRec(ilRowNo).tBof.iAdfCode = 0
                tgSBofRec(ilRowNo).tBof.lSChfCode = 0
                tgSBofRec(ilRowNo).sAdfName = ""
                tgSBofRec(ilRowNo).lSCntrNo = 0
            End If
            gFindMatch Trim$(smSSave(8, ilRowNo)), 0, lbcSAdvt
            If gLastFound(lbcSAdvt) > 0 Then
                slNameCode = tgAdvertiser(gLastFound(lbcSAdvt) - 1).sKey  'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tgSBofRec(ilRowNo).tBof.iRAdfCode = Val(slCode)
                tgSBofRec(ilRowNo).tBof.lRChfCode = lmSSave(2, ilRowNo)
                tgSBofRec(ilRowNo).sRAdfName = smSSave(8, ilRowNo)
                tgSBofRec(ilRowNo).lRCntrNo = Val(smSSave(12, ilRowNo))
                tgSBofRec(ilRowNo).tBof.lCifCode = lmSSave(3, ilRowNo)
                tgSBofRec(ilRowNo).tBof.lSifCode = lmSSave(4, ilRowNo)
            Else
                tgSBofRec(ilRowNo).tBof.iRAdfCode = 0
                tgSBofRec(ilRowNo).tBof.lRChfCode = 0
                tgSBofRec(ilRowNo).tBof.lCifCode = 0
                tgSBofRec(ilRowNo).tBof.lSifCode = 0
                tgSBofRec(ilRowNo).sRAdfName = ""
                tgSBofRec(ilRowNo).lRCntrNo = 0
            End If
            'Set Vehicle
            gFindMatch Trim$(smSSave(3, ilRowNo)), 1, lbcVehicle
            If gLastFound(lbcVehicle) > 0 Then
                slNameCode = tgUserVehicle(gLastFound(lbcVehicle) - 1).sKey  'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tgSBofRec(ilRowNo).tBof.iVefCode = Val(slCode)
                tgSBofRec(ilRowNo).sVefName = smSSave(3, ilRowNo)
            Else
                tgSBofRec(ilRowNo).tBof.iVefCode = 0
                tgSBofRec(ilRowNo).sVefName = ""
            End If
            'Length
            If (smSSave(9, ilRowNo) <> "") And (smSSave(9, ilRowNo) <> "[All]") Then
                tgSBofRec(ilRowNo).tBof.iLen = Val(smSSave(9, ilRowNo))
            Else
                tgSBofRec(ilRowNo).tBof.iLen = 0
            End If
            'Dates
            gPackDate smSSave(4, ilRowNo), tgSBofRec(ilRowNo).tBof.iStartDate(0), tgSBofRec(ilRowNo).tBof.iStartDate(1)
            gPackDate smSSave(5, ilRowNo), tgSBofRec(ilRowNo).tBof.iEndDate(0), tgSBofRec(ilRowNo).tBof.iEndDate(1)
            ilPos = InStr(1, smSSave(10, ilRowNo), "-", 1)
            slStr = UCase$(Left$(smSSave(10, ilRowNo), 2))
            Select Case slStr
                Case "MO"
                    ilFDay = 0
                Case "TU"
                    ilFDay = 1
                Case "WE"
                    ilFDay = 2
                Case "TH"
                    ilFDay = 3
                Case "FR"
                    ilFDay = 4
                Case "SA"
                    ilFDay = 5
                Case "SU"
                    ilFDay = 6
            End Select
            If ilPos > 0 Then
                slStr = UCase$(Mid$(smSSave(10, ilRowNo), ilPos + 1))
                Select Case slStr
                    Case "MO"
                        ilTDay = 0
                    Case "TU"
                        ilTDay = 1
                    Case "WE"
                        ilTDay = 2
                    Case "TH"
                        ilTDay = 3
                    Case "FR"
                        ilTDay = 4
                    Case "SA"
                        ilTDay = 5
                    Case "SU"
                        ilTDay = 6
                End Select
            Else
                ilTDay = ilFDay
            End If
            For ilDay = 0 To 6 Step 1
                If (ilDay >= ilFDay) And (ilDay <= ilTDay) Then
                    tgSBofRec(ilRowNo).tBof.sDays(ilDay) = "Y"
                Else
                    tgSBofRec(ilRowNo).tBof.sDays(ilDay) = "N"
                End If
            Next ilDay
            gPackTime smSSave(6, ilRowNo), tgSBofRec(ilRowNo).tBof.iStartTime(0), tgSBofRec(ilRowNo).tBof.iStartTime(1)
            gPackTime smSSave(7, ilRowNo), tgSBofRec(ilRowNo).tBof.iEndTime(0), tgSBofRec(ilRowNo).tBof.iEndTime(1)
'            tgSBofRec(ilRowNo).tBof.lSifCode = 0
'            tgSBofRec(ilRowNo).tBof.lCifCode = 0
        Else
            gFindMatch Trim$(smSSave(1, ilRowNo)), 0, lbcSAdvt
            If gLastFound(lbcSAdvt) >= 0 Then
                slNameCode = tgAdvertiser(gLastFound(lbcSAdvt)).sKey  'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tgSBofRec(ilRowNo).tBof.iAdfCode = Val(slCode)
                tgSBofRec(ilRowNo).sAdfName = smSSave(1, ilRowNo)
            End If
            'Set Short Title
            If (smSSave(2, ilRowNo) <> "[None]") And (Trim$(smSSave(2, ilRowNo)) <> "") Then
                mShtTitlePop ilRowNo
                gFindMatch Trim$(smSSave(2, ilRowNo)), 1, lbcShtTitle
                If gLastFound(lbcShtTitle) > 0 Then
                    slNameCode = tmShtTitleCode(gLastFound(lbcShtTitle) - 1).sKey  'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tgSBofRec(ilRowNo).tBof.lSifCode = Val(slCode)
                    tgSBofRec(ilRowNo).sShtTitle = smSSave(2, ilRowNo)
                End If
            Else
                tgSBofRec(ilRowNo).tBof.lSifCode = 0
                tgSBofRec(ilRowNo).sShtTitle = ""
            End If
            'Set Vehicle
            gFindMatch Trim$(smSSave(3, ilRowNo)), 0, lbcVehicle
            If gLastFound(lbcVehicle) >= 0 Then
                slNameCode = tgUserVehicle(gLastFound(lbcVehicle)).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tgSBofRec(ilRowNo).tBof.iVefCode = Val(slCode)
                tgSBofRec(ilRowNo).sVefName = smSSave(3, ilRowNo)
            End If
            'Dates
            gPackDate smSSave(4, ilRowNo), tgSBofRec(ilRowNo).tBof.iStartDate(0), tgSBofRec(ilRowNo).tBof.iStartDate(1)
            gPackDate smSSave(5, ilRowNo), tgSBofRec(ilRowNo).tBof.iEndDate(0), tgSBofRec(ilRowNo).tBof.iEndDate(1)
            For ilDay = 0 To 6 Step 1
                If imSSave(ilDay + 1, ilRowNo) = 0 Then
                    tgSBofRec(ilRowNo).tBof.sDays(ilDay) = "N"
                Else
                    tgSBofRec(ilRowNo).tBof.sDays(ilDay) = "Y"
                End If
            Next ilDay
            gPackTime smSSave(6, ilRowNo), tgSBofRec(ilRowNo).tBof.iStartTime(0), tgSBofRec(ilRowNo).tBof.iStartTime(1)
            gPackTime smSSave(7, ilRowNo), tgSBofRec(ilRowNo).tBof.iEndTime(0), tgSBofRec(ilRowNo).tBof.iEndTime(1)
            tgSBofRec(ilRowNo).tBof.lCifCode = 0
            tgSBofRec(ilRowNo).tBof.lSChfCode = 0
            tgSBofRec(ilRowNo).tBof.iRAdfCode = 0
            tgSBofRec(ilRowNo).tBof.lRChfCode = 0
            tgSBofRec(ilRowNo).tBof.iLen = 0
        End If
    Next ilRowNo
    igView = 1
    For ilRowNo = LBound(tgRBofRec) To UBound(tgRBofRec) - 1 Step 1
        If (smFromLog = "Y") Then
            'Set Advertiser
            gFindMatch Trim$(smRSave(1, ilRowNo)), 0, lbcRAdvt
            If gLastFound(lbcRAdvt) >= 0 Then
                slNameCode = tgAdvertiser(gLastFound(lbcRAdvt)).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tgRBofRec(ilRowNo).tBof.iAdfCode = Val(slCode)
                tgRBofRec(ilRowNo).tBof.lRChfCode = lmRSave(3, ilRowNo)
                tgRBofRec(ilRowNo).sAdfName = smRSave(1, ilRowNo)
                tgRBofRec(ilRowNo).lRCntrNo = Val(smRSave(11, ilRowNo))
            End If
            tgRBofRec(ilRowNo).tBof.iRAdfCode = 0
            tgRBofRec(ilRowNo).tBof.lSChfCode = 0
            tgRBofRec(ilRowNo).lSCntrNo = 0
            'Set Cart
            'mCartPop ilRowNo
            Screen.MousePointer = vbHourglass
            tgRBofRec(ilRowNo).tBof.lCifCode = lmRSave(1, ilRowNo)
            tgRBofRec(ilRowNo).tBof.lSifCode = lmRSave(2, ilRowNo)
            'Set Vehicle
            If Trim$(smRSave(3, ilRowNo)) <> "[All]" Then
                gFindMatch Trim$(smRSave(3, ilRowNo)), 1, lbcVehicle
                If gLastFound(lbcVehicle) > 0 Then
                    slNameCode = tgUserVehicle(gLastFound(lbcVehicle) - 1).sKey  'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tgRBofRec(ilRowNo).tBof.iVefCode = Val(slCode)
                    tgRBofRec(ilRowNo).sVefName = smRSave(3, ilRowNo)
                Else
                    tgRBofRec(ilRowNo).tBof.iVefCode = 0
                    tgRBofRec(ilRowNo).sVefName = ""
                End If
            Else
                tgRBofRec(ilRowNo).tBof.iVefCode = 0
                tgRBofRec(ilRowNo).sVefName = ""
            End If
            'Product Protection
            tgRBofRec(ilRowNo).tBof.iMnfComp(0) = 0
            If smSplitFill <> "Y" Then
                gFindMatch Trim$(smRSave(4, ilRowNo)), 0, lbcComp(0)
                If gLastFound(lbcComp(0)) > 0 Then
                    slNameCode = tgCompCode(gLastFound(lbcComp(0)) - 1).sKey  'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tgRBofRec(ilRowNo).tBof.iMnfComp(0) = Val(slCode)
                End If
            End If
            tgRBofRec(ilRowNo).tBof.iMnfComp(1) = 0
            If smSplitFill <> "Y" Then
                gFindMatch Trim$(smRSave(5, ilRowNo)), 0, lbcComp(1)
                If gLastFound(lbcComp(1)) > 0 Then
                    slNameCode = tgCompCode(gLastFound(lbcComp(1)) - 1).sKey  'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tgRBofRec(ilRowNo).tBof.iMnfComp(1) = Val(slCode)
                End If
            End If
            'Dates
            gPackDate smRSave(6, ilRowNo), tgRBofRec(ilRowNo).tBof.iStartDate(0), tgRBofRec(ilRowNo).tBof.iStartDate(1)
            gPackDate smRSave(7, ilRowNo), tgRBofRec(ilRowNo).tBof.iEndDate(0), tgRBofRec(ilRowNo).tBof.iEndDate(1)
            'For ilDay = 0 To 6 Step 1
            '    If imRSave(ilDay + 1, ilRowNo) = 0 Then
            '        tgRBofRec(ilRowNo).tBof.sDays(ilDay) = "N"
            '    Else
            '        tgRBofRec(ilRowNo).tBof.sDays(ilDay) = "Y"
            '    End If
            'Next ilDay
            ilPos = InStr(1, smRSave(10, ilRowNo), "-", 1)
            slStr = UCase$(Left$(smRSave(10, ilRowNo), 2))
            Select Case slStr
                Case "MO"
                    ilFDay = 0
                Case "TU"
                    ilFDay = 1
                Case "WE"
                    ilFDay = 2
                Case "TH"
                    ilFDay = 3
                Case "FR"
                    ilFDay = 4
                Case "SA"
                    ilFDay = 5
                Case "SU"
                    ilFDay = 6
            End Select
            If ilPos > 0 Then
                slStr = UCase$(Mid$(smRSave(10, ilRowNo), ilPos + 1))
                Select Case slStr
                    Case "MO"
                        ilTDay = 0
                    Case "TU"
                        ilTDay = 1
                    Case "WE"
                        ilTDay = 2
                    Case "TH"
                        ilTDay = 3
                    Case "FR"
                        ilTDay = 4
                    Case "SA"
                        ilTDay = 5
                    Case "SU"
                        ilTDay = 6
                End Select
            Else
                ilTDay = ilFDay
            End If
            For ilDay = 0 To 6 Step 1
                If (ilDay >= ilFDay) And (ilDay <= ilTDay) Then
                    tgRBofRec(ilRowNo).tBof.sDays(ilDay) = "Y"
                Else
                    tgRBofRec(ilRowNo).tBof.sDays(ilDay) = "N"
                End If
            Next ilDay
            gPackTime smRSave(8, ilRowNo), tgRBofRec(ilRowNo).tBof.iStartTime(0), tgRBofRec(ilRowNo).tBof.iStartTime(1)
            gPackTime smRSave(9, ilRowNo), tgRBofRec(ilRowNo).tBof.iEndTime(0), tgRBofRec(ilRowNo).tBof.iEndTime(1)
            tgRBofRec(ilRowNo).tBof.iLen = 0
        Else
            'Set Advertiser
            gFindMatch Trim$(smRSave(1, ilRowNo)), 0, lbcRAdvt
            If gLastFound(lbcRAdvt) >= 0 Then
                slNameCode = tgAdvertiser(gLastFound(lbcRAdvt)).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tgRBofRec(ilRowNo).tBof.iAdfCode = Val(slCode)
                tgRBofRec(ilRowNo).sAdfName = smRSave(1, ilRowNo)
            End If
            tgRBofRec(ilRowNo).tBof.lSChfCode = 0
            tgRBofRec(ilRowNo).lSCntrNo = 0
            tgRBofRec(ilRowNo).tBof.iRAdfCode = 0
            tgRBofRec(ilRowNo).tBof.lRChfCode = 0
            tgRBofRec(ilRowNo).lRCntrNo = 0
            'Set Cart
            'mCartPop ilRowNo
            Screen.MousePointer = vbHourglass
            tgRBofRec(ilRowNo).tBof.lCifCode = lmRSave(1, ilRowNo)
            tgRBofRec(ilRowNo).tBof.lSifCode = lmRSave(2, ilRowNo)
            ''Set Short Title
            'Set Vehicle
            tgRBofRec(ilRowNo).tBof.iVefCode = 0
            tgRBofRec(ilRowNo).sVefName = ""
            'Product Protection
            tgRBofRec(ilRowNo).tBof.iMnfComp(0) = 0
            gFindMatch Trim$(smRSave(4, ilRowNo)), 0, lbcComp(0)
            If gLastFound(lbcComp(0)) > 0 Then
                slNameCode = tgCompCode(gLastFound(lbcComp(0)) - 1).sKey  'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tgRBofRec(ilRowNo).tBof.iMnfComp(0) = Val(slCode)
            End If
            tgRBofRec(ilRowNo).tBof.iMnfComp(1) = 0
            gFindMatch Trim$(smRSave(5, ilRowNo)), 0, lbcComp(1)
            If gLastFound(lbcComp(1)) > 0 Then
                slNameCode = tgCompCode(gLastFound(lbcComp(1)) - 1).sKey  'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tgRBofRec(ilRowNo).tBof.iMnfComp(1) = Val(slCode)
            End If
            'Dates
            gPackDate smRSave(6, ilRowNo), tgRBofRec(ilRowNo).tBof.iStartDate(0), tgRBofRec(ilRowNo).tBof.iStartDate(1)
            gPackDate smRSave(7, ilRowNo), tgRBofRec(ilRowNo).tBof.iEndDate(0), tgRBofRec(ilRowNo).tBof.iEndDate(1)
            For ilDay = 0 To 6 Step 1
                If imRSave(ilDay + 1, ilRowNo) = 0 Then
                    tgRBofRec(ilRowNo).tBof.sDays(ilDay) = "N"
                Else
                    tgRBofRec(ilRowNo).tBof.sDays(ilDay) = "Y"
                End If
            Next ilDay
            gPackTime smRSave(8, ilRowNo), tgRBofRec(ilRowNo).tBof.iStartTime(0), tgRBofRec(ilRowNo).tBof.iStartTime(1)
            gPackTime smRSave(9, ilRowNo), tgRBofRec(ilRowNo).tBof.iEndTime(0), tgRBofRec(ilRowNo).tBof.iEndTime(1)
            tgRBofRec(ilRowNo).tBof.iLen = 0
        End If
    Next ilRowNo
    igView = ilView
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
Private Sub mMoveRecToCtrl(slInType As String)
'
'   mMoveRecToCtrl slInType
'   Where:
'   Where:
'       slInType(I)-S=Suppression; R=Replacement; B=Both
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slFDay As String
    Dim slTDay As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilRowNo As Integer
    Dim ilPass As Integer
    Dim ilSPass As Integer
    Dim ilEPass As Integer
    Dim ilUpper As Integer
    Dim ilDay As Integer
    Dim ilView As Integer
    Dim slName As String
    Dim slSDate As String
    Dim slEDate As String
    If (slInType = "S") Then
        ilSPass = 1
        ilEPass = 1
    End If
    If (slInType = "R") Then
        ilSPass = 2
        ilEPass = 2
    End If
    If (slInType = "B") Then
        ilSPass = 1
        ilEPass = 2
    End If
    ilView = igView
    'Add direct to each advertiser if co
    If (smFromLog = "Y") And (ilSPass = 1) Then
        For ilRowNo = LBound(tgSBofRec) To UBound(tgSBofRec) - 1 Step 1
            ilLoop = gBinarySearchAdf(tgSBofRec(ilRowNo).tBof.iAdfCode)
            If ilLoop <> -1 Then
                If tgCommAdf(ilLoop).sBillAgyDir = "D" Then
                    tgSBofRec(ilRowNo).sAdfName = Trim$(tgSBofRec(ilRowNo).sAdfName) & "/Direct"
                End If
            End If
        Next ilRowNo
    End If
    If (smFromLog = "Y") And (ilEPass = 2) Then
        For ilRowNo = LBound(tgRBofRec) To UBound(tgRBofRec) - 1 Step 1
            ilLoop = gBinarySearchAdf(tgRBofRec(ilRowNo).tBof.iAdfCode)
            If ilLoop <> -1 Then
                If tgCommAdf(ilLoop).sBillAgyDir = "D" Then
                    tgRBofRec(ilRowNo).sAdfName = Trim$(tgRBofRec(ilRowNo).sAdfName) & "/Direct"
                End If
            End If
        Next ilRowNo
    End If
    For ilPass = ilSPass To ilEPass Step 1
        If ilPass = 1 Then
            igView = 0
            ilUpper = UBound(tgSBofRec)
'            ReDim Preserve smSShow(1 To 14, 1 To ilUpper) As String 'Values shown in program area
'            ReDim Preserve smSSave(1 To 12, 1 To ilUpper) As String    'Values saved (program name) in program area
'            ReDim Preserve imSSave(1 To 7, 1 To ilUpper) As Integer 'Values saved (program name) in program area
'            ReDim Preserve lmSSave(1 To 4, 1 To ilUpper) As Long 'Values saved (program name) in program area
            ReDim Preserve smSShow(0 To 14, 0 To ilUpper) As String 'Values shown in program area
            ReDim Preserve smSSave(0 To 12, 0 To ilUpper) As String    'Values saved (program name) in program area
            ReDim Preserve imSSave(0 To 7, 0 To ilUpper) As Integer 'Values saved (program name) in program area
            ReDim Preserve lmSSave(0 To 4, 0 To ilUpper) As Long 'Values saved (program name) in program area

            For ilLoop = LBound(smSShow, 1) To UBound(smSShow, 1) Step 1
                smSShow(ilLoop, ilUpper) = ""
            Next ilLoop
            For ilLoop = LBound(smSSave, 1) To UBound(smSSave, 1) Step 1
                smSSave(ilLoop, ilUpper) = ""
            Next ilLoop
            For ilLoop = LBound(imSSave, 1) To UBound(imSSave, 1) Step 1
                imSSave(ilLoop, ilUpper) = -1
            Next ilLoop
            For ilLoop = LBound(lmSSave, 1) To UBound(lmSSave, 1) Step 1
                lmSSave(ilLoop, ilUpper) = -1
            Next ilLoop
            For ilRowNo = LBound(tgSBofRec) To UBound(tgSBofRec) - 1 Step 1
                If smFromLog = "Y" Then
                    'Get Advertiser Name
                    If Trim$(tgSBofRec(ilRowNo).sAdfName) <> "" Then
                        smSSave(1, ilRowNo) = Trim$(tgSBofRec(ilRowNo).sAdfName)
                    Else
                        smSSave(1, ilRowNo) = "[All]"
                    End If
                    gSetShow pbcSuppression(imSRIndex), smSSave(1, ilRowNo), tmSCtrls(SADVTINDEX)
                    smSShow(SADVTINDEX, ilRowNo) = tmSCtrls(SADVTINDEX).sShow
                    'Suppress Contract
                    lmSSave(1, ilRowNo) = tgSBofRec(ilRowNo).tBof.lSChfCode
                    If tgSBofRec(ilRowNo).lSCntrNo > 0 Then
                        slStr = Trim$(str$(tgSBofRec(ilRowNo).lSCntrNo))
                    Else
                        slStr = ""
                    End If
                    smSSave(11, ilRowNo) = slStr
                    gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(SCNTRINDEX)
                    smSShow(SCNTRINDEX, ilRowNo) = tmSCtrls(SCNTRINDEX).sShow
                    'Get Replace Advertiser Name
                    If Trim$(tgSBofRec(ilRowNo).sRAdfName) <> "" Then
                        smSSave(8, ilRowNo) = Trim$(tgSBofRec(ilRowNo).sRAdfName)
                    Else
                        smSSave(8, ilRowNo) = "[None]"
                    End If
                    gSetShow pbcSuppression(imSRIndex), smSSave(8, ilRowNo), tmSCtrls(SRADVTINDEX)
                    smSShow(SRADVTINDEX, ilRowNo) = tmSCtrls(SRADVTINDEX).sShow
                    'Replace Contract
                    lmSSave(2, ilRowNo) = tgSBofRec(ilRowNo).tBof.lRChfCode
                    If tgSBofRec(ilRowNo).lRCntrNo > 0 Then
                        slStr = Trim$(str$(tgSBofRec(ilRowNo).lRCntrNo))
                    Else
                        slStr = ""
                    End If
                    smSSave(12, ilRowNo) = slStr
                    gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(SRCNTRINDEX)
                    smSShow(SRCNTRINDEX, ilRowNo) = tmSCtrls(SRCNTRINDEX).sShow
                    'Get Cart
                    smSSave(2, ilRowNo) = ""
                    lmSSave(3, ilRowNo) = tgSBofRec(ilRowNo).tBof.lCifCode
                    lmSSave(4, ilRowNo) = tgSBofRec(ilRowNo).tBof.lSifCode
                    tmCifSrchKey.lCode = tgSBofRec(ilRowNo).tBof.lCifCode
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                            If tmMcf.iCode <> tmCif.iMcfCode Then
                                tmMcfSrchKey.iCode = tmCif.iMcfCode
                                ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo mMoveRecToCtrlErr
                                gBtrvErrorMsg ilRet, "mMoveRecToCtrl (btrGetEqual):" & "Mcf.Btr", Blackout
                                On Error GoTo 0
                            End If
                            slName = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                            If (Len(Trim$(tmCif.sCut)) <> 0) Then
                                slName = slName & "-" & tmCif.sCut
                            End If
                            If tmCif.lcpfCode > 0 Then
                                tmCpfSrchKey.lCode = tmCif.lcpfCode
                                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo mMoveRecToCtrlErr
                                gBtrvErrorMsg ilRet, "mMoveRecToCtrl (btrGetEqual):" & "Cpf.Btr", Blackout
                                On Error GoTo 0
                                If Trim$(tmCpf.sISCI) <> "" Then
                                    slName = slName & " " & Trim$(tmCpf.sISCI)
                                End If
                            End If
                        Else
                            If tmCif.lcpfCode > 0 Then
                                tmCpfSrchKey.lCode = tmCif.lcpfCode
                                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo mMoveRecToCtrlErr
                                gBtrvErrorMsg ilRet, "mMoveRecToCtrl (btrGetEqual):" & "Cpf.Btr", Blackout
                                On Error GoTo 0
                                If Trim$(tmCpf.sISCI) <> "" Then
                                    slName = Trim$(tmCpf.sISCI)
                                Else
                                    slName = "ISCI Missing"
                                End If
                            Else
                                slName = "ISCI Missing"
                            End If
                        End If
                        slName = Trim$(str(tmCif.iLen)) & " " & slName
                        gUnpackDate tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), slSDate
                        gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slEDate
                        smSSave(2, ilRowNo) = slName & " " & slSDate & " " & slEDate
                    End If
                    gSetShow pbcSuppression(imSRIndex), smSSave(2, ilRowNo), tmSCtrls(SRCARTINDEX)
                    smSShow(SRCARTINDEX, ilRowNo) = tmSCtrls(SRCARTINDEX).sShow
                    'Get Vehicle
                    If Trim$(tgSBofRec(ilRowNo).sVefName) <> "" Then
                        smSSave(3, ilRowNo) = Trim$(tgSBofRec(ilRowNo).sVefName)
                    Else
                        smSSave(3, ilRowNo) = "[All]"
                    End If
                    gSetShow pbcSuppression(imSRIndex), smSSave(3, ilRowNo), tmSCtrls(SVEHINDEX)
                    smSShow(SVEHINDEX, ilRowNo) = tmSCtrls(SVEHINDEX).sShow
                    'Length
                    If tgSBofRec(ilRowNo).tBof.iLen > 0 Then
                        smSSave(9, ilRowNo) = Trim$(str$(tgSBofRec(ilRowNo).tBof.iLen))
                    Else
                        smSSave(9, ilRowNo) = "[All]"
                    End If
                    gSetShow pbcSuppression(imSRIndex), smSSave(9, ilRowNo), tmSCtrls(SLENINDEX)
                    smSShow(SLENINDEX, ilRowNo) = tmSCtrls(SLENINDEX).sShow

                    'Start Date
                    gUnpackDate tgSBofRec(ilRowNo).tBof.iStartDate(0), tgSBofRec(ilRowNo).tBof.iStartDate(1), smSSave(4, ilRowNo)
                    gSetShow pbcSuppression(imSRIndex), smSSave(4, ilRowNo), tmSCtrls(SSTARTDATEINDEX)
                    smSShow(SSTARTDATEINDEX, ilRowNo) = tmSCtrls(SSTARTDATEINDEX).sShow
                    'End Date
                    gUnpackDate tgSBofRec(ilRowNo).tBof.iEndDate(0), tgSBofRec(ilRowNo).tBof.iEndDate(1), smSSave(5, ilRowNo)
                    If smSSave(5, ilRowNo) <> "" Then
                        gSetShow pbcSuppression(imSRIndex), smSSave(5, ilRowNo), tmSCtrls(SENDDATEINDEX)
                    Else
                        gSetShow pbcSuppression(imSRIndex), "TFN", tmSCtrls(SENDDATEINDEX)
                    End If
                    smSShow(SENDDATEINDEX, ilRowNo) = tmSCtrls(SENDDATEINDEX).sShow
                    'Days
                    slFDay = ""
                    slTDay = ""
                    For ilDay = 0 To 6 Step 1
                        If tgSBofRec(ilRowNo).tBof.sDays(ilDay) <> "N" Then
                            If slFDay = "" Then
                                Select Case ilDay
                                    Case 0
                                        slFDay = "Mo"
                                    Case 1
                                        slFDay = "Tu"
                                    Case 2
                                        slFDay = "We"
                                    Case 3
                                        slFDay = "Th"
                                    Case 4
                                        slFDay = "Fr"
                                    Case 5
                                        slFDay = "Sa"
                                    Case 6
                                        slFDay = "Su"
                                End Select
                            End If
                            Select Case ilDay
                                Case 0
                                    slTDay = "Mo"
                                Case 1
                                    slTDay = "Tu"
                                Case 2
                                    slTDay = "We"
                                Case 3
                                    slTDay = "Th"
                                Case 4
                                    slTDay = "Fr"
                                Case 5
                                    slTDay = "Sa"
                                Case 6
                                    slTDay = "Su"
                            End Select
                        End If
                    Next ilDay
                    If slFDay = slTDay Then
                        slStr = slFDay
                    Else
                        slStr = slFDay & "-" & slTDay
                    End If
                    smSSave(10, ilRowNo) = slStr
                    gSetShow pbcSuppression(imSRIndex), smSSave(10, ilRowNo), tmSCtrls(SDAYINDEX)
                    smSShow(SDAYINDEX, ilRowNo) = tmSCtrls(SDAYINDEX).sShow
                    'Start Time
                    gUnpackTime tgSBofRec(ilRowNo).tBof.iStartTime(0), tgSBofRec(ilRowNo).tBof.iStartTime(1), "A", "1", smSSave(6, ilRowNo)
                    gSetShow pbcSuppression(imSRIndex), smSSave(6, ilRowNo), tmSCtrls(SSTARTTIMEINDEX)
                    smSShow(SSTARTTIMEINDEX, ilRowNo) = tmSCtrls(SSTARTTIMEINDEX).sShow
                    'End Time
                    gUnpackTime tgSBofRec(ilRowNo).tBof.iEndTime(0), tgSBofRec(ilRowNo).tBof.iEndTime(1), "A", "1", smSSave(7, ilRowNo)
                    gSetShow pbcSuppression(imSRIndex), smSSave(7, ilRowNo), tmSCtrls(SENDTIMEINDEX)
                    smSShow(SENDTIMEINDEX, ilRowNo) = tmSCtrls(SENDTIMEINDEX).sShow
                Else
                    'Get Advertiser Name
                    If Trim$(tgSBofRec(ilRowNo).sAdfName) <> "" Then
                        smSSave(1, ilRowNo) = Trim$(tgSBofRec(ilRowNo).sAdfName)
                    Else
                        smSSave(1, ilRowNo) = "[None]"
                    End If
                    gSetShow pbcSuppression(imSRIndex), smSSave(1, ilRowNo), tmSCtrls(SADVTINDEX)
                    smSShow(SADVTINDEX, ilRowNo) = tmSCtrls(SADVTINDEX).sShow
                    'Get Short Title
                    smSSave(2, ilRowNo) = Trim$(tgSBofRec(ilRowNo).sShtTitle)
                    gSetShow pbcSuppression(imSRIndex), smSSave(2, ilRowNo), tmSCtrls(SSHORTTITLEINDEX)
                    smSShow(SSHORTTITLEINDEX, ilRowNo) = tmSCtrls(SSHORTTITLEINDEX).sShow
                    'Get Vehicle
                    smSSave(3, ilRowNo) = Trim$(tgSBofRec(ilRowNo).sVefName)
                    gSetShow pbcSuppression(imSRIndex), smSSave(3, ilRowNo), tmSCtrls(SVEHINDEX)
                    smSShow(SVEHINDEX, ilRowNo) = tmSCtrls(SVEHINDEX).sShow
                    'Start Date
                    gUnpackDate tgSBofRec(ilRowNo).tBof.iStartDate(0), tgSBofRec(ilRowNo).tBof.iStartDate(1), smSSave(4, ilRowNo)
                    gSetShow pbcSuppression(imSRIndex), smSSave(4, ilRowNo), tmSCtrls(SSTARTDATEINDEX)
                    smSShow(SSTARTDATEINDEX, ilRowNo) = tmSCtrls(SSTARTDATEINDEX).sShow
                    'End Date
                    gUnpackDate tgSBofRec(ilRowNo).tBof.iEndDate(0), tgSBofRec(ilRowNo).tBof.iEndDate(1), smSSave(5, ilRowNo)
                    If smSSave(5, ilRowNo) <> "" Then
                        gSetShow pbcSuppression(imSRIndex), smSSave(5, ilRowNo), tmSCtrls(SENDDATEINDEX)
                    Else
                        gSetShow pbcSuppression(imSRIndex), "TFN", tmSCtrls(SENDDATEINDEX)
                    End If
                    smSShow(SENDDATEINDEX, ilRowNo) = tmSCtrls(SENDDATEINDEX).sShow
                    'Days
                    For ilDay = 0 To 6 Step 1
                        If tgSBofRec(ilRowNo).tBof.sDays(ilDay) = "N" Then
                            imSSave(ilDay + 1, ilRowNo) = 0
                            gSetShow pbcSuppression(imSRIndex), "  ", tmSCtrls(SDAYINDEX + ilDay)
                        Else
                            imSSave(ilDay + 1, ilRowNo) = 1
                            gSetShow pbcSuppression(imSRIndex), "4", tmSCtrls(SDAYINDEX + ilDay)
                        End If
                        smSShow(SDAYINDEX + ilDay, ilRowNo) = tmSCtrls(SDAYINDEX + ilDay).sShow
                    Next ilDay
                    'Start Time
                    gUnpackTime tgSBofRec(ilRowNo).tBof.iStartTime(0), tgSBofRec(ilRowNo).tBof.iStartTime(1), "A", "1", smSSave(6, ilRowNo)
                    gSetShow pbcSuppression(imSRIndex), smSSave(6, ilRowNo), tmSCtrls(SSTARTTIMEINDEX)
                    smSShow(SSTARTTIMEINDEX, ilRowNo) = tmSCtrls(SSTARTTIMEINDEX).sShow
                    'End Time
                    gUnpackTime tgSBofRec(ilRowNo).tBof.iEndTime(0), tgSBofRec(ilRowNo).tBof.iEndTime(1), "A", "1", smSSave(7, ilRowNo)
                    gSetShow pbcSuppression(imSRIndex), smSSave(7, ilRowNo), tmSCtrls(SENDTIMEINDEX)
                    smSShow(SENDTIMEINDEX, ilRowNo) = tmSCtrls(SENDTIMEINDEX).sShow
                End If
            Next ilRowNo
        Else
            igView = 1
            ilUpper = UBound(tgRBofRec)
'            ReDim Preserve smRShow(1 To 16, 1 To ilUpper) As String 'Values shown in program area
'            ReDim Preserve smRSave(1 To 11, 1 To ilUpper) As String    'Values saved (program name) in program area
'            ReDim Preserve imRSave(1 To 7, 1 To ilUpper) As Integer 'Values saved (program name) in program area
'            ReDim Preserve lmRSave(1 To 3, 1 To ilUpper) As Long 'Values saved (program name) in program area
            ReDim Preserve smRShow(0 To 16, 0 To ilUpper) As String 'Values shown in program area
            ReDim Preserve smRSave(0 To 11, 0 To ilUpper) As String    'Values saved (program name) in program area
            ReDim Preserve imRSave(0 To 7, 0 To ilUpper) As Integer 'Values saved (program name) in program area
            ReDim Preserve lmRSave(0 To 3, 0 To ilUpper) As Long 'Values saved (program name) in program area

            For ilLoop = LBound(smRShow, 1) To UBound(smRShow, 1) Step 1
                smRShow(ilLoop, ilUpper) = ""
            Next ilLoop
            For ilLoop = LBound(smRSave, 1) To UBound(smRSave, 1) Step 1
                smRSave(ilLoop, ilUpper) = ""
            Next ilLoop
            For ilLoop = LBound(imRSave, 1) To UBound(imRSave, 1) Step 1
                imRSave(ilLoop, ilUpper) = -1
            Next ilLoop

            For ilRowNo = LBound(tgRBofRec) To UBound(tgRBofRec) - 1 Step 1
                If smFromLog = "Y" Then
                    'Get Advertiser Name
                    smRSave(1, ilRowNo) = Trim$(tgRBofRec(ilRowNo).sAdfName)
                    gSetShow pbcReplacement(imSRIndex), smRSave(1, ilRowNo), tmRCtrls(RADVTINDEX)
                    smRShow(RADVTINDEX, ilRowNo) = tmRCtrls(RADVTINDEX).sShow
                    'Replace Contract
                    lmRSave(3, ilRowNo) = tgRBofRec(ilRowNo).tBof.lRChfCode
                    If tgRBofRec(ilRowNo).lRCntrNo > 0 Then
                        slStr = Trim$(str$(tgRBofRec(ilRowNo).lRCntrNo))
                    Else
                        slStr = ""
                    End If
                    smRSave(11, ilRowNo) = slStr
                    gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(RCNTRINDEX)
                    smRShow(RCNTRINDEX, ilRowNo) = tmRCtrls(RCNTRINDEX).sShow
                    'Get Cart
                    smRSave(2, ilRowNo) = ""
                    smRSave(3, ilRowNo) = ""
                    lmRSave(1, ilRowNo) = tgRBofRec(ilRowNo).tBof.lCifCode
                    lmRSave(2, ilRowNo) = tgRBofRec(ilRowNo).tBof.lSifCode
                    tmCifSrchKey.lCode = tgRBofRec(ilRowNo).tBof.lCifCode
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                            If tmMcf.iCode <> tmCif.iMcfCode Then
                                tmMcfSrchKey.iCode = tmCif.iMcfCode
                                ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo mMoveRecToCtrlErr
                                gBtrvErrorMsg ilRet, "mMoveRecToCtrl (btrGetEqual):" & "Mcf.Btr", Blackout
                                On Error GoTo 0
                            End If
                            slName = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                            If (Len(Trim$(tmCif.sCut)) <> 0) Then
                                slName = slName & "-" & tmCif.sCut
                            End If
                            If tmCif.lcpfCode > 0 Then
                                tmCpfSrchKey.lCode = tmCif.lcpfCode
                                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo mMoveRecToCtrlErr
                                gBtrvErrorMsg ilRet, "mMoveRecToCtrl (btrGetEqual):" & "Cpf.Btr", Blackout
                                On Error GoTo 0
                                If Trim$(tmCpf.sISCI) <> "" Then
                                    slName = slName & " " & Trim$(tmCpf.sISCI)
                                End If
                            End If
                        Else
                            If tmCif.lcpfCode > 0 Then
                                tmCpfSrchKey.lCode = tmCif.lcpfCode
                                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo mMoveRecToCtrlErr
                                gBtrvErrorMsg ilRet, "mMoveRecToCtrl (btrGetEqual):" & "Cpf.Btr", Blackout
                                On Error GoTo 0
                                If Trim$(tmCpf.sISCI) <> "" Then
                                    slName = Trim$(tmCpf.sISCI)
                                Else
                                    slName = "ISCI Missing"
                                End If
                            Else
                                slName = "ISCI Missing"
                            End If
                        End If
                        slName = Trim$(str(tmCif.iLen)) & " " & slName
                        'Not including Short Title or Product as part of Cart Name because Product is coming
                        'from Contract and it sets sifcode to zero so we never match
                        'If tgSpf.sUseProdSptScr = "P" Then  'Short Title
                        '    tmSifSrchKey.lCode = tgRBofRec(ilRowNo).tBof.lSifCode
                        '    ilRet = btrGetEqual(hmSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        '    If ilRet = BTRV_ERR_NONE Then
                        '        slName = slName & " " & Trim$(tmSif.sName)
                        '        smRSave(3, ilRowNo) = Trim$(tmSif.sName)
                        '    End If
                        'Else
                        '    tmPrfSrchKey.lCode = tgRBofRec(ilRowNo).tBof.lSifCode
                        '    ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        '    If ilRet = BTRV_ERR_NONE Then
                        '        slName = slName & " " & Trim$(tmPrf.sName)
                        '        smRSave(3, ilRowNo) = Trim$(tmPrf.sName)
                        '    End If
                        'End If
                        gUnpackDate tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), slSDate
                        gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slEDate
                        smRSave(2, ilRowNo) = slName & " " & slSDate & " " & slEDate
                    End If
                    gSetShow pbcReplacement(imSRIndex), smRSave(2, ilRowNo), tmRCtrls(RCARTINDEX)
                    smRShow(RCARTINDEX, ilRowNo) = tmRCtrls(RCARTINDEX).sShow
                    'Get Vehicle
                    If Trim$(tgRBofRec(ilRowNo).sVefName) <> "" Then
                        smRSave(3, ilRowNo) = Trim$(tgRBofRec(ilRowNo).sVefName)
                    Else
                        smRSave(3, ilRowNo) = "[All]"
                    End If
                    gSetShow pbcReplacement(imSRIndex), smRSave(3, ilRowNo), tmRCtrls(RVEHINDEX)
                    smRShow(RVEHINDEX, ilRowNo) = tmRCtrls(RVEHINDEX).sShow
                    smRSave(4, ilRowNo) = ""
                    If smSplitFill <> "Y" Then
                        For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                            slNameCode = tgCompCode(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = tgRBofRec(ilRowNo).tBof.iMnfComp(0) Then
                                ilRet = gParseItem(slNameCode, 1, "\", smRSave(4, ilRowNo))
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    gSetShow pbcReplacement(imSRIndex), smRSave(4, ilRowNo), tmRCtrls(RPPINDEX)
                    smRShow(RPPINDEX, ilRowNo) = tmRCtrls(RPPINDEX).sShow
                    smRSave(5, ilRowNo) = ""
                    If smSplitFill <> "Y" Then
                        For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                            slNameCode = tgCompCode(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = tgRBofRec(ilRowNo).tBof.iMnfComp(1) Then
                                ilRet = gParseItem(slNameCode, 1, "\", smRSave(5, ilRowNo))
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    gSetShow pbcReplacement(imSRIndex), smRSave(5, ilRowNo), tmRCtrls(RPPINDEX + 1)
                    smRShow(RPPINDEX + 1, ilRowNo) = tmRCtrls(RPPINDEX + 1).sShow

                    'Start Date
                    gUnpackDate tgRBofRec(ilRowNo).tBof.iStartDate(0), tgRBofRec(ilRowNo).tBof.iStartDate(1), smRSave(6, ilRowNo)
                    gSetShow pbcReplacement(imSRIndex), smRSave(6, ilRowNo), tmRCtrls(RSTARTDATEINDEX)
                    smRShow(RSTARTDATEINDEX, ilRowNo) = tmRCtrls(RSTARTDATEINDEX).sShow
                    'End Date
                    gUnpackDate tgRBofRec(ilRowNo).tBof.iEndDate(0), tgRBofRec(ilRowNo).tBof.iEndDate(1), smRSave(7, ilRowNo)
                    gSetShow pbcReplacement(imSRIndex), smRSave(7, ilRowNo), tmRCtrls(RENDDATEINDEX)
                    smRShow(RENDDATEINDEX, ilRowNo) = tmRCtrls(RENDDATEINDEX).sShow
                    'Days
                    slFDay = ""
                    slTDay = ""
                    For ilDay = 0 To 6 Step 1
                        If tgRBofRec(ilRowNo).tBof.sDays(ilDay) <> "N" Then
                            If slFDay = "" Then
                                Select Case ilDay
                                    Case 0
                                        slFDay = "Mo"
                                    Case 1
                                        slFDay = "Tu"
                                    Case 2
                                        slFDay = "We"
                                    Case 3
                                        slFDay = "Th"
                                    Case 4
                                        slFDay = "Fr"
                                    Case 5
                                        slFDay = "Sa"
                                    Case 6
                                        slFDay = "Su"
                                End Select
                            End If
                            Select Case ilDay
                                Case 0
                                    slTDay = "Mo"
                                Case 1
                                    slTDay = "Tu"
                                Case 2
                                    slTDay = "We"
                                Case 3
                                    slTDay = "Th"
                                Case 4
                                    slTDay = "Fr"
                                Case 5
                                    slTDay = "Sa"
                                Case 6
                                    slTDay = "Su"
                            End Select
                        End If
                    Next ilDay
                    If slFDay = slTDay Then
                        slStr = slFDay
                    Else
                        slStr = slFDay & "-" & slTDay
                    End If
                    smRSave(10, ilRowNo) = slStr
                    gSetShow pbcReplacement(imSRIndex), smRSave(10, ilRowNo), tmRCtrls(RDAYINDEX)
                    smRShow(RDAYINDEX, ilRowNo) = tmRCtrls(RDAYINDEX).sShow
                    'Start Time
                    gUnpackTime tgRBofRec(ilRowNo).tBof.iStartTime(0), tgRBofRec(ilRowNo).tBof.iStartTime(1), "A", "1", smRSave(8, ilRowNo)
                    gSetShow pbcReplacement(imSRIndex), smRSave(8, ilRowNo), tmRCtrls(RSTARTTIMEINDEX)
                    smRShow(RSTARTTIMEINDEX, ilRowNo) = tmRCtrls(RSTARTTIMEINDEX).sShow
                    'End Time
                    gUnpackTime tgRBofRec(ilRowNo).tBof.iEndTime(0), tgRBofRec(ilRowNo).tBof.iEndTime(1), "A", "1", smRSave(9, ilRowNo)
                    gSetShow pbcReplacement(imSRIndex), smRSave(9, ilRowNo), tmRCtrls(RENDTIMEINDEX)
                    smRShow(RENDTIMEINDEX, ilRowNo) = tmRCtrls(RENDTIMEINDEX).sShow
                Else
                    'Get Advertiser Name
                    smRSave(1, ilRowNo) = Trim$(tgRBofRec(ilRowNo).sAdfName)
                    gSetShow pbcReplacement(imSRIndex), smRSave(1, ilRowNo), tmRCtrls(RADVTINDEX)
                    smRShow(RADVTINDEX, ilRowNo) = tmRCtrls(RADVTINDEX).sShow
                    'Get Cart
                    smRSave(2, ilRowNo) = ""
                    smRSave(3, ilRowNo) = ""
                    lmRSave(1, ilRowNo) = tgRBofRec(ilRowNo).tBof.lCifCode
                    lmRSave(2, ilRowNo) = tgRBofRec(ilRowNo).tBof.lSifCode
                    tmCifSrchKey.lCode = tgRBofRec(ilRowNo).tBof.lCifCode
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                            If tmMcf.iCode <> tmCif.iMcfCode Then
                                tmMcfSrchKey.iCode = tmCif.iMcfCode
                                ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo mMoveRecToCtrlErr
                                gBtrvErrorMsg ilRet, "mMoveRecToCtrl (btrGetEqual):" & "Mcf.Btr", Blackout
                                On Error GoTo 0
                            End If
                            slName = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                            If (Len(Trim$(tmCif.sCut)) <> 0) Then
                                slName = slName & "-" & tmCif.sCut
                            End If
                            If tmCif.lcpfCode > 0 Then
                                tmCpfSrchKey.lCode = tmCif.lcpfCode
                                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo mMoveRecToCtrlErr
                                gBtrvErrorMsg ilRet, "mMoveRecToCtrl (btrGetEqual):" & "Cpf.Btr", Blackout
                                On Error GoTo 0
                                If Trim$(tmCpf.sISCI) <> "" Then
                                    slName = slName & " " & Trim$(tmCpf.sISCI)
                                End If
                            End If
                        Else
                            If tmCif.lcpfCode > 0 Then
                                tmCpfSrchKey.lCode = tmCif.lcpfCode
                                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo mMoveRecToCtrlErr
                                gBtrvErrorMsg ilRet, "mMoveRecToCtrl (btrGetEqual):" & "Cpf.Btr", Blackout
                                On Error GoTo 0
                                If Trim$(tmCpf.sISCI) <> "" Then
                                    slName = Trim$(tmCpf.sISCI)
                                Else
                                    slName = "ISCI Missing"
                                End If
                            Else
                                slName = "ISCI Missing"
                            End If
                        End If
                        slName = Trim$(str(tmCif.iLen)) & " " & slName
                        If tgSpf.sUseProdSptScr = "P" Then  'Short Title
                            tmSifSrchKey.lCode = tgRBofRec(ilRowNo).tBof.lSifCode
                            ilRet = btrGetEqual(hmSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                slName = slName & " " & Trim$(tmSif.sName)
                                smRSave(3, ilRowNo) = Trim$(tmSif.sName)
                            End If
                        Else
                            tmPrfSrchKey.lCode = tgRBofRec(ilRowNo).tBof.lSifCode
                            ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                slName = slName & " " & Trim$(tmPrf.sName)
                                smRSave(3, ilRowNo) = Trim$(tmPrf.sName)
                            End If
                        End If
                        gUnpackDate tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), slSDate
                        gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slEDate
                        smRSave(2, ilRowNo) = slName & " " & slSDate & " " & slEDate
                    End If
                    gSetShow pbcReplacement(imSRIndex), smRSave(2, ilRowNo), tmRCtrls(RCARTINDEX)
                    smRShow(RCARTINDEX, ilRowNo) = tmRCtrls(RCARTINDEX).sShow
                    'Get Short Title/Product
                    gSetShow pbcReplacement(imSRIndex), smRSave(3, ilRowNo), tmRCtrls(RSHORTTITLEINDEX)
                    smRShow(RSHORTTITLEINDEX, ilRowNo) = tmRCtrls(RSHORTTITLEINDEX).sShow
                    'Product Protection
                    smRSave(4, ilRowNo) = ""
                    For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                        slNameCode = tgCompCode(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = tgRBofRec(ilRowNo).tBof.iMnfComp(0) Then
                            ilRet = gParseItem(slNameCode, 1, "\", smRSave(4, ilRowNo))
                            Exit For
                        End If
                    Next ilLoop
                    gSetShow pbcReplacement(imSRIndex), smRSave(4, ilRowNo), tmRCtrls(RPPINDEX)
                    smRShow(RPPINDEX, ilRowNo) = tmRCtrls(RPPINDEX).sShow
                    smRSave(5, ilRowNo) = ""
                    For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                        slNameCode = tgCompCode(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = tgRBofRec(ilRowNo).tBof.iMnfComp(1) Then
                            ilRet = gParseItem(slNameCode, 1, "\", smRSave(5, ilRowNo))
                            Exit For
                        End If
                    Next ilLoop
                    gSetShow pbcReplacement(imSRIndex), smRSave(5, ilRowNo), tmRCtrls(RPPINDEX + 1)
                    smRShow(RPPINDEX + 1, ilRowNo) = tmRCtrls(RPPINDEX + 1).sShow

                    'Start Date
                    gUnpackDate tgRBofRec(ilRowNo).tBof.iStartDate(0), tgRBofRec(ilRowNo).tBof.iStartDate(1), smRSave(6, ilRowNo)
                    gSetShow pbcReplacement(imSRIndex), smRSave(6, ilRowNo), tmRCtrls(RSTARTDATEINDEX)
                    smRShow(RSTARTDATEINDEX, ilRowNo) = tmRCtrls(RSTARTDATEINDEX).sShow
                    'End Date
                    gUnpackDate tgRBofRec(ilRowNo).tBof.iEndDate(0), tgRBofRec(ilRowNo).tBof.iEndDate(1), smRSave(7, ilRowNo)
                    gSetShow pbcReplacement(imSRIndex), smRSave(7, ilRowNo), tmRCtrls(RENDDATEINDEX)
                    smRShow(RENDDATEINDEX, ilRowNo) = tmRCtrls(RENDDATEINDEX).sShow
                    'Days
                    For ilDay = 0 To 6 Step 1
                        If tgRBofRec(ilRowNo).tBof.sDays(ilDay) = "N" Then
                            imRSave(ilDay + 1, ilRowNo) = 0
                            gSetShow pbcReplacement(imSRIndex), "  ", tmRCtrls(RDAYINDEX + ilDay)
                        Else
                            imRSave(ilDay + 1, ilRowNo) = 1
                            gSetShow pbcReplacement(imSRIndex), "4", tmRCtrls(RDAYINDEX + ilDay)
                        End If
                        smRShow(RDAYINDEX + ilDay, ilRowNo) = tmRCtrls(RDAYINDEX + ilDay).sShow
                    Next ilDay
                    'Start Time
                    gUnpackTime tgRBofRec(ilRowNo).tBof.iStartTime(0), tgRBofRec(ilRowNo).tBof.iStartTime(1), "A", "1", smRSave(8, ilRowNo)
                    gSetShow pbcReplacement(imSRIndex), smRSave(8, ilRowNo), tmRCtrls(RSTARTTIMEINDEX)
                    smRShow(RSTARTTIMEINDEX, ilRowNo) = tmRCtrls(RSTARTTIMEINDEX).sShow
                    'End Time
                    gUnpackTime tgRBofRec(ilRowNo).tBof.iEndTime(0), tgRBofRec(ilRowNo).tBof.iEndTime(1), "A", "1", smRSave(9, ilRowNo)
                    gSetShow pbcReplacement(imSRIndex), smRSave(9, ilRowNo), tmRCtrls(RENDTIMEINDEX)
                    smRShow(RENDTIMEINDEX, ilRowNo) = tmRCtrls(RENDTIMEINDEX).sShow
                End If
            Next ilRowNo
        End If
    Next ilPass
    igView = ilView
    mSetMinMax
    Exit Sub
mMoveRecToCtrlErr:
    On Error GoTo 0
    imTerminate = True
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
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slHelpSystem As String

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
        imShowHelpMsg = True
        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
            imShowHelpMsg = False
        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    smFromLog = "N"
    smSplitFill = "N"
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get user name
    If StrComp(slStr, "Log", 1) = 0 Then
        smFromLog = "Y"
    End If
    If StrComp(slStr, "SplitFill", 1) = 0 Then
        smFromLog = "Y"
        smSplitFill = "Y"
    End If
    'gInitStdAlone Blackout, slStr, ilTestSystem
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRCntrPop                        *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Contract Population            *
'*                                                     *
'*******************************************************
Private Sub mRCntrPop(ilRowNo As Integer)
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilCurrent As Integer
    Dim ilShow As Integer
    Dim ilState As Integer
    Dim ilAAS As Integer
    Dim ilAASCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    'If Not imPropPopReqd Then
    '    Exit Sub
    'End If
    Screen.MousePointer = vbHourglass
    ilIndex = lbcRCntr.ListIndex
    If ilIndex >= 0 Then
        slName = lbcRCntr.List(ilIndex)
    End If
    slCntrStatus = "OH" 'Hold and Orders
    'slCntrType = "C" 'Standard only
    slCntrType = "CTRQSM" 'Standard only
    ilCurrent = 0
    ilShow = 2
    ilState = 1
    'ilRet = gPopCntrForAASBox(CntrProj, -1, ilAASCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcPropNo, Traffic!lbcSCntrCode)
    ilAAS = -1
    ilAASCode = 0
    'If ilRowNo > 0 Then
    If ilRowNo >= 0 Then
        gFindMatch smRSave(1, ilRowNo), 0, lbcRAdvt
        If gLastFound(lbcRAdvt) >= 0 Then
            slNameCode = tgAdvertiser(gLastFound(lbcRAdvt)).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcSAdvt) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilAAS = 0   'By Advertiser
            ilAASCode = Val(slCode)
        Else
            lbcRCntr.Clear
            sgPlannerTag = ""
            ReDim tgPlanner(0 To 0) As SORTCODE
            '9/5/06: Disallow [all] as mRTestSaveFields requires contract to be defined.
            'If tgSpf.sUseProdSptScr <> "P" Then
            '    lbcRCntr.AddItem "[All]", 0
            'End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        lbcRCntr.Clear
        '9/5/06: Disallow [all] as mRTestSaveFields requires contract to be defined.
        'If tgSpf.sUseProdSptScr <> "P" Then
        '    lbcRCntr.AddItem "[All]", 0
        'End If
        sgPlannerTag = ""
        ReDim tgPlanner(0 To 0) As SORTCODE
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ilRet = gPopCntrForAASBox(Blackout, ilAAS, ilAASCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcRCntr, tgPlanner(), sgPlannerTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mRCntrErr
        gCPErrorMsg ilRet, "mPropNo (gPopCntrForAASBox)", Blackout
        On Error GoTo 0
        '9/5/06: Disallow [all] as mRTestSaveFields requires contract to be defined.
        'If tgSpf.sUseProdSptScr <> "P" Then
        '    lbcRCntr.AddItem "[All]", 0
        'End If
        imChgMode = True
        If ilIndex >= 0 Then
            gFindMatch slName, 0, lbcRCntr
            If gLastFound(lbcRCntr) >= 0 Then
                lbcRCntr.ListIndex = gLastFound(lbcRCntr)
            Else
                lbcRCntr.ListIndex = -1
            End If
        Else
            lbcRCntr.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mRCntrErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mREnableBox                     *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mREnableBox(ilBoxNo As Integer)
'
'   mREnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim ilSDay As Integer
    Dim ilEDay As Integer
    If (ilBoxNo < imLBRCtrls) Or (ilBoxNo > UBound(tmRCtrls)) Then
        Exit Sub
    End If

    If (imRRowNo < vbcSR.Value) Or (imRRowNo >= vbcSR.Value + vbcSR.LargeChange + 1) Then
        'mSetShow ilBoxNo
        pbcArrow.Visible = False
        lacSFrame(imSRIndex).Visible = False
        Exit Sub
    End If
    lacRFrame(imSRIndex).Move 0, tmRCtrls(RADVTINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15) - 30
    lacRFrame(imSRIndex).Visible = True
    pbcArrow.Move pbcArrow.Left, plcBlackout.Top + tmRCtrls(RADVTINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    If smFromLog = "Y" Then
        Select Case ilBoxNo 'Branch on box type (control)
            Case RADVTINDEX 'Advertiser
                mAdvtPop
                If imTerminate Then
                    Exit Sub
                End If
                lbcRAdvt.height = gListBoxHeight(lbcRAdvt.ListCount, 10)
                edcDropDown.Width = tmRCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 40
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(ilBoxNo).fBoxX, tmRCtrls(ilBoxNo).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcRAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imRRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcRAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcRAdvt.Move edcDropDown.Left, edcDropDown.Top - lbcRAdvt.height
                End If
                imChgMode = True
                slStr = Trim$(smRSave(1, imRRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcRAdvt
                    If gLastFound(lbcRAdvt) >= 0 Then
                        lbcRAdvt.ListIndex = gLastFound(lbcRAdvt)
                    Else
                        lbcRAdvt.ListIndex = 0
                    End If
                Else
                    lbcRAdvt.ListIndex = 0
                End If
                If lbcRAdvt.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcRAdvt.List(lbcRAdvt.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RCNTRINDEX
                mRCntrPop imRRowNo
                If imTerminate Then
                    Exit Sub
                End If
                lbcRCntr.height = gListBoxHeight(lbcRCntr.ListCount, 10)
                edcDropDown.Width = tmRCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 0
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(ilBoxNo).fBoxX, tmRCtrls(ilBoxNo).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcRCntr.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imRRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcRCntr.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcRCntr.Move edcDropDown.Left, edcDropDown.Top - lbcRCntr.height
                End If
                imChgMode = True
                If lmRSave(3, imRRowNo) >= 0 Then
                    ilFound = False
                    For ilLoop = 0 To UBound(tgPlanner) - 1 Step 1
                        slNameCode = tgPlanner(ilLoop).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcSAdvt) - 1)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = lmRSave(3, imRRowNo) Then
                            If tgSpf.sUseProdSptScr = "P" Then
                                lbcRCntr.ListIndex = ilLoop
                            Else
                                '9/5/06: Disallow [all] as mRTestSaveFields requires contract to be defined.
                                'lbcRCntr.ListIndex = ilLoop + 1
                                lbcRCntr.ListIndex = ilLoop
                            End If
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If (Not ilFound) And (lbcRCntr.ListCount > 0) Then
                        lbcRCntr.ListIndex = 0
                    End If
                Else
                    If lbcRCntr.ListCount > 0 Then
                        lbcRCntr.ListIndex = 0
                    End If
                End If
                If lbcRCntr.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcRCntr.List(lbcRCntr.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RCARTINDEX
                mCartPop imRRowNo
                If imTerminate Then
                    Exit Sub
                End If
                lbcCart.height = gListBoxHeight(lbcCart.ListCount, 10)
                'edcDropDown.Width = tmRCtrls(ilBoxNo).fBoxW + tmRCtrls(RVEHINDEX).fBoxW + tmRCtrls(RPPINDEX).fBoxW + tmRCtrls(RPPINDEX + 1).fBoxW + tmRCtrls(RSTARTDATEINDEX).fBoxW + tmRCtrls(RENDDATEINDEX).fBoxW + tmRCtrls(RDAYINDEX).fBoxW - cmcDropDown.Width
                edcDropDown.Width = lbcCart.Width - cmcDropDown.Width
                edcDropDown.MaxLength = 0
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(ilBoxNo).fBoxX, tmRCtrls(ilBoxNo).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcCart.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imRRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcCart.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcCart.Move edcDropDown.Left, edcDropDown.Top - lbcCart.height
                End If
                imChgMode = True
                slStr = Trim$(smRSave(2, imRRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcCart
                    If gLastFound(lbcCart) >= 0 Then
                        lbcCart.ListIndex = gLastFound(lbcCart)
                    Else
                        lbcCart.ListIndex = 0
                    End If
                Else
                    If lbcCart.ListCount >= 1 Then
                        lbcCart.ListIndex = 0
                    End If
                End If
                If lbcCart.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcCart.List(lbcCart.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RPPINDEX, RPPINDEX + 1
                mCompPop
                If imTerminate Then
                    Exit Sub
                End If
                lbcComp(ilBoxNo - RPPINDEX).height = gListBoxHeight(lbcComp(ilBoxNo - RPPINDEX).ListCount, 10)
                edcDropDown.Width = tmRCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 20
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(ilBoxNo).fBoxX, tmRCtrls(ilBoxNo).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcComp(ilBoxNo - RPPINDEX).Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imRRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcComp(ilBoxNo - RPPINDEX).Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcComp(ilBoxNo - RPPINDEX).Move edcDropDown.Left, edcDropDown.Top - lbcComp(ilBoxNo - RPPINDEX).height
                End If
                imChgMode = True
                slStr = Trim$(smRSave(4 + ilBoxNo - RPPINDEX, imRRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcComp(ilBoxNo - RPPINDEX)
                    If gLastFound(lbcComp(ilBoxNo - RPPINDEX)) >= 0 Then
                        lbcComp(ilBoxNo - RPPINDEX).ListIndex = gLastFound(lbcComp(ilBoxNo - RPPINDEX))
                    Else
                        lbcComp(ilBoxNo - RPPINDEX).ListIndex = 0
                    End If
                Else
                    lbcComp(ilBoxNo - RPPINDEX).ListIndex = 0
                End If
                If lbcComp(ilBoxNo - RPPINDEX).ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcComp(ilBoxNo - RPPINDEX).List(lbcComp(ilBoxNo - RPPINDEX).ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RVEHINDEX
                mVehPop
                If imTerminate Then
                    Exit Sub
                End If
                lbcVehicle.height = gListBoxHeight(lbcVehicle.ListCount, 10)
                edcDropDown.Width = tmRCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                If tgSpf.iVehLen <= 40 Then
                    edcDropDown.MaxLength = tgSpf.iVehLen
                Else
                    edcDropDown.MaxLength = 20
                End If
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(ilBoxNo).fBoxX, tmRCtrls(ilBoxNo).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imRRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.height
                End If
                imChgMode = True
                slStr = Trim$(smRSave(3, imRRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcVehicle
                    If gLastFound(lbcVehicle) >= 0 Then
                        lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                    Else
                        lbcVehicle.ListIndex = 0
                    End If
                Else
                    lbcVehicle.ListIndex = 0
                End If
                If lbcVehicle.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RSTARTDATEINDEX 'Start date
                edcDropDown.Width = tmRCtrls(RSTARTDATEINDEX).fBoxW + cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(RSTARTDATEINDEX).fBoxX, tmRCtrls(RSTARTDATEINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcCalendar.height < cmcDone.Top Then
                    plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.height
                End If
                If smRSave(6, imRRowNo) = "" Then
                    slStr = Trim$(smRSave(2, imRRowNo))
                    If (slStr <> "") And (slStr <> "[None]") Then
                        gFindMatch slStr, 0, lbcCart
                        If gLastFound(lbcCart) >= 0 Then
                            lbcCart.ListIndex = gLastFound(lbcCart)
                            '9/5/06: Disallow [None] if from Log and SplitFill
                            If smSplitFill <> "Y" Then
                                slNameCode = tmCartCode(gLastFound(lbcCart) - 1).sKey  'Traffic!lbcRAdvt.List(gLastFound(lbcRAdvt) - 1)
                            Else
                                slNameCode = tmCartCode(gLastFound(lbcCart)).sKey  'Traffic!lbcRAdvt.List(gLastFound(lbcRAdvt) - 1)
                            End If
                            ilRet = gParseItem(slNameCode, 5, "\", slStr)
                            If slStr = "" Then
                                slStr = gObtainMondayFromToday()
                            End If
                        Else
                            slStr = gObtainMondayFromToday()
                        End If
                    Else
                        slStr = gObtainMondayFromToday()
                    End If
                Else
                    slStr = smRSave(6, imRRowNo)
                End If
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
                edcDropDown.Text = slStr
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                If smRSave(6, imRRowNo) = "" Then
                    pbcCalendar.Visible = True
                End If
                edcDropDown.SetFocus
            Case RENDDATEINDEX 'Start date
                edcDropDown.Width = tmRCtrls(RENDDATEINDEX).fBoxW + cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(RENDDATEINDEX).fBoxX, tmRCtrls(RENDDATEINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcCalendar.height < cmcDone.Top Then
                    plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top + edcDropDown.height
                Else
                    plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top - plcCalendar.height
                End If
                If smRSave(7, imRRowNo) <> "" Then
                    slStr = smRSave(7, imRRowNo)
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint
                Else
                    slStr = Trim$(smRSave(2, imRRowNo))
                    If (slStr <> "") And (slStr <> "[None]") Then
                        gFindMatch slStr, 0, lbcCart
                        If gLastFound(lbcCart) >= 0 Then
                            lbcCart.ListIndex = gLastFound(lbcCart)
                            '9/5/06: Disallow [None] if from Log and SplitFill
                            If smSplitFill <> "Y" Then
                                slNameCode = tmCartCode(gLastFound(lbcCart) - 1).sKey  'Traffic!lbcRAdvt.List(gLastFound(lbcRAdvt) - 1)
                            Else
                                slNameCode = tmCartCode(gLastFound(lbcCart)).sKey  'Traffic!lbcRAdvt.List(gLastFound(lbcRAdvt) - 1)
                            End If
                            ilRet = gParseItem(slNameCode, 6, "\", slStr)
                            If slStr = "" Then
                                slStr = Format$(gDateValue(smRSave(6, imRRowNo)) + 6, "m/d/yy")
                            End If
                        Else
                            slStr = Format$(gDateValue(smRSave(6, imRRowNo)) + 6, "m/d/yy")
                        End If
                    Else
                        slStr = gObtainMondayFromToday()
                        slStr = Format$(gDateValue(slStr) + 6, "m/d/yy")
                    End If
                End If
                edcDropDown.Text = slStr
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RDAYINDEX  'Day index
                lbcDays.height = gListBoxHeight(lbcDays.ListCount, 10)
                edcDropDown.Width = tmRCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 5
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(ilBoxNo).fBoxX, tmRCtrls(ilBoxNo).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcDays.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imRRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcDays.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcDays.Move edcDropDown.Left, edcDropDown.Top - lbcDays.height
                End If
                imChgMode = True
                slStr = Trim$(smRSave(10, imRRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcDays
                    If gLastFound(lbcDays) >= 0 Then
                        lbcDays.ListIndex = gLastFound(lbcDays)
                    Else
                        lbcDays.ListIndex = 0
                    End If
                Else
                    lbcDays.ListIndex = 0
                End If
                If lbcDays.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcDays.List(lbcDays.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RSTARTTIMEINDEX 'Start time
                edcDropDown.Width = tmRCtrls(RSTARTTIMEINDEX).fBoxW - cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(RSTARTTIMEINDEX).fBoxX, tmRCtrls(RSTARTTIMEINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcTme.height < cmcDone.Top Then
                    plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.height
                End If
                If smRSave(8, imRRowNo) <> "" Then
                    edcDropDown.Text = smRSave(8, imRRowNo)
                Else
                    edcDropDown.Text = "12M"
                End If
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                If smRSave(8, imRRowNo) = "" Then
                    plcTme.Visible = True
                End If
                edcDropDown.SetFocus
            Case RENDTIMEINDEX 'Start time
                edcDropDown.Width = tmRCtrls(RENDTIMEINDEX).fBoxW - cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(RENDTIMEINDEX).fBoxX, tmRCtrls(RENDTIMEINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcTme.height < cmcDone.Top Then
                    plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.height
                End If
                If smRSave(9, imRRowNo) <> "" Then
                    edcDropDown.Text = smRSave(9, imRRowNo)
                Else
                    If (Trim$(smRSave(3, imRRowNo)) = "[All]") Or (Trim$(smRSave(3, imRRowNo)) = "") Or (smFromLog <> "Y") Then
                        edcDropDown.Text = "12M"
                    Else
                        edcDropDown.Text = smRSave(8, imRRowNo) '"12M" jim request to set EndTime = StartTime
                    End If
                End If
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                If smRSave(9, imRRowNo) = "" Then
                    plcTme.Visible = True
                End If
                edcDropDown.SetFocus
        End Select
    Else
        Select Case ilBoxNo 'Branch on box type (control)
            Case RADVTINDEX 'Advertiser
                mAdvtPop
                If imTerminate Then
                    Exit Sub
                End If
                lbcRAdvt.height = gListBoxHeight(lbcRAdvt.ListCount, 10)
                edcDropDown.Width = tmRCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 40
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(ilBoxNo).fBoxX, tmRCtrls(ilBoxNo).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcRAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imRRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcRAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcRAdvt.Move edcDropDown.Left, edcDropDown.Top - lbcRAdvt.height
                End If
                imChgMode = True
                slStr = Trim$(smRSave(1, imRRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcRAdvt
                    If gLastFound(lbcRAdvt) >= 0 Then
                        lbcRAdvt.ListIndex = gLastFound(lbcRAdvt)
                    Else
                        lbcRAdvt.ListIndex = 0
                    End If
                Else
                    lbcRAdvt.ListIndex = 0
                End If
                If lbcRAdvt.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcRAdvt.List(lbcRAdvt.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RCARTINDEX
                mCartPop imRRowNo
                If imTerminate Then
                    Exit Sub
                End If
                lbcCart.height = gListBoxHeight(lbcCart.ListCount, 10)
                edcDropDown.Width = tmRCtrls(ilBoxNo).fBoxW + tmRCtrls(RSHORTTITLEINDEX).fBoxW + tmRCtrls(RPPINDEX).fBoxW + tmRCtrls(RPPINDEX + 1).fBoxW + tmRCtrls(RSTARTDATEINDEX).fBoxW + tmRCtrls(RENDDATEINDEX).fBoxW + 7 * tmRCtrls(RDAYINDEX).fBoxW - cmcDropDown.Width
                edcDropDown.MaxLength = 0
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(ilBoxNo).fBoxX, tmRCtrls(ilBoxNo).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcCart.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imRRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcCart.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcCart.Move edcDropDown.Left, edcDropDown.Top - lbcCart.height
                End If
                imChgMode = True
                slStr = Trim$(smRSave(2, imRRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcCart
                    If gLastFound(lbcCart) >= 0 Then
                        lbcCart.ListIndex = gLastFound(lbcCart)
                    Else
                        lbcCart.ListIndex = 0
                    End If
                Else
                    lbcCart.ListIndex = 0
                End If
                If lbcCart.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcCart.List(lbcCart.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RSHORTTITLEINDEX
            '    mShtTitlePop imRRowNo
            '    If imTerminate Then
            '        Exit Sub
            '    End If
            '    lbcShtTitle.Height = gListBoxHeight(lbcShtTitle.ListCount, 10)
            '    edcDropDown.Width = tmRCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
            '    edcDropDown.MaxLength = 15
            '    gMoveTableCtrl pbcReplacement, edcDropDown, tmRCtrls(ilBoxNo).fBoxX, tmRCtrls(ilBoxNo).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
             '   cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            '    lbcShtTitle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            '    If imRRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
            '        lbcShtTitle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            '    Else
            '        lbcShtTitle.Move edcDropDown.Left, edcDropDown.Top - lbcShtTitle.Height
            '    End If
            '    imChgMode = True
            '    slStr = Trim$(smRSave(3, imRRowNo))
            '    If slStr <> "" Then
            '        gFindMatch slStr, 1, lbcShtTitle
            '        If gLastFound(lbcShtTitle) > 0 Then
            '            lbcShtTitle.ListIndex = gLastFound(lbcShtTitle)
            '        Else
            '            lbcShtTitle.ListIndex = 0
            '        End If
            '    Else
            '        lbcShtTitle.ListIndex = 0
            '    End If
            '    If lbcShtTitle.ListIndex < 0 Then
            '        edcDropDown.Text = ""
            '    Else
            '        edcDropDown.Text = lbcShtTitle.List(lbcShtTitle.ListIndex)
            '    End If
            '    imChgMode = False
            '    edcDropDown.SelStart = 0
            '    edcDropDown.SelLength = Len(edcDropDown.Text)
            '    edcDropDown.Visible = True
            '    cmcDropDown.Visible = True
            '    edcDropDown.SetFocus
            Case RPPINDEX, RPPINDEX + 1
                mCompPop
                If imTerminate Then
                    Exit Sub
                End If
                lbcComp(ilBoxNo - RPPINDEX).height = gListBoxHeight(lbcComp(ilBoxNo - RPPINDEX).ListCount, 10)
                edcDropDown.Width = tmRCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 20
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(ilBoxNo).fBoxX, tmRCtrls(ilBoxNo).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcComp(ilBoxNo - RPPINDEX).Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imRRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcComp(ilBoxNo - RPPINDEX).Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcComp(ilBoxNo - RPPINDEX).Move edcDropDown.Left, edcDropDown.Top - lbcComp(ilBoxNo - RPPINDEX).height
                End If
                imChgMode = True
                slStr = Trim$(smRSave(4 + ilBoxNo - RPPINDEX, imRRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcComp(ilBoxNo - RPPINDEX)
                    If gLastFound(lbcComp(ilBoxNo - RPPINDEX)) >= 0 Then
                        lbcComp(ilBoxNo - RPPINDEX).ListIndex = gLastFound(lbcComp(ilBoxNo - RPPINDEX))
                    Else
                        lbcComp(ilBoxNo - RPPINDEX).ListIndex = 0
                    End If
                Else
                    lbcComp(ilBoxNo - RPPINDEX).ListIndex = 0
                End If
                If lbcComp(ilBoxNo - RPPINDEX).ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcComp(ilBoxNo - RPPINDEX).List(lbcComp(ilBoxNo - RPPINDEX).ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RSTARTDATEINDEX 'Start date
                edcDropDown.Width = tmRCtrls(RSTARTDATEINDEX).fBoxW + cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(RSTARTDATEINDEX).fBoxX, tmRCtrls(RSTARTDATEINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcCalendar.height < cmcDone.Top Then
                    plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.height
                End If
                If smRSave(6, imRRowNo) = "" Then
                    slStr = Trim$(smRSave(2, imRRowNo))
                    If (slStr <> "") And (slStr <> "[None]") Then
                        gFindMatch slStr, 0, lbcCart
                        If gLastFound(lbcCart) >= 0 Then
                            lbcCart.ListIndex = gLastFound(lbcCart)
                            slNameCode = tmCartCode(gLastFound(lbcCart)).sKey    'Traffic!lbcRAdvt.List(gLastFound(lbcRAdvt) - 1)
                            ilRet = gParseItem(slNameCode, 5, "\", slStr)
                            If slStr = "" Then
                                slStr = gObtainMondayFromToday()
                            End If
                        Else
                            slStr = gObtainMondayFromToday()
                        End If
                    Else
                        slStr = gObtainMondayFromToday()
                    End If
                Else
                    slStr = smRSave(6, imRRowNo)
                End If
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
                edcDropDown.Text = slStr
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                If smRSave(6, imRRowNo) = "" Then
                    pbcCalendar.Visible = True
                End If
                edcDropDown.SetFocus
            Case RENDDATEINDEX 'Start date
                edcDropDown.Width = tmRCtrls(RENDDATEINDEX).fBoxW + cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(RENDDATEINDEX).fBoxX, tmRCtrls(RENDDATEINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcCalendar.height < cmcDone.Top Then
                    plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top + edcDropDown.height
                Else
                    plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top - plcCalendar.height
                End If
                If smRSave(7, imRRowNo) <> "" Then
                    slStr = smRSave(7, imRRowNo)
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint
                Else
                    slStr = Trim$(smRSave(2, imRRowNo))
                    If (slStr <> "") And (slStr <> "[None]") Then
                        gFindMatch slStr, 0, lbcCart
                        If gLastFound(lbcCart) >= 0 Then
                            lbcCart.ListIndex = gLastFound(lbcCart)
                            slNameCode = tmCartCode(gLastFound(lbcCart)).sKey    'Traffic!lbcRAdvt.List(gLastFound(lbcRAdvt) - 1)
                            ilRet = gParseItem(slNameCode, 6, "\", slStr)
                            If slStr = "" Then
                                slStr = Format$(gDateValue(smRSave(6, imRRowNo)) + 6, "m/d/yy")
                            End If
                        Else
                            slStr = Format$(gDateValue(smRSave(6, imRRowNo)) + 6, "m/d/yy")
                        End If
                    Else
                        slStr = gObtainMondayFromToday()
                        slStr = Format$(gDateValue(slStr) + 6, "m/d/yy")
                    End If
                End If
                edcDropDown.Text = slStr
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RDAYINDEX To RDAYINDEX + 6 'Day index
                gMoveTableCtrl pbcReplacement(imSRIndex), pbcDays, tmRCtrls(ilBoxNo).fBoxX, tmRCtrls(ilBoxNo).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                If imRSave(ilBoxNo - RDAYINDEX + 1, imRRowNo) = 1 Then
                    ckcDay.Value = vbChecked
                ElseIf imRSave(ilBoxNo - RDAYINDEX + 1, imRRowNo) = -1 Then
                    If (Trim$(smRSave(6, imRRowNo)) = "") Or (Trim$(smRSave(7, imRRowNo)) = "") Then
                        ckcDay.Value = vbChecked
                    Else
                        ilSDay = gWeekDayStr(smRSave(6, imRRowNo))
                        ilEDay = gWeekDayStr(smRSave(7, imRRowNo))
                        If ilSDay <= ilEDay Then
                            If ((ilBoxNo - RDAYINDEX) >= ilSDay) And ((ilBoxNo - RDAYINDEX) <= ilEDay) Then
                                ckcDay.Value = vbChecked
                            Else
                                ckcDay.Value = vbUnchecked
                            End If
                        Else
                            If ((ilBoxNo - RDAYINDEX) >= ilSDay) And ((ilBoxNo - RDAYINDEX) <= 6) Then
                                ckcDay.Value = vbChecked
                            Else
                                If ((ilBoxNo - RDAYINDEX) >= 0) And ((ilBoxNo - RDAYINDEX) <= ilEDay) Then
                                    ckcDay.Value = vbChecked
                                Else
                                    ckcDay.Value = vbUnchecked
                                End If
                            End If
                        End If
                    End If
                Else
                    ckcDay.Value = vbUnchecked
                End If
                pbcDays.Visible = True  'Set visibility
                ckcDay.SetFocus
            Case RSTARTTIMEINDEX 'Start time
                edcDropDown.Width = tmRCtrls(RSTARTTIMEINDEX).fBoxW - cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(RSTARTTIMEINDEX).fBoxX, tmRCtrls(RSTARTTIMEINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcTme.height < cmcDone.Top Then
                    plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.height
                End If
                If smRSave(8, imRRowNo) <> "" Then
                    edcDropDown.Text = smRSave(8, imRRowNo)
                Else
                    edcDropDown.Text = "12M"
                End If
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                If smRSave(8, imRRowNo) = "" Then
                    plcTme.Visible = True
                End If
                edcDropDown.SetFocus
            Case RENDTIMEINDEX 'Start time
                edcDropDown.Width = tmRCtrls(RENDTIMEINDEX).fBoxW - cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcReplacement(imSRIndex), edcDropDown, tmRCtrls(RENDTIMEINDEX).fBoxX, tmRCtrls(RENDTIMEINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcTme.height < cmcDone.Top Then
                    plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.height
                End If
                If smRSave(9, imRRowNo) <> "" Then
                    edcDropDown.Text = smRSave(9, imRRowNo)
                Else
                    If (Trim$(smRSave(3, imRRowNo)) = "[All]") Or (Trim$(smRSave(3, imRRowNo)) = "") Or (smFromLog <> "Y") Then
                        edcDropDown.Text = "12M"
                    Else
                        edcDropDown.Text = smRSave(8, imRRowNo) '"12M" jim request to set EndTime = StartTime
                    End If
                End If
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                If smRSave(9, imRRowNo) = "" Then
                    plcTme.Visible = True
                End If
                edcDropDown.SetFocus
        End Select
    End If
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mRSetFocus(ilBoxNo As Integer)
'
'   mRSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBRCtrls) Or (ilBoxNo > UBound(tmRCtrls)) Then
        Exit Sub
    End If

    If (imRRowNo < vbcSR.Value) Or (imRRowNo >= vbcSR.Value + vbcSR.LargeChange + 1) Then
        'mSetShow ilBoxNo
        pbcArrow.Visible = False
        lacSFrame(imSRIndex).Visible = False
        Exit Sub
    End If
    lacRFrame(imSRIndex).Move 0, tmRCtrls(RADVTINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15) - 30
    lacRFrame(imSRIndex).Visible = True
    pbcArrow.Move pbcArrow.Left, plcBlackout.Top + tmRCtrls(RADVTINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    If smFromLog = "Y" Then
        Select Case ilBoxNo 'Branch on box type (control)
            Case RADVTINDEX 'Advertiser
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RCNTRINDEX 'Contract
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RCARTINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RPPINDEX, RPPINDEX + 1
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RVEHINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RSTARTDATEINDEX 'Start date
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RENDDATEINDEX 'Start date
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RDAYINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RSTARTTIMEINDEX 'Start time
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RENDTIMEINDEX 'Start time
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
        End Select
    Else
        Select Case ilBoxNo 'Branch on box type (control)
            Case RADVTINDEX 'Advertiser
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RCARTINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RSHORTTITLEINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RPPINDEX, RPPINDEX + 1
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RSTARTDATEINDEX 'Start date
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RENDDATEINDEX 'Start date
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RDAYINDEX To RDAYINDEX + 6 'Day index
                pbcDays.Visible = True  'Set visibility
                ckcDay.SetFocus
            Case RSTARTTIMEINDEX 'Start time
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case RENDTIMEINDEX 'Start time
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
        End Select
    End If
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRSetShow                       *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mRSetShow(ilBoxNo As Integer)
'
'   mRSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    pbcArrow.Visible = False
    lacSFrame(imSRIndex).Visible = False
    lacRFrame(imSRIndex).Visible = False
    If (ilBoxNo < imLBRCtrls) Or (ilBoxNo > UBound(tmRCtrls)) Then
        Exit Sub
    End If

    If smFromLog = "Y" Then
        Select Case ilBoxNo 'Branch on box type (control)
            Case RADVTINDEX
                lbcRAdvt.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                If Trim$(smRSave(1, imRRowNo)) <> slStr Then
                    imRChg = True
                    smRSave(1, imRRowNo) = slStr
                    lmRSave(1, imRRowNo) = 0
                    lmRSave(2, imRRowNo) = 0
                    smRSave(2, imRRowNo) = ""   'Clear Cart
                    smRShow(RCARTINDEX, imRRowNo) = ""
                    lmRSave(3, imRRowNo) = 0
                    smRSave(11, imRRowNo) = ""
                    smRShow(RCNTRINDEX, imRRowNo) = ""
                End If
            Case RCNTRINDEX
                lbcRCntr.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                '9/5/06: Disallow [all] as mRTestSaveFields requires contract to be defined.
                'If ((lbcRCntr.ListIndex > 0) And (tgSpf.sUseProdSptScr <> "P")) Or ((lbcRCntr.ListIndex >= 0) And (tgSpf.sUseProdSptScr = "P")) Then
                If ((lbcRCntr.ListIndex >= 0) And (tgSpf.sUseProdSptScr <> "P")) Or ((lbcRCntr.ListIndex >= 0) And (tgSpf.sUseProdSptScr = "P")) Then
                    If tgSpf.sUseProdSptScr = "P" Then
                        slNameCode = tgPlanner(lbcRCntr.ListIndex).sKey 'Traffic!lbcAdvertiser.List(gLastFound(lbcSAdvt) - 1)
                    Else
                        '9/5/06: Disallow [all] as mRTestSaveFields requires contract to be defined.
                        'slNameCode = tgPlanner(lbcRCntr.ListIndex - 1).sKey 'Traffic!lbcAdvertiser.List(gLastFound(lbcSAdvt) - 1)
                        slNameCode = tgPlanner(lbcRCntr.ListIndex).sKey 'Traffic!lbcAdvertiser.List(gLastFound(lbcSAdvt) - 1)
                    End If
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) <> lmRSave(3, imRRowNo) Then
                        imRChg = True
                        lmRSave(3, imRRowNo) = Val(slCode)
                        ilPos = InStr(1, slNameCode, " ", 1)
                        If ilPos > 0 Then
                            smRSave(11, imRRowNo) = Left$(slNameCode, ilPos - 1)
                        Else
                            ilRet = gParseItem(slNameCode, 1, "\", smRSave(11, imRRowNo))
                        End If
                        mGetCntrPP imRRowNo
                    End If
                Else
                    If lmRSave(3, imRRowNo) <> 0 Then
                        imRChg = True
                        lmRSave(3, imRRowNo) = 0
                        smRSave(11, imRRowNo) = ""
                        mGetCntrPP imRRowNo
                    End If
                End If
            Case RCARTINDEX
                lbcCart.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                If Trim$(smRSave(2, imRRowNo)) <> slStr Then
                    imRChg = True
                    smRSave(2, imRRowNo) = slStr
                    If (lbcCart.ListIndex >= 0) And (slStr <> "[None]") Then
                        '9/5/06: Disallow [None] if from Log and SplitFill
                        If smSplitFill <> "Y" Then
                            slNameCode = tmCartCode(lbcCart.ListIndex - 1).sKey 'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                        Else
                            slNameCode = tmCartCode(lbcCart.ListIndex).sKey 'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                        End If
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        lmRSave(1, imRRowNo) = Val(slCode)
                        ilRet = gParseItem(slNameCode, 3, "\", slName)
                        ilRet = gParseItem(slNameCode, 4, "\", slCode)
                        lmRSave(2, imRRowNo) = Val(slCode)
                    Else
                        lmRSave(1, imRRowNo) = 0
                        lmRSave(2, imRRowNo) = 0
                    End If
                End If
            Case RPPINDEX, RPPINDEX + 1
                lbcComp(ilBoxNo - RPPINDEX).Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                If Trim$(smRSave(4 + ilBoxNo - RPPINDEX, imRRowNo)) <> slStr Then
                    If (slStr <> "[None]") Or (smRSave(ilBoxNo, imRRowNo) <> "") Then
                        imRChg = True
                        smRSave(4 + ilBoxNo - RPPINDEX, imRRowNo) = slStr
                    End If
                End If
            Case RVEHINDEX
                lbcVehicle.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                If Trim$(smRSave(3, imRRowNo)) <> slStr Then
                    imRChg = True
                    smRSave(3, imRRowNo) = slStr
                End If
            Case RSTARTDATEINDEX 'Start date
                plcCalendar.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gValidDate(slStr) Then
                    gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                    smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                    If Trim$(smRSave(6, imRRowNo)) <> slStr Then
                        imRChg = True
                        smRSave(6, imRRowNo) = slStr
                    End If
                Else
                    Beep
                End If
            Case RENDDATEINDEX 'Start date
                plcCalendar.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If StrComp(slStr, "TFN", 1) <> 0 Then
                    If gValidDate(slStr) Then
                        If Trim$(smRSave(7, imRRowNo)) <> slStr Then
                            imRChg = True
                            smRSave(7, imRRowNo) = slStr
                        End If
                        slStr = gFormatDate(slStr)
                        gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                        smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                Else
                    If Trim$(smRSave(7, imRRowNo)) <> "" Then
                        imRChg = True
                    End If
                    smRSave(7, imRRowNo) = ""
                    slStr = "TFN"
                    gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                    smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                End If
            Case RDAYINDEX 'Day index
                lbcDays.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                If Trim$(smRSave(10, imRRowNo)) <> slStr Then
                    imRChg = True
                    smRSave(10, imRRowNo) = slStr
                End If
            Case RSTARTTIMEINDEX 'Start time
                cmcDropDown.Visible = False
                plcTme.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If gValidTime(slStr) Then
                        If Trim$(smRSave(8, imRRowNo)) <> slStr Then
                            imRChg = True
                            smRSave(8, imRRowNo) = slStr
                        End If
                        slStr = gFormatTime(slStr, "A", "1")
                        gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                        smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                End If
            Case RENDTIMEINDEX 'Start time
                cmcDropDown.Visible = False
                plcTme.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If gValidTime(slStr) Then
                        If Trim$(smRSave(9, imRRowNo)) <> slStr Then
                            imRChg = True
                            smRSave(9, imRRowNo) = slStr
                        End If
                        slStr = gFormatTime(slStr, "A", "1")
                        gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                        smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                End If
        End Select
    Else
        Select Case ilBoxNo 'Branch on box type (control)
            Case RADVTINDEX
                lbcRAdvt.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                If Trim$(smRSave(1, imRRowNo)) <> slStr Then
                    imRChg = True
                    smRSave(1, imRRowNo) = slStr
                    smRSave(2, imRRowNo) = ""   'Clear Cart
                End If
            Case RCARTINDEX
                lbcCart.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                If Trim$(smRSave(2, imRRowNo)) <> slStr Then
                    imRChg = True
                    smRSave(2, imRRowNo) = slStr
                    If (lbcCart.ListIndex >= 0) And (slStr <> "[None]") Then
                        slNameCode = tmCartCode(lbcCart.ListIndex).sKey  'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        lmRSave(1, imRRowNo) = Val(slCode)
                        ilRet = gParseItem(slNameCode, 3, "\", slName)
                        'If smFromLog <> "Y" Then
                            smRSave(3, imRRowNo) = slName
                            gSetShow pbcReplacement(imSRIndex), slName, tmRCtrls(RSHORTTITLEINDEX)
                            smRShow(RSHORTTITLEINDEX, imRRowNo) = tmRCtrls(RSHORTTITLEINDEX).sShow
                        'End If
                        ilRet = gParseItem(slNameCode, 4, "\", slCode)
                        lmRSave(2, imRRowNo) = Val(slCode)
                    Else
                        lmRSave(1, imRRowNo) = 0
                        lmRSave(2, imRRowNo) = 0
                    End If
                End If
            Case RSHORTTITLEINDEX
            Case RPPINDEX, RPPINDEX + 1
                lbcComp(ilBoxNo - RPPINDEX).Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                If Trim$(smRSave(4 + ilBoxNo - RPPINDEX, imRRowNo)) <> slStr Then
                    If (slStr <> "[None]") Or (smRSave(ilBoxNo, imRRowNo) <> "") Then
                        imRChg = True
                        smRSave(4 + ilBoxNo - RPPINDEX, imRRowNo) = slStr
                    End If
                End If
            Case RSTARTDATEINDEX 'Start date
                plcCalendar.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gValidDate(slStr) Then
                    gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                    smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                    If Trim$(smRSave(6, imRRowNo)) <> slStr Then
                        imRChg = True
                        smRSave(6, imRRowNo) = slStr
                    End If
                Else
                    Beep
                End If
            Case RENDDATEINDEX 'Start date
                plcCalendar.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If StrComp(slStr, "TFN", 1) <> 0 Then
                    If gValidDate(slStr) Then
                        If Trim$(smRSave(7, imRRowNo)) <> slStr Then
                            imRChg = True
                            smRSave(7, imRRowNo) = slStr
                        End If
                        slStr = gFormatDate(slStr)
                        gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                        smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                Else
                    If Trim$(smRSave(7, imRRowNo)) <> "" Then
                        imRChg = True
                    End If
                    smRSave(7, imRRowNo) = ""
                    slStr = "TFN"
                    gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                    smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                End If
            Case RDAYINDEX To RDAYINDEX + 6 'Day index
                pbcDays.Visible = False  'Set visibility
                If ckcDay.Value = vbChecked Then
                    If imRSave(ilBoxNo - RDAYINDEX + 1, imRRowNo) <> 1 Then
                        imRChg = True
                    End If
                    slStr = "4"
                    imRSave(ilBoxNo - RDAYINDEX + 1, imRRowNo) = 1
                Else
                    If imRSave(ilBoxNo - RDAYINDEX + 1, imRRowNo) <> 0 Then
                        imRChg = True
                    End If
                    slStr = "  "
                    imRSave(ilBoxNo - RDAYINDEX + 1, imRRowNo) = 0
                End If
                gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
            Case RSTARTTIMEINDEX 'Start time
                cmcDropDown.Visible = False
                plcTme.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If gValidTime(slStr) Then
                        If Trim$(smRSave(8, imRRowNo)) <> slStr Then
                            imRChg = True
                            smRSave(8, imRRowNo) = slStr
                        End If
                        slStr = gFormatTime(slStr, "A", "1")
                        gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                        smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                End If
            Case RENDTIMEINDEX 'Start time
                cmcDropDown.Visible = False
                plcTme.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If gValidTime(slStr) Then
                        If Trim$(smRSave(9, imRRowNo)) <> slStr Then
                            imRChg = True
                            smRSave(9, imRRowNo) = slStr
                        End If
                        slStr = gFormatTime(slStr, "A", "1")
                        gSetShow pbcReplacement(imSRIndex), slStr, tmRCtrls(ilBoxNo)
                        smRShow(ilBoxNo, imRRowNo) = tmRCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                End If
        End Select
    End If
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRTestSaveFields                *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mRTestSaveFields(ilRowNo As Integer, ilMsg As Integer) As Integer
'
'   iRet = mRTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilOneDay As Integer
    Dim ilDay As Integer
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilRet As Integer

    If Trim$(smRSave(1, ilRowNo)) = "" Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Advertiser must be specified", vbOKOnly + vbExclamation, "In Replace")
        End If
        imRBoxNo = RADVTINDEX
        mRTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smRSave(2, ilRowNo)) = "" Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Cart must be specified for" & Trim$(smRSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Replace")
        End If
        imRBoxNo = RCARTINDEX
        mRTestSaveFields = NO
        Exit Function
    End If
    If (Trim$(smRSave(2, ilRowNo)) = "[None]") And (imSRIndex = 0) Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Cart must be specified for " & Trim$(smRSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Replace")
        End If
        imRBoxNo = RCARTINDEX
        mRTestSaveFields = NO
        Exit Function
    End If
    'Use Short Title, need contract to obtain short title when generating Logs
    If tgSpf.sUseProdSptScr = "P" Then
        If (Trim$(smRSave(11, ilRowNo)) = "") And (imSRIndex = 1) Then
            Beep
            If ilMsg Then
                ilRes = MsgBox("Contract must be specified for " & Trim$(smRSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Replace")
            End If
            imRBoxNo = RCNTRINDEX
            mRTestSaveFields = NO
            Exit Function
        End If
        If (Trim$(smRSave(2, ilRowNo)) = "[None]") And (imSRIndex = 1) Then
            Beep
            If ilMsg Then
                ilRes = MsgBox("Cart must be specified for " & Trim$(smRSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Replace")
            End If
            imRBoxNo = RCARTINDEX
            mRTestSaveFields = NO
            Exit Function
        End If
    End If
    'Use Advt/Prod
    If tgSpf.sUseProdSptScr <> "P" Then
'        If (Trim$(smRSave(11, ilRowNo)) = "") And (Trim$(smRSave(2, ilRowNo)) = "[None]") And (imSRIndex = 1) Then
'            Beep
'            If ilMsg Then
'                ilRes = MsgBox("Contract or Cart must be specified for " & Trim$(smRSave(1, ilRowNo)), vbOkOnly + vbExclamation, "In Replace")
'            End If
'            imRBoxNo = RCARTINDEX
'            mRTestSaveFields = NO
'            Exit Function
'        End If
        If (Trim$(smRSave(11, ilRowNo)) = "") And (imSRIndex = 1) Then
            Beep
            If ilMsg Then
                ilRes = MsgBox("Contract must be specified for " & Trim$(smRSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Replace")
            End If
            imRBoxNo = RCNTRINDEX
            mRTestSaveFields = NO
            Exit Function
        End If
    End If
'    If (Trim$(smRSave(11, ilRowNo)) = "") And (imSRIndex = 1) And (tgSpf.sUseProdSptScr = "P") Then
'        'Test if cart has short title associated with it
'        tmCifSrchKey.lCode = lmRSave(1, ilRowNo)
'        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            tmCpfSrchKey.lCode = tmCif.lCpfCode
'            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'            If ilRet = BTRV_ERR_NONE Then
'                If Trim$(tmCpf.sName) = "" Then
'                    Beep
'                    If ilMsg Then
'                        ilRes = MsgBox("Cart specified does not have Short Title, enter Contract for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOkOnly + vbExclamation, "In Replace")
'                    End If
'                    imRBoxNo = RCARTINDEX
'                    mRTestSaveFields = NO
'                    Exit Function
'                End If
'            Else
'            End If
'        Else
'            Beep
'            If ilMsg Then
'                ilRes = MsgBox("Contract or Cart must be specified for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOkOnly + vbExclamation, "In Replace")
'            End If
'            imRBoxNo = RCARTINDEX
'            mRTestSaveFields = NO
'            Exit Function
'        End If
'    End If
    'If Trim$(smRSave(3, ilRowNo)) = "" Then
    '    Beep
    '    ilRes = MsgBox("Short Title must be specified", vbOkOnly + vbExclamation, "Incomplete")
    '    imRBoxNo = RSHORTTITLEINDEX
    '    mRTestSaveFields = NO
    '    Exit Function
    'End If
    'If Trim$(smRSave(4, ilRowNo)) = "" Then
    '    Beep
    '    ilRes = MsgBox("Product Protection must be specified", vbOkOnly + vbExclamation, "Incomplete")
    '    imRBoxNo = RPPINDEX
    '    mRTestSaveFields = NO
    '    Exit Function
    'End If
    'If Trim$(smRSave(5, ilRowNo)) = "" Then
    '    Beep
    '    ilRes = MsgBox("Product Protection must be specified", vbOkOnly + vbExclamation, "Incomplete")
    '    imRBoxNo = RPPINDEX+1
    '    mRTestSaveFields = NO
    '    Exit Function
    'End If
    If Trim$(smRSave(6, ilRowNo)) = "" Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Start Date must be specified for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOKOnly + vbExclamation, "In Replace")
        End If
        imRBoxNo = RSTARTDATEINDEX
        mRTestSaveFields = NO
        Exit Function
    Else
        If Not gValidDate(smRSave(6, ilRowNo)) Then
            Beep
            If ilMsg Then
                ilRes = MsgBox("Start Date must be valid for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOKOnly + vbExclamation, "In Replacee")
            End If
            imRBoxNo = RSTARTDATEINDEX
            mRTestSaveFields = NO
            Exit Function
        End If
    End If
    If Trim$(smRSave(7, ilRowNo)) = "" Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("End Date must be specified for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOKOnly + vbExclamation, "In Replace")
        End If
        imRBoxNo = RENDDATEINDEX
        mRTestSaveFields = NO
        Exit Function
    Else
        If Not gValidDate(smRSave(7, ilRowNo)) Then
            Beep
            If ilMsg Then
                ilRes = MsgBox("End Date must be valid for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOKOnly + vbExclamation, "In Replace")
            End If
            imRBoxNo = RENDDATEINDEX
            mRTestSaveFields = NO
            Exit Function
        End If
    End If
    'Test that Dates are within Rotation Dates
    If (ilMsg) And (lmRSave(1, ilRowNo) > 0) Then
        tmCifSrchKey.lCode = lmRSave(1, ilRowNo)
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        gUnpackDateLong tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), llSDate
        gUnpackDateLong tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), llEDate
        If (gDateValue(smRSave(6, ilRowNo)) < llSDate) Or (gDateValue(smRSave(7, ilRowNo)) > llEDate) Then
            Beep
            ilRes = MsgBox("Cart dates not covering specified dates for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOKOnly + vbExclamation, "In Replace")
            imRBoxNo = RSTARTDATEINDEX
            mRTestSaveFields = NO
            Exit Function
        End If
    End If
    ilOneDay = False
    If smFromLog = "Y" Then
        If smRSave(10, ilRowNo) <> "" Then
            ilOneDay = True
        End If
    Else
        For ilDay = 0 To 6 Step 1
            If imRSave(ilDay + 1, ilRowNo) > 0 Then
                ilOneDay = True
                Exit For
            End If
        Next ilDay
    End If
    If Not ilOneDay Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("One Day must be specified for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOKOnly + vbExclamation, "In Replace")
        End If
        imRBoxNo = RDAYINDEX
        mRTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smRSave(8, ilRowNo)) = "" Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Start Time must be specified for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOKOnly + vbExclamation, "In Replace")
        End If
        imRBoxNo = RSTARTTIMEINDEX
        mRTestSaveFields = NO
        Exit Function
    Else
        If Not gValidTime(smRSave(8, ilRowNo)) Then
            Beep
            If ilMsg Then
                ilRes = MsgBox("Start Time must be valid for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOKOnly + vbExclamation, "In Replace")
            End If
            imRBoxNo = RSTARTTIMEINDEX
            mRTestSaveFields = NO
            Exit Function
        End If
    End If
    If Trim$(smRSave(9, ilRowNo)) = "" Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Start Time must be specified for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOKOnly + vbExclamation, "In Replace")
        End If
        imRBoxNo = RSTARTTIMEINDEX
        mRTestSaveFields = NO
        Exit Function
    Else
        If Not gValidTime(smRSave(9, ilRowNo)) Then
            Beep
            If ilMsg Then
                ilRes = MsgBox("Start Time must be valid for " & Trim$(smRSave(1, ilRowNo)) & ", " & smRSave(2, ilRowNo), vbOKOnly + vbExclamation, "In Replace")
            End If
            imRBoxNo = RENDTIMEINDEX
            mRTestSaveFields = NO
            Exit Function
        End If
    End If
    mRTestSaveFields = YES
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
    Dim ilRet As Integer
    Dim slMsg As String
    Dim ilBof As Integer
    Dim ilRowNo As Integer
    Dim tlBof As BOF
    Dim tlBof1 As MOVEREC
    Dim tlBof2 As MOVEREC
    mSSetShow imSBoxNo
    imSBoxNo = -1
    mRSetShow imRBoxNo
    imRBoxNo = -1
    For ilRowNo = 1 To UBound(smSSave, 2) - 1 Step 1
        If mSTestSaveFields(ilRowNo, True) = NO Then
            mSaveRec = False
            imSRowNo = ilRowNo
            Exit Function
        End If
    Next ilRowNo
    For ilRowNo = 1 To UBound(smRSave, 2) - 1 Step 1
        If mRTestSaveFields(ilRowNo, True) = NO Then
            mSaveRec = False
            imRRowNo = ilRowNo
            Exit Function
        End If
    Next ilRowNo
    Screen.MousePointer = vbHourglass  'Wait
    mMoveCtrlToRec
    Screen.MousePointer = vbHourglass  'Wait
    If imTerminate Then
        Screen.MousePointer = vbDefault    'Default
        mSaveRec = False
        Exit Function
    End If
    For ilBof = LBound(tgSBofRec) To UBound(tgSBofRec) - 1 Step 1
        Do  'Loop until record updated or added
            If (tgSBofRec(ilBof).iStatus = 0) Then  'New selected
                tgSBofRec(ilBof).tBof.lCode = 0
                tgSBofRec(ilBof).tBof.sType = "S"
                If smFromLog <> "Y" Then
                    tgSBofRec(ilBof).tBof.sSource = "N"
                Else
                    tgSBofRec(ilBof).tBof.sSource = "L"
                End If
                tgSBofRec(ilBof).tBof.iUrfCode = tgUrf(0).iCode
                ilRet = btrInsert(hmBof, tgSBofRec(ilBof).tBof, imBofRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert: Blackout)"
                tgSBofRec(ilBof).iStatus = 1
                ilRet = btrGetPosition(hmBof, tgSBofRec(ilBof).lRecPos)
            Else 'Old record-Update
                slMsg = "mSaveRec (btrGetDirect: Blackout)"
                ilRet = btrGetDirect(hmBof, tlBof, imBofRecLen, tgSBofRec(ilBof).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, Blackout
                On Error GoTo 0
                'tmRec = tlBof
                'ilRet = gGetByKeyForUpdate("BOF", hmBof, tmRec)
                'tlBof = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
                '    imTerminate = True
                '    mSaveRec = False
                '    Exit Function
                'End If
                LSet tlBof1 = tlBof
                LSet tlBof2 = tgSBofRec(ilBof).tBof
                If StrComp(tlBof1.sChar, tlBof2.sChar, 0) <> 0 Then
                    If smFromLog <> "Y" Then
                        tgSBofRec(ilBof).tBof.sSource = "N"
                    Else
                        tgSBofRec(ilBof).tBof.sSource = "L"
                    End If
                    tgSBofRec(ilBof).tBof.iUrfCode = tgUrf(0).iCode
                    ilRet = btrUpdate(hmBof, tgSBofRec(ilBof).tBof, imBofRecLen)
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrUpdate: Blackout)"
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, Blackout
        On Error GoTo 0
    Next ilBof
    For ilBof = LBound(tgRBofRec) To UBound(tgRBofRec) - 1 Step 1
        Do  'Loop until record updated or added
            If (tgRBofRec(ilBof).iStatus = 0) Then  'New selected
                tgRBofRec(ilBof).tBof.lCode = 0
                tgRBofRec(ilBof).tBof.sType = "R"
                If smFromLog <> "Y" Then
                    tgRBofRec(ilBof).tBof.sSource = "N"
                Else
                    tgRBofRec(ilBof).tBof.sSource = "L"
                End If
                tgRBofRec(ilBof).tBof.iUrfCode = tgUrf(0).iCode
                ilRet = btrInsert(hmBof, tgRBofRec(ilBof).tBof, imBofRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert: Blackout)"
                tgRBofRec(ilBof).iStatus = 1
                ilRet = btrGetPosition(hmBof, tgRBofRec(ilBof).lRecPos)
            Else 'Old record-Update
                slMsg = "mSaveRec (btrGetDirect: Blackout)"
                ilRet = btrGetDirect(hmBof, tlBof, imBofRecLen, tgRBofRec(ilBof).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, Blackout
                On Error GoTo 0
                'tmRec = tlBof
                'ilRet = gGetByKeyForUpdate("BOF", hmBof, tmRec)
                'tlBof = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
                '    imTerminate = True
                '    mSaveRec = False
                '    Exit Function
                'End If
                LSet tlBof1 = tlBof
                LSet tlBof2 = tgRBofRec(ilBof).tBof
                If StrComp(tlBof1.sChar, tlBof2.sChar, 0) <> 0 Then
                    If smFromLog <> "Y" Then
                        tgRBofRec(ilBof).tBof.sSource = "N"
                    Else
                        tgRBofRec(ilBof).tBof.sSource = "L"
                    End If
                    tgRBofRec(ilBof).tBof.iUrfCode = tgUrf(0).iCode
                    ilRet = btrUpdate(hmBof, tgRBofRec(ilBof).tBof, imBofRecLen)
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrUpdate: Blackout)"
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, Blackout
        On Error GoTo 0
    Next ilBof
    For ilBof = LBound(tgBofDel) To UBound(tgBofDel) - 1 Step 1
        If tgBofDel(ilBof).iStatus = 1 Then
            Do
                slMsg = "mSaveRec (btrGetDirect: Blackout)"
                ilRet = btrGetDirect(hmBof, tlBof, imBofRecLen, tgBofDel(ilBof).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later", vbOKOnly + vbExclamation, "Invoice")
                    Exit Function
                End If
                'tmRec = tlBof
                'ilRet = gGetByKeyForUpdate("BOF", hmBof, tmRec)
                'tlBof = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    Screen.MousePointer = vbDefault
                '    ilRet = MsgBox("Update Not Completed, Try Later", vbOkOnly + vbExclamation, "Invoice")
                '    Exit Function
                'End If
                ilRet = btrDelete(hmBof)
                slMsg = "mSaveRec (btrDelete: Sales History)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Update Not Completed, Try Later", vbOKOnly + vbExclamation, "Invoice")
                Exit Function
            End If
        End If
    Next ilBof
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
'*             Created:6/29/93       By:D. LeVine      *
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
'    If (igRcfChg Or imRpfChg Or imRgfChg) And ((imRowNo < UBound(tgRpfI) + 1) Or ((imRowNo = UBound(tgRpfI) + 1) And (imBoxNo = 0))) Then
    If imSChg Or imRChg Then
        If ilAsk Then
            slMess = "Save Changes and Additions"
            ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
            If ilRes = vbCancel Then
                mSaveRecChg = False
                If igView = 0 Then
                    pbcSuppression_Paint imSRIndex
                Else
                    pbcReplacement_Paint imSRIndex
                End If
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
'*      Procedure Name:mCntrPop                        *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Contract Population            *
'*                                                     *
'*******************************************************
Private Sub mSCntrPop(ilRowNo As Integer)
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilCurrent As Integer
    Dim ilShow As Integer
    Dim ilState As Integer
    Dim ilAAS As Integer
    Dim ilAASCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    'If Not imPropPopReqd Then
    '    Exit Sub
    'End If
    Screen.MousePointer = vbHourglass
    ilIndex = lbcSCntr.ListIndex
    If ilIndex > 0 Then
        slName = lbcSCntr.List(ilIndex)
    End If
    slCntrStatus = "OH" 'Hold and Orders
    slCntrType = "CTRQSM" 'All but Reservation
    ilCurrent = 0
    ilShow = 2
    ilState = 1
    'ilRet = gPopCntrForAASBox(CntrProj, -1, ilAASCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcPropNo, Traffic!lbcSCntrCode)
    ilAAS = -1
    ilAASCode = 0
    'If ilRowNo > 0 Then
    If ilRowNo >= 0 Then
        gFindMatch smSSave(1, ilRowNo), 1, lbcSAdvt
        If gLastFound(lbcSAdvt) > 0 Then
            slNameCode = tgAdvertiser(gLastFound(lbcSAdvt) - 1).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcSAdvt) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilAAS = 0   'By Advertiser
            ilAASCode = Val(slCode)
        Else
            lbcSCntr.Clear
            sgCntrCodeTag = ""
            ReDim tgCntrCode(0 To 0) As SORTCODE
            lbcSCntr.AddItem "[All]", 0
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        lbcSCntr.Clear
        sgCntrCodeTag = ""
        ReDim tgCntrCode(0 To 0) As SORTCODE
        lbcSCntr.AddItem "[All]", 0
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ilRet = gPopCntrForAASBox(Blackout, ilAAS, ilAASCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcSCntr, tgCntrCode(), sgCntrCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSCntrErr
        gCPErrorMsg ilRet, "mPropNo (gPopCntrForAASBox)", Blackout
        On Error GoTo 0
        lbcSCntr.AddItem "[All]", 0
        imChgMode = True
        If ilIndex >= 1 Then
            gFindMatch slName, 1, lbcSCntr
            If gLastFound(lbcSCntr) > 0 Then
                lbcSCntr.ListIndex = gLastFound(lbcSCntr)
            Else
                lbcSCntr.ListIndex = -1
            End If
        Else
            lbcSCntr.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mSCntrErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSEnableBox                     *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSEnableBox(ilBoxNo As Integer)
'
'   mSEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    Dim ilSDay As Integer
    Dim ilEDay As Integer
    Dim ilFound As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    If (ilBoxNo < imLBSCtrls) Or (ilBoxNo > UBound(tmSCtrls)) Then
        Exit Sub
    End If

    If (imSRowNo < vbcSR.Value) Or (imSRowNo >= vbcSR.Value + vbcSR.LargeChange + 1) Then
        'mSetShow ilBoxNo
        pbcArrow.Visible = False
        lacSFrame(imSRIndex).Visible = False
        Exit Sub
    End If
    lacSFrame(imSRIndex).Move 0, tmSCtrls(SADVTINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15) - 30
    lacSFrame(imSRIndex).Visible = True
    pbcArrow.Move pbcArrow.Left, plcBlackout.Top + tmSCtrls(SADVTINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    If smFromLog = "Y" Then
        Select Case ilBoxNo 'Branch on box type (control)
            Case SADVTINDEX 'Advertiser
                mAdvtPop
                If imTerminate Then
                    Exit Sub
                End If
                If smFromLog = "Y" Then
                    slStr = lbcSAdvt.List(0)
                    If slStr = "[None]" Then
                        lbcSAdvt.List(0) = "[All]"
                    End If
                End If
                lbcSAdvt.height = gListBoxHeight(lbcSAdvt.ListCount, 10)
                edcDropDown.Width = tmSCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 40
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcSAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imSRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcSAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcSAdvt.Move edcDropDown.Left, edcDropDown.Top - lbcSAdvt.height
                End If
                imChgMode = True
                slStr = Trim$(smSSave(1, imSRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcSAdvt
                    If gLastFound(lbcSAdvt) >= 0 Then
                        lbcSAdvt.ListIndex = gLastFound(lbcSAdvt)
                    Else
                        lbcSAdvt.ListIndex = 0
                    End If
                Else
                    lbcSAdvt.ListIndex = 0
                End If
                If lbcSAdvt.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcSAdvt.List(lbcSAdvt.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SCNTRINDEX
                mSCntrPop imSRowNo
                If imTerminate Then
                    Exit Sub
                End If
                lbcSCntr.height = gListBoxHeight(lbcSCntr.ListCount, 10)
                edcDropDown.Width = tmSCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 0
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcSCntr.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imSRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcSCntr.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcSCntr.Move edcDropDown.Left, edcDropDown.Top - lbcSCntr.height
                End If
                imChgMode = True
                If lmSSave(1, imSRowNo) > 0 Then
                    ilFound = False
                    For ilLoop = 0 To UBound(tgCntrCode) - 1 Step 1
                        slNameCode = tgCntrCode(ilLoop).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcSAdvt) - 1)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = lmSSave(1, imSRowNo) Then
                            lbcSCntr.ListIndex = ilLoop + 1
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        lbcSCntr.ListIndex = 0
                    End If
                Else
                    lbcSCntr.ListIndex = 0
                End If
                If lbcSCntr.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcSCntr.List(lbcSCntr.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SRADVTINDEX 'Advertiser
                mAdvtPop
                If imTerminate Then
                    Exit Sub
                End If
                If smFromLog = "Y" Then
                    slStr = lbcSAdvt.List(0)
                    If slStr = "[All]" Then
                        lbcSAdvt.List(0) = "[None]"
                    End If
                End If
                lbcSAdvt.height = gListBoxHeight(lbcSAdvt.ListCount, 10)
                edcDropDown.Width = tmSCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 40
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcSAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imSRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcSAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcSAdvt.Move edcDropDown.Left, edcDropDown.Top - lbcSAdvt.height
                End If
                imChgMode = True
                slStr = Trim$(smSSave(8, imSRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcSAdvt
                    If gLastFound(lbcSAdvt) >= 0 Then
                        lbcSAdvt.ListIndex = gLastFound(lbcSAdvt)
                    Else
                        lbcSAdvt.ListIndex = 0
                    End If
                Else
                    lbcSAdvt.ListIndex = 0
                End If
                If lbcSAdvt.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcSAdvt.List(lbcSAdvt.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SRCNTRINDEX
                mSRCntrPop imSRowNo
                If imTerminate Then
                    Exit Sub
                End If
                lbcSRCntr.height = gListBoxHeight(lbcSRCntr.ListCount, 10)
                edcDropDown.Width = tmSCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 0
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcSRCntr.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imSRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcSRCntr.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcSRCntr.Move edcDropDown.Left, edcDropDown.Top - lbcSRCntr.height
                End If
                imChgMode = True
                If lmSSave(2, imSRowNo) > 0 Then
                    ilFound = False
                    For ilLoop = 0 To UBound(tgManager) - 1 Step 1
                        slNameCode = tgManager(ilLoop).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcSAdvt) - 1)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = lmSSave(2, imSRowNo) Then
                            lbcSRCntr.ListIndex = ilLoop
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        lbcSRCntr.ListIndex = -1
                    End If
                Else
                    lbcSRCntr.ListIndex = -1
                End If
                If lbcSRCntr.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcSRCntr.List(lbcSRCntr.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SRCARTINDEX
                mCartPop imSRowNo
                If imTerminate Then
                    Exit Sub
                End If
                lbcCart.height = gListBoxHeight(lbcCart.ListCount, 10)
                'edcDropDown.Width = tmSCtrls(ilBoxNo).fBoxW + tmSCtrls(SVEHINDEX).fBoxW + tmSCtrls(SLENINDEX).fBoxW + tmSCtrls(SSTARTDATEINDEX).fBoxW + tmSCtrls(SENDDATEINDEX).fBoxW + tmSCtrls(SDAYINDEX).fBoxW - cmcDropDown.Width
                edcDropDown.Width = lbcCart.Width - cmcDropDown.Width
                edcDropDown.MaxLength = 0
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcCart.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imSRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcCart.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcCart.Move edcDropDown.Left, edcDropDown.Top - lbcCart.height
                End If
                imChgMode = True
                slStr = Trim$(smSSave(2, imSRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcCart
                    If gLastFound(lbcCart) >= 0 Then
                        lbcCart.ListIndex = gLastFound(lbcCart)
                    Else
                        If lbcCart.ListCount > 0 Then
                            lbcCart.ListIndex = 0
                        End If
                    End If
                Else
                    If lbcCart.ListCount > 0 Then
                        lbcCart.ListIndex = 0
                    End If
                End If
                If lbcCart.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcCart.List(lbcCart.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SVEHINDEX
                mVehPop
                If imTerminate Then
                    Exit Sub
                End If
                lbcVehicle.height = gListBoxHeight(lbcVehicle.ListCount, 10)
                edcDropDown.Width = tmSCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                If tgSpf.iVehLen <= 40 Then
                    edcDropDown.MaxLength = tgSpf.iVehLen
                Else
                    edcDropDown.MaxLength = 20
                End If
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imSRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.height
                End If
                imChgMode = True
                slStr = Trim$(smSSave(3, imSRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcVehicle
                    If gLastFound(lbcVehicle) >= 0 Then
                        lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                    Else
                        lbcVehicle.ListIndex = 0
                    End If
                Else
                    lbcVehicle.ListIndex = 0
                End If
                If lbcVehicle.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SLENINDEX
                mLenPop imSRowNo
                If imTerminate Then
                    Exit Sub
                End If
                lbcLen.height = gListBoxHeight(lbcLen.ListCount, 10)
                edcDropDown.Width = tmSCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 5
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcLen.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imSRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcLen.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcLen.Move edcDropDown.Left, edcDropDown.Top - lbcLen.height
                End If
                imChgMode = True
                slStr = Trim$(smSSave(9, imSRowNo))
                If (slStr = "") And (lmSSave(3, imSRowNo) > 0) Then
                    tmCifSrchKey.lCode = lmSSave(3, imSRowNo)
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        slStr = Trim$(str$(tmCif.iLen))
                    End If
                End If
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcLen
                    If gLastFound(lbcLen) >= 0 Then
                        lbcLen.ListIndex = gLastFound(lbcLen)
                    Else
                        lbcLen.ListIndex = 0
                    End If
                Else
                    lbcLen.ListIndex = 0
                End If
                If lbcLen.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcLen.List(lbcLen.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SSTARTDATEINDEX 'Start date
                edcDropDown.Width = tmSCtrls(SSTARTDATEINDEX).fBoxW + cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(SSTARTDATEINDEX).fBoxX, tmSCtrls(SSTARTDATEINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcCalendar.height < cmcDone.Top Then
                    plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.height
                End If
                If smSSave(4, imSRowNo) = "" Then
                    slStr = gObtainMondayFromToday()
                Else
                    slStr = smSSave(4, imSRowNo)
                End If
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
                edcDropDown.Text = slStr
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                If smSSave(4, imSRowNo) = "" Then
                    pbcCalendar.Visible = True
                End If
                edcDropDown.SetFocus
            Case SENDDATEINDEX 'Start date
                edcDropDown.Width = tmSCtrls(SENDDATEINDEX).fBoxW + cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(SENDDATEINDEX).fBoxX, tmSCtrls(SENDDATEINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcCalendar.height < cmcDone.Top Then
                    plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top + edcDropDown.height
                Else
                    plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top - plcCalendar.height
                End If
                If smSSave(5, imSRowNo) <> "" Then
                    slStr = smSSave(5, imSRowNo)
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint
                Else
                    'Preset calendar
                    If smSSave(4, imSRowNo) <> "" Then
                        slStr = smSSave(4, imSRowNo)
                        gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                        pbcCalendar_Paint
                    End If
                    slStr = "TFN"
                End If
                edcDropDown.Text = slStr
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SDAYINDEX  'Day index
                lbcDays.height = gListBoxHeight(lbcDays.ListCount, 10)
                edcDropDown.Width = tmSCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 5
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcDays.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imSRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcDays.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcDays.Move edcDropDown.Left, edcDropDown.Top - lbcDays.height
                End If
                imChgMode = True
                slStr = Trim$(smSSave(10, imSRowNo))
                If slStr = "" Then
                    slStr = "Mo-Su"
                End If
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcDays
                    If gLastFound(lbcDays) >= 0 Then
                        lbcDays.ListIndex = gLastFound(lbcDays)
                    Else
                        lbcDays.ListIndex = 0
                    End If
                Else
                    lbcDays.ListIndex = 0
                End If
                If lbcDays.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcDays.List(lbcDays.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SSTARTTIMEINDEX 'Start time
                edcDropDown.Width = tmSCtrls(SSTARTTIMEINDEX).fBoxW - cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(SSTARTTIMEINDEX).fBoxX, tmSCtrls(SSTARTTIMEINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcTme.height < cmcDone.Top Then
                    plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.height
                End If
                If smSSave(6, imSRowNo) <> "" Then
                    edcDropDown.Text = smSSave(6, imSRowNo)
                Else
                    edcDropDown.Text = "12M"
                End If
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                If smSSave(6, imSRowNo) = "" Then
                    plcTme.Visible = True
                End If
                edcDropDown.SetFocus
            Case SENDTIMEINDEX 'Start time
                edcDropDown.Width = tmSCtrls(SENDTIMEINDEX).fBoxW - cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(SENDTIMEINDEX).fBoxX, tmSCtrls(SENDTIMEINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcTme.height < cmcDone.Top Then
                    plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.height
                End If
                If smSSave(7, imSRowNo) <> "" Then
                    edcDropDown.Text = smSSave(7, imSRowNo)
                Else
                    If (Trim$(smSSave(1, imSRowNo)) = "[All]") Or (Trim$(smSSave(1, imSRowNo)) = "[None]") Or (Trim$(smSSave(1, imSRowNo)) = "") Then
                        edcDropDown.Text = smSSave(6, imSRowNo)
                    Else
                        edcDropDown.Text = "12M"
                    End If
                End If
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                If smSSave(7, imSRowNo) = "" Then
                    plcTme.Visible = True
                End If
                edcDropDown.SetFocus
        End Select
    Else
        Select Case ilBoxNo 'Branch on box type (control)
            Case SADVTINDEX 'Advertiser
                mAdvtPop
                If imTerminate Then
                    Exit Sub
                End If
                lbcSAdvt.height = gListBoxHeight(lbcSAdvt.ListCount, 10)
                edcDropDown.Width = tmSCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                edcDropDown.MaxLength = 40
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcSAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imSRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcSAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcSAdvt.Move edcDropDown.Left, edcDropDown.Top - lbcSAdvt.height
                End If
                imChgMode = True
                slStr = Trim$(smSSave(1, imSRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcSAdvt
                    If gLastFound(lbcSAdvt) >= 0 Then
                        lbcSAdvt.ListIndex = gLastFound(lbcSAdvt)
                    Else
                        lbcSAdvt.ListIndex = 0
                    End If
                Else
                    lbcSAdvt.ListIndex = 0
                End If
                If lbcSAdvt.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcSAdvt.List(lbcSAdvt.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SSHORTTITLEINDEX
                mShtTitlePop imSRowNo
                If imTerminate Then
                    Exit Sub
                End If
                lbcShtTitle.height = gListBoxHeight(lbcShtTitle.ListCount, 10)
                edcDropDown.Width = tmSCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                If tgSpf.sUseProdSptScr = "P" Then  'Short Title
                    edcDropDown.MaxLength = 15
                Else
                    edcDropDown.MaxLength = 35
                End If
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcShtTitle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imSRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcShtTitle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcShtTitle.Move edcDropDown.Left, edcDropDown.Top - lbcShtTitle.height
                End If
                imChgMode = True
                slStr = Trim$(smSSave(2, imSRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 1, lbcShtTitle
                    If gLastFound(lbcShtTitle) > 0 Then
                        lbcShtTitle.ListIndex = gLastFound(lbcShtTitle)
                    Else
                        lbcShtTitle.ListIndex = 0
                    End If
                Else
                    lbcShtTitle.ListIndex = 0
                End If
                If lbcShtTitle.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcShtTitle.List(lbcShtTitle.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SVEHINDEX
                mVehPop
                If imTerminate Then
                    Exit Sub
                End If
                lbcVehicle.height = gListBoxHeight(lbcVehicle.ListCount, 10)
                edcDropDown.Width = tmSCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
                If tgSpf.iVehLen <= 40 Then
                    edcDropDown.MaxLength = tgSpf.iVehLen
                Else
                    edcDropDown.MaxLength = 20
                End If
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                If imSRowNo - vbcSR.Value <= vbcSR.LargeChange \ 2 Then
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.height
                End If
                imChgMode = True
                slStr = Trim$(smSSave(3, imSRowNo))
                If slStr <> "" Then
                    gFindMatch slStr, 0, lbcVehicle
                    If gLastFound(lbcVehicle) >= 0 Then
                        lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                    Else
                        lbcVehicle.ListIndex = 0
                    End If
                Else
                    lbcVehicle.ListIndex = 0
                End If
                If lbcVehicle.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SSTARTDATEINDEX 'Start date
                edcDropDown.Width = tmSCtrls(SSTARTDATEINDEX).fBoxW + cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(SSTARTDATEINDEX).fBoxX, tmSCtrls(SSTARTDATEINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcCalendar.height < cmcDone.Top Then
                    plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.height
                End If
                If smSSave(4, imSRowNo) = "" Then
                    slStr = gObtainMondayFromToday()
                Else
                    slStr = smSSave(4, imSRowNo)
                End If
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
                edcDropDown.Text = slStr
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                If smSSave(4, imSRowNo) = "" Then
                    pbcCalendar.Visible = True
                End If
                edcDropDown.SetFocus
            Case SENDDATEINDEX 'Start date
                edcDropDown.Width = tmSCtrls(SENDDATEINDEX).fBoxW + cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(SENDDATEINDEX).fBoxX, tmSCtrls(SENDDATEINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcCalendar.height < cmcDone.Top Then
                    plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top + edcDropDown.height
                Else
                    plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top - plcCalendar.height
                End If
                If smSSave(5, imSRowNo) <> "" Then
                    slStr = smSSave(5, imSRowNo)
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint
                Else
                    'Preset calendar
                    If smSSave(4, imSRowNo) <> "" Then
                        slStr = smSSave(4, imSRowNo)
                        gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                        pbcCalendar_Paint
                    End If
                    slStr = "TFN"
                End If
                edcDropDown.Text = slStr
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SDAYINDEX To SDAYINDEX + 6 'Day index
                gMoveTableCtrl pbcSuppression(imSRIndex), pbcDays, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                If imSSave(ilBoxNo - SDAYINDEX + 1, imSRowNo) = 1 Then
                    ckcDay.Value = vbChecked
                ElseIf imSSave(ilBoxNo - SDAYINDEX + 1, imSRowNo) = -1 Then
                    If (Trim$(smSSave(4, imSRowNo)) = "") Or (Trim$(smSSave(5, imSRowNo)) = "") Then
                        ckcDay.Value = vbChecked
                    Else
                        ilSDay = gWeekDayStr(smSSave(4, imSRowNo))
                        ilEDay = gWeekDayStr(smSSave(5, imSRowNo))
                        If ilSDay <= ilEDay Then
                            If ((ilBoxNo - SDAYINDEX) >= ilSDay) And ((ilBoxNo - SDAYINDEX) <= ilEDay) Then
                                ckcDay.Value = vbChecked
                            Else
                                ckcDay.Value = vbUnchecked
                            End If
                        Else
                            If ((ilBoxNo - SDAYINDEX) >= ilSDay) And ((ilBoxNo - SDAYINDEX) <= 6) Then
                                ckcDay.Value = vbChecked
                            Else
                                If ((ilBoxNo - SDAYINDEX) >= 0) And ((ilBoxNo - SDAYINDEX) <= ilEDay) Then
                                    ckcDay.Value = vbChecked
                                Else
                                    ckcDay.Value = vbUnchecked
                                End If
                            End If
                        End If
                    End If
                Else
                    ckcDay.Value = vbUnchecked
                End If
                pbcDays.Visible = True  'Set visibility
                ckcDay.SetFocus
            Case SSTARTTIMEINDEX 'Start time
                edcDropDown.Width = tmSCtrls(SSTARTTIMEINDEX).fBoxW - cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(SSTARTTIMEINDEX).fBoxX, tmSCtrls(SSTARTTIMEINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcTme.height < cmcDone.Top Then
                    plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.height
                End If
                If smSSave(6, imSRowNo) <> "" Then
                    edcDropDown.Text = smSSave(6, imSRowNo)
                Else
                    edcDropDown.Text = "12M"
                End If
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                If smSSave(6, imSRowNo) = "" Then
                    plcTme.Visible = True
                End If
                edcDropDown.SetFocus
            Case SENDTIMEINDEX 'Start time
                edcDropDown.Width = tmSCtrls(SENDTIMEINDEX).fBoxW - cmcDropDown.Width
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcSuppression(imSRIndex), edcDropDown, tmSCtrls(SENDTIMEINDEX).fBoxX, tmSCtrls(SENDTIMEINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15)
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                If edcDropDown.Top + edcDropDown.height + plcTme.height < cmcDone.Top Then
                    plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                Else
                    plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.height
                End If
                If smSSave(7, imSRowNo) <> "" Then
                    edcDropDown.Text = smSSave(7, imSRowNo)
                Else
                    If (Trim$(smSSave(1, imSRowNo)) = "[None]") Or (Trim$(smSSave(1, imSRowNo)) = "") Then
                        edcDropDown.Text = smSSave(6, imSRowNo)
                    Else
                        edcDropDown.Text = "12M"
                    End If
                End If
                edcDropDown.Visible = True  'Set visibility
                cmcDropDown.Visible = True
                If smSSave(7, imSRowNo) = "" Then
                    plcTme.Visible = True
                End If
                edcDropDown.SetFocus
        End Select
    End If
    mSetCommands
End Sub
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
    'If (imBypassSetting) Or (Not imUpdateAllowed) Then
    '    Exit Sub
    'End If
    ilAltered = imSChg
    If Not ilAltered Then
        ilAltered = imRChg
    End If
    If (Not ilAltered) And (UBound(tgBofDel) > LBound(tgBofDel)) Then
        ilAltered = True
    End If
    'Update button set if all mandatory fields have data and any field altered
    'If (mTestFields() = YES) And (ilAltered) And ((UBound(tgSBofRec) > 1) Or (UBound(tgRBofRec) > 1) Or (UBound(tgBofDel) > LBound(tgBofDel))) Then
    If (mTestFields() = YES) And (ilAltered) And ((UBound(tgSBofRec) > 0) Or (UBound(tgRBofRec) > 0) Or (UBound(tgBofDel) > LBound(tgBofDel))) Then
        If imUpdateAllowed Then
            cmcUpdate.Enabled = True
        Else
            cmcUpdate.Enabled = False
        End If
    Else
        cmcUpdate.Enabled = False
    End If
    If imSChg Or imRChg Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
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
    If igView = 0 Then
        vbcSR.Min = LBound(smSShow, 2)
        imSettingValue = True
        If UBound(smSShow, 2) - 1 <= vbcSR.LargeChange + 1 Then ' + 1 Then
            vbcSR.Max = LBound(smSShow, 2)
        Else
            vbcSR.Max = UBound(smSShow, 2) - vbcSR.LargeChange
        End If
        imSettingValue = True
        If vbcSR.Value = vbcSR.Min Then
            vbcSR_Change
        Else
            vbcSR.Value = vbcSR.Min
        End If
        imSettingValue = False
    Else
        vbcSR.Min = LBound(smRShow, 2)
        imSettingValue = True
        If UBound(smRShow, 2) - 1 <= vbcSR.LargeChange + 1 Then ' + 1 Then
            vbcSR.Max = LBound(smRShow, 2)
        Else
            vbcSR.Max = UBound(smRShow, 2) - vbcSR.LargeChange
        End If
        imSettingValue = True
        If vbcSR.Value = vbcSR.Min Then
            vbcSR_Change
        Else
            vbcSR.Value = vbcSR.Min
        End If
        imSettingValue = False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mShtTitlePop                    *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate advertiser short title*
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mShtTitlePop(ilRowNo As Integer)
'
'   mShtTitlePop
'   Where:
'       ilAdvtCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilAdfCode As Integer
    ilIndex = lbcShtTitle.ListIndex
    If ilIndex > 0 Then
        slName = lbcShtTitle.List(ilIndex)
    End If
    'If ilRowNo <= 0 Then
    If ilRowNo < 0 Then
        Exit Sub
    End If
    If igView = 0 Then
        slStr = Trim$(smSSave(1, ilRowNo))
        If smFromLog = "Y" Then
            gFindMatch slStr, 0, lbcSAdvt
            If gLastFound(lbcSAdvt) <= 0 Then
                Exit Sub
            End If
            slNameCode = tgAdvertiser(gLastFound(lbcSAdvt) - 1).sKey  'Traffic!lbcAdvt.List(imAdvtIndex)
        Else
            gFindMatch slStr, 0, lbcSAdvt
            If gLastFound(lbcSAdvt) < 0 Then
                Exit Sub
            End If
            slNameCode = tgAdvertiser(gLastFound(lbcSAdvt)).sKey    'Traffic!lbcAdvt.List(imAdvtIndex)
        End If
    Else
        slStr = Trim$(smRSave(1, ilRowNo))
        gFindMatch slStr, 0, lbcRAdvt
        If gLastFound(lbcRAdvt) < 0 Then
            Exit Sub
        End If
        slNameCode = tgAdvertiser(gLastFound(lbcRAdvt)).sKey    'Traffic!lbcAdvt.List(imAdvtIndex)
    End If
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilAdfCode = Val(slCode)   '
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtProdBox(Copy, ilAdfCode, lbcShtTitle, lbcShtTitleCode)
    'ilRet = gPopShortTitleBox(Blackout, ilAdfCode, lbcShtTitle, tmShtTitleCode(), smShtTitleCodeTag)
    If tgSpf.sUseProdSptScr = "P" Then
        ilRet = gPopShortTitleBox(Blackout, ilAdfCode, lbcShtTitle, tmShtTitleCode(), smShtTitleCodeTag)
    Else
        ilRet = gPopAdvtProdBox(Blackout, ilAdfCode, lbcShtTitle, tmShtTitleCode(), smShtTitleCodeTag)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mShtTitlePopErr
        gCPErrorMsg ilRet, "mShtTitlePop (gIMoveListBox)", Blackout
        On Error GoTo 0
        lbcShtTitle.AddItem "[None]", 0  'Force as first item on list
'            lbcShtTitle.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcShtTitle
            If gLastFound(lbcShtTitle) > 0 Then
                lbcShtTitle.ListIndex = gLastFound(lbcShtTitle)
            Else
                lbcShtTitle.ListIndex = -1
            End If
        Else
            lbcShtTitle.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mShtTitlePopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCntrPop                        *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Contract Population            *
'*                                                     *
'*******************************************************
Private Sub mSRCntrPop(ilRowNo As Integer)
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilCurrent As Integer
    Dim ilShow As Integer
    Dim ilState As Integer
    Dim ilAAS As Integer
    Dim ilAASCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    'If Not imPropPopReqd Then
    '    Exit Sub
    'End If
    Screen.MousePointer = vbHourglass
    ilIndex = lbcSRCntr.ListIndex
    If ilIndex >= 0 Then
        slName = lbcSRCntr.List(ilIndex)
    End If
    slCntrStatus = "OH" 'Hold and Orders
    slCntrType = "CTRQSM" 'Standard only
    ilCurrent = 0
    ilShow = 2
    ilState = 1
    'ilRet = gPopCntrForAASBox(CntrProj, -1, ilAASCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcPropNo, Traffic!lbcSCntrCode)
    ilAAS = -1
    ilAASCode = 0
    'If ilRowNo > 0 Then
    If ilRowNo >= 0 Then
        gFindMatch smSSave(8, ilRowNo), 1, lbcSAdvt
        If gLastFound(lbcSAdvt) > 0 Then
            slNameCode = tgAdvertiser(gLastFound(lbcSAdvt) - 1).sKey  'Traffic!lbcAdvertiser.List(gLastFound(lbcSAdvt) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilAAS = 0   'By Advertiser
            ilAASCode = Val(slCode)
        Else
            lbcSRCntr.Clear
            sgManagerTag = ""
            ReDim tgManager(0 To 0) As SORTCODE
            'lbcSRCntr.AddItem "[All]", 0
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        lbcSRCntr.Clear
        sgManagerTag = ""
        ReDim tgManager(0 To 0) As SORTCODE
        'lbcSRCntr.AddItem "[All]", 0
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ilRet = gPopCntrForAASBox(Blackout, ilAAS, ilAASCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcSRCntr, tgManager(), sgManagerTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSRCntrErr
        gCPErrorMsg ilRet, "mPropNo (gPopCntrForAASBox)", Blackout
        On Error GoTo 0
        'lbcSRCntr.AddItem "[All]", 0
        imChgMode = True
        If ilIndex >= 0 Then
            gFindMatch slName, 0, lbcSRCntr
            If gLastFound(lbcSRCntr) > 0 Then
                lbcSRCntr.ListIndex = gLastFound(lbcSRCntr)
            Else
                lbcSRCntr.ListIndex = -1
            End If
        Else
            lbcSRCntr.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mSRCntrErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSSetFocus(ilBoxNo As Integer)
'
'   mSSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBSCtrls) Or (ilBoxNo > UBound(tmSCtrls)) Then
        Exit Sub
    End If

    If (imSRowNo < vbcSR.Value) Or (imSRowNo >= vbcSR.Value + vbcSR.LargeChange + 1) Then
        'mSetShow ilBoxNo
        pbcArrow.Visible = False
        lacSFrame(imSRIndex).Visible = False
        Exit Sub
    End If
    lacSFrame(imSRIndex).Move 0, tmSCtrls(SADVTINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15) - 30
    lacSFrame(imSRIndex).Visible = True
    pbcArrow.Move pbcArrow.Left, plcBlackout.Top + tmSCtrls(SADVTINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    If smFromLog = "Y" Then
        Select Case ilBoxNo 'Branch on box type (control)
            Case SADVTINDEX 'Advertiser
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SCNTRINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SRADVTINDEX 'Advertiser
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SRCNTRINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SRCARTINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SVEHINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SLENINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SSTARTDATEINDEX 'Start date
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SENDDATEINDEX 'Start date
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SDAYINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SSTARTTIMEINDEX 'Start time
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SENDTIMEINDEX 'Start time
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
        End Select
    Else
        Select Case ilBoxNo 'Branch on box type (control)
            Case SADVTINDEX 'Advertiser
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SSHORTTITLEINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SVEHINDEX
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SSTARTDATEINDEX 'Start date
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SENDDATEINDEX 'Start date
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SDAYINDEX To SDAYINDEX + 6 'Day index
                pbcDays.Visible = True  'Set visibility
                ckcDay.SetFocus
            Case SSTARTTIMEINDEX 'Start time
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            Case SENDTIMEINDEX 'Start time
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
        End Select
    End If
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSSetShow                       *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSSetShow(ilBoxNo As Integer)
'
'   mSSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    pbcArrow.Visible = False
    lacSFrame(imSRIndex).Visible = False
    lacRFrame(imSRIndex).Visible = False
    If (ilBoxNo < imLBSCtrls) Or (ilBoxNo > UBound(tmSCtrls)) Then
        Exit Sub
    End If
    If smFromLog = "Y" Then
        Select Case ilBoxNo 'Branch on box type (control)
            Case SADVTINDEX
                lbcSAdvt.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                If Trim$(smSSave(1, imSRowNo)) <> slStr Then
                    imSChg = True
                    smSSave(1, imSRowNo) = slStr
                    lmSSave(1, imSRowNo) = 0
                    smSSave(11, imSRowNo) = ""
                    smSShow(SCNTRINDEX, imSRowNo) = ""
                End If
                If smFromLog = "Y" Then
                    slStr = lbcSAdvt.List(0)
                    If slStr = "[All]" Then
                        lbcSAdvt.List(0) = "[None]"
                    End If
                End If
            Case SCNTRINDEX
                lbcSCntr.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                If lbcSCntr.ListIndex > 0 Then
                    slNameCode = tgCntrCode(lbcSCntr.ListIndex - 1).sKey 'Traffic!lbcAdvertiser.List(gLastFound(lbcSAdvt) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) <> lmSSave(1, imSRowNo) Then
                        imSChg = True
                        lmSSave(1, imSRowNo) = Val(slCode)
                        ilPos = InStr(1, slNameCode, " ", 1)
                        If ilPos > 0 Then
                            smSSave(11, imSRowNo) = Left$(slNameCode, ilPos - 1)
                        Else
                            ilRet = gParseItem(slNameCode, 1, "\", smSSave(11, imSRowNo))
                        End If
                    End If
                Else
                    If lmSSave(1, imSRowNo) <> 0 Then
                        imSChg = True
                        lmSSave(1, imSRowNo) = 0
                        smSSave(11, imSRowNo) = ""
                    End If
                End If
            Case SRADVTINDEX
                lbcSAdvt.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                If Trim$(smSSave(8, imSRowNo)) <> slStr Then
                    imSChg = True
                    smSSave(8, imSRowNo) = slStr
                    lmSSave(2, imSRowNo) = 0
                    smSSave(12, imSRowNo) = ""
                    smSShow(SRCNTRINDEX, imSRowNo) = ""
                    lmSSave(3, imSRowNo) = 0
                    lmSSave(4, imSRowNo) = 0
                    smSSave(2, imSRowNo) = ""   'Clear Cart
                    smSShow(SRCARTINDEX, imSRowNo) = ""
                End If
            Case SRCNTRINDEX
                lbcSRCntr.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                If lbcSRCntr.ListIndex >= 0 Then
                    slNameCode = tgManager(lbcSRCntr.ListIndex).sKey 'Traffic!lbcAdvertiser.List(gLastFound(lbcSAdvt) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) <> lmSSave(2, imSRowNo) Then
                        imSChg = True
                        lmSSave(2, imSRowNo) = Val(slCode)
                        ilPos = InStr(1, slNameCode, " ", 1)
                        If ilPos > 0 Then
                            smSSave(12, imSRowNo) = Left$(slNameCode, ilPos - 1)
                        Else
                            ilRet = gParseItem(slNameCode, 1, "\", smSSave(12, imSRowNo))
                        End If
                    End If
                Else
                    If lmSSave(2, imSRowNo) <> 0 Then
                        imSChg = True
                        lmSSave(2, imSRowNo) = 0
                        smSSave(12, imSRowNo) = ""
                    End If
                End If
            Case SRCARTINDEX
                lbcCart.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                If Trim$(smSSave(2, imSRowNo)) <> slStr Then
                    imRChg = True
                    smSSave(2, imSRowNo) = slStr
                    If (lbcCart.ListIndex >= 0) And (slStr <> "[None]") Then
                        slNameCode = tmCartCode(lbcCart.ListIndex - 1).sKey 'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        lmSSave(3, imSRowNo) = Val(slCode)
                        ilRet = gParseItem(slNameCode, 3, "\", slName)
                        ilRet = gParseItem(slNameCode, 4, "\", slCode)
                        lmSSave(4, imSRowNo) = Val(slCode)
                    Else
                        lmSSave(3, imSRowNo) = 0
                        lmSSave(4, imSRowNo) = 0
                    End If
                End If
            Case SVEHINDEX
                lbcVehicle.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                If Trim$(smSSave(3, imSRowNo)) <> slStr Then
                    imSChg = True
                    smSSave(3, imSRowNo) = slStr
                End If
            Case SLENINDEX
                lbcLen.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                If Trim$(smSSave(9, imSRowNo)) <> slStr Then
                    imSChg = True
                    smSSave(9, imSRowNo) = slStr
                End If
            Case SSTARTDATEINDEX 'Start date
                plcCalendar.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gValidDate(slStr) Then
                    gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                    smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                    If Trim$(smSSave(4, imSRowNo)) <> slStr Then
                        imSChg = True
                        smSSave(4, imSRowNo) = slStr
                    End If
                Else
                    Beep
                End If
            Case SENDDATEINDEX 'Start date
                plcCalendar.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If StrComp(slStr, "TFN", 1) <> 0 Then
                    If gValidDate(slStr) Then
                        If Trim$(smSSave(5, imSRowNo)) <> slStr Then
                            imSChg = True
                            smSSave(5, imSRowNo) = slStr
                        End If
                        slStr = gFormatDate(slStr)
                        gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                        smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                Else
                    If Trim$(smSSave(5, imSRowNo)) <> "" Then
                        imSChg = True
                    End If
                    smSSave(5, imSRowNo) = ""
                    slStr = "TFN"
                    gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                    smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                End If
            Case SDAYINDEX   'Day index
                lbcDays.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                If Trim$(smSSave(10, imSRowNo)) <> slStr Then
                    imSChg = True
                    smSSave(10, imSRowNo) = slStr
                End If
            Case SSTARTTIMEINDEX 'Start time
                cmcDropDown.Visible = False
                plcTme.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If gValidTime(slStr) Then
                        If Trim$(smSSave(6, imSRowNo)) <> slStr Then
                            imSChg = True
                            smSSave(6, imSRowNo) = slStr
                        End If
                        slStr = gFormatTime(slStr, "A", "1")
                        gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                        smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                End If
            Case SENDTIMEINDEX 'Start time
                cmcDropDown.Visible = False
                plcTme.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If gValidTime(slStr) Then
                        If Trim$(smSSave(7, imSRowNo)) <> slStr Then
                            imSChg = True
                            smSSave(7, imSRowNo) = slStr
                        End If
                        slStr = gFormatTime(slStr, "A", "1")
                        gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                        smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                End If
        End Select
    Else
        Select Case ilBoxNo 'Branch on box type (control)
            Case SADVTINDEX
                lbcSAdvt.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                If Trim$(smSSave(1, imSRowNo)) <> slStr Then
                    imSChg = True
                    smSSave(1, imSRowNo) = slStr
                    smSSave(2, imSRowNo) = ""   'Clear short title/Product
                    smSShow(SSHORTTITLEINDEX, imSRowNo) = ""
                End If
            Case SSHORTTITLEINDEX
                lbcShtTitle.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                If Trim$(smSSave(2, imSRowNo)) <> slStr Then
                    imSChg = True
                    smSSave(2, imSRowNo) = slStr
                End If
            Case SVEHINDEX
                lbcVehicle.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                If Trim$(smSSave(3, imSRowNo)) <> slStr Then
                    imSChg = True
                    smSSave(3, imSRowNo) = slStr
                End If
            Case SSTARTDATEINDEX 'Start date
                plcCalendar.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gValidDate(slStr) Then
                    gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                    smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                    If Trim$(smSSave(4, imSRowNo)) <> slStr Then
                        imSChg = True
                        smSSave(4, imSRowNo) = slStr
                    End If
                Else
                    Beep
                End If
            Case SENDDATEINDEX 'Start date
                plcCalendar.Visible = False
                cmcDropDown.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If StrComp(slStr, "TFN", 1) <> 0 Then
                    If gValidDate(slStr) Then
                        If Trim$(smSSave(5, imSRowNo)) <> slStr Then
                            imSChg = True
                            smSSave(5, imSRowNo) = slStr
                        End If
                        slStr = gFormatDate(slStr)
                        gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                        smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                Else
                    If Trim$(smSSave(5, imSRowNo)) <> "" Then
                        imSChg = True
                    End If
                    smSSave(5, imSRowNo) = ""
                    slStr = "TFN"
                    gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                    smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                End If
            Case SDAYINDEX To SDAYINDEX + 6 'Day index
                pbcDays.Visible = False  'Set visibility
                If ckcDay.Value = vbChecked Then
                    If imSSave(ilBoxNo - SDAYINDEX + 1, imSRowNo) <> 1 Then
                        imSChg = True
                    End If
                    slStr = "4"
                    imSSave(ilBoxNo - SDAYINDEX + 1, imSRowNo) = 1
                Else
                    If imSSave(ilBoxNo - SDAYINDEX + 1, imSRowNo) <> 0 Then
                        imSChg = True
                    End If
                    slStr = "  "
                    imSSave(ilBoxNo - SDAYINDEX + 1, imSRowNo) = 0
                End If
                gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
            Case SSTARTTIMEINDEX 'Start time
                cmcDropDown.Visible = False
                plcTme.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If gValidTime(slStr) Then
                        If Trim$(smSSave(6, imSRowNo)) <> slStr Then
                            imSChg = True
                            smSSave(6, imSRowNo) = slStr
                        End If
                        slStr = gFormatTime(slStr, "A", "1")
                        gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                        smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                End If
            Case SENDTIMEINDEX 'Start time
                cmcDropDown.Visible = False
                plcTme.Visible = False
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If gValidTime(slStr) Then
                        If Trim$(smSSave(7, imSRowNo)) <> slStr Then
                            imSChg = True
                            smSSave(7, imSRowNo) = slStr
                        End If
                        slStr = gFormatTime(slStr, "A", "1")
                        gSetShow pbcSuppression(imSRIndex), slStr, tmSCtrls(ilBoxNo)
                        smSShow(ilBoxNo, imSRowNo) = tmSCtrls(ilBoxNo).sShow
                    Else
                        Beep
                    End If
                End If
        End Select
    End If
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSTestSaveFields                *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mSTestSaveFields(ilRowNo As Integer, ilMsg As Integer) As Integer
'
'   iRet = mSTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilOneDay As Integer
    Dim ilDay As Integer
    Dim ilRet As Integer

    If Trim$(smSSave(1, ilRowNo)) = "" Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Advertiser must be specified", vbOKOnly + vbExclamation, "In Suppress")
        End If
        imSBoxNo = SADVTINDEX
        mSTestSaveFields = NO
        Exit Function
    End If
    'If Trim$(smSSave(2, ilRowNo)) = "" Then
    '    Beep
    '    ilRes = MsgBox("Short Title must be specified", vbOkOnly + vbExclamation, "Incomplete")
    '    imSBoxNo = SSHORTTITLEINDEX
    '    mSTestSaveFields = NO
    '    Exit Function
    'End If
    If (Trim$(smSSave(8, ilRowNo)) <> "") And (Trim$(smSSave(8, ilRowNo)) <> "[None]") And (lmSSave(2, ilRowNo) <= 0) And (imSRIndex = 1) Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Replace Contract must be specified for " & Trim$(smSSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Suppress")
        End If
        imSBoxNo = SRCNTRINDEX
        mSTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smSSave(3, ilRowNo)) = "" Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Vehicle must be specified for " & Trim$(smSSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Suppress")
        End If
        imSBoxNo = SVEHINDEX
        mSTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smSSave(4, ilRowNo)) = "" Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Start Date must be specified for " & Trim$(smSSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Suppress")
        End If
        imSBoxNo = SSTARTDATEINDEX
        mSTestSaveFields = NO
        Exit Function
    Else
        If Not gValidDate(smSSave(4, ilRowNo)) Then
            Beep
            If ilMsg Then
                ilRes = MsgBox("Start Date must be valid for " & Trim$(smSSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Suppress")
            End If
            imSBoxNo = SSTARTDATEINDEX
            mSTestSaveFields = NO
            Exit Function
        End If
    End If
    If Trim$(smSSave(5, ilRowNo)) <> "" Then
        If Not gValidDate(smSSave(5, ilRowNo)) Then
            Beep
            If ilMsg Then
                ilRes = MsgBox("End Date must be valid for " & Trim$(smSSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Suppress")
            End If
            imSBoxNo = SENDDATEINDEX
            mSTestSaveFields = NO
            Exit Function
        End If
    End If
    ilOneDay = False
    If smFromLog = "Y" Then
        If smSSave(10, ilRowNo) <> "" Then
            ilOneDay = True
        End If
    Else
        For ilDay = 0 To 6 Step 1
            If imSSave(ilDay + 1, ilRowNo) > 0 Then
                ilOneDay = True
                Exit For
            End If
        Next ilDay
    End If
    If Not ilOneDay Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("One Day must be specified for " & Trim$(smSSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Suppress")
        End If
        imSBoxNo = SDAYINDEX
        mSTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smSSave(6, ilRowNo)) = "" Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Start Time must be specified for " & Trim$(smSSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Suppress")
        End If
        imSBoxNo = SSTARTTIMEINDEX
        mSTestSaveFields = NO
        Exit Function
    Else
        If Not gValidTime(smSSave(6, ilRowNo)) Then
            Beep
            If ilMsg Then
                ilRes = MsgBox("Start Time must be valid for " & Trim$(smSSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Suppress")
            End If
            imSBoxNo = SSTARTTIMEINDEX
            mSTestSaveFields = NO
            Exit Function
        End If
    End If
    If Trim$(smSSave(7, ilRowNo)) = "" Then
        Beep
        If ilMsg Then
            ilRes = MsgBox("Start Time must be specified for " & Trim$(smSSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Suppress")
        End If
        imSBoxNo = SSTARTTIMEINDEX
        mSTestSaveFields = NO
        Exit Function
    Else
        If Not gValidTime(smSSave(7, ilRowNo)) Then
            Beep
            If ilMsg Then
                ilRes = MsgBox("Start Time must be valid for " & Trim$(smSSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Suppress")
            End If
            imSBoxNo = SENDTIMEINDEX
            mSTestSaveFields = NO
            Exit Function
        End If
    End If
    If smFromLog = "Y" Then
        If lmSSave(3, ilRowNo) > 0 Then
            'If (smSSave(9, ilRowNo) <> "") And (smSSave(9, ilRowNo) <> "[All]") Then
                tmCifSrchKey.lCode = lmSSave(3, ilRowNo)
                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    If Val(smSSave(9, ilRowNo)) <> tmCif.iLen Then
                        Beep
                        If ilMsg Then
                            ilRes = MsgBox("Copy Length and Suppress Length not matching for " & Trim$(smSSave(1, ilRowNo)), vbOKOnly + vbExclamation, "In Suppress")
                        End If
                        imSBoxNo = SLENINDEX
                        mSTestSaveFields = NO
                        Exit Function
                    End If
                End If
            'End If
        End If
    End If
    mSTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
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


    sgDoneMsg = "Done"
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload Blackout
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
    For ilRowNo = LBound(smSSave, 2) To UBound(smSSave, 2) - 1 Step 1
        If smSSave(1, ilRowNo) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If (smSSave(8, ilRowNo) <> "") And (smSSave(8, ilRowNo) <> "[None]") And (lmSSave(2, ilRowNo) <= 0) And (imSRIndex = 1) Then
            mTestFields = NO
            Exit Function
        End If
        'If Trim$(smSSave(2, ilRowNo)) = "" Then
        '    mTestFields = NO
        '    Exit Function
        'End If
        If Trim$(smSSave(3, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSSave(4, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        'If Trim$(smSSave(5, ilRowNo)) = "" Then
        '    mTestFields = NO
        '    Exit Function
        'End If
        If Trim$(smSSave(6, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSSave(7, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
    Next ilRowNo
    For ilRowNo = LBound(smRSave, 2) To UBound(smRSave, 2) - 1 Step 1
        If smRSave(1, ilRowNo) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If (Trim$(smRSave(2, ilRowNo)) = "") And (imSRIndex = 0) Then
            mTestFields = NO
            Exit Function
        End If
        If (Trim$(smRSave(2, ilRowNo)) = "") And (Trim$(smRSave(11, ilRowNo)) = "") And (imSRIndex = 1) Then
            mTestFields = NO
            Exit Function
        End If
        'If Trim$(smRSave(3, ilRowNo)) = "" Then
        '    mTestFields = NO
        '    Exit Function
        'End If
        'If Trim$(smRSave(4, ilRowNo)) = "" Then
        '    mTestFields = NO
        '    Exit Function
        'End If
        'If Trim$(smRSave(5, ilRowNo)) = "" Then
        '    mTestFields = NO
        '    Exit Function
        'End If
        If Trim$(smRSave(6, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        'If Trim$(smRSave(7, ilRowNo)) = "" Then
        '    mTestFields = NO
        '    Exit Function
        'End If
        If Trim$(smRSave(8, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smRSave(9, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
    Next ilRowNo
    mTestFields = YES
End Function
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
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim llFilter As Long
    If (smFromLog = "Y") Then
        If smSplitFill = "Y" Then
            '10/13/10:  Add Airing vehicle
            llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH + VEHEXCLUDEPODNOPRGM  ' all conventional vehicles since Split can only be defined for conventional vehicles, 1/25/21 Exlude CPM vehicles, 2/8/21 - exclude podcast w/o program based on ad server.
        Else
            llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH + VEHEXCLUDEPODNOPRGM  ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 Exlude CPM vehicles, 2/8/21 - exclude podcast w/o program based on ad server.
        End If
    Else
        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH + VEHEXCLUDEPODNOPRGM  ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 Exlude CPM vehicles, 2/8/21 - exclude podcast w/o program based on ad server.
    End If
    'ilRet = gPopUserVehicleBox(Budget, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, cbcCtrl, Traffic!lbcUserVehicle)
    ilRet = gPopUserVehicleBox(Blackout, llFilter, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    'ilRet = gPopUserVehComboBox(Blackout, cbcCtrl, Traffic!lbcUserVehicle, lbcCombo)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        'gCPErrorMsg ilRet, "mVehPop (gPopUserVehComboBox: Vehicle/Combo)", Blackout
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", Blackout
        On Error GoTo 0
        If smFromLog = "Y" Then
            lbcVehicle.AddItem "[All]", 0
        End If
    End If
    ''ReDim tmUserVeh(1 To UBound(tgUserVehicle) + 1) As USERVEH
    'ReDim tmUserVeh(0 To UBound(tgUserVehicle) + 1) As USERVEH
    'For ilLoop = 0 To UBound(tgUserVehicle) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
    '    slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
    '    ilRet = gParseItem(slNameCode, 1, "\", slName)
    '    ilRet = gParseItem(slName, 3, "|", tmUserVeh(ilLoop + 1).sName)
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '    tmUserVeh(ilLoop + 1).iCode = Val(slCode)
    'Next ilLoop
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub pbcArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
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
                edcDropDown.Text = Format$(llDate, "m/d/yy")
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                imBypassFocus = True
                edcDropDown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcDropDown.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    imSRowNo = -1
    mRSetShow imRBoxNo
    imRBoxNo = -1
    imRRowNo = -1
    pbcArrow.Visible = False
    lacSFrame(imSRIndex).Visible = False
    lacRFrame(imSRIndex).Visible = False
    gCtrlGotFocus ActiveControl
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
Private Sub pbcReplacement_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcReplacement_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If Button = 2 Then
        imButtonIndex = -1
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    ilCompRow = vbcSR.LargeChange + 1
    If UBound(tgRBofRec) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tgRBofRec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBRCtrls To UBound(tmRCtrls) Step 1
            If (X >= tmRCtrls(ilBox).fBoxX) And (X <= (tmRCtrls(ilBox).fBoxX + tmRCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCtrls(ilBox).fBoxY + tmRCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcSR.Value - 1
                    If ilRowNo > UBound(smRSave, 2) Then
                        Beep
                        mRSetFocus imRBoxNo
                        Exit Sub
                    End If
                    If (ilBox > RADVTINDEX) And smRSave(1, ilRowNo) = "" Then
                        Beep
                        mRSetFocus imRBoxNo
                        Exit Sub
                    End If
                    If smFromLog = "Y" Then
                        If (ilBox = RCNTRINDEX) And (smRSave(1, ilRowNo) = "[None]") Then
                            Beep
                            mRSetFocus imRBoxNo
                            Exit Sub
                        End If
                        If ((ilBox = RPPINDEX) Or (ilBox = RPPINDEX + 1)) And ((lmRSave(3, ilRowNo) > 0) Or (smSplitFill = "Y")) Then
                            Beep
                            mRSetFocus imRBoxNo
                            Exit Sub
                        End If
                    Else
                        If (ilBox = RSHORTTITLEINDEX) Then
                            Beep
                            mRSetFocus imRBoxNo
                            Exit Sub
                        End If
                        If ilBox >= RCARTINDEX Then
                            mCartPop ilRowNo
                        End If
                        If (ilBox >= RCARTINDEX) And (lbcCart.List(0) = "[None]") Then
                            ilBox = RADVTINDEX
                        End If
                    End If
                    mRSetShow imRBoxNo
                    imRRowNo = ilRow + vbcSR.Value - 1
                    imRBoxNo = ilBox
                    mREnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mRSetFocus imRBoxNo
End Sub
Private Sub pbcReplacement_Paint(Index As Integer)
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim slFont As String
    Dim llColor As Long

    'mPaintTitle 1
    mPaintBlackoutTitle
    ilStartRow = vbcSR.Value '+ 1  'Top location
    ilEndRow = vbcSR.Value + vbcSR.LargeChange ' + 1
    If ilEndRow > UBound(smRSave, 2) Then
        ilEndRow = UBound(smRSave, 2) 'include blank row as it might have data
    End If
    llColor = pbcReplacement(imSRIndex).ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBRCtrls To UBound(tmRCtrls) Step 1
            If ilRow = UBound(smRSave, 2) Then
                pbcReplacement(imSRIndex).ForeColor = DARKPURPLE
            Else
                pbcReplacement(imSRIndex).ForeColor = llColor
            End If
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(ilBox).fBoxX + fgBoxInsetX
            If (ilBox >= RDAYINDEX) And (ilBox <= RDAYINDEX + 6) And (imSRIndex = 0) Then
                pbcReplacement(imSRIndex).CurrentY = tmRCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) + 15 '+ fgBoxInsetY
                slFont = pbcReplacement(imSRIndex).FontName
                pbcReplacement(imSRIndex).FontName = "Monotype Sorts"
                pbcReplacement(imSRIndex).FontBold = False
                pbcReplacement(imSRIndex).Print smRShow(ilBox, ilRow)
                pbcReplacement(imSRIndex).FontName = slFont
                pbcReplacement(imSRIndex).FontBold = True
            Else
                pbcReplacement(imSRIndex).CurrentY = tmRCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                slStr = smRShow(ilBox, ilRow)
                pbcReplacement(imSRIndex).Print slStr
            End If
        Next ilBox
        pbcReplacement(imSRIndex).ForeColor = llColor
    Next ilRow
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    If igView = 0 Then
        Select Case imSBoxNo
            Case -1 'Tab from control prior to form area
                If (UBound(smSSave, 2) = 1) Then
                    imTabDirection = 0  'Set-Left to right
                    imSRowNo = 1
                    mInitNew imSRowNo
                Else
                    If UBound(smSSave, 2) <= vbcSR.LargeChange Then 'was <=
                        vbcSR.Max = LBound(smSSave, 2)
                    Else
                        vbcSR.Max = UBound(smSSave, 2) - vbcSR.LargeChange '- 1
                    End If
                    imSRowNo = 1
                    If imSRowNo >= UBound(smSSave, 2) Then
                        mInitNew imSRowNo
                    End If
                    imSettingValue = True
                    vbcSR.Value = vbcSR.Min
                    imSettingValue = False
                End If
                ilBox = SADVTINDEX
                imSBoxNo = ilBox
                mSEnableBox ilBox
                Exit Sub
            Case SADVTINDEX, 0
                mSSetShow imSBoxNo
                If (imSBoxNo < 1) And (imSRowNo < 1) Then 'Modelled from Proposal
                    Exit Sub
                End If
                ilBox = SENDTIMEINDEX
                If imSRowNo <= 1 Then
                    imSBoxNo = -1
                    imSRowNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imSRowNo = imSRowNo - 1
                If imSRowNo < vbcSR.Value Then
                    imSettingValue = True
                    vbcSR.Value = vbcSR.Value - 1
                    imSettingValue = False
                End If
                imSBoxNo = ilBox
                mSEnableBox ilBox
                Exit Sub
            Case SRADVTINDEX
                If Trim$(smSSave(1, imSRowNo)) = "[None]" Then
                    ilBox = SADVTINDEX
                Else
                    ilBox = imSBoxNo - 1
                End If
            Case SVEHINDEX
                If smFromLog = "Y" Then
                    If Trim$(smSSave(8, imSRowNo)) = "[None]" Then
                        ilBox = SRADVTINDEX
                    Else
                        ilBox = imSBoxNo - 1
                    End If
                Else
                    If Trim$(smSSave(1, imSRowNo)) = "[None]" Then
                        ilBox = SADVTINDEX
                    Else
                        ilBox = imSBoxNo - 1
                    End If
                End If
            Case SSTARTDATEINDEX
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
                ilBox = imSBoxNo - 1
            Case SENDDATEINDEX
                slStr = edcDropDown.Text
                If (slStr <> "") And (StrComp(slStr, "TFN", 1) <> 0) Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                End If
                ilBox = imSBoxNo - 1
            Case SSTARTTIMEINDEX
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
                ilBox = imSBoxNo - 1
            Case SENDTIMEINDEX
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
                ilBox = imSBoxNo - 1
            Case Else
                ilBox = imSBoxNo - 1
        End Select
        mSSetShow imSBoxNo
        imSBoxNo = ilBox
        mSEnableBox ilBox
    Else
        Select Case imRBoxNo
            Case -1 'Tab from control prior to form area
                If (UBound(smRSave, 2) = 1) Then
                    imTabDirection = 0  'Set-Left to right
                    imRRowNo = 1
                    mInitNew imRRowNo
                Else
                    If UBound(smRSave, 2) <= vbcSR.LargeChange Then 'was <=
                        vbcSR.Max = LBound(smRSave, 2)
                    Else
                        vbcSR.Max = UBound(smRSave, 2) - vbcSR.LargeChange '- 1
                    End If
                    imRRowNo = 1
                    If imRRowNo >= UBound(smRSave, 2) Then
                        mInitNew imRRowNo
                    End If
                    imSettingValue = True
                    vbcSR.Value = vbcSR.Min
                    imSettingValue = False
                End If
                ilBox = RADVTINDEX
                imRBoxNo = ilBox
                mREnableBox ilBox
                Exit Sub
            Case RADVTINDEX, 0
                mRSetShow imRBoxNo
                If (imRBoxNo < 1) And (imRRowNo < 1) Then 'Modelled from Proposal
                    Exit Sub
                End If
                ilBox = RENDTIMEINDEX
                If imRRowNo <= 1 Then
                    imRBoxNo = -1
                    imRRowNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imRRowNo = imRRowNo - 1
                If imRRowNo < vbcSR.Value Then
                    imSettingValue = True
                    vbcSR.Value = vbcSR.Value - 1
                    imSettingValue = False
                End If
                mCartPop imRRowNo
                If (lbcCart.List(0) = "[None]") And (imSRIndex = 0) Then
                    ilBox = RADVTINDEX
                End If
                imRBoxNo = ilBox
                mREnableBox ilBox
                Exit Sub
            Case RVEHINDEX
                If (lmRSave(3, imRRowNo) > 0) Or (smSplitFill = "Y") Then
                    ilBox = imRBoxNo - 3
                Else
                    ilBox = imRBoxNo - 1
                End If
            Case RSTARTDATEINDEX
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
                ilBox = imRBoxNo - 1
            Case RENDDATEINDEX
                slStr = edcDropDown.Text
                If (slStr <> "") And (StrComp(slStr, "TFN", 1) <> 0) Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                End If
                ilBox = imRBoxNo - 1
            Case RPPINDEX
                If smFromLog = "Y" Then
                    ilBox = imRBoxNo - 1
                Else
                    ilBox = imRBoxNo - 2    'Bypass short title
                End If
            Case RSTARTTIMEINDEX
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
                ilBox = imRBoxNo - 1
            Case RENDTIMEINDEX
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
                ilBox = imRBoxNo - 1
            Case Else
                ilBox = imRBoxNo - 1
        End Select
        mRSetShow imRBoxNo
        imRBoxNo = ilBox
        mREnableBox ilBox
    End If
End Sub
Private Sub pbcSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcSuppression_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcSuppression_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If Button = 2 Then
        imButtonIndex = -1
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    ilCompRow = vbcSR.LargeChange + 1
    If UBound(tgSBofRec) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tgSBofRec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBSCtrls To UBound(tmSCtrls) Step 1
            If (X >= tmSCtrls(ilBox).fBoxX) And (X <= (tmSCtrls(ilBox).fBoxX + tmSCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmSCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmSCtrls(ilBox).fBoxY + tmSCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcSR.Value - 1
                    If ilRowNo > UBound(smSSave, 2) Then
                        Beep
                        mSSetFocus imSBoxNo
                        Exit Sub
                    End If
                    If (ilBox > SADVTINDEX) And (smSSave(1, ilRowNo) = "") Then
                        Beep
                        mSSetFocus imSBoxNo
                        Exit Sub
                    End If
                    If smFromLog = "Y" Then
                        If (ilBox = SCNTRINDEX) And (smSSave(1, ilRowNo) = "[All]") Then
                            Beep
                            mSSetFocus imSBoxNo
                            Exit Sub
                        End If
                        If (ilBox = SRCNTRINDEX) And ((smSSave(8, ilRowNo) = "[None]") Or (smSSave(8, ilRowNo) = "")) Then
                            Beep
                            mSSetFocus imSBoxNo
                            Exit Sub
                        End If
                    Else
                        'If (ilBox = SSHORTTITLEINDEX) And (smSSave(1, ilRowNo) = "[None]") Then
                        '    Beep
                        '    mSSetFocus imSBoxNo
                        '    Exit Sub
                        'End If
                    End If
                    mSSetShow imSBoxNo
                    imSRowNo = ilRow + vbcSR.Value - 1
                    imSBoxNo = ilBox
                    mSEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mSSetFocus imSBoxNo
End Sub
Private Sub pbcSuppression_Paint(Index As Integer)
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim slFont As String
    Dim llColor As Long

    'mPaintTitle 0
    mPaintBlackoutTitle
    ilStartRow = vbcSR.Value '+ 1  'Top location
    ilEndRow = vbcSR.Value + vbcSR.LargeChange ' + 1
    If ilEndRow > UBound(smSSave, 2) Then
        ilEndRow = UBound(smSSave, 2) 'include blank row as it might have data
    End If
    llColor = pbcSuppression(imSRIndex).ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBSCtrls To UBound(tmSCtrls) Step 1
            If ilRow = UBound(smSSave, 2) Then
                pbcSuppression(imSRIndex).ForeColor = DARKPURPLE
            Else
                pbcSuppression(imSRIndex).ForeColor = llColor
            End If
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(ilBox).fBoxX + fgBoxInsetX
            If (ilBox >= SDAYINDEX) And (ilBox <= SDAYINDEX + 6) And (imSRIndex = 0) Then
                pbcSuppression(imSRIndex).CurrentY = tmSCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) + 15 '+ fgBoxInsetY
                slFont = pbcSuppression(imSRIndex).FontName
                pbcSuppression(imSRIndex).FontName = "Monotype Sorts"
                pbcSuppression(imSRIndex).FontBold = False
                pbcSuppression(imSRIndex).Print smSShow(ilBox, ilRow)
                pbcSuppression(imSRIndex).FontName = slFont
                pbcSuppression(imSRIndex).FontBold = True
            Else
                pbcSuppression(imSRIndex).CurrentY = tmSCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                slStr = smSShow(ilBox, ilRow)
                pbcSuppression(imSRIndex).Print slStr
            End If
        Next ilBox
        pbcSuppression(imSRIndex).ForeColor = llColor
    Next ilRow
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim slStr As String

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    If igView = 0 Then
        Select Case imSBoxNo
            Case -1 'Tab from control prior to form area
                imTabDirection = -1  'Set-Right to left
                imSRowNo = UBound(smSSave, 2) - 1
                imSettingValue = True
                If imSRowNo <= vbcSR.LargeChange + 1 Then
                    vbcSR.Value = 1
                Else
                    vbcSR.Value = imSRowNo - vbcSR.LargeChange - 1
                End If
                imSettingValue = False
                ilBox = SENDTIMEINDEX
            Case SADVTINDEX
                slStr = edcDropDown.Text
                If slStr = "[None]" Then
                    If smFromLog = "Y" Then
                        ilBox = SRADVTINDEX
                    Else
                        ilBox = SVEHINDEX
                    End If
                Else
                    ilBox = imSBoxNo + 1
                End If
            Case SRADVTINDEX
                slStr = edcDropDown.Text
                If slStr = "[None]" Then
                    ilBox = SVEHINDEX
                Else
                    ilBox = imSBoxNo + 1
                End If
            Case SRCNTRINDEX
                slStr = edcDropDown.Text
                If slStr = "" Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                Else
                    ilBox = imSBoxNo + 1
                End If
            Case SSTARTDATEINDEX
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
                ilBox = imSBoxNo + 1
            Case SENDDATEINDEX
                slStr = edcDropDown.Text
                If (slStr <> "") And (StrComp(slStr, "TFN", 1) <> 0) Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                End If
                ilBox = imSBoxNo + 1
            Case SSTARTTIMEINDEX
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
                ilBox = imSBoxNo + 1
            Case SENDTIMEINDEX
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
                mSSetShow imSBoxNo
                If mSTestSaveFields(imSRowNo, False) = NO Then
                    mSEnableBox imSBoxNo
                    Exit Sub
                End If
                If imSRowNo >= UBound(smSSave, 2) Then
                    imSChg = True
'                    ReDim Preserve smSShow(1 To 14, 1 To imSRowNo + 1) As String 'Values shown in program area
'                    ReDim Preserve smSSave(1 To 12, 1 To imSRowNo + 1) As String 'Values saved (program name) in program area
'                    ReDim Preserve imSSave(1 To 7, 1 To imSRowNo + 1) As Integer 'Values saved (program name) in program area
'                    ReDim Preserve lmSSave(1 To 4, 1 To imSRowNo + 1) As Long 'Values saved (program name) in program area
                    ReDim Preserve smSShow(0 To 14, 0 To imSRowNo + 1) As String 'Values shown in program area
                    ReDim Preserve smSSave(0 To 12, 0 To imSRowNo + 1) As String 'Values saved (program name) in program area
                    ReDim Preserve imSSave(0 To 7, 0 To imSRowNo + 1) As Integer 'Values saved (program name) in program area
                    ReDim Preserve lmSSave(0 To 4, 0 To imSRowNo + 1) As Long 'Values saved (program name) in program area

                    For ilLoop = LBound(smSShow, 1) To UBound(smSShow, 1) Step 1
                        smSShow(ilLoop, imSRowNo + 1) = ""
                    Next ilLoop
                    For ilLoop = LBound(smSSave, 1) To UBound(smSSave, 1) Step 1
                        smSSave(ilLoop, imSRowNo + 1) = ""
                    Next ilLoop
                    For ilLoop = LBound(imSSave, 1) To UBound(imSSave, 1) Step 1
                        imSSave(ilLoop, imSRowNo + 1) = -1
                    Next ilLoop
                    For ilLoop = LBound(lmSSave, 1) To UBound(lmSSave, 1) Step 1
                        lmSSave(ilLoop, imSRowNo + 1) = -1
                    Next ilLoop
                    'ReDim Preserve tgSBofRec(1 To UBound(tgSBofRec) + 1) As BOFREC
                    ReDim Preserve tgSBofRec(0 To UBound(tgSBofRec) + 1) As BOFREC
                    tgSBofRec(UBound(tgSBofRec)).iStatus = 0
                    tgSBofRec(UBound(tgSBofRec)).lRecPos = 0
                End If
                If imSRowNo >= UBound(smSSave, 2) - 1 Then
                    imSRowNo = imSRowNo + 1
                    mInitNew imSRowNo
                    If UBound(smSSave, 2) <= vbcSR.LargeChange Then 'was <=
                        vbcSR.Max = LBound(smSSave, 2) '- 1
                    Else
                        vbcSR.Max = UBound(smSSave, 2) - vbcSR.LargeChange '- 1
                    End If
                Else
                    imSRowNo = imSRowNo + 1
                End If
                If imSRowNo > vbcSR.Value + vbcSR.LargeChange Then
                    imSettingValue = True
                    vbcSR.Value = vbcSR.Value + 1
                    imSettingValue = False
                End If
                If imSRowNo >= UBound(smSSave, 2) Then
                    imSBoxNo = 0
                    mSetCommands
                    'lacFrame.Move 0, tmCtrls(PROPNOINDEX).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15) - 30
                    'lacFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcBlackout.Top + tmSCtrls(SADVTINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                    Exit Sub
                Else
                    ilBox = SADVTINDEX
                End If
                imSBoxNo = ilBox
                mSEnableBox ilBox
                Exit Sub
            Case 0
                ilBox = SADVTINDEX
            Case Else
                ilBox = imSBoxNo + 1
        End Select
        mSSetShow imSBoxNo
        imSBoxNo = ilBox
        mSEnableBox ilBox
    Else
        Select Case imRBoxNo
            Case -1 'Tab from control prior to form area
                imTabDirection = -1  'Set-Right to left
                imRRowNo = UBound(smRSave, 2) - 1
                imSettingValue = True
                If imRRowNo <= vbcSR.LargeChange + 1 Then
                    vbcSR.Value = 1
                Else
                    vbcSR.Value = imRRowNo - vbcSR.LargeChange - 1
                End If
                imSettingValue = False
                ilBox = RENDTIMEINDEX
            Case RCARTINDEX
                If (lbcCart.List(0) = "[None]") And (imSRIndex = 0) Then
                    ilBox = RADVTINDEX
                Else
                    'Go to Product protection
                    If smFromLog = "Y" Then
                        If (lmRSave(3, imRRowNo) > 0) Or (smSplitFill = "Y") Then
                            ilBox = imRBoxNo + 3
                        Else
                            ilBox = imRBoxNo + 1
                        End If
                    Else
                        ilBox = imRBoxNo + 2    'Bypass short title
                    End If
                End If
            Case RSTARTDATEINDEX
                slStr = edcDropDown.Text
                If slStr <> "" Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
                ilBox = imRBoxNo + 1
            Case RENDDATEINDEX
                slStr = edcDropDown.Text
                If (slStr <> "") And (StrComp(slStr, "TFN", 1) <> 0) Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                End If
                ilBox = imRBoxNo + 1
            Case RSTARTTIMEINDEX
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
                ilBox = imRBoxNo + 1
            Case RENDTIMEINDEX
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
                mRSetShow imRBoxNo
                If mRTestSaveFields(imRRowNo, False) = NO Then
                    mREnableBox imRBoxNo
                    Exit Sub
                End If
                If imRRowNo >= UBound(smRSave, 2) Then
                    imRChg = True
'                    ReDim Preserve smRShow(1 To 16, 1 To imRRowNo + 1) As String 'Values shown in program area
'                    ReDim Preserve smRSave(1 To 11, 1 To imRRowNo + 1) As String 'Values saved (program name) in program area
'                    ReDim Preserve imRSave(1 To 7, 1 To imRRowNo + 1) As Integer 'Values saved (program name) in program area
'                    ReDim Preserve lmRSave(1 To 3, 1 To imRRowNo + 1) As Long 'Values saved (program name) in program area
                    ReDim Preserve smRShow(0 To 16, 0 To imRRowNo + 1) As String 'Values shown in program area
                    ReDim Preserve smRSave(0 To 11, 0 To imRRowNo + 1) As String 'Values saved (program name) in program area
                    ReDim Preserve imRSave(0 To 7, 0 To imRRowNo + 1) As Integer 'Values saved (program name) in program area
                    ReDim Preserve lmRSave(0 To 3, 0 To imRRowNo + 1) As Long 'Values saved (program name) in program area

                    For ilLoop = LBound(smRShow, 1) To UBound(smRShow, 1) Step 1
                        smRShow(ilLoop, imRRowNo + 1) = ""
                    Next ilLoop
                    For ilLoop = LBound(smRSave, 1) To UBound(smRSave, 1) Step 1
                        smRSave(ilLoop, imRRowNo + 1) = ""
                    Next ilLoop
                    For ilLoop = LBound(imRSave, 1) To UBound(imRSave, 1) Step 1
                        imRSave(ilLoop, imRRowNo + 1) = -1
                    Next ilLoop
                    'ReDim Preserve tgRBofRec(1 To UBound(tgRBofRec) + 1) As BOFREC
                    ReDim Preserve tgRBofRec(0 To UBound(tgRBofRec) + 1) As BOFREC
                    tgRBofRec(UBound(tgRBofRec)).iStatus = 0
                    tgRBofRec(UBound(tgRBofRec)).lRecPos = 0
                End If
                If imRRowNo >= UBound(smRSave, 2) - 1 Then
                    imRRowNo = imRRowNo + 1
                    mInitNew imRRowNo
                    If UBound(smRSave, 2) <= vbcSR.LargeChange Then 'was <=
                        vbcSR.Max = LBound(smRSave, 2) '- 1
                    Else
                        vbcSR.Max = UBound(smRSave, 2) - vbcSR.LargeChange '- 1
                    End If
                Else
                    imRRowNo = imRRowNo + 1
                End If
                If imRRowNo > vbcSR.Value + vbcSR.LargeChange Then
                    imSettingValue = True
                    vbcSR.Value = vbcSR.Value + 1
                    imSettingValue = False
                End If
                If imRRowNo >= UBound(smRSave, 2) Then
                    imRBoxNo = 0
                    mSetCommands
                    'lacFrame.Move 0, tmCtrls(PROPNOINDEX).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15) - 30
                    'lacFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcBlackout.Top + tmRCtrls(RADVTINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                    Exit Sub
                Else
                    ilBox = RADVTINDEX
                End If
                imRBoxNo = ilBox
                mREnableBox ilBox
                Exit Sub
            Case 0
                ilBox = SADVTINDEX
            Case Else
                ilBox = imRBoxNo + 1
        End Select
        mRSetShow imRBoxNo
        imRBoxNo = ilBox
        mREnableBox ilBox
    End If
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
                    If igView = 0 Then
                        Select Case imSBoxNo
                            Case SSTARTTIMEINDEX
                                imBypassFocus = True    'Don't change select text
                                edcDropDown.SetFocus
                                'SendKeys slKey
                                gSendKeys edcDropDown, slKey
                            Case SENDTIMEINDEX
                                imBypassFocus = True    'Don't change select text
                                edcDropDown.SetFocus
                                'SendKeys slKey
                                gSendKeys edcDropDown, slKey
                        End Select
                    Else
                        Select Case imRBoxNo
                            Case RSTARTTIMEINDEX
                                imBypassFocus = True    'Don't change select text
                                edcDropDown.SetFocus
                                'SendKeys slKey
                                gSendKeys edcDropDown, slKey
                            Case RENDTIMEINDEX
                                imBypassFocus = True    'Don't change select text
                                edcDropDown.SetFocus
                                'SendKeys slKey
                                gSendKeys edcDropDown, slKey
                        End Select
                    End If
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcSelect_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub rbcSR_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSR(Index).Value
    'End of coded added
    If Value Then
        If Index = 0 Then
            igView = 0
            'mClearCtrlFields
            pbcSuppression(imSRIndex).Visible = True
            pbcReplacement(imSRIndex).Visible = False
        Else
            igView = 1
            'mClearCtrlFields
            pbcReplacement(imSRIndex).Visible = True
            pbcSuppression(imSRIndex).Visible = False
        End If
        mSetMinMax
    End If
End Sub
Private Sub rbcSR_GotFocus(Index As Integer)
    If imFirstTime Then 'Test if coming from sales source- if so, branch to first control
        imFirstTime = False
    End If
    mSSetShow imSBoxNo
    imSBoxNo = -1
    imSRowNo = -1
    mRSetShow imRBoxNo
    imRBoxNo = -1
    imRRowNo = -1
    pbcArrow.Visible = False
    lacSFrame(imSRIndex).Visible = False
    lacRFrame(imSRIndex).Visible = False
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If igView = 0 Then
        If smFromLog = "Y" Then
            Select Case imSBoxNo
                Case SADVTINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcSAdvt, edcDropDown, imChgMode, imLbcArrowSetting
                Case SCNTRINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcSCntr, edcDropDown, imChgMode, imLbcArrowSetting
                Case SRADVTINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcSAdvt, edcDropDown, imChgMode, imLbcArrowSetting
                Case SRCNTRINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcSRCntr, edcDropDown, imChgMode, imLbcArrowSetting
                Case SRCARTINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcCart, edcDropDown, imChgMode, imLbcArrowSetting
                Case SVEHINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
                Case SLENINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcLen, edcDropDown, imChgMode, imLbcArrowSetting
                Case SDAYINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcDays, edcDropDown, imChgMode, imLbcArrowSetting
            End Select
        Else
            Select Case imSBoxNo
                Case SADVTINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcSAdvt, edcDropDown, imChgMode, imLbcArrowSetting
                Case SSHORTTITLEINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcShtTitle, edcDropDown, imChgMode, imLbcArrowSetting
                Case SVEHINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
            End Select
        End If
    Else
        If smFromLog = "Y" Then
            Select Case imRBoxNo
                Case RADVTINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcRAdvt, edcDropDown, imChgMode, imLbcArrowSetting
                Case RCNTRINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcRCntr, edcDropDown, imChgMode, imLbcArrowSetting
                Case RCARTINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcCart, edcDropDown, imChgMode, imLbcArrowSetting
                Case RVEHINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
                Case RPPINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcComp(0), edcDropDown, imChgMode, imLbcArrowSetting
                Case RPPINDEX + 1
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcComp(1), edcDropDown, imChgMode, imLbcArrowSetting
                Case RDAYINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcDays, edcDropDown, imChgMode, imLbcArrowSetting
            End Select
        Else
            Select Case imRBoxNo
                Case RADVTINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcRAdvt, edcDropDown, imChgMode, imLbcArrowSetting
                Case RCARTINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcCart, edcDropDown, imChgMode, imLbcArrowSetting
                Case RSHORTTITLEINDEX
                Case RPPINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcComp(0), edcDropDown, imChgMode, imLbcArrowSetting
                Case RPPINDEX + 1
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcComp(1), edcDropDown, imChgMode, imLbcArrowSetting
            End Select
        End If
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
            ilCompRow = vbcSR.LargeChange + 1
            If igView = 0 Then
                If UBound(smSSave, 2) > ilCompRow Then
                    ilMaxRow = ilCompRow
                Else
                    ilMaxRow = UBound(smSSave, 2)
                End If
                For ilRow = 1 To ilMaxRow Step 1
                    If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmSCtrls(SADVTINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmSCtrls(SADVTINDEX).fBoxY + tmSCtrls(SADVTINDEX).fBoxH)) Then
                        mSSetShow imSBoxNo
                        imSBoxNo = -1
                        imSRowNo = -1
                        imSRowNo = ilRow + vbcSR.Value - 1
                        lacSFrame(imSRIndex).DragIcon = IconTraf!imcIconStd.DragIcon
                        lacSFrame(imSRIndex).Move 0, tmSCtrls(SADVTINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15) - 30
                        'If gInvertArea call then remove visible setting
                        lacSFrame(imSRIndex).Visible = True
                        pbcArrow.Move pbcArrow.Left, plcBlackout.Top + tmSCtrls(SADVTINDEX).fBoxY + (imSRowNo - vbcSR.Value) * (fgBoxGridH + 15) + 45
                        pbcArrow.Visible = True
                        imcTrash.Enabled = True
                        lacSFrame(imSRIndex).Drag vbBeginDrag
                        lacSFrame(imSRIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
                        Exit Sub
                    End If
                Next ilRow
            Else
                If UBound(smRSave, 2) > ilCompRow Then
                    ilMaxRow = ilCompRow
                Else
                    ilMaxRow = UBound(smRSave, 2)
                End If
                For ilRow = 1 To ilMaxRow Step 1
                    If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCtrls(RADVTINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCtrls(RADVTINDEX).fBoxY + tmRCtrls(RADVTINDEX).fBoxH)) Then
                        mRSetShow imRBoxNo
                        imRBoxNo = -1
                        imRRowNo = -1
                        imRRowNo = ilRow + vbcSR.Value - 1
                        lacRFrame(imSRIndex).DragIcon = IconTraf!imcIconStd.DragIcon
                        lacRFrame(imSRIndex).Move 0, tmRCtrls(RADVTINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15) - 30
                        'If gInvertArea call then remove visible setting
                        lacRFrame(imSRIndex).Visible = True
                        pbcArrow.Move pbcArrow.Left, plcBlackout.Top + tmRCtrls(RADVTINDEX).fBoxY + (imRRowNo - vbcSR.Value) * (fgBoxGridH + 15) + 45
                        pbcArrow.Visible = True
                        imcTrash.Enabled = True
                        lacRFrame(imSRIndex).Drag vbBeginDrag
                        lacRFrame(imSRIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
                        Exit Sub
                    End If
                Next ilRow
            End If
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If Not imTerminate Then
        mMoveRecToCtrl "B"
        'Add clear since auto select removed
        'mClearCtrlFields
        'Remove auto select until getting data is faster
        'If cbcSelect.ListCount <= 1 Then
        '    cbcSelect.ListIndex = 0 'This will generate a select_change event
        'Else
        '    cbcSelect.ListIndex = 1
        'End If
        'mSetCommands
    End If
End Sub
Private Sub vbcSR_Change()
    If imSettingValue Then
        If igView = 0 Then
            pbcSuppression(imSRIndex).Cls
            pbcSuppression_Paint imSRIndex
        Else
            pbcReplacement(imSRIndex).Cls
            pbcReplacement_Paint imSRIndex
        End If
        imSettingValue = False
    Else
        If igView = 0 Then
            mSSetShow imSBoxNo
            pbcSuppression(imSRIndex).Cls
            pbcSuppression_Paint imSRIndex
            mSEnableBox imSBoxNo
        Else
            mRSetShow imRBoxNo
            pbcReplacement(imSRIndex).Cls
            pbcReplacement_Paint imSRIndex
            mREnableBox imRBoxNo
        End If
    End If
End Sub
Private Sub vbcSR_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    imSRowNo = -1
    mRSetShow imRBoxNo
    imRBoxNo = -1
    imRRowNo = -1
    pbcArrow.Visible = False
    lacSFrame(imSRIndex).Visible = False
    lacRFrame(imSRIndex).Visible = False
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    If smSplitFill = "Y" Then
        plcScreen.Print "Split Fill"
    Else
        plcScreen.Print "Blackout"
    End If
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
Private Sub mPaintBlackoutTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer

    If rbcSR(0).Value = True Then
        llColor = pbcSuppression(imSRIndex).ForeColor
        slFontName = pbcSuppression(imSRIndex).FontName
        flFontSize = pbcSuppression(imSRIndex).FontSize
        ilFillStyle = pbcSuppression(imSRIndex).FillStyle
        llFillColor = pbcSuppression(imSRIndex).FillColor
        ilHalfY = tmSCtrls(SADVTINDEX).fBoxY / 2
        Do While ilHalfY Mod 15 <> 0
            ilHalfY = ilHalfY - 1
        Loop
        pbcSuppression(imSRIndex).ForeColor = BLUE
        pbcSuppression(imSRIndex).FontBold = False
        pbcSuppression(imSRIndex).FontSize = 7
        pbcSuppression(imSRIndex).FontName = "Arial"
        pbcSuppression(imSRIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        If imSRIndex = 1 Then
            pbcSuppression(imSRIndex).Line (tmSCtrls(SADVTINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SADVTINDEX).fBoxW + tmSCtrls(SCNTRINDEX).fBoxW + 30, tmSCtrls(SADVTINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SADVTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Suppress"
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SADVTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = ilHalfY
            pbcSuppression(imSRIndex).Print "Advertiser"
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SCNTRINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = ilHalfY
            pbcSuppression(imSRIndex).Print "Contract"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SRADVTINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SRADVTINDEX).fBoxW + tmSCtrls(SRCNTRINDEX).fBoxW + tmSCtrls(SRCARTINDEX).fBoxW + 45, tmSCtrls(SADVTINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SRADVTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Replace"
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SRADVTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = ilHalfY
            pbcSuppression(imSRIndex).Print "Advertiser"
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SRCNTRINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = ilHalfY
            pbcSuppression(imSRIndex).Print "Contract"
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SRCARTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = ilHalfY
            pbcSuppression(imSRIndex).Print "Copy"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SVEHINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SVEHINDEX).fBoxW + 15, tmSCtrls(SVEHINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SVEHINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Vehicle"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SLENINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SLENINDEX).fBoxW + 15, tmSCtrls(SLENINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SLENINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Len"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SSTARTDATEINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SSTARTDATEINDEX).fBoxW + 15, tmSCtrls(SSTARTDATEINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SSTARTDATEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Start Date"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SENDDATEINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SENDDATEINDEX).fBoxW + 15, tmSCtrls(SENDDATEINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SENDDATEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "End Date"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SDAYINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SDAYINDEX).fBoxW + 15, tmSCtrls(SDAYINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SDAYINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Days"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SSTARTTIMEINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SSTARTTIMEINDEX).fBoxW + 15, tmSCtrls(SSTARTTIMEINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SSTARTTIMEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Start Time"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SENDTIMEINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SENDTIMEINDEX).fBoxW + 15, tmSCtrls(SENDTIMEINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SENDTIMEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "End Time"
        Else
            pbcSuppression(imSRIndex).Line (tmSCtrls(SADVTINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SADVTINDEX).fBoxW + 15, tmSCtrls(SADVTINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SADVTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Advertiser"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SSHORTTITLEINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SSHORTTITLEINDEX).fBoxW + 15, tmSCtrls(SSHORTTITLEINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SSHORTTITLEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            If tgSpf.sUseProdSptScr = "P" Then  'Short Title
                pbcSuppression(imSRIndex).Print "Short Title"
            Else
                pbcSuppression(imSRIndex).Print "Product"
            End If
            pbcSuppression(imSRIndex).Line (tmSCtrls(SVEHINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SVEHINDEX).fBoxW + 15, tmSCtrls(SVEHINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SVEHINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Vehicle"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SSTARTDATEINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SSTARTDATEINDEX).fBoxW + 15, tmSCtrls(SSTARTDATEINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SSTARTDATEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Start Date"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SENDDATEINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SENDDATEINDEX).fBoxW + 15, tmSCtrls(SENDDATEINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SENDDATEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "End Date"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SDAYINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SSTARTTIMEINDEX).fBoxX - tmSCtrls(SDAYINDEX).fBoxW + 15, tmSCtrls(SDAYINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SDAYINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Mo"
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SDAYINDEX + 1).fBoxX + 15 'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Tu"
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SDAYINDEX + 2).fBoxX + 15 'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "We"
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SDAYINDEX + 3).fBoxX + 15 'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Th"
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SDAYINDEX + 4).fBoxX + 15 'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Fr"
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SDAYINDEX + 5).fBoxX + 15 'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Sa"
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SDAYINDEX + 6).fBoxX + 15 'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Su"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SSTARTTIMEINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SSTARTTIMEINDEX).fBoxW + 15, tmSCtrls(SSTARTTIMEINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SSTARTTIMEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "Start Time"
            pbcSuppression(imSRIndex).Line (tmSCtrls(SENDTIMEINDEX).fBoxX - 15, 15)-Step(tmSCtrls(SENDTIMEINDEX).fBoxW + 15, tmSCtrls(SENDTIMEINDEX).fBoxY - 30), BLUE, B
            pbcSuppression(imSRIndex).CurrentX = tmSCtrls(SENDTIMEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSuppression(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcSuppression(imSRIndex).Print "End Time"
        End If
        ilLineCount = 0
        llTop = tmSCtrls(1).fBoxY
        Do
            For ilLoop = imLBSCtrls To UBound(tmSCtrls) Step 1
                pbcSuppression(imSRIndex).Line (tmSCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmSCtrls(ilLoop).fBoxW + 15, tmSCtrls(ilLoop).fBoxH + 15), BLUE, B
            Next ilLoop
            ilLineCount = ilLineCount + 1
            llTop = llTop + tmSCtrls(1).fBoxH + 15
        Loop While llTop + tmSCtrls(1).fBoxH < pbcSuppression(imSRIndex).height
        vbcSR.LargeChange = ilLineCount - 1
        pbcSuppression(imSRIndex).FontSize = flFontSize
        pbcSuppression(imSRIndex).FontName = slFontName
        pbcSuppression(imSRIndex).FontSize = flFontSize
        pbcSuppression(imSRIndex).ForeColor = llColor
        pbcSuppression(imSRIndex).FontBold = True
    Else
        llColor = pbcReplacement(imSRIndex).ForeColor
        slFontName = pbcReplacement(imSRIndex).FontName
        flFontSize = pbcReplacement(imSRIndex).FontSize
        ilFillStyle = pbcReplacement(imSRIndex).FillStyle
        llFillColor = pbcReplacement(imSRIndex).FillColor
        ilHalfY = tmRCtrls(SADVTINDEX).fBoxY / 2
        Do While ilHalfY Mod 15 <> 0
            ilHalfY = ilHalfY - 1
        Loop
        pbcReplacement(imSRIndex).ForeColor = BLUE
        pbcReplacement(imSRIndex).FontBold = False
        pbcReplacement(imSRIndex).FontSize = 7
        pbcReplacement(imSRIndex).FontName = "Arial"
        pbcReplacement(imSRIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        If imSRIndex = 1 Then
            pbcReplacement(imSRIndex).Line (tmRCtrls(RADVTINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RADVTINDEX).fBoxW + 15, tmRCtrls(RADVTINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RADVTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Advertiser"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RCNTRINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RCNTRINDEX).fBoxW + 15, tmRCtrls(RCNTRINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RCNTRINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Contract"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RCARTINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RCARTINDEX).fBoxW + 15, tmRCtrls(RCARTINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RCARTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Cart"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RPPINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RPPINDEX).fBoxW + tmRCtrls(RPPINDEX + 1).fBoxW + 30, tmRCtrls(RPPINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RPPINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Product"
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RPPINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = ilHalfY
            pbcReplacement(imSRIndex).Print "Protection"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RVEHINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RVEHINDEX).fBoxW + 15, tmRCtrls(RVEHINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RVEHINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Vehicle"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RSTARTDATEINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RSTARTDATEINDEX).fBoxW + 15, tmRCtrls(RSTARTDATEINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RSTARTDATEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Start Date"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RENDDATEINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RENDDATEINDEX).fBoxW + 15, tmRCtrls(RENDDATEINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RENDDATEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "End Date"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RDAYINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RDAYINDEX).fBoxW + 15, tmRCtrls(RDAYINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RDAYINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Days"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RSTARTTIMEINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RSTARTTIMEINDEX).fBoxW + 15, tmRCtrls(RSTARTTIMEINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RSTARTTIMEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Start Time"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RENDTIMEINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RENDTIMEINDEX).fBoxW + 15, tmRCtrls(RENDTIMEINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RENDTIMEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "End Time"
        Else
            pbcReplacement(imSRIndex).Line (tmRCtrls(RADVTINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RADVTINDEX).fBoxW + 15, tmRCtrls(RADVTINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RADVTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Advertiser"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RCARTINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RCARTINDEX).fBoxW + 15, tmRCtrls(RCARTINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RCARTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Cart"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RSHORTTITLEINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RSHORTTITLEINDEX).fBoxW + 15, tmRCtrls(RSHORTTITLEINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).Line (tmRCtrls(RSHORTTITLEINDEX).fBoxX, 30)-Step(tmRCtrls(RSHORTTITLEINDEX).fBoxW - 15, tmRCtrls(RSHORTTITLEINDEX).fBoxY - 60), LIGHTYELLOW, BF
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RSHORTTITLEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            If tgSpf.sUseProdSptScr = "P" Then  'Short Title
                pbcReplacement(imSRIndex).Print "Short Title"
            Else
                pbcReplacement(imSRIndex).Print "Product"
            End If
            pbcReplacement(imSRIndex).Line (tmRCtrls(RPPINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RPPINDEX).fBoxW + tmRCtrls(RPPINDEX + 1).fBoxW + 30, tmRCtrls(RPPINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RPPINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Product Protection"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RSTARTDATEINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RSTARTDATEINDEX).fBoxW + 15, tmRCtrls(RSTARTDATEINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RSTARTDATEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Start Date"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RENDDATEINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RENDDATEINDEX).fBoxW + 15, tmRCtrls(RENDDATEINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RENDDATEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "End Date"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RDAYINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RSTARTTIMEINDEX).fBoxX - tmRCtrls(RDAYINDEX).fBoxW + 15, tmRCtrls(RDAYINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RDAYINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Mo"
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RDAYINDEX + 1).fBoxX + 15 'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Tu"
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RDAYINDEX + 2).fBoxX + 15 'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "We"
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RDAYINDEX + 3).fBoxX + 15 'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Th"
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RDAYINDEX + 4).fBoxX + 15 'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Fr"
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RDAYINDEX + 5).fBoxX + 15 'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Sa"
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RDAYINDEX + 6).fBoxX + 15 'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Su"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RSTARTTIMEINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RSTARTTIMEINDEX).fBoxW + 15, tmRCtrls(RSTARTTIMEINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RSTARTTIMEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "Start Time"
            pbcReplacement(imSRIndex).Line (tmRCtrls(RENDTIMEINDEX).fBoxX - 15, 15)-Step(tmRCtrls(RENDTIMEINDEX).fBoxW + 15, tmRCtrls(RENDTIMEINDEX).fBoxY - 30), BLUE, B
            pbcReplacement(imSRIndex).CurrentX = tmRCtrls(RENDTIMEINDEX).fBoxX + 15  'fgBoxInsetX
            pbcReplacement(imSRIndex).CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcReplacement(imSRIndex).Print "End Time"
        End If
        ilLineCount = 0
        llTop = tmRCtrls(1).fBoxY
        Do
            For ilLoop = imLBRCtrls To UBound(tmRCtrls) Step 1
                If (imSRIndex <> 1) And (ilLoop = RSHORTTITLEINDEX) Then
                    pbcReplacement(imSRIndex).FillStyle = 0 'Solid
                    pbcReplacement(imSRIndex).FillColor = LIGHTYELLOW
                End If
                pbcReplacement(imSRIndex).Line (tmRCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmRCtrls(ilLoop).fBoxW + 15, tmRCtrls(ilLoop).fBoxH + 15), BLUE, B
                If (imSRIndex <> 1) And (ilLoop = RSHORTTITLEINDEX) Then
                    pbcReplacement(imSRIndex).FillStyle = ilFillStyle
                    pbcReplacement(imSRIndex).FillColor = llFillColor
                End If
            Next ilLoop
            ilLineCount = ilLineCount + 1
            llTop = llTop + tmRCtrls(1).fBoxH + 15
        Loop While llTop + tmRCtrls(1).fBoxH < pbcReplacement(imSRIndex).height
        vbcSR.LargeChange = ilLineCount - 1
        pbcReplacement(imSRIndex).FontSize = flFontSize
        pbcReplacement(imSRIndex).FontName = slFontName
        pbcReplacement(imSRIndex).FontSize = flFontSize
        pbcReplacement(imSRIndex).ForeColor = llColor
        pbcReplacement(imSRIndex).FontBold = True
    End If
End Sub
