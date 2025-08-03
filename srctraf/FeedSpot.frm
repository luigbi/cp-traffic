VERSION 5.00
Begin VB.Form FeedSpot 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
   ClientLeft      =   855
   ClientTop       =   1470
   ClientWidth     =   9465
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   9465
   Begin VB.CommandButton cmcPledge 
      Appearance      =   0  'Flat
      Caption         =   "&Pledge"
      Height          =   285
      Left            =   6750
      TabIndex        =   48
      Top             =   5250
      Width           =   1050
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   8070
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   3525
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "FeedSpot.frx":0000
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   47
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
            Picture         =   "FeedSpot.frx":0CBE
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   6
      Left            =   9330
      Top             =   4995
   End
   Begin VB.CheckBox ckcAirDay 
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
      Left            =   7890
      TabIndex        =   19
      Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
      Top             =   3690
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.ListBox lbcDW 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "FeedSpot.frx":0FC8
      Left            =   1245
      List            =   "FeedSpot.frx":0FCF
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ListBox lbcRun 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "FeedSpot.frx":0FDA
      Left            =   3975
      List            =   "FeedSpot.frx":0FE1
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcLen 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "FeedSpot.frx":0FED
      Left            =   3600
      List            =   "FeedSpot.frx":0FF4
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3150
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lbcComp 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   1
      ItemData        =   "FeedSpot.frx":1000
      Left            =   2805
      List            =   "FeedSpot.frx":1007
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2625
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcComp 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      ItemData        =   "FeedSpot.frx":1014
      Left            =   1605
      List            =   "FeedSpot.frx":101B
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcCart 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   7860
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3225
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ListBox lbcProd 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   750
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3420
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.ListBox lbcAdvt 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   720
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3060
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9090
      Top             =   5490
   End
   Begin VB.PictureBox plcNum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   4005
      ScaleHeight     =   1140
      ScaleWidth      =   1095
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3660
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcNum 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1050
         Left            =   45
         Picture         =   "FeedSpot.frx":1028
         ScaleHeight     =   1050
         ScaleWidth      =   1020
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcNumOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            Top             =   15
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcNumInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   300
            Picture         =   "FeedSpot.frx":1B9A
            Top             =   255
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox pbcArrow 
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
      Height          =   180
      Left            =   60
      Picture         =   "FeedSpot.frx":1EA4
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1245
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1095
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3915
      Visible         =   0   'False
      Width           =   2670
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
      Left            =   2040
      Picture         =   "FeedSpot.frx":21AE
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
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
      Left            =   1020
      MaxLength       =   20
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1020
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
      Left            =   5460
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   23
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
         Left            =   1635
         TabIndex        =   28
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
         Left            =   45
         TabIndex        =   26
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
         Left            =   45
         Picture         =   "FeedSpot.frx":22A8
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   29
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
            TabIndex        =   30
            Top             =   405
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
         TabIndex        =   27
         Top             =   45
         Width           =   1305
      End
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
      Left            =   9240
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5025
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
      Left            =   9195
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   5055
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
      Left            =   9225
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5295
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Height          =   285
      Left            =   5655
      TabIndex        =   38
      Top             =   5250
      Width           =   1050
   End
   Begin VB.CommandButton cmcSchedule 
      Appearance      =   0  'Flat
      Caption         =   "&Schedule"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   37
      Top             =   5250
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
      Left            =   -15
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4560
      Width           =   75
   End
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   3570
      TabIndex        =   36
      Top             =   5250
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
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   32
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
      TabIndex        =   6
      Top             =   525
      Width           =   105
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2475
      TabIndex        =   35
      Top             =   5250
      Width           =   1050
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   975
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1380
      TabIndex        =   34
      Top             =   5250
      Width           =   1050
   End
   Begin VB.PictureBox pbcFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   4080
      Index           =   2
      Left            =   75
      Picture         =   "FeedSpot.frx":50C2
      ScaleHeight     =   4080
      ScaleWidth      =   8940
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   8940
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
         Index           =   2
         Left            =   15
         TabIndex        =   45
         Top             =   585
         Visible         =   0   'False
         Width           =   8910
      End
   End
   Begin VB.PictureBox pbcFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   4080
      Index           =   1
      Left            =   135
      Picture         =   "FeedSpot.frx":81E8C
      ScaleHeight     =   4080
      ScaleWidth      =   8940
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1155
      Visible         =   0   'False
      Width           =   8940
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
         Index           =   1
         Left            =   15
         TabIndex        =   44
         Top             =   585
         Visible         =   0   'False
         Width           =   8910
      End
   End
   Begin VB.PictureBox pbcFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   4080
      Index           =   0
      Left            =   195
      Picture         =   "FeedSpot.frx":FF0CE
      ScaleHeight     =   4080
      ScaleWidth      =   8940
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   660
      Width           =   8940
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
         Index           =   0
         Left            =   15
         TabIndex        =   43
         Top             =   585
         Visible         =   0   'False
         Width           =   8910
      End
   End
   Begin VB.VScrollBar vbcFeed 
      Height          =   4110
      LargeChange     =   18
      Left            =   9150
      Max             =   1
      Min             =   1
      TabIndex        =   33
      Top             =   675
      Value           =   1
      Width           =   240
   End
   Begin VB.PictureBox plcFeed 
      BackColor       =   &H00FFFFFF&
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
      Height          =   4215
      Left            =   180
      ScaleHeight     =   4155
      ScaleWidth      =   9180
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   615
      Width           =   9240
   End
   Begin VB.PictureBox plcSelect 
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
      Height          =   390
      Left            =   1170
      ScaleHeight     =   330
      ScaleWidth      =   8175
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   15
      Width           =   8235
      Begin VB.CommandButton cmcDate 
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
         Left            =   7965
         Picture         =   "FeedSpot.frx":17C310
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   60
         Width           =   195
      End
      Begin VB.TextBox edcDate 
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
         Left            =   6705
         MaxLength       =   10
         TabIndex        =   4
         Top             =   60
         Width           =   1245
      End
      Begin VB.ComboBox cbcVehicle 
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
         Left            =   3180
         TabIndex        =   3
         Top             =   0
         Width           =   3195
      End
      Begin VB.ComboBox cbcFeed 
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
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.Label lacType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   195
      TabIndex        =   49
      Top             =   480
      Width           =   2970
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8595
      Picture         =   "FeedSpot.frx":17C40A
      Top             =   5115
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   180
      Top             =   5145
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "FeedSpot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of FeedSpot.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: FeedSpot.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Commission input screen code


'Save disallowed if Copy defined
'Code placed into mTestFields at bottom


Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim tmCtrls()  As FIELDAREA
Dim tmLCtrls(0 To 10)  As FIELDAREA
Dim imLBLCtrls As Integer
Dim tmPCtrls(0 To 12)  As FIELDAREA
Dim imLBPCtrls As Integer
Dim tmICtrls(0 To 20)  As FIELDAREA
Dim imLBICtrls As Integer
Dim imBoxNoMap(1 To 20) As Integer
Dim imPaintIndex As Integer '1=Lock Box; 0=EDI Service
Dim imBoxNo As Integer   'Current event name Box
Dim imRowNo As Integer  'Current event row
Dim smSave() As String  '1=Ref #; 2=Advertiser; 3=Product; 4=Protection 1; 5=Protection 2
                        '6=Length; 7=Start Date; 8=End Date; 9=Run Every; 10=Daily/Weekly
                        '11=Spots/Week; 12-18=either # spots or air Day (Y or N);
                        '19=Start Time; 20=End Time; 21=ISCI; 22=Creative Title; 23=Cart; 24=Schedule Status
                        '
Dim smShow() As String

Dim tmFsf As FSF        'Fsf record image
Dim tmFSFSrchKey As LONGKEY0    'Fsf key record image
Dim tmFsfSrchKey3 As FSFKEY3    'Fsf key record image
Dim hmFsf As Integer    'Sale Commission file handle
Dim imFsfRecLen As Integer        'Fsf record length
Dim imFsfChg As Integer     'Indicates if field changed

Dim hmFnf As Integer 'Feed name file handle
Dim tmFnf As FNF        'Fnf record image
Dim tmFnfSrchKey As INTKEY0    'Fnf key record image
Dim imFnfRecLen As Integer        'Fnf record length

Dim hmPrf As Integer            'Product file handle
Dim tmPrfSrchKey0 As LONGKEY0            'PRF record image
Dim imPrfRecLen As Integer        'PRF record length
Dim tmPrf As PRF

Dim hmAdf As Integer            'Advertiser file handle
Dim tmAdfSrchKey As INTKEY0            'ADF record image
Dim imAdfRecLen As Integer        'ADF record length
Dim tmAdf As ADF
Dim imAdfCode As Integer

Dim hmMnf As Integer            'Advertiser file handle
Dim tmMnfSrchKey As INTKEY0            'ADF record image
Dim imMnfRecLen As Integer        'ADF record length
Dim tmMnf As MNF
Dim imProdRowNo As Integer
Dim smComp1 As String
Dim smComp2 As String

Dim hmMcf As Integer            'Media Code file handle
Dim tmMcfSrchKey0 As INTKEY0            'MCF record image
Dim imMcfRecLen As Integer        'MCF record length
Dim tmMcf As MCF

Dim hmCif As Integer            'Copy Inventory file handle
Dim tmCifSrchKey0 As LONGKEY0            'CIF record image
Dim imCifRecLen As Integer        'CIF record length
Dim tmCif As CIF

Dim hmCpf As Integer            'Copy Product file handle
Dim tmCpfSrchKey0 As LONGKEY0            'CPF record image
Dim imCpfRecLen As Integer        'CPF record length
Dim tmCpf As CPF

Dim tmFeedNameCode() As SORTCODE
Dim smFeedNameCodeTag As String

Dim tmFeedAdvertiser() As SORTCODE
Dim smFeedtAdvertiserTag As String

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcVehicle was populated
Dim imBypassSetting As Integer      'In cbcVehicle--- bypass mSetCommands (when user entering new name, don't want cbcVehicle disabled)
Dim imChgMode As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imVehSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imFdNmSelectedIndex As Integer
Dim imComboBoxIndex As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imSettingValue As Integer
Dim smNowDate As String
'Calendar variables
Dim tmCDCtrls(0 To 7) As FIELDAREA  'Field area image
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer        'Month of displayed calendar
Dim imCalMonth As Integer       'Year of displayed calendar
Dim lmCalStartDate As Long      'Start date of displayed calendar
Dim lmCalEndDate As Long        'End date of displayed calendar
Dim imCalType As Integer        'Calendar type
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer

Dim tmUserVehicle() As SORTCODE
Dim smUserVehicleTag As String

Const IREFNOINDEX = 1
Const IADVTINDEX = 2
Const IPRODINDEX = 3
Const IPROT1INDEX = 4
Const IPROT2INDEX = 5
Const ILENINDEX = 6
Const ISDATEINDEX = 7
Const IEDATEINDEX = 8
Const IRUNINDEX = 9
Const IDWINDEX = 10
Const ISPOTSINDEX = 11
Const IMOINDEX = 12
Const ITUINDEX = 13
Const IWEINDEX = 14
Const ITHINDEX = 15
Const IFRINDEX = 16
Const ISAINDEX = 17
Const ISUINDEX = 18
Const ISTIMEINDEX = 19
Const IETIMEINDEX = 20

Const LADVTINDEX = 1
Const LPRODINDEX = 2
Const LPROT1INDEX = 3
Const LPROT2INDEX = 4
Const LLENINDEX = 5
Const LDATEINDEX = 6
Const LTIMEINDEX = 7
Const LISCIINDEX = 8
Const LCREATIVEINDEX = 9
Const LCARTINDEX = 10

Const PADVTINDEX = 1
Const PPRODINDEX = 2
Const PPROT1INDEX = 3
Const PPROT2INDEX = 4
Const PLENINDEX = 5
Const PSDATEINDEX = 6
Const PEDATEINDEX = 7
Const PSTIMEINDEX = 8
Const PETIMEINDEX = 9
Const PISCIINDEX = 10
Const PCREATIVEINDEX = 11
Const PCARTINDEX = 12


'*******************************************************
'*                                                     *
'*      Procedure Name:mGetProduct                     *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get advertiser                *
'*                                                     *
'*******************************************************
Private Sub mGetProd(ilRowNo As Integer)
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    If ilRowNo = imProdRowNo Then
        Exit Sub
    End If
    imProdRowNo = ilRowNo
    mProdPop ilRowNo
    gFindMatch smSave(3, ilRowNo), 2, lbcProd
    If gLastFound(lbcProd) >= 2 Then
        slNameCode = tgProdCode(gLastFound(lbcProd) - 2).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        imAdfCode = Val(slCode)
        tmPrfSrchKey0.lCode = Val(slCode)
        If tmPrf.lCode <> tmPrfSrchKey0.lCode Then
            ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        End If
        tmMnfSrchKey.iCode = tmPrf.iMnfComp(0)
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smComp1 = tmMnf.sName
        Else
            smComp1 = ""
        End If
        tmMnfSrchKey.iCode = tmPrf.iMnfComp(1)
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smComp2 = tmMnf.sName
        Else
            smComp2 = ""
        End If
    Else
        smComp1 = ""
        smComp2 = ""
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
Private Sub mLenPop()
    'Dim ilRet As Integer
    Dim ilVpfIndex As Integer
    Dim ilVefCode As Integer
    Dim ilLoop As Integer
    Dim ilLen As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer

    ilVpfIndex = -1
    If imVehSelectedIndex >= 0 Then
        slNameCode = tmUserVehicle(imVehSelectedIndex).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        ilVpfIndex = gVpfFind(FeedSpot, ilVefCode)
    End If
    lbcLen.Clear
    If ilVpfIndex >= 0 Then
        For ilLoop = LBound(tgVpf(ilVpfIndex).iSLen) To UBound(tgVpf(ilVpfIndex).iSLen) Step 1
            If tgVpf(ilVpfIndex).iSLen(ilLoop) <> 0 Then
                lbcLen.AddItem Trim$(Str$(tgVpf(ilVpfIndex).iSLen(ilLoop)))
            End If
        Next ilLoop
    Else
        For ilLoop = LBound(tgVpf) To UBound(tgVpf) Step 1
            For ilLen = LBound(tgVpf(ilLoop).iSLen) To UBound(tgVpf(ilLoop).iSLen) Step 1
                If tgVpf(ilLoop).iSLen(ilLen) <> 0 Then
                    gFindMatch Trim$(Str$(tgVpf(ilLoop).iSLen(ilLen))), 0, lbcLen
                    If gLastFound(lbcLen) < 0 Then
                        lbcLen.AddItem Trim$(Str$(tgVpf(ilLoop).iSLen(ilLen)))
                    End If
                End If
            Next ilLen
        Next ilLoop
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetAdvt                        *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get advertiser                *
'*                                                     *
'*******************************************************
Private Sub mGetAdvt(slAdvtName As String)
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String
    'If Trim$(slAdvtName) = Trim$(tmAdf.sName) Then
    If (Trim$(slAdvtName) = Trim$(tmAdf.sName)) And (imAdfCode > 0) And (imAdfCode = tmAdf.iCode) Then
        Exit Sub
    End If
    gFindMatch slAdvtName, 1, lbcAdvt
    If gLastFound(lbcAdvt) >= 1 Then
        slNameCode = tmFeedAdvertiser(gLastFound(lbcAdvt) - 1).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        imAdfCode = Val(slCode)
        tmAdfSrchKey.iCode = imAdfCode
        If tmAdf.iCode <> imAdfCode Then
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            slStr = tmAdf.sProduct
        End If
    Else
        imAdfCode = 0
        tmAdf.sProduct = ""
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
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcProd, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mProdBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(ADVTPRODEXE)) Then
    '    imDoubleClickName = False
    '    mProdBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    Screen.MousePointer = vbHourglass  'Wait
    igAdvtProdCallSource = CALLSOURCEFEED
    sgAdvtProdName = smSave(2, imRowNo)
    If edcDropDown.Text = "[New]" Then
        sgAdvtProdName = sgAdvtProdName & "\" & " "
    Else
        sgAdvtProdName = sgAdvtProdName & "\" & Trim$(edcDropDown.Text)
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Invoice!edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (igShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Invoice^Test\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
        Else
            slStr = "Invoice^Prod\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Invoice^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    '    Else
    '        slStr = "Invoice^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "AdvtProd.Exe " & slStr, 1)
    'SaleHist.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    AdvtProd.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgAdvtProdName)
    igAdvtProdCallSource = Val(sgAdvtProdName)
    ilParse = gParseItem(slStr, 2, "\", sgAdvtProdName)
    'SaleHist.Enabled = True
    'Invoice!edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mProdBranch = True
    imUpdateAllowed = ilUpdateAllowed
    If igAdvtProdCallSource = CALLDONE Then  'Done
        igAdvtProdCallSource = CALLNONE
'        gSetMenuState True
        lbcProd.Clear
        sgProdCodeTag = ""
        mProdPop imRowNo
        If imTerminate Then
            mProdBranch = False
            Exit Function
        End If
        gFindMatch sgAdvtProdName, 1, lbcProd
        If gLastFound(lbcProd) > 0 Then
            imChgMode = True
            lbcProd.ListIndex = gLastFound(lbcProd)
            edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
            imChgMode = False
            mProdBranch = False
        Else
            imChgMode = True
            lbcProd.ListIndex = -1
            edcDropDown.Text = sgAdvtProdName
            imChgMode = False
            edcDropDown.SetFocus
            sgAdvtProdName = ""
            Exit Function
        End If
        sgAdvtProdName = ""
    End If
    If igAdvtProdCallSource = CALLCANCELLED Then  'Cancelled
        igAdvtProdCallSource = CALLNONE
        sgAdvtProdName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igAdvtProdCallSource = CALLTERMINATED Then
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
Private Sub mProdPop(ilRowNo As Integer)
'
'   mProdPop
'   Where:
'       ilAdvtCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer

    mGetAdvt smSave(2, ilRowNo)
    If imAdfCode <= 0 Then
        lbcProd.Clear
        'lbcProdCode.Clear
        'lbcProdCode.Tag = ""
        ReDim tgProdCode(0 To 0) As SORTCODE
        sgProdCodeTag = ""
        lbcProd.AddItem "[None]", 0  'Force as first item on list
        lbcProd.AddItem "[New]", 0  'Force as first item on list
        Exit Sub
    End If
    ilIndex = lbcProd.ListIndex
    If ilIndex > 1 Then
        slName = lbcProd.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtProdBox(SaleHist, ilAdfCode, lbcProduct, lbcProdCode)
    ilRet = gPopAdvtProdBox(FeedSpot, imAdfCode, lbcProd, tgProdCode(), sgProdCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mProdPopErr
        gCPErrorMsg ilRet, "mProdPop (gPopAdvtProdBox)", FeedSpot
        On Error GoTo 0
        lbcProd.AddItem "[None]", 0  'Force as first item on list
        lbcProd.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcProd
            If gLastFound(lbcProd) > 1 Then
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
    If igWinStatus(COMPETITIVESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mCompBranch = True
        mSetFocus imBoxNo
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
    igMNmCallSource = CALLSOURCEFEED
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
            slStr = "Feed^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Feed^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Feed^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "Feed^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Advt.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Advt.Enabled = True
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
        ilRet = csiSetStamp("COMPMNF", sgCompMnfStamp)
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
        Else
            imChgMode = True
            lbcComp(ilIndex).ListIndex = 1
            edcDropDown.Text = lbcComp(ilIndex).List(1)
            imChgMode = False
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
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffset(0) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    ReDim slComp(0 To 1) As String      'Competitive name, saved to determine if changed
    ReDim ilComp(0 To 1) As Integer      'Competitive name, saved to determine if changed
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "C"
    ilOffset(0) = gFieldOffset("Mnf", "MnfType") '2
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
    'ilRet = gIMoveListBox(Feed, lbcComp(0), lbcCompCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    ilRet = gIMoveListBox(FeedSpot, lbcComp(0), tgCompCode(), sgCompCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffset())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mCompPopErr
        gCPErrorMsg ilRet, "mCompPop (gIMoveListBox)", FeedSpot
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
'*      Procedure Name:mAdvtBranch                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      advertiser and process         *
'*                      communication back from        *
'*                      advertiser                     *
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
Private Function mAdvtBranch() As Integer
'
'   ilRet = mAdvtBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    'Dim ilNoForms As Integer
    'Dim ilUsage As Integer
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    slStr = Trim$(edcDropDown.Text)
    If slStr = "" Then
        imDoubleClickName = False
        mAdvtBranch = False
        Exit Function
    End If
    ilRet = gOptionalLookAhead(edcDropDown, lbcAdvt, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        imDoubleClickName = False
        mAdvtBranch = False
        Exit Function
    End If
    If igWinStatus(ADVERTISERSLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mAdvtBranch = True
        mSetFocus imBoxNo
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(ADVERTISERSLIST)) Then
    '    imDoubleClickName = False
    '    mAdvtBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    Screen.MousePointer = vbHourglass  'Wait
    igAdvtCallSource = CALLSOURCEFEED
    If edcDropDown.Text = "[New]" Then
        sgAdvtName = ""
    Else
        sgAdvtName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Traffic!edcLinkSrceHelpMsg.Text = ""
    If igTestSystem Then
        slStr = "Feed^Test\" & sgUserName & "\" & Trim$(Str$(igAdvtCallSource)) & "\" & sgAdvtName
    Else
        slStr = "Feed^Prod\" & sgUserName & "\" & Trim$(Str$(igAdvtCallSource)) & "\" & sgAdvtName
    End If
    'lgShellRet = Shell(sgExePath & "Advt.Exe " & slStr, 1)
    'Traffic.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    Advt.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgAdvtName)
    igAdvtCallSource = Val(sgAdvtName)
    ilParse = gParseItem(slStr, 2, "\", sgAdvtName)
    'Traffic.Enabled = True
    'Traffic!edcLinkSrceHelpMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mAdvtBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    If igAdvtCallSource = CALLDONE Then  'Done
        igAdvtCallSource = CALLNONE
'        gSetMenuState True
        'slFilter = cbcFilter.Text
        'ilFilter = imFilterSelectedIndex
'        slStamp = FileDateTime(sgDBPath & "Adf.Btr")
'        If StrComp(slStamp, Traffic!lbcAdvt.Tag, 1) <> 0 Then
        lbcAdvt.Clear
        smFeedtAdvertiserTag = ""    'Traffic!lbcAdvt.Tag = ""
        mAdvtPop
        If imTerminate Then
            mAdvtBranch = False
            Exit Function
        End If
'        End If
        gFindMatch sgAdvtName, 1, lbcAdvt
        If gLastFound(lbcAdvt) > 0 Then
            imChgMode = True
            lbcAdvt.ListIndex = gLastFound(lbcAdvt)
            edcDropDown.Text = lbcAdvt.List(lbcAdvt.ListIndex)
            imChgMode = False
            mAdvtBranch = False
        Else
            imChgMode = True
            lbcAdvt.ListIndex = 1
            edcDropDown.Text = lbcAdvt.List(lbcAdvt.ListIndex)
            imChgMode = False
            edcDropDown.SetFocus
            sgAdvtName = ""
            Exit Function
        End If
        sgAdvtName = ""
    End If
    If igAdvtCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igAdvtCallSource = CALLNONE
        sgAdvtName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igAdvtCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igAdvtCallSource = CALLNONE
        sgAdvtName = ""
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
    ilIndex = lbcAdvt.ListIndex
    If ilIndex > 0 Then
        slName = lbcAdvt.List(ilIndex)
    End If
    igPopAdfAgfDormant = True
    'Repopulate if required- if sales source changed by another user while in this screen
    ilRet = gPopAdvtBox(FeedSpot, lbcAdvt, tmFeedAdvertiser(), smFeedtAdvertiserTag) 'Traffic!lbcAdvt)
    igPopAdfAgfDormant = True
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", FeedSpot
        On Error GoTo 0
        lbcAdvt.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcAdvt
            If gLastFound(lbcAdvt) > 0 Then
                lbcAdvt.ListIndex = gLastFound(lbcAdvt)
            Else
                lbcAdvt.ListIndex = -1
            End If
        Else
            lbcAdvt.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mFnfReadRec                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mFnfReadRec() As Integer
'
'   iRet = ENmRead(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    Dim ilLoop As Integer

    slNameCode = tmFeedNameCode(imFdNmSelectedIndex).sKey    'lbcNameCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mFnfReadRecErr
    gCPErrorMsg ilRet, "mFnfReadRec (gParseItem field 2)", FeedSpot
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmFnfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmFnf, tmFnf, imFnfRecLen, tmFnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mFnfReadRecErr
    gBtrvErrorMsg ilRet, "mFnfReadRec (btrGetEqual)", FeedSpot
    On Error GoTo 0
    If tmFnf.sPledgeTime = "I" Then
        imPaintIndex = 2
        ReDim tmCtrls(0 To UBound(tmICtrls)) As FIELDAREA
        For ilLoop = 1 To UBound(tmICtrls) Step 1
            tmCtrls(ilLoop) = tmICtrls(ilLoop)
        Next ilLoop
        pbcFeed(2).Visible = True
        pbcFeed(0).Visible = False
        pbcFeed(1).Visible = False
        cmcPledge.Enabled = False
        lacType.Caption = "Insertion Order"
    ElseIf tmFnf.sPledgeTime = "L" Then
        imPaintIndex = 1
        ReDim tmCtrls(0 To UBound(tmLCtrls)) As FIELDAREA
        For ilLoop = 1 To UBound(tmLCtrls) Step 1
            tmCtrls(ilLoop) = tmLCtrls(ilLoop)
        Next ilLoop
        pbcFeed(1).Visible = True
        pbcFeed(0).Visible = False
        pbcFeed(2).Visible = False
        cmcPledge.Enabled = True
        lacType.Caption = "Log Needs Conversion"
    Else
        imPaintIndex = 0
        ReDim tmCtrls(0 To UBound(tmPCtrls)) As FIELDAREA
        For ilLoop = 1 To UBound(tmPCtrls) Step 1
            tmCtrls(ilLoop) = tmPCtrls(ilLoop)
        Next ilLoop
        pbcFeed(0).Visible = True
        pbcFeed(1).Visible = False
        pbcFeed(2).Visible = False
        cmcPledge.Enabled = False
        lacType.Caption = "Pre-Log Converted"
    End If
    If tmFnf.iMcfCode > 0 Then
        tmMcfSrchKey0.iCode = tmFnf.iMcfCode
        ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mFnfReadRecErr
        gBtrvErrorMsg ilRet, "mFnfReadRec (btrGetEqual:MCF)", FeedSpot
        On Error GoTo 0
    Else
        tmMcf.iCode = 0
    End If
    mBuildMap
    mFnfReadRec = True
    Exit Function
mFnfReadRecErr:
    On Error GoTo 0
    mFnfReadRec = False
    Exit Function
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
    ilRet = gPopUserVehicleBox(FeedSpot, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, cbcVehicle, tmUserVehicle(), smUserVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", FeedSpot
        On Error GoTo 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub cbcFeed_Change()
    Dim ilRet As Integer
    If imChgMode = False Then
        imChgMode = True
        Screen.MousePointer = vbHourglass  'Wait
        If cbcFeed.Text <> "" Then
            gManLookAhead cbcFeed, imBSMode, imComboBoxIndex
        End If
        imFdNmSelectedIndex = cbcFeed.ListIndex
        mClearCtrlFields
        ilRet = mFnfReadRec()
        If cbcVehicle.ListCount > 1 Then
            cbcVehicle.ListIndex = -1
            imVehSelectedIndex = -1
        End If
        ilRet = mReadRec()
        pbcFeed(imPaintIndex).Cls
        mMoveRecToCtrl
        mInitShow
        mSetMinMax
        mSetCommands
        imChgMode = False
        Screen.MousePointer = vbDefault    'Default
    End If
End Sub

Private Sub cbcFeed_Click()
    cbcFeed_Change    'Process change as change event is not generated by VB
End Sub

Private Sub cbcFeed_GotFocus()

    plcCalendar.Visible = False
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then
        imFirstFocus = False
    End If
    If cbcFeed.Text = "" Then
        If cbcFeed.ListCount > 0 Then
            cbcFeed.ListIndex = 0
        End If
    Else
        cbcFeed.ListIndex = imFdNmSelectedIndex
    End If
    gCtrlGotFocus cbcFeed
    imComboBoxIndex = cbcFeed.ListIndex
    imFdNmSelectedIndex = imComboBoxIndex
    If cbcFeed.ListCount = 1 Then
        cbcVehicle.SetFocus
    End If
    'tmcClick.Enabled = False

End Sub

Private Sub cbcFeed_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcFeed_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcFeed.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cbcVehicle_Change()
    Dim ilRet As Integer
    If imChgMode = False Then
        imChgMode = True
        Screen.MousePointer = vbHourglass  'Wait
        If cbcVehicle.Text <> "" Then
            gManLookAhead cbcVehicle, imBSMode, imComboBoxIndex
        End If
        imVehSelectedIndex = cbcVehicle.ListIndex
        mClearCtrlFields
        ilRet = mReadRec()
        pbcFeed(imPaintIndex).Cls
        mMoveRecToCtrl
        mInitShow
        mSetMinMax
        mSetCommands
        imChgMode = False
        Screen.MousePointer = vbDefault    'Default
    End If
End Sub
Private Sub cbcVehicle_Click()
    cbcVehicle_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcVehicle_GotFocus()

    plcCalendar.Visible = False
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then
        imFirstFocus = False
    End If
    imComboBoxIndex = imVehSelectedIndex
    gCtrlGotFocus cbcVehicle
    If cbcFeed.Text = "" Then
        If cbcVehicle.ListCount > 0 Then
            cbcVehicle.ListIndex = 0
        End If
    Else
        cbcVehicle.ListIndex = imVehSelectedIndex
    End If
    gCtrlGotFocus cbcVehicle
    imComboBoxIndex = cbcVehicle.ListIndex
    imVehSelectedIndex = imComboBoxIndex
    If cbcVehicle.ListCount = 1 Then
        edcDate.SetFocus
    End If
    Exit Sub
End Sub
Private Sub cbcVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcVehicle_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcVehicle.SelLength <> 0 Then    'avoid deleting two characters
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
    If imBoxNo > 0 Then
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
        edcDropDown.SetFocus
    Else
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
        edcDate.SetFocus
    End If
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    If imBoxNo > 0 Then
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
        edcDropDown.SetFocus
    Else
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
        edcDate.SetFocus
    End If
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    If imBoxNo = -1 Then
        plcCalendar.Visible = False
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    plcCalendar.Move plcSelect.Left + plcSelect.Width - fgBevelX - plcCalendar.Width, plcSelect.Top + edcDate.Height + fgBevelY
    edcDate.SelStart = 0
    edcDate.SelLength = Len(edcDate.Text)
    edcDate.SetFocus
End Sub

Private Sub cmcDate_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    gCtrlGotFocus ActiveControl
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
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    If imBoxNo = -1 Then
        plcCalendar.Visible = False
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNoMap(imBoxNo)
        Case IADVTINDEX
            lbcAdvt.Visible = Not lbcAdvt.Visible
        Case IPRODINDEX
            lbcProd.Visible = Not lbcProd.Visible
        Case IPROT1INDEX
            lbcComp(0).Visible = Not lbcComp(0).Visible
        Case IPROT2INDEX
            lbcComp(1).Visible = Not lbcComp(1).Visible
        Case ILENINDEX
            lbcLen.Visible = Not lbcLen.Visible
        Case ISDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case IEDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case IRUNINDEX
            lbcRun.Visible = Not lbcRun.Visible
        Case IDWINDEX
            lbcDW.Visible = Not lbcDW.Visible
        Case ISTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case IETIMEINDEX
            plcTme.Visible = Not plcTme.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcPledge_Click()
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String

    If imFdNmSelectedIndex < 0 Then
        Exit Sub
    End If
    If imVehSelectedIndex < 0 Then
        Exit Sub
    End If
    slNameCode = tmFeedNameCode(imFdNmSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    igPledgeFnfCode = Val(slCode)
    slNameCode = tmUserVehicle(imVehSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    igPledgeVefCode = Val(slCode)
    FeedPlge.Show vbModal
End Sub

Private Sub cmcReport_Click()
    Dim slStr As String
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = FEEDJOB
    igRptType = 0
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "FeedSpot^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
        Else
            slStr = "FeedSpot^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "FeedSpot^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    Else
    '        slStr = "FeedSpot^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'FeedSpot.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'FeedSpot.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    sgCommandStr = slStr
    RptList.Show vbModal
    ''Screen.MousePointer = vbDefault    'Default
End Sub

Private Sub cmcReport_GotFocus()
    If imBoxNo = -1 Then
        plcCalendar.Visible = False
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub

Private Sub cmcReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcSave_Click()
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
    ReDim tgFsfDel(1 To 1) As FSFREC
    pbcFeed(imPaintIndex).Cls
    pbcFeed_Paint imPaintIndex
    imBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
    imFsfChg = False
    mSetCommands
    'pbcSTab.SetFocus
End Sub

Private Sub cmcSave_GotFocus()
    If imBoxNo = -1 Then
        plcCalendar.Visible = False
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub

Private Sub cmcSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcSchedule_Click()
'    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
'    Dim ilExtLen As Integer
'    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
'    Dim llNoRec As Long         'Number of records in Sof
'    Dim ilUpper As Integer
'    Dim ilLoop As Integer
'    ReDim llFsfCode(0 To 0) As Long
'
'    btrExtClear hmFsf   'Clear any previous extend operation
'    ilExtLen = Len(tmFsf)  'Extract operation record size
'    ilRet = btrGetFirst(hmFsf, tmFsf, imFsfRecLen, INDEXKEY2, BTRV_LOCK_NONE)
'    If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
'        ilRet = BTRV_ERR_END_OF_FILE
'    Else
'        If ilRet <> BTRV_ERR_NONE Then
'            mReadRec = False
'            Exit Function
'        End If
'    End If
'    If ilRet <> BTRV_ERR_END_OF_FILE Then
'        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
'        Call btrExtSetBounds(hmFsf, llNoRec, -1, "UC", "FSF", "") '"EG") 'Set extract limits (all records)
'        ilOffset = gFieldOffset("FSF", "FSFSCHSTATUS") 'GetOffSetForInt(tmFsf, tmFsf.iSlfCode)
'        tlCharTypeBuff.sType = "F"
'        ilRet = btrExtAddLogicConst(hmFsf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_NE, BTRV_EXT_LAST, tlCharTypeBuff, 1)
'        On Error GoTo mReadRecErr
'        gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Fsf.Btr", FeedSpot
'        On Error GoTo 0
'        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
'        ilRet = btrExtGetNext(hmFsf, tmFsf, ilExtLen, llRecPos)
'        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'            On Error GoTo mReadRecErr
'            gBtrvErrorMsg ilRet, "mReadRec (btrExtGetNextExt):" & "Fsf.Btr", FeedSpot
'            On Error GoTo 0
'            'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
'            ilExtLen = Len(tmFsf)  'Extract operation record size
'            Do While ilRet = BTRV_ERR_REJECT_COUNT
'                ilRet = btrExtGetNext(hmFsf, tmFsf, ilExtLen, llRecPos)
'            Loop
'            Do While ilRet = BTRV_ERR_NONE
'                ilUpper = UBound(llFsfCode)
'                llFsfCode(ilUpper) = tmFsf.lCode
'                ReDim Preserve llFsfCode(1 To ilUpper) As Long
'                ilRet = btrExtGetNext(hmFsf, tmFsf, ilExtLen, llRecPos)
'                Do While ilRet = BTRV_ERR_REJECT_COUNT
'                    ilRet = btrExtGetNext(hmFsf, tmFsf, ilExtLen, llRecPos)
'                Loop
'            Loop
'        End If
'    End If
'    If UBound(llFsfCode) > LBound(llFsfCode) Then
'        For ilLoop = LBound(llFsfCode) To UBound(llFsfCode) - 1 Step 1
'            ilRet = gFeedSchSpots(False, llFsfCode(ilLoop))
'        Next ilLoop
'    End If
    sgGenMsg = sgLF & sgCR & "Scheduling Feed Spots..."
    igDefCMC = 1
    GenSch.Show vbModal
    If UBound(tgFsfRec) > LBound(tgFsfRec) Then
        mClearCtrlFields
        ilRet = mReadRec()
        pbcFeed(imPaintIndex).Cls
        mMoveRecToCtrl
        mInitShow
        mSetMinMax
        mSetCommands
    End If
End Sub

Private Sub cmcSchedule_GotFocus()
    If imBoxNo = -1 Then
        plcCalendar.Visible = False
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub

Private Sub edcDate_Change()
    Dim slStr As String
    Dim ilRet As Integer

    slStr = edcDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    mClearCtrlFields
    lacDate.Visible = True
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    ilRet = mReadRec()
    pbcFeed(imPaintIndex).Cls
    mMoveRecToCtrl
    mInitShow
    mSetMinMax
    mSetCommands
End Sub

Private Sub edcDate_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
        plcCalendar.Move plcSelect.Left + plcSelect.Width - fgBevelX - plcCalendar.Width, plcSelect.Top + edcDate.Height + fgBevelY
    End If
    imBypassFocus = False
End Sub

Private Sub edcDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcDate_KeyPress(KeyAscii As Integer)
    Dim ilKeyAscii As Integer
    ilKeyAscii = KeyAscii
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String

    If (KeyCode = KeyUp) Or (KeyCode = KeyDown) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KeyUp Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcDate.Text = slDate
            End If
        End If
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcDate.Text = slDate
            End If
        End If
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
    End If
End Sub

Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNoMap(imBoxNo)
        Case IADVTINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcAdvt, imBSMode, slStr)
            If ilRet = 1 Then
                lbcAdvt.ListIndex = 1
            End If
        Case IPRODINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcProd, imBSMode, slStr)
            If ilRet = 1 Then
                lbcProd.ListIndex = 1
            End If
        Case IPROT1INDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcComp(0), imBSMode, slStr)
            If ilRet = 1 Then
                lbcComp(0).ListIndex = 1
            End If
        Case IPROT2INDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcComp(1), imBSMode, slStr)
            If ilRet = 1 Then
                lbcComp(1).ListIndex = 1
            End If
        Case ILENINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcLen, imBSMode, imComboBoxIndex
        Case ISDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case IEDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case IRUNINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcRun, imBSMode, imComboBoxIndex
        Case IDWINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcDW, imBSMode, imComboBoxIndex
    End Select
    imLbcArrowSetting = False
End Sub

Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub edcDropDown_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
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
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilKeyAscii As Integer

    ilKeyAscii = KeyAscii
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imBoxNoMap(imBoxNo)
        Case ISDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case IEDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case IMOINDEX To ISUINDEX
            If smSave(10, imRowNo) = "Daily" Then
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        Case ISTIMEINDEX, IETIMEINDEX
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
    If (KeyCode = KeyUp) Or (KeyCode = KeyDown) Then
        Select Case imBoxNoMap(imBoxNo)
            Case IADVTINDEX
                gProcessArrowKey Shift, KeyCode, lbcAdvt, imLbcArrowSetting
            Case IPRODINDEX
                gProcessArrowKey Shift, KeyCode, lbcProd, imLbcArrowSetting
            Case IPROT1INDEX
                gProcessArrowKey Shift, KeyCode, lbcComp(0), imLbcArrowSetting
            Case IPROT2INDEX
                gProcessArrowKey Shift, KeyCode, lbcComp(1), imLbcArrowSetting
            Case ILENINDEX
                gProcessArrowKey Shift, KeyCode, lbcLen, imLbcArrowSetting
            Case ISDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KeyUp Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case IEDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KeyUp Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case IRUNINDEX
                gProcessArrowKey Shift, KeyCode, lbcRun, imLbcArrowSetting
            Case IDWINDEX
                gProcessArrowKey Shift, KeyCode, lbcDW, imLbcArrowSetting
            Case ISTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case IETIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNoMap(imBoxNo)
            Case ISDATEINDEX
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
            Case IEDATEINDEX
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
        End Select
    End If
End Sub

Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case imBoxNoMap(imBoxNo)
        Case IADVTINDEX
            If imTabDirection = -1 Then  'Right To Left
                pbcSTab.SetFocus
            Else
                pbcTab.SetFocus
            End If
            Exit Sub
        Case IPRODINDEX
            If imTabDirection = -1 Then  'Right To Left
                pbcSTab.SetFocus
            Else
                pbcTab.SetFocus
            End If
            Exit Sub
        Case IPROT1INDEX
            If imTabDirection = -1 Then  'Right To Left
                pbcSTab.SetFocus
            Else
                pbcTab.SetFocus
            End If
            Exit Sub
        Case IPROT2INDEX
            If imTabDirection = -1 Then  'Right To Left
                pbcSTab.SetFocus
            Else
                pbcTab.SetFocus
            End If
            Exit Sub
        Case ILENINDEX
            If imTabDirection = -1 Then  'Right To Left
                pbcSTab.SetFocus
            Else
                pbcTab.SetFocus
            End If
            Exit Sub
        Case ISDATEINDEX
            If imTabDirection = -1 Then  'Right To Left
                pbcSTab.SetFocus
            Else
                pbcTab.SetFocus
            End If
            Exit Sub
        Case IEDATEINDEX
            If imTabDirection = -1 Then  'Right To Left
                pbcSTab.SetFocus
            Else
                pbcTab.SetFocus
            End If
            Exit Sub
        Case IRUNINDEX
            If imTabDirection = -1 Then  'Right To Left
                pbcSTab.SetFocus
            Else
                pbcTab.SetFocus
            End If
            Exit Sub
        Case IDWINDEX
            If imTabDirection = -1 Then  'Right To Left
                pbcSTab.SetFocus
            Else
                pbcTab.SetFocus
            End If
            Exit Sub
    End Select
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True  'To get Alt J and Alt L keys
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(FEEDJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcFeed(imPaintIndex).Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcFeed(imPaintIndex).Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    Me.ZOrder 0 'Send to front
    FeedSpot.Refresh
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
        plcCalendar.Visible = False
        plcNum.Visible = False
        If (cbcVehicle.Enabled) And (imBoxNo > 0) Then
            cbcFeed.Enabled = False
            cbcVehicle.Enabled = False
            cmcDate.Enabled = False
            edcDate.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            cbcFeed.Enabled = True
            cbcVehicle.Enabled = True
            cmcDate.Enabled = True
            edcDate.Enabled = True
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
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    'btrStopAppl
    'End
    igJobShowing(FEEDJOB) = False
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub imcHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub imcTrash_Click()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilUpperBound As Integer
    Dim ilRowNo As Integer
    Dim ilFsf As Integer
    If (imRowNo < vbcFeed.Value) Or (imRowNo > vbcFeed.Value + vbcFeed.LargeChange) Then
        Exit Sub
    End If
    ilRowNo = imRowNo
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame(imPaintIndex).Visible = False
    gCtrlGotFocus ActiveControl
    ilUpperBound = UBound(smSave, 2)
    ilFsf = ilRowNo
    If ilFsf = ilUpperBound Then
        mInitNew ilFsf
    Else
        If ilFsf > 0 Then
            If tgFsfRec(ilFsf).iStatus = 1 Then
                tgFsfDel(UBound(tgFsfDel)).tFsf = tgFsfRec(ilFsf).tFsf
                tgFsfDel(UBound(tgFsfDel)).iStatus = tgFsfRec(ilFsf).iStatus
                tgFsfDel(UBound(tgFsfDel)).lRecPos = tgFsfRec(ilFsf).lRecPos
                ReDim Preserve tgFsfDel(1 To UBound(tgFsfDel) + 1) As FSFREC
            End If
            ilFsf = ilRowNo
            'Remove record from tgRjf1Rec- Leave tgPjf2Rec
            For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
                tgFsfRec(ilLoop) = tgFsfRec(ilLoop + 1)
            Next ilLoop
            ReDim Preserve tgFsfRec(1 To UBound(tgFsfRec) - 1) As FSFREC
        End If
        For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
            For ilIndex = 1 To UBound(smSave, 1) Step 1
                smSave(ilIndex, ilLoop) = smSave(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(smShow, 1) Step 1
                smShow(ilIndex, ilLoop) = smShow(ilIndex, ilLoop + 1)
            Next ilIndex
        Next ilLoop
        ilUpperBound = UBound(smSave, 2)
        ReDim Preserve smShow(1 To 20, 1 To ilUpperBound - 1) As String 'Values shown in program area
        ReDim Preserve smSave(1 To 24, 1 To ilUpperBound - 1) As String    'Values saved (program name) in program area
        imFsfChg = True
    End If
    mSetCommands
    lacFrame(imPaintIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    imSettingValue = True
    vbcFeed.Min = LBound(smShow, 2)
    imSettingValue = True
    If UBound(smShow, 2) - 1 <= vbcFeed.LargeChange + 1 Then ' + 1 Then
        vbcFeed.Max = LBound(smShow, 2)
    Else
        vbcFeed.Max = UBound(smShow, 2) - vbcFeed.LargeChange
    End If
    imSettingValue = True
    vbcFeed.Value = vbcFeed.Min
    imSettingValue = True
    pbcFeed(imPaintIndex).Cls
    pbcFeed_Paint imPaintIndex
End Sub
Private Sub imcTrash_DragDrop(Source As Control, X As Single, Y As Single)
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        lacFrame(imPaintIndex).DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        lacFrame(imPaintIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub lbcAdvt_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcAdvt, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcAdvt_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcAdvt_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcAdvt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcAdvt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcAdvt, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
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

Private Sub lbcDW_Click()
    gProcessLbcClick lbcDW, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcLen_Click()
    gProcessLbcClick lbcLen, edcDropDown, imChgMode, imLbcArrowSetting
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

Private Sub lbcRun_Click()
    gProcessLbcClick lbcRun, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcVehicle_Click()
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
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

    If imBoxNo <> -1 Then
        slStr = edcDropDown.Text
    Else
        slStr = edcDate.Text
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

    imFsfChg = False
    lbcVehicle.ListIndex = -1
    ReDim tgFsfRec(1 To 1) As FSFREC
    tgFsfRec(1).iStatus = -1
    tgFsfRec(1).lRecPos = 0
    tgFsfRec(1).iDateChg = False
    ReDim tgFsfDel(1 To 1) As FSFREC
    tgFsfDel(1).iStatus = -1
    tgFsfDel(1).lRecPos = 0
    ReDim smShow(1 To 20, 1 To 1) As String 'Values shown in program area
    ReDim smSave(1 To 24, 1 To 1) As String 'Values saved (program name) in program area
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, 1) = ""
    Next ilLoop
    vbcFeed.Min = LBound(smShow, 2)
    imSettingValue = True
    vbcFeed.Max = LBound(smShow, 2)
    imSettingValue = False
    If imFdNmSelectedIndex < 0 Then
        lacType.Caption = ""
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String


    If ilBoxNo < LBound(tmCtrls) Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If
    If (imRowNo < vbcFeed.Value) Or (imRowNo >= vbcFeed.Value + vbcFeed.LargeChange + 1) Then
        mSetShow ilBoxNo
        Exit Sub
    End If
    lacFrame(imPaintIndex).Move 0, tmCtrls(1).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15) - 30
    lacFrame(imPaintIndex).Visible = True
    pbcArrow.Move pbcArrow.Left, plcFeed.Top + tmCtrls(1).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True

    Select Case imBoxNoMap(ilBoxNo) 'Branch on box type (control)
        Case IREFNOINDEX
            edcDropDown.Width = (3 * tmCtrls(ilBoxNo).fBoxW / 2)
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(1, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case IADVTINDEX
            mAdvtPop
            If imTerminate Then
                Exit Sub
            End If
            lbcAdvt.Height = gListBoxHeight(lbcAdvt.ListCount, 10)
            edcDropDown.Width = (3 * tmCtrls(ilBoxNo).fBoxW) / 2
            edcDropDown.MaxLength = 30
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If imRowNo - vbcFeed.Value <= vbcFeed.LargeChange \ 2 Then
                lbcAdvt.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcAdvt.Move edcDropDown.Left, edcDropDown.Top - lbcAdvt.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(2, imRowNo))
            If (slStr = "") And (imRowNo > 1) Then
                slStr = Trim$(smSave(2, imRowNo - 1))
            End If
            If slStr <> "" Then
                gFindMatch slStr, 0, lbcAdvt
                If gLastFound(lbcAdvt) > 0 Then
                    lbcAdvt.ListIndex = gLastFound(lbcAdvt)
                Else
                    lbcAdvt.ListIndex = 0
                End If
            Else
                lbcAdvt.ListIndex = 0
            End If
            If lbcAdvt.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcAdvt.List(lbcAdvt.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IPRODINDEX
            mProdPop imRowNo
            If imTerminate Then
                Exit Sub
            End If
            lbcProd.Height = gListBoxHeight(lbcProd.ListCount, 10)
            edcDropDown.Width = 2 * tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 35    'tgSpf.iAProd
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcProd.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If imRowNo - vbcFeed.Value <= vbcFeed.LargeChange \ 2 Then
                lbcProd.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcProd.Move edcDropDown.Left, edcDropDown.Top - lbcProd.Height
            End If
            imChgMode = True
            gFindMatch smSave(3, imRowNo), 1, lbcProd
            If gLastFound(lbcProd) >= 1 Then
                lbcProd.ListIndex = gLastFound(lbcProd)
                edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
            Else
                If imRowNo > 1 Then
                    If smSave(2, imRowNo) = smSave(2, imRowNo - 1) Then
                        If smSave(3, imRowNo - 1) <> "" Then
                            lbcProd.ListIndex = -1
                            edcDropDown.Text = smSave(3, imRowNo - 1)
                        Else
                            'If imProdFirstTime Then
                                gFindMatch tmAdf.sProduct, 1, lbcProd
                                If gLastFound(lbcProd) >= 1 Then
                                    lbcProd.ListIndex = gLastFound(lbcProd)
                                Else
                                    lbcProd.ListIndex = 1
                                End If
                            'Else
                            '    lbcProd.ListIndex = 0
                            'End If
                            edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
                        End If
                    Else
                        lbcProd.ListIndex = 1
                        edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
                    End If
                Else
                    gFindMatch tmAdf.sProduct, 1, lbcProd
                    If gLastFound(lbcProd) >= 1 Then
                        lbcProd.ListIndex = gLastFound(lbcProd)
                    Else
                        lbcProd.ListIndex = 1
                    End If
                    edcDropDown.Text = lbcProd.List(lbcProd.ListIndex)
                End If
            End If
            imChgMode = False
            'imProdFirstTime = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IPROT1INDEX
            mCompPop
            If imTerminate Then
                Exit Sub
            End If
            lbcComp(0).Height = gListBoxHeight(lbcComp(0).ListCount, 10)
            edcDropDown.Width = 4 * tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcComp(0).Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If imRowNo - vbcFeed.Value <= vbcFeed.LargeChange \ 2 Then
                lbcComp(0).Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcComp(0).Move edcDropDown.Left, edcDropDown.Top - lbcComp(0).Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(4, imRowNo))
            If slStr = "" Then
                mGetProd imRowNo
                slStr = smComp1
            End If
            If slStr <> "" Then
                gFindMatch slStr, 1, lbcComp(0)
                If gLastFound(lbcComp(0)) > 0 Then
                    lbcComp(0).ListIndex = gLastFound(lbcComp(0))
                Else
                    lbcComp(0).ListIndex = 1
                End If
            Else
                lbcComp(0).ListIndex = 1
            End If
            If lbcComp(0).ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcComp(0).List(lbcComp(0).ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IPROT2INDEX
            mCompPop
            If imTerminate Then
                Exit Sub
            End If
            lbcComp(1).Height = gListBoxHeight(lbcComp(1).ListCount, 10)
            edcDropDown.Width = 4 * tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 30
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcComp(1).Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If imRowNo - vbcFeed.Value <= vbcFeed.LargeChange \ 2 Then
                lbcComp(1).Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcComp(1).Move edcDropDown.Left, edcDropDown.Top - lbcComp(1).Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(5, imRowNo))
            If slStr = "" Then
                mGetProd imRowNo
                slStr = smComp2
            End If
            If slStr <> "" Then
                gFindMatch slStr, 1, lbcComp(1)
                If gLastFound(lbcComp(1)) > 0 Then
                    lbcComp(1).ListIndex = gLastFound(lbcComp(1))
                Else
                    lbcComp(1).ListIndex = 1
                End If
            Else
                lbcComp(1).ListIndex = 1
            End If
            If lbcComp(1).ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcComp(1).List(lbcComp(1).ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ILENINDEX
            lbcLen.Height = gListBoxHeight(lbcLen.ListCount, 10)
            edcDropDown.Width = (3 * tmCtrls(ilBoxNo).fBoxW) / 2
            edcDropDown.MaxLength = 30
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcLen.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If imRowNo - vbcFeed.Value <= vbcFeed.LargeChange \ 2 Then
                lbcLen.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcLen.Move edcDropDown.Left, edcDropDown.Top - lbcLen.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(6, imRowNo))
            If (slStr = "") And (imRowNo > 1) Then
                slStr = Trim$(smSave(6, imRowNo - 1))
            End If
            If slStr <> "" Then
                gFindMatch slStr, 0, lbcLen
                If gLastFound(lbcLen) > 0 Then
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
        Case ISDATEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.Height
            End If
            slStr = Trim$(smSave(7, imRowNo))
            If (slStr = "") And (imRowNo > 1) Then
                If (smSave(7, imRowNo - 1) <> "") Then
                    slStr = Trim$(smSave(7, imRowNo - 1))
                End If
            End If
            If slStr = "" Then
                'Set to beginning of the next week
                slStr = edcDate.Text
            End If
            edcDropDown.Text = slStr
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            If Trim$(smSave(7, imRowNo)) = "" Then
                pbcCalendar.Visible = True
            End If
            edcDropDown.SetFocus
        Case IEDATEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + 2 * cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.Height
            End If
            slStr = Trim$(smSave(8, imRowNo))
            If slStr = "" Then
                If imPaintIndex = 2 Then    'Insertion Order
                    slStr = gObtainNextSunday(smSave(7, imRowNo))
                Else
                    slStr = smSave(7, imRowNo)
                End If
            End If
            edcDropDown.Text = slStr
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            If Trim$(smSave(8, imRowNo)) = "" Then
                pbcCalendar.Visible = True
            End If
            edcDropDown.SetFocus
        Case IRUNINDEX
            lbcRun.Height = gListBoxHeight(lbcRun.ListCount, 10)
            edcDropDown.Width = 2 * tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 4
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcRun.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If imRowNo - vbcFeed.Value <= vbcFeed.LargeChange \ 2 Then
                lbcRun.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcRun.Move edcDropDown.Left, edcDropDown.Top - lbcRun.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(9, imRowNo))
            If slStr <> "" Then
                gFindMatch slStr, 0, lbcRun
                If gLastFound(lbcRun) > 0 Then
                    lbcRun.ListIndex = gLastFound(lbcRun)
                Else
                    lbcRun.ListIndex = 0
                End If
            Else
                lbcRun.ListIndex = 0
            End If
            If lbcRun.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcRun.List(lbcRun.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IDWINDEX
            lbcDW.Height = gListBoxHeight(lbcDW.ListCount, 10)
            edcDropDown.Width = 4 * tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 6
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcDW.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If imRowNo - vbcFeed.Value <= vbcFeed.LargeChange \ 2 Then
                lbcDW.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcDW.Move edcDropDown.Left, edcDropDown.Top - lbcDW.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(10, imRowNo))
            If slStr <> "" Then
                gFindMatch slStr, 0, lbcDW
                If gLastFound(lbcDW) > 0 Then
                    lbcDW.ListIndex = gLastFound(lbcDW)
                Else
                    If tgSpf.sAllowDailyBuys = "Y" Then
                        lbcDW.ListIndex = 1
                    Else
                        lbcDW.ListIndex = 0
                    End If
                End If
            Else
                If tgSpf.sAllowDailyBuys = "Y" Then
                    lbcDW.ListIndex = 1
                Else
                    lbcDW.ListIndex = 0
                End If
            End If
            If lbcDW.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcDW.List(lbcDW.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ISPOTSINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 3
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(11, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case IMOINDEX To ISUINDEX
            If smSave(10, imRowNo) = "Daily" Then
                edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
                edcDropDown.MaxLength = 3
                gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
                slStr = Trim$(smSave(12 + ilBoxNo - IMOINDEX, imRowNo))
                edcDropDown.Text = slStr
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
            Else
                gMoveTableCtrl pbcFeed(imPaintIndex), ckcAirDay, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
                slStr = Trim$(smSave(12 + ilBoxNo - IMOINDEX, imRowNo))
                If slStr = "" Or slStr = "Y" Then
                    ckcAirDay.Value = vbChecked
                Else
                    ckcAirDay.Value = vbUnchecked
                End If
                ckcAirDay.Visible = True
                ckcAirDay.SetFocus
            End If
        Case ISTIMEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 11
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            edcDropDown.Left = edcDropDown.Left - cmcDropDown.Width
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            slStr = Trim$(smSave(19, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            If imRowNo > UBound(smSave, 2) - 1 Then
                plcTme.Visible = True
            End If
            edcDropDown.SetFocus
        Case IETIMEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 11
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            edcDropDown.Left = edcDropDown.Left - cmcDropDown.Width
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - plcTme.Width, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - plcTme.Width, edcDropDown.Top - plcTme.Height
            End If
            slStr = Trim$(smSave(20, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            If imRowNo > UBound(smSave, 2) - 1 Then
                plcTme.Visible = True
            End If
            edcDropDown.SetFocus
        Case 21 'ISCI
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 20
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(21, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case 22 'Creative Title
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 20
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(22, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case 23 'Cart #
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 5
            gMoveTableCtrl pbcFeed(imPaintIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(23, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
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
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    imTerminate = False
    imFirstActivate = True
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165

    Screen.MousePointer = vbHourglass
    imLBLCtrls = 1
    imLBPCtrls = 1
    imLBICtrls = 1
    imLBCDCtrls = 1
    igJobShowing(FEEDJOB) = True
    imFirstActivate = True
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    FeedSpot.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    ReDim tmCtrls(1 To 10)
    mInitBox
    smNowDate = Format$(Now, "m/d/yy")
    gCenterForm FeedSpot
    FeedSpot.Show
    Screen.MousePointer = vbHourglass
    ReDim tgFsfRec(1 To 1) As FSFREC
    tgFsfRec(1).iStatus = -1
    tgFsfRec(1).lRecPos = 0
    ReDim tgFsfDel(1 To 1) As FSFREC
    tgFsfDel(1).iStatus = -1
    tgFsfDel(1).lRecPos = 0
    ReDim smShow(1 To 20, 1 To 1) As String 'Values shown in program area
    ReDim smSave(1 To 24, 1 To 1) As String 'Values saved (program name) in program area
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, 1) = ""
    Next ilLoop
'    mInitDDE
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    imFirstFocus = True
    imDoubleClickName = False
    imLbcMouseDown = False
    imCalType = 0               'Standard type
    imBoxNo = -1                'Initialize current Box to N/A
    imRowNo = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imChgMode = False
    imBSMode = False
    imBypassFocus = False
    imBypassSetting = False
    imVehSelectedIndex = -1
    imFsfChg = False
    hmFsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Fsf.Btr)", FeedSpot
    On Error GoTo 0
    imFsfRecLen = Len(tmFsf)

    hmFnf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmFnf, "", sgDBPath & "Fnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Fnf.Btr)", FeedSpot
    On Error GoTo 0
    imFnfRecLen = Len(tmFnf)


    hmAdf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", FeedSpot
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)

    hmPrf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Prf.Btr)", FeedSpot
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)

    hmMnf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", FeedSpot
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)

    hmMcf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mcf.Btr)", FeedSpot
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)

    hmCif = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", FeedSpot
    On Error GoTo 0
    imCifRecLen = Len(tmCif)

    hmCpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", FeedSpot
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)


    cbcFeed.Clear
    mFeedNamePop
    If imTerminate Then
        Exit Sub
    End If
    ilRet = gObtainVef()
    cbcVehicle.Clear 'Force list box to be populated
    mVehPop
    If imTerminate Then
        Exit Sub
    End If
    slDate = Format$(gNow(), "m/d/yy")
    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
    lbcAdvt.Clear
    mAdvtPop
    lbcComp(0).Clear
    lbcComp(1).Clear
    mCompPop
    lbcRun.Clear
    lbcRun.AddItem "Week"
    lbcRun.AddItem "2nd"
    lbcRun.AddItem "3rd"
    lbcRun.AddItem "4th"
    lbcDW.Clear
    If tgSpf.sAllowDailyBuys = "Y" Then
        lbcDW.AddItem "Daily"
    End If
    lbcDW.AddItem "Weekly"

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
'*             Created:5/17/93       By:D. LeVine      *
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
    flTextHeight = pbcFeed(imPaintIndex).TextHeight("1") - 35
    'Position panel and picture areas with panel
    'plcSelect.Move 3555, 120
    imPaintIndex = 0
    plcSelect.Move 1155, 15
    lacType.Move 195, 420
    plcFeed.Move 180, 615, pbcFeed(imPaintIndex).Width + fgPanelAdj + vbcFeed.Width, pbcFeed(imPaintIndex).Height + fgPanelAdj
    pbcFeed(imPaintIndex).Move plcFeed.Left + fgBevelX, plcFeed.Top + fgBevelY
    pbcFeed(1).Move pbcFeed(imPaintIndex).Left, pbcFeed(imPaintIndex).Top
    pbcFeed(2).Move pbcFeed(imPaintIndex).Left, pbcFeed(imPaintIndex).Top
    vbcFeed.Move pbcFeed(imPaintIndex).Left + pbcFeed(imPaintIndex).Width - 15, pbcFeed(imPaintIndex).Top


    'Ref #
    gSetCtrl tmICtrls(IREFNOINDEX), 30, 375, 615, fgBoxGridH
    'Advertiser
    gSetCtrl tmICtrls(IADVTINDEX), 660, tmICtrls(IREFNOINDEX).fBoxY, 1230, fgBoxGridH
    'Product
    gSetCtrl tmICtrls(IPRODINDEX), 1905, tmICtrls(IREFNOINDEX).fBoxY, 1005, fgBoxGridH
    'Protection 1
    gSetCtrl tmICtrls(IPROT1INDEX), 2925, tmICtrls(IREFNOINDEX).fBoxY, 420, fgBoxGridH
    'Protection 2
    gSetCtrl tmICtrls(IPROT2INDEX), 3360, tmICtrls(IREFNOINDEX).fBoxY, 420, fgBoxGridH
    'Length
    gSetCtrl tmICtrls(ILENINDEX), 3795, tmICtrls(IREFNOINDEX).fBoxY, 270, fgBoxGridH
    'Start Date
    gSetCtrl tmICtrls(ISDATEINDEX), 4080, tmICtrls(IREFNOINDEX).fBoxY, 510, fgBoxGridH
    'End Date
    gSetCtrl tmICtrls(IEDATEINDEX), 4605, tmICtrls(IREFNOINDEX).fBoxY, 510, fgBoxGridH
    'Run
    gSetCtrl tmICtrls(IRUNINDEX), 5130, tmICtrls(IREFNOINDEX).fBoxY, 285, fgBoxGridH
    'Daily/Weekly
    gSetCtrl tmICtrls(IDWINDEX), 5430, tmICtrls(IREFNOINDEX).fBoxY, 225, fgBoxGridH
    'Weekly Spots
    gSetCtrl tmICtrls(ISPOTSINDEX), 5670, tmICtrls(IREFNOINDEX).fBoxY, 510, fgBoxGridH
    'Monday Spots
    gSetCtrl tmICtrls(IMOINDEX), 6195, tmICtrls(IREFNOINDEX).fBoxY, 225, fgBoxGridH
    'Tuesday Spots
    gSetCtrl tmICtrls(ITUINDEX), 6435, tmICtrls(IREFNOINDEX).fBoxY, 225, fgBoxGridH
    'Wednesday Spots
    gSetCtrl tmICtrls(IWEINDEX), 6675, tmICtrls(IREFNOINDEX).fBoxY, 225, fgBoxGridH
    'Thursday Spots
    gSetCtrl tmICtrls(ITHINDEX), 6915, tmICtrls(IREFNOINDEX).fBoxY, 225, fgBoxGridH
    'Friday Spots
    gSetCtrl tmICtrls(IFRINDEX), 7155, tmICtrls(IREFNOINDEX).fBoxY, 225, fgBoxGridH
    'Saturday Spots
    gSetCtrl tmICtrls(ISAINDEX), 7395, tmICtrls(IREFNOINDEX).fBoxY, 225, fgBoxGridH
    'Sunday Spots
    gSetCtrl tmICtrls(ISUINDEX), 7635, tmICtrls(IREFNOINDEX).fBoxY, 225, fgBoxGridH
    'Start Time
    gSetCtrl tmICtrls(ISTIMEINDEX), 7875, tmICtrls(IREFNOINDEX).fBoxY, 510, fgBoxGridH
    'End Time
    gSetCtrl tmICtrls(IETIMEINDEX), 8400, tmICtrls(IREFNOINDEX).fBoxY, 510, fgBoxGridH

    'Advertiser
    gSetCtrl tmLCtrls(LADVTINDEX), 30, 375, 1665, fgBoxGridH
    'Product
    gSetCtrl tmLCtrls(LPRODINDEX), 1710, tmLCtrls(LADVTINDEX).fBoxY, 1125, fgBoxGridH
    'Protection 1
    gSetCtrl tmLCtrls(LPROT1INDEX), 2850, tmLCtrls(LADVTINDEX).fBoxY, 600, fgBoxGridH
    'Protection 2
    gSetCtrl tmLCtrls(LPROT2INDEX), 3465, tmLCtrls(LADVTINDEX).fBoxY, 600, fgBoxGridH
    'Length
    gSetCtrl tmLCtrls(LLENINDEX), 4080, tmLCtrls(LADVTINDEX).fBoxY, 270, fgBoxGridH
    'Date
    gSetCtrl tmLCtrls(LDATEINDEX), 4365, tmLCtrls(LADVTINDEX).fBoxY, 765, fgBoxGridH
    'Time
    gSetCtrl tmLCtrls(LTIMEINDEX), 5145, tmLCtrls(LADVTINDEX).fBoxY, 765, fgBoxGridH
    'ISCI
    gSetCtrl tmLCtrls(LISCIINDEX), 5925, tmLCtrls(LADVTINDEX).fBoxY, 1170, fgBoxGridH
    'Creative Title
    gSetCtrl tmLCtrls(LCREATIVEINDEX), 7110, tmLCtrls(LADVTINDEX).fBoxY, 1125, fgBoxGridH
    'Cart
    gSetCtrl tmLCtrls(LCARTINDEX), 8250, tmLCtrls(LADVTINDEX).fBoxY, 675, fgBoxGridH

    'Advertiser
    gSetCtrl tmPCtrls(PADVTINDEX), 30, 375, 1230, fgBoxGridH
    'Product
    gSetCtrl tmPCtrls(PPRODINDEX), 1275, tmPCtrls(PADVTINDEX).fBoxY, 1125, fgBoxGridH
    'Protection 1
    gSetCtrl tmPCtrls(PPROT1INDEX), 2415, tmPCtrls(PADVTINDEX).fBoxY, 420, fgBoxGridH
    'Protection 2
    gSetCtrl tmPCtrls(PPROT2INDEX), 2850, tmPCtrls(PADVTINDEX).fBoxY, 420, fgBoxGridH
    'Length
    gSetCtrl tmPCtrls(PLENINDEX), 3285, tmPCtrls(PADVTINDEX).fBoxY, 270, fgBoxGridH
    'Start Date
    gSetCtrl tmPCtrls(PSDATEINDEX), 3570, tmPCtrls(PADVTINDEX).fBoxY, 705, fgBoxGridH
    'End Date
    gSetCtrl tmPCtrls(PEDATEINDEX), 4290, tmPCtrls(PADVTINDEX).fBoxY, 705, fgBoxGridH
    'Time
    gSetCtrl tmPCtrls(PSTIMEINDEX), 5010, tmPCtrls(PADVTINDEX).fBoxY, 705, fgBoxGridH
    'Time
    gSetCtrl tmPCtrls(PETIMEINDEX), 5730, tmPCtrls(PADVTINDEX).fBoxY, 705, fgBoxGridH
    'ISCI
    gSetCtrl tmPCtrls(PISCIINDEX), 6450, tmPCtrls(PADVTINDEX).fBoxY, 1170, fgBoxGridH
    'Creative Title
    gSetCtrl tmPCtrls(PCREATIVEINDEX), 7635, tmPCtrls(PADVTINDEX).fBoxY, 600, fgBoxGridH
    'Cart
    gSetCtrl tmPCtrls(PCARTINDEX), 8250, tmPCtrls(PADVTINDEX).fBoxY, 675, fgBoxGridH


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

    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, ilRowNo) = ""
    Next ilLoop
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, ilRowNo) = ""
    Next ilLoop
    tgFsfRec(ilRowNo).iStatus = 0
    tgFsfRec(ilRowNo).lRecPos = 0
    tgFsfRec(ilRowNo).iDateChg = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitShow                       *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mInitShow()
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    For ilRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        For ilBoxNo = LBound(tmCtrls) To UBound(tmCtrls) Step 1
            Select Case imBoxNoMap(ilBoxNo)
                Case IREFNOINDEX
                    slStr = Trim$(smSave(1, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case IADVTINDEX
                    slStr = Trim$(smSave(2, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case IPRODINDEX
                    slStr = Trim$(smSave(3, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case IPROT1INDEX
                    slStr = Trim$(smSave(4, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case IPROT2INDEX
                    slStr = Trim$(smSave(5, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case ILENINDEX
                    slStr = Trim$(smSave(6, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case ISDATEINDEX
                    slStr = Trim$(smSave(7, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case IEDATEINDEX
                    slStr = Trim$(smSave(8, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case IRUNINDEX
                    slStr = Trim$(smSave(9, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case IDWINDEX
                    slStr = Trim$(smSave(10, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case ISPOTSINDEX
                    slStr = Trim$(smSave(11, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case IMOINDEX To ISUINDEX
                    slStr = smSave(12 + ilBoxNo - IMOINDEX, ilRowNo)
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case ISTIMEINDEX
                    slStr = Trim$(smSave(19, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case IETIMEINDEX
                    slStr = Trim$(smSave(20, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case 21
                    slStr = Trim$(smSave(21, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case 22
                    slStr = Trim$(smSave(22, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case 23
                    slStr = Trim$(smSave(23, ilRowNo))
                    gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
            End Select
        Next ilBoxNo
    Next ilRowNo
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
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilRowNo As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilFnfCode As Integer
    Dim ilDay As Integer
    Dim ilSDay As Integer
    Dim ilEDay As Integer

    slNameCode = tmFeedNameCode(imFdNmSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilFnfCode = Val(slCode)
    slNameCode = tmUserVehicle(imVehSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilVefCode = Val(slCode)

    For ilRowNo = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
        tgFsfRec(ilRowNo).tFsf.iFnfCode = ilFnfCode
        tgFsfRec(ilRowNo).tFsf.iVefCode = ilVefCode
        tgFsfRec(ilRowNo).tFsf.sRefID = smSave(1, ilRowNo)
        gFindMatch smSave(2, ilRowNo), 1, lbcAdvt
        If gLastFound(lbcAdvt) >= 1 Then
            slNameCode = tmFeedAdvertiser(gLastFound(lbcAdvt) - 1).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgFsfRec(ilRowNo).tFsf.iAdfCode = Val(slCode)
        Else
            tgFsfRec(ilRowNo).tFsf.iAdfCode = 0
        End If
        mProdPop ilRowNo
        gFindMatch smSave(3, ilRowNo), 2, lbcProd
        If gLastFound(lbcProd) >= 2 Then
            slNameCode = tgProdCode(gLastFound(lbcProd) - 2).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgFsfRec(ilRowNo).tFsf.lPrfCode = Val(slCode)
        Else
            tgFsfRec(ilRowNo).tFsf.lPrfCode = 0
        End If
        gFindMatch smSave(4, ilRowNo), 2, lbcComp(0)
        If gLastFound(lbcComp(0)) >= 2 Then
            slNameCode = tgCompCode(gLastFound(lbcComp(0)) - 2).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgFsfRec(ilRowNo).tFsf.iMnfComp1 = Val(slCode)
        Else
            tgFsfRec(ilRowNo).tFsf.iMnfComp1 = 0
        End If
        gFindMatch smSave(5, ilRowNo), 2, lbcComp(1)
        If gLastFound(lbcComp(1)) >= 2 Then
            slNameCode = tgCompCode(gLastFound(lbcComp(1)) - 2).sKey    'Traffic!lbcAdvt.List(gLastFound(lbcAdvt) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgFsfRec(ilRowNo).tFsf.iMnfComp2 = Val(slCode)
        Else
            tgFsfRec(ilRowNo).tFsf.iMnfComp2 = 0
        End If
        tgFsfRec(ilRowNo).tFsf.iLen = Val(smSave(6, ilRowNo))
        gPackDate smSave(7, ilRowNo), tgFsfRec(ilRowNo).tFsf.iStartDate(0), tgFsfRec(ilRowNo).tFsf.iStartDate(1)
        If imPaintIndex = 1 Then
            gPackDate smSave(7, ilRowNo), tgFsfRec(ilRowNo).tFsf.iEndDate(0), tgFsfRec(ilRowNo).tFsf.iEndDate(1)
        Else
            gPackDate smSave(8, ilRowNo), tgFsfRec(ilRowNo).tFsf.iEndDate(0), tgFsfRec(ilRowNo).tFsf.iEndDate(1)
        End If
        tgFsfRec(ilRowNo).tFsf.iRunEvery = Val(smSave(9, ilRowNo))
        If (imPaintIndex = 2) Then
            If (smSave(10, ilRowNo) = "Daily") Then
                tgFsfRec(ilRowNo).tFsf.sDyWk = "D"
                tgFsfRec(ilRowNo).tFsf.iNoSpots = 0
            Else
                tgFsfRec(ilRowNo).tFsf.sDyWk = "W"
                tgFsfRec(ilRowNo).tFsf.iNoSpots = Val(smSave(11, ilRowNo))
            End If
        End If
        If imPaintIndex = 0 Then
            tgFsfRec(ilRowNo).tFsf.sDyWk = "W"
            tgFsfRec(ilRowNo).tFsf.iNoSpots = 1
            For ilLoop = 0 To 6 Step 1
                tgFsfRec(ilRowNo).tFsf.iDays(ilLoop) = 0
            Next ilLoop
            ilSDay = gWeekDayStr(smSave(7, ilRowNo))
            ilEDay = gWeekDayStr(smSave(8, ilRowNo))
            For ilDay = ilSDay To ilEDay Step 1
                tgFsfRec(ilRowNo).tFsf.iDays(ilDay) = 1
            Next ilDay
        ElseIf imPaintIndex = 1 Then
            tgFsfRec(ilRowNo).tFsf.sDyWk = "W"
            tgFsfRec(ilRowNo).tFsf.iNoSpots = 1
            ilDay = gWeekDayStr(smSave(7, ilRowNo))
            For ilLoop = 0 To 6 Step 1
                If ilLoop <> ilDay Then
                    tgFsfRec(ilRowNo).tFsf.iDays(ilLoop) = 0
                Else
                    tgFsfRec(ilRowNo).tFsf.iDays(ilLoop) = 1
                End If
            Next ilLoop
        Else
            If tgFsfRec(ilRowNo).tFsf.sDyWk = "D" Then
                For ilLoop = 0 To 6 Step 1
                    tgFsfRec(ilRowNo).tFsf.iDays(ilLoop) = Val(smSave(12 + ilLoop, ilRowNo))
                Next ilLoop
            Else
                For ilLoop = 0 To 6 Step 1
                    If smSave(12 + ilLoop, ilRowNo) = "N" Then
                        tgFsfRec(ilRowNo).tFsf.iDays(ilLoop) = 0
                    Else
                        tgFsfRec(ilRowNo).tFsf.iDays(ilLoop) = 1
                    End If
                Next ilLoop
            End If
        End If
        gPackTime smSave(19, ilRowNo), tgFsfRec(ilRowNo).tFsf.iStartTime(0), tgFsfRec(ilRowNo).tFsf.iStartTime(1)
        If imPaintIndex = 1 Then
            gPackTime smSave(19, ilRowNo), tgFsfRec(ilRowNo).tFsf.iEndTime(0), tgFsfRec(ilRowNo).tFsf.iEndTime(1)
        Else
            gPackTime smSave(20, ilRowNo), tgFsfRec(ilRowNo).tFsf.iEndTime(0), tgFsfRec(ilRowNo).tFsf.iEndTime(1)
        End If
        tgFsfRec(ilRowNo).tFsf.lCifCode = 0
        slStr = Format(gNow(), "m/d/yy")
        gPackDate slStr, tgFsfRec(ilRowNo).tFsf.iEnterDate(0), tgFsfRec(ilRowNo).tFsf.iEnterDate(1)
        slStr = Format(gNow(), "h:mm:ssAM/PM")
        gPackTime slStr, tgFsfRec(ilRowNo).tFsf.iEnterTime(0), tgFsfRec(ilRowNo).tFsf.iEnterTime(1)
        tgFsfRec(ilRowNo).tFsf.sSchStatus = "N"
        tgFsfRec(ilRowNo).tFsf.iUrfCode = tgUrf(0).iCode
    Next ilRowNo
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
Private Sub mMoveRecToCtrl()
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilRowNo As Integer
    Dim ilUpper As Integer

    ilUpper = UBound(tgFsfRec)
    ReDim smShow(1 To 20, 1 To ilUpper) As String 'Values shown in program area
    ReDim smSave(1 To 24, 1 To ilUpper) As String 'Values saved (program name) in program area
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, ilUpper) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, ilUpper) = ""
    Next ilLoop
    'Init value in the case that no records are associated with the salesperson
    If ilUpper = LBound(tgFsfRec) Then
        ilRowNo = imRowNo
        imRowNo = 1
        mInitNew imRowNo
        imRowNo = ilRowNo
    End If
    For ilRowNo = LBound(tgFsfRec) To UBound(tgFsfRec) - 1 Step 1
        smSave(1, ilRowNo) = Trim$(tgFsfRec(ilRowNo).tFsf.sRefID)
        ilLoop = gBinarySearchAdf(tgFsfRec(ilRowNo).tFsf.iAdfCode)
        If ilLoop <> -1 Then
            smSave(2, ilRowNo) = Trim$(tgCommAdf(ilLoop).sName)
            mGetAdvt smSave(2, ilRowNo)
        End If
        smSave(3, ilRowNo) = ""
        mProdPop ilRowNo
        slRecCode = Trim$(Str$(tgFsfRec(ilRowNo).tFsf.lPrfCode))
        For ilLoop = 0 To UBound(tgProdCode) - 1 Step 1 'lbcInvSortCode.ListCount - 1 Step 1
            slNameCode = tgProdCode(ilLoop).sKey   'lbcInvSortCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", FeedSpot
            On Error GoTo 0
            If slRecCode = slCode Then
                smSave(3, ilRowNo) = lbcProd.List(ilLoop + 2)
                Exit For
            End If
        Next ilLoop

        smSave(4, ilRowNo) = ""
        slRecCode = Trim$(Str$(tgFsfRec(ilRowNo).tFsf.iMnfComp1))
        For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1 'lbcCompCode.ListCount - 1 Step 1
            slNameCode = tgCompCode(ilLoop).sKey   'lbcCompCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", FeedSpot
            On Error GoTo 0
            If slRecCode = slCode Then
                smSave(4, ilRowNo) = lbcComp(0).List(ilLoop + 2)
                Exit For
            End If
        Next ilLoop
        smSave(5, ilRowNo) = ""
        slRecCode = Trim$(Str$(tgFsfRec(ilRowNo).tFsf.iMnfComp2))
        For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1 'lbcCompCode.ListCount - 1 Step 1
            slNameCode = tgCompCode(ilLoop).sKey   'lbcCompCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", FeedSpot
            On Error GoTo 0
            If slRecCode = slCode Then
                smSave(5, ilRowNo) = lbcComp(1).List(ilLoop + 2)
                Exit For
            End If
        Next ilLoop
        smSave(6, ilRowNo) = Trim$(Str$(tgFsfRec(ilRowNo).tFsf.iLen))
        'Get Start Date
        gUnpackDate tgFsfRec(ilRowNo).tFsf.iStartDate(0), tgFsfRec(ilRowNo).tFsf.iStartDate(1), smSave(7, ilRowNo)
        'Get End Date
        gUnpackDate tgFsfRec(ilRowNo).tFsf.iEndDate(0), tgFsfRec(ilRowNo).tFsf.iEndDate(1), smSave(8, ilRowNo)
        smSave(9, ilRowNo) = lbcRun.List(tgFsfRec(ilRowNo).tFsf.iRunEvery)
        If tgFsfRec(ilRowNo).tFsf.sDyWk = "D" Then
            smSave(10, ilRowNo) = "Daily"
        Else
            smSave(10, ilRowNo) = "Weekly"
        End If
        smSave(11, ilRowNo) = Trim$(Str$(tgFsfRec(ilRowNo).tFsf.iNoSpots))
        If tgFsfRec(ilRowNo).tFsf.sDyWk = "D" Then
            For ilLoop = 0 To 6 Step 1
                smSave(12 + ilLoop, ilRowNo) = Trim$(Str$(tgFsfRec(ilRowNo).tFsf.iDays(ilLoop)))
            Next ilLoop
        Else
            For ilLoop = 0 To 6 Step 1
                If tgFsfRec(ilRowNo).tFsf.iDays(ilLoop) = 0 Then
                    smSave(12 + ilLoop, ilRowNo) = "N"
                Else
                    smSave(12 + ilLoop, ilRowNo) = "Y"
                End If
            Next ilLoop
        End If
        'Get Start Date
        gUnpackTime tgFsfRec(ilRowNo).tFsf.iStartTime(0), tgFsfRec(ilRowNo).tFsf.iStartTime(1), "A", "1", smSave(19, ilRowNo)
        'Get End Date
        gUnpackTime tgFsfRec(ilRowNo).tFsf.iEndTime(0), tgFsfRec(ilRowNo).tFsf.iEndTime(1), "A", "1", smSave(20, ilRowNo)
        'ISCI, Creative Title and Cart missing
        smSave(21, ilRowNo) = ""
        smSave(22, ilRowNo) = ""
        smSave(23, ilRowNo) = ""
        'If (tgFsfRec(ilRowNo).tFsf.lCifCode > 0) And (tmMcf.iCode > 0) Then
        If (tgFsfRec(ilRowNo).tFsf.lCifCode > 0) Then
            tmCifSrchKey0.lCode = tgFsfRec(ilRowNo).tFsf.lCifCode
            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                If (tgSpf.sUseCartNo <> "N") And (tmMcf.iCode > 0) Then
                    smSave(23, ilRowNo) = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                    If (Len(Trim$(tmCif.sCut)) <> 0) Then
                        smSave(23, ilRowNo) = smSave(23, ilRowNo) & "-" & tmCif.sCut
                    End If
                End If
                tmCpfSrchKey0.lCode = tmCif.lCpfCode
                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    smSave(21, ilRowNo) = Trim$(tmCpf.sISCI)
                    smSave(22, ilRowNo) = Trim$(tmCpf.sCreative)
                End If
            End If
        End If
        'Schedule status
        smSave(24, ilRowNo) = tgFsfRec(ilRowNo).tFsf.sSchStatus
    Next ilRowNo
    Exit Sub
mMoveRecToCtrlErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec() As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFnfCode As Integer
    Dim ilVefCode As Integer
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim llNoRec As Long
    Dim ilOffset As Integer
    Dim slStr As String
    Dim slDate As String
    Dim llTime As Long
    Dim slTime As String
    Dim ilFsf As Integer
    Dim ilPass As Integer
    Dim ilIndex As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    ReDim tgFsfRec(1 To 1) As FSFREC
    tgFsfRec(1).iStatus = -1
    tgFsfRec(1).lRecPos = 0
    tgFsfRec(1).iDateChg = False
    ReDim tgFsfDel(1 To 1) As FSFREC
    tgFsfDel(1).iStatus = -1
    tgFsfDel(1).lRecPos = 0
    ilUpper = 1
    If imFdNmSelectedIndex < 0 Then
        mReadRec = False
        Exit Function
    End If
    If imVehSelectedIndex < 0 Then
        mReadRec = False
        Exit Function
    End If
    mLenPop
    slDate = Trim$(edcDate.Text)
    If Not gValidDate(slDate) Or slDate = "" Then
        mReadRec = False
        Exit Function
    End If
    For ilPass = 0 To 1 Step 1
        slDate = Trim$(edcDate.Text)
        btrExtClear hmFsf   'Clear any previous extend operation
        ilExtLen = Len(tgFsfRec(1).tFsf)  'Extract operation record size
        slNameCode = tmFeedNameCode(imFdNmSelectedIndex).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilFnfCode = Val(slCode)
        slNameCode = tmUserVehicle(imVehSelectedIndex).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)

        tmFsfSrchKey3.iFnfCode = ilFnfCode
        tmFsfSrchKey3.iVefCode = ilVefCode
        'gPackDate slDate, tmFsfSrchKey3.iStartDate(0), tmFsfSrchKey3.iStartDate(1)
        gPackDate "", tmFsfSrchKey3.iStartDate(0), tmFsfSrchKey3.iStartDate(1)
        ilRet = btrGetGreaterOrEqual(hmFsf, tmFsf, imFsfRecLen, tmFsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
            ilRet = BTRV_ERR_END_OF_FILE
        Else
            If ilRet <> BTRV_ERR_NONE Then
                mReadRec = False
                Exit Function
            End If
        End If
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
            Call btrExtSetBounds(hmFsf, llNoRec, -1, "UC", "FSF", "") '"EG") 'Set extract limits (all records)
            ilOffset = gFieldOffset("FSF", "FSFFNFCODE") 'GetOffSetForInt(tmFsf, tmFsf.iSlfCode)
            tlIntTypeBuff.iType = ilFnfCode
            ilRet = btrExtAddLogicConst(hmFsf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Fsf.Btr", FeedSpot
            On Error GoTo 0
            ilOffset = gFieldOffset("FSF", "FSFVEFCODE") 'GetOffSetForInt(tmFsf, tmFsf.iSlfCode)
            tlIntTypeBuff.iType = ilVefCode
            ilRet = btrExtAddLogicConst(hmFsf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Fsf.Btr", FeedSpot
            On Error GoTo 0
            If ilPass = 0 Then
                ilOffset = gFieldOffset("FSF", "FsfStartDate")
                gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                ilRet = btrExtAddLogicConst(hmFsf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                On Error GoTo mReadRecErr
                gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Fsf.Btr", FeedSpot
                On Error GoTo 0
                ilOffset = gFieldOffset("FSF", "FsfEndDate")
                gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                ilRet = btrExtAddLogicConst(hmFsf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                On Error GoTo mReadRecErr
                gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Fsf.Btr", FeedSpot
                On Error GoTo 0
            Else
                ilOffset = gFieldOffset("FSF", "FsfStartDate")
                gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                ilRet = btrExtAddLogicConst(hmFsf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlDateTypeBuff, 4)
                On Error GoTo mReadRecErr
                gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Fsf.Btr", FeedSpot
                On Error GoTo 0
                ilOffset = gFieldOffset("FSF", "FsfEndDate")
                slDate = gDecOneDay(slDate)
                gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                ilRet = btrExtAddLogicConst(hmFsf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                On Error GoTo mReadRecErr
                gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Fsf.Btr", FeedSpot
                On Error GoTo 0
            End If
            ilRet = btrExtAddField(hmFsf, 0, ilExtLen) 'Extract the whole record
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrExtAddField):" & "Fsf.Btr", FeedSpot
            On Error GoTo 0
            'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
            ilRet = btrExtGetNext(hmFsf, tgFsfRec(ilUpper).tFsf, ilExtLen, tgFsfRec(ilUpper).lRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo mReadRecErr
                gBtrvErrorMsg ilRet, "mReadRec (btrExtGetNextExt):" & "Fsf.Btr", FeedSpot
                On Error GoTo 0
                'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
                ilExtLen = Len(tgFsfRec(1).tFsf)  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmFsf, tgFsfRec(ilUpper).tFsf, ilExtLen, tgFsfRec(ilUpper).lRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    slStr = ""
                    gUnpackDateForSort tgFsfRec(ilUpper).tFsf.iStartDate(0), tgFsfRec(ilUpper).tFsf.iStartDate(1), slDate
                    gUnpackTimeLong tgFsfRec(ilUpper).tFsf.iStartTime(0), tgFsfRec(ilUpper).tFsf.iStartTime(1), False, llTime
                    slTime = Trim$(Str$(llTime))
                    Do While Len(slTime) < 5
                        slTime = "0" & slTime
                    Loop
                    tgFsfRec(ilUpper).sKey = slDate & slTime
                    tgFsfRec(ilUpper).iStatus = 1
                    ilUpper = ilUpper + 1
                    ReDim Preserve tgFsfRec(1 To ilUpper) As FSFREC
                    tgFsfRec(ilUpper).iStatus = -1
                    tgFsfRec(ilUpper).lRecPos = 0
                    ilRet = btrExtGetNext(hmFsf, tgFsfRec(ilUpper).tFsf, ilExtLen, tgFsfRec(ilUpper).lRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmFsf, tgFsfRec(ilUpper).tFsf, ilExtLen, tgFsfRec(ilUpper).lRecPos)
                    Loop
                Loop
            End If
        End If
    Next ilPass
    'Retain latest version only
    For ilLoop = 1 To UBound(tgFsfRec) - 1 Step 1
        If tgFsfRec(ilLoop).tFsf.lPrevFsfCode = 0 Then
            'Loop for record that references it
            ilFsf = ilLoop
            ilIndex = mFindNextVer(tgFsfRec(ilLoop).tFsf.lCode)
            Do While ilIndex > 0
                'Mark for deletion
                tgFsfRec(ilFsf).iStatus = -2
                ilFsf = ilIndex
                ilIndex = mFindNextVer(tgFsfRec(ilIndex).tFsf.lCode)
            Loop
        End If
    Next ilLoop
    'Remove all tagged records
    ilFsf = UBound(tgFsfRec) - 1
    Do While ilFsf >= 1
        If tgFsfRec(ilFsf).iStatus = -2 Then
            For ilLoop = ilFsf To UBound(tgFsfRec) - 1 Step 1
                tgFsfRec(ilLoop) = tgFsfRec(ilLoop + 1)
            Next ilLoop
            If UBound(tgFsfRec) > 1 Then
                ReDim Preserve tgFsfRec(1 To UBound(tgFsfRec) - 1) As FSFREC
            End If
        End If
        ilFsf = ilFsf - 1
    Loop
    'Remove cancel before start
    If (Trim$(tgUrf(0).sName) = sgCPName) Or ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
    Else
        'remove CBS
        ilFsf = UBound(tgFsfRec) - 1
        Do While ilFsf >= 1
            gUnpackDateLong tgFsfRec(ilFsf).tFsf.iStartDate(0), tgFsfRec(ilFsf).tFsf.iStartDate(1), llStartDate
            gUnpackDateLong tgFsfRec(ilFsf).tFsf.iEndDate(0), tgFsfRec(ilFsf).tFsf.iEndDate(1), llEndDate
            If llEndDate < llStartDate Then
                For ilLoop = ilFsf To UBound(tgFsfRec) - 1 Step 1
                    tgFsfRec(ilLoop) = tgFsfRec(ilLoop + 1)
                Next ilLoop
                If UBound(tgFsfRec) > 1 Then
                    ReDim Preserve tgFsfRec(1 To UBound(tgFsfRec) - 1) As FSFREC
                End If
            End If
            ilFsf = ilFsf - 1
        Loop
    End If
    ilUpper = UBound(tgFsfRec)
    If ilUpper > 1 Then
        ArraySortTyp fnAV(tgFsfRec(), 1), UBound(tgFsfRec) - 1, 0, LenB(tgFsfRec(1)), 0, LenB(tgFsfRec(1).sKey), 0
    End If
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
    Dim ilLoop As Integer   'For loop control
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilRowNo As Integer
    Dim slMsg As String
    Dim ilFsf As Integer
    Dim tlFsf As FSF
    Dim tlSvFsf As FSF
    Dim tlFsf1 As MOVEREC
    Dim tlFsf2 As MOVEREC
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    For ilRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        If mTestSaveFields(ilRowNo) = NO Then
            mSaveRec = False
            imRowNo = ilRowNo
            Exit Function
        End If
    Next ilRowNo
    mMoveCtrlToRec
    Screen.MousePointer = vbHourglass  'Wait
    If imTerminate Then
        Screen.MousePointer = vbDefault    'Default
        mSaveRec = False
        Exit Function
    End If
    ilRet = btrBeginTrans(hmFsf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 1", vbOkOnly + vbExclamation, "Invoice")
        Exit Function
    End If
    ilLoop = 0
    For ilFsf = LBound(tgFsfRec) To UBound(tgFsfRec) - 1 Step 1
        Do  'Loop until record updated or added
            If (tgFsfRec(ilFsf).iStatus = 0) Then  'New selected
                'User
                tgFsfRec(ilFsf).tFsf.lCode = 0
                tgFsfRec(ilFsf).tFsf.iRevNo = 0
                tgFsfRec(ilFsf).tFsf.lPrevFsfCode = 0
                tgFsfRec(ilFsf).tFsf.lAstCode = 0
                ilRet = btrInsert(hmFsf, tgFsfRec(ilFsf).tFsf, imFsfRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmFsf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 2", vbOkOnly + vbExclamation, "Invoice")
                    Exit Function
                End If
                slMsg = "mSaveRec (btrInsert: Feed)"
                ilRet = btrGetPosition(hmFsf, tgFsfRec(ilFsf).lRecPos)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmFsf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 3", vbOkOnly + vbExclamation, "Invoice")
                    Exit Function
                End If
                tgFsfRec(ilFsf).iStatus = 1
            ElseIf (tgFsfRec(ilFsf).iStatus = 1) Then  'Old record-Update
                slMsg = "mSaveRec (btrGetDirect: Feed)"
                tmFSFSrchKey.lCode = tgFsfRec(ilFsf).tFsf.lCode
                ilRet = btrGetEqual(hmFsf, tlFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmFsf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 4", vbOkOnly + vbExclamation, "Invoice")
                    Exit Function
                End If
                'Set values that should not be compared
                tlSvFsf = tlFsf
                tlFsf.sSchStatus = tgFsfRec(ilFsf).tFsf.sSchStatus
                tlFsf.iUrfCode = tgFsfRec(ilFsf).tFsf.iUrfCode
                tlFsf.sRefID = tgFsfRec(ilFsf).tFsf.sRefID
                tlFsf.iEnterDate(0) = tgFsfRec(ilFsf).tFsf.iEnterDate(0)
                tlFsf.iEnterDate(1) = tgFsfRec(ilFsf).tFsf.iEnterDate(1)
                tlFsf.iEnterTime(0) = tgFsfRec(ilFsf).tFsf.iEnterTime(0)
                tlFsf.iEnterTime(1) = tgFsfRec(ilFsf).tFsf.iEnterTime(1)
                tlFsf.sUnused = tgFsfRec(ilFsf).tFsf.sUnused
                tlFsf.iRevNo = tgFsfRec(ilFsf).tFsf.iRevNo
                tlFsf1 = tlFsf2
                tlFsf1 = tlFsf
                tlFsf2 = tgFsfRec(ilFsf).tFsf

                If StrComp(tlFsf1.sChar, tlFsf2.sChar, 0) <> 0 Then
                    If tlSvFsf.sSchStatus = "F" Then
                        tgFsfRec(ilFsf).tFsf.lCode = 0
                        tgFsfRec(ilFsf).tFsf.iRevNo = tlSvFsf.iRevNo + 1
                        tgFsfRec(ilFsf).tFsf.lPrevFsfCode = tlSvFsf.lCode
                        tgFsfRec(ilFsf).tFsf.iUrfCode = tgUrf(0).iCode
                        ilRet = btrInsert(hmFsf, tgFsfRec(ilFsf).tFsf, imFsfRecLen, INDEXKEY0)
                    Else
                        tgFsfRec(ilFsf).tFsf.iRevNo = tlSvFsf.iRevNo
                        tgFsfRec(ilFsf).tFsf.sSchStatus = tlSvFsf.sSchStatus
                        tgFsfRec(ilFsf).tFsf.lPrevFsfCode = tlSvFsf.lPrevFsfCode
                        tgFsfRec(ilFsf).tFsf.iUrfCode = tgUrf(0).iCode
                        ilRet = btrUpdate(hmFsf, tgFsfRec(ilFsf).tFsf, imFsfRecLen)
                    End If
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrUpdate: Feed)"
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            ilCRet = btrAbortTrans(hmFsf)
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 5", vbOkOnly + vbExclamation, "Invoice")
            Exit Function
        End If
    Next ilFsf
    For ilFsf = LBound(tgFsfDel) To UBound(tgFsfDel) - 1 Step 1
        If tgFsfDel(ilFsf).iStatus = 1 Then
            Do
                slMsg = "mSaveRec (btrGetEqual: Feed)"
                tmFSFSrchKey.lCode = tgFsfDel(ilFsf).tFsf.lCode
                ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmFsf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 6", vbOkOnly + vbExclamation, "Invoice")
                    Exit Function
                End If
                If tmFsf.sSchStatus = "N" Then
                    ilRet = btrDelete(hmFsf)
                    slMsg = "mSaveRec (btrDelete: Feed)"
                Else
                    tgFsfRec(ilFsf).tFsf.sSchStatus = "D"
                    ilRet = btrUpdate(hmFsf, tgFsfRec(ilFsf).tFsf, imFsfRecLen)
                    slMsg = "mSaveRec (btrUpdate: Feed-Deleted)"
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                ilCRet = btrAbortTrans(hmFsf)
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & Str$(ilRet) & " at 7", vbOkOnly + vbExclamation, "Invoice")
                Exit Function
            End If
        End If
    Next ilFsf
    ilRet = btrEndTrans(hmFsf)
    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function

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
'*             Created:9/24/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
    Dim ilNew As Integer
    If imFsfChg And (UBound(tgFsfRec) > LBound(tgFsfRec)) Or (UBound(tgFsfDel) > LBound(tgFsfDel)) Then
        If ilAsk Then
            ilNew = True
            For ilLoop = LBound(tgFsfRec) To UBound(tgFsfRec) - 1 Step 1
                If tgFsfRec(ilLoop).iStatus <> 0 Then
                    ilNew = False
                    Exit For
                End If
            Next ilLoop
            For ilLoop = LBound(tgFsfDel) To UBound(tgFsfDel) - 1 Step 1
                If tgFsfDel(ilLoop).iStatus <> 0 Then
                    ilNew = False
                    Exit For
                End If
            Next ilLoop
            If Not ilNew Then
                slMess = "Save Changes"
            Else
                slMess = "Add Changes"
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
    If ilBoxNo < LBound(tmCtrls) Or ilBoxNo > UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
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
    If (imBypassSetting) Or (Not imUpdateAllowed) Then
        Exit Sub
    End If
    If (UBound(tgFsfRec) <= LBound(tgFsfRec)) And (UBound(tgFsfDel) <= LBound(tgFsfDel)) Then
        cmcSchedule.Enabled = True
        Exit Sub
    End If
    ilAltered = imFsfChg
    If (Not ilAltered) And (UBound(tgFsfDel) > LBound(tgFsfDel)) Then
        ilAltered = True
    End If
    If ilAltered Then
        pbcFeed(imPaintIndex).Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        cbcFeed.Enabled = False
        cbcVehicle.Enabled = False
        cmcDate.Enabled = False
        edcDate.Enabled = False
    Else
        If (imFdNmSelectedIndex < 0) Or (imVehSelectedIndex < 0) Then
            pbcFeed(imPaintIndex).Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            cbcFeed.Enabled = True
            cbcVehicle.Enabled = True
            cmcDate.Enabled = True
            edcDate.Enabled = True
        Else
            pbcFeed(imPaintIndex).Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            cbcFeed.Enabled = True
            cbcVehicle.Enabled = True
            cmcDate.Enabled = True
            edcDate.Enabled = True
        End If
    End If
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields() = YES) And (ilAltered) And (UBound(tgFsfRec) > 1) And (imUpdateAllowed) Then
        cmcSave.Enabled = True
        cmcSchedule.Enabled = False
        cmcPledge.Enabled = False
    Else
        cmcSave.Enabled = False
        cmcSchedule.Enabled = True
        If (imPaintIndex = 1) And (imFdNmSelectedIndex >= 0) And (imVehSelectedIndex >= 0) Then
            cmcPledge.Enabled = True
        Else
            cmcPledge.Enabled = False
        End If
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
    If (ilBoxNo < LBound(tmCtrls)) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case imBoxNoMap(ilBoxNo) 'Branch on box type (control)
        Case IREFNOINDEX
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case IADVTINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IPRODINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IPROT1INDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IPROT2INDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ILENINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ISDATEINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IEDATEINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IRUNINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IDWINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ISPOTSINDEX
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case IMOINDEX To ISUINDEX
            If smSave(10, imRowNo) = "Daily" Then
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
            Else
                ckcAirDay.Visible = True
                ckcAirDay.SetFocus
            End If
        Case ISTIMEINDEX
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IETIMEINDEX
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case 21 'ISCI
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case 22 'Creative Title
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case 23 'Cart #
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
    End Select
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
    vbcFeed.Min = LBound(smShow, 2)
    imSettingValue = True
    If UBound(smShow, 2) - 1 <= vbcFeed.LargeChange + 1 Then ' + 1 Then
        vbcFeed.Max = LBound(smShow, 2)
    Else
        vbcFeed.Max = UBound(smShow, 2) - vbcFeed.LargeChange
    End If
    imSettingValue = True
    If vbcFeed.Value = vbcFeed.Min Then
        vbcFeed_Change
    Else
        vbcFeed.Value = vbcFeed.Min
    End If
    imSettingValue = False
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
    pbcArrow.Visible = False
    lacFrame(imPaintIndex).Visible = False
    If (ilBoxNo < LBound(tmCtrls)) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case imBoxNoMap(ilBoxNo) 'Branch on box type (control)
        Case IREFNOINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(1, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(1, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If

        Case IADVTINDEX
            lbcAdvt.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(2, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(2, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If
        Case IPRODINDEX
            lbcProd.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(3, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(3, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If
        Case IPROT1INDEX
            lbcComp(0).Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(4, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(4, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If
        Case IPROT2INDEX
            lbcComp(1).Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(5, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(5, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If
        Case ILENINDEX
            lbcLen.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(6, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(6, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If
        Case ISDATEINDEX
            plcCalendar.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidDate(slStr) Then
                gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
                If Trim$(smSave(7, imRowNo)) <> slStr Then
                    imFsfChg = True
                    smSave(7, imRowNo) = slStr
                    tgFsfRec(imRowNo).iDateChg = True
                    smSave(24, imRowNo) = ""
                End If
            End If
        Case IEDATEINDEX
            plcCalendar.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidDate(slStr) Then
                gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
                If Trim$(smSave(8, imRowNo)) <> slStr Then
                    imFsfChg = True
                    smSave(8, imRowNo) = slStr
                    tgFsfRec(imRowNo).iDateChg = True
                    smSave(24, imRowNo) = ""
                End If
            End If
        Case IRUNINDEX
            lbcRun.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(9, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(9, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If
        Case IDWINDEX
            lbcDW.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(10, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(10, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If
        Case ISPOTSINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(11, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(11, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If
        Case IMOINDEX To ISUINDEX
            If smSave(10, imRowNo) = "Daily" Then
                edcDropDown.Visible = False
                slStr = edcDropDown.Text
                gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
                If Trim$(smSave(12 + ilBoxNo - IMOINDEX, imRowNo)) <> slStr Then
                    imFsfChg = True
                    smSave(12 + ilBoxNo - IMOINDEX, imRowNo) = slStr
                    smSave(24, imRowNo) = ""
                End If
            Else
                ckcAirDay.Visible = False
                If ckcAirDay.Value = vbChecked Then
                    If smSave(12 + ilBoxNo - IMOINDEX, imRowNo) <> "Y" Then
                        imFsfChg = True
                        smSave(12 + ilBoxNo - IMOINDEX, imRowNo) = "Y"
                        smSave(24, imRowNo) = ""
                    End If
                Else
                    If smSave(12 + ilBoxNo - IMOINDEX, imRowNo) <> "N" Then
                        imFsfChg = True
                        smSave(12 + ilBoxNo - IMOINDEX, imRowNo) = "N"
                        smSave(24, imRowNo) = ""
                    End If
                End If
                slStr = smSave(12 + ilBoxNo - IMOINDEX, imRowNo)
                gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            End If
        Case ISTIMEINDEX
            plcTme.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidTime(slStr) Then
                gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
                If Trim$(smSave(19, imRowNo)) <> slStr Then
                    imFsfChg = True
                    smSave(19, imRowNo) = slStr
                    smSave(24, imRowNo) = ""
                End If
            End If
        Case IETIMEINDEX
            plcTme.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidTime(slStr) Then
                gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
                If Trim$(smSave(20, imRowNo)) <> slStr Then
                    imFsfChg = True
                    smSave(20, imRowNo) = slStr
                    smSave(24, imRowNo) = ""
                End If
            End If
        Case 21
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(21, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(21, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If
        Case 22
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(22, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(22, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If
        Case 23
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcFeed(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(23, imRowNo)) <> slStr Then
                imFsfChg = True
                smSave(23, imRowNo) = slStr
                smSave(24, imRowNo) = ""
            End If

    End Select
    mSetCommands
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
    Dim ilRet As Integer
    Erase tgFsfRec
    Erase tgFsfDel

    Erase tmFeedNameCode
    smFeedNameCodeTag = ""

    Erase tmFeedAdvertiser
    smFeedtAdvertiserTag = ""

    Erase tmUserVehicle
    smUserVehicleTag = ""

    btrExtClear hmCpf   'Clear any previous extend operation
    ilRet = btrClose(hmCpf)
	btrDestroy hmCpf

    btrExtClear hmCif   'Clear any previous extend operation
    ilRet = btrClose(hmCif)
	btrDestroy hmCif

    btrExtClear hmMcf   'Clear any previous extend operation
    ilRet = btrClose(hmMcf)
	btrDestroy hmMcf
	
	btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrClose(hmMnf)
	btrDestroy hmMnf
	
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
	btrDestroy hmAdf

    btrExtClear hmPrf   'Clear any previous extend operation
    ilRet = btrClose(hmPrf)
	btrDestroy hmPrf

    btrExtClear hmFnf   'Clear any previous extend operation
    ilRet = btrClose(hmFnf)
	btrDestroy hmFnf

    btrExtClear hmFsf   'Clear any previous extend operation
    ilRet = btrClose(hmFsf)
	btrDestroy hmFsf

    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload FeedSpot
    Set FeedSpot = Nothing
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
    Dim ilSpotFound As Integer
    Dim ilDayFound As Integer
    Dim ilLoop As Integer

    For ilRowNo = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
        If Trim$(smSave(2, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSave(6, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSave(7, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSave(19, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If (imPaintIndex = 2) Or (imPaintIndex = 0) Then
            If Trim$(smSave(8, ilRowNo)) = "" Then
                mTestFields = NO
                Exit Function
            End If
            If Trim$(smSave(20, ilRowNo)) = "" Then
                mTestFields = NO
                Exit Function
            End If
        End If
        If (imPaintIndex = 2) Then
            If Trim$(smSave(9, ilRowNo)) = "" Then
                mTestFields = NO
                Exit Function
            End If
            If Trim$(smSave(10, ilRowNo)) = "" Then
                mTestFields = NO
                Exit Function
            End If
            If Trim$(smSave(10, ilRowNo)) = "Daily" Then
                ilSpotFound = False
                For ilLoop = 0 To 6 Step 1
                    If Trim$(smSave(12 + ilLoop, ilRowNo)) <> "" Then
                        ilSpotFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilSpotFound Then
                    mTestFields = NO
                    Exit Function
                End If
            Else
                If Trim$(smSave(11, ilRowNo)) = "" Then
                    mTestFields = NO
                    Exit Function
                End If
                ilDayFound = False
                For ilLoop = 0 To 6 Step 1
                    If Trim$(smSave(12 + ilLoop, ilRowNo)) <> "" Then
                        ilDayFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilDayFound Then
                    mTestFields = NO
                    Exit Function
                End If
            End If
        End If
        'Temporary patch to disallow save if copy defined
        If (Trim$(smSave(21, ilRowNo)) <> "") Or (Trim$(smSave(22, ilRowNo)) <> "") Or (Trim$(smSave(23, ilRowNo)) <> "") Then
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
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields(ilRowNo As Integer) As Integer
'
'   iRet = mTestSaveFields(ilRowNo)
'   Where:
'       ilRowNo (I)- row number to be checked
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilLoop As Integer
    Dim ilSpotFound As Integer
    Dim ilDayFound As Integer

    'Test date
    If (Not gValidDate(smSave(7, ilRowNo))) Or (Trim$(smSave(7, ilRowNo)) = "") Then
        Beep
        ilRes = MsgBox("Start Date must be specified correctly", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = ISDATEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If (Not gValidTime(smSave(19, ilRowNo))) Or (Trim$(smSave(19, ilRowNo)) = "") Then
        Beep
        ilRes = MsgBox("Start Time must be specified correctly", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = ISTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If (imPaintIndex = 2) Or (imPaintIndex = 0) Then
        If (Not gValidDate(smSave(8, ilRowNo))) Or (Trim$(smSave(8, ilRowNo)) = "") Then
            Beep
            ilRes = MsgBox("End Date must be specified correctly", vbOkOnly + vbExclamation, "Incomplete")
            imBoxNo = ISDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
        If (Not gValidTime(smSave(20, ilRowNo))) Or (Trim$(smSave(20, ilRowNo)) = "") Then
            Beep
            ilRes = MsgBox("End Time must be specified correctly", vbOkOnly + vbExclamation, "Incomplete")
            imBoxNo = ISTIMEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
        If gDateValue(smSave(7, ilRowNo)) > gDateValue(smSave(8, ilRowNo)) Then
            Beep
            ilRes = MsgBox("Start Date must be prior to End Date", vbOkOnly + vbExclamation, "Incomplete")
            imBoxNo = ISDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
        'Check times
        If gTimeToLong(smSave(19, ilRowNo), False) > gTimeToLong(smSave(20, ilRowNo), True) Then
            Beep
            ilRes = MsgBox("Start Time must be prior to End Time", vbOkOnly + vbExclamation, "Incomplete")
            imBoxNo = ISTIMEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If (imPaintIndex = 0) Then
        If gWeekDayStr(smSave(8, ilRowNo)) < gWeekDayStr(smSave(7, ilRowNo)) Then
            Beep
            ilRes = MsgBox("Dates can't cross week", vbOkOnly + vbExclamation, "Incomplete")
            imBoxNo = ISDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If Trim$(smSave(2, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("Advertiser must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = IADVTINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smSave(6, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("Spot Length must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = ILENINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If (imPaintIndex = 2) Then
        If Trim$(smSave(9, ilRowNo)) = "" Then
            Beep
            ilRes = MsgBox("Run Every must be specified", vbOkOnly + vbExclamation, "Incomplete")
            imBoxNo = IRUNINDEX
            mTestSaveFields = NO
            Exit Function
        End If
        If Trim$(smSave(10, ilRowNo)) = "" Then
            Beep
            ilRes = MsgBox("Daily/Weekly must be specified", vbOkOnly + vbExclamation, "Incomplete")
            imBoxNo = IDWINDEX
            mTestSaveFields = NO
            Exit Function
        End If
        If Trim$(smSave(10, ilRowNo)) = "Daily" Then
            ilSpotFound = False
            For ilLoop = 0 To 6 Step 1
                If Trim$(smSave(12 + ilLoop, ilRowNo)) <> "" Then
                    ilSpotFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilSpotFound Then
                Beep
                ilRes = MsgBox("Daily Spots must be specified", vbOkOnly + vbExclamation, "Incomplete")
                imBoxNo = IMOINDEX
                mTestSaveFields = NO
                Exit Function
            End If
        Else
            If Trim$(smSave(11, ilRowNo)) = "" Then
                Beep
                ilRes = MsgBox("Number of Spots per Week must be specified", vbOkOnly + vbExclamation, "Incomplete")
                imBoxNo = ISPOTSINDEX
                mTestSaveFields = NO
                Exit Function
            End If
            ilDayFound = False
            For ilLoop = 0 To 6 Step 1
                If smSave(12 + ilLoop, ilRowNo) <> "N" Then
                    ilDayFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilDayFound Then
                Beep
                ilRes = MsgBox("Week Day must be specified", vbOkOnly + vbExclamation, "Incomplete")
                imBoxNo = IMOINDEX
                mTestSaveFields = NO
                Exit Function
            End If
        End If
    End If

    mTestSaveFields = YES
End Function
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
                If imBoxNo <> -1 Then
                    edcDropDown.Text = Format$(llDate, "m/d/yy")
                    edcDropDown.SelStart = 0
                    edcDropDown.SelLength = Len(edcDropDown.Text)
                    imBypassFocus = True
                    edcDropDown.SetFocus
                    Exit Sub
                Else
                    edcDate.Text = Format$(llDate, "m/d/yy")
                    edcDate.SelStart = 0
                    edcDate.SelLength = Len(edcDate.Text)
                    imBypassFocus = True
                    edcDate.SetFocus
                    Exit Sub
                End If
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate

    If imBoxNo <> -1 Then
        edcDropDown.SetFocus
    Else
        edcDate.SetFocus
    End If
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(Str$(imCalMonth)) & "/15/" & Trim$(Str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    plcCalendar.Visible = False
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
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

Private Sub pbcFeed_GotFocus(Index As Integer)
    If imBoxNo = -1 Then
        plcCalendar.Visible = False
    End If
End Sub

Private Sub pbcFeed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim slDate As String

    If imFdNmSelectedIndex < 0 Then
        Exit Sub
    End If
    If imVehSelectedIndex < 0 Then
        Exit Sub
    End If
    slDate = Trim$(edcDate.Text)
    If Not gValidDate(slDate) Or slDate = "" Then
        Exit Sub
    End If
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcFeed_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim slDate As String

    If Button = 2 Then
        Exit Sub
    End If
    If imFdNmSelectedIndex < 0 Then
        Exit Sub
    End If
    If imVehSelectedIndex < 0 Then
        Exit Sub
    End If
    slDate = Trim$(edcDate.Text)
    If Not gValidDate(slDate) Or slDate = "" Then
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    ilCompRow = vbcFeed.LargeChange + 1
    If UBound(tgFsfRec) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tgFsfRec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = LBound(tmCtrls) To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcFeed.Value - 1
                    If ilRowNo > UBound(smSave, 2) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If imPaintIndex = 2 Then
                        If (ilBox > 1) And (Trim$(smSave(2, ilRowNo)) = "") Then
                            Beep
                            mSetFocus imBoxNo
                            Exit Sub
                        End If
                        If ilBox = IRUNINDEX Then
                            If gDateValue(smSave(8, ilRowNo)) - gDateValue(smSave(7, ilRowNo) < 14) Then
                                Beep
                                mSetFocus imBoxNo
                                Exit Sub
                            End If
                        End If
                        If (smSave(10, ilRowNo) = "Daily") And (ilBox = ISPOTSINDEX) Then
                            Beep
                            mSetFocus imBoxNo
                            Exit Sub
                        End If
                    Else
                        If (ilBox > 1) And (Trim$(smSave(2, ilRowNo)) = "") Then
                            Beep
                            mSetFocus imBoxNo
                            Exit Sub
                        End If
                    End If
                    mSetShow imBoxNo
                    imRowNo = ilRow + vbcFeed.Value - 1
                    If (imRowNo = UBound(smSave, 2)) And (Trim$(smSave(2, imRowNo)) = "") Then
                        mInitNew imRowNo
                    End If
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mSetFocus imBoxNo
End Sub
Private Sub pbcFeed_Paint(Index As Integer)
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long
    Dim llSDate As Long
    Dim llEDate As Long

    pbcFeed(Index).Cls
    ilStartRow = vbcFeed.Value '+ 1  'Top location
    ilEndRow = vbcFeed.Value + vbcFeed.LargeChange ' + 1
    If ilEndRow > UBound(smSave, 2) Then
        If imPaintIndex = 2 Then
            If (Trim$(smShow(1, UBound(smShow, 2))) <> "") Or (Trim$(smShow(2, UBound(smShow, 2))) <> "") Then
                ilEndRow = UBound(smSave, 2) 'include blank row as it might have data
            Else
                ilEndRow = UBound(smSave, 2) - 1
            End If
        Else
            If Trim$(smShow(1, UBound(smShow, 2))) <> "" Then
                ilEndRow = UBound(smSave, 2) 'include blank row as it might have data
            Else
                ilEndRow = UBound(smSave, 2) - 1
            End If
        End If
    End If
    llColor = pbcFeed(Index).ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        If ilRow = UBound(smSave, 2) Then
            pbcFeed(Index).ForeColor = DARKPURPLE
        Else
            llSDate = gDateValue(smSave(7, ilRow))
            llEDate = gDateValue(smSave(8, ilRow))
            If llEDate < llSDate Then
                pbcFeed(Index).ForeColor = vbRed
            Else
                If smSave(24, ilRow) = "F" Then
                    pbcFeed(Index).ForeColor = DARKGRAY
                Else
                    pbcFeed(Index).ForeColor = llColor
                End If
            End If
        End If
        For ilBox = LBound(tmCtrls) To UBound(tmCtrls) Step 1
            pbcFeed(imPaintIndex).CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcFeed(imPaintIndex).CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            slStr = Trim$(smShow(ilBox, ilRow))
            pbcFeed(imPaintIndex).Print slStr
        Next ilBox
    Next ilRow
    pbcFeed(Index).ForeColor = llColor
End Sub

Private Sub pbcNum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    imcNumInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 4 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            For ilColNo = 1 To 3 Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcNumInv.Move flX, flY
                    imcNumInv.Visible = True
                    imcNumOutline.Move flX - 15, flY - 15
                    imcNumOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcNum_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    imcNumInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 4 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            For ilColNo = 1 To 3 Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcNumInv.Move flX, flY
                    imcNumOutline.Move flX - 15, flY - 15
                    imcNumOutline.Visible = True
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
                                Case 3
                                    slKey = "."
                            End Select
                    End Select
                    imBypassFocus = True    'Don't change select text
                    edcDropDown.SetFocus
                    'SendKeys slKey
                    gSendKeys edcDropDown, slKey
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    Dim slDate As String

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If imFdNmSelectedIndex < 0 Then
        Exit Sub
    End If
    If imVehSelectedIndex < 0 Then
        Exit Sub
    End If
    slDate = Trim$(edcDate.Text)
    If Not gValidDate(slDate) Or slDate = "" Then
        Exit Sub
    End If
    If imBoxNo = -1 Then
        plcCalendar.Visible = False
    End If
    imTabDirection = -1 'Set- Right to left
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            plcCalendar.Visible = False
            If (UBound(smSave, 2) = 1) Then
                imTabDirection = 0  'Set-Left to right
                imRowNo = 1
                mInitNew imRowNo
            Else
                If UBound(smSave, 2) <= vbcFeed.LargeChange Then 'was <=
                    vbcFeed.Max = LBound(smSave, 2)
                Else
                    vbcFeed.Max = UBound(smSave, 2) - vbcFeed.LargeChange '- 1
                End If
                imRowNo = 1
                If imRowNo >= UBound(smSave, 2) Then
                    mInitNew imRowNo
                End If
                imSettingValue = True
                vbcFeed.Value = vbcFeed.Min
                imSettingValue = False
            End If
            ilBox = 1
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case 1, 0
            mSetShow imBoxNo
            If (imBoxNo < 1) And (imRowNo < 1) Then 'Modelled from Proposal
                Exit Sub
            End If
            ilBox = UBound(tmCtrls)   'Show be max
            If imRowNo <= 1 Then
                imBoxNo = -1
                imRowNo = -1
                cmcDone.SetFocus
                Exit Sub
            End If
            imRowNo = imRowNo - 1
            If imRowNo < vbcFeed.Value Then
                imSettingValue = True
                vbcFeed.Value = vbcFeed.Value - 1
                imSettingValue = False
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
    End Select
    If imBoxNoMap(imBoxNo) = IADVTINDEX Then
        If mAdvtBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNoMap(imBoxNo) = IPRODINDEX Then
        If mProdBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNoMap(imBoxNo) = IPROT1INDEX Then
        If mCompBranch(0) Then
            Exit Sub
        End If
    End If
    If imBoxNoMap(imBoxNo) = IPROT2INDEX Then
        If mCompBranch(1) Then
            Exit Sub
        End If
    End If
    Select Case imBoxNoMap(imBoxNo)
        Case ISDATEINDEX
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
            ilBox = imBoxNo - 1
        Case IEDATEINDEX
            slStr = edcDropDown.Text
            If (slStr <> "") Then
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
            ilBox = imBoxNo - 1
        Case IDWINDEX
            If gDateValue(smSave(8, imRowNo)) - gDateValue(smSave(7, imRowNo) >= 14) Then
                ilBox = imBoxNo - 1
            Else
                ilBox = imBoxNo - 2
                If smSave(9, imRowNo) = "" Then
                    smSave(9, imRowNo) = lbcRun.List(0)
                End If
            End If
        Case IMOINDEX
            If smSave(10, imRowNo) = "Daily" Then
                ilBox = imBoxNo - 2
            Else
                ilBox = imBoxNo - 1
            End If
        Case ISTIMEINDEX
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
            ilBox = imBoxNo - 1
        Case IETIMEINDEX
            slStr = edcDropDown.Text
            If (slStr <> "") Then
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
            ilBox = imBoxNo - 1
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDate As String

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If imFdNmSelectedIndex < 0 Then
        Exit Sub
    End If
    If imVehSelectedIndex < 0 Then
        Exit Sub
    End If
    slDate = Trim$(edcDate.Text)
    If Not gValidDate(slDate) Or slDate = "" Then
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            imRowNo = UBound(smSave, 2) - 1
            imSettingValue = True
            If imRowNo <= vbcFeed.LargeChange + 1 Then
                vbcFeed.Value = 1
            Else
                vbcFeed.Value = imRowNo - vbcFeed.LargeChange - 1
            End If
            imSettingValue = False
            ilBox = 1
        Case UBound(tmCtrls)
            mSetShow imBoxNo
            If mTestSaveFields(imRowNo) = NO Then
                mEnableBox imBoxNo
                Exit Sub
            End If
            If imRowNo >= UBound(smSave, 2) Then
                imFsfChg = True
                ReDim Preserve smShow(1 To 20, 1 To imRowNo + 1) As String 'Values shown in program area
                ReDim Preserve smSave(1 To 24, 1 To imRowNo + 1) As String 'Values saved (program name) in program area
                For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                    smShow(ilLoop, imRowNo + 1) = ""
                Next ilLoop
                For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                    smSave(ilLoop, imRowNo + 1) = ""
                Next ilLoop
                ReDim Preserve tgFsfRec(1 To UBound(tgFsfRec) + 1) As FSFREC
                tgFsfRec(UBound(tgFsfRec)).iStatus = 0
                tgFsfRec(UBound(tgFsfRec)).lRecPos = 0
            End If
            If imRowNo >= UBound(smSave, 2) - 1 Then
                imRowNo = imRowNo + 1
                mInitNew imRowNo
                If UBound(smSave, 2) <= vbcFeed.LargeChange Then 'was <=
                    vbcFeed.Max = LBound(smSave, 2) '- 1
                Else
                    vbcFeed.Max = UBound(smSave, 2) - vbcFeed.LargeChange '- 1
                End If
            Else
                imRowNo = imRowNo + 1
            End If
            If imRowNo > vbcFeed.Value + vbcFeed.LargeChange Then
                imSettingValue = True
                vbcFeed.Value = vbcFeed.Value + 1
                imSettingValue = False
            End If
            If imRowNo >= UBound(smSave, 2) Then
                imBoxNo = 0
                mSetCommands
                'lacFrame(imPaintIndex).Move 0, tmCtrls(PROPNOINDEX).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15) - 30
                'lacFrame(imPaintIndex).Visible = True
                pbcArrow.Move pbcArrow.Left, plcFeed.Top + tmCtrls(1).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15) + 45
                pbcArrow.Visible = True
                pbcArrow.SetFocus
                Exit Sub
            Else
                ilBox = 1
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case 0
            ilBox = 1
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
    End Select
    If imBoxNo > 0 Then
        If imBoxNoMap(imBoxNo) = IADVTINDEX Then
            If mAdvtBranch() Then
                Exit Sub
            End If
        End If
        If imBoxNoMap(imBoxNo) = IPRODINDEX Then
            If mProdBranch() Then
                Exit Sub
            End If
        End If
        If imBoxNoMap(imBoxNo) = IPROT1INDEX Then
            If mCompBranch(0) Then
                Exit Sub
            End If
        End If
        If imBoxNoMap(imBoxNo) = IPROT2INDEX Then
            If mCompBranch(1) Then
                Exit Sub
            End If
        End If
        Select Case imBoxNoMap(imBoxNo)
            Case ISDATEINDEX
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
                ilBox = imBoxNo + 1
            Case IEDATEINDEX
                slStr = edcDropDown.Text
                If (slStr <> "") Then
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
                If imPaintIndex = 2 Then
                    If gDateValue(slStr) - gDateValue(smSave(7, imRowNo)) >= 14 Then
                        ilBox = imBoxNo + 1
                    Else
                        ilBox = imBoxNo + 2
                        If smSave(9, imRowNo) = "" Then
                            smSave(9, imRowNo) = lbcRun.List(0)
                        End If
                    End If
                Else
                    ilBox = imBoxNo + 1
                End If
            Case IDWINDEX
                If edcDropDown.Text = "Daily" Then
                    ilBox = imBoxNo + 2
                Else
                    ilBox = imBoxNo + 1
                End If
            Case ISTIMEINDEX
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
                ilBox = imBoxNo + 1
            Case IETIMEINDEX
                slStr = edcDropDown.Text
                If (slStr <> "") Then
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
                ilBox = imBoxNo + 1
            Case Else
                ilBox = imBoxNo + 1
        End Select
        mSetShow imBoxNo
        imBoxNo = ilBox
        mEnableBox ilBox
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
                    Select Case imBoxNo
                        Case ISTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                         Case IETIMEINDEX
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

Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNoMap(imBoxNo)
        Case IADVTINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcAdvt, edcDropDown, imChgMode, imLbcArrowSetting
        Case IPRODINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcProd, edcDropDown, imChgMode, imLbcArrowSetting
        Case IPROT1INDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcComp(0), edcDropDown, imChgMode, imLbcArrowSetting
        Case IPROT2INDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcComp(1), edcDropDown, imChgMode, imLbcArrowSetting
    End Select

End Sub


Private Sub tmcDrag_Timer()
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Dim slDate As String

    tmcDrag.Enabled = False
    If imFdNmSelectedIndex < 0 Then
        Exit Sub
    End If
    If imVehSelectedIndex < 0 Then
        Exit Sub
    End If
    slDate = Trim$(edcDate.Text)
    If Not gValidDate(slDate) Or slDate = "" Then
        Exit Sub
    End If
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            ilCompRow = vbcFeed.LargeChange + 1
            If UBound(smSave, 2) > ilCompRow Then
                ilMaxRow = ilCompRow
            Else
                ilMaxRow = UBound(smSave, 2)
            End If
            For ilRow = 1 To ilMaxRow Step 1
                If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(1).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(1).fBoxY + tmCtrls(1).fBoxH)) Then
                    mSetShow imBoxNo
                    imBoxNo = -1
                    imRowNo = -1
                    imRowNo = ilRow + vbcFeed.Value - 1
                    lacFrame(imPaintIndex).DragIcon = IconTraf!imcIconStd.DragIcon
                    lacFrame(imPaintIndex).Move 0, tmCtrls(1).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15) - 30
                    'If gInvertArea call then remove visible setting
                    lacFrame(imPaintIndex).Visible = True
                    pbcArrow.Move pbcArrow.Left, plcFeed.Top + tmCtrls(1).fBoxY + (imRowNo - vbcFeed.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    lacFrame(imPaintIndex).Drag vbBeginDrag
                    lacFrame(imPaintIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
                    Exit Sub
                End If
            Next ilRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub vbcFeed_Change()
    If imSettingValue Then
        pbcFeed(imPaintIndex).Cls
        pbcFeed_Paint imPaintIndex
        imSettingValue = False
    Else
        mSetShow imBoxNo
        imBoxNo = -1
        imRowNo = -1
        pbcFeed(imPaintIndex).Cls
        pbcFeed_Paint imPaintIndex
    End If
End Sub
Private Sub vbcFeed_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Feed"
End Sub

Private Sub mFeedNamePop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffset(0) As Integer

    imPopReqd = False
    ilfilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffset(0) = 0
    'ilRet = gIMoveListBox(FeedName, cbcVehicle, lbcNameCode, "Fnf.btr", gFieldOffset("Fnf", "FnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(FeedSpot, cbcFeed, tmFeedNameCode(), smFeedNameCodeTag, "Fnf.btr", gFieldOffset("Fnf", "FnfName"), 20, ilfilter(), slFilter(), ilOffset())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mFeedNameErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", FeedSpot
        On Error GoTo 0
        imPopReqd = True
    End If
    Exit Sub
mFeedNameErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mBuildMap()
    If imPaintIndex = 1 Then
        imBoxNoMap(LADVTINDEX) = IADVTINDEX
        imBoxNoMap(LPRODINDEX) = IPRODINDEX
        imBoxNoMap(LPROT1INDEX) = IPROT1INDEX
        imBoxNoMap(LPROT2INDEX) = IPROT2INDEX
        imBoxNoMap(LLENINDEX) = ILENINDEX
        imBoxNoMap(LDATEINDEX) = ISDATEINDEX
        imBoxNoMap(LTIMEINDEX) = ISTIMEINDEX
        imBoxNoMap(LISCIINDEX) = 21
        imBoxNoMap(LCREATIVEINDEX) = 22
        imBoxNoMap(LCARTINDEX) = 23
    ElseIf imPaintIndex = 2 Then
        imBoxNoMap(IREFNOINDEX) = IREFNOINDEX
        imBoxNoMap(IADVTINDEX) = IADVTINDEX
        imBoxNoMap(IPRODINDEX) = IPRODINDEX
        imBoxNoMap(IPROT1INDEX) = IPROT1INDEX
        imBoxNoMap(IPROT2INDEX) = IPROT2INDEX
        imBoxNoMap(ILENINDEX) = ILENINDEX
        imBoxNoMap(ISDATEINDEX) = ISDATEINDEX
        imBoxNoMap(IEDATEINDEX) = IEDATEINDEX
        imBoxNoMap(IRUNINDEX) = IRUNINDEX
        imBoxNoMap(IDWINDEX) = IDWINDEX
        imBoxNoMap(ISPOTSINDEX) = ISPOTSINDEX
        imBoxNoMap(IMOINDEX) = IMOINDEX
        imBoxNoMap(ITUINDEX) = ITUINDEX
        imBoxNoMap(IWEINDEX) = IWEINDEX
        imBoxNoMap(ITHINDEX) = ITHINDEX
        imBoxNoMap(IFRINDEX) = IFRINDEX
        imBoxNoMap(ISAINDEX) = ISAINDEX
        imBoxNoMap(ISUINDEX) = ISUINDEX
        imBoxNoMap(ISTIMEINDEX) = ISTIMEINDEX
        imBoxNoMap(IETIMEINDEX) = IETIMEINDEX
    Else
        imBoxNoMap(PADVTINDEX) = IADVTINDEX
        imBoxNoMap(PPRODINDEX) = IPRODINDEX
        imBoxNoMap(PPROT1INDEX) = IPROT1INDEX
        imBoxNoMap(PPROT2INDEX) = IPROT2INDEX
        imBoxNoMap(PLENINDEX) = ILENINDEX
        imBoxNoMap(PSDATEINDEX) = ISDATEINDEX
        imBoxNoMap(PEDATEINDEX) = IEDATEINDEX
        imBoxNoMap(PSTIMEINDEX) = ISTIMEINDEX
        imBoxNoMap(PETIMEINDEX) = IETIMEINDEX
        imBoxNoMap(PISCIINDEX) = 21
        imBoxNoMap(PCREATIVEINDEX) = 22
        imBoxNoMap(PCARTINDEX) = 23
    End If
End Sub

Private Function mFindNextVer(llFsfCode As Long) As Integer
    Dim ilLoop As Integer
    For ilLoop = LBound(tgFsfRec) To UBound(tgFsfRec) - 1 Step 1
        If llFsfCode = tgFsfRec(ilLoop).tFsf.lPrevFsfCode Then
            mFindNextVer = ilLoop
            Exit Function
        End If
    Next ilLoop
    mFindNextVer = 0
End Function
