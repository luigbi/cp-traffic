VERSION 5.00
Begin VB.Form PEvent 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   105
   ClientTop       =   1365
   ClientWidth     =   9360
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   9360
   Begin VB.PictureBox pbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   15
      Picture         =   "Pevent.frx":0000
      ScaleHeight     =   525
      ScaleWidth      =   3285
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   585
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   15
      Picture         =   "Pevent.frx":5B96
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1365
      Visible         =   0   'False
      Width           =   105
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
      Left            =   7365
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   3555
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
         Picture         =   "Pevent.frx":5EA0
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "Pevent.frx":6B5E
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
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   5475
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8085
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   5475
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcLibType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3240
      ScaleHeight     =   180
      ScaleWidth      =   1185
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lbcExcl 
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
      Index           =   0
      Left            =   -225
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3855
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcExcl 
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
      Index           =   1
      Left            =   -495
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3150
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox plclen 
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
      Left            =   7035
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2070
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcLen 
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
         Picture         =   "Pevent.frx":6E68
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcLenInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "Pevent.frx":7B26
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image imcLenOutline 
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
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1320
      Top             =   5460
   End
   Begin VB.PictureBox pbcIconMove 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragIcon        =   "Pevent.frx":7E30
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4380
      ScaleHeight     =   165
      ScaleWidth      =   150
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   45
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox pbcIconTrash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragIcon        =   "Pevent.frx":813A
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3990
      ScaleHeight     =   165
      ScaleWidth      =   150
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   45
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox pbcIconStd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3615
      ScaleHeight     =   165
      ScaleWidth      =   150
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2175
      Top             =   60
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   45
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5820
      Width           =   45
   End
   Begin VB.ListBox lbcEvtNameCode 
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
      Index           =   0
      Left            =   1335
      Sorted          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   270
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.ListBox lbcEvtAvail 
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
      Left            =   5130
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1245
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcEvtName 
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
      Index           =   0
      Left            =   1905
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1125
      Visible         =   0   'False
      Width           =   3210
   End
   Begin VB.ListBox lbcEvtType 
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
      Left            =   -480
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2295
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.TextBox edcComment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   540
      HelpContextID   =   8
      Left            =   2415
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   1110
      Visible         =   0   'False
      Width           =   5625
   End
   Begin VB.PictureBox pbcTrueTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6315
      ScaleHeight     =   210
      ScaleWidth      =   375
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmcEvtDropDown 
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
      Left            =   2205
      Picture         =   "Pevent.frx":8444
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1575
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcEvtDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1170
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox edcEvtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   315
      MaxLength       =   20
      TabIndex        =   17
      Top             =   1650
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.ListBox lbcLibName 
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   2220
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
      Left            =   5400
      Picture         =   "Pevent.frx":853E
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcSpecDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4380
      MaxLength       =   20
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   60
      ScaleHeight     =   105
      ScaleWidth      =   75
      TabIndex        =   28
      Top             =   5415
      Width           =   75
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   105
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   14
      Top             =   885
      Width           =   60
   End
   Begin VB.PictureBox pbcSpecTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   45
      ScaleHeight     =   45
      ScaleWidth      =   90
      TabIndex        =   12
      Top             =   855
      Width           =   90
   End
   Begin VB.PictureBox pbcSpecSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   525
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   5
      Top             =   345
      Width           =   60
   End
   Begin VB.PictureBox pbcLibSpec 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   2940
      Picture         =   "Pevent.frx":8638
      ScaleHeight     =   375
      ScaleWidth      =   6240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   540
      Width           =   6240
   End
   Begin VB.VScrollBar vbcEvents 
      Height          =   4260
      LargeChange     =   19
      Left            =   8955
      TabIndex        =   29
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox pbcEvents 
      Appearance      =   0  'Flat
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
      Height          =   4290
      Left            =   225
      Picture         =   "Pevent.frx":B172
      ScaleHeight     =   4290
      ScaleWidth      =   8730
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1065
      Width           =   8730
      Begin VB.Label lacEvtFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -315
         TabIndex        =   36
         Top             =   2265
         Visible         =   0   'False
         Width           =   8700
      End
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
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   3015
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   3015
      Begin VB.TextBox edcLinkDestDoneMsg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2205
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.PictureBox plcSelect 
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   3090
      ScaleHeight     =   360
      ScaleWidth      =   6090
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   6150
      Begin VB.CheckBox ckcShowVersion 
         Caption         =   "Show All Versions"
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
         Height          =   195
         Left            =   1815
         TabIndex        =   3
         Top             =   75
         Width           =   1845
      End
      Begin VB.ComboBox cbcType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   60
         TabIndex        =   2
         Top             =   15
         Width           =   1650
      End
      Begin VB.ComboBox cbcSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   3660
         TabIndex        =   4
         Top             =   15
         Width           =   2430
      End
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
      Left            =   1875
      TabIndex        =   30
      Top             =   5685
      Width           =   945
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
      Left            =   2970
      TabIndex        =   31
      Top             =   5685
      Width           =   945
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
      HelpContextID   =   3
      Left            =   4065
      TabIndex        =   32
      Top             =   5685
      Width           =   945
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
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
      HelpContextID   =   4
      Left            =   5160
      TabIndex        =   33
      Top             =   5685
      Width           =   945
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
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
      HelpContextID   =   5
      Left            =   6255
      TabIndex        =   34
      Top             =   5685
      Width           =   945
   End
   Begin VB.PictureBox plcLibSpec 
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   165
      ScaleHeight     =   420
      ScaleWidth      =   9075
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   495
      Width           =   9135
      Begin VB.OptionButton rbcView 
         Caption         =   "Clock View"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1500
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   135
         Width           =   1290
      End
      Begin VB.OptionButton rbcView 
         Caption         =   "Tabular View"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   48
         Top             =   135
         Width           =   1455
      End
      Begin VB.CommandButton cmcDates 
         Appearance      =   0  'Flat
         Caption         =   "D&ates"
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
         HelpContextID   =   6
         Left            =   8070
         TabIndex        =   13
         Top             =   90
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.PictureBox pbcEatTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Index           =   0
      Left            =   135
      ScaleHeight     =   45
      ScaleWidth      =   60
      TabIndex        =   44
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox pbcEatTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   1
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   30
      TabIndex        =   45
      Top             =   0
      Width           =   30
   End
   Begin VB.PictureBox pbcClock 
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
      Height          =   4275
      Left            =   240
      ScaleHeight     =   4245
      ScaleWidth      =   8940
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1065
      Visible         =   0   'False
      Width           =   8970
      Begin VB.PictureBox pbcCEvents 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2475
         Left            =   4425
         Picture         =   "Pevent.frx":33E74
         ScaleHeight     =   2475
         ScaleWidth      =   1980
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   750
         Width           =   1980
      End
      Begin VB.PictureBox plcHour 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3135
         ScaleHeight     =   165
         ScaleWidth      =   870
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   90
         Width           =   930
      End
      Begin VB.HScrollBar hbcHour 
         Height          =   240
         Left            =   4080
         TabIndex        =   55
         Top             =   90
         Width           =   1695
      End
      Begin VB.PictureBox pbcCLibType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   6990
         ScaleHeight     =   180
         ScaleWidth      =   1185
         TabIndex        =   54
         Top             =   4005
         Width           =   1215
      End
      Begin VB.ListBox lbcCLibName 
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
         Left            =   6690
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2205
         Width           =   2055
      End
      Begin VB.ListBox lbcCEvtType 
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
         Left            =   6690
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lacColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Other Events"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   5
         Left            =   5085
         TabIndex        =   61
         Top             =   3945
         Width           =   1215
      End
      Begin VB.Label lacColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Other Avail"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Index           =   4
         Left            =   3825
         TabIndex        =   60
         Top             =   3945
         Width           =   1020
      End
      Begin VB.Label lacColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "BB Avail"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   5085
         TabIndex        =   59
         Top             =   3750
         Width           =   810
      End
      Begin VB.Label lacColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Promo Avail"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Index           =   2
         Left            =   3825
         TabIndex        =   58
         Top             =   3750
         Width           =   1065
      End
      Begin VB.Label lacColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "PSA Avail"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   5085
         TabIndex        =   57
         Top             =   3540
         Width           =   900
      End
      Begin VB.Label lacColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Contract Avail"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   0
         Left            =   3825
         TabIndex        =   56
         Top             =   3540
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcEvents 
      ForeColor       =   &H00000000&
      Height          =   4395
      Left            =   165
      ScaleHeight     =   4335
      ScaleWidth      =   9030
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1020
      Width           =   9090
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   30
      Picture         =   "Pevent.frx":39396
      Top             =   270
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   5520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imcCalc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      Picture         =   "Pevent.frx":396A0
      Top             =   5535
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8790
      Picture         =   "Pevent.frx":39882
      Top             =   5490
      Width           =   480
   End
End
Attribute VB_Name = "PEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Pevent.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PEvent.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Program library input screen code
Option Explicit
Option Compare Text
'Passed values
Dim smLibName As String
Dim lmLibCode As Long
Dim smVehName As String
Dim imVefCode As Integer
Dim smLibTypeEnabled As String
Dim smVersionChecked As String
Dim tmSpecCtrls(0 To 5)  As FIELDAREA
Dim imLBSpecCtrls As Integer
Dim tmEvtCtrls(0 To 11)  As FIELDAREA
Dim imLBEvtCtrls As Integer
Dim imSpecBoxNo As Integer      'Current Specification Box
Dim imEvtBoxNo As Integer       'Current event box
Dim imEvtRowNo As Integer       'Current row number in event area (start at 0)
Dim hmLtf As Integer            'Log title library file handle
Dim tmLtf As LTF                'LTF record image
Dim tmLtfSrchKey As INTKEY0     'LTF key record image
Dim imLtfRecLen As Integer      'LTF record length
Dim lmLtfRecPos As Long         'LTF record position
Dim hmLvf As Integer            'Log version library file handle
Dim tmLvf As LVF                'LVF record image
Dim tmLvfSrchKey As LONGKEY0     'LVF key record image
Dim tmLvf1SrchKey As LVFKEY1    'LVF key record image
Dim imLvfRecLen As Integer      'LVF record length
Dim lmLvfRecPos As Long         'LVF record position
Dim hmLef As Integer            'Log event file handle
Dim tmLef() As LEF              'Lef record images
Dim lmEvtRecPos() As Long       'Lef record position
Dim tmLefSrchKey As LEFKEY0     'Lef key record image
Dim imLefRecLen As Integer         'Lef record length
Dim hmEnf As Integer    'Event name file handle
Dim tmEnf As ENF        'ENF record image
Dim tmEnfSrchKey As INTKEY0    'ENF key record image
Dim imEnfRecLen As Integer        'ENF record length
Dim hmCef As Integer    'Comment file handle
Dim tmCef As CEF        'CEF record image
Dim tmCefSrchKey As LONGKEY0    'CEF key record image
Dim imCefRecLen As Integer        'CEF record length
Dim hmVef As Integer    'Vehicle file handle
Dim tmVef As VEF        'VEF record image
Dim tmVefSrchKey As INTKEY0    'VEF key record image
Dim imVefRecLen As Integer        'VEF record length
'Dim tmRec As LPOPREC
Dim smSpecSave(4) As String     'Values saved (1=Lib name; 2=Variation; 3=length; 4=Base time {blank=relative})
Dim imSpecSave(1) As Integer    'Value saved (1=Lib Type)
Dim smShow() As String  'Values shown in event area
Dim smSave() As String  'Values saved (1=Time{always stored as relative}; 2=Event type; 3=Event name; 4=Avail name or exlusion 1;
                        '5=Exclusion 2; 6=Units; 7=Length; 8=Comment; 9=Type of event; 10=lRecPos; 11=Event ID) in program area
Dim tmLibNameCode() As SORTCODE
Dim smLibNameCodeTag As String
Dim tmCLibNameCode() As SORTCODE
Dim smCLibNameCodeTag As String
Dim tmEvtTypeCode() As SORTCODE
Dim smEvtTypeCodeTag As String
Dim tmEvtAvailCode() As SORTCODE
Dim smEvtAvailCodeTag As String
Dim tmSelectCode() As SORTCODE
Dim smSelectCodeTag As String
Dim imSave() As Integer 'Values saved (1= True Time) in program area
Dim imUpdateAllowed As Integer
Dim imEvtNameIndex As Integer   'Which lbcEvtName index for a row (from event type)
Dim imLtfChg As Integer
Dim imLvfChg As Integer
Dim imLefChg As Integer  'True=Event value changed; False=No Event value changed
Dim imDateAssign As Integer    'True=Dates assigned
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imTypeChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imTypeSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imComboBoxIndex As Integer
Dim imFirstFocus As Integer
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imSettingValue As Integer   'True=Don't enable any box woth change
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imVpfIndex As Integer   'Vehicle option index
Dim imAllAnsw As Integer    'Used to indicate if all specification questions anwswered
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer  'Shift state when mouse down event occurrs
Dim imFirstTimeType As Integer  'True=first time at cbcType control- branch to other control
                            'Either Lib name(if new) or Time (if old)
Dim imBypassFocus As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim smEvtPrgName As String  'Program name- far coded to "Programs"
Dim smEvtAvName As String   'Contract Avail- far coded to"Contract Avails"
Dim imTimeRelative As Integer   'True=Time input is relative; False=Time input is from base length
Dim tmCEvtCtrls(0 To 11)  As FIELDAREA
Dim imLBCEvtCtrls As Integer
Dim imViewLibType As Integer    '0=Regular; 1=Special; 2=Sport; 3=Std Format
Dim fmPI As Single
Dim lmArcTimes() As Long    'Index 1 = Start Time; 2= End Time; 3= Radius; 4=smSave Index
Dim lmXCenter As Long
Dim lmYCenter As Long
Dim lmBaseRadius As Long
Dim imButton As Integer
Dim imIgnoreRightMove As Integer
Dim imLastArcPainted As Integer
Dim smHourCaption As String
Dim imShowHelpMsg As Integer    'True=Show help messages; False=Ignore help message system
Dim lm3600 As Long

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor
'8298
Dim bmTestEventID As Boolean
'10933
Dim bmEventIsCueZone As Boolean
Private Const ZONEEVENTMAX = 4

Const LBONE = 1
Const SPECLIBTYPEINDEX = 1  'Library type control/field
Const SPECLIBNAMEINDEX = 2  'Library name control/field
Const SPECVARINDEX = 3      'Variation control/field
Const SPECLENGTHINDEX = 4   'Length control/field
Const SPECBASETIMEINDEX = 5   'Base Time control/field
Const TIMEINDEX = 1     'Time control/field
Const EVTTYPEINDEX = 2  'Event type control/field
Const EVTNAMEINDEX = 3  'Event name control/field
Const AVAILINDEX = 4    'Avail or exclusion control/field
Const UNITSINDEX = 5    'Units date control/field
Const LENGTHINDEX = 6   'Length time control/field
Const TRUETIMEINDEX = 7 'True time control/field
Const EVTIDINDEX = 8  'Comment control/field
Const COMMENTINDEX = 9  'Comment control/field
Const EXCL1INDEX = 10    'Exclusion 1- exclusions painted within avail area
Const EXCL2INDEX = 11
Private Sub cbcSelect_Change()
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True    'Set change mode to avoid infinite loop
        Screen.MousePointer = vbHourglass  'Wait
        ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
        If ilRet = 0 Then
            ilIndex = cbcSelect.ListIndex
            If Not mReadRec(ilIndex, False) Then
                GoTo cbcSelectErr
            End If
            If Not mReadLefRec() Then
                GoTo cbcSelectErr
            End If
        Else
            If ilRet = 1 Then
                If cbcSelect.ListCount > 0 Then
                    cbcSelect.ListIndex = 0
                End If
            End If
            ilRet = 1   'Clear fields as no match name found
        End If
        pbcLibSpec.Cls
        pbcEvents.Cls
        pbcClock.Cls
        If ilRet = 0 Then
            imSelectedIndex = cbcSelect.ListIndex
            mMoveRecToCtrl
            mMoveEvtRecToCtrl False
            mInitEvtShow
        Else
            imSelectedIndex = 0
            mClearCtrlFields
            If slStr <> "[New]" Then
                smSpecSave(1) = slStr
                tmSpecCtrls(SPECLIBNAMEINDEX).iChg = True
            End If
        End If
        mInitSpecShow
        pbcLibSpec_Paint
        pbcEvents_Paint
        pbcClock_Paint
        Screen.MousePointer = vbDefault
        imChgMode = False
        ReDim tgPrg(0 To 0) As PRGDATE  'Time/Dates
        imDateAssign = False
    End If
    mSetCommands
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    Exit Sub
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
    mSetCommands
End Sub
Private Sub cbcSelect_DropDown()
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        pbcLibSpec_Paint
    End If
    If imTerminate Then
        Exit Sub
    End If

    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    slSvText = cbcSelect.Text
'    gSetIndexFromText cbcSelect
'    ilSvIndex = cbcSelect.ListIndex
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields
        pbcSpecSTab.SetFocus
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
Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcType_Change()
    If imTypeChgMode = False Then
        imTypeChgMode = True
        imChgMode = True
        If cbcType.Text <> "" Then
            gManLookAhead cbcType, imBSMode, imComboBoxIndex
        End If
        imTypeSelectedIndex = cbcType.ListIndex
        pbcLibSpec.Cls
        pbcEvents.Cls
        mClearCtrlFields
        cbcSelect.Clear 'Force population
        imTypeChgMode = False
    End If
End Sub
Private Sub cbcType_Click()
    imComboBoxIndex = cbcType.ListIndex
    imTypeSelectedIndex = imComboBoxIndex
    pbcLibSpec.Cls
    pbcEvents.Cls
    mClearCtrlFields
    cbcSelect.Clear 'Force population
    mSetCommands
End Sub
Private Sub cbcType_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        pbcLibSpec_Paint
    End If
    If imFirstTimeType Then
        imFirstTimeType = False
        pbcEvents.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        If imSelectedIndex > 0 Then
            pbcSTab.SetFocus
        Else
            pbcSpecSTab.SetFocus
        End If
        Exit Sub
    End If
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    gCtrlGotFocus cbcType
    imComboBoxIndex = imTypeSelectedIndex
End Sub
Private Sub cbcType_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcType_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcType.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub ckcShowVersion_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
    End If
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDates_Click()
    Dim ilLoop As Integer
    Dim ilLibType As Integer
    'If Not gWinRoom(igNoExeWinRes(PRGDATESEXE)) Then
    '    Exit Sub
    'End If
    'MousePointer = vbHourGlass  'Wait
    cmcCancel.Enabled = False
    cmcDone.Enabled = False
    ilLibType = igLibType
    igLibType = imSpecSave(1)
    lgLibLength = 0 'Test past midnight within this code- as library length might change
    PrgDates.Show vbModal
    igLibType = ilLibType
    'MousePointer = vbDefault    'Default
    cmcCancel.Enabled = True
    cmcDone.Enabled = True
    imDateAssign = False
    For ilLoop = LBound(tgPrg) To UBound(tgPrg) - 1 Step 1
        If (tgPrg(ilLoop).sStartTime <> "") And (tgPrg(ilLoop).sStartDate <> "") Then
            imDateAssign = True
        End If
    Next ilLoop
    If pbcSTab.Enabled Then
        pbcSTab.SetFocus
    Else
        pbcSpecSTab.SetFocus
    End If
End Sub
Private Sub cmcDates_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
        If imSpecBoxNo > 0 Then
            mSpecEnableBox imSpecBoxNo
        ElseIf imEvtBoxNo > 0 Then
            mEvtEnableBox imEvtBoxNo
        End If
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim slStamp As String   'Date/Time stamp for file
    Dim llCode As Long
    Dim slMsg As String
    Dim ilLef As Integer
    Dim llRecPos As Long
    Dim tlLtf As LTF
    Dim tlLvf As LVF
    Dim tlLef As LEF
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imSelectedIndex > 0 Then
        Screen.MousePointer = vbHourglass
        llCode = tmLvf.lCode
        ilRet = gLibCodeRefExist(PEvent, "B", llCode)
        'Check that record is not referenced-Code missing
'        llCode = tmLnf.lCode
'        ilOffset = gFieldOffset("LCF", "LCFLNF1")
'        For ilIndex = LBound(tlLcf.lLnfCode) To UBound(tlLcf.lLnfCode) Step 1
            'Check that record is not referenced-Code missing
'            ilRet = gLLCodeRefExistOffset(PEvent, llCode, "Lcf.Btr", ilOffset)
            If ilRet Then
                Screen.MousePointer = vbDefault    'Default
                slMsg = "Cannot erase - a Log Calendar Day references this Name"
                ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
'            ilOffset = ilOffset + 4
'        Next ilIndex
        Screen.MousePointer = vbDefault    'Default
        ilRet = MsgBox("OK to remove " & tmLtf.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        slStamp = gFileDateTime(sgDBPath & "Lvf.Btr")
        ilRet = btrBeginTrans(hmLtf, 1000)
        For ilLef = LBound(tmLef) To UBound(tmLef) - 1 Step 1
            llRecPos = CLng(smSave(10, ilLef + 1))
            ilRet = btrGetDirect(hmLef, tlLef, imLefRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmLtf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
            'tmRec = tlLef
            'ilRet = gGetByKeyForUpdate("LEF", hmLef, tmRec)
            'tlLef = tmRec
            'If ilRet <> BTRV_ERR_NONE Then
            '    ilRet = btrAbortTrans(hmLtf)
            '    Screen.MousePointer = vbDefault    'Default
            '    ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
            '    Exit Sub
            'End If
            Do
                If mReadCefRec(tlLef.lEvtIDCefCode, SETFORWRITE) Then
                    If tlLef.lEvtIDCefCode <> 0 Then
                        ilRet = btrDelete(hmCef)
                    End If
                Else
                    ilRet = btrAbortTrans(hmLtf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            Do
                If mReadCefRec(tlLef.lCefCode, SETFORWRITE) Then
                    If tlLef.lCefCode <> 0 Then
                        ilRet = btrDelete(hmCef)
                    End If
                Else
                    ilRet = btrAbortTrans(hmLtf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmLtf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
            Do
                ilRet = btrDelete(hmLef)
                If ilRet = BTRV_ERR_CONFLICT Then
                    ilCRet = btrGetDirect(hmLef, tlLef, imLefRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilCRet <> BTRV_ERR_NONE Then
                        ilRet = btrAbortTrans(hmLtf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                        Exit Sub
                    End If
                    'tmRec = tlLef
                    'ilCRet = gGetByKeyForUpdate("LEF", hmLef, tmRec)
                    'tlLef = tmRec
                    'If ilCRet <> BTRV_ERR_NONE Then
                    '    ilRet = btrAbortTrans(hmLtf)
                    '    Screen.MousePointer = vbDefault    'Default
                    '    ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
                    '    Exit Sub
                    'End If
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmLtf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
        Next ilLef
        ilRet = btrGetDirect(hmLvf, tlLvf, imLvfRecLen, lmLvfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrAbortTrans(hmLtf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        'tmRec = tlLvf
        'ilRet = gGetByKeyForUpdate("LVF", hmLvf, tmRec)
        'tlLvf = tmRec
        'If ilRet <> BTRV_ERR_NONE Then
        '    ilRet = btrAbortTrans(hmLtf)
        '    Screen.MousePointer = vbDefault    'Default
        '    ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
        '    Exit Sub
        'End If
        Do
            ilRet = btrDelete(hmLvf)
            If ilRet = BTRV_ERR_CONFLICT Then
                ilCRet = btrGetDirect(hmLvf, tlLvf, imLvfRecLen, lmLvfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilCRet <> BTRV_ERR_NONE Then
                    ilRet = btrAbortTrans(hmLtf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                'tmRec = tlLvf
                'ilCRet = gGetByKeyForUpdate("LVF", hmLvf, tmRec)
                'tlLvf = tmRec
                'If ilCRet <> BTRV_ERR_NONE Then
                '    ilRet = btrAbortTrans(hmLtf)
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
                '    Exit Sub
                'End If
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilCRet <> BTRV_ERR_NONE Then
            ilRet = btrAbortTrans(hmLtf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = btrGetDirect(hmLtf, tlLtf, imLtfRecLen, lmLtfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrAbortTrans(hmLtf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        'Only remove the title if all versions are removed
        tmLvf1SrchKey.iLtfCode = tlLtf.iCode
        tmLvf1SrchKey.iVersion = 32000
        ilRet = btrGetGreaterOrEqual(hmLvf, tlLvf, imLvfRecLen, tmLvf1SrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If (ilRet <> BTRV_ERR_NONE) Or (tlLvf.iLtfCode <> tlLtf.iCode) Then
            Do
                'tmRec = tlLtf
                'ilCRet = gGetByKeyForUpdate("LTF", hmLtf, tmRec)
                'tlLtf = tmRec
                'If ilCRet <> BTRV_ERR_NONE Then
                '    ilRet = btrAbortTrans(hmLtf)
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
                '    Exit Sub
                'End If
                ilRet = btrDelete(hmLtf)
                If ilRet = BTRV_ERR_CONFLICT Then
                    ilCRet = btrGetDirect(hmLtf, tlLtf, imLtfRecLen, lmLtfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmLtf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
        End If
        ilRet = btrEndTrans(hmLtf)
        'If lbcSelectCode.Tag <> "" Then
        '    If slStamp = lbcSelectCode.Tag Then
        '        lbcSelectCode.Tag = FileDateTime(sgDBPath & "Lvf.Btr")
        '    End If
        'End If
        If smSelectCodeTag <> "" Then
            If slStamp = smSelectCodeTag Then
                smSelectCodeTag = gFileDateTime(sgDBPath & "Lvf.Btr")
            End If
        End If
        'lbcSelectCode.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tmSelectCode()
        cbcSelect.RemoveItem imSelectedIndex
        Screen.MousePointer = vbDefault    'Default
    End If
    'Remove focus from control and make invisible
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcLibSpec.Cls
    pbcEvents.Cls
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Resume Next
End Sub
Private Sub cmcErase_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcErase_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcEvtDropDown_Click()
    Select Case imEvtBoxNo
        Case TIMEINDEX
            If imTimeRelative Then
                plclen.Visible = Not plclen.Visible
            Else
                plcTme.Visible = Not plcTme.Visible
            End If
        Case EVTTYPEINDEX
            lbcEvtType.Visible = Not lbcEvtType.Visible
        Case EVTNAMEINDEX
            lbcEvtName(imEvtNameIndex).Visible = Not lbcEvtName(imEvtNameIndex).Visible
        Case AVAILINDEX
            lbcEvtAvail.Visible = Not lbcEvtAvail.Visible
        Case EXCL1INDEX
            lbcExcl(0).Visible = Not lbcExcl(0).Visible
        Case EXCL2INDEX
            lbcExcl(1).Visible = Not lbcExcl(1).Visible
        Case LENGTHINDEX
            plclen.Visible = Not plclen.Visible
    End Select
    edcEvtDropDown.SelStart = 0
    edcEvtDropDown.SelLength = Len(edcEvtDropDown.Text)
    edcEvtDropDown.SetFocus
End Sub
Private Sub cmcEvtDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcSpecDropDown_Click()
    Select Case imSpecBoxNo
        Case SPECLIBNAMEINDEX
            lbcLibName.Visible = Not lbcLibName.Visible
        Case SPECLENGTHINDEX
            plclen.Visible = Not plclen.Visible
        Case SPECBASETIMEINDEX
            plcTme.Visible = Not plcTme.Visible
    End Select
    edcSpecDropDown.SelStart = 0
    edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    edcSpecDropDown.SetFocus
End Sub
Private Sub cmcSpecDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUndo_Click()
    Dim ilIndex As Integer
    pbcLibSpec.Cls
    pbcEvents.Cls
    imLtfChg = False
    imLvfChg = False
    imLefChg = False
    imDateAssign = False
    ReDim tgPrg(0 To 0) As PRGDATE  'Time/Dates
    imDateAssign = False
    ilIndex = imSelectedIndex
    If ilIndex > 0 Then
        Screen.MousePointer = vbHourglass  'Wait
        If Not mReadRec(ilIndex, igPrgDupl) Then
            GoTo cmcUndoErr
        End If
        If Not mReadLefRec() Then
            GoTo cmcUndoErr
        End If
        pbcLibSpec.Cls
        pbcEvents.Cls
        mMoveRecToCtrl
        mMoveEvtRecToCtrl False
        mInitSpecShow
        mInitEvtShow
        pbcLibSpec_Paint
        pbcEvents_Paint
        Screen.MousePointer = vbDefault
        mSetCommands
        imSpecBoxNo = -1
        pbcSpecSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    mClearCtrlFields 'If coming from [New], select_change will not be generated
    If cbcSelect.ListCount > 0 Then
        cbcSelect.ListIndex = 0
        mSetCommands
        cbcSelect.SetFocus
    Else
        mSetCommands
        cbcType.SetFocus
    End If
    Exit Sub
cmcUndoErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub cmcUndo_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcUpdate_Click()
    Dim imSvSelectedIndex As Integer
    Dim llCode As Long
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    imSvSelectedIndex = imSelectedIndex
'    If imSelectedIndex > 0 Then
'        slName = cbcSelect.Text
'    Else
'        slName = ""
'    End If
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        If imSpecBoxNo > 0 Then
            mSpecEnableBox imSpecBoxNo
        ElseIf imEvtBoxNo > 0 Then
            mEvtEnableBox imEvtBoxNo
        End If
        Exit Sub
    End If
'    'Must reset display so altered flag is cleared and setcommand will turn select on
'    If imSvSelectedIndex <> 0 Then
''        cbcSelect.Text = slName
'        cbcSelect.ListIndex = 1 'latest version added as item 1
'    Else
'        cbcSelect.ListIndex = 0
'    End If
'    cbcSelect_Change    'Call change so picture area repainted
    llCode = tmLvf.lCode
    cbcSelect.Clear
    smSelectCodeTag = ""
    mPopulate
    If imSvSelectedIndex <> 0 Then
        For ilLoop = 0 To UBound(tmSelectCode) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
            slNameCode = tmSelectCode(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = llCode Then
                If cbcSelect.ListIndex = ilLoop + 1 Then
                    cbcSelect_Change
                Else
                    cbcSelect.ListIndex = ilLoop + 1
                End If
                Exit For
            End If
        Next ilLoop
    Else
        cbcSelect.ListIndex = 0
    End If
    imLtfChg = False
    imLvfChg = False
    imLefChg = False
    mSetCommands
    cbcSelect.SetFocus
End Sub
Private Sub cmcUpdate_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUpdate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub edcComment_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcComment_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    'If Not gCheckKeyAscii(ilKey) Then
    
    '11-16-10 need unique check of special characters due to Audio Vault with ";" in comments field
    If Not mCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcEvtDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imEvtBoxNo
        Case EVTTYPEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcEvtDropDown, lbcEvtType, imBSMode, slStr)
            If ilRet = 1 Then
                lbcEvtType.ListIndex = 0
            End If
        Case EVTNAMEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcEvtDropDown, lbcEvtName(imEvtNameIndex), imBSMode, slStr)
            If ilRet = 1 Then
                lbcEvtName(imEvtNameIndex).ListIndex = 0
            End If
        Case AVAILINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcEvtDropDown, lbcEvtAvail, imBSMode, slStr)
            If ilRet = 1 Then
                lbcEvtAvail.ListIndex = 0
            End If
        Case EXCL1INDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcEvtDropDown, lbcExcl(0), imBSMode, slStr)
            If ilRet = 1 Then
                lbcExcl(0).ListIndex = 0
            End If
        Case EXCL2INDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcEvtDropDown, lbcExcl(1), imBSMode, slStr)
            If ilRet = 1 Then
                lbcExcl(1).ListIndex = 0
            End If
    End Select
End Sub
Private Sub edcEvtDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcEvtDropDown_GotFocus()
    Select Case imEvtBoxNo
        Case EVTTYPEINDEX
            If lbcEvtType.ListCount = 1 Then
                lbcEvtType.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case EVTNAMEINDEX
'            If lbcEvtName(imEvtNameIndex).ListCount = 1 Then
'                lbcEvtName(imEvtNameIndex).ListIndex = 0
'                If imTabDirection = -1 Then  'Right To Left
'                    pbcSTab.SetFocus
'                Else
'                    pbcTab.SetFocus
'                End If
'                Exit Sub
'            End If
        Case AVAILINDEX
            If lbcEvtAvail.ListCount = 1 Then
                lbcEvtAvail.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case EXCL1INDEX
        Case EXCL2INDEX
    End Select
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcEvtDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcEvtDropDown_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilKey As Integer
    If imEvtBoxNo <> EVTIDINDEX Then
        If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
            If edcEvtDropDown.SelLength <> 0 Then    'avoid deleting two characters
                imBSMode = True 'Force deletion of character prior to selected text
            End If
        End If
        ilKey = KeyAscii
        If Not gCheckKeyAscii(ilKey) Then
            KeyAscii = 0
            Exit Sub
        End If
        If ((imEvtBoxNo = TIMEINDEX) And (imTimeRelative)) Or (imEvtBoxNo = LENGTHINDEX) Then
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                ilFound = False
                For ilLoop = LBound(igLegalLength) To UBound(igLegalLength) Step 1
                    If KeyAscii = igLegalLength(ilLoop) Then
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
            gLengthOutLine KeyAscii, imcLenOutline
        End If
        If (imEvtBoxNo = TIMEINDEX) And (Not imTimeRelative) Then
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
        End If
    Else
        ilKey = KeyAscii
        If Not gCheckKeyAscii(ilKey) Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Private Sub edcEvtDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If imEvtBoxNo <> EVTIDINDEX Then
        If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KeyDown) Then
            imDirProcess = KeyCode 'mDirection 0
            pbcTab.SetFocus
            Exit Sub
        End If
        If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
            Select Case imEvtBoxNo
                Case TIMEINDEX
                    If (Shift And vbAltMask) > 0 Then
                        If imTimeRelative Then
                            plclen.Visible = Not plclen.Visible
                        Else
                            plcTme.Visible = Not plcTme.Visible
                        End If
                    End If
                Case EVTTYPEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcEvtType, imLbcArrowSetting
                Case EVTNAMEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcEvtName(imEvtNameIndex), imLbcArrowSetting
                Case AVAILINDEX
                    gProcessArrowKey Shift, KeyCode, lbcEvtAvail, imLbcArrowSetting
                Case EXCL1INDEX
                    gProcessArrowKey Shift, KeyCode, lbcExcl(0), imLbcArrowSetting
                Case EXCL2INDEX
                    gProcessArrowKey Shift, KeyCode, lbcExcl(1), imLbcArrowSetting
                Case LENGTHINDEX
                    If (Shift And vbAltMask) > 0 Then
                        plclen.Visible = Not plclen.Visible
                    End If
            End Select
            edcEvtDropDown.SelStart = 0
            edcEvtDropDown.SelLength = Len(edcEvtDropDown.Text)
        End If
    End If
End Sub
Private Sub edcEvtDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imEvtBoxNo
            Case EVTTYPEINDEX, EVTNAMEINDEX, AVAILINDEX, EXCL1INDEX, EXCL2INDEX
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
Private Sub edcEvtEdit_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcEvtEdit_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case imEvtBoxNo
        Case UNITSINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcEvtEdit.Text
            slStr = Left$(slStr, edcEvtEdit.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcEvtEdit.SelStart - edcEvtEdit.SelLength)
            If gCompNumberStr(slStr, "30") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub
Private Sub edcEvtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KeyDown) Then
        imDirProcess = KeyCode 'mDirection 0
        pbcTab.SetFocus
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcSpecDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imSpecBoxNo
        Case SPECLIBNAMEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcSpecDropDown, lbcLibName, imBSMode, slStr)
'            If ilRet = 1 Then   'input was ""
'                 lbcLibName.ListIndex = 0
'            End If
        Case SPECLENGTHINDEX
        Case SPECBASETIMEINDEX
    End Select
End Sub
Private Sub edcSpecDropDown_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcSpecDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcSpecDropDown_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilKey As Integer
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
    If imSpecBoxNo = SPECLIBNAMEINDEX Then
        If (KeyAscii = KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    If imSpecBoxNo = SPECLENGTHINDEX Then
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            ilFound = False
            For ilLoop = LBound(igLegalLength) To UBound(igLegalLength) Step 1
                If KeyAscii = igLegalLength(ilLoop) Then
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
        gLengthOutLine KeyAscii, imcLenOutline
    End If
    If imSpecBoxNo = SPECBASETIMEINDEX Then
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
            Case SPECLENGTHINDEX
                If (Shift And vbAltMask) > 0 Then
                    plclen.Visible = Not plclen.Visible
                End If
            Case SPECBASETIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
        End Select
        edcSpecDropDown.SelStart = 0
        edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
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
    If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcLibSpec.Enabled = False
        pbcSpecSTab.Enabled = False
        pbcSpecTab.Enabled = False
        pbcEvents.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcLibSpec.Enabled = True
        pbcSpecSTab.Enabled = True
        pbcSpecTab.Enabled = True
        pbcEvents.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    PEvent.Refresh
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
        If (imSpecBoxNo > 0) Or (imEvtBoxNo > 0) Then
            plcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imSpecBoxNo > 0 Then
            mSpecEnableBox imSpecBoxNo
        ElseIf imEvtBoxNo > 0 Then
            mEvtEnableBox imEvtBoxNo
        End If
        If ilReSet Then
            plcSelect.Enabled = True
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
        mSpecSetShow imSpecBoxNo
        imSpecBoxNo = -1
        mEvtSetShow imEvtBoxNo
        imEvtBoxNo = -1
        imEvtRowNo = -1
        pbcArrow.Visible = False
        lacEvtFrame.Visible = False
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            If imSpecBoxNo <> -1 Then
                mSpecEnableBox imSpecBoxNo
            ElseIf imEvtBoxNo <> -1 Then
                mEvtEnableBox imEvtBoxNo
            End If
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    ilRet = btrClose(hmEnf)
    btrDestroy hmEnf
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    ilRet = btrClose(hmLef)
    btrDestroy hmLef
    ilRet = btrClose(hmLvf)
    btrDestroy hmLvf
    ilRet = btrClose(hmLtf)
    btrDestroy hmLtf
    Erase tmLef
    Erase lmEvtRecPos
    Erase smShow
    Erase smSave
    Erase imSave
    Erase tmLibNameCode
    Erase tmCLibNameCode
    Erase tmEvtTypeCode
    Erase tmEvtAvailCode
    Erase tmSelectCode
    
    Set PEvent = Nothing   'Remove data segment
    
End Sub
Private Sub hbcHour_Change()
    If hbcHour.Value = 1 Then
        smHourCaption = "1st Hour"
    ElseIf hbcHour.Value = 2 Then
        smHourCaption = "2nd Hour"
    ElseIf hbcHour.Value = 3 Then
        smHourCaption = "3rd Hour"
    ElseIf hbcHour.Value = 21 Then
        smHourCaption = "21st Hour"
    ElseIf hbcHour.Value = 22 Then
        smHourCaption = "22nd Hour"
    ElseIf hbcHour.Value = 23 Then
        smHourCaption = "23rd Hour"
    Else
        smHourCaption = Trim$(str$(hbcHour.Value)) & "th Hour"
    End If
    plcHour.Cls
    plcHour_Paint
    pbcClock_Paint
End Sub
Private Sub hbcHour_Scroll()
    If hbcHour.Value = 1 Then
        smHourCaption = "1st Hour"
    ElseIf hbcHour.Value = 2 Then
        smHourCaption = "2nd Hour"
    ElseIf hbcHour.Value = 3 Then
        smHourCaption = "3rd Hour"
    ElseIf hbcHour.Value = 21 Then
        smHourCaption = "21st Hour"
    ElseIf hbcHour.Value = 22 Then
        smHourCaption = "22nd Hour"
    ElseIf hbcHour.Value = 23 Then
        smHourCaption = "23rd Hour"
    Else
        smHourCaption = Trim$(str$(hbcHour.Value)) & "th Hour"
    End If
    plcHour.Cls
    plcHour_Paint
End Sub
Private Sub imcCalc_Click()
    Dim slStr As String        'General string
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(CALCEXE)) Then
    '    Exit Sub
    'End If
    'Screen.MousePointer = vbHourglass  'Wait
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "PEvent^Test\" & sgUserName & "\" & Trim$(str$(CInt(fgCalcLeft))) & "\" & Trim$(str$(CInt(fgCalcTop)))
        Else
            slStr = "PEvent^Prod\" & sgUserName & "\" & Trim$(str$(CInt(fgCalcLeft))) & "\" & Trim$(str$(CInt(fgCalcTop)))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "PEvent^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(CInt(fgCalcLeft))) & "\" & Trim$(Str$(CInt(fgCalcTop)))
    '    Else
    '        slStr = "PEvent^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(CInt(fgCalcLeft))) & "\" & Trim$(Str$(CInt(fgCalcTop)))
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MathCalc.Exe " & slStr, 1)
    'PEvent.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MathCalc.Show vbModal
    ilParse = gParseItem(sgDoneMsg, 2, "|", slStr)
    fgCalcLeft = Val(slStr)
    ilParse = gParseItem(sgDoneMsg, 3, "|", slStr)
    fgCalcTop = Val(slStr)
    'PEvent.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
End Sub
Private Sub imcCalc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
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
    If (imEvtRowNo < vbcEvents.Value + 1) Or (imEvtRowNo >= vbcEvents.Value + vbcEvents.LargeChange + 2) Then
        Exit Sub
    End If
    ilRowNo = imEvtRowNo
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    gCtrlGotFocus ActiveControl
    ilUpperBound = UBound(smSave, 2)
    For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
        For ilIndex = 1 To UBound(smSave, 1) Step 1
            smSave(ilIndex, ilLoop) = smSave(ilIndex, ilLoop + 1)
        Next ilIndex
        For ilIndex = 1 To UBound(imSave, 1) Step 1
            imSave(ilIndex, ilLoop) = imSave(ilIndex, ilLoop + 1)
        Next ilIndex
        For ilIndex = 1 To UBound(smShow, 1) Step 1
            smShow(ilIndex, ilLoop) = smShow(ilIndex, ilLoop + 1)
        Next ilIndex
    Next ilLoop
    ilUpperBound = UBound(smSave, 2)
    ReDim Preserve smSave(0 To 11, 0 To ilUpperBound - 1) As String
    ReDim Preserve imSave(0 To 1, 0 To ilUpperBound - 1) As Integer
    ReDim Preserve smShow(0 To COMMENTINDEX, 0 To ilUpperBound - 1) As String
    If UBound(smSave, 2) < vbcEvents.LargeChange Then 'was <=
        vbcEvents.Max = LBONE - 1   'LBound(smSave, 2) - 1
    Else
        vbcEvents.Max = UBound(smSave, 2) - vbcEvents.LargeChange - 1
    End If
    imLefChg = True
    mSetCommands
    lacEvtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcEvents.Cls
    pbcEvents_Paint
End Sub
Private Sub imcTrash_DragDrop(Source As control, X As Single, Y As Single)
'    lacEvtFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        lacEvtFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
'        lacEvtFrame.DragIcon = IconTraf!imcIconTrash.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        lacEvtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub lbcEvtAvail_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        pbcEatTab(1).Enabled = True
        pbcEatTab(0).Enabled = True
        pbcEatTab(0).SetFocus
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcEvtAvail, edcEvtDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcEvtAvail_DblClick()
    tmcClick.Enabled = False
    pbcEatTab(1).Enabled = False
    pbcEatTab(0).Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcEvtAvail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcEvtAvail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcEvtAvail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcEvtAvail, edcEvtDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcEvtName_Click(Index As Integer)
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        pbcEatTab(1).Enabled = True
        pbcEatTab(0).Enabled = True
        pbcEatTab(0).SetFocus
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcEvtName(Index), edcEvtDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcEvtName_DblClick(Index As Integer)
    tmcClick.Enabled = False
    pbcEatTab(1).Enabled = False
    pbcEatTab(0).Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcEvtName_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcEvtName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcEvtName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcEvtName(Index), edcEvtDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcEvtType_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        pbcEatTab(1).Enabled = True
        pbcEatTab(0).Enabled = True
        pbcEatTab(0).SetFocus
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcEvtType, edcEvtDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcEvtType_DblClick()
    tmcClick.Enabled = False
    pbcEatTab(1).Enabled = False
    pbcEatTab(0).Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcEvtType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcEvtType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcEvtType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcEvtType, edcEvtDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcExcl_Click(Index As Integer)
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        pbcEatTab(1).Enabled = True
        pbcEatTab(0).Enabled = True
        pbcEatTab(0).SetFocus
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcExcl(Index), edcEvtDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcExcl_DblClick(Index As Integer)
    tmcClick.Enabled = False
    pbcEatTab(1).Enabled = False
    pbcEatTab(0).Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcExcl_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcExcl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcExcl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcExcl(Index), edcEvtDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcLibName_Click()
    gProcessLbcClick lbcLibName, edcSpecDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcLibName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
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
'       ilOnlyAddr (I)- Clear only fields after address
'
    Dim ilLoop As Integer
    imSpecSave(1) = -1  'Library type
    smSpecSave(1) = ""  'Library name
    smSpecSave(2) = ""  'Variation
    smSpecSave(3) = ""  'Length
    smSpecSave(4) = ""  'Base Time
    tmLtf.iVefCode = 0  'Force this field to be reset in mMoveCtrlToRec
    For ilLoop = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        tmSpecCtrls(ilLoop).sShow = ""
    Next ilLoop
    mMoveCtrlToRec False
    For ilLoop = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        tmSpecCtrls(ilLoop).iChg = False
    Next ilLoop
    ReDim tmLef(0 To 0) As LEF              'Lef record images
    ReDim lmEvtRecPos(0 To 0) As Long
    ReDim smSave(0 To 11, 0 To 1) As String
    ReDim imSave(0 To 1, 0 To 1) As Integer
    ReDim smShow(0 To COMMENTINDEX, 0 To 1) As String
    lmLtfRecPos = -1    'Indicator of [New] contract
    lmLvfRecPos = -1    'Indicator of [New] contract
    imLtfChg = False
    imLvfChg = False
    imLefChg = False
    imAllAnsw = False
    vbcEvents.Value = vbcEvents.Min
    vbcEvents.Max = vbcEvents.Min
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCLibPop                        *
'*                                                     *
'*             Created:10/11/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection library *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mCLibPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slType As String
    Screen.MousePointer = vbHourglass  'Wait
    If imViewLibType = 3 Then 'Std Format
        slType = "F"
    ElseIf imViewLibType = 2 Then 'Sports
        slType = "P"
    ElseIf imViewLibType = 1 Then 'Special
        slType = "S"
    Else    'Regular
        slType = "R"
    End If
    'ilRet = gPopProgLibBox(PEvent, LATESTLIB, slType, imVefCode, lbcCLibName, lbcCLibNameCode)
    ilRet = gPopProgLibBox(PEvent, LATESTLIB, slType, imVefCode, lbcCLibName, tmCLibNameCode(), smCLibNameCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mCLibPopErr
        gCPErrorMsg ilRet, "mCLibPope (gPopProgLibBox)", PEvent
        On Error GoTo 0
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mCLibPopErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCompVector                     *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Given X,Y compute Angle (time) *
'*                      and radius                     *
'*                                                     *
'*******************************************************
Private Sub mCompVector(llX As Long, llY As Long, llTime As Long, llRadius As Long)
    Dim llXc As Long
    Dim llYc As Long
    Dim flDelta As Single
    llXc = llX - lmXCenter
    llYc = lmYCenter - llY
    flDelta = CSng(llXc) / CSng(llYc)
    If llXc = 0 Then
        If llYc >= 0 Then
            llTime = 0
        Else
            llTime = 1800
        End If
    ElseIf llYc = 0 Then
        If llYc >= 0 Then
            llTime = 900
        Else
            llTime = 2700
        End If
    ElseIf (llXc > 0) And (llYc > 0) Then
        llTime = 1800 * Atn(flDelta) / fmPI
    ElseIf (llXc > 0) And (llYc < 0) Then
        llTime = 1800 + 1800 * Atn(flDelta) / fmPI
    ElseIf (llXc < 0) And (llYc < 0) Then
        llTime = 1800 + 1800 * Atn(flDelta) / fmPI
    ElseIf (llXc < 0) And (llYc > 0) Then
        llTime = lm3600 + 1800 * Atn(flDelta) / fmPI
    End If
    llTime = llTime + lm3600 * (hbcHour.Value - hbcHour.Min)
    llRadius = Sqr(llXc * llXc + llYc * llYc)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: mDeleteEvtNameCtrl             *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Reduce control array to 1       *
'*                     (index =0)                      *
'*                                                     *
'*******************************************************
Private Sub mDeleteEvtNameCtrl()
    Dim ilMaxNoCtrls As Integer
    Dim ilLoop As Integer
    ilMaxNoCtrls = 0
    On Error GoTo gDetMaxCtrlCntErr
    Do While (err = 0)
        ilMaxNoCtrls = lbcEvtName(ilMaxNoCtrls).Index
        ilMaxNoCtrls = ilMaxNoCtrls + 1
    Loop
gDetMaxCtrlCntErr:
    On Error GoTo 0
    ilMaxNoCtrls = ilMaxNoCtrls - 1
    For ilLoop = ilMaxNoCtrls To 1 Step -1
        Unload lbcEvtName(ilLoop)
        Unload lbcEvtNameCode(ilLoop)
    Next ilLoop
    Exit Sub    'Require to avoid error as no resume was executed
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtAvailBranch                 *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Avail  *
'*                      names and process communication*
'*                      back from avail names          *
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
Private Function mEvtAvailBranch()
'
'   ilRet = mEvtAvailBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    'Test if [New] Or new name specified
    ilRet = gOptionalLookAhead(edcEvtDropDown, lbcEvtAvail, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcEvtDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mEvtAvailBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(AVAILNAMESLIST)) Then
    '    imDoubleClickName = False
    '    mEvtAvailBranch = True
    '    mEvtEnableBox imEvtBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourglass  'Wait
    igANmCallSource = CALLSOURCEPEVENT
    If edcEvtDropDown.Text = "[New]" Then
        sgANmName = ""
    Else
        sgANmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "PEvent^Test\" & sgUserName & "\" & Trim$(str$(igANmCallSource)) & "\" & sgANmName
        Else
            slStr = "PEvent^Prod\" & sgUserName & "\" & Trim$(str$(igANmCallSource)) & "\" & sgANmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "PEvent^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
    '    Else
    '        slStr = "PEvent^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "AName.Exe " & slStr, 1)
    'PEvent.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    AName.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgANmName)
    igANmCallSource = Val(sgANmName)
    ilParse = gParseItem(slStr, 2, "\", sgANmName)
    'PEvent.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mEvtAvailBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igANmCallSource = CALLDONE Then  'Done
        igANmCallSource = CALLNONE
        lbcEvtAvail.Clear
        smEvtAvailCodeTag = ""
        sgAvailAnfStamp = ""
        mEvtAvailPop
        If imTerminate Then
            mEvtAvailBranch = False
            Exit Function
        End If
        gFindMatch sgANmName, 1, lbcEvtAvail
        If gLastFound(lbcEvtAvail) > 0 Then
            imChgMode = True
            lbcEvtAvail.ListIndex = gLastFound(lbcEvtAvail)
            edcEvtDropDown.Text = lbcEvtAvail.List(lbcEvtAvail.ListIndex)
            imChgMode = False
            mEvtAvailBranch = False
        Else
            imChgMode = True
            lbcEvtAvail.ListIndex = 1
            edcEvtDropDown.Text = lbcEvtAvail.List(1)
            imChgMode = False
            edcEvtDropDown.SetFocus
            sgANmName = ""
            Exit Function
        End If
        sgANmName = ""
    End If
    If igANmCallSource = CALLCANCELLED Then  'Cancelled
        igANmCallSource = CALLNONE
        sgANmName = ""
        mEvtEnableBox imEvtBoxNo
        Exit Function
    End If
    If igANmCallSource = CALLTERMINATED Then
        igANmCallSource = CALLNONE
        sgANmName = ""
        mEvtEnableBox imEvtBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtAvailPop                    *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Avail Pop the selection Avail  *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mEvtAvailPop()
'
'   mEvtAvailPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilLp As Integer
    Dim slStr As String
    ilIndex = lbcEvtAvail.ListIndex
    If ilIndex > 0 Then
        slName = lbcEvtAvail.List(ilIndex)
    End If
    ilFilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffSet(0) = 0
    'ilRet = gIMoveListBox(PEvent, lbcEvtAvail, lbcEvtAvailCode, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(PEvent, lbcEvtAvail, tmEvtAvailCode(), smEvtAvailCodeTag, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        'Remove "Post Log" avail name
        For ilLoop = 0 To lbcEvtAvail.ListCount - 1 Step 1
            slStr = Trim$(lbcEvtAvail.List(ilLoop))
            If StrComp(slStr, "Post Log", 1) = 0 Then
                lbcEvtAvail.RemoveItem ilLoop
                For ilLp = ilLoop To UBound(tmEvtAvailCode) - 1 Step 1
                    tmEvtAvailCode(ilLp) = tmEvtAvailCode(ilLp + 1)
                Next ilLp
                ReDim Preserve tmEvtAvailCode(LBound(tmEvtAvailCode) To UBound(tmEvtAvailCode) - 1) As SORTCODE
                Exit For
            End If
        Next ilLoop
        On Error GoTo mEvtAvailPopErr
        gCPErrorMsg ilRet, "mEvtAvailPop (gIMoveListBox)", PEvent
        On Error GoTo 0
        lbcEvtAvail.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcEvtAvail
            If gLastFound(lbcEvtAvail) > 0 Then
                lbcEvtAvail.ListIndex = gLastFound(lbcEvtAvail)
            Else
                lbcEvtAvail.ListIndex = -1
            End If
        Else
            lbcEvtAvail.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mEvtAvailPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtDirection                    *
'*                                                     *
'*             Created:9/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move to box indicated by       *
'*                      user direction                 *
'*                                                     *
'*******************************************************
Private Sub mEvtDirection(ilMoveDir As Integer)
'
'   mEvtDirection ilMove
'   Where:
'       ilMove (I)- 0=Up; 1= down; 2= left; 3= right
'
    mEvtSetShow imEvtBoxNo
    Select Case ilMoveDir
        Case KEYUP  'Up
            If imEvtRowNo > 1 Then
                imEvtRowNo = imEvtRowNo - 1
                If imEvtRowNo < vbcEvents.Value + 1 Then
                    imSettingValue = True
                    vbcEvents.Value = vbcEvents.Value - 1
                End If
            Else
                imEvtRowNo = UBound(smSave, 2)
                imSettingValue = True
                If imEvtRowNo <= vbcEvents.LargeChange + 1 Then
                    vbcEvents.Value = 0
                Else
                    vbcEvents.Value = imEvtRowNo - vbcEvents.LargeChange - 1
                End If
            End If
        Case KeyDown  'Down
            If imEvtRowNo < UBound(smSave, 2) Then
                imEvtRowNo = imEvtRowNo + 1
                If imEvtRowNo > vbcEvents.Value + vbcEvents.LargeChange + 1 Then
                    imSettingValue = True
                    vbcEvents.Value = vbcEvents.Value + 1
                End If
            Else
                imEvtRowNo = 1
                imSettingValue = True
                vbcEvents.Value = 0
            End If
        Case KEYLEFT  'Left
            If imEvtBoxNo > TIMEINDEX Then
                imEvtBoxNo = imEvtBoxNo - 1
            Else
                imEvtBoxNo = EVTIDINDEX  'COMMENTINDEX
            End If
        Case KEYRIGHT  'Right
            If imEvtBoxNo < COMMENTINDEX Then
                imEvtBoxNo = imEvtBoxNo + 1
            Else
                imEvtBoxNo = TIMEINDEX
            End If
    End Select
    imSettingValue = False
    mEvtEnableBox imEvtBoxNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtEnableBox                   *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEvtEnableBox(ilBoxNo As Integer)
'
'   mEvtEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilFound1 As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim slXMid As String

    If ilBoxNo < imLBEvtCtrls Or ilBoxNo > UBound(tmEvtCtrls) Then
        Exit Sub
    End If

    If (imEvtRowNo < vbcEvents.Value + 1) Or (imEvtRowNo >= vbcEvents.Value + vbcEvents.LargeChange + 2) Then
        mEvtSetShow ilBoxNo
        pbcArrow.Visible = False
        lacEvtFrame.Visible = False
        Exit Sub
    End If
    lacEvtFrame.Move 0, tmEvtCtrls(TIMEINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) - 30
    lacEvtFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcEvents.Top + tmEvtCtrls(TIMEINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    If ilBoxNo > EVTTYPEINDEX Then
        gFindMatch smSave(2, imEvtRowNo), 1, lbcEvtType
        If gLastFound(lbcEvtType) > 0 Then
            imEvtNameIndex = gLastFound(lbcEvtType)
        Else
            Exit Sub
        End If
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case TIMEINDEX
            edcEvtDropDown.Width = tmEvtCtrls(ilBoxNo).fBoxW
            If imTimeRelative Then
                edcEvtDropDown.MaxLength = 9
            Else
                edcEvtDropDown.MaxLength = 10
            End If
            gMoveTableCtrl pbcEvents, edcEvtDropDown, tmEvtCtrls(ilBoxNo).fBoxX, tmEvtCtrls(ilBoxNo).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            cmcEvtDropDown.Move edcEvtDropDown.Left + edcEvtDropDown.Width, edcEvtDropDown.Top
            If imTimeRelative Then
                If edcEvtDropDown.Top + edcEvtDropDown.height + plclen.height < cmcDone.Top Then
                    plclen.Move edcEvtDropDown.Left, edcEvtDropDown.Top + edcEvtDropDown.height
                Else
                    plclen.Move edcEvtDropDown.Left, edcEvtDropDown.Top - plclen.height
                End If
            Else
                If edcEvtDropDown.Top + edcEvtDropDown.height + plcTme.height < cmcDone.Top Then
                    plcTme.Move edcEvtDropDown.Left, edcEvtDropDown.Top + edcEvtDropDown.height
                Else
                    plcTme.Move edcEvtDropDown.Left, edcEvtDropDown.Top - plcTme.height
                End If
            End If
            If smSave(1, imEvtRowNo) = "" Then
                If imEvtRowNo <= 1 Then
                    If imTimeRelative Then
                        edcEvtDropDown.Text = "0s"
                    Else
                        edcEvtDropDown.Text = smSpecSave(4)
                    End If
                Else
                    gFindMatch smSave(2, imEvtRowNo - 1), 0, lbcEvtType
                    If gLastFound(lbcEvtType) >= 0 Then
                        slNameCode = tmEvtTypeCode(gLastFound(lbcEvtType) - 1).sKey  'lbcEvtTypeCode.List(gLastFound(lbcEvtType) - 1)
                        ilRet = gParseItem(slNameCode, 3, "\", slCode)
                        If ilRet = CP_MSG_NONE Then
                            ilCode = Val(slCode)
                            If (ilCode = 1) Or (ilCode = 10) Or (ilCode = 11) Or (ilCode = 12) Or (ilCode = 13) Then
                                If imTimeRelative Then
                                    edcEvtDropDown.Text = smSave(1, imEvtRowNo - 1)
                                Else
                                    gAddTimeLength smSpecSave(4), smSave(1, imEvtRowNo - 1), "A", "1", slStr, slXMid
                                    edcEvtDropDown.Text = slStr
                                End If
                            Else
                                If imTimeRelative Then
                                    gAddLengths smSave(1, imEvtRowNo - 1), smSave(7, imEvtRowNo - 1), "3", slStr
                                    edcEvtDropDown.Text = slStr
                                Else
                                    gAddTimeLength smSpecSave(4), smSave(1, imEvtRowNo - 1), "A", "1", slStr, slXMid
                                    gAddTimeLength slStr, smSave(7, imEvtRowNo - 1), "A", "1", slStr, slXMid
                                    edcEvtDropDown.Text = slStr
                                End If
                            End If
                        Else
                            edcEvtDropDown.Text = ""
                        End If
                    Else
                        edcEvtDropDown.Text = ""
                    End If
                End If
            Else
                If imTimeRelative Then
                    edcEvtDropDown.Text = Trim$(smSave(1, imEvtRowNo))
                Else
                    gAddTimeLength smSpecSave(4), smSave(1, imEvtRowNo), "A", "1", slStr, slXMid
                    edcEvtDropDown.Text = slStr
                End If
            End If
            edcEvtDropDown.SelStart = 0
            edcEvtDropDown.SelLength = Len(edcEvtDropDown.Text)
            edcEvtDropDown.Visible = True  'Set visibility
            cmcEvtDropDown.Visible = True
            If imEvtRowNo >= UBound(smSave, 2) Then   'New lines set after all fields entered
                If imTimeRelative Then
                    plclen.Visible = True
                Else
                    plcTme.Visible = True
                End If
            End If
            edcEvtDropDown.SetFocus
        Case EVTTYPEINDEX 'Event Type
            mEvtTypePop
            If imTerminate Then
                Exit Sub
            End If
            lbcEvtType.height = gListBoxHeight(lbcEvtType.ListCount, 10)
            edcEvtDropDown.Width = tmEvtCtrls(EVTTYPEINDEX).fBoxW
            edcEvtDropDown.MaxLength = 20
            gMoveTableCtrl pbcEvents, edcEvtDropDown, tmEvtCtrls(ilBoxNo).fBoxX, tmEvtCtrls(ilBoxNo).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15)
            cmcEvtDropDown.Move edcEvtDropDown.Left + edcEvtDropDown.Width, edcEvtDropDown.Top
            imChgMode = True
            gFindMatch smSave(2, imEvtRowNo), 0, lbcEvtType
            If gLastFound(lbcEvtType) >= 0 Then
                lbcEvtType.ListIndex = gLastFound(lbcEvtType)
                edcEvtDropDown.Text = lbcEvtType.List(lbcEvtType.ListIndex)
            Else
                If lbcEvtType.ListCount > 1 Then
                    If imEvtRowNo > 1 Then
                        gFindMatch smEvtAvName, 1, lbcEvtType
                        If gLastFound(lbcEvtType) > 0 Then
                            lbcEvtType.ListIndex = gLastFound(lbcEvtType)    'Avails
                        Else
                            lbcEvtType.ListIndex = 1
                        End If
                        edcEvtDropDown.Text = lbcEvtType.List(lbcEvtType.ListIndex)
                    Else
                        lbcEvtType.ListIndex = 1
                    End If
                    edcEvtDropDown.Text = lbcEvtType.List(lbcEvtType.ListIndex)
                End If
            End If
            imChgMode = False
            If edcEvtDropDown.Top + edcEvtDropDown.height + lbcEvtType.height < cmcDone.Top Then
                lbcEvtType.Move edcEvtDropDown.Left, edcEvtDropDown.Top + edcEvtDropDown.height
            Else
                lbcEvtType.Move edcEvtDropDown.Left, edcEvtDropDown.Top - lbcEvtType.height
            End If
            edcEvtDropDown.SelStart = 0
            edcEvtDropDown.SelLength = Len(edcEvtDropDown.Text)
            edcEvtDropDown.Visible = True
            cmcEvtDropDown.Visible = True
            edcEvtDropDown.SetFocus
        Case EVTNAMEINDEX 'Event Name
            mEvtNamePop imEvtNameIndex, lbcEvtName(imEvtNameIndex), lbcEvtNameCode(imEvtNameIndex)
            If imTerminate Then
                Exit Sub
            End If
            gFindMatch smSave(2, imEvtRowNo), 1, lbcEvtType
            If gLastFound(lbcEvtType) <= 0 Then
                pbcSTab.SetFocus 'Go back to event type
                Exit Sub
            End If
            imEvtNameIndex = gLastFound(lbcEvtType)
            lbcEvtName(imEvtNameIndex).height = gListBoxHeight(lbcEvtName(imEvtNameIndex).ListCount, 10)
            edcEvtDropDown.Width = tmEvtCtrls(ilBoxNo).fBoxW
            edcEvtDropDown.MaxLength = 30
            gMoveTableCtrl pbcEvents, edcEvtDropDown, tmEvtCtrls(ilBoxNo).fBoxX, tmEvtCtrls(ilBoxNo).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15)
            cmcEvtDropDown.Move edcEvtDropDown.Left + edcEvtDropDown.Width, edcEvtDropDown.Top
            imChgMode = True
            gFindMatch smSave(3, imEvtRowNo), 0, lbcEvtName(imEvtNameIndex)
            If gLastFound(lbcEvtName(imEvtNameIndex)) >= 0 Then
                lbcEvtName(imEvtNameIndex).ListIndex = gLastFound(lbcEvtName(imEvtNameIndex))
                edcEvtDropDown.Text = lbcEvtName(imEvtNameIndex).List(lbcEvtName(imEvtNameIndex).ListIndex)
            Else
                If imEvtRowNo > 1 Then
                    ilFound1 = False
                    For ilLoop = imEvtRowNo - 1 To 1 Step -1
                        If StrComp(smSave(2, imEvtRowNo), smSave(2, ilLoop), 1) = 0 Then
                            gFindMatch smSave(3, ilLoop), 0, lbcEvtName(imEvtNameIndex)
                            If gLastFound(lbcEvtName(imEvtNameIndex)) >= 0 Then
                                lbcEvtName(imEvtNameIndex).ListIndex = gLastFound(lbcEvtName(imEvtNameIndex))
                                edcEvtDropDown.Text = lbcEvtName(imEvtNameIndex).List(lbcEvtName(imEvtNameIndex).ListIndex)
                                ilFound1 = True
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If Not ilFound1 Then
                        If lbcEvtName(imEvtNameIndex).ListCount <= 1 Then
                            lbcEvtName(imEvtNameIndex).ListIndex = 0
                            edcEvtDropDown.Text = "[New]"
                        Else
                            lbcEvtName(imEvtNameIndex).ListIndex = 1
                            edcEvtDropDown.Text = lbcEvtName(imEvtNameIndex).List(1)
                        End If
                    End If
                Else
                    If lbcEvtName(imEvtNameIndex).ListCount <= 1 Then
                        lbcEvtName(imEvtNameIndex).ListIndex = 0
                        edcEvtDropDown.Text = lbcEvtName(imEvtNameIndex).List(0)
                    Else
                        lbcEvtName(imEvtNameIndex).ListIndex = 1
                        edcEvtDropDown.Text = lbcEvtName(imEvtNameIndex).List(1)
                    End If
                End If
            End If
            imChgMode = False
            If edcEvtDropDown.Top + edcEvtDropDown.height + lbcEvtName(imEvtNameIndex).height < cmcDone.Top Then
                lbcEvtName(imEvtNameIndex).Move edcEvtDropDown.Left, edcEvtDropDown.Top + edcEvtDropDown.height
            Else
                lbcEvtName(imEvtNameIndex).Move edcEvtDropDown.Left, edcEvtDropDown.Top - lbcEvtName(imEvtNameIndex).height
            End If
            lbcEvtName(imEvtNameIndex).ZOrder vbBringToFront
            edcEvtDropDown.SelStart = 0
            edcEvtDropDown.SelLength = Len(edcEvtDropDown.Text)
            edcEvtDropDown.Visible = True
            cmcEvtDropDown.Visible = True
            edcEvtDropDown.SetFocus
        Case AVAILINDEX 'Avail
            mEvtAvailPop
            If imTerminate Then
                Exit Sub
            End If
            If Asc(smSave(9, imEvtRowNo)) = Asc("A") Then
                lbcEvtAvail.List(0) = "[None]"
            Else
                lbcEvtAvail.List(0) = "[New]"
            End If
            lbcEvtAvail.height = gListBoxHeight(lbcEvtAvail.ListCount, 10)
            edcEvtDropDown.Width = tmEvtCtrls(ilBoxNo).fBoxW
            edcEvtDropDown.MaxLength = 20
            gMoveTableCtrl pbcEvents, edcEvtDropDown, tmEvtCtrls(ilBoxNo).fBoxX, tmEvtCtrls(ilBoxNo).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15)
            cmcEvtDropDown.Move edcEvtDropDown.Left + edcEvtDropDown.Width, edcEvtDropDown.Top
            gFindMatch smSave(4, imEvtRowNo), 0, lbcEvtAvail
            imChgMode = True
            If gLastFound(lbcEvtAvail) >= 0 Then
                lbcEvtAvail.ListIndex = gLastFound(lbcEvtAvail)
                edcEvtDropDown.Text = lbcEvtAvail.List(lbcEvtAvail.ListIndex)
            Else
                If imEvtRowNo > 1 Then
                    ilFound1 = False
                    For ilLoop = imEvtRowNo - 1 To 1 Step -1
                        If StrComp(smSave(2, imEvtRowNo), smSave(2, ilLoop), 1) = 0 Then
                            gFindMatch smSave(4, ilLoop), 0, lbcEvtAvail
                            If gLastFound(lbcEvtAvail) >= 0 Then
                                lbcEvtAvail.ListIndex = gLastFound(lbcEvtAvail)
                                edcEvtDropDown.Text = lbcEvtAvail.List(lbcEvtAvail.ListIndex)
                                ilFound1 = True
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If Not ilFound1 Then
                        If lbcEvtAvail.ListCount <= 1 Then
                            lbcEvtAvail.ListIndex = -1
                            edcEvtDropDown.Text = ""
                        Else
                            If Asc(smSave(9, imEvtRowNo)) = Asc("A") Then
                                lbcEvtAvail.ListIndex = 0
                                edcEvtDropDown.Text = lbcEvtAvail.List(0)
                            Else
                                lbcEvtAvail.ListIndex = 1
                                edcEvtDropDown.Text = lbcEvtAvail.List(1)
                            End If
                        End If
                    End If
                Else
                    If lbcEvtAvail.ListCount <= 1 Then
                        lbcEvtAvail.ListIndex = -1
                        edcEvtDropDown.Text = ""
                    Else
                        lbcEvtAvail.ListIndex = 1
                        edcEvtDropDown.Text = lbcEvtAvail.List(1)
                    End If
                End If
            End If
            imChgMode = False
            If edcEvtDropDown.Top + edcEvtDropDown.height + lbcEvtAvail.height < cmcDone.Top Then
                lbcEvtAvail.Move edcEvtDropDown.Left, edcEvtDropDown.Top + edcEvtDropDown.height
            Else
                lbcEvtAvail.Move edcEvtDropDown.Left, edcEvtDropDown.Top - lbcEvtAvail.height
            End If
            edcEvtDropDown.SelStart = 0
            edcEvtDropDown.SelLength = Len(edcEvtDropDown.Text)
            edcEvtDropDown.Visible = True
            cmcEvtDropDown.Visible = True
            edcEvtDropDown.SetFocus
        Case UNITSINDEX
            edcEvtEdit.Width = tmEvtCtrls(ilBoxNo).fBoxW
            edcEvtEdit.MaxLength = 2
            gMoveTableCtrl pbcEvents, edcEvtEdit, tmEvtCtrls(ilBoxNo).fBoxX, tmEvtCtrls(ilBoxNo).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            edcEvtEdit.Text = Trim$(smSave(6, imEvtRowNo))
            If (smSave(6, imEvtRowNo) = "") And (imEvtRowNo > 1) Then
                For ilLoop = imEvtRowNo - 1 To 1 Step -1
                    If StrComp(smSave(2, imEvtRowNo), smSave(2, ilLoop), 1) = 0 Then
                        edcEvtEdit.Text = Trim$(smSave(6, ilLoop))
                        Exit For
                    End If
                Next ilLoop
            End If
            If (edcEvtEdit.Text = "") And (tgVpf(imVpfIndex).sSSellOut = "U") Then
                edcEvtEdit.Text = "1"
            End If
            edcEvtEdit.Visible = True  'Set visibility
            edcEvtEdit.SetFocus
        Case LENGTHINDEX
            edcEvtDropDown.Width = tmEvtCtrls(ilBoxNo).fBoxW
            edcEvtDropDown.MaxLength = 9
            gMoveTableCtrl pbcEvents, edcEvtDropDown, tmEvtCtrls(ilBoxNo).fBoxX, tmEvtCtrls(ilBoxNo).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            cmcEvtDropDown.Move edcEvtDropDown.Left + edcEvtDropDown.Width, edcEvtDropDown.Top
            If edcEvtDropDown.Top + edcEvtDropDown.height + plclen.height < cmcDone.Top Then
                plclen.Move edcEvtDropDown.Left, edcEvtDropDown.Top + edcEvtDropDown.height
            Else
                plclen.Move edcEvtDropDown.Left, edcEvtDropDown.Top - plclen.height
            End If
            edcEvtDropDown.Text = Trim$(smSave(7, imEvtRowNo))
            If (smSave(7, imEvtRowNo) = "") And (imEvtRowNo > 1) Then
                For ilLoop = imEvtRowNo - 1 To 1 Step -1
                    If StrComp(smSave(2, imEvtRowNo), smSave(2, ilLoop), 1) = 0 Then
                        edcEvtDropDown.Text = Trim$(smSave(7, ilLoop))
                        Exit For
                    End If
                Next ilLoop
            End If
            edcEvtDropDown.SelStart = 0
            edcEvtDropDown.SelLength = Len(edcEvtDropDown.Text)
            edcEvtDropDown.Visible = True  'Set visibility
            cmcEvtDropDown.Visible = True
            plclen.Visible = True
            edcEvtDropDown.SetFocus
        Case TRUETIMEINDEX    'Source Index
            gMoveTableCtrl pbcEvents, pbcTrueTime, tmEvtCtrls(ilBoxNo).fBoxX, tmEvtCtrls(ilBoxNo).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            pbcTrueTime_Paint
            pbcTrueTime.Visible = True
            pbcTrueTime.SetFocus
        Case EVTIDINDEX
            edcEvtDropDown.Width = tmEvtCtrls(ilBoxNo).fBoxW
            edcEvtDropDown.MaxLength = 0
            gMoveTableCtrl pbcEvents, edcEvtDropDown, tmEvtCtrls(ilBoxNo).fBoxX, tmEvtCtrls(ilBoxNo).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            edcEvtDropDown.Text = Trim$(smSave(11, imEvtRowNo))
            edcEvtDropDown.Visible = True  'Set visibility
            edcEvtDropDown.SetFocus
        Case COMMENTINDEX
            If (smSave(8, imEvtRowNo) = "^") Then
                smSave(8, imEvtRowNo) = ""
                'Obtain the default comment from the event name
                If (Asc(smSave(9, imEvtRowNo)) < Asc("A")) Or (Asc(smSave(9, imEvtRowNo)) > Asc("D")) Then
                    If mReadEnfRec(imEvtRowNo) Then
                        If mReadCefRec(tmEnf.lCefCode, SETFORREADONLY) Then
                            'If tmCef.iStrLen > 0 Then
                            '    smSave(8, imEvtRowNo) = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                                smSave(8, imEvtRowNo) = gStripChr0(tmCef.sComment)
                            'End If
                        End If
                    End If
                End If
            End If
            gMoveTableCtrl pbcEvents, edcComment, tmEvtCtrls(ilBoxNo).fBoxX, tmEvtCtrls(ilBoxNo).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            edcComment.Text = Trim$(smSave(8, imEvtRowNo))
            edcComment.Visible = True  'Set visibility
            edcComment.SetFocus
        Case EXCL1INDEX
            mExclPop
            If imTerminate Then
                Exit Sub
            End If
            lbcExcl(0).height = gListBoxHeight(lbcExcl(0).ListCount, 10)
            edcEvtDropDown.Width = tmEvtCtrls(AVAILINDEX).fBoxW
            edcEvtDropDown.MaxLength = 20
            gMoveTableCtrl pbcEvents, edcEvtDropDown, tmEvtCtrls(AVAILINDEX).fBoxX, tmEvtCtrls(AVAILINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15)
            cmcEvtDropDown.Move edcEvtDropDown.Left + edcEvtDropDown.Width, edcEvtDropDown.Top
            imChgMode = True
            gFindMatch smSave(4, imEvtRowNo), 0, lbcExcl(0) 'Avail name or exclusion
            If gLastFound(lbcExcl(0)) >= 0 Then
                lbcExcl(0).ListIndex = gLastFound(lbcExcl(0))
                edcEvtDropDown.Text = lbcExcl(0).List(lbcExcl(0).ListIndex)
            Else
                lbcExcl(0).ListIndex = 1    '[None]
                edcEvtDropDown.Text = lbcExcl(0).List(1)
            End If
            imChgMode = False
            If edcEvtDropDown.Top + edcEvtDropDown.height + lbcExcl(0).height < cmcDone.Top Then
                lbcExcl(0).Move edcEvtDropDown.Left, edcEvtDropDown.Top + edcEvtDropDown.height
            Else
                lbcExcl(0).Move edcEvtDropDown.Left, edcEvtDropDown.Top - lbcExcl(0).height
            End If
            edcEvtDropDown.SelStart = 0
            edcEvtDropDown.SelLength = Len(edcEvtDropDown.Text)
            edcEvtDropDown.Visible = True
            cmcEvtDropDown.Visible = True
            edcEvtDropDown.SetFocus
        Case EXCL2INDEX
            mExclPop
            If imTerminate Then
                Exit Sub
            End If
            lbcExcl(1).height = gListBoxHeight(lbcExcl(1).ListCount, 10)
            edcEvtDropDown.Width = tmEvtCtrls(AVAILINDEX).fBoxW
            edcEvtDropDown.MaxLength = 20
            gMoveTableCtrl pbcEvents, edcEvtDropDown, tmEvtCtrls(AVAILINDEX).fBoxX + tmEvtCtrls(AVAILINDEX).fBoxW / 4, tmEvtCtrls(AVAILINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15)
            cmcEvtDropDown.Move edcEvtDropDown.Left + edcEvtDropDown.Width, edcEvtDropDown.Top
            imChgMode = True
            gFindMatch smSave(5, imEvtRowNo), 0, lbcExcl(1)
            If gLastFound(lbcExcl(1)) >= 0 Then
                lbcExcl(1).ListIndex = gLastFound(lbcExcl(1))
                edcEvtDropDown.Text = lbcExcl(1).List(lbcExcl(1).ListIndex)
            Else
                lbcExcl(1).ListIndex = 1    '[None]
                edcEvtDropDown.Text = lbcExcl(1).List(1)
            End If
            imChgMode = False
            If edcEvtDropDown.Top + edcEvtDropDown.height + lbcExcl(1).height < cmcDone.Top Then
                lbcExcl(1).Move edcEvtDropDown.Left, edcEvtDropDown.Top + edcEvtDropDown.height
            Else
                lbcExcl(1).Move edcEvtDropDown.Left, edcEvtDropDown.Top - lbcExcl(1).height
            End If
            edcEvtDropDown.SelStart = 0
            edcEvtDropDown.SelLength = Len(edcEvtDropDown.Text)
            edcEvtDropDown.Visible = True
            cmcEvtDropDown.Visible = True
            edcEvtDropDown.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtNameBranch                  *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to event  *
'*                      Name and process               *
'*                      communication back from event  *
'*                      Name                           *
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
Private Function mEvtNameBranch() As Integer
'
'   ilRet = mEvtTypeBranch()
'   Where:
'       ilRet (O)- True = event type started
'                  False = Event type over or not started
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcEvtDropDown, lbcEvtName(imEvtNameIndex), imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcEvtDropDown.Text = "[None]") Then
        mEvtNameBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(EVENTNAMESLIST)) Then
    '    imDoubleClickName = False
    '    mEvtNameBranch = True
    '    mEvtEnableBox imEvtBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourglass  'Wait
    igENameCallSource = CALLSOURCEPEVENT
    If edcEvtDropDown.Text = "[New]" Then
        sgENameName = smVehName & "\" & smSave(2, imEvtRowNo)
    Else
        sgENameName = smVehName & "\" & smSave(2, imEvtRowNo) & "\" & slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "PEvent^Test\" & sgUserName & "\" & Trim$(str$(igENameCallSource)) & "\" & sgENameName
        Else
            slStr = "PEvent^Prod\" & sgUserName & "\" & Trim$(str$(igENameCallSource)) & "\" & sgENameName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "PEvent^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igENameCallSource)) & "\" & sgENameName
    '    Else
    '        slStr = "PEvent^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igENameCallSource)) & "\" & sgENameName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "EName.Exe " & slStr, 1)
    'PEvent.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    EName.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgENameName)
    igENameCallSource = Val(sgENameName)
    ilParse = gParseItem(slStr, 2, "\", sgENameName)
    'PEvent.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mEvtNameBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igENameCallSource = CALLDONE Then  'Done
        igENameCallSource = CALLNONE
'        gSetMenuState True
        'Clear Event type instead of Event name as the user
        'might have altered event types when in event names.
        lbcEvtType.Clear
        smEvtTypeCodeTag = ""
        mEvtTypePop
'        lbcEvtName(imEvtNameIndex).Clear
'        mEvtNamePop imEvtNameIndex, lbcEvtName(imEvtNameIndex), lbcEvtNameCode(imEvtNameIndex)
        If imTerminate Then
            mEvtNameBranch = False
            Exit Function
        End If
        gFindMatch sgENameName, 2, lbcEvtName(imEvtNameIndex)
        If gLastFound(lbcEvtName(imEvtNameIndex)) > 1 Then
            imChgMode = True
            lbcEvtName(imEvtNameIndex).ListIndex = gLastFound(lbcEvtName(imEvtNameIndex))
            edcEvtDropDown.Text = lbcEvtName(imEvtNameIndex).List(lbcEvtName(imEvtNameIndex).ListIndex)
            imChgMode = False
            mEvtNameBranch = False
        Else
            imChgMode = True
            lbcEvtName(imEvtNameIndex).ListIndex = 1
            edcEvtDropDown.Text = lbcEvtName(imEvtNameIndex).List(1)
            imChgMode = False
            edcEvtDropDown.SetFocus
            sgENameName = ""
            Exit Function
        End If
        sgENameName = ""
    End If
    If igENameCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igENameCallSource = CALLNONE
        sgENameName = ""
        mEvtEnableBox imEvtBoxNo
        Exit Function
    End If
    If igENameCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igENameCallSource = CALLNONE
        sgENameName = ""
        mEvtEnableBox imEvtBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtNamePop                     *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: EvtNamePop the selection event *
'*                      name box                       *
'*                                                     *
'*******************************************************
Private Sub mEvtNamePop(ilEvtNameIndex As Integer, lbcName As control, lbcNameCode As control)
'
'   mEvtNamePop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slNameCode As String
    Dim slName As String
    Dim ilLoop As Integer
    Dim slCode As String
    ReDim ilFilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    If ilEvtNameIndex >= 1 Then 'Event index
        ilFilter(0) = INTEGERFILTER
        slFilter(0) = Trim$(str$(imVefCode))
        ilOffSet(0) = gFieldOffset("Enf", "EnfVefCode") '2

        slNameCode = tmEvtTypeCode(ilEvtNameIndex - 1).sKey    'lbcEvtTypeCode.List(ilEvtNameIndex - 1)
        ilRet = gParseItem(slNameCode, 3, "\", slCode)
        On Error GoTo mEvtNamePopErr
        gCPErrorMsg ilRet, "mEvtNamePop (gParseItem field 3)", PEvent
        On Error GoTo 0
        ilFilter(1) = INTEGERFILTER
        slFilter(1) = slCode
        ilOffSet(1) = gFieldOffset("Enf", "EnfEtfCode") '4
        ReDim tgTmpSort(0 To 0) As SORTCODE
        For ilLoop = 0 To lbcNameCode.ListCount - 1 Step 1
            slName = lbcNameCode.List(ilLoop)
            gAddItemToSortCode slName, tgTmpSort(), True
        Next ilLoop
        sgTmpSortTag = lbcNameCode.Tag
        'ilRet = gIMoveListBox(PEvent, lbcName, lbcNameCode, "Enf.Btr", gFieldOffset("Enf", "EnfName"), 30, ilFilter(), slFilter(), ilOffset())
        ilRet = gIMoveListBox(PEvent, lbcName, tgTmpSort(), sgTmpSortTag, "Enf.Btr", gFieldOffset("Enf", "EnfName"), 30, ilFilter(), slFilter(), ilOffSet())
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mEvtNamePopErr
            gCPErrorMsg ilRet, "mEvtNamePop (gIMoveListBox)", PEvent
            On Error GoTo 0
            lbcName.AddItem "[None]", 0  'Force as first item on list
            lbcName.AddItem "[New]", 0  'Force as first item on list
            lbcNameCode.Clear
            For ilLoop = 0 To UBound(tgTmpSort) - 1 Step 1
                lbcNameCode.AddItem Trim$(tgTmpSort(ilLoop).sKey), ilLoop
            Next ilLoop
            lbcNameCode.Tag = sgTmpSortTag
        End If
    End If
    Exit Sub
mEvtNamePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtSetFocus                    *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mEvtSetFocus(ilBoxNo As Integer)
'
'   mEvtSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBEvtCtrls Or ilBoxNo > UBound(tmEvtCtrls) Then
        Exit Sub
    End If

    If (imEvtRowNo < vbcEvents.Value + 1) Or (imEvtRowNo >= vbcEvents.Value + vbcEvents.LargeChange + 2) Then
        mEvtSetShow ilBoxNo
        pbcArrow.Visible = False
        lacEvtFrame.Visible = False
        Exit Sub
    End If
    lacEvtFrame.Move 0, tmEvtCtrls(TIMEINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) - 30
    lacEvtFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcEvents.Top + tmEvtCtrls(TIMEINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case TIMEINDEX
            If edcEvtDropDown.Enabled Then
                edcEvtDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case EVTTYPEINDEX 'Event Type
            If edcEvtDropDown.Enabled Then
                edcEvtDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case EVTNAMEINDEX 'Event Name
            If edcEvtDropDown.Enabled Then
                edcEvtDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case AVAILINDEX 'Avail
            If edcEvtDropDown.Enabled Then
                edcEvtDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case UNITSINDEX
            If edcEvtEdit.Enabled Then
                edcEvtEdit.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case LENGTHINDEX
            If edcEvtDropDown.Enabled Then
                edcEvtDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case TRUETIMEINDEX    'Source Index
            If pbcTrueTime.Enabled Then
                pbcTrueTime.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case EVTIDINDEX
            If edcEvtDropDown.Enabled Then
                edcEvtDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case COMMENTINDEX
            If edcComment.Enabled Then
                edcComment.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case EXCL1INDEX
            If edcEvtDropDown.Enabled Then
                edcEvtDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case EXCL2INDEX
            If edcEvtDropDown.Enabled Then
                edcEvtDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtSetShow                     *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mEvtSetShow(ilBoxNo As Integer)
'
'   mEvtSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim flWidth As Single
    Dim slXMid As String

    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    If (ilBoxNo < imLBEvtCtrls) Or (ilBoxNo > UBound(tmEvtCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case TIMEINDEX 'Time index
            If imTimeRelative Then
                plclen.Visible = False
            Else
                plcTme.Visible = False
            End If
            cmcEvtDropDown.Visible = False
            edcEvtDropDown.Visible = False  'Set visibility
            slStr = edcEvtDropDown.Text
            If imTimeRelative Then
                If gValidLength(slStr) Then
                    slStr = gFormatLength(slStr, "3", False)
                    gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                    smShow(ilBoxNo, imEvtRowNo) = tmEvtCtrls(ilBoxNo).sShow
                    If StrComp(smSave(1, imEvtRowNo), edcEvtDropDown.Text, 1) <> 0 Then
                        If imEvtRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                            imLefChg = True
                        End If
                    End If
                    smSave(1, imEvtRowNo) = edcEvtDropDown.Text
                Else
                    Beep
                    edcEvtDropDown.Text = smSave(1, imEvtRowNo)
                End If
            Else
                If gValidTime(slStr) Then
                    slStr = gFormatTime(slStr, "A", "1")
                    gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                    smShow(ilBoxNo, imEvtRowNo) = tmEvtCtrls(ilBoxNo).sShow
                    'Convert to relative time
                    slStr = gCurrencyToLength(gTimeToCurrency(slStr, False) - gTimeToCurrency(smSpecSave(4), False))
                    If gLengthToCurrency(smSave(1, imEvtRowNo)) <> gLengthToCurrency(slStr) Then
                        If imEvtRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                            imLefChg = True
                        End If
                    End If
                    smSave(1, imEvtRowNo) = slStr
                Else
                    Beep
                    If imTimeRelative Then
                        edcEvtDropDown.Text = smSave(1, imEvtRowNo)
                    Else
                        gAddTimeLength smSpecSave(4), smSave(1, imEvtRowNo), "A", "1", slStr, slXMid
                        edcEvtDropDown.Text = slStr
                    End If
                End If
            End If
        Case EVTTYPEINDEX 'Event Type
            lbcEvtType.Visible = False
            edcEvtDropDown.Visible = False
            cmcEvtDropDown.Visible = False
            If lbcEvtType.ListIndex > 0 Then
                slStr = edcEvtDropDown.Text
            Else
                slStr = ""
            End If
            gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
            smShow(ilBoxNo, imEvtRowNo) = tmEvtCtrls(ilBoxNo).sShow
            If StrComp(smSave(2, imEvtRowNo), edcEvtDropDown.Text, 1) <> 0 Then
                If lbcEvtType.ListIndex > 0 Then
                    smSave(2, imEvtRowNo) = lbcEvtType.List(lbcEvtType.ListIndex)
                    slNameCode = tmEvtTypeCode(lbcEvtType.ListIndex - 1).sKey    'lbcEvtTypeCode.List(imEvtNameIndex - 1)
                    ilRet = gParseItem(slNameCode, 1, "\", smSave(9, imEvtRowNo))
                    If imEvtRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                        imLefChg = True
                    End If
                Else
                    smSave(2, imEvtRowNo) = ""
                    smSave(9, imEvtRowNo) = ""
                    imLefChg = True
                End If
                'Clear other fields as event type changed
                smSave(3, imEvtRowNo) = ""  'Event Name
                slStr = ""
                gSetShow pbcEvents, slStr, tmEvtCtrls(EVTNAMEINDEX)
                smShow(EVTNAMEINDEX, imEvtRowNo) = tmEvtCtrls(EVTNAMEINDEX).sShow
                smSave(4, imEvtRowNo) = ""  'Avail or exclusion
                smSave(5, imEvtRowNo) = ""  'Exclusion
                slStr = ""
                gSetShow pbcEvents, slStr, tmEvtCtrls(AVAILINDEX)
                smShow(AVAILINDEX, imEvtRowNo) = tmEvtCtrls(AVAILINDEX).sShow
                smSave(6, imEvtRowNo) = ""  'Units
                slStr = smSave(6, imEvtRowNo)
                gSetShow pbcEvents, slStr, tmEvtCtrls(UNITSINDEX)
                smShow(UNITSINDEX, imEvtRowNo) = tmEvtCtrls(UNITSINDEX).sShow
                smSave(7, imEvtRowNo) = ""  'Length
                slStr = smSave(7, imEvtRowNo)
                gSetShow pbcEvents, slStr, tmEvtCtrls(LENGTHINDEX)
                smShow(LENGTHINDEX, imEvtRowNo) = tmEvtCtrls(LENGTHINDEX).sShow
                imSave(1, imEvtRowNo) = 1   'True Time (default to No)
                slStr = ""
                gSetShow pbcEvents, slStr, tmEvtCtrls(TRUETIMEINDEX)
                smShow(TRUETIMEINDEX, imEvtRowNo) = tmEvtCtrls(TRUETIMEINDEX).sShow
                smSave(11, imEvtRowNo) = ""  'Event ID
                slStr = ""
                gSetShow pbcEvents, slStr, tmEvtCtrls(EVTIDINDEX)
                smShow(EVTIDINDEX, imEvtRowNo) = tmEvtCtrls(EVTIDINDEX).sShow
                smSave(8, imEvtRowNo) = "^"  'Comment-flag as not set
                slStr = ""
                gSetShow pbcEvents, slStr, tmEvtCtrls(COMMENTINDEX)
                smShow(COMMENTINDEX, imEvtRowNo) = tmEvtCtrls(COMMENTINDEX).sShow
                pbcEvents_Paint
            End If
            imEvtNameIndex = lbcEvtType.ListIndex
            If imEvtNameIndex > 0 Then
                slNameCode = tmEvtTypeCode(imEvtNameIndex - 1).sKey    'lbcEvtTypeCode.List(imEvtNameIndex - 1)
                ilRet = gParseItem(slNameCode, 1, "\", smSave(9, imEvtRowNo))
            End If
        Case EVTNAMEINDEX 'Event Name
            lbcEvtName(imEvtNameIndex).Visible = False
            edcEvtDropDown.Visible = False
            cmcEvtDropDown.Visible = False
            If lbcEvtName(imEvtNameIndex).ListIndex > 0 Then
                slStr = edcEvtDropDown.Text
            Else
                slStr = ""
            End If
            gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
            smShow(ilBoxNo, imEvtRowNo) = tmEvtCtrls(ilBoxNo).sShow
            If StrComp(smSave(3, imEvtRowNo), edcEvtDropDown.Text, 1) <> 0 Then
                If lbcEvtName(imEvtNameIndex).ListIndex > 0 Then
                    smSave(3, imEvtRowNo) = lbcEvtName(imEvtNameIndex).List(lbcEvtName(imEvtNameIndex).ListIndex)
                Else
                    smSave(3, imEvtRowNo) = ""
                End If
                If imEvtRowNo < UBound(smSave, 2) Then   'New lines set after all fields entered
                    imLefChg = True
                End If
            End If
        Case AVAILINDEX 'Avail
            lbcEvtAvail.Visible = False
            edcEvtDropDown.Visible = False
            cmcEvtDropDown.Visible = False
            If lbcEvtAvail.ListIndex > 0 Then
                slStr = edcEvtDropDown.Text
            Else
                slStr = ""
            End If
            gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
            smShow(ilBoxNo, imEvtRowNo) = tmEvtCtrls(ilBoxNo).sShow
            If StrComp(smSave(4, imEvtRowNo), edcEvtDropDown.Text, 1) <> 0 Then
                If lbcEvtAvail.ListIndex > 0 Then
                    smSave(4, imEvtRowNo) = lbcEvtAvail.List(lbcEvtAvail.ListIndex)
                Else
                    smSave(4, imEvtRowNo) = ""
                End If
                If imEvtRowNo < UBound(smSave, 2) Then    'New lines set after all fields entered
                    imLefChg = True
                End If
            End If
            lbcEvtAvail.List(0) = "[New]"
        Case UNITSINDEX 'Unit index
            edcEvtEdit.Visible = False  'Set visibility
            slStr = edcEvtEdit.Text
            gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
            smShow(ilBoxNo, imEvtRowNo) = tmEvtCtrls(ilBoxNo).sShow
            If smSave(6, imEvtRowNo) <> edcEvtEdit.Text Then
                smSave(6, imEvtRowNo) = edcEvtEdit.Text
                If imEvtRowNo < UBound(smSave, 2) Then    'New lines set after all fields entered
                    imLefChg = True
                End If
            End If
        Case LENGTHINDEX 'Unit index
            plclen.Visible = False
            cmcEvtDropDown.Visible = False
            edcEvtDropDown.Visible = False  'Set visibility
            slStr = edcEvtDropDown.Text
            If gValidLength(slStr) Then
                If smSave(9, imEvtRowNo) = "1" Then 'Program
                    slStr = gFormatLength(slStr, "3", False)
                Else
                    slStr = gFormatLength(slStr, "3", True)
                End If
                gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                smShow(ilBoxNo, imEvtRowNo) = tmEvtCtrls(ilBoxNo).sShow
                If smSave(7, imEvtRowNo) <> edcEvtDropDown.Text Then
                    smSave(7, imEvtRowNo) = edcEvtDropDown.Text
                    If imEvtRowNo < UBound(smSave, 2) Then    'New lines set after all fields entered
                        imLefChg = True
                    End If
                End If
                'If (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                If (tgVpf(imVpfIndex).sSSellOut = "M") Then
                    If (smSave(9, imEvtRowNo) = "2") Or (smSave(9, imEvtRowNo) = "6") Or (smSave(9, imEvtRowNo) = "7") Or (smSave(9, imEvtRowNo) = "8") Or (smSave(9, imEvtRowNo) = "9") Then
                        If smSave(6, imEvtRowNo) = "" Then
                            smSave(6, imEvtRowNo) = "1"
                            slStr = "1"
                            gSetShow pbcEvents, slStr, tmEvtCtrls(UNITSINDEX)
                            smShow(UNITSINDEX, imEvtRowNo) = tmEvtCtrls(UNITSINDEX).sShow
                            If imEvtRowNo < UBound(smSave, 2) Then    'New lines set after all fields entered
                                imLefChg = True
                            End If
                        End If
                    End If
                End If
            Else
                Beep
                edcEvtDropDown.Text = smSave(7, imEvtRowNo)
            End If
        Case TRUETIMEINDEX
            pbcTrueTime.Visible = False
            If imSave(1, imEvtRowNo) = 0 Then
                slStr = "Yes"
            Else
                slStr = "No"
            End If
            gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
            smShow(ilBoxNo, imEvtRowNo) = tmEvtCtrls(ilBoxNo).sShow
        Case EVTIDINDEX 'Event ID
            edcEvtDropDown.Visible = False  'Set visibility
            slStr = edcEvtDropDown.Text
            gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
            smShow(ilBoxNo, imEvtRowNo) = tmEvtCtrls(ilBoxNo).sShow
            If smSave(11, imEvtRowNo) <> edcEvtDropDown.Text Then
                smSave(11, imEvtRowNo) = edcEvtDropDown.Text
                If imEvtRowNo < UBound(smSave, 2) Then    'New lines set after all fields entered
                    imLefChg = True
                End If
            End If
        Case COMMENTINDEX 'Comment
            edcComment.Visible = False  'Set visibility
            slStr = edcComment.Text
            gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
            smShow(ilBoxNo, imEvtRowNo) = tmEvtCtrls(ilBoxNo).sShow
            If smSave(8, imEvtRowNo) <> edcComment.Text Then
                smSave(8, imEvtRowNo) = edcComment.Text
                If imEvtRowNo < UBound(smSave, 2) Then    'New lines set after all fields entered
                    imLefChg = True
                End If
            End If
        Case EXCL1INDEX 'Avail
            lbcExcl(0).Visible = False
            edcEvtDropDown.Visible = False
            cmcEvtDropDown.Visible = False
            If smSave(5, imEvtRowNo) = "" Then
                If lbcExcl(0).ListIndex > 0 Then
                    slStr = edcEvtDropDown.Text
                Else
                    slStr = ""
                End If
                gSetShow pbcEvents, slStr, tmEvtCtrls(AVAILINDEX)
                smShow(AVAILINDEX, imEvtRowNo) = tmEvtCtrls(AVAILINDEX).sShow
            Else
                flWidth = tmEvtCtrls(AVAILINDEX).fBoxW
                tmEvtCtrls(AVAILINDEX).fBoxW = tmEvtCtrls(AVAILINDEX).fBoxW / 2
                If lbcExcl(0).ListIndex > 0 Then
                    slStr = edcEvtDropDown.Text
                Else
                    slStr = ""
                End If
                gSetShow pbcEvents, slStr, tmEvtCtrls(AVAILINDEX)
                tmEvtCtrls(AVAILINDEX).fBoxW = flWidth
                slStr = tmEvtCtrls(AVAILINDEX).sShow & "/" & smSave(5, imEvtRowNo)
                gSetShow pbcEvents, slStr, tmEvtCtrls(AVAILINDEX)
                smShow(AVAILINDEX, imEvtRowNo) = tmEvtCtrls(AVAILINDEX).sShow
            End If
            If StrComp(smSave(4, imEvtRowNo), edcEvtDropDown.Text, 1) <> 0 Then
                If lbcExcl(0).ListIndex > 0 Then
                    smSave(4, imEvtRowNo) = lbcExcl(0).List(lbcExcl(0).ListIndex)
                Else
                    smSave(4, imEvtRowNo) = ""
                End If
                If imEvtRowNo < UBound(smSave, 2) Then    'New lines set after all fields entered
                    imLefChg = True
                End If
            End If
        Case EXCL2INDEX 'Avail
            lbcExcl(1).Visible = False
            edcEvtDropDown.Visible = False
            cmcEvtDropDown.Visible = False
            If smSave(4, imEvtRowNo) = "" Then
                If lbcExcl(1).ListIndex > 0 Then
                    slStr = edcEvtDropDown.Text
                Else
                    slStr = ""
                End If
                gSetShow pbcEvents, slStr, tmEvtCtrls(AVAILINDEX)
                smShow(AVAILINDEX, imEvtRowNo) = tmEvtCtrls(AVAILINDEX).sShow
            Else
                flWidth = tmEvtCtrls(AVAILINDEX).fBoxW
                tmEvtCtrls(AVAILINDEX).fBoxW = tmEvtCtrls(AVAILINDEX).fBoxW / 2
                If lbcExcl(1).ListIndex > 0 Then
                    slStr = smSave(4, imEvtRowNo)
                Else
                    slStr = ""
                End If
                gSetShow pbcEvents, slStr, tmEvtCtrls(AVAILINDEX)
                tmEvtCtrls(AVAILINDEX).fBoxW = flWidth
                slStr = tmEvtCtrls(AVAILINDEX).sShow & "/" & edcEvtDropDown.Text
                gSetShow pbcEvents, slStr, tmEvtCtrls(AVAILINDEX)
                smShow(AVAILINDEX, imEvtRowNo) = tmEvtCtrls(AVAILINDEX).sShow
            End If
            If StrComp(smSave(5, imEvtRowNo), edcEvtDropDown.Text, 1) <> 0 Then
                If lbcExcl(1).ListIndex > 0 Then
                    smSave(5, imEvtRowNo) = lbcExcl(1).List(lbcExcl(1).ListIndex)
                Else
                    smSave(5, imEvtRowNo) = ""
                End If
                If imEvtRowNo < UBound(smSave, 2) Then    'New lines set after all fields entered
                    imLefChg = True
                End If
            End If
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtTestSaveFields              *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mEvtTestSaveFields(ilRowNo As Integer) As Integer
'
'   iRet = mEvtTestSaveFields(ilRowNo)
'   Where:
'       ilRowNo (I)- row number to be checked
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim slEvtLen As String
    Dim slLibLen As String
    Dim slLen As String
    If smSave(1, ilRowNo) = "" Then
        ilRes = MsgBox("Time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imEvtBoxNo = TIMEINDEX
        mEvtTestSaveFields = NO
        Exit Function
    Else
        If Not gValidLength(smSave(1, ilRowNo)) Then
            ilRes = MsgBox("Time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imEvtBoxNo = TIMEINDEX
            mEvtTestSaveFields = NO
            Exit Function
        End If
    End If
    'Test that time is not larger then lib time
    slEvtLen = gFormatLength(smSave(1, ilRowNo), "3", False)
    slLibLen = gFormatLength(smSpecSave(3), "3", False)
    If gLengthToCurrency(slEvtLen) >= gLengthToCurrency(slLibLen) Then
        ilRes = MsgBox("Time exceeds Library time", vbOKOnly + vbExclamation, "Incomplete")
        imEvtBoxNo = TIMEINDEX
        mEvtTestSaveFields = NO
        Exit Function
    End If
    If smSave(2, ilRowNo) = "" Then
        ilRes = MsgBox("Event Type must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imEvtBoxNo = EVTTYPEINDEX
        mEvtTestSaveFields = NO
        Exit Function
    End If
    If (((Asc(smSave(9, ilRowNo)) >= Asc("1")) And (Asc(smSave(9, ilRowNo)) <= Asc("9"))) Or (smSave(9, ilRowNo) = "Y")) Then
        If smSave(3, ilRowNo) = "" Then
            'ilRes = MsgBox("Event Name must be specified", vbOkOnly + vbExclamation, "Incomplete")
            'imEvtBoxNo = EVTNAMEINDEX
            'mEvtTestSaveFields = NO
            'Exit Function
        End If
    End If
    Select Case smSave(9, ilRowNo)
        Case "1"  'Program
            If smSave(7, ilRowNo) = "" Then
                ilRes = MsgBox("Length must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = LENGTHINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            If Not gValidLength(smSave(7, ilRowNo)) Then
                ilRes = MsgBox("Length must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = LENGTHINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            slLen = gFormatLength(smSave(7, ilRowNo), "3", False)
            If gLengthToCurrency(slEvtLen) + gLengthToCurrency(slLen) > gLengthToCurrency(slLibLen) Then
                ilRes = MsgBox("End Time of Event exceeds Library time", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = LENGTHINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
        Case "2"  'Contract Avail
            
            If smSave(4, ilRowNo) = "" Then
                ilRes = MsgBox("Avail Name must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = AVAILINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            If (smSave(6, ilRowNo) = "") And ((tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M")) Then
                smSave(6, ilRowNo) = "1"
            End If
            If (smSave(6, ilRowNo) = "") Then
                ilRes = MsgBox("Units must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = UNITSINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                If smSave(7, ilRowNo) = "" Then
                    ilRes = MsgBox("Length must be specified", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
                If Not gValidLength(smSave(7, ilRowNo)) Then
                    ilRes = MsgBox("Length must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
                slLen = gFormatLength(smSave(7, ilRowNo), "3", False)
                If gLengthToCurrency(slEvtLen) + gLengthToCurrency(slLen) > gLengthToCurrency(slLibLen) Then
                    ilRes = MsgBox("End Time of Event exceeds Library time", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
                If gLengthToCurrency(slLen) > 28800 Then
                    ilRes = MsgBox("Length can't exceed 8 hours", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
            End If
            '8298
            If mEventIdTest(smSave(11, ilRowNo)) = False Then
                If Len(Trim$(smSave(11, ilRowNo))) = 0 Then
                    ilRes = MsgBox("Event Id not formatted properly on line #" & ilRowNo & ": Cannot be blank", vbOKOnly + vbExclamation, "Incomplete")
                Else
                    ilRes = MsgBox("Event Id not formatted properly on line #" & ilRowNo & ": " & smSave(11, ilRowNo), vbOKOnly + vbExclamation, "Incomplete")
                End If
                imEvtBoxNo = LENGTHINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
        Case "3", "4", "5"  'Open BB/Floating/Close BB
            If smSave(4, ilRowNo) = "" Then
                ilRes = MsgBox("Avail Name must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = AVAILINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
        Case "6"  'Cmml Promo
            If smSave(4, ilRowNo) = "" Then
                ilRes = MsgBox("Avail Name must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = AVAILINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            If (smSave(6, ilRowNo) = "") And ((tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M")) Then
                smSave(6, ilRowNo) = "1"
            End If
            If (smSave(6, ilRowNo) = "") Then
                ilRes = MsgBox("Units must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = UNITSINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                If smSave(7, ilRowNo) = "" Then
                    ilRes = MsgBox("Length must be specified", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
                If Not gValidLength(smSave(7, ilRowNo)) Then
                    ilRes = MsgBox("Length must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
                slLen = gFormatLength(smSave(7, ilRowNo), "3", False)
                If gLengthToCurrency(slEvtLen) + gLengthToCurrency(slLen) > gLengthToCurrency(slLibLen) Then
                    ilRes = MsgBox("End Time of Event exceeds Library time", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
            End If
        Case "7"  'Feed avail
            If smSave(4, ilRowNo) = "" Then
                ilRes = MsgBox("Avail Name must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = AVAILINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            If (smSave(6, ilRowNo) = "") And ((tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M")) Then
                smSave(6, ilRowNo) = "1"
            End If
            If smSave(6, ilRowNo) = "" Then
                ilRes = MsgBox("Units must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = UNITSINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                If smSave(7, ilRowNo) = "" Then
                    ilRes = MsgBox("Length must be specified", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
                If Not gValidLength(smSave(7, ilRowNo)) Then
                    ilRes = MsgBox("Length must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
                slLen = gFormatLength(smSave(7, ilRowNo), "3", False)
                If gLengthToCurrency(slEvtLen) + gLengthToCurrency(slLen) > gLengthToCurrency(slLibLen) Then
                    ilRes = MsgBox("End Time of Event exceeds Library time", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
            End If
        Case "8", "9"  'PSA/Promo (Avail)
            If smSave(4, ilRowNo) = "" Then
                ilRes = MsgBox("Avail Name must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = AVAILINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            If (smSave(6, ilRowNo) = "") And ((tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M")) Then
                smSave(6, ilRowNo) = "1"
            End If
            If smSave(6, ilRowNo) = "" Then
                ilRes = MsgBox("Units must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = UNITSINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                If smSave(7, ilRowNo) = "" Then
                    ilRes = MsgBox("Length must be specified", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
                If Not gValidLength(smSave(7, ilRowNo)) Then
                    ilRes = MsgBox("Length must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
                slLen = gFormatLength(smSave(7, ilRowNo), "3", False)
                If gLengthToCurrency(slEvtLen) + gLengthToCurrency(slLen) > gLengthToCurrency(slLibLen) Then
                    ilRes = MsgBox("End Time of Event exceeds Library time", vbOKOnly + vbExclamation, "Incomplete")
                    imEvtBoxNo = LENGTHINDEX
                    mEvtTestSaveFields = NO
                    Exit Function
                End If
            End If
        Case "A", "B", "C", "D"  'Page eject, Line space 1, 2 or 3
        Case Else   'Other
            If smSave(7, ilRowNo) = "" Then
                ilRes = MsgBox("Length must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = LENGTHINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            If Not gValidLength(smSave(7, ilRowNo)) Then
                ilRes = MsgBox("Length must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = LENGTHINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
            slLen = gFormatLength(smSave(7, ilRowNo), "3", False)
            If gLengthToCurrency(slEvtLen) + gLengthToCurrency(slLen) > gLengthToCurrency(slLibLen) Then
                ilRes = MsgBox("End Time of Event exceeds Library time", vbOKOnly + vbExclamation, "Incomplete")
                imEvtBoxNo = LENGTHINDEX
                mEvtTestSaveFields = NO
                Exit Function
            End If
        End Select
    mEvtTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtTypeBranch                  *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to event  *
'*                      type and process               *
'*                      communication back from event  *
'*                      type                           *
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
Private Function mEvtTypeBranch() As Integer
'
'   ilRet = mEvtTypeBranch()
'   Where:
'       ilRet (O)- True = event type started
'                  False = Event type over or not started
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcEvtDropDown, lbcEvtType, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mEvtTypeBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(EVENTTYPESLIST)) Then
    '    imDoubleClickName = False
    '    mEvtTypeBranch = True
    '    mEvtEnableBox imEvtBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourglass  'Wait
    igETypeCallSource = CALLSOURCEPEVENT
    If edcEvtDropDown.Text = "[New]" Then
        sgETypeName = ""
    Else
        sgETypeName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "PEvent^Test\" & sgUserName & "\" & Trim$(str$(igETypeCallSource)) & "\" & sgETypeName
        Else
            slStr = "PEvent^Prod\" & sgUserName & "\" & Trim$(str$(igETypeCallSource)) & "\" & sgETypeName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "PEvent^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igETypeCallSource)) & "\" & sgETypeName
    '    Else
    '        slStr = "PEvent^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igETypeCallSource)) & "\" & sgETypeName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "EType.Exe " & slStr, 1)
    'PEvent.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    EType.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgETypeName)
    igETypeCallSource = Val(sgETypeName)
    ilParse = gParseItem(slStr, 2, "\", sgETypeName)
    'PEvent.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mEvtTypeBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igETypeCallSource = CALLDONE Then  'Done
        igETypeCallSource = CALLNONE
'        gSetMenuState True
        lbcEvtType.Clear
        smEvtTypeCodeTag = ""
        mEvtTypePop
        If imTerminate Then
            mEvtTypeBranch = False
            Exit Function
        End If
        gFindMatch sgETypeName, 1, lbcEvtType
        If gLastFound(lbcEvtType) > 0 Then
            imChgMode = True
            lbcEvtType.ListIndex = gLastFound(lbcEvtType)
            edcEvtDropDown.Text = lbcEvtType.List(lbcEvtType.ListIndex)
            imChgMode = False
            mEvtTypeBranch = False
        Else
            imChgMode = True
            lbcEvtType.ListIndex = 0
            edcEvtDropDown.Text = lbcEvtType.List(0)
            imChgMode = False
            edcEvtDropDown.SetFocus
            sgETypeName = ""
            Exit Function
        End If
        sgETypeName = ""
 '       slNameCode = lbcETypeCode.List(cbcEType.ListIndex - 1)
 '       ilRet = gParseItem(slNameCode, 2, "\", slCode)
'       On Error GoTo mETypeBranchErr
'        gCPErrorMsg ilRet, "mETypeBranch (gParseItem field 2)", EName
'        On Error GoTo 0
'        slCode = Trim$(slCode)
'        tmEtfSrchKey.iCode = Val(slCode)
'        ilRet = btrGetEqual(hmEtf, tmEtf, lmEtfRecLen, tmEtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'        On Error GoTo mETypeBranchErr
'        gBtrvErrorMsg ilRet, "mETypeBranch (btrGetEqual)", EName
'        On Error GoTo 0
    End If
    If igETypeCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igETypeCallSource = CALLNONE
        sgETypeName = ""
        mEvtEnableBox imEvtBoxNo
        Exit Function
    End If
    If igETypeCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igETypeCallSource = CALLNONE
        sgETypeName = ""
        mEvtEnableBox imEvtBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mETypePop                       *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection event   *
'*                      type box                       *
'*                                                     *
'*******************************************************
Private Sub mEvtTypePop()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcEvtType.ListIndex
    If ilIndex > 0 Then
        slName = lbcEvtType.List(ilIndex)
    End If
    'ilRet = gPopEvtNmByTypeBox(PEvent, True, True, lbcEvtType, lbcEvtTypeCode)
    ilRet = gPopEvtNmByTypeBox(PEvent, True, True, lbcEvtType, tmEvtTypeCode(), smEvtTypeCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mEvtTypePopErr
        gCPErrorMsg ilRet, "mEvtTypePop (gIMoveListBox: EvtType)", PEvent
        On Error GoTo 0
        lbcEvtType.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        mDeleteEvtNameCtrl
        For ilLoop = 1 To lbcEvtType.ListCount - 1 Step 1
            Load lbcEvtName(ilLoop) 'Create list box
            Load lbcEvtNameCode(ilLoop)
            lbcEvtName(ilLoop).Clear
            lbcEvtNameCode(ilLoop).Tag = "" 'Force population
            mEvtNamePop ilLoop, lbcEvtName(ilLoop), lbcEvtNameCode(ilLoop)
            If imTerminate Then
                Exit Sub
            End If
        Next ilLoop
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcEvtType
            If gLastFound(lbcEvtType) > 0 Then
                lbcEvtType.ListIndex = gLastFound(lbcEvtType)
            Else
                lbcEvtType.ListIndex = -1
            End If
        Else
            lbcEvtType.ListIndex = ilIndex
        End If
        lbcCEvtType.Clear
        For ilLoop = 1 To lbcEvtType.ListCount - 1 Step 1
            lbcCEvtType.AddItem lbcEvtType.List(ilLoop)
        Next ilLoop
        imChgMode = False
    End If
    Exit Sub
mEvtTypePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mExclBranch                     *
'*                                                     *
'*             Created:5/8/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      exclusion and process          *
'*                      communication back from        *
'*                      exclusion                      *
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
Private Function mExclBranch(ilIndex As Integer) As Integer
'
'   ilRet = mExclBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcEvtDropDown, lbcExcl(ilIndex), imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcEvtDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mExclBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(EXCLUSIONSLIST)) Then
    '    imDoubleClickName = False
    '    mExclBranch = True
    '    mEvtEnableBox imEvtBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourglass  'Wait
    sgMnfCallType = "X"
    igMNmCallSource = CALLSOURCEPEVENT
    If edcEvtDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "PEvent^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "PEvent^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "PEvent^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "PEvent^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'PEvent.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'PEvent.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mExclBranch = True
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
        lbcExcl(ilIndex).Clear
        sgExclCodeTag = ""
        sgExclMnfStamp = ""
        mExclPop
        If imTerminate Then
            mExclBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcExcl(ilIndex)
        If gLastFound(lbcExcl(ilIndex)) > 0 Then
            imChgMode = True
            lbcExcl(ilIndex).ListIndex = gLastFound(lbcExcl(ilIndex))
            edcEvtDropDown.Text = lbcExcl(ilIndex).List(lbcExcl(ilIndex).ListIndex)
            imChgMode = False
            mExclBranch = False
        Else
            imChgMode = True
            lbcExcl(ilIndex).ListIndex = 1
            edcEvtDropDown.Text = lbcExcl(ilIndex).List(1)
            imChgMode = False
            edcEvtDropDown.SetFocus
            sgMNmName = ""
            Exit Function
        End If
        sgMNmName = ""
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEvtEnableBox imEvtBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEvtEnableBox imEvtBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mExclPop                        *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate exclusion list        *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mExclPop()
'
'   mExclPop
'   Where:
'
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    ReDim slExcl(0 To 1) As String      'Exclusion name, saved to determine if changed
    ReDim ilExcl(0 To 1) As Integer      'Exclusion name, saved to determine if changed
    'Repopulate if required- if sales source changed by another user while in this screen
    ilFilter(0) = CHARFILTER
    slFilter(0) = "X"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilExcl(0) = lbcExcl(0).ListIndex
    ilExcl(1) = lbcExcl(1).ListIndex
    If ilExcl(0) > 1 Then
        slExcl(0) = lbcExcl(0).List(ilExcl(0))
    End If
    If ilExcl(1) > 1 Then
        slExcl(1) = lbcExcl(1).List(ilExcl(1))
    End If
    If lbcExcl(0).ListCount <> lbcExcl(1).ListCount Then
        lbcExcl(0).Clear
    End If
    'ilRet = gIMoveListBox(PEvent, lbcExcl(0), lbcExclCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(PEvent, lbcExcl(0), tgExclCode(), sgExclCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mExclPopErr
        gCPErrorMsg ilRet, "mExclPop (gIMoveListBox)", PEvent
        On Error GoTo 0
        lbcExcl(0).AddItem "[None]", 0
        lbcExcl(0).AddItem "[New]", 0  'Force as first item on list
        lbcExcl(1).Clear
        For ilLoop = lbcExcl(0).ListCount - 1 To 0 Step -1
            lbcExcl(1).AddItem lbcExcl(0).List(ilLoop), 0
        Next ilLoop
        imChgMode = True
        If ilExcl(0) > 1 Then
            gFindMatch slExcl(0), 2, lbcExcl(0)
            If gLastFound(lbcExcl(0)) > 1 Then
                lbcExcl(0).ListIndex = gLastFound(lbcExcl(0))
            Else
                lbcExcl(0).ListIndex = -1
            End If
        Else
            lbcExcl(0).ListIndex = ilExcl(0)
        End If
        If ilExcl(1) > 1 Then
            gFindMatch slExcl(1), 2, lbcExcl(1)
            If gLastFound(lbcExcl(1)) > 1 Then
                lbcExcl(1).ListIndex = gLastFound(lbcExcl(1))
            Else
                lbcExcl(1).ListIndex = -1
            End If
        Else
            lbcExcl(1).ListIndex = ilExcl(1)
        End If
        imChgMode = False
    End If
    Exit Sub
mExclPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
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
    Dim ilRet As Integer 'btrieve status
    Dim ilLoop As Integer
    Dim llLength As Long
    Dim slNameCode As String
    Dim slCode As String
    ReDim tmLef(0 To 0) As LEF              'Lef record images
    ReDim lmEvtRecPos(0 To 0) As Long
    ReDim smSave(0 To 11, 0 To 1) As String
    ReDim imSave(0 To 1, 0 To 1) As Integer
    ReDim smShow(0 To COMMENTINDEX, 0 To 1) As String
    ReDim tgPrg(0 To 0) As PRGDATE  'Time/Dates
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    imLBEvtCtrls = 1
    imLBSpecCtrls = 1
    imLBCEvtCtrls = 1
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.height = 165
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    'PEvent.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone PEvent
    'PEvent.Show
    Screen.MousePointer = vbHourglass
    'DoEvents
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    imFirstFocus = True
    imDateAssign = False
    imDragType = -1
    imPopReqd = False
    imLtfRecLen = Len(tmLtf)  'Get and save LTF record length
    imLvfRecLen = Len(tmLvf)  'Get and save LVF record length
    imLefRecLen = Len(tmLef(0))  'Get and save LEF record length
    imSpecBoxNo = -1 'Initialize current Box to N/A
    imEvtBoxNo = -1 'Initialize current Box to N/A
    imEvtRowNo = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imChgMode = False
    imTypeChgMode = False
    imBSMode = False
    imLtfChg = False
    imLvfChg = False
    imLefChg = False
    imAllAnsw = False
    imSettingValue = False
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    lm3600 = 3600
    lmXCenter = 2145
    lmYCenter = 2125
    lmBaseRadius = 1995
    imButton = 0
    imIgnoreRightMove = False
    imLastArcPainted = -1
    imFirstTimeType = True
    smEvtPrgName = "Programs"   'Used to test if event is a program
    smEvtAvName = "Contract Avails" 'Used to test if event is an avail
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    'imcCalc.Picture = IconTraf!imcCalc.Picture
    hmLtf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmLtf, "", sgDBPath & "Ltf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ltf.Btr)", PEvent
    On Error GoTo 0
    hmLvf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lvf.Btr)", PEvent
    On Error GoTo 0
    hmLef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmLef, "", sgDBPath & "Lef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lef.btr)", PEvent
    On Error GoTo 0
    hmEnf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmEnf, "", sgDBPath & "Enf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Enf.btr)", PEvent
    On Error GoTo 0
    imEnfRecLen = Len(tmEnf)
    hmCef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cef.btr)", PEvent
    On Error GoTo 0
    If imVefCode > 0 Then
        hmVef = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.btr)", PEvent
        On Error GoTo 0
        imVefRecLen = Len(tmVef)
        tmVefSrchKey.iCode = imVefCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrGetEqual: Vef.btr)", PEvent
        On Error GoTo 0
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        smVehName = Trim$(tmVef.sName)
        imVpfIndex = gVpfFind(PEvent, imVefCode)
        'Determine if Date button can be enabled
        'imAllowDate = True
        'If (tmVef.sType = "S") Or (tmVef.sType = "A") Then
        '    hmLcf = CBtrvTable()
        '    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        '    On Error GoTo mInitErr
        '    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.btr)", PEvent
        '    On Error GoTo 0
        '    imLcfRecLen = Len(tmLcf)
        '    tmLcfSrchKey.sType = "O"
        '    tmLcfSrchKey.sStatus = "P"
        '    tmLcfSrchKey.iVefCode = imVefCode
        '    tmLcfSrchKey.iLogDate(0) = 0
        '    tmLcfSrchKey.iLogDate(1) = 0
        '    tmLcfSrchKey.iSeqNo = 0
        '    ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        '    If (ilRet = BTRV_ERR_NONE) And (tmLcf.sType = "O") And (tmLcf.sStatus = "P") And (tmLcf.iVefCode = imVefCode) Then
        '        imAllowDate = True
        '    Else
        '        tmLcfSrchKey.sType = "O"
        '        tmLcfSrchKey.sStatus = "D"
        '        tmLcfSrchKey.iVefCode = imVefCode
        '        tmLcfSrchKey.iLogDate(0) = 0
        '        tmLcfSrchKey.iLogDate(1) = 0
        '        tmLcfSrchKey.iSeqNo = 0
        '        ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        '        'If (ilRet = BTRV_ERR_NONE) And (tmLcf.sType = "O") And (tmLcf.sStatus = "D") And (tmLcf.iVefCode = imVefCode) Then
        '        '    imAllowDate = True
        '        'Else
        '            'If any links define- disallow date button (this should test for the vehicle-but it takes
        '            'to long
        '            'If tmVef.sType = "S" Then
        '            '    If gCodeChrRefExist(PEvent, "Vlf.Btr", imVefCode, "VLFSELLCODE", "P", "VLFSTATUS") Then
        '            '        imAllowDate = False
        '            '        cmcDates.Enabled = False
        '            '    End If
        '            'Else
        '            '    If gCodeChrRefExist(PEvent, "Vlf.Btr", imVefCode, "VLFAIRCODE", "P", "VLFSTATUS") Then
        '            '        imAllowDate = False
        '            '        cmcDates.Enabled = False
        '            '    End If
        '            'End If
        '        'End If
        '    End If
        '    ilRet = btrClose(hmLcf)
        '    btrDestroy hmLcf
        'End If
    Else
        Screen.MousePointer = vbDefault
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    'plcScreen.Caption = "Library Events- " & smVehName
    lbcEvtType.Clear
    mEvtTypePop
    If imTerminate Then
        Exit Sub
    End If
    lbcEvtType.ListIndex = -1
    lbcEvtAvail.Clear
    mEvtAvailPop
    If imTerminate Then
        Exit Sub
    End If
    lbcExcl(0).Clear 'Force list box to be populated
    lbcExcl(1).Clear 'Force list box to be populated
    mExclPop
    If imTerminate Then
        Exit Sub
    End If
    'PEvent.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'gCenterModalForm PEvent
    'Traffic!plcHelp.Caption = ""
    'mInitBox
    cbcType.Clear
    cbcType.AddItem "Regular"
    If smLibTypeEnabled = "Y" Then
        cbcType.AddItem "Special"
        cbcType.AddItem "Sport"
        cbcType.AddItem "Std Format"
    End If
    cbcType.ListIndex = igLibType
    imTypeSelectedIndex = igLibType
'    imSpecSave(1) = igLibType
    lbcLibName.Clear
    If smVersionChecked = "Y" Then
        ckcShowVersion.Value = vbChecked
    Else
        ckcShowVersion.Value = vbUnchecked
    End If
'    mLibPop 'Populate as the setting of ckcShowVersion might not cause population
'    If imTerminate Then
'        Exit Sub
'    End If
    cbcSelect.Clear  'Force list box to be populated
    mPopulate
    If Not imTerminate Then
        If lmLibCode <= 0 Then
            cbcSelect.ListIndex = 0 'This will generate a select_change event
        Else
            If Not igPrgDupl Then
                cbcSelect.ListIndex = 0 'This will generate a select_change event
                For ilLoop = 0 To UBound(tmSelectCode) - 1 Step 1 'lbcSelectCode.ListCount - 1 Step 1
                    slNameCode = tmSelectCode(ilLoop).sKey 'lbcSelectCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Library Name
                    If Val(slCode) = lmLibCode Then
                        cbcSelect.ListIndex = ilLoop + 1 'This will generate a select_change event
                        Exit For
                    End If
                Next ilLoop
            Else
                ilRet = mReadRec(-1, True)  'Use lmLibCode instead of index
                ilRet = mReadLefRec()
                tmLtf.iVefCode = 0  'Remove vehicle
                pbcEvents.Cls
                mMoveRecToCtrl
                mMoveEvtRecToCtrl True
                mInitEvtShow
                pbcEvents_Paint


                imSelectedIndex = 0 'Act as if new
                imChgMode = True    'Don't branch into cbcSelect code as init values will be removed
                cbcSelect.ListIndex = 0 'New
                imChgMode = False
                If imSpecSave(1) = 3 Then   'If duplicating std format- change to regular
                    imSpecSave(1) = 0
                    tmSpecCtrls(SPECLIBTYPEINDEX).iChg = True
                End If
                smSpecSave(1) = ""
                smSpecSave(2) = ""
                mInitSpecShow
                pbcLibSpec.Cls
                pbcLibSpec_Paint
            End If
        End If
        mSetCommands
        'Process View Clock information
        fmPI = 3.14149265
        imViewLibType = igLibType
        mCLibPop
        If imTerminate Then
            Exit Sub
        End If
        If smSpecSave(3) <> "" Then
            llLength = CLng(gLengthToCurrency(smSpecSave(3)))
            hbcHour.Min = 1
            If (llLength Mod lm3600) = 0 Then
                hbcHour.Max = llLength \ lm3600
            Else
                hbcHour.Max = llLength \ lm3600 + 1
            End If
            hbcHour.Value = 1
        Else
            hbcHour.Min = 1
            hbcHour.Max = 1
            hbcHour.Value = 1
        End If
        rbcView(0).Value = True
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
'*      Procedure Name:mInitBox                      *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
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

    flTextHeight = pbcLibSpec.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcLibSpec.Move 180, 480    ', pbcEvents.Width + vbcEvents.Width + fgPanelAdj, pbcLibSpec.Height + fgPanelAdj
    pbcLibSpec.Move rbcView(1).Left + rbcView(1).Width + plcLibSpec.Left + fgBevelX, plcLibSpec.Top + fgBevelY
    'plcLibSpec.Move 180, 480, rbcView(1).Left + rbcView(1).Width + pbcLibSpec.Width + 2 * fgBevelX, pbcLibSpec.Height + 2 * fgBevelY ' + fgPanelAdj
    plcLibSpec.Move 180, 480
    pbcKey.Move 15, 585
    'cmcDates.Move plcLibSpec.Left + plcLibSpec.Width - cmcDates.Width - 2 * fgPanelAdj
    'Library Type
    gSetCtrl tmSpecCtrls(SPECLIBTYPEINDEX), 30, 30, 1230, fgBoxStH
    'Library Name
    gSetCtrl tmSpecCtrls(SPECLIBNAMEINDEX), 1275, tmSpecCtrls(SPECLIBTYPEINDEX).fBoxY, 2175, fgBoxStH
    'Variation- show field only
    gSetCtrl tmSpecCtrls(SPECVARINDEX), 3465, tmSpecCtrls(SPECLIBTYPEINDEX).fBoxY, 660, fgBoxStH
    'Length
    gSetCtrl tmSpecCtrls(SPECLENGTHINDEX), 4140, tmSpecCtrls(SPECLIBTYPEINDEX).fBoxY, 1035, fgBoxStH
    'Base Time
    gSetCtrl tmSpecCtrls(SPECBASETIMEINDEX), 5190, tmSpecCtrls(SPECLIBTYPEINDEX).fBoxY, 1035, fgBoxStH
    tmSpecCtrls(SPECBASETIMEINDEX).iReq = False
    'Start Date
'    gSetCtrl tmSpecCtrls(SPECSTARTDATEINDEX), 3270, tmSpecCtrls(SPECLIBNAMEINDEX).fBoxY, 945, fgBoxSth
'    tmSpecCtrls(SPECSTARTDATEINDEX).iReq = False
    'End Date
'    gSetCtrl tmSpecCtrls(SPECENDDATEINDEX), 4230, tmSpecCtrls(SPECLIBNAMEINDEX).fBoxY, 945, fgBoxSth
'    tmSpecCtrls(SPECSTARTDATEINDEX).iReq = False
    'Start Time
'    gSetCtrl tmSpecCtrls(SPECSTARTTIMEINDEX), 5190, tmSpecCtrls(SPECLIBNAMEINDEX).fBoxY, 1215, fgBoxSth
'    tmSpecCtrls(SPECSTARTTIMEINDEX).iReq = False
    'Days of the week
'    For ilLoop = 0 To 6 Step 1
'        gSetCtrl tmSpecCtrls(SPECDAYSINDEX + ilLoop), 6420 + 255 * (ilLoop), tmSpecCtrls(SPECLIBNAMEINDEX).fBoxY, 240, fgBoxSth
'        tmSpecCtrls(SPECDAYSINDEX + ilLoop).iReq = False
'    Next ilLoop
    'Event area
    plcEvents.Move 180, 1035, pbcEvents.Width + vbcEvents.Width + fgPanelAdj, pbcEvents.height + fgPanelAdj
    pbcEvents.Move plcEvents.Left + fgBevelX, plcEvents.Top + fgBevelY
    vbcEvents.Move pbcEvents.Left + pbcEvents.Width + 15, pbcEvents.Top
    pbcArrow.Move plcEvents.Left - pbcArrow.Width - 15    'set arrow    'Vehicle
    pbcClock.Move pbcEvents.Left, pbcEvents.Top
    'Time
    gSetCtrl tmEvtCtrls(TIMEINDEX), 30, 375, 780, fgBoxGridH
    'Event Type
    gSetCtrl tmEvtCtrls(EVTTYPEINDEX), 825, tmEvtCtrls(TIMEINDEX).fBoxY, 1350, fgBoxGridH
    'Event Name
    gSetCtrl tmEvtCtrls(EVTNAMEINDEX), 2190, tmEvtCtrls(TIMEINDEX).fBoxY, 1350, fgBoxGridH
    tmEvtCtrls(EVTNAMEINDEX).iReq = False
    'Avails
    gSetCtrl tmEvtCtrls(AVAILINDEX), 3555, tmEvtCtrls(TIMEINDEX).fBoxY, 1380, fgBoxGridH
    tmEvtCtrls(AVAILINDEX).iReq = False
    'Units
    gSetCtrl tmEvtCtrls(UNITSINDEX), 4950, tmEvtCtrls(TIMEINDEX).fBoxY, 390, fgBoxGridH
    tmEvtCtrls(UNITSINDEX).iReq = False
    'Length
    gSetCtrl tmEvtCtrls(LENGTHINDEX), 5355, tmEvtCtrls(TIMEINDEX).fBoxY, 780, fgBoxGridH
    tmEvtCtrls(LENGTHINDEX).iReq = False
    'True Time
    gSetCtrl tmEvtCtrls(TRUETIMEINDEX), 6150, tmEvtCtrls(TIMEINDEX).fBoxY, 375, fgBoxGridH
    tmEvtCtrls(TRUETIMEINDEX).iReq = False
    'Event ID
    gSetCtrl tmEvtCtrls(EVTIDINDEX), 6540, tmEvtCtrls(TIMEINDEX).fBoxY, 1065, fgBoxGridH
    tmEvtCtrls(EVTIDINDEX).iReq = False
    'Comment
    gSetCtrl tmEvtCtrls(COMMENTINDEX), 7620, tmEvtCtrls(TIMEINDEX).fBoxY, 1080, fgBoxGridH
    tmEvtCtrls(COMMENTINDEX).iReq = False
    'Exclusion- shown in avail area
    gSetCtrl tmEvtCtrls(EXCL1INDEX), 9000, tmEvtCtrls(TIMEINDEX).fBoxY, 1380, fgBoxGridH
    tmEvtCtrls(EXCL1INDEX).iReq = False
    gSetCtrl tmEvtCtrls(EXCL2INDEX), 9000, tmEvtCtrls(TIMEINDEX).fBoxY, 1380, fgBoxGridH
    tmEvtCtrls(EXCL2INDEX).iReq = False
    'Clock events
    'Time
    gSetCtrl tmCEvtCtrls(TIMEINDEX), 30, 30, 1890, fgBoxStH
    'Event Type
    gSetCtrl tmCEvtCtrls(EVTTYPEINDEX), tmCEvtCtrls(TIMEINDEX).fBoxX, tmCEvtCtrls(TIMEINDEX).fBoxY + fgStDeltaY, 1890, fgBoxStH
    'Event Name
    gSetCtrl tmCEvtCtrls(EVTNAMEINDEX), tmCEvtCtrls(TIMEINDEX).fBoxX, tmCEvtCtrls(EVTTYPEINDEX).fBoxY + fgStDeltaY, 1890, fgBoxStH
    'Avails
    gSetCtrl tmCEvtCtrls(AVAILINDEX), tmCEvtCtrls(TIMEINDEX).fBoxX, tmCEvtCtrls(EVTNAMEINDEX).fBoxY + fgStDeltaY, 1890, fgBoxStH
    tmCEvtCtrls(AVAILINDEX).iReq = False
    'Units
    gSetCtrl tmCEvtCtrls(UNITSINDEX), tmCEvtCtrls(TIMEINDEX).fBoxX, tmCEvtCtrls(AVAILINDEX).fBoxY + fgStDeltaY, 630, fgBoxStH
    tmCEvtCtrls(UNITSINDEX).iReq = False
    'Length
    gSetCtrl tmCEvtCtrls(LENGTHINDEX), 675, tmCEvtCtrls(UNITSINDEX).fBoxY, 1245, fgBoxStH
    tmCEvtCtrls(LENGTHINDEX).iReq = False
    'True Time
    gSetCtrl tmCEvtCtrls(TRUETIMEINDEX), tmCEvtCtrls(TIMEINDEX).fBoxX, tmCEvtCtrls(UNITSINDEX).fBoxY + fgStDeltaY, 630, fgBoxStH
    tmCEvtCtrls(TRUETIMEINDEX).iReq = False
    'Event ID
    gSetCtrl tmCEvtCtrls(EVTIDINDEX), 675, tmCEvtCtrls(TRUETIMEINDEX).fBoxY, 1245, fgBoxStH
    tmCEvtCtrls(EVTIDINDEX).iReq = False
    'Comment
    gSetCtrl tmCEvtCtrls(COMMENTINDEX), tmCEvtCtrls(TIMEINDEX).fBoxX, tmCEvtCtrls(TRUETIMEINDEX).fBoxY + fgStDeltaY, 1890, fgBoxStH
    tmCEvtCtrls(COMMENTINDEX).iReq = False
    'Exclusion- shown in avail area
    gSetCtrl tmCEvtCtrls(EXCL1INDEX), 9000, tmCEvtCtrls(AVAILINDEX).fBoxY, 1890, fgBoxStH
    tmCEvtCtrls(EXCL1INDEX).iReq = False
    gSetCtrl tmCEvtCtrls(EXCL2INDEX), 9000, tmCEvtCtrls(AVAILINDEX).fBoxY, 1890, fgBoxStH
    tmCEvtCtrls(EXCL2INDEX).iReq = False

    llMax = 0
    For ilLoop = imLBEvtCtrls To UBound(tmEvtCtrls) Step 1
        tmEvtCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmEvtCtrls(ilLoop).fBoxW)
        Do While (tmEvtCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmEvtCtrls(ilLoop).fBoxW = tmEvtCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmEvtCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmEvtCtrls(ilLoop).fBoxX)
            Do While (tmEvtCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmEvtCtrls(ilLoop).fBoxX = tmEvtCtrls(ilLoop).fBoxX + 1
            Loop
            If (tmEvtCtrls(ilLoop).fBoxX > 90) And (ilLoop < EXCL1INDEX) Then
                Do
                    If tmEvtCtrls(ilLoop - 1).fBoxX + tmEvtCtrls(ilLoop - 1).fBoxW + 15 < tmEvtCtrls(ilLoop).fBoxX Then
                        tmEvtCtrls(ilLoop - 1).fBoxW = tmEvtCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmEvtCtrls(ilLoop - 1).fBoxX + tmEvtCtrls(ilLoop - 1).fBoxW + 15 > tmEvtCtrls(ilLoop).fBoxX Then
                        tmEvtCtrls(ilLoop - 1).fBoxW = tmEvtCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If ilLoop < EXCL1INDEX Then
            If tmEvtCtrls(ilLoop).fBoxX + tmEvtCtrls(ilLoop).fBoxW + 15 > llMax Then
                llMax = tmEvtCtrls(ilLoop).fBoxX + tmEvtCtrls(ilLoop).fBoxW + 15
            End If
        Else
            tmEvtCtrls(ilLoop).fBoxW = tmEvtCtrls(AVAILINDEX).fBoxW
        End If
    Next ilLoop

    pbcEvents.Picture = LoadPicture("")
    pbcEvents.Width = llMax
    plcEvents.Width = llMax + vbcEvents.Width + 2 * fgBevelX + 15
    lacEvtFrame.Width = llMax - 15
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    cmcDone.Left = (PEvent.Width - 5 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
    cmcUpdate.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    cmcErase.Left = cmcUpdate.Left + cmcUpdate.Width + ilSpaceBetweenButtons
    cmcUndo.Left = cmcErase.Left + cmcErase.Width + ilSpaceBetweenButtons
    cmcDone.Top = PEvent.height - (3 * cmcDone.height) / 2
    cmcCancel.Top = cmcDone.Top
    cmcUpdate.Top = cmcDone.Top
    cmcErase.Top = cmcDone.Top
    cmcUndo.Top = cmcDone.Top
    imcTrash.Top = cmcDone.Top - imcTrash.height / 2
    imcTrash.Left = PEvent.Width - (3 * imcTrash.Width) / 2
    llAdjTop = imcTrash.Top - plcLibSpec.Top - plcLibSpec.height - 120 - tmEvtCtrls(1).fBoxY
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    Do While plcEvents.Top + llAdjTop + 2 * fgBevelY + 240 < imcTrash.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    plcEvents.height = llAdjTop + 2 * fgBevelY
    pbcEvents.Left = plcEvents.Left + fgBevelX
    pbcEvents.Top = plcEvents.Top + fgBevelY
    pbcEvents.height = plcEvents.height - 2 * fgBevelY
    vbcEvents.Left = pbcEvents.Left + pbcEvents.Width + 15
    vbcEvents.Top = pbcEvents.Top
    vbcEvents.height = pbcEvents.height
    If fmAdjFactorW >= 1.2 Then
        plcSelect.Width = CLng(1.2 * plcSelect.Width)
        Do While (plcSelect.Width Mod 15) <> 0
            plcSelect.Width = plcSelect.Width + 1
        Loop
        cbcSelect.Width = plcSelect.Width - cbcSelect.Left - fgBevelX
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitEvtShow                    *
'*                                                     *
'*             Created:9/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show values         *
'*                                                     *
'*******************************************************
Private Sub mInitEvtShow()
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    Dim slStr As String
    Dim flWidth As Single
    Dim ilSvEvtRowNo As Integer
    Dim slXMid As String

    ilSvEvtRowNo = imEvtRowNo
    For ilRowNo = LBound(tmLef) + 1 To UBound(tmLef) Step 1
        For ilBoxNo = TIMEINDEX To COMMENTINDEX Step 1
            Select Case ilBoxNo 'Branch on box type (control)
                Case TIMEINDEX 'Time index
                    slStr = smSave(1, ilRowNo)
                    If imTimeRelative Then
                        slStr = gFormatLength(slStr, "3", False)
                        gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                        smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                    Else
                        gAddTimeLength smSpecSave(4), slStr, "A", "1", slStr, slXMid
                        gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                        smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                    End If
                Case EVTTYPEINDEX 'Event Type
                    slStr = smSave(2, ilRowNo)
                    gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                Case EVTNAMEINDEX 'Event Name
                    slStr = smSave(3, ilRowNo)
                    gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                Case AVAILINDEX 'Avail
                    If smSave(9, ilRowNo) = "1" Then    'program
                        If smSave(5, ilRowNo) = "" Then
                            slStr = smSave(4, ilRowNo)
                            gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                            smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                        ElseIf smSave(4, ilRowNo) = "" Then
                            slStr = smSave(5, ilRowNo)
                            gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                            smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                        Else
                            flWidth = tmEvtCtrls(ilBoxNo).fBoxW
                            tmEvtCtrls(ilBoxNo).fBoxW = tmEvtCtrls(ilBoxNo).fBoxW / 2
                            slStr = smSave(4, ilRowNo)
                            gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                            tmEvtCtrls(ilBoxNo).fBoxW = flWidth
                            slStr = tmEvtCtrls(ilBoxNo).sShow & "/" & smSave(5, ilRowNo)
                            gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                            smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                        End If
                    Else
                        slStr = smSave(4, ilRowNo)
                        gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                        smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                    End If
                Case UNITSINDEX 'Unit index
                    slStr = smSave(6, ilRowNo)
                    gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                Case LENGTHINDEX 'Unit index
                    slStr = smSave(7, ilRowNo)
                    If smSave(9, ilRowNo) = "1" Then 'Program
                        slStr = gFormatLength(slStr, "3", False)
                    Else
                        slStr = gFormatLength(slStr, "3", True)
                    End If
                    gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                Case TRUETIMEINDEX
                    If smSave(9, ilRowNo) = "1" Then    'program
                        If imSave(1, ilRowNo) = 0 Then
                            slStr = "Yes"
                        Else
                            slStr = "No"
                        End If
                    Else
                        slStr = ""
                    End If
                    gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                Case EVTIDINDEX 'Event ID
                    slStr = smSave(11, ilRowNo)
                    gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
                Case COMMENTINDEX 'Unit index
                    slStr = smSave(8, ilRowNo)
                    If slStr = "^" Then
                        slStr = ""
                    End If
                    gSetShow pbcEvents, slStr, tmEvtCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmEvtCtrls(ilBoxNo).sShow
            End Select
        Next ilBoxNo
    Next ilRowNo
    imEvtRowNo = ilSvEvtRowNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNewLine                    *
'*                                                     *
'*             Created:8/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize line                *
'*                                                     *
'*******************************************************
Private Sub mInitNewEvent()
    smSave(1, imEvtRowNo) = ""  'Time
    smSave(2, imEvtRowNo) = ""  'Event type
    smSave(3, imEvtRowNo) = ""  'Event Name
    smSave(4, imEvtRowNo) = ""  'Avail or Exclusion
    smSave(5, imEvtRowNo) = ""  'Exclusion
    smSave(6, imEvtRowNo) = ""  'Units
    smSave(7, imEvtRowNo) = ""  'Length
    imSave(1, imEvtRowNo) = 1   'True Time (default to No)
    smSave(8, imEvtRowNo) = "^"  'Comment-the symbol ^ is used to indicate first time and no comment
    smSave(9, imEvtRowNo) = ""  'Type of event
    smSave(10, imEvtRowNo) = "" 'Rec position
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitSpecShow                   *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mInitSpecShow()
'
'   mInitSpecShow
'   Where:
'
    Dim slStr As String
    Dim ilBoxNo As Integer
    For ilBoxNo = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        Select Case ilBoxNo 'Branch on box type (control)
            Case SPECLIBTYPEINDEX
                Select Case imSpecSave(1)
                    Case 0  'Regular
                        slStr = "Regular"
                    Case 1
                        slStr = "Special"
                    Case 2
                        slStr = "Sport"
                    Case 3
                        slStr = "Std Format"
                    Case Else
                        slStr = ""
                End Select
                gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case SPECLIBNAMEINDEX
                slStr = smSpecSave(1)
                gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case SPECVARINDEX
                If Val(smSpecSave(2)) <> 0 Then
                    slStr = smSpecSave(2)
                Else
                    slStr = ""
                End If
                gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case SPECLENGTHINDEX
                slStr = smSpecSave(3)
                slStr = gFormatLength(slStr, "3", False)
                gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
            Case SPECBASETIMEINDEX
                slStr = smSpecSave(4)
                slStr = gFormatTime(slStr, "A", "1")
                gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
'            Case SPECSTARTDATEINDEX
'                slStr = smSpecSave(3)
'                slStr = gFormatDate(slStr)
'                gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
'            Case SPECENDDATEINDEX
'                slStr = smSpecSave(4)
'                slStr = gFormatDate(slStr)
'                gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
'            Case SPECSTARTTIMEINDEX
'                slStr = smSpecSave(5)
'                slStr = gFormatTime(slStr, "A", "1")
'                gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
'            Case SPECDAYSINDEX To SPECDAYSINDEX + 6
'                If smSpecSave(ilBoxNo) = "Y" Then
'                    slStr = smSpecSave(3)
'                Else
'                    slStr = "  "
'                End If
'                gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
        End Select
    Next ilBoxNo
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLibPop                         *
'*                                                     *
'*             Created:10/11/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection library *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mLibPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slType As String
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcLibName.ListIndex
    If ilIndex > 0 Then
        slName = lbcLibName.List(ilIndex)
    End If
    Screen.MousePointer = vbHourglass  'Wait
    If imSpecSave(1) = 3 Then 'Std Format
        slType = "F"
    ElseIf imSpecSave(1) = 2 Then 'Sports
        slType = "P"
    ElseIf imSpecSave(1) = 1 Then 'Special
        slType = "S"
    Else    'Regular
        slType = "R"
    End If
    'ilRet = gPopProgLibBox(PEvent, LATESTLIB, slType, imVefCode, lbcLibName, lbcLibNameCode)
    ilRet = gPopProgLibBox(PEvent, LATESTLIB, slType, imVefCode, lbcLibName, tmLibNameCode(), smLibNameCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLibPopErr
        gCPErrorMsg ilRet, "mLibPope (gPopProgLibBox)", PEvent
        On Error GoTo 0
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcLibName
            If gLastFound(lbcLibName) > 0 Then
                lbcLibName.ListIndex = gLastFound(lbcLibName)
            Else
                lbcLibName.ListIndex = -1
            End If
        Else
            lbcLibName.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mLibPopErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilTestChg As Integer)
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    'Vehicle can't be changed
    If tmLtf.iVefCode = 0 Then  'This field set in mClearCtrlFields
        tmLtf.iVefCode = imVefCode
    End If
    If Not ilTestChg Or tmSpecCtrls(SPECLIBTYPEINDEX).iChg Then
        Select Case imSpecSave(1)
            Case 0  'Regular
                tmLtf.sType = "R"
            Case 1  'Special
                tmLtf.sType = "S"
            Case 2  'Sports
                tmLtf.sType = "P"
            Case 3  'Std Format
                tmLtf.sType = "F"
        End Select
    End If
    If Not ilTestChg Or tmSpecCtrls(SPECLIBNAMEINDEX).iChg Or (Trim$(tmLtf.sName) = "") Then
        tmLtf.sName = smSpecSave(1)
    End If
    tmLtf.iVar = Val(smSpecSave(2))
    If Not ilTestChg Or tmSpecCtrls(SPECLENGTHINDEX).iChg Then
        gPackLength smSpecSave(3), tmLvf.iLen(0), tmLvf.iLen(1)
    End If
    If Not ilTestChg Or tmSpecCtrls(SPECBASETIMEINDEX).iChg Then
        gPackTime smSpecSave(4), tmLvf.iBaseTime(0), tmLvf.iBaseTime(1)
    End If
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveEvt                        *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move events from a row to      *
'*                      another row                    *
'*                                                     *
'*******************************************************
Private Sub mMoveEvt(ilInsertAfter As Integer, ilInsertAtRowNo As Integer, ilOrigRowNo As Integer)
'
'   mMoveEvt ilInsertAfter, ilInsertAtRowNo, ilOrigRowNo
'   Where:
'       ilInsertAfter (I)- True = insert row after row; False= insert prior to row
'       ilInsertAtRowNo (I)- row number that event is to be moved into
'       ilOrigRowNo (I)- Original row number of the event to be moved
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    ReDim slMoveSave(0 To UBound(smSave, 1)) As String
    ReDim ilMoveSave(0 To UBound(imSave, 1)) As Integer
    ReDim slMoveShow(0 To COMMENTINDEX) As String
    For ilLoop = 1 To UBound(smSave, 1) Step 1
        slMoveSave(ilLoop) = smSave(ilLoop, ilOrigRowNo)
    Next ilLoop
    For ilLoop = 1 To UBound(imSave, 1) Step 1
        ilMoveSave(ilLoop) = imSave(ilLoop, ilOrigRowNo)
    Next ilLoop
    For ilLoop = 1 To UBound(smShow, 1) Step 1
        slMoveShow(ilLoop) = smShow(ilLoop, ilOrigRowNo)
    Next ilLoop
    If ilInsertAtRowNo > ilOrigRowNo Then
        If Not ilInsertAfter Then
            ilInsertAtRowNo = ilInsertAtRowNo - 1
        End If
        For ilLoop = ilOrigRowNo To ilInsertAtRowNo - 1 Step 1
            For ilIndex = 1 To UBound(smSave, 1) Step 1
                smSave(ilIndex, ilLoop) = smSave(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(imSave, 1) Step 1
                imSave(ilIndex, ilLoop) = imSave(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(smShow, 1) Step 1
                smShow(ilIndex, ilLoop) = smShow(ilIndex, ilLoop + 1)
            Next ilIndex
        Next ilLoop
    Else
        If ilInsertAfter Then
            ilInsertAtRowNo = ilInsertAtRowNo + 1
        End If
        For ilLoop = ilOrigRowNo To ilInsertAtRowNo + 1 Step -1
            For ilIndex = 1 To UBound(smSave, 1) Step 1
                smSave(ilIndex, ilLoop) = smSave(ilIndex, ilLoop - 1)
            Next ilIndex
            For ilIndex = 1 To UBound(imSave, 1) Step 1
                imSave(ilIndex, ilLoop) = imSave(ilIndex, ilLoop - 1)
            Next ilIndex
            For ilIndex = 1 To UBound(smShow, 1) Step 1
                smShow(ilIndex, ilLoop) = smShow(ilIndex, ilLoop - 1)
            Next ilIndex
        Next ilLoop
    End If
    For ilLoop = 1 To UBound(smSave, 1) Step 1
        smSave(ilLoop, ilInsertAtRowNo) = slMoveSave(ilLoop)
    Next ilLoop
    For ilLoop = 1 To UBound(imSave, 1) Step 1
        imSave(ilLoop, ilInsertAtRowNo) = ilMoveSave(ilLoop)
    Next ilLoop
    For ilLoop = 1 To UBound(smShow, 1) Step 1
        smShow(ilLoop, ilInsertAtRowNo) = slMoveShow(ilLoop)
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveEvtCtrlToRec               *
'*                                                     *
'*             Created:9/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move controls values to record *
'*                                                     *
'*******************************************************
Private Sub mMoveEvtCtrlToRec()
'
'   mMoveEvtCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slCode As String
    Dim ilEvtNameIndex As Integer
    ReDim tmLef(0 To UBound(smSave, 2) - 1) As LEF
    For ilLoop = 0 To UBound(smSave, 2) - 2 Step 1
        gPackLength smSave(1, ilLoop + 1), tmLef(ilLoop).iStartTime(0), tmLef(ilLoop).iStartTime(1)
        If ilLoop > 0 Then
            If (tmLef(ilLoop - 1).iStartTime(0) = tmLef(ilLoop).iStartTime(0)) And (tmLef(ilLoop - 1).iStartTime(1) = tmLef(ilLoop).iStartTime(1)) Then
                tmLef(ilLoop).iSeqNo = tmLef(ilLoop - 1).iSeqNo + 1
            Else
                tmLef(ilLoop).iSeqNo = 1
            End If
        Else
            tmLef(ilLoop).iSeqNo = 1
        End If
        ilEvtNameIndex = -1
        If smSave(2, ilLoop + 1) <> "" Then
            gFindMatch smSave(2, ilLoop + 1), 1, lbcEvtType
            If gLastFound(lbcEvtType) > 0 Then
                ilEvtNameIndex = gLastFound(lbcEvtType)
                slNameCode = tmEvtTypeCode(gLastFound(lbcEvtType) - 1).sKey  'lbcEvtTypeCode.List(gLastFound(lbcEvtType) - 1)
                ilRet = gParseItem(slNameCode, 3, "\", slCode)
                On Error GoTo mMoveEvtCtrlToRecErr
                gCPErrorMsg ilRet, "mMoveEvtCtrlToRec (gParseItem field 3)", PEvent
                On Error GoTo 0
                slCode = Trim$(slCode)
                tmLef(ilLoop).iEtfCode = CInt(slCode)
            Else
                tmLef(ilLoop).iEtfCode = 0
            End If
        Else
            tmLef(ilLoop).iEtfCode = 0
        End If
        tmLef(ilLoop).iMaxUnits = 0
        tmLef(ilLoop).iLen(0) = 0
        tmLef(ilLoop).iLen(1) = 0
        tmLef(ilLoop).sTrue = " "
        If (((Asc(smSave(9, ilLoop + 1)) >= Asc("1")) And (Asc(smSave(9, ilLoop + 1)) <= Asc("9"))) Or (smSave(9, ilLoop + 1) = "Y")) And (ilEvtNameIndex > 0) And (smSave(3, ilLoop + 1) <> "") Then
            gFindMatch smSave(3, ilLoop + 1), 2, lbcEvtName(ilEvtNameIndex)
            If gLastFound(lbcEvtName(ilEvtNameIndex)) > 1 Then
                slNameCode = lbcEvtNameCode(ilEvtNameIndex).List(gLastFound(lbcEvtName(ilEvtNameIndex)) - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveEvtCtrlToRecErr
                gCPErrorMsg ilRet, "mMoveEvtCtrlToRec (gParseItem field 2)", PEvent
                On Error GoTo 0
                slCode = Trim$(slCode)
                tmLef(ilLoop).iEnfCode = CInt(slCode)
            Else
                tmLef(ilLoop).iEnfCode = 0
            End If
        Else
            tmLef(ilLoop).iEnfCode = 0
        End If
        'Exclusion
        If Asc(smSave(9, ilLoop + 1)) = Asc("1") Then
            If smSave(4, ilLoop + 1) <> "" Then
                gFindMatch smSave(4, ilLoop + 1), 2, lbcExcl(0)
                If gLastFound(lbcExcl(0)) >= 2 Then
                    slNameCode = tgExclCode(gLastFound(lbcExcl(0)) - 2).sKey    'lbcEvtAvailCode.List(gLastFound(lbcEvtAvail) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    On Error GoTo mMoveEvtCtrlToRecErr
                    gCPErrorMsg ilRet, "mMoveEvtCtrlToRec (gParseItem field 2)", PEvent
                    On Error GoTo 0
                    slCode = Trim$(slCode)
                    tmLef(ilLoop).iMnfExcl(0) = CInt(slCode)
                Else
                    tmLef(ilLoop).iMnfExcl(0) = 0
                End If
            Else
                tmLef(ilLoop).iMnfExcl(0) = 0
            End If
            If smSave(5, ilLoop + 1) <> "" Then
                gFindMatch smSave(5, ilLoop + 1), 2, lbcExcl(1)
                If gLastFound(lbcExcl(1)) >= 2 Then
                    slNameCode = tgExclCode(gLastFound(lbcExcl(1)) - 2).sKey    'lbcEvtAvailCode.List(gLastFound(lbcEvtAvail) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    On Error GoTo mMoveEvtCtrlToRecErr
                    gCPErrorMsg ilRet, "mMoveEvtCtrlToRec (gParseItem field 2)", PEvent
                    On Error GoTo 0
                    slCode = Trim$(slCode)
                    tmLef(ilLoop).iMnfExcl(1) = CInt(slCode)
                Else
                    tmLef(ilLoop).iMnfExcl(1) = 0
                End If
            Else
                tmLef(ilLoop).iMnfExcl(1) = 0
            End If
        Else
            tmLef(ilLoop).iMnfExcl(0) = 0
            tmLef(ilLoop).iMnfExcl(1) = 0
        End If
        'Avail
        If ((Asc(smSave(9, ilLoop + 1)) >= Asc("2")) And (Asc(smSave(9, ilLoop + 1)) <= Asc("9"))) Or (smSave(9, ilLoop + 1) = "A") Then
            If smSave(4, ilLoop + 1) <> "" Then
                gFindMatch smSave(4, ilLoop + 1), 1, lbcEvtAvail
                If gLastFound(lbcEvtAvail) > 0 Then
                    slNameCode = tmEvtAvailCode(gLastFound(lbcEvtAvail) - 1).sKey    'lbcEvtAvailCode.List(gLastFound(lbcEvtAvail) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    On Error GoTo mMoveEvtCtrlToRecErr
                    gCPErrorMsg ilRet, "mMoveEvtCtrlToRec (gParseItem field 2)", PEvent
                    On Error GoTo 0
                    slCode = Trim$(slCode)
                    tmLef(ilLoop).ianfCode = CInt(slCode)
                Else
                    tmLef(ilLoop).ianfCode = 0
                End If
            Else
                tmLef(ilLoop).ianfCode = 0
            End If
        Else
            tmLef(ilLoop).ianfCode = 0
        End If
        Select Case smSave(9, ilLoop + 1)
            Case "1"  'Program
                gPackLength smSave(7, ilLoop + 1), tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1)
                If imSave(1, ilLoop + 1) = 0 Then
                    tmLef(ilLoop).sTrue = "Y"
                Else
                    tmLef(ilLoop).sTrue = "N"
                End If
            'Case "2", "3", "4", "5"  'Contract Avail
            Case "2"  'Contract Avail
                tmLef(ilLoop).iMaxUnits = Val(smSave(6, ilLoop + 1))
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                    gPackLength smSave(7, ilLoop + 1), tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1)
                Else
                    gPackLength "", tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1)
                End If
            Case "3", "4", "5"  'Open BB/Close BB
                tmLef(ilLoop).iMaxUnits = 0
                gPackLength "", tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1)
            Case "6"  'Cmml Promo
                tmLef(ilLoop).iMaxUnits = Val(smSave(6, ilLoop + 1))
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                    gPackLength smSave(7, ilLoop + 1), tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1)
                Else
                    gPackLength "", tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1)
                End If
            Case "7"  'Feed avail
                tmLef(ilLoop).iMaxUnits = Val(smSave(6, ilLoop + 1))
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                    gPackLength smSave(7, ilLoop + 1), tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1)
                Else
                    gPackLength "", tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1)
                End If
            Case "8", "9"  'PSA/Promo (Avail)
                tmLef(ilLoop).iMaxUnits = Val(smSave(6, ilLoop + 1))
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                    gPackLength smSave(7, ilLoop + 1), tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1)
                Else
                    gPackLength "", tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1)
                End If
            Case "A", "B", "C", "D"  'Page eject, Line space 1, 2 or 3
            Case Else   'Other
                gPackLength smSave(7, ilLoop + 1), tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1)
        End Select
    Next ilLoop
    Exit Sub
mMoveEvtCtrlToRecErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveEvtRecToCtrl               *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveEvtRecToCtrl(ilCreateEvtName As Integer)
'
'   mMoveEvtRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilEvtNameIndex As Integer
    For ilLoop = 0 To UBound(tmLef) - 1 Step 1
        smSave(10, ilLoop + 1) = Trim$(str$(lmEvtRecPos(ilLoop)))  'Save record position so copy can be retained
        gUnpackLength tmLef(ilLoop).iStartTime(0), tmLef(ilLoop).iStartTime(1), "3", False, smSave(1, ilLoop + 1)
        smSave(2, ilLoop + 1) = ""  'Event type
        smSave(9, ilLoop + 1) = ""  'Type of event
        ilEvtNameIndex = -1
        slRecCode = Trim$(str$(tmLef(ilLoop).iEtfCode))
        For ilTest = 0 To UBound(tmEvtTypeCode) - 1 Step 1  'lbcEvtTypeCode.ListCount - 1 Step 1
            slNameCode = tmEvtTypeCode(ilTest).sKey    'lbcEvtTypeCode.List(ilTest)
            ilRet = gParseItem(slNameCode, 3, "\", slCode)
            On Error GoTo mMoveEvtRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveEvtRecToCtrl (gParseItem field 3: Event Type)", PEvent
            On Error GoTo 0
            If slRecCode = slCode Then
                ilEvtNameIndex = ilTest + 1
                ilRet = gParseItem(slNameCode, 2, "\", smSave(2, ilLoop + 1))
                On Error GoTo mMoveEvtRecToCtrlErr
                gCPErrorMsg ilRet, "mMoveEvtRecToCtrl (gParseItem field 2: Event Type)", PEvent
                On Error GoTo 0
                ilRet = gParseItem(slNameCode, 1, "\", smSave(9, ilLoop + 1))
                On Error GoTo mMoveEvtRecToCtrlErr
                gCPErrorMsg ilRet, "mMoveEvtRecToCtrl (gParseItem field 1: Event Type)", PEvent
                On Error GoTo 0
                Exit For
            End If
        Next ilTest
        'Event Name
        smSave(3, ilLoop + 1) = ""
        If (((Asc(smSave(9, ilLoop + 1)) >= Asc("1")) And (Asc(smSave(9, ilLoop + 1)) <= Asc("9"))) Or (smSave(9, ilLoop + 1) = "Y")) And (ilEvtNameIndex >= 0) Then
            If tmLef(ilLoop).iEnfCode <> 0 Then
                slRecCode = Trim$(str$(tmLef(ilLoop).iEnfCode))
                For ilTest = 0 To lbcEvtNameCode(ilEvtNameIndex).ListCount - 1 Step 1
                    slNameCode = lbcEvtNameCode(ilEvtNameIndex).List(ilTest)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    On Error GoTo mMoveEvtRecToCtrlErr
                    gCPErrorMsg ilRet, "mMoveEvtRecToCtrl (gParseItem field 2: Event Name)", PEvent
                    On Error GoTo 0
                    If slRecCode = slCode Then
                        ilRet = gParseItem(slNameCode, 1, "\", smSave(3, ilLoop + 1))
                        On Error GoTo mMoveEvtRecToCtrlErr
                        gCPErrorMsg ilRet, "mMoveEvtRecToCtrl (gParseItem field 1: Event Name)", PEvent
                        On Error GoTo 0
                        Exit For
                    End If
                Next ilTest
                If (smSave(3, ilLoop + 1) = "") And (smSave(9, ilLoop + 1) <> "1") And (ilCreateEvtName) Then
                    'Obtain Event name- test if same name exist
                    tmEnfSrchKey.iCode = tmLef(ilLoop).iEnfCode
                    ilRet = btrGetEqual(hmEnf, tmEnf, imEnfRecLen, tmEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        For ilTest = 0 To lbcEvtNameCode(ilEvtNameIndex).ListCount - 1 Step 1
                            slNameCode = lbcEvtNameCode(ilEvtNameIndex).List(ilTest)
                            ilRet = gParseItem(slNameCode, 1, "\", slName)
                            If StrComp(Trim$(slName), Trim$(tmEnf.sName), 1) = 0 Then
                                smSave(3, ilLoop + 1) = slName
                                Exit For
                            End If
                        Next ilTest
                        If smSave(3, ilLoop + 1) = "" Then
                            'If name not found- add to enf
                            'First check if comment must be inserted
                            ilRet = mReadCefRec(tmEnf.lCefCode, SETFORWRITE)
                            imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
                            'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
                            If gStripChr0(tmCef.sComment) <> "" Then
                                tmCef.lCode = 0 'Autoincrement
                                ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                                If ilRet <> BTRV_ERR_NONE Then
                                    tmCef.lCode = 0
                                End If
                            Else
                                tmCef.lCode = 0
                            End If
                            tmEnf.lCefCode = tmCef.lCode
                            tmEnf.iCode = 0
                            tmEnf.iVefCode = imVefCode
                            ilRet = btrInsert(hmEnf, tmEnf, imEnfRecLen, INDEXKEY0)
                            lbcEvtName(ilEvtNameIndex).Clear
                            lbcEvtNameCode(ilEvtNameIndex).Clear
                            lbcEvtNameCode(ilEvtNameIndex).Tag = ""
                            mEvtNamePop ilEvtNameIndex, lbcEvtName(ilEvtNameIndex), lbcEvtNameCode(ilEvtNameIndex)
                            smSave(3, ilLoop + 1) = Trim$(tmEnf.sName)
                        End If
                    End If
                End If
            End If
        End If
        'Avail Name or Exclusions
        smSave(4, ilLoop + 1) = ""
        smSave(5, ilLoop + 1) = ""
        If smSave(9, ilLoop + 1) <> "1" Then    'Not program
            If ((Asc(smSave(9, ilLoop + 1)) >= Asc("2")) And (Asc(smSave(9, ilLoop + 1)) <= Asc("9"))) Or (smSave(9, ilLoop + 1) = "A") Then
                slRecCode = Trim$(str$(tmLef(ilLoop).ianfCode))
                For ilTest = 0 To UBound(tmEvtAvailCode) - 1 Step 1 'lbcEvtAvailCode.ListCount - 1 Step 1
                    slNameCode = tmEvtAvailCode(ilTest).sKey   'lbcEvtAvailCode.List(ilTest)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    On Error GoTo mMoveEvtRecToCtrlErr
                    gCPErrorMsg ilRet, "mMoveEvtRecToCtrl (gParseItem field 2: Avail Code)", PEvent
                    On Error GoTo 0
                    If slRecCode = slCode Then
                        ilRet = gParseItem(slNameCode, 1, "\", smSave(4, ilLoop + 1))
                        On Error GoTo mMoveEvtRecToCtrlErr
                        gCPErrorMsg ilRet, "mMoveEvtRecToCtrl (gParseItem field 1: Avail Name)", PEvent
                        On Error GoTo 0
                        Exit For
                    End If
                Next ilTest
            End If
        Else
            slRecCode = Trim$(str$(tmLef(ilLoop).iMnfExcl(0)))
            For ilTest = 0 To UBound(tgExclCode) - 1 Step 1 'lbcExclCode.ListCount - 1 Step 1
                slNameCode = tgExclCode(ilTest).sKey   'lbcExclCode.List(ilTest)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveEvtRecToCtrlErr
                gCPErrorMsg ilRet, "mMoveEvtRecToCtrl (gParseItem field 2: Exclusion 1 Code)", PEvent
                On Error GoTo 0
                If slRecCode = slCode Then
                    ilRet = gParseItem(slNameCode, 1, "\", smSave(4, ilLoop + 1))
                    On Error GoTo mMoveEvtRecToCtrlErr
                    gCPErrorMsg ilRet, "mMoveEvtRecToCtrl (gParseItem field 1: Exclusion 1 Name)", PEvent
                    On Error GoTo 0
                    Exit For
                End If
            Next ilTest
            slRecCode = Trim$(str$(tmLef(ilLoop).iMnfExcl(1)))
            For ilTest = 0 To UBound(tgExclCode) - 1 Step 1 'lbcExclCode.ListCount - 1 Step 1
                slNameCode = tgExclCode(ilTest).sKey   'lbcExclCode.List(ilTest)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveEvtRecToCtrlErr
                gCPErrorMsg ilRet, "mMoveEvtRecToCtrl (gParseItem field 2: Exclusion 2 Code)", PEvent
                On Error GoTo 0
                If slRecCode = slCode Then
                    ilRet = gParseItem(slNameCode, 1, "\", smSave(5, ilLoop + 1))
                    On Error GoTo mMoveEvtRecToCtrlErr
                    gCPErrorMsg ilRet, "mMoveEvtRecToCtrl (gParseItem field 1: Exclusion 2 Name)", PEvent
                    On Error GoTo 0
                    Exit For
                End If
            Next ilTest
        End If
        Select Case smSave(9, ilLoop + 1)
            Case "1"  'Program
                gUnpackLength tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1), "3", False, smSave(7, ilLoop + 1)
                If tmLef(ilLoop).sTrue = "Y" Then
                    imSave(1, ilLoop + 1) = 0
                Else
                    imSave(1, ilLoop + 1) = 1
                End If
            'Case "2", "3", "4", "5"  'Contract Avail
            Case "2"  'Contract Avail
                smSave(6, ilLoop + 1) = Trim(str$(tmLef(ilLoop).iMaxUnits))
                gUnpackLength tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1), "3", True, smSave(7, ilLoop + 1)
            Case "3", "4", "5"  'Open BB/floating/Close BB
                smSave(6, ilLoop + 1) = ""
                smSave(7, ilLoop + 1) = ""
            Case "6"  'Cmml Promo
                smSave(6, ilLoop + 1) = Trim$(str$(tmLef(ilLoop).iMaxUnits))
                gUnpackLength tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1), "3", True, smSave(7, ilLoop + 1)
            Case "7"  'Feed avail
                smSave(6, ilLoop + 1) = Trim$(str$(tmLef(ilLoop).iMaxUnits))
                gUnpackLength tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1), "3", True, smSave(7, ilLoop + 1)
            Case "8", "9"  'PSA/Promo (Avail)
                smSave(6, ilLoop + 1) = Trim$(str$(tmLef(ilLoop).iMaxUnits))
                gUnpackLength tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1), "3", True, smSave(7, ilLoop + 1)
            Case "A", "B", "C", "D"  'Page eject, Line space 1, 2 or 3
            Case Else   'Other
                gUnpackLength tmLef(ilLoop).iLen(0), tmLef(ilLoop).iLen(1), "3", True, smSave(7, ilLoop + 1)
        End Select
        If mReadCefRec(tmLef(ilLoop).lEvtIDCefCode, SETFORREADONLY) Then
            'If tmCef.iStrLen > 0 Then
            '    smSave(11, ilLoop + 1) = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                smSave(11, ilLoop + 1) = gStripChr0(tmCef.sComment)
            'Else
            '    smSave(11, ilLoop + 1) = ""
            'End If
        Else
            smSave(11, ilLoop + 1) = ""
        End If
        If mReadCefRec(tmLef(ilLoop).lCefCode, SETFORREADONLY) Then
            'If tmCef.iStrLen > 0 Then
            '    smSave(8, ilLoop + 1) = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                smSave(8, ilLoop + 1) = gStripChr0(tmCef.sComment)
            'Else
            '    smSave(8, ilLoop + 1) = ""
            'End If
        Else
            smSave(8, ilLoop + 1) = ""
        End If
    Next ilLoop
    Exit Sub
mMoveEvtRecToCtrlErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
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
    Select Case tmLtf.sType
        Case "R"    'Regular
            imSpecSave(1) = 0
        Case "S"    'Special
            imSpecSave(1) = 1
        Case "P"    'Sports
            imSpecSave(1) = 2
        Case "F"    'Std format
            imSpecSave(1) = 3
    End Select
    smSpecSave(1) = Trim$(tmLtf.sName)
    smSpecSave(2) = Trim$(str$(tmLtf.iVar))
    gUnpackLength tmLvf.iLen(0), tmLvf.iLen(1), "3", False, smSpecSave(3)
    gUnpackTime tmLvf.iBaseTime(0), tmLvf.iBaseTime(1), "A", "1", smSpecSave(4)
    If Len(smSpecSave(4)) = 0 Then
        imTimeRelative = True
    Else
        imTimeRelative = False
    End If
    For ilLoop = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        tmSpecCtrls(ilLoop).iChg = False
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
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName()
    Dim slStr As String
    Dim slNameVer As String
    Dim slName As String
    Dim ilRet As Integer
    Dim ilLastFound As Integer
    Dim slFndName As String
    Dim ilLoop As Integer
    If smSpecSave(1) <> "" Then    'Test name
        If Val(smSpecSave(2)) <> 0 Then
            slFndName = smSpecSave(1) & "-" & smSpecSave(2)  'Determine if name exist
        Else
            slFndName = smSpecSave(1)   'Determine if name exist
        End If
        slFndName = Trim$(slFndName)
        If ckcShowVersion.Value = vbChecked Then
            ilLastFound = -1
            For ilLoop = 0 To cbcSelect.ListCount - 1 Step 1
                slNameVer = cbcSelect.List(ilLoop)
                ilRet = gParseItem(slNameVer, 1, "/", slName)
                If StrComp(slFndName, slName, 1) = 0 Then
                    ilLastFound = ilLoop
                    Exit For
                End If
            Next ilLoop
        Else
            ilLastFound = -1
            gFindMatch slFndName, 0, cbcSelect  'Determine if name exist
            If gLastFound(cbcSelect) >= 0 Then
                ilLastFound = gLastFound(cbcSelect)
            End If
        End If
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If Val(smSpecSave(2)) <> 0 Then
                    slStr = smSpecSave(1) & "-" & smSpecSave(2)
                Else
                    slStr = smSpecSave(1)
                End If
                If ckcShowVersion.Value = vbChecked Then
                    slNameVer = cbcSelect.List(gLastFound(cbcSelect))
                    ilRet = gParseItem(slNameVer, 1, "/", slName)
                Else
                    slName = cbcSelect.List(gLastFound(cbcSelect))
                End If
                If Trim$(slStr) = Trim$(slName) Then
                    Beep
                    MsgBox "Library name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    mEvtSetShow imEvtBoxNo
                    imEvtBoxNo = -1
                    imEvtRowNo = -1
                    mSpecSetShow imSpecBoxNo
                    imSpecBoxNo = SPECLIBNAMEINDEX
                    mSpecEnableBox imSpecBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
'        Select Case imSpecSave(1)
'            Case 0  'Regular
'                slType = "R"
'            Case 1  'Special
'                slType = "S"
'            Case 2  'Sports
'                slType = "P"
'            Case 3  'Std Format
'                slType = "F"
'        End Select
'        ilFound = True
'        tmLnf1SrchKey.sName = smSpecSave(1)
'        tmLnf1SrchKey.iVar = Val(smSpecSave(2))
'        tmLnf1SrchKey.iVersion = 0
'        ilRet = btrGetGreaterOrEqual(hmLnf, tlLnf, imLnfRecLen, tmLnf1SrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'        Do
'            If (ilRet = BTRV_ERR_NONE) And (Trim$(tlLnf.sName) = Trim$(smSpecSave(1))) And (tlLnf.iVar = Val(smSpecSave(2))) Then
'                If tlLnf.iVefCode = imVefCode Then
'                    If tlLnf.sType <> slType Then
'                        Beep
'                        Select Case tlLnf.sType
'                            Case "R"  'Regular
'                                slType = "Regulars"
'                            Case "S"  'Special
'                                slType = "Specials"
'                            Case "P"  'Sports
'                                slType = "Sports"
'                            Case "F"  'Std Format
'                                slType = "Std Formats"
'                        End Select
'                        MsgBox "Library name already defined within " & slType & ", enter a different name", vbOkOnly + vbExclamation + vbApplicationModal, "Error"
'                        mEvtSetShow imEvtBoxNo
'                        imEvtBoxNo = -1
'                        imEvtRowNo = -1
'                        mSpecSetShow imSpecBoxNo
'                        imSpecBoxNo = SPECLIBNAMEINDEX
'                        mSpecEnableBox imSpecBoxNo
'                        mOKName = False
'                        Exit Function
'                    Else
'                        ilRet = btrGetNext(hmLnf, tlLnf, imLnfRecLen, BTRV_LOCK_NONE)
'                    End If
'                Else
'                    ilRet = btrGetNext(hmLnf, tlLnf, imLnfRecLen, BTRV_LOCK_NONE)
'                End If
'            Else
'                ilFound = False
'            End If
'        Loop While ilFound
    End If
    mOKName = True
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
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slHelpSystem As String
    slCommand = sgCommandStr
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
    'gInitStdAlone PEvent, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    'Program^Test^NOHELP\Guide\Alternative Services\33\Alternative\4\Y\N\0\0\0
    '    'Program^Prod^NOHELP\Guide\MON-FRI\119\KRLA-AM\33\Y\N\0\0\0
    '    smLibName = "MON-FRI" '"Alternative Services"  '"M=F" '"Afternoon" '"M-F"
    '    lmLibCode = 119 '33                      '314 '25 '468
    '    smVehName = "KRLA-AM" '"Alternative"           '"Galaxy"    '"Gold"  '"ABC Business Report"
    '    imVefCode = 33 '4                       '1   '10  '17
    '    smLibTypeEnabled = "Y"
    '    smVersionChecked = "N"
    '    igLibType = 0
    '    igViewType = 0
    '    igPrgDupl = 0
    'Else
        ilRet = gParseItem(slCommand, 3, "\", smLibName)    'Library Name
        ilRet = gParseItem(slCommand, 4, "\", slStr)    'Library Name Code
        lmLibCode = Val(slStr)
        ilRet = gParseItem(slCommand, 5, "\", smVehName)    'Vehicle Name
        ilRet = gParseItem(slCommand, 6, "\", slStr)    'Library Name Code
        imVefCode = Val(slStr)
        ilRet = gParseItem(slCommand, 7, "\", smLibTypeEnabled)    'Library Name
        ilRet = gParseItem(slCommand, 8, "\", smVersionChecked)    'Library Name
        ilRet = gParseItem(slCommand, 9, "\", slStr)    'Library Name
        igLibType = Val(slStr)
        ilRet = gParseItem(slCommand, 10, "\", slStr)    'Library Name
        igViewType = Val(slStr)
        ilRet = gParseItem(slCommand, 11, "\", slStr)    'Library Name
        igPrgDupl = Val(slStr)
    'End If
    '8298    '10933
       ' bmTestEventID = mTestEventIDVehicle(imVefCode)
        mTestEventIDVehicle imVefCode
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection library *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slType As String
    Dim ilVer As Integer

    imPopReqd = False
    Screen.MousePointer = vbHourglass  'Wait
    Select Case imTypeSelectedIndex
        Case 0  'Regular
            slType = "R"
        Case 1  'Special
            slType = "S"
        Case 2  'Sports
            slType = "P"
        Case 3  'Std Format
            slType = "F"
    End Select
    If ckcShowVersion.Value = vbChecked Then
        ilVer = ALLLIB
    Else
        ilVer = LATESTLIB
    End If
    'ilRet = gPopProgLibBox(PEvent, ilVer, slType, imVefCode, cbcSelect, lbcSelectCode)
    ilRet = gPopProgLibBox(PEvent, ilVer, slType, imVefCode, cbcSelect, tmSelectCode(), smSelectCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopProgLibBox)", PEvent
        On Error GoTo 0
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
        imPopReqd = True
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadCefRec                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a comment record          *
'*                                                     *
'*******************************************************
Private Function mReadCefRec(llCefCode As Long, ilForUpdate As Integer) As Integer
'
'   iRet = mReadCefRec()
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    tmCefSrchKey.lCode = llCefCode
    If llCefCode <> 0 Then
        tmCef.sComment = ""
        imCefRecLen = Len(tmCef)    '1009
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
        On Error GoTo mReadCefRecErr
        gBtrvErrorMsg ilRet, "mReadCefRec (btrGetEqual:Comment)", PEvent
        On Error GoTo 0
    Else
        tmCef.lCode = 0
        'tmCef.iStrLen = 0
        tmCef.sComment = ""
    End If
    mReadCefRec = True
    Exit Function
mReadCefRecErr:
    On Error GoTo 0
    tmCef.lCode = 0
    'tmCef.iStrLen = 0
    tmCef.sComment = ""
    mReadCefRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadEnfRec                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a comment record          *
'*                                                     *
'*******************************************************
Private Function mReadEnfRec(ilRowNo As Integer) As Integer
'
'   iRet = mReadEnfRec()
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    Dim ilEnfCode As Integer
    Dim ilEvtNameIndex As Integer
    tmEnf.iCode = 0
    tmEnf.lCefCode = 0
    mReadEnfRec = False
    ilEvtNameIndex = -1
    If smSave(2, ilRowNo) <> "" Then
        gFindMatch smSave(2, ilRowNo), 1, lbcEvtType
        If gLastFound(lbcEvtType) > 0 Then
            ilEvtNameIndex = gLastFound(lbcEvtType)
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    gFindMatch smSave(3, ilRowNo), 2, lbcEvtName(ilEvtNameIndex)
    If gLastFound(lbcEvtName(ilEvtNameIndex)) > 1 Then
        slNameCode = lbcEvtNameCode(ilEvtNameIndex).List(gLastFound(lbcEvtName(ilEvtNameIndex)) - 2)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If ilRet <> CP_MSG_NONE Then
            Exit Function
        End If
        slCode = Trim$(slCode)
        ilEnfCode = CInt(slCode)
    Else
        Exit Function
    End If
    tmEnfSrchKey.iCode = ilEnfCode
    If ilEnfCode <> 0 Then
        ilRet = btrGetEqual(hmEnf, tmEnf, imEnfRecLen, tmEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mReadEnfRecErr
        gBtrvErrorMsg ilRet, "mReadEnfRec (btrGetEqual:Event Name)", PEvent
        On Error GoTo 0
    Else
        Exit Function
    End If
    mReadEnfRec = True
    Exit Function
mReadEnfRecErr:
    On Error GoTo 0
    mReadEnfRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadLefRec                     *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadLefRec() As Integer
'
'   iRet = mReadLefRec
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpperBound As Integer
    Dim llRecPos As Long
    ReDim tmLef(0 To 0) As LEF
    ReDim lmEvtRecPos(0 To 0) As Long
    ReDim lmEvtIDCefCode(0 To 0) As Long
    ReDim lmCefCode(0 To 0) As Long
    ilUpperBound = UBound(tmLef)
    lmEvtRecPos(ilUpperBound) = 0
    lmEvtIDCefCode(ilUpperBound) = 0
    lmCefCode(ilUpperBound) = 0
    tmLefSrchKey.lLvfCode = tmLvf.lCode
    tmLefSrchKey.iStartTime(0) = 0
    tmLefSrchKey.iStartTime(1) = 0
    tmLefSrchKey.iSeqNo = 0
    ilRet = btrGetGreaterOrEqual(hmLef, tmLef(ilUpperBound), imLefRecLen, tmLefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmLef(ilUpperBound).lLvfCode = tmLvf.lCode)
        ilRet = btrGetPosition(hmLef, llRecPos)
        lmEvtRecPos(ilUpperBound) = llRecPos
        ilUpperBound = ilUpperBound + 1
        ReDim Preserve tmLef(0 To ilUpperBound) As LEF
        ReDim Preserve lmEvtRecPos(0 To ilUpperBound) As Long
        lmEvtRecPos(ilUpperBound) = 0
        ilRet = btrGetNext(hmLef, tmLef(ilUpperBound), imLefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_KEY_NOT_FOUND) Then
        On Error GoTo mReadLefRecErr
        gBtrvErrorMsg ilRet, "mReadLefRec (btrGetNext):" & "Lef.Btr", PEvent
        On Error GoTo 0
    End If
    vbcEvents.Min = LBound(tmLef)
    If UBound(tmLef) <= vbcEvents.LargeChange Then
        vbcEvents.Max = LBound(tmLef)
    Else
        vbcEvents.Max = UBound(tmLef) - vbcEvents.LargeChange ' - 1
    End If
    vbcEvents.Value = vbcEvents.Min
    ReDim smSave(0 To 11, 0 To UBound(tmLef) + 1) As String
    ReDim imSave(0 To 1, 0 To UBound(tmLef) + 1) As Integer
    ReDim smShow(0 To COMMENTINDEX, 0 To UBound(tmLef) + 1) As String
'    mInitContractCtrls
    mReadLefRec = True
    Exit Function
mReadLefRecErr:
    On Error GoTo 0
    mReadLefRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer, ilUseLibCode As Integer) As Integer
'
'   iRet = mReadRec(ilSelectIndex, ilUseLibCode)
'   Where:
'       ilSelectIndex (I) - list box index
'       ilUseLibCode(I)- True if from program duplicate, else use list index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    If Not ilUseLibCode Then
        slNameCode = tmSelectCode(ilSelectIndex - 1).sKey  'lbcSelectCode.List(ilSelectIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mReadRecErr
        gCPErrorMsg ilRet, "mReadRecErr (gParseItem field 2)", PEvent
        On Error GoTo 0
        slCode = Trim$(slCode)
        tmLvfSrchKey.lCode = CLng(slCode)
    Else
        tmLvfSrchKey.lCode = lmLibCode
    End If
    ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetEqual: Log Version Library)", PEvent
    On Error GoTo 0
    ilRet = btrGetPosition(hmLvf, lmLvfRecPos)  'Save position so correct record will be updated
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetPosition: Log Version Library)", PEvent
    On Error GoTo 0
    tmLtfSrchKey.iCode = tmLvf.iLtfCode
    ilRet = btrGetEqual(hmLtf, tmLtf, imLtfRecLen, tmLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetEqual: Log Title Library)", PEvent
    On Error GoTo 0
    ilRet = btrGetPosition(hmLtf, lmLtfRecPos)  'Save position so correct record will be updated
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetPosition: Log Title Library)", PEvent
    On Error GoTo 0
    mReadRec = True
    ReDim tmLef(0 To 0) As LEF
    ReDim lmEvtRecPos(0 To 0) As Long
    imLtfChg = False
    imLvfChg = False
    imLefChg = False
    imAllAnsw = False
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
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update LTF if name change only. *
'*                     Added LVF as a new version      *
'*                     Added LTF as a new version if   *
'*                        more changed then just the   *
'*                        name                         *
'*                     Added LEF as a new version      *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRowNo As Integer
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim ilLef As Integer
    Dim ilNewLtf As Integer
    Dim lRecPos As Long
    Dim tlLtf As LTF
    Dim tlLvf As LVF
    Dim tlLef As LEF
    Dim tlNmLtf As LTF
    Dim llNmLtfRecPos As Long
    Dim slType As String
    Dim ilFound As Integer
    Dim ilRet1 As Integer
    Dim ilCRet As Integer
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    If mSpecTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        mSaveRec = False
        Exit Function
    End If
    If Not mTestLibTimes() Then
        mSaveRec = False
        'cmcDates.SetFocus
        imSpecBoxNo = SPECLENGTHINDEX
        Exit Function
    End If
    For ilRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        If mEvtTestSaveFields(ilRowNo) = NO Then
            mSaveRec = False
            imEvtRowNo = ilRowNo
            Exit Function
        End If
    Next ilRowNo
    If mTestDuplAvailTimes() = YES Then
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    mMoveCtrlToRec True
    If imTerminate Then
        Screen.MousePointer = vbDefault    'Default
        mSaveRec = False
        Exit Function
    End If
    mMoveEvtCtrlToRec
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = btrBeginTrans(hmLtf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault    'Default
        ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
        imTerminate = True
        mSaveRec = False
        Exit Function
    End If
    If (imSelectedIndex <> 0) And (tmLtf.iCode <> 0) Then 'New selected
        ilRet = btrGetDirect(hmLtf, tlLtf, imLtfRecLen, lmLtfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrAbortTrans(hmLtf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
            imTerminate = True
            mSaveRec = False
            Exit Function
        End If
        'tmRec = tlLtf
        'ilRet = gGetByKeyForUpdate("LTF", hmLtf, tmRec)
        'tlLtf = tmRec
        'If ilRet <> BTRV_ERR_NONE Then
        '    ilRet = btrAbortTrans(hmLtf)
        '    Screen.MousePointer = vbDefault    'Default
        '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
        '    imTerminate = True
        '    mSaveRec = False
        '    Exit Function
        'End If
        ilRet = btrGetDirect(hmLvf, tlLvf, imLvfRecLen, lmLvfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrAbortTrans(hmLtf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
            imTerminate = True
            mSaveRec = False
            Exit Function
        End If
        'tmRec = tlLvf
        'ilRet = gGetByKeyForUpdate("LVF", hmLvf, tmRec)
        'tlLvf = tmRec
        'If ilRet <> BTRV_ERR_NONE Then
        '    ilRet = btrAbortTrans(hmLtf)
        '    Screen.MousePointer = vbDefault    'Default
        '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
        '    imTerminate = True
        '    mSaveRec = False
        '    Exit Function
        'End If
        'Test if name changed- if so, change all names
        If imLtfChg And (Trim$(tlLtf.sName) <> Trim$(tmLtf.sName)) And (tlLtf.iVar = tmLtf.iVar) And (tlLtf.sType = tmLtf.sType) Then
            'Only the name changed
            Do
                mMoveCtrlToRec True
                ilRet = btrUpdate(hmLtf, tmLtf, imLtfRecLen)
                If ilRet = BTRV_ERR_CONFLICT Then
                    ilCRet = btrGetDirect(hmLtf, tlLtf, imLtfRecLen, lmLtfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilCRet <> BTRV_ERR_NONE Then
                        ilRet = btrAbortTrans(hmLtf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    'tmRec = tlLtf
                    'ilCRet = gGetByKeyForUpdate("LTF", hmLtf, tmRec)
                    'tlLtf = tmRec
                    'If ilCRet <> BTRV_ERR_NONE Then
                    '    ilRet = btrAbortTrans(hmLtf)
                    '    Screen.MousePointer = vbDefault    'Default
                    '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                    '    imTerminate = True
                    '    mSaveRec = False
                    '    Exit Function
                    'End If
                    ilCRet = btrGetDirect(hmLvf, tlLvf, imLvfRecLen, lmLvfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilCRet <> BTRV_ERR_NONE Then
                        ilRet = btrAbortTrans(hmLtf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    'tmRec = tlLvf
                    'ilCRet = gGetByKeyForUpdate("LVF", hmLvf, tmRec)
                    'tlLvf = tmRec
                    'If ilCRet <> BTRV_ERR_NONE Then
                    '    ilRet = btrAbortTrans(hmLtf)
                    '    Screen.MousePointer = vbDefault    'Default
                    '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                    '    imTerminate = True
                    '    mSaveRec = False
                    '    Exit Function
                    'End If
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmLtf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        Else    'Variation or type changed- add as new
            If imLtfChg And ((tlLtf.iVar <> tmLtf.iVar) Or (tlLtf.sType <> tmLtf.sType)) Then
                imSelectedIndex = 0 'Force to add as new
                'Name in all variations not being updated- fix later
                If imLtfChg And (Trim$(tlLtf.sName) <> Trim$(tmLtf.sName)) And (tlLtf.iVar <> tmLtf.iVar) And (tlLtf.sType = tmLtf.sType) Then
'                    Would have to read all LTF and test if name matched the old name and type matched- then update
                    ilRet = btrGetFirst(hmLtf, tlNmLtf, imLtfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    Do
                       If ilRet = BTRV_ERR_NONE Then
                           If (Trim$(tlLtf.sName) = Trim$(tlNmLtf.sName)) And (tlLtf.sType = tlNmLtf.sType) And (tlLtf.iVefCode = tlNmLtf.iVefCode) Then
                               ilRet = btrGetPosition(hmLtf, llNmLtfRecPos)
                               Do
                                   tlNmLtf.sName = tmLtf.sName
                                   ilRet = btrUpdate(hmLtf, tlNmLtf, imLtfRecLen)
                                   If ilRet = BTRV_ERR_CONFLICT Then
                                        ilRet1 = btrGetDirect(hmLtf, tlNmLtf, imLtfRecLen, llNmLtfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                        'tmRec = tlNmLtf
                                        'ilRet1 = gGetByKeyForUpdate("LTF", hmLtf, tmRec)
                                        'tlNmLtf = tmRec
                                        'If ilRet1 <> BTRV_ERR_NONE Then
                                        '    ilRet = btrAbortTrans(hmLtf)
                                        '    Screen.MousePointer = vbDefault    'Default
                                        '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                                        '    imTerminate = True
                                        '    mSaveRec = False
                                        '    Exit Function
                                        'End If
                                   End If
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                If ilRet <> BTRV_ERR_NONE Then
                                    ilRet = btrAbortTrans(hmLtf)
                                    Screen.MousePointer = vbDefault    'Default
                                    ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                                    imTerminate = True
                                    mSaveRec = False
                                    Exit Function
                                End If
                           End If
                           ilRet = btrGetNext(hmLtf, tlNmLtf, imLtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                       End If
                    Loop While ilRet = BTRV_ERR_NONE
                    'If ilRet <> BTRV_ERR_NONE Then
                    If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_END_OF_FILE) Then
                        ilRet = btrAbortTrans(hmLtf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    ilRet = btrGetDirect(hmLtf, tlLtf, imLtfRecLen, lmLtfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilRet = btrAbortTrans(hmLtf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    'tmRec = tlLtf
                    'ilRet = gGetByKeyForUpdate("LTF", hmLtf, tmRec)
                    'tlLtf = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    ilRet = btrAbortTrans(hmLtf)
                    '    Screen.MousePointer = vbDefault    'Default
                    '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                    '    imTerminate = True
                    '    mSaveRec = False
                    '    Exit Function
                    'End If
                End If
            End If
        End If
    End If
    If imLtfChg Or imLvfChg Or imLefChg Then
        Do  'Loop until record updated or added
            slStamp = gFileDateTime(sgDBPath & "Lvf.Btr")
            'If Len(lbcSelectCode.Tag) > Len(slStamp) Then
            '    slStamp = slStamp & Right$(lbcSelectCode.Tag, Len(lbcSelectCode.Tag) - Len(slStamp))
            'End If
            If Len(smSelectCodeTag) > Len(slStamp) Then
                slStamp = slStamp & right$(smSelectCodeTag, Len(smSelectCodeTag) - Len(slStamp))
            End If
            tmLvf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
            If (imSelectedIndex = 0) Or (tmLtf.iCode = 0) Then 'New selected
                ilNewLtf = True
                tmLtf.iCode = 0  'Autoincrement
    '            tmLnf.iVar = 0
    '            tmLnf.iVersion = 1
                ilRet = btrInsert(hmLtf, tmLtf, imLtfRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = btrAbortTrans(hmLtf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                tmLvf.lCode = 0
                tmLvf.iLtfCode = tmLtf.iCode
                tmLvf.iVersion = 1
                'tmLvf.iBaseTime(0) = 1  'Relative time
                'tmLvf.iBaseTime(1) = 0
                ilRet = btrInsert(hmLvf, tmLvf, imLvfRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert: Library)"
            Else 'Old record-Update
                ilNewLtf = False
                'Determine next version number
                ilFound = True
                tmLvf1SrchKey.iLtfCode = tmLtf.iCode
                tmLvf1SrchKey.iVersion = 32000
                ilRet = btrGetGreaterOrEqual(hmLvf, tlLvf, imLvfRecLen, tmLvf1SrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If (ilRet = BTRV_ERR_NONE) And (tlLvf.iLtfCode = tmLtf.iCode) Then
                    tmLvf.iVersion = tlLvf.iVersion + 1
                    tmLvf.lCode = 0  'Autoincrement
                    ilRet = btrInsert(hmLvf, tmLvf, imLvfRecLen, INDEXKEY0)
                    slMsg = "mSaveRec (btrInsert: Library)"
                Else
                    ilRet = BTRV_ERR_NONE
                End If
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrAbortTrans(hmLtf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
            imTerminate = True
            mSaveRec = False
            Exit Function
        End If
        For ilLef = LBound(tmLef) To UBound(tmLef) - 1 Step 1
            tmLef(ilLef).iUrfCode = tgUrf(0).iCode 'Use first record retained for user
            tmLef(ilLef).lLvfCode = tmLvf.lCode
            lRecPos = Val(smSave(10, ilLef + 1))
            If (imSelectedIndex = 0) Or ilNewLtf Or (lRecPos = 0) Then 'New selected
                'tmLef(ilLef).lCifCode = 0
                tmLef(ilLef).sPreFinal = " "
            Else
                ilRet = btrGetDirect(hmLef, tlLef, imLefRecLen, lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = btrAbortTrans(hmLtf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                'tmRec = tlLef
                'ilRet = gGetByKeyForUpdate("LEF", hmLef, tmRec)
                'tlLef = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilRet = btrAbortTrans(hmLtf)
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                '    imTerminate = True
                '    mSaveRec = False
                '    Exit Function
                'End If
                ''tmLef(ilLef).lCifCode = tlLef.lCifCode
                tmLef(ilLef).sPreFinal = tlLef.sPreFinal
            End If

            'tmCef.iStrLen = Len(Trim$(smSave(11, ilLef + 1)))
            tmCef.sComment = Trim$(smSave(11, ilLef + 1)) & Chr$(0) '& Chr$(0) 'sgTB
            Do  'Loop until record updated or added
                imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
                'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
                If Trim$(smSave(11, ilLef + 1)) <> "" Then
                    tmCef.lCode = 0 'Autoincrement
                    ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                Else
                    tmCef.lCode = 0
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrInsert: Event ID)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmLtf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
            tmLef(ilLef).lEvtIDCefCode = tmCef.lCode
            If smSave(8, ilLef + 1) = "^" Then  'Comment-the symbol ^ is used to indicate first time and no comment
                smSave(8, ilLef + 1) = ""
            End If
            'tmCef.iStrLen = Len(Trim$(smSave(8, ilLef + 1)))
            tmCef.sComment = Trim$(smSave(8, ilLef + 1)) & Chr$(0) '& Chr$(0) 'sgTB
            Do  'Loop until record updated or added
                imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
                'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
                If Trim$(smSave(8, ilLef + 1)) <> "" Then
                    tmCef.lCode = 0 'Autoincrement
                    ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                Else
                    tmCef.lCode = 0
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrInsert: Comment)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmLtf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
            tmLef(ilLef).lCefCode = tmCef.lCode
        Next ilLef
        For ilLef = LBound(tmLef) To UBound(tmLef) - 1 Step 1
            Do  'Loop until record updated or added
                tmLef(ilLef).lCode = 0
                ilRet = btrInsert(hmLef, tmLef(ilLef), imLefRecLen, INDEXKEY2)
                slMsg = "mSaveRec (btrInsert: Library Events)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmLtf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        Next ilLef
    End If
    ilRet = btrEndTrans(hmLtf)
    'Add to log calendar-pending
    If igViewType = 1 Then
        slType = "A"
    Else
        slType = "O"
    End If
    'If cmcDates.Enabled Then
    '    gPrgToPend PEvent, tmLvf, slType
    'End If
    'If lbcSelectCode.Tag <> "" Then
    '    If slStamp = lbcSelectCode.Tag Then
    '        lbcSelectCode.Tag = FileDateTime(sgDBPath & "Lvf.Btr")
    '        If Len(slStamp) > Len(lbcSelectCode.Tag) Then
    '            lbcSelectCode.Tag = lbcSelectCode.Tag & Right$(slStamp, Len(slStamp) - Len(lbcSelectCode.Tag))
    '        End If
    '    End If
    'End If
'    If smSelectCodeTag <> "" Then
'        If slStamp = smSelectCodeTag Then
'            smSelectCodeTag = gFileDateTime(sgDBPath & "Lvf.Btr")
'            If Len(slStamp) > Len(smSelectCodeTag) Then
'                smSelectCodeTag = smSelectCodeTag & right$(slStamp, Len(slStamp) - Len(smSelectCodeTag))
'            End If
'        End If
'    End If
'    'If (Not ckcShowVersion.Value) Then
'        If imSelectedIndex <> 0 Then
'            'lbcSelectCode.RemoveItem imSelectedIndex - 1
'            gRemoveItemFromSortCode imSelectedIndex - 1, tmSelectCode()
'            cbcSelect.RemoveItem imSelectedIndex
'        End If
'    'End If
'    'Add to top of list- the sort was such that the newest should be at top
'    cbcSelect.RemoveItem 0 'Remove [New]
'    If tmLtf.iVar <> 0 Then
'        slName = Trim$(tmLtf.sName) & "-" & Trim$(Str$(tmLtf.iVar))
'    Else
'        slName = Trim$(tmLtf.sName)
'    End If
'    If ckcShowVersion.Value = vbChecked Then
'        slNameVer = slName & "/" & Trim$(Str$(tmLvf.iVersion))
'        cbcSelect.AddItem slNameVer, 0
'    Else
'        cbcSelect.AddItem slName, 0
'    End If
'    cbcSelect.AddItem "[New]", 0
'    slName = slName + "\" + LTrim$(Str$(tmLvf.lCode))
'    'lbcSelectCode.AddItem slName, 0
'    gAddItemToSortCode slName, tmSelectCode(), True
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
    If (mSpecTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO) Or imLefChg Then
        If imLtfChg Or imLvfChg Or imLefChg Or imDateAssign Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    If Val(smSpecSave(2)) <> 0 Then
                        slMess = "Add " & smSpecSave(1) & "-" & smSpecSave(2)
                    Else
                        slMess = "Add " & smSpecSave(1)
                    End If
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcLibSpec_Paint
                    Exit Function
                End If
                If ilRes = vbYes Then
                    ilRes = mSaveRec()
                    mSaveRecChg = ilRes
                    Exit Function
                End If
                If ilRes = vbNo Then
                    cmcUndo_Click
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
'*             Created:9/22/93       By:D. LeVine      *
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
    If imLtfChg Or imLvfChg Or imLefChg Or imDateAssign Then
        cbcType.Enabled = False
        cbcSelect.Enabled = False
    Else
        cbcType.Enabled = True
        cbcSelect.Enabled = True
    End If
    If (imLtfChg Or imLvfChg Or imLefChg Or imDateAssign) And (UBound(smSave, 2) >= 2) Then  'At least one event added
        If imUpdateAllowed Then
            cmcUpdate.Enabled = True
        Else
            cmcUpdate.Enabled = False
        End If
    Else
        cmcUpdate.Enabled = False
    End If
    If ((imSelectedIndex = 0) And (smSpecSave(1) <> "")) Or (imSelectedIndex > 0) Then
        pbcEvents.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
    Else
        pbcEvents.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
    End If
    'Revert button set if any field changed
    If imLtfChg Or imLvfChg Or imLefChg Or imDateAssign Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    If (imSelectedIndex > 0) Or imLtfChg Or imLvfChg Or imLefChg Or imDateAssign Then
        If imUpdateAllowed Then
            cmcErase.Enabled = True
        Else
            cmcErase.Enabled = False
        End If
    Else
        cmcErase.Enabled = False
    End If
'    If (imEvtBoxNo >= imLBEvtCtrls) And (imEvtBoxNo <= UBound(tmEvtCtrls)) Then
'        If (imEvtRowNo >= vbcEvents.Value + 1) And (imEvtRowNo < vbcEvents.Value + vbcEvents.LargeChange + 2) Then
'            imcTrash.Enabled = True
'        Else
'            imcTrash.Enabled = False
'        End If
'    Else
'        imcTrash.Enabled = False
'    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSortByRowTime                  *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move row if time of row is out *
'*                      of order                       *
'*                                                     *
'*******************************************************
Private Function mSortByRowTime(ilCurRowNo As Integer) As Integer
'
'   ilNewRow = mSortByRowTime (ilCurRowNo)
'   Where:
'       ilCurRowNo (I)- Current row number to be checked
'       ilNewRow (O) - Row number moved into
'
    Dim ilLoop As Integer
    Dim ilLastRowNo As Integer
    Dim ilTestRowNo As Integer
    Dim slEvtStart As String
    Dim slTestEvtStart As String
    mSortByRowTime = ilCurRowNo
    'If (ilCurRowNo < LBound(smSave, 2)) Or (ilCurRowNo > UBound(smSave, 2)) Then
    If (ilCurRowNo < LBONE) Or (ilCurRowNo > UBound(smSave, 2)) Then
        Exit Function
    End If
    If (ilCurRowNo = UBound(smSave, 2)) And (Len(smSave(1, ilCurRowNo)) = 0) Then
        Exit Function
    End If
    slEvtStart = gFormatLength(smSave(1, ilCurRowNo), "3", False)
    ilLastRowNo = UBound(smSave, 2) - 1
    'For ilLoop = LBound(smSave, 2) To ilLastRowNo Step 1
    For ilLoop = LBONE To ilLastRowNo Step 1
        If ilLoop <> ilCurRowNo Then    'Don't test same row
            slTestEvtStart = gFormatLength(smSave(1, ilLoop), "3", False)
            If gLengthToCurrency(slEvtStart) < gLengthToCurrency(slTestEvtStart) Then
                If ilLoop <> ilCurRowNo + 1 Then    'If adjacent rows- move not required
                    If ilLoop < ilCurRowNo Then
                        mSortByRowTime = ilLoop
                    Else
                        mSortByRowTime = ilLoop - 1
                    End If
                    mMoveEvt False, ilLoop, ilCurRowNo
                    pbcEvents.Cls
                    pbcEvents_Paint
                End If
                Exit Function
            ElseIf gLengthToCurrency(slEvtStart) = gLengthToCurrency(slTestEvtStart) Then
                If ilCurRowNo + 1 = ilLoop Then
                    Exit Function
                End If
                'If program event being sorted-place it prior to other events at same time except
                'other programs
                If Asc(smSave(9, ilCurRowNo)) = Asc("1") Then
                    ilTestRowNo = ilLoop
                    Do
                        If ilTestRowNo > ilLastRowNo Then
                            ilTestRowNo = ilTestRowNo - 1
                            mSortByRowTime = ilTestRowNo
                            mMoveEvt True, ilTestRowNo, ilCurRowNo
                            pbcEvents.Cls
                            pbcEvents_Paint
                            Exit Function
                        End If
                        If ilTestRowNo = ilCurRowNo Then
                            Exit Function
                        End If
                        If Asc(smSave(9, ilTestRowNo)) <> Asc("1") Then
                            If ilTestRowNo < ilCurRowNo Then
                                mSortByRowTime = ilTestRowNo
                            Else
                                mSortByRowTime = ilTestRowNo - 1
                            End If
                            mMoveEvt False, ilTestRowNo, ilCurRowNo
                            pbcEvents.Cls
                            pbcEvents_Paint
                            Exit Function
                        End If
                        slTestEvtStart = gFormatLength(smSave(1, ilTestRowNo), "3", False)
                        If gLengthToCurrency(slEvtStart) <> gLengthToCurrency(slTestEvtStart) Then
                            If ilTestRowNo < ilCurRowNo Then
                                mSortByRowTime = ilTestRowNo
                            Else
                                mSortByRowTime = ilTestRowNo - 1
                            End If
                            mMoveEvt False, ilTestRowNo, ilCurRowNo
                            pbcEvents.Cls
                            pbcEvents_Paint
                            Exit Function
                        End If
                        ilTestRowNo = ilTestRowNo + 1
                    Loop
                    Exit Function
                End If
                ilTestRowNo = ilLoop + 1
                Do
                    If ilTestRowNo > ilLastRowNo Then
                        ilTestRowNo = ilTestRowNo - 1
                        mSortByRowTime = ilTestRowNo
                        mMoveEvt True, ilTestRowNo, ilCurRowNo
                        pbcEvents.Cls
                        pbcEvents_Paint
                        Exit Function
                    End If
                    If ilTestRowNo = ilCurRowNo Then
                        Exit Function
                    End If
                    slTestEvtStart = gFormatLength(smSave(1, ilTestRowNo), "3", False)
                    If gLengthToCurrency(slEvtStart) <> gLengthToCurrency(slTestEvtStart) Then
                        If ilTestRowNo < ilCurRowNo Then
                            mSortByRowTime = ilTestRowNo
                        Else
                            mSortByRowTime = ilTestRowNo - 1
                        End If
                        mMoveEvt False, ilTestRowNo, ilCurRowNo
                        pbcEvents.Cls
                        pbcEvents_Paint
                        Exit Function
                    End If
                    ilTestRowNo = ilTestRowNo + 1
                Loop
                Exit Function
            End If
        End If
    Next ilLoop
    If ilLastRowNo <> ilCurRowNo Then
        If gLengthToCurrency(slEvtStart) >= gLengthToCurrency(slTestEvtStart) Then
            mSortByRowTime = ilLastRowNo
            mMoveEvt True, ilLastRowNo, ilCurRowNo
            pbcEvents.Cls
            pbcEvents_Paint
            Exit Function
        End If
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecDirection                  *
'*                                                     *
'*             Created:9/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move to box indicated by       *
'*                      user direction                 *
'*                                                     *
'*******************************************************
Private Sub mSpecDirection(ilMoveDir As Integer)
'
'   mSpecDirection ilMove
'   Where:
'       ilMove (I)- 0=Up; 1= down; 2= left; 3= right
'
    mEvtSetShow imEvtBoxNo
    Select Case ilMoveDir
        Case KEYUP  'Up
        Case KeyDown  'Down
        Case KEYLEFT  'Left
            If imSpecBoxNo > SPECLIBNAMEINDEX Then
                imSpecBoxNo = imSpecBoxNo - 1
            Else
                imSpecBoxNo = SPECLENGTHINDEX
            End If
        Case KEYRIGHT  'Right
            If imSpecBoxNo < SPECLENGTHINDEX Then
                imSpecBoxNo = imSpecBoxNo + 1
            Else
                imSpecBoxNo = SPECLIBNAMEINDEX
            End If
    End Select
    mEvtEnableBox imEvtBoxNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecEnableBox                  *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
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
    If ilBoxNo < imLBSpecCtrls Or ilBoxNo > UBound(tmSpecCtrls) Then
        Exit Sub
    End If

    imAllAnsw = False
    Select Case ilBoxNo 'Branch on box type (control)
        Case SPECLIBTYPEINDEX
            If imSpecSave(1) < 0 Then
                If (igPrgDupl) And (imTypeSelectedIndex = 3) Then
                    imSpecSave(1) = 0   'Change Std Format to regular if duplicating
                    'If imAllowDate Then
                    '    cmcDates.Enabled = True
                    'End If
                Else
                    imSpecSave(1) = imTypeSelectedIndex
                    'If imSpecSave(1) = 3 Then   'Std Format
                    '    cmcDates.Enabled = False
                    'Else
                    '    If imAllowDate Then
                    '        cmcDates.Enabled = True
                    '    End If
                    'End If
                End If
                imLtfChg = True
                tmSpecCtrls(ilBoxNo).iChg = True
            End If
            pbcLibType.Width = tmSpecCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcLibSpec, pbcLibType, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            pbcLibType_Paint
            pbcLibType.Visible = True
            pbcLibType.SetFocus
        Case SPECLIBNAMEINDEX
            mLibPop
            If imTerminate Then
                Exit Sub
            End If
            lbcLibName.height = gListBoxHeight(lbcLibName.ListCount, 10)
            edcSpecDropDown.Width = tmSpecCtrls(ilBoxNo).fBoxW - cmcSpecDropDown.Width
            edcSpecDropDown.MaxLength = 20
            gMoveFormCtrl pbcLibSpec, edcSpecDropDown, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            imChgMode = True
            gFindMatch smSpecSave(1), 0, lbcLibName
            If gLastFound(lbcLibName) >= 0 Then
                lbcLibName.ListIndex = gLastFound(lbcLibName)
                edcSpecDropDown.Text = lbcLibName.List(lbcLibName.ListIndex)
            Else
                lbcLibName.ListIndex = -1    '[None]
                edcSpecDropDown.Text = smSpecSave(1)
            End If
            imChgMode = False
            lbcLibName.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.height
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            cmcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        Case SPECVARINDEX
        Case SPECLENGTHINDEX
            edcSpecDropDown.Width = tmSpecCtrls(ilBoxNo).fBoxW - cmcSpecDropDown.Width
            edcSpecDropDown.MaxLength = 9
            gMoveFormCtrl pbcLibSpec, edcSpecDropDown, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            plclen.Move edcSpecDropDown.Left, edcSpecDropDown.Top + edcSpecDropDown.height
            edcSpecDropDown.Text = Trim$(smSpecSave(3))
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True  'Set visibility
            cmcSpecDropDown.Visible = True
            plclen.Visible = True
            edcSpecDropDown.SetFocus
        Case SPECBASETIMEINDEX
            edcSpecDropDown.Width = tmSpecCtrls(ilBoxNo).fBoxW - cmcSpecDropDown.Width
            edcSpecDropDown.MaxLength = 10
            gMoveFormCtrl pbcLibSpec, edcSpecDropDown, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            cmcSpecDropDown.Move edcSpecDropDown.Left + edcSpecDropDown.Width, edcSpecDropDown.Top
            plcTme.Move cmcSpecDropDown.Left + cmcSpecDropDown.Width - plcTme.Width, edcSpecDropDown.Top + edcSpecDropDown.height
            edcSpecDropDown.Text = Trim$(smSpecSave(4))
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True  'Set visibility
            cmcSpecDropDown.Visible = True
            plcTme.Visible = True
            edcSpecDropDown.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecSetFocus                   *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
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
    If ilBoxNo < imLBSpecCtrls Or ilBoxNo > UBound(tmSpecCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case SPECLIBTYPEINDEX
            If pbcLibType.Enabled Then
                pbcLibType.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case SPECLIBNAMEINDEX
            If edcSpecDropDown.Enabled Then
                edcSpecDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case SPECVARINDEX
        Case SPECLENGTHINDEX
            If edcSpecDropDown.Enabled Then
                edcSpecDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
        Case SPECBASETIMEINDEX
            If edcSpecDropDown.Enabled Then
                edcSpecDropDown.SetFocus
            Else
                pbcClickFocus.SetFocus
            End If
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecSetShow                    *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSpecSetShow(ilBoxNo As Integer)
'
'   mSpecSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilRowNo As Integer
    Dim slXMid As String

    If ilBoxNo < imLBSpecCtrls Or ilBoxNo > UBound(tmSpecCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case SPECLIBTYPEINDEX
            pbcLibType.Visible = False  'Set visibility
            Select Case imSpecSave(1)
                Case 0
                    slStr = "Regular"
                Case 1
                    slStr = "Special"
                Case 2
                    slStr = "Sport"
                Case 3
                    slStr = "Std Format"
                Case Else
                    slStr = ""
            End Select
            gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
        Case SPECLIBNAMEINDEX
            lbcLibName.Visible = False
            edcSpecDropDown.Visible = False
            cmcSpecDropDown.Visible = False
            slStr = edcSpecDropDown.Text
            gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
            If smSpecSave(1) <> edcSpecDropDown.Text Then
                imLtfChg = True
                tmSpecCtrls(ilBoxNo).iChg = True
            End If
            smSpecSave(1) = edcSpecDropDown.Text
            gFindMatch smSpecSave(1), 0, lbcLibName
            If gLastFound(lbcLibName) >= 0 Then
                ilPos = InStr(smSpecSave(1), "-")
                If ilPos > 0 Then
                    slStr = Mid$(smSpecSave(1), ilPos + 1)
                    smSpecSave(2) = Trim$(str$(Val(slStr) + 1))
                Else
                    smSpecSave(2) = "1"
                End If
            Else
                smSpecSave(2) = ""
            End If
            slStr = smSpecSave(2)
            gSetShow pbcLibSpec, slStr, tmSpecCtrls(SPECVARINDEX)
        Case SPECVARINDEX
        Case SPECLENGTHINDEX
            plclen.Visible = False
            cmcSpecDropDown.Visible = False
            edcSpecDropDown.Visible = False  'Set visibility
            slStr = edcSpecDropDown.Text
            If gValidLength(slStr) Then
                slStr = gFormatLength(slStr, "3", False)
                gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
                If smSpecSave(3) <> edcSpecDropDown.Text Then
                    imLvfChg = True
                    tmSpecCtrls(ilBoxNo).iChg = True
                End If
                smSpecSave(3) = edcSpecDropDown.Text
            Else
                Beep
                edcSpecDropDown.Text = smSpecSave(3)
            End If
        Case SPECBASETIMEINDEX
            plcTme.Visible = False
            cmcSpecDropDown.Visible = False
            edcSpecDropDown.Visible = False  'Set visibility
            slStr = edcSpecDropDown.Text
            If Len(slStr) > 0 Then
                If gValidTime(slStr) Then
                    slStr = gFormatTime(slStr, "A", "1")
                    gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
                    If (gTimeToCurrency(smSpecSave(4), False) <> gTimeToCurrency(slStr, False)) Or (imTimeRelative) Then
                        imTimeRelative = False
                        imLvfChg = True
                        tmSpecCtrls(ilBoxNo).iChg = True
                        'Reset times for events
                        smSpecSave(4) = edcSpecDropDown.Text
                        ''For ilRowNo = LBound(tmLef) + 1 To UBound(tmLef) Step 1
                        'For ilRowNo = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
                        For ilRowNo = LBONE To UBound(smSave, 2) - 1 Step 1
                            slStr = smSave(1, ilRowNo)
                            gAddTimeLength smSpecSave(4), slStr, "A", "1", slStr, slXMid
                            gSetShow pbcEvents, slStr, tmEvtCtrls(TIMEINDEX)
                            smShow(TIMEINDEX, ilRowNo) = tmEvtCtrls(TIMEINDEX).sShow
                        Next ilRowNo
                        pbcEvents.Cls
                        pbcEvents_Paint
                    End If
                    imTimeRelative = False
                    smSpecSave(4) = edcSpecDropDown.Text
                Else
                    Beep
                    edcSpecDropDown.Text = smSpecSave(4)
                End If
            Else
                If Not imTimeRelative Then
                    imTimeRelative = True
                    imLvfChg = True
                    tmSpecCtrls(ilBoxNo).iChg = True
                    slStr = ""
                    gSetShow pbcLibSpec, slStr, tmSpecCtrls(ilBoxNo)
                    'For ilRowNo = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
                    For ilRowNo = LBONE To UBound(smSave, 2) - 1 Step 1
                        slStr = smSave(1, ilRowNo)
                        slStr = gFormatLength(slStr, "3", False)
                        gSetShow pbcEvents, slStr, tmEvtCtrls(TIMEINDEX)
                        smShow(TIMEINDEX, ilRowNo) = tmEvtCtrls(TIMEINDEX).sShow
                    Next ilRowNo
                    pbcEvents.Cls
                    pbcEvents_Paint
                End If
                smSpecSave(4) = ""
            End If
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecTestFields                     *
'*                                                     *
'*             Created:9/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mSpecTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mSpecTestFields(iTest, iState)
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

    If (ilCtrlNo = SPECLIBTYPEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        Select Case imSpecSave(1)
            Case 0
                slStr = "Regular"
            Case 1
                slStr = "Special"
            Case 2
                slStr = "Sport"
            Case 3
                slStr = "Std Format"
            Case Else
                slStr = ""
        End Select
        If gFieldDefinedStr(slStr, "", "Library Type must be specified", tmSpecCtrls(SPECLIBTYPEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imSpecBoxNo = SPECLIBTYPEINDEX
            End If
            mSpecTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SPECLIBNAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSpecSave(1), "", "Library Name must be specified", tmSpecCtrls(SPECLIBNAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imSpecBoxNo = SPECLIBNAMEINDEX
            End If
            mSpecTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SPECLENGTHINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSpecSave(3), "", "Length must be specified", tmSpecCtrls(SPECLENGTHINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imSpecBoxNo = SPECLENGTHINDEX
            End If
            mSpecTestFields = NO
            Exit Function
        End If
        If Not gValidLength(smSpecSave(3)) Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                ilRes = MsgBox("Length must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
                imSpecBoxNo = SPECLENGTHINDEX
            End If
            mSpecTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SPECLENGTHINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If Not gValidTime(smSpecSave(4)) Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                ilRes = MsgBox("Time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
                imSpecBoxNo = SPECLENGTHINDEX
            End If
            mSpecTestFields = NO
            Exit Function
        End If
    End If
    mSpecTestFields = YES
End Function
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


    sgDoneMsg = Trim$(str$(igAdvtCallSource)) & "\" & sgAdvtName
    Screen.MousePointer = vbDefault
    'Unload IconTraf
    igManUnload = YES
    'Unload Traffic
    Unload PEvent
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestDuplAvailTimes             *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test for duplicate avail times *
'*                                                     *
'*******************************************************
Private Function mTestDuplAvailTimes() As Integer
    Dim ilRowNo As Integer
    Dim ilLoop As Integer
    Dim slEvtLen As String
    Dim slLen As String
    Dim ilRes As Integer
    For ilRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        If (Asc(smSave(9, ilRowNo)) >= Asc("2")) And (Asc(smSave(9, ilRowNo)) <= Asc("9")) Then
            slEvtLen = gFormatLength(smSave(1, ilRowNo), "3", False)
            For ilLoop = ilRowNo + 1 To UBound(smSave, 2) - 1 Step 1
                If (Asc(smSave(9, ilLoop)) >= Asc("2")) And (Asc(smSave(9, ilLoop)) <= Asc("9")) Then
                    slLen = gFormatLength(smSave(1, ilLoop), "3", False)
                    If gLengthToCurrency(slEvtLen) = gLengthToCurrency(slLen) Then
                        ilRes = MsgBox("Avail Times Matching", vbOKOnly + vbExclamation, "Incomplete")
                        imEvtBoxNo = TIMEINDEX
                        imEvtRowNo = ilLoop
                        mTestDuplAvailTimes = YES
                        Exit Function
                    End If
                End If
            Next ilLoop
        End If
    Next ilRowNo
    mTestDuplAvailTimes = NO
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestLibTimes                   *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if start time plus   *
'*                      library length exceeds 12Mid   *
'*                                                     *
'*******************************************************
Private Function mTestLibTimes() As Integer
    Dim ilLoop As Integer
    Dim llStartTime As Long
    Dim llLength As Long
    Dim ilRes As Integer
    For ilLoop = LBound(tgPrg) To UBound(tgPrg) - 1 Step 1
        If (tgPrg(ilLoop).sStartTime <> "") And (tgPrg(ilLoop).sStartDate <> "") Then
            llStartTime = CLng(gTimeToCurrency(tgPrg(ilLoop).sStartTime, False))
            llLength = CLng(gLengthToCurrency(smSpecSave(3)))
            If llStartTime + llLength > 86400 Then
                ilRes = MsgBox("Library end time exceeds 12Midnight", vbOKOnly + vbExclamation, "Error")
                mTestLibTimes = False
                Exit Function
            End If
        End If
    Next ilLoop
    mTestLibTimes = True
    Exit Function
End Function
Private Sub pbcArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcCLibType_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(" ") Then
        If imViewLibType = 0 Then
            imViewLibType = 1
        ElseIf imViewLibType = 1 Then
            imViewLibType = 2
        ElseIf imViewLibType = 2 Then
            imViewLibType = 3
        Else
            imViewLibType = 0
        End If
        pbcCLibType_Paint
        mCLibPop
    End If
End Sub
Private Sub pbcCLibType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imViewLibType = 0 Then
        imViewLibType = 1
    ElseIf imViewLibType = 1 Then
        imViewLibType = 2
    ElseIf imViewLibType = 2 Then
        imViewLibType = 3
    Else
        imViewLibType = 0
    End If
    pbcCLibType_Paint
    mCLibPop
End Sub
Private Sub pbcCLibType_Paint()
    pbcCLibType.Cls
    pbcCLibType.CurrentX = fgBoxInsetX
    pbcCLibType.CurrentY = -15 'fgBoxInsetY
    If imViewLibType = 0 Then
        pbcCLibType.Print "Regular"
    ElseIf imViewLibType = 1 Then
        pbcCLibType.Print "Special"
    ElseIf imViewLibType = 2 Then
        pbcCLibType.Print "Sport"
    ElseIf imViewLibType = 3 Then
        pbcCLibType.Print "Std Format"
    End If
End Sub
Private Sub pbcClickFocus_GotFocus()
    Dim slStr As String
    If (imEvtBoxNo = TIMEINDEX) And (edcEvtDropDown.Text <> "") Then
        If imTimeRelative Then
            slStr = edcEvtDropDown.Text
            If Not gValidLength(slStr) Then
                Beep
                edcEvtDropDown.SetFocus
                Exit Sub
            End If
        Else
            slStr = edcEvtDropDown.Text
            If Not gValidTime(slStr) Then
                Beep
                edcEvtDropDown.SetFocus
                Exit Sub
            End If
        End If
    End If
    If imEvtBoxNo = LENGTHINDEX Then
        slStr = edcEvtDropDown.Text
        If Not gValidLength(slStr) Then
            Beep
            edcEvtDropDown.SetFocus
            Exit Sub
        End If
    End If
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
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
Private Sub pbcClock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim llTime As Long
    Dim llRadius As Long
    Dim ilFound As Integer
    Dim ilArc As Integer
    imButton = Button
    If (imButton And &H2) = 2 Then  'Right mouse- show values
        imIgnoreRightMove = True
        mCompVector CLng(X), CLng(Y), llTime, llRadius
        'Determine if in an event
        ilFound = -1
        For ilArc = LBound(lmArcTimes, 2) To UBound(lmArcTimes, 2) - 1 Step 1
            If (llTime >= lmArcTimes(1, ilArc) - 5) And (llTime <= lmArcTimes(2, ilArc) + 5) And llRadius <= lmArcTimes(3, ilArc) Then
                If ilFound = -1 Then
                    ilFound = ilArc
                Else
                    If lmArcTimes(3, ilArc) < lmArcTimes(3, ilFound) Then
                        ilFound = ilArc
                    End If
                End If
            End If
        Next ilArc
        imLastArcPainted = ilFound
        pbcCEvents.Cls
        If ilFound >= 0 Then
            ilFound = lmArcTimes(4, ilFound)
            For ilBox = imLBCEvtCtrls To COMMENTINDEX Step 1
                tmCEvtCtrls(ilBox).sShow = smShow(ilBox, ilFound)
            Next ilBox
            For ilBox = imLBCEvtCtrls To COMMENTINDEX Step 1
                pbcCEvents.CurrentX = tmCEvtCtrls(ilBox).fBoxX + fgBoxInsetX
                pbcCEvents.CurrentY = tmCEvtCtrls(ilBox).fBoxY + fgBoxInsetY
                pbcCEvents.Print tmCEvtCtrls(ilBox).sShow
            Next ilBox
        End If
        imIgnoreRightMove = False
    End If
End Sub
Private Sub pbcClock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim llTime As Long
    Dim llRadius As Long
    Dim ilFound As Integer
    Dim ilArc As Integer

    If imIgnoreRightMove Then
        Exit Sub
    End If
    imButton = Button
    If (imButton And &H2) = 2 Then  'Right mouse- show values
        imIgnoreRightMove = True
        mCompVector CLng(X), CLng(Y), llTime, llRadius
        'Determine if in an event
        ilFound = -1
        For ilArc = LBound(lmArcTimes, 2) To UBound(lmArcTimes, 2) - 1 Step 1
            If (llTime >= lmArcTimes(1, ilArc) - 5) And (llTime <= lmArcTimes(2, ilArc) + 5) And llRadius <= lmArcTimes(3, ilArc) Then
                If ilFound = -1 Then
                    ilFound = ilArc
                Else
                    If lmArcTimes(3, ilArc) < lmArcTimes(3, ilFound) Then
                        ilFound = ilArc
                    End If
                End If
            End If
        Next ilArc
        If imLastArcPainted <> ilFound Then
            imLastArcPainted = ilFound
            pbcCEvents.Cls
            If ilFound >= 0 Then
                ilFound = lmArcTimes(4, ilFound)
                For ilBox = imLBCEvtCtrls To COMMENTINDEX Step 1
                    tmCEvtCtrls(ilBox).sShow = smShow(ilBox, ilFound)
                Next ilBox
                For ilBox = imLBCEvtCtrls To COMMENTINDEX Step 1
                    pbcCEvents.CurrentX = tmCEvtCtrls(ilBox).fBoxX + fgBoxInsetX
                    pbcCEvents.CurrentY = tmCEvtCtrls(ilBox).fBoxY + fgBoxInsetY
                    pbcCEvents.Print tmCEvtCtrls(ilBox).sShow
                Next ilBox
            End If
        End If
        imIgnoreRightMove = False
    End If
End Sub
Private Sub pbcClock_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (imButton And &H2) = 2 Then  'Right mouse- show values
        pbcCEvents.Cls
        imButton = 0
        imLastArcPainted = -1
        imIgnoreRightMove = False
    End If
End Sub
Private Sub pbcClock_Paint()
    Dim flStartArc As Single
    Dim flEndArc As Single
    Dim llHour As Long
    Dim ilLoop As Integer
    Dim ilEvtNameIndex As Integer
    Dim ilRet As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilCode As Integer
    Dim fldegree As Single
    Dim ilArc As Integer
    Dim ilInset As Integer
    Dim ilUpper As Integer
    Dim ilMatchTimeAdj As Integer
    ilMatchTimeAdj = 1
    ReDim lmArcTimes(0 To 4, 0 To 0) As Long
    pbcClock.ForeColor = BLACK
    pbcClock.FillStyle = 0      'Solid
    pbcClock.FillColor = WHITE
    'Outside circle
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 60
    'Inside circle
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 15
    '3 o'clock arc
    flStartArc = (358 * fmPI) / 180
    flEndArc = (2 * fmPI) / 180
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    '2 o'clock arc
    flStartArc = (28 * fmPI) / 180
    flEndArc = (32 * fmPI) / 180
    pbcClock.DrawWidth = 2
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    '1 o'clock arc
    flStartArc = (58 * fmPI) / 180
    flEndArc = (62 * fmPI) / 180
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    '12 o'clock arc
    flStartArc = (88 * fmPI) / 180
    flEndArc = (92 * fmPI) / 180
    pbcClock.DrawWidth = 1
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    '11 o'clock arc
    flStartArc = (118 * fmPI) / 180
    flEndArc = (122 * fmPI) / 180
    pbcClock.DrawWidth = 2
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    '10 o'clock arc
    flStartArc = (148 * fmPI) / 180
    flEndArc = (152 * fmPI) / 180
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    '9 o'clock arc
    flStartArc = (178 * fmPI) / 180
    flEndArc = (182 * fmPI) / 180
    pbcClock.DrawWidth = 1
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    '8 o'clock arc
    flStartArc = (208 * fmPI) / 180
    flEndArc = (212 * fmPI) / 180
    pbcClock.DrawWidth = 2
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    '7 o'clock arc
    flStartArc = (238 * fmPI) / 180
    flEndArc = (242 * fmPI) / 180
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    '6 o'clock arc
    flStartArc = (268 * fmPI) / 180
    flEndArc = (272 * fmPI) / 180
    pbcClock.DrawWidth = 1
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    '5 o'clock arc
    flStartArc = (298 * fmPI) / 180
    flEndArc = (302 * fmPI) / 180
    pbcClock.DrawWidth = 2
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    '4 o'clock arc
    flStartArc = (328 * fmPI) / 180
    flEndArc = (332 * fmPI) / 180
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 45, , flStartArc, flEndArc
    pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius + 30, , flStartArc, flEndArc
    pbcClock.DrawWidth = 1
    'Ignore Programs; Page Skips; Line Skips
    '   Event               Color   Type
    '   Program             ignore  1
    '   Contract Avail      Green   2
    '   Open BB             Blue    3
    '   Close BB            Blue    5
    '   Cmml Promo          Magneta 6
    '   Feed Avail          Magenta 7
    '   PSA                 Red     8
    '   Promo               Brown   9
    '   Page Skip           ignore  A
    '   Line Skip           ignore  B
    '   Line Skip           ignore  C
    '   Line Skip           ignore  D
    '   Other               Light Yellow
    '
    llHour = hbcHour.Value - 1
    For ilLoop = 1 To UBound(smSave, 2) Step 1
        llStartTime = CLng(gLengthToCurrency(smSave(1, ilLoop)))
        gFindMatch smSave(2, ilLoop), 1, lbcEvtType
        If gLastFound(lbcEvtType) > 0 Then
            ilEvtNameIndex = gLastFound(lbcEvtType)
            slNameCode = tmEvtTypeCode(ilEvtNameIndex - 1).sKey    'lbcEvtTypeCode.List(ilEvtNameIndex - 1)
            ilRet = gParseItem(slNameCode, 3, "\", slCode)
            ilCode = Val(slCode)
            If (ilCode = 2) Or (ilCode = 3) Or (ilCode = 5) Or (ilCode = 6) Or (ilCode = 7) Or (ilCode = 8) Or (ilCode = 9) Then
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                    llEndTime = llStartTime + CLng(gLengthToCurrency(smSave(7, ilLoop)))
                Else
                    llEndTime = llStartTime
                End If
            ElseIf ilCode > 13 Then
                llEndTime = llStartTime + CLng(gLengthToCurrency(smSave(7, ilLoop)))
            End If
            If llEndTime = llStartTime Then
                llEndTime = llStartTime + ilMatchTimeAdj 'This is required so user can click on event
            End If
            If (llHour = (llStartTime \ lm3600)) Or (llHour = (llEndTime \ lm3600)) Or (((llStartTime \ lm3600) < llHour) And ((llEndTime \ lm3600) > llHour)) Then
                If llStartTime < llHour * lm3600 Then
                    llStartTime = llHour * lm3600
                End If
                fldegree = 360 - (llStartTime Mod lm3600) / 10
                If fldegree >= 270 Then
                    fldegree = fldegree - 270
                Else
                    fldegree = fldegree + 90
                End If
                flStartArc = (fldegree * fmPI) / 180
                If flStartArc = 0# Then
                    flStartArc = 0.0001
                End If
                If llEndTime > hbcHour.Value * lm3600 Then
                    llEndTime = hbcHour.Value * lm3600
                End If
                fldegree = 360 - (llEndTime Mod lm3600) / 10
                If fldegree >= 270 Then
                    fldegree = fldegree - 270
                Else
                    fldegree = fldegree + 90
                End If
                flEndArc = (fldegree * fmPI) / 180
                'Test if any arc overlap
                If (ilCode = 2) Or (ilCode = 3) Or (ilCode = 5) Or (ilCode = 6) Or (ilCode = 7) Or (ilCode = 8) Or (ilCode = 9) Or (ilCode > 13) Then
                    ilInset = 0
                    ilUpper = UBound(lmArcTimes, 2)
                    For ilArc = 0 To ilUpper - 1 Step 1
                        If (llEndTime > lmArcTimes(1, ilArc)) And (llStartTime < lmArcTimes(2, ilArc)) Then
                            ilInset = ilInset + 45
                        End If
                    Next ilArc
                    lmArcTimes(1, ilUpper) = llStartTime
                    lmArcTimes(2, ilUpper) = llEndTime
                    lmArcTimes(3, ilUpper) = 1995 - ilInset
                    lmArcTimes(4, ilUpper) = ilLoop
                    ReDim Preserve lmArcTimes(0 To 4, ilUpper + 1) As Long
                End If
                If ilCode = 2 Then  'Contract Avail
                    If llStartTime + ilMatchTimeAdj = llEndTime Then
                        pbcClock.ForeColor = GREEN  'GREEN
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = GREEN
                        pbcClock.Line (lmXCenter, lmYCenter)-Step((lmBaseRadius - ilInset) * Cos(flStartArc), -1 * (lmBaseRadius - ilInset) * Sin(flStartArc))
                    Else
                        pbcClock.ForeColor = BLACK  'GREEN
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = GREEN
                        pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius - ilInset, , -flEndArc, -flStartArc
                    End If
                ElseIf (ilCode = 3) Or (ilCode = 5) Then
                    If llStartTime + ilMatchTimeAdj = llEndTime Then
                        pbcClock.ForeColor = BLUE
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = BLUE
                        pbcClock.Line (lmXCenter, lmYCenter)-Step((lmBaseRadius - ilInset) * Cos(flStartArc), -1 * (lmBaseRadius - ilInset) * Sin(flStartArc))
                    Else
                        pbcClock.ForeColor = BLACK  'BLUE
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = BLUE
                        pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius - ilInset, , -flEndArc, -flStartArc
                    End If
                ElseIf (ilCode = 6) Or (ilCode = 7) Then
                    If llStartTime + ilMatchTimeAdj = llEndTime Then
                        pbcClock.ForeColor = MAGENTA
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = MAGENTA
                        pbcClock.Line (lmXCenter, lmYCenter)-Step((lmBaseRadius - ilInset) * Cos(flStartArc), -1 * (lmBaseRadius - ilInset) * Sin(flStartArc))
                    Else
                        pbcClock.ForeColor = BLACK  'MAGENTA
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = MAGENTA
                        pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius - ilInset, , -flEndArc, -flStartArc
                    End If
                ElseIf (ilCode = 8) Then
                    If llStartTime + ilMatchTimeAdj = llEndTime Then
                        pbcClock.ForeColor = Red
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = Red
                        pbcClock.Line (lmXCenter, lmYCenter)-Step((lmBaseRadius - ilInset) * Cos(flStartArc), -1 * (lmBaseRadius - ilInset) * Sin(flStartArc))
                    Else
                        pbcClock.ForeColor = BLACK  'RED
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = Red
                        pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius - ilInset, , -flEndArc, -flStartArc
                    End If
                ElseIf (ilCode = 9) Then
                    If llStartTime + ilMatchTimeAdj = llEndTime Then
                        pbcClock.ForeColor = CYAN
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = CYAN
                        pbcClock.Line (lmXCenter, lmYCenter)-Step((lmBaseRadius - ilInset) * Cos(flStartArc), -1 * (lmBaseRadius - ilInset) * Sin(flStartArc))
                    Else
                        pbcClock.ForeColor = BLACK  'CYAN
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = CYAN
                        pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius - ilInset, , -flEndArc, -flStartArc
                    End If
                ElseIf ilCode > 13 Then
                    If llStartTime + ilMatchTimeAdj = llEndTime Then
                        pbcClock.ForeColor = Yellow
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = Yellow
                        pbcClock.Line (lmXCenter, lmYCenter)-Step((lmBaseRadius - ilInset) * Cos(flStartArc), -1 * (lmBaseRadius - ilInset) * Sin(flStartArc))
                    Else
                        pbcClock.ForeColor = BLACK  'YELLOW
                        pbcClock.FillStyle = 0      'Solid
                        pbcClock.FillColor = Yellow
                        pbcClock.Circle (lmXCenter, lmYCenter), lmBaseRadius - ilInset, , -flEndArc, -flStartArc
                    End If
                End If
            End If
        End If
    Next ilLoop
End Sub
Private Sub pbcEatTab_GotFocus(Index As Integer)
    If (tmcClick.Enabled) And (Index = 1) Then
        pbcEatTab(0).SetFocus
    End If
End Sub
Private Sub pbcEvents_DragDrop(Source As control, X As Single, Y As Single)
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Dim ilFound As Integer
    Dim ilNewRowNo As Integer
    Dim slStr As String
    Dim slXMid As String

    ilFound = False
    'If (imEvtRowNo < LBound(smSave, 2)) Or (imEvtRowNo > UBound(smSave, 2)) Then
    If (imEvtRowNo < LBONE) Or (imEvtRowNo > UBound(smSave, 2)) Then
        Exit Sub
    End If
    ilCompRow = vbcEvents.LargeChange + 1
    If UBound(smSave, 2) > ilCompRow Then
        ilMaxRow = ilCompRow
        If ilMaxRow + vbcEvents.Value >= UBound(smSave, 2) Then
            ilMaxRow = UBound(smSave, 2) - 1 - vbcEvents.Value
        End If
    Else
        ilMaxRow = UBound(smSave, 2) - 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmEvtCtrls(TIMEINDEX).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmEvtCtrls(TIMEINDEX).fBoxY + tmEvtCtrls(TIMEINDEX).fBoxH)) Then
            If Asc(smSave(9, imEvtRowNo)) >= Asc("2") Then
                If (ilRow + vbcEvents.Value <> imEvtRowNo) And (ilRow + vbcEvents.Value <> imEvtRowNo - 1) Then
                    ilNewRowNo = ilRow + vbcEvents.Value    'Min = 0
                    ilFound = True
                End If
            Else
                If (ilRow + vbcEvents.Value <> imEvtRowNo) And (ilRow + vbcEvents.Value <> imEvtRowNo + 1) Then
                    ilNewRowNo = ilRow + vbcEvents.Value    'Min = 0
                    ilFound = True
                End If
            End If
        End If
    Next ilRow
    If ilFound Then
        If Asc(smSave(9, imEvtRowNo)) >= Asc("2") Then
            mMoveEvt True, ilNewRowNo, imEvtRowNo
            If ilNewRowNo = 1 Then
                If (Asc(smSave(9, ilNewRowNo)) >= Asc("2")) And (Asc(smSave(9, ilNewRowNo)) <= Asc("9")) And (smSave(7, ilNewRowNo) <> "") Then
                    If imTimeRelative Then
                        gAddLengths smSave(1, ilNewRowNo), smSave(7, ilNewRowNo), "3", smSave(1, ilNewRowNo + 1)
                        gSetShow pbcEvents, smSave(1, ilNewRowNo + 1), tmEvtCtrls(TIMEINDEX)
                        smShow(TIMEINDEX, ilNewRowNo + 1) = tmEvtCtrls(TIMEINDEX).sShow
                    Else
                        gAddLengths smSave(1, ilNewRowNo), smSave(7, ilNewRowNo), "3", smSave(1, ilNewRowNo + 1)
                        gAddTimeLength smSpecSave(4), smSave(1, ilNewRowNo + 1), "A", "1", slStr, slXMid
                        gSetShow pbcEvents, slStr, tmEvtCtrls(TIMEINDEX)
                        smShow(TIMEINDEX, ilNewRowNo + 1) = tmEvtCtrls(TIMEINDEX).sShow
                    End If
                End If
            Else
                If (Asc(smSave(1, ilNewRowNo - 1)) >= Asc("2")) And (Asc(smSave(1, ilNewRowNo - 1)) <= Asc("9")) And (smSave(7, ilNewRowNo - 1) <> "") Then
                    If imTimeRelative Then
                        gAddLengths smSave(1, ilNewRowNo - 1), smSave(7, ilNewRowNo - 1), "3", smSave(1, ilNewRowNo)
                        gSetShow pbcEvents, smSave(1, ilNewRowNo), tmEvtCtrls(TIMEINDEX)
                        smShow(TIMEINDEX, ilNewRowNo) = tmEvtCtrls(TIMEINDEX).sShow
                    Else
                        gAddLengths smSave(1, ilNewRowNo - 1), smSave(7, ilNewRowNo - 1), "3", smSave(1, ilNewRowNo)
                        gAddTimeLength smSpecSave(4), smSave(1, ilNewRowNo), "A", "1", slStr, slXMid
                        gSetShow pbcEvents, slStr, tmEvtCtrls(TIMEINDEX)
                        smShow(TIMEINDEX, ilNewRowNo) = tmEvtCtrls(TIMEINDEX).sShow
                    End If
                Else
                    smSave(1, ilNewRowNo) = smSave(1, ilNewRowNo - 1)
                    smShow(1, ilNewRowNo) = smShow(1, ilNewRowNo - 1)
                End If
            End If
        Else
            mMoveEvt False, ilNewRowNo, imEvtRowNo
            smSave(1, ilNewRowNo) = smSave(1, ilNewRowNo + 1)
            smShow(1, ilNewRowNo) = smShow(1, ilNewRowNo + 1)
        End If
        imEvtRowNo = ilNewRowNo
        pbcEvents.Cls
        pbcEvents_Paint
        imLefChg = True
        imEvtBoxNo = 0
        lacEvtFrame.Move 0, tmEvtCtrls(TIMEINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) - 30
        lacEvtFrame.Visible = True
        pbcArrow.Move pbcArrow.Left, plcEvents.Top + tmEvtCtrls(TIMEINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) + 45
        pbcArrow.Visible = True
        pbcArrow.SetFocus
        mSetCommands
    End If
End Sub
Private Sub pbcEvents_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    If State = vbEnter Then
        lacEvtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        Exit Sub
    End If
    If State = vbLeave Then
        lacEvtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        Exit Sub
    End If
    ilCompRow = vbcEvents.LargeChange + 1
    If UBound(smSave, 2) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smSave, 2) - 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmEvtCtrls(TIMEINDEX).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmEvtCtrls(TIMEINDEX).fBoxY + tmEvtCtrls(TIMEINDEX).fBoxH)) Then
            If (ilRow + vbcEvents.Value = imEvtRowNo) Or (ilRow + vbcEvents.Value = imEvtRowNo - 1) Then
                lacEvtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
            Else
                lacEvtFrame.DragIcon = IconTraf!imcIconMove.DragIcon
            End If
            Exit Sub
        End If
    Next ilRow
    lacEvtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
End Sub
Private Sub pbcEvents_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcEvents_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim slStr As String
    Dim ilNewRowNo As Integer
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If vbcEvents.Value = -1 Then
        Exit Sub
    End If
    ilCompRow = vbcEvents.LargeChange + 1
    If UBound(smSave, 2) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smSave, 2)
    End If
    If Not imAllAnsw Then
        mSpecSetShow imSpecBoxNo
        imSpecBoxNo = -1
        If mSpecTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
            mSpecEnableBox imSpecBoxNo
            Exit Sub
        End If
    End If
    imAllAnsw = True
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBEvtCtrls To COMMENTINDEX Step 1
            If (X >= tmEvtCtrls(ilBox).fBoxX) And (X <= (tmEvtCtrls(ilBox).fBoxX + tmEvtCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmEvtCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmEvtCtrls(ilBox).fBoxY + tmEvtCtrls(ilBox).fBoxH)) Then
                    If (imEvtBoxNo = TIMEINDEX) And (edcEvtDropDown.Text <> "") Then
                        slStr = edcEvtDropDown.Text
                        If imTimeRelative Then
                            If Not gValidLength(slStr) Then
                                Beep
                                edcEvtDropDown.SetFocus
                                Exit Sub
                            End If
                        Else
                            If Not gValidTime(slStr) Then
                                Beep
                                edcEvtDropDown.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                    If imEvtBoxNo = LENGTHINDEX Then
                        slStr = edcEvtDropDown.Text
                        If Not gValidLength(slStr) Then
                            Beep
                            edcEvtDropDown.SetFocus
                            Exit Sub
                        End If
                    End If
                    ilRowNo = ilRow + vbcEvents.Value
                    If imEvtBoxNo = EVTTYPEINDEX Then
                        mEvtSetShow imEvtBoxNo
                    End If
                    If (ilBox > EVTTYPEINDEX) And ((smSave(2, ilRowNo) = "") Or (smSave(9, ilRowNo) = "")) Then
                        Beep
                        ilBox = EVTTYPEINDEX
                        mSpecSetShow imSpecBoxNo
                        imSpecBoxNo = -1
                        mEvtSetShow imEvtBoxNo
                        If imEvtRowNo < UBound(smSave, 2) Then
                            ilNewRowNo = mSortByRowTime(imEvtRowNo)
                            If (ilRowNo > imEvtRowNo) And (ilRowNo <= ilNewRowNo) Then
                                ilRowNo = ilRowNo - 1
                            ElseIf (ilRowNo < imEvtRowNo) And (ilRowNo > ilNewRowNo) Then
                                ilRowNo = ilRowNo + 1
                            End If
                        End If
                        imEvtRowNo = ilRowNo
                        imEvtBoxNo = ilBox
                        If (UBound(smSave, 2) = 1) And (smSave(1, 1) = "") Then
                            mInitNewEvent
                        End If
                        mEvtEnableBox ilBox
                        Exit Sub
                    End If
'                    If (ilRowNo = UBound(smSave, 2)) And (imEvtBoxNo = TIMEINDEX) Then
'                        mClearLastEvent
'                    End If
                    imTabDirection = 0  'Set-Left to right
                    If ilBox > EVTNAMEINDEX Then
                        Select Case smSave(9, ilRowNo)
                            Case "1"  'Program
                                If (ilBox = UNITSINDEX) Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                                If ilBox = AVAILINDEX Then
                                    ilBox = EXCL1INDEX
                                End If
                            'Case "2", "3", "4", "5" 'Contract Avail
                            Case "2" 'Contract Avail
                                If ilBox = TRUETIMEINDEX Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                                'If (ilBox = UNITSINDEX) And ((tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M")) Then
                                If (ilBox = UNITSINDEX) And (tgVpf(imVpfIndex).sSSellOut = "M") Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                                If (ilBox = LENGTHINDEX) And ((tgVpf(imVpfIndex).sSSellOut <> "B") And (tgVpf(imVpfIndex).sSSellOut <> "U") And (tgVpf(imVpfIndex).sSSellOut <> "M")) Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                            Case "3", "4", "5"  'Open BB/Floating/Close BB
                                If (ilBox = UNITSINDEX) Or (ilBox = LENGTHINDEX) Or (ilBox = TRUETIMEINDEX) Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                            Case "6"  'Cmml Promo
                                If ilBox = TRUETIMEINDEX Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                                'If (ilBox = UNITSINDEX) And ((tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M")) Then
                                If (ilBox = UNITSINDEX) And (tgVpf(imVpfIndex).sSSellOut = "M") Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                                If (ilBox = LENGTHINDEX) And ((tgVpf(imVpfIndex).sSSellOut <> "B") And (tgVpf(imVpfIndex).sSSellOut <> "U") And (tgVpf(imVpfIndex).sSSellOut <> "M")) Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                            Case "7"  'Feed avail
                                If ilBox = TRUETIMEINDEX Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                                'If (ilBox = UNITSINDEX) And ((tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M")) Then
                                If (ilBox = UNITSINDEX) And (tgVpf(imVpfIndex).sSSellOut = "M") Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                                If (ilBox = LENGTHINDEX) And ((tgVpf(imVpfIndex).sSSellOut <> "B") And (tgVpf(imVpfIndex).sSSellOut <> "U") And (tgVpf(imVpfIndex).sSSellOut <> "M")) Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                            Case "8", "9"  'PSA/Promo (Avail)
                                If ilBox = TRUETIMEINDEX Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                                'If (ilBox = UNITSINDEX) And ((tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M")) Then
                                If (ilBox = UNITSINDEX) And (tgVpf(imVpfIndex).sSSellOut = "M") Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                                If (ilBox = LENGTHINDEX) And ((tgVpf(imVpfIndex).sSSellOut <> "B") And (tgVpf(imVpfIndex).sSSellOut <> "U") And (tgVpf(imVpfIndex).sSSellOut <> "M")) Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                            Case "A", "B", "C", "D"  'Page eject, Line space 1, 2 or 3
                                'If (ilBox = EVTNAMEINDEX) Or (ilBox = AVAILINDEX) Or (ilBox = UNITSINDEX) Or (ilBox = LENGTHINDEX) Or (ilBox = TRUETIMEINDEX) Then
                                If (ilBox = EVTNAMEINDEX) Or (ilBox = UNITSINDEX) Or (ilBox = LENGTHINDEX) Or (ilBox = TRUETIMEINDEX) Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                            Case Else   'Other
                                If (ilBox = AVAILINDEX) Or (ilBox = UNITSINDEX) Or (ilBox = TRUETIMEINDEX) Then
                                    Beep
                                    If imSpecBoxNo > 0 Then
                                        mSpecSetFocus imSpecBoxNo
                                    ElseIf imEvtBoxNo > 0 Then
                                        mEvtSetFocus imEvtBoxNo
                                    End If
                                    Exit Sub
                                End If
                        End Select
                    ElseIf ilBox = EVTNAMEINDEX Then
                        Select Case smSave(9, ilRowNo)
                            Case "A", "B", "C", "D"  'Page eject, Line space 1, 2 or 3
                                Beep
                                If imSpecBoxNo > 0 Then
                                    mSpecSetFocus imSpecBoxNo
                                ElseIf imEvtBoxNo > 0 Then
                                    mEvtSetFocus imEvtBoxNo
                                End If
                                Exit Sub
                        End Select
                    End If
                    mSpecSetShow imSpecBoxNo
                    imSpecBoxNo = -1
                    mEvtSetShow imEvtBoxNo
                    If imEvtRowNo < UBound(smSave, 2) Then
                        ilNewRowNo = mSortByRowTime(imEvtRowNo)
                        If (ilRowNo > imEvtRowNo) And (ilRowNo <= ilNewRowNo) Then
                            ilRowNo = ilRowNo - 1
                        ElseIf (ilRowNo < imEvtRowNo) And (ilRowNo > ilNewRowNo) Then
                            ilRowNo = ilRowNo + 1
                        End If
                    End If
                    imEvtRowNo = ilRowNo
                    imEvtBoxNo = ilBox
                    If (UBound(smSave, 2) = 1) And (smSave(1, 1) = "") Then
                        mInitNewEvent
                    End If
                    mEvtEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    If imSpecBoxNo > 0 Then
        mSpecSetFocus imSpecBoxNo
    ElseIf imEvtBoxNo > 0 Then
        mEvtSetFocus imEvtBoxNo
    End If
End Sub
Private Sub pbcEvents_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim llColor As Long

    mPaintEvtTitle
    If vbcEvents.Value = -1 Then
        Exit Sub
    End If
    ilStartRow = vbcEvents.Value + 1  'Top location
    ilEndRow = vbcEvents.Value + vbcEvents.LargeChange + 1
    If ilEndRow > UBound(smSave, 2) Then
        ilEndRow = UBound(smSave, 2) 'include blank row as it might have data
    End If
    'If Not imTimeRelative Then  'Remove the word (Relative)
    '    gPaintArea pbcEvents, tmEvtCtrls(TIMEINDEX).fBoxX, tmEvtCtrls(TIMEINDEX).fBoxY - fgBoxGridH - 45, tmEvtCtrls(TIMEINDEX).fBoxW - 15, 195, WHITE
    'End If
    llColor = pbcEvents.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        If ilRow = UBound(smSave, 2) Then
            pbcEvents.ForeColor = DARKPURPLE
        Else
            pbcEvents.ForeColor = llColor
        End If
        For ilBox = imLBEvtCtrls To COMMENTINDEX Step 1
'            gPaintArea pbcEvents, tmEvtCtrls(ilBox).fBoxX, tmEvtCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmEvtCtrls(ilBox).fBoxW - 15, tmEvtCtrls(ilBox).fBoxH - 15
            pbcEvents.CurrentX = tmEvtCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcEvents.CurrentY = tmEvtCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            pbcEvents.Print smShow(ilBox, ilRow)
        Next ilBox
    Next ilRow
    pbcEvents.ForeColor = llColor
End Sub
Private Sub pbcLen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcLenInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 5 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcLenInv.Move flX, flY
                    imcLenInv.Visible = True
                    imcLenOutline.Move flX - 15, flY - 15
                    imcLenOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcLen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcLenInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 5 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcLenInv.Move flX, flY
                    imcLenOutline.Move flX - 15, flY - 15
                    imcLenOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "H"
                                Case 2
                                    slKey = "M"
                                Case 3
                                    slKey = "S"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                    End Select
                    Select Case imSpecBoxNo
                        Case SPECLENGTHINDEX
                            imBypassFocus = True    'Don't change select text
                            edcSpecDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcSpecDropDown, slKey
                    End Select
                    Select Case imEvtBoxNo
                        Case TIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcEvtDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcEvtDropDown, slKey
                        Case LENGTHINDEX
                            imBypassFocus = True    'Don't change select text
                            edcEvtDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcEvtDropDown, slKey
                    End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcLibSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim slStr As String
    For ilBox = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        If (X >= tmSpecCtrls(ilBox).fBoxX) And (X <= tmSpecCtrls(ilBox).fBoxX + tmSpecCtrls(ilBox).fBoxW) Then
            If (Y >= tmSpecCtrls(ilBox).fBoxY) And (Y <= tmSpecCtrls(ilBox).fBoxY + tmSpecCtrls(ilBox).fBoxH) Then
                If (imSpecBoxNo = SPECLENGTHINDEX) And (edcEvtDropDown.Text <> "") Then
                    slStr = edcSpecDropDown.Text
                    If Not gValidLength(slStr) Then
                        Beep
                        edcSpecDropDown.SetFocus
                        Exit Sub
                    End If
                End If
                If (imSpecBoxNo = SPECBASETIMEINDEX) And (edcEvtDropDown.Text <> "") Then
                    slStr = edcSpecDropDown.Text
                    If Not gValidTime(slStr) Then
                        Beep
                        edcSpecDropDown.SetFocus
                        Exit Sub
                    End If
                End If
                If (imEvtBoxNo = TIMEINDEX) And (edcEvtDropDown.Text <> "") Then
                    slStr = edcEvtDropDown.Text
                    If imTimeRelative Then
                        If Not gValidLength(slStr) Then
                            Beep
                            edcEvtDropDown.SetFocus
                            Exit Sub
                        End If
                    Else
                        If Not gValidTime(slStr) Then
                            Beep
                            edcEvtDropDown.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
                If (imEvtBoxNo = LENGTHINDEX) And (edcEvtDropDown.Text <> "") Then
                    slStr = edcEvtDropDown.Text
                    If Not gValidLength(slStr) Then
                        Beep
                        edcEvtDropDown.SetFocus
                        Exit Sub
                    End If
                End If
'                If (ilBox = SPECLIBNAMEINDEX) And (imSelectedIndex <> 0) Then
'                    Beep
'                    If imSpecBoxNo > 0 Then
'                        mSpecSetFocus imSpecBoxNo
'                    ElseIf imEvtBoxNo > 0 Then
'                        mEvtSetFocus imEvtBoxNo
'                    End If
'                    Exit Sub
'                End If
                If ilBox = SPECVARINDEX Then
                    Beep
                    If imSpecBoxNo > 0 Then
                        mSpecSetFocus imSpecBoxNo
                    ElseIf imEvtBoxNo > 0 Then
                        mEvtSetFocus imEvtBoxNo
                    End If
                    Exit Sub
                End If
                mEvtSetShow imEvtBoxNo
                imEvtBoxNo = -1
                imEvtRowNo = -1
                pbcArrow.Visible = False
                lacEvtFrame.Visible = False
                mSpecSetShow imSpecBoxNo
                imSpecBoxNo = ilBox
                mSpecEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    If imSpecBoxNo > 0 Then
        mSpecSetFocus imSpecBoxNo
    ElseIf imEvtBoxNo > 0 Then
        mEvtSetFocus imEvtBoxNo
    End If
End Sub
Private Sub pbcLibSpec_Paint()
    Dim ilBox As Integer
    pbcLibSpec.Cls
    For ilBox = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        'Remove variation title if variation is not defined
        If (ilBox = SPECVARINDEX) And (Trim$(tmSpecCtrls(ilBox).sShow) = "") Then
            gPaintArea pbcLibSpec, tmSpecCtrls(ilBox).fBoxX, tmSpecCtrls(ilBox).fBoxY, tmSpecCtrls(ilBox).fBoxW - 15, tmSpecCtrls(ilBox).fBoxH - 15, WHITE
        Else
            pbcLibSpec.CurrentX = tmSpecCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcLibSpec.CurrentY = tmSpecCtrls(ilBox).fBoxY + fgBoxInsetY
            pbcLibSpec.Print tmSpecCtrls(ilBox).sShow
        End If
    Next ilBox
End Sub
Private Sub pbcLibType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcLibType_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(" ") Then
        If cbcType.ListCount > 1 Then
            If imSpecSave(1) = 0 Then
                imSpecSave(1) = 1
                imLtfChg = True
                tmSpecCtrls(SPECLIBTYPEINDEX).iChg = True
                'If imAllowDate Then
                '    cmcDates.Enabled = True
                'End If
            ElseIf imSpecSave(1) = 1 Then
                imSpecSave(1) = 2
                imLtfChg = True
                tmSpecCtrls(SPECLIBTYPEINDEX).iChg = True
                'If imAllowDate Then
                '    cmcDates.Enabled = True
                'End If
            ElseIf imSpecSave(1) = 2 Then
                imSpecSave(1) = 3
                imLtfChg = True
                tmSpecCtrls(SPECLIBTYPEINDEX).iChg = True
                'cmcDates.Enabled = False
            Else
                imSpecSave(1) = 0
                imLtfChg = True
                tmSpecCtrls(SPECLIBTYPEINDEX).iChg = True
                'If imAllowDate Then
                '    cmcDates.Enabled = True
                'End If
            End If
            pbcLibType_Paint
        End If
    End If
End Sub
Private Sub pbcLibType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cbcType.ListCount > 1 Then
        If imSpecSave(1) = 0 Then
            imSpecSave(1) = 1
            imLtfChg = True
            tmSpecCtrls(SPECLIBTYPEINDEX).iChg = True
            'If imAllowDate Then
            '    cmcDates.Enabled = True
            'End If
        ElseIf imSpecSave(1) = 1 Then
            imSpecSave(1) = 2
            imLtfChg = True
            tmSpecCtrls(SPECLIBTYPEINDEX).iChg = True
            'If imAllowDate Then
            '    cmcDates.Enabled = True
            'End If
        ElseIf imSpecSave(1) = 2 Then
            imSpecSave(1) = 3
            imLtfChg = True
            tmSpecCtrls(SPECLIBTYPEINDEX).iChg = True
            'cmcDates.Enabled = False
        Else
            imSpecSave(1) = 0
            imLtfChg = True
            tmSpecCtrls(SPECLIBTYPEINDEX).iChg = True
            'If imAllowDate Then
            '    cmcDates.Enabled = True
            'End If
        End If
        pbcLibType_Paint
    End If
End Sub
Private Sub pbcLibType_Paint()
    pbcLibType.Cls
    pbcLibType.CurrentX = fgBoxInsetX
    pbcLibType.CurrentY = -15 'fgBoxInsetY
    If imSpecSave(1) = 0 Then
        pbcLibType.Print "Regular"
    ElseIf imSpecSave(1) = 1 Then
        pbcLibType.Print "Special"
    ElseIf imSpecSave(1) = 2 Then
        pbcLibType.Print "Sport"
    ElseIf imSpecSave(1) = 3 Then
        pbcLibType.Print "Std Format"
    End If
End Sub
Private Sub pbcSpecSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSpecSTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = -1  'Set-Right to left
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    If (imSpecBoxNo >= imLBSpecCtrls) And (imSpecBoxNo <= UBound(tmSpecCtrls)) Then
        If (imSpecBoxNo <> SPECLIBNAMEINDEX) Or (Not cbcSelect.Enabled) Then
            If mSpecTestFields(imSpecBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mSpecEnableBox imSpecBoxNo
                Exit Sub
            End If
        End If
    End If
    Select Case imSpecBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            If (imSelectedIndex = 0) Then 'And (cbcSelect.Text = "[New]") Then
                ilBox = SPECLIBTYPEINDEX
            Else
                ilBox = SPECBASETIMEINDEX
            End If
        Case SPECLIBTYPEINDEX 'Type (first control within header)
            mSpecSetShow imSpecBoxNo
            imSpecBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            ilBox = SPECLIBNAMEINDEX
        Case SPECBASETIMEINDEX
            If imSelectedIndex <> 0 Then
                mSpecSetShow imSpecBoxNo
                imSpecBoxNo = -1
                If cbcSelect.Enabled Then
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                ilBox = SPECBASETIMEINDEX
            Else
                ilBox = SPECLIBNAMEINDEX
            End If
        Case Else
            ilBox = imSpecBoxNo - 1
    End Select
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = ilBox
    mSpecEnableBox ilBox
End Sub
Private Sub pbcSpecSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcSpecTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSpecTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = 0  'Set-Left to right
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    mSpecSetShow imSpecBoxNo
    If (imSpecBoxNo >= imLBSpecCtrls) And (imSpecBoxNo <= UBound(tmSpecCtrls)) Then
        If mSpecTestFields(imSpecBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            imDirProcess = -1
            mSpecEnableBox imSpecBoxNo
            Exit Sub
        End If
    End If
    If imDirProcess >= 0 Then
        mSpecDirection imDirProcess
        imDirProcess = -1
        Exit Sub
    End If
    Select Case imSpecBoxNo
        Case -1 'Shift tab from button
            imTabDirection = -1  'Set-Right to left
            ilBox = SPECBASETIMEINDEX
        Case SPECBASETIMEINDEX    'last control
            imSpecBoxNo = -1
            If mSpecTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
                Beep
                mSpecEnableBox imSpecBoxNo
                Exit Sub
            End If
            imAllAnsw = True
            'If cmcDates.Enabled Then
            '    cmcDates.SetFocus
            'Else
                If pbcSTab.Enabled Then
                    pbcSTab.SetFocus
                Else
                    pbcSpecSTab.SetFocus
                End If
            'End If
            Exit Sub
        Case SPECLIBNAMEINDEX
            ilBox = SPECLENGTHINDEX
        Case Else
            ilBox = imSpecBoxNo + 1
    End Select
    imSpecBoxNo = ilBox
    mSpecEnableBox ilBox
End Sub
Private Sub pbcSpecTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    Dim ilNewRowNo As Integer
    Dim ilRowNo As Integer
    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = -1  'Set-Right to left
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    If imEvtBoxNo = EVTTYPEINDEX Then
        If mEvtTypeBranch() Then
            Exit Sub
        End If
    End If
    If imEvtBoxNo = EVTNAMEINDEX Then
        If mEvtNameBranch() Then
            Exit Sub
        End If
    End If
    If imEvtBoxNo = AVAILINDEX Then
        If mEvtAvailBranch() Then
            Exit Sub
        End If
    End If
    If imEvtBoxNo = EXCL1INDEX Then
        If mExclBranch(0) Then
            Exit Sub
        End If
    End If
    If imEvtBoxNo = EXCL2INDEX Then
        If mExclBranch(1) Then
            Exit Sub
        End If
    End If
    Select Case imEvtBoxNo
        Case -1 'Tab from control prior to form area
            If (UBound(smSave, 2) = 1) And (smSave(1, 1) = "") Then
                imTabDirection = 0  'Set-Left to right
                imEvtRowNo = 1
                mInitNewEvent
            Else
                If UBound(smSave, 2) < vbcEvents.LargeChange Then 'was <=
                    vbcEvents.Max = LBONE - 1 'LBound(smSave, 2) - 1
                Else
                    vbcEvents.Max = UBound(smSave, 2) - vbcEvents.LargeChange - 1
                End If
                imEvtRowNo = 0
'                Do While imEvtRowNo < UBound(smSave, 2) - 1
'                    imEvtRowNo = imEvtRowNo + 1
'                    If imEvtRowNo > vbcEvents.Value + vbcEvents.LargeChange + 1 Then
'                        imSettingValue = True
'                        vbcEvents.Value = vbcEvents.Value + 1
'                    End If
'                Loop
                imEvtRowNo = imEvtRowNo + 1
                If imEvtRowNo >= UBound(smSave, 2) Then
                    mInitNewEvent
                End If
                imSettingValue = True
                vbcEvents.Value = vbcEvents.Min
                imSettingValue = False
'                If imEvtRowNo > vbcEvents.Value + vbcEvents.LargeChange + 1 Then
'                    imSettingValue = True
'                    vbcEvents.Value = vbcEvents.Value + 1
 '               End If
            End If
            ilBox = TIMEINDEX
            imEvtBoxNo = ilBox
            mEvtEnableBox ilBox
            Exit Sub
        Case TIMEINDEX, 0 'Name (first control within header)
            If imEvtBoxNo = TIMEINDEX Then
                If edcEvtDropDown.Text <> "" Then
                    slStr = edcEvtDropDown.Text
                    If imTimeRelative Then
                        If Not gValidLength(slStr) Then
                            Beep
                            edcEvtDropDown.SetFocus
                            Exit Sub
                        End If
                    Else
                        If Not gValidTime(slStr) Then
                            Beep
                            edcEvtDropDown.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
                mEvtSetShow imEvtBoxNo
            End If
            ilBox = COMMENTINDEX
            If imEvtRowNo <= 1 Then
                imEvtBoxNo = -1
                imEvtRowNo = -1
                pbcSpecTab.SetFocus
                Exit Sub
            End If
            imEvtRowNo = imEvtRowNo - 1
            If imEvtRowNo < vbcEvents.Value + 1 Then
                imSettingValue = True
                vbcEvents.Value = vbcEvents.Value - 1
                imSettingValue = False
            End If
            ilRowNo = imEvtRowNo
            ilNewRowNo = mSortByRowTime(imEvtRowNo + 1)
            If (ilRowNo > imEvtRowNo + 1) And (ilRowNo <= ilNewRowNo) Then
                imEvtRowNo = ilRowNo - 1
            ElseIf (ilRowNo < imEvtRowNo + 1) And (ilRowNo > ilNewRowNo) Then
                imEvtRowNo = ilRowNo + 1
            End If
            imEvtBoxNo = ilBox
            mEvtEnableBox ilBox
            Exit Sub
        Case EVTIDINDEX 'COMMENTINDEX
            Select Case smSave(9, imEvtRowNo)
                Case "1"  'Program
                    ilBox = TRUETIMEINDEX
                'Case "2", "3", "4", "5" 'Contract avail
                Case "2" 'Contract avail
                    If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                        ilBox = LENGTHINDEX
                    Else
                        ilBox = UNITSINDEX
                    End If
                Case "3", "4", "5"   'BB
                    ilBox = AVAILINDEX
                Case "6", "7", "8", "9"
                    If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                        ilBox = LENGTHINDEX
                    Else
                        ilBox = UNITSINDEX
                    End If
                Case "A"
                    ilBox = AVAILINDEX
                'Case "A", "B", "C", "D"
                Case "A", "B", "C", "D"
                    ilBox = EVTTYPEINDEX
                Case Else
                    ilBox = LENGTHINDEX
            End Select
        Case LENGTHINDEX
            slStr = edcEvtDropDown.Text
            If (smSave(9, imEvtRowNo) = "2") And (slStr <> "") Then '2=Avail
                If InStr(1, UCase(slStr), "H", vbTextCompare) <= 0 Then
                    If InStr(1, UCase(slStr), "M", vbTextCompare) <= 0 Then
                        If InStr(1, UCase(slStr), "S", vbTextCompare) <= 0 Then
                            Beep
                            edcEvtDropDown.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            End If
            If Not gValidLength(slStr) Then
                Beep
                edcEvtDropDown.SetFocus
                Exit Sub
            End If
            Select Case smSave(9, imEvtRowNo)
                Case "1"  'Program
                    ilBox = EXCL2INDEX
                Case "2"  'Contract avail
                    If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                        ilBox = UNITSINDEX
                    Else    'tgVpf(imVpfIndex).sSSellOut = "M"
                        ilBox = AVAILINDEX
                    End If
                Case "3", "4", "5"  'BB
                Case "6", "7", "8", "9"
                    If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                        ilBox = UNITSINDEX
                    Else    'tgVpf(imVpfIndex).sSSellOut = "M"
                        ilBox = AVAILINDEX
                    End If
                Case "A", "B", "C", "D"
                Case Else
                    ilBox = EVTNAMEINDEX
            End Select
        Case AVAILINDEX
            If smSave(9, imEvtRowNo) = "A" Then
                ilBox = EVTTYPEINDEX
            Else
                ilBox = imEvtBoxNo - 1
            End If
        Case Else
            ilBox = imEvtBoxNo - 1
    End Select
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = ilBox
    mEvtEnableBox ilBox
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
    Dim ilNewRowNo As Integer
    Dim ilRowNo As Integer

    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = 0  'Set-Left to right
    If imDirProcess >= 0 Then
        mEvtDirection imDirProcess
        imDirProcess = -1
        Exit Sub
    End If
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    If imEvtBoxNo = EVTTYPEINDEX Then
        If mEvtTypeBranch() Then
            Exit Sub
        End If
    End If
    If imEvtBoxNo = EVTNAMEINDEX Then
        If mEvtNameBranch() Then
            Exit Sub
        End If
    End If
    If imEvtBoxNo = AVAILINDEX Then
        If mEvtAvailBranch() Then
            Exit Sub
        End If
    End If
    If imEvtBoxNo = EXCL1INDEX Then
        If mExclBranch(0) Then
            Exit Sub
        End If
    End If
    If imEvtBoxNo = EXCL2INDEX Then
        If mExclBranch(1) Then
            Exit Sub
        End If
    End If
    Select Case imEvtBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            imEvtRowNo = UBound(smSave, 2)
            imSettingValue = True
            If imEvtRowNo <= vbcEvents.LargeChange Then
                vbcEvents.Value = 0
            Else
                vbcEvents.Value = imEvtRowNo - vbcEvents.LargeChange - 1
            End If
            imSettingValue = False
            ilBox = 1
        Case TIMEINDEX
            If (imEvtRowNo >= UBound(smSave, 2)) And (edcEvtDropDown.Text = "") Then
                mEvtSetShow imEvtBoxNo
                For ilLoop = TIMEINDEX To COMMENTINDEX Step 1
                    slStr = ""
                    gSetShow pbcEvents, slStr, tmEvtCtrls(ilLoop)
                    smShow(ilLoop, imEvtRowNo) = tmEvtCtrls(ilLoop).sShow
                Next ilLoop
                imEvtBoxNo = -1
                imEvtRowNo = -1
                pbcEvents_Paint
                If cmcUpdate.Enabled Then
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
            End If
            slStr = edcEvtDropDown.Text
            If imTimeRelative Then
                If Not gValidLength(slStr) Then
                    Beep
                    edcEvtDropDown.SetFocus
                    Exit Sub
                End If
            Else
                If Not gValidTime(slStr) Then
                    Beep
                    edcEvtDropDown.SetFocus
                    Exit Sub
                End If
            End If
            ilBox = imEvtBoxNo + 1
        Case EVTTYPEINDEX
            mEvtSetShow imEvtBoxNo
            If (smSave(2, imEvtRowNo) = "") Or (smSave(9, imEvtRowNo) = "") Then
                ilBox = EVTTYPEINDEX
            Else
                'If (Asc(smSave(9, imEvtRowNo)) >= Asc("A")) And (Asc(smSave(9, imEvtRowNo)) <= Asc("D")) Then
                If (Asc(smSave(9, imEvtRowNo)) >= Asc("B")) And (Asc(smSave(9, imEvtRowNo)) <= Asc("D")) Then
                    ilBox = EVTIDINDEX  'COMMENTINDEX
                ElseIf (Asc(smSave(9, imEvtRowNo)) = Asc("A")) Then
                    ilBox = AVAILINDEX
                Else
                    ilBox = EVTNAMEINDEX
                End If
            End If
            imEvtBoxNo = ilBox
            mEvtEnableBox ilBox
            Exit Sub
        Case EVTNAMEINDEX
            If smSave(9, imEvtRowNo) = "1" Then 'Program
                ilBox = EXCL1INDEX
            ElseIf smSave(9, imEvtRowNo) = "Y" Then
                ilBox = LENGTHINDEX
            Else
                ilBox = AVAILINDEX
            End If
        Case AVAILINDEX
            'No units associated with BB
'            If (Asc(smSave(9, imEvtRowNo)) >= Asc("3")) And (Asc(smSave(9, imEvtRowNo)) <= Asc("5")) Then
'                ilBox = COMMENTINDEX
'            Else
            'If (tgVpf(imVpfIndex).sSSellOut <> "U") And (tgVpf(imVpfIndex).sSSellOut <> "M") Then
            If (Asc(smSave(9, imEvtRowNo)) = Asc("A")) Then
                ilBox = EVTIDINDEX  'COMMENTINDEX
            ElseIf (smSave(9, imEvtRowNo) = "3") Or (smSave(9, imEvtRowNo) = "4") Or (smSave(9, imEvtRowNo) = "5") Then
                ilBox = EVTIDINDEX  'COMMENTINDEX
            Else
                If (tgVpf(imVpfIndex).sSSellOut <> "M") Then
                    ilBox = UNITSINDEX
                Else
                    ilBox = LENGTHINDEX
                End If
            End If
'            End If
        Case EXCL2INDEX
            ilBox = LENGTHINDEX
        Case UNITSINDEX
            If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                ilBox = LENGTHINDEX
            Else
                ilBox = EVTIDINDEX  'COMMENTINDEX
            End If
        Case LENGTHINDEX
            slStr = edcEvtDropDown.Text
            If (smSave(9, imEvtRowNo) = "2") And (slStr <> "") Then '2=Avail
                If InStr(1, UCase(slStr), "H", vbTextCompare) <= 0 Then
                    If InStr(1, UCase(slStr), "M", vbTextCompare) <= 0 Then
                        If InStr(1, UCase(slStr), "S", vbTextCompare) <= 0 Then
                            Beep
                            edcEvtDropDown.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            End If
            If Not gValidLength(slStr) Then
                Beep
                edcEvtDropDown.SetFocus
                Exit Sub
            End If
            If smSave(9, imEvtRowNo) <> "1" Then
                ilBox = EVTIDINDEX  'COMMENTINDEX
            Else
                ilBox = TRUETIMEINDEX
            End If
        Case COMMENTINDEX 'Last control within header
            mEvtSetShow imEvtBoxNo
            If mEvtTestSaveFields(imEvtRowNo) = NO Then
                mEvtEnableBox imEvtBoxNo
                Exit Sub
            End If
            If imEvtRowNo >= UBound(smSave, 2) Then
                If smSave(10, imEvtRowNo) = "" Then
                    imLefChg = True
                End If
                smSave(10, imEvtRowNo) = "0"    'Record position
                ReDim Preserve smSave(0 To 11, 0 To imEvtRowNo + 1) As String
                ReDim Preserve imSave(0 To 1, 0 To imEvtRowNo + 1) As Integer
                ReDim Preserve smShow(0 To COMMENTINDEX, 0 To imEvtRowNo + 1) As String
            End If
            If imEvtRowNo >= UBound(smSave, 2) - 1 Then
                imEvtRowNo = imEvtRowNo + 1
                mInitNewEvent
                If UBound(smSave, 2) < vbcEvents.LargeChange Then 'was <=
                    vbcEvents.Max = LBONE - 1   'LBound(smSave, 2) - 1
                Else
                    vbcEvents.Max = UBound(smSave, 2) - vbcEvents.LargeChange - 1
                End If
            Else
                imEvtRowNo = imEvtRowNo + 1
            End If
            If imEvtRowNo > vbcEvents.Value + vbcEvents.LargeChange + 1 Then
                imSettingValue = True
                vbcEvents.Value = vbcEvents.Value + 1
                imSettingValue = False
            End If
            ilRowNo = imEvtRowNo
            ilNewRowNo = mSortByRowTime(imEvtRowNo - 1)
            If (ilRowNo > imEvtRowNo - 1) And (ilRowNo <= ilNewRowNo) Then
                imEvtRowNo = ilRowNo - 1
            ElseIf (ilRowNo < imEvtRowNo - 1) And (ilRowNo > ilNewRowNo) Then
                imEvtRowNo = ilRowNo + 1
            End If
            If imEvtRowNo >= UBound(smSave, 2) Then
                imEvtBoxNo = 0
                mSetCommands
                lacEvtFrame.Move 0, tmEvtCtrls(TIMEINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) - 30
                lacEvtFrame.Visible = True
                pbcArrow.Move pbcArrow.Left, plcEvents.Top + tmEvtCtrls(TIMEINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) + 45
                pbcArrow.Visible = True
                pbcArrow.SetFocus
                Exit Sub
            Else
                ilBox = TIMEINDEX
            End If
            imEvtBoxNo = ilBox
            mEvtEnableBox ilBox
            Exit Sub
        Case Else
            ilBox = imEvtBoxNo + 1
    End Select
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = ilBox
    mEvtEnableBox ilBox
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
                    Select Case imSpecBoxNo
                        Case SPECBASETIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcSpecDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcSpecDropDown, slKey
                    End Select
                    Select Case imEvtBoxNo
                        Case TIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcEvtDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcEvtDropDown, slKey
                    End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcTrueTime_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcTrueTime_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        If imSave(1, imEvtRowNo) <> 0 Then
            imLefChg = True
        End If
        imSave(1, imEvtRowNo) = 0
        pbcTrueTime_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imSave(1, imEvtRowNo) <> 1 Then
            imLefChg = True
        End If
        imSave(1, imEvtRowNo) = 1
        pbcTrueTime_Paint
    End If
End Sub
Private Sub pbcTrueTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imSave(1, imEvtRowNo) = 0 Then
        imLefChg = True
        imSave(1, imEvtRowNo) = 1
    Else
        imLefChg = True
        imSave(1, imEvtRowNo) = 0
    End If
    pbcTrueTime_Paint
End Sub
Private Sub pbcTrueTime_Paint()
    pbcTrueTime.Cls
    pbcTrueTime.CurrentX = fgBoxInsetX
    pbcTrueTime.CurrentY = 0 'fgBoxInsetY
    If imSave(1, imEvtRowNo) = 0 Then
        pbcTrueTime.Print "Yes"
    Else
        pbcTrueTime.Print "No"
    End If
End Sub
Private Sub plcEvents_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub plcHour_Paint()
    plcHour.CurrentX = 0
    plcHour.CurrentY = 0
    plcHour.Print smHourCaption
End Sub

Private Sub plcLibSpec_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcLibSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
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
Private Sub rbcView_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcView(Index).Value
    'End of coded added
    Dim llLength As Long
    If Value Then
        If Index = 0 Then   'Tabular
            plcEvents.height = pbcEvents.height + 2 * fgBevelY
            plcEvents.Width = pbcEvents.Width + vbcEvents.Width + 2 * fgBevelX + 15
            pbcClock.Visible = False
            pbcEvents.Visible = True
            vbcEvents.Visible = True
        Else
            plcEvents.height = pbcClock.height + 2 * fgBevelY
            plcEvents.Width = pbcClock.Width + 2 * fgBevelX
            pbcEvents.Visible = False
            vbcEvents.Visible = False
            pbcClock.Visible = True
            If smSpecSave(3) <> "" Then
                llLength = CLng(gLengthToCurrency(smSpecSave(3)))
                hbcHour.Min = 1
                If (llLength Mod lm3600) = 0 Then
                    hbcHour.Max = llLength \ lm3600
                Else
                    hbcHour.Max = llLength \ lm3600 + 1
                End If
                hbcHour.Value = 1
            Else
                hbcHour.Min = 1
                hbcHour.Max = 1
                hbcHour.Value = 1
            End If
        End If
    End If
End Sub
Private Sub rbcView_GotFocus(Index As Integer)
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    pbcArrow.Visible = False
    lacEvtFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imEvtBoxNo
        Case EVTTYPEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcEvtType, edcEvtDropDown, imChgMode, imLbcArrowSetting
        Case EVTNAMEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcEvtName(imEvtNameIndex), edcEvtDropDown, imChgMode, imLbcArrowSetting
        Case AVAILINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcEvtAvail, edcEvtDropDown, imChgMode, imLbcArrowSetting
        Case EXCL1INDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcExcl(0), edcEvtDropDown, imChgMode, imLbcArrowSetting
        Case EXCL2INDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcExcl(1), edcEvtDropDown, imChgMode, imLbcArrowSetting
    End Select
    pbcEatTab(1).Enabled = False
    pbcEatTab(0).Enabled = False
End Sub
Private Sub tmcDrag_Timer()
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            ilCompRow = vbcEvents.LargeChange + 1
            If UBound(smSave, 2) > ilCompRow Then
                ilMaxRow = ilCompRow
            Else
                ilMaxRow = UBound(smSave, 2)
            End If
            For ilRow = 1 To ilMaxRow Step 1
                If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmEvtCtrls(TIMEINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmEvtCtrls(TIMEINDEX).fBoxY + tmEvtCtrls(TIMEINDEX).fBoxH)) Then
                    mSpecSetShow imSpecBoxNo
                    imSpecBoxNo = -1
                    mEvtSetShow imEvtBoxNo
                    imEvtBoxNo = -1
                    imEvtRowNo = -1
                    imEvtRowNo = ilRow + vbcEvents.Value
                    lacEvtFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                    lacEvtFrame.Move 0, tmEvtCtrls(TIMEINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) - 30
                    'If gInvertArea call then remove visible setting
                    lacEvtFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcEvents.Top + tmEvtCtrls(TIMEINDEX).fBoxY + (imEvtRowNo - vbcEvents.Value - 1) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    lacEvtFrame.Drag vbBeginDrag
                    lacEvtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                    Exit Sub
                End If
            Next ilRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub vbcEvents_Change()
    If imSettingValue Then
        pbcEvents.Cls
        pbcEvents_Paint
        imSettingValue = False
    Else
        mEvtSetShow imEvtBoxNo
        imEvtBoxNo = -1
        imEvtRowNo = -1
        pbcEvents.Cls
        pbcEvents_Paint
        'If (igWinStatus(PROGRAMMINGJOB) = 2) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        '    mEvtEnableBox imEvtBoxNo
        'End If
    End If
End Sub
Private Sub vbcEvents_GotFocus()
    mEvtSetShow imEvtBoxNo
    imEvtBoxNo = -1
    imEvtRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Library Events- " & smVehName
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
Private Sub mPaintEvtTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer

    llColor = pbcEvents.ForeColor
    slFontName = pbcEvents.FontName
    flFontSize = pbcEvents.FontSize
    ilFillStyle = pbcEvents.FillStyle
    llFillColor = pbcEvents.FillColor
    pbcEvents.ForeColor = BLUE
    pbcEvents.FontBold = False
    pbcEvents.FontSize = 7
    pbcEvents.FontName = "Arial"
    pbcEvents.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    ilHalfY = tmEvtCtrls(TIMEINDEX).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop
    pbcEvents.Line (tmEvtCtrls(TIMEINDEX).fBoxX - 15, 15)-Step(tmEvtCtrls(TIMEINDEX).fBoxW + 15, tmEvtCtrls(TIMEINDEX).fBoxY - 30), BLUE, B
    pbcEvents.CurrentX = tmEvtCtrls(TIMEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcEvents.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcEvents.Print "Time"
    If imTimeRelative Then
        pbcEvents.CurrentX = tmEvtCtrls(TIMEINDEX).fBoxX + 15  'fgBoxInsetX
        pbcEvents.CurrentY = ilHalfY + 15
        pbcEvents.Print "(Relative)"
    End If
    pbcEvents.Line (tmEvtCtrls(EVTTYPEINDEX).fBoxX - 15, 15)-Step(tmEvtCtrls(EVTTYPEINDEX).fBoxW + 15, tmEvtCtrls(EVTTYPEINDEX).fBoxY - 30), BLUE, B
    pbcEvents.CurrentX = tmEvtCtrls(EVTTYPEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcEvents.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcEvents.Print "Event Type"
    pbcEvents.Line (tmEvtCtrls(EVTNAMEINDEX).fBoxX - 15, 15)-Step(tmEvtCtrls(EVTNAMEINDEX).fBoxW + 15, tmEvtCtrls(EVTNAMEINDEX).fBoxY - 30), BLUE, B
    pbcEvents.CurrentX = tmEvtCtrls(EVTNAMEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcEvents.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcEvents.Print "Event Name"
    pbcEvents.Line (tmEvtCtrls(AVAILINDEX).fBoxX - 15, 15)-Step(tmEvtCtrls(AVAILINDEX).fBoxW + 15, tmEvtCtrls(AVAILINDEX).fBoxY - 30), BLUE, B
    pbcEvents.CurrentX = tmEvtCtrls(AVAILINDEX).fBoxX + 15  'fgBoxInsetX
    pbcEvents.CurrentY = 15
    pbcEvents.Print "Program Exclusions"
    pbcEvents.CurrentX = tmEvtCtrls(AVAILINDEX).fBoxX + 15  'fgBoxInsetX
    pbcEvents.CurrentY = ilHalfY + 15
    pbcEvents.Print "or Avail Name"
    pbcEvents.Line (tmEvtCtrls(UNITSINDEX).fBoxX - 15, 15)-Step(tmEvtCtrls(UNITSINDEX).fBoxW + 15, tmEvtCtrls(UNITSINDEX).fBoxY - 30), BLUE, B
    pbcEvents.CurrentX = tmEvtCtrls(UNITSINDEX).fBoxX + 15  'fgBoxInsetX
    pbcEvents.CurrentY = 15
    pbcEvents.Print "Units"
    pbcEvents.Line (tmEvtCtrls(LENGTHINDEX).fBoxX - 15, 15)-Step(tmEvtCtrls(LENGTHINDEX).fBoxW + 15, tmEvtCtrls(LENGTHINDEX).fBoxY - 30), BLUE, B
    pbcEvents.CurrentX = tmEvtCtrls(LENGTHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcEvents.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcEvents.Print "Length"
    pbcEvents.Line (tmEvtCtrls(TRUETIMEINDEX).fBoxX - 15, 15)-Step(tmEvtCtrls(TRUETIMEINDEX).fBoxW + 15, tmEvtCtrls(TRUETIMEINDEX).fBoxY - 30), BLUE, B
    pbcEvents.CurrentX = tmEvtCtrls(TRUETIMEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcEvents.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcEvents.Print "True"
    pbcEvents.CurrentX = tmEvtCtrls(TRUETIMEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcEvents.CurrentY = ilHalfY + 15
    pbcEvents.Print "Time"
    pbcEvents.Line (tmEvtCtrls(EVTIDINDEX).fBoxX - 15, 15)-Step(tmEvtCtrls(EVTIDINDEX).fBoxW + 15, tmEvtCtrls(EVTIDINDEX).fBoxY - 30), BLUE, B
    pbcEvents.CurrentX = tmEvtCtrls(EVTIDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcEvents.CurrentY = 15
    pbcEvents.Print "Event ID"
    pbcEvents.Line (tmEvtCtrls(COMMENTINDEX).fBoxX - 15, 15)-Step(tmEvtCtrls(COMMENTINDEX).fBoxW + 15, tmEvtCtrls(COMMENTINDEX).fBoxY - 30), BLUE, B
    pbcEvents.CurrentX = tmEvtCtrls(COMMENTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcEvents.CurrentY = 15
    pbcEvents.Print "Comment"

    ilLineCount = 0
    llTop = tmEvtCtrls(1).fBoxY
    Do
        For ilLoop = imLBEvtCtrls To UBound(tmEvtCtrls) Step 1
            pbcEvents.Line (tmEvtCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmEvtCtrls(ilLoop).fBoxW + 15, tmEvtCtrls(ilLoop).fBoxH + 15), BLUE, B
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmEvtCtrls(1).fBoxH + 15
    Loop While llTop + tmEvtCtrls(1).fBoxH < pbcEvents.height
    vbcEvents.LargeChange = ilLineCount - 1
    pbcEvents.FontSize = flFontSize
    pbcEvents.FontName = slFontName
    pbcEvents.FontSize = flFontSize
    pbcEvents.ForeColor = llColor
    pbcEvents.FontBold = True
End Sub
'only if exporting AudioVault PRS will a ";" be allowed.  This is a separator in the commnets field to create one
'export event for each text separated by ";"
Function mCheckKeyAscii(ilKeyAscii As Integer) As Integer
    '92=Backslash (\); 94=Caret (^); 124=Verical Bar (|); 91=Sq Bracket ([); 93=Sq Bracket (]); 59=Semi-colon(;)
    'If (ilKeyAscii < 32) Or (ilKeyAscii > 126) Or (ilKeyAscii = 92) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Then
    'Allow LF, CR and Backspace.  Just remove test for < 32
    If (ilKeyAscii > 126) Or (ilKeyAscii = 92) Or (ilKeyAscii = 94) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Then     'Or (ilKeyAscii = 59) Then
        Beep
        mCheckKeyAscii = False
        Exit Function
    End If
    
    'if Audio Vault PRS then allow a ";", otherwise disallow
    'must also be a special comment named "RPS Comment", EtfType = "Y"
    If (((Asc(tgSpf.sAutoType2) And AUDIOVAULTRPS) = AUDIOVAULTRPS)) And (smSave(2, imEvtRowNo) = "RPS Comment") And ((smSave(9, imEvtRowNo) > "D") And (smSave(9, imEvtRowNo) <= "Z")) Then
        'Audio Vault RPS comment allowed to have ";" for multiple event comments on export
        mCheckKeyAscii = True
    Else            'illegal character for this event (semi colon ; or vertical bar |)
        If ilKeyAscii = 59 Or (ilKeyAscii = 124) Then
            Beep
            mCheckKeyAscii = False
            Exit Function
        End If
    End If

    mCheckKeyAscii = True
End Function
'8298 10933
'Private Function mEventIdTest(slLine As String) As Boolean
'    Dim blRet As Boolean
'    Dim i As Integer
'    Dim ilCountCommas As Integer
'    Dim ilCountTildes As Integer
'    'returns true unless IS an 'event' vehicle AND then fails test
'    blRet = True
'    ilCountCommas = 0
'    ilCountTildes = 0
'    If bmTestEventID Then
'        If UCase(slLine) <> "-XDS" Then
'            For i = 1 To Len(slLine)
'                If Mid$(slLine, i, 1) = ":" Then
'                    ilCountCommas = ilCountCommas + 1
'                End If
'                If Mid$(slLine, i, 1) = "~" Then
'                    ilCountTildes = ilCountTildes + 1
'                End If
'            Next
'            If ilCountCommas <> ilCountTildes + 1 Then
'                blRet = False
'            End If
'        End If
'    End If
'    mEventIdTest = blRet
'End Function
Private Function mEventIdTest(slLine As String) As Boolean
    Dim blRet As Boolean
    Dim i As Integer
    Dim ilCountCommas As Integer
    Dim ilCountTildes As Integer
    'returns true unless IS an 'event' vehicle AND then fails test
    blRet = True
    ilCountCommas = 0
    ilCountTildes = 0
    If bmTestEventID Then
        If UCase(slLine) <> "-XDS" Then
            For i = 1 To Len(slLine)
                If Mid$(slLine, i, 1) = ":" Then
                    ilCountCommas = ilCountCommas + 1
                End If
                If Mid$(slLine, i, 1) = "~" Then
                    ilCountTildes = ilCountTildes + 1
                End If
            Next
            If bmEventIsCueZone Then
                '4 zones? 3 tildes
                If ilCountTildes <> ZONEEVENTMAX - 1 Then
                    blRet = False
                ElseIf ilCountCommas <> 1 Then
                    blRet = False
                End If
            ElseIf ilCountCommas <> ilCountTildes + 1 Then
                blRet = False
            End If
        End If
    End If
    mEventIdTest = blRet
End Function
'Private Function mTestEventIDVehicle(ilVefCode As Integer) As Boolean
'    Dim blRet As Boolean
'    Dim rst As ADODB.Recordset
'
'    blRet = False
'    SQLQuery = "select vffXDProgCodeID from vff_Vehicle_Features  WHERE vffvefcode = " & ilVefCode
'    Set rst = cnn.Execute(SQLQuery)
'    If Not rst.EOF Then
'        If Trim$(rst!vffXDProgCodeId) = "EVENT" Then
'            blRet = True
'        End If
'    End If
'    mTestEventIDVehicle = blRet
'End Function
'10933
'Private Function mTestEventIDVehicle(ilVefCode As Integer) As Boolean
'    Dim blRet As Boolean
'    Dim ilVff As Integer
'
'    blRet = False
'    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
'        If ilVefCode = tgVff(ilVff).iVefCode Then
'            If Trim$(tgVff(ilVff).sXDProgCodeID) = "EVENT" Then
'                blRet = True
'            End If
'            Exit For
'        End If
'    Next ilVff
'    mTestEventIDVehicle = blRet
'End Function
Private Sub mTestEventIDVehicle(ilVefCode As Integer)
    Dim ilVff As Integer
    
    bmTestEventID = False
    bmEventIsCueZone = False
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If ilVefCode = tgVff(ilVff).iVefCode Then
            If Trim$(tgVff(ilVff).sXDProgCodeID) = "EVENT" Then
                bmTestEventID = True
            End If
            If Trim$(tgVff(ilVff).sXDEventZone) = "Y" Then
                bmEventIsCueZone = True
            End If
            Exit For
        End If
    Next ilVff
End Sub
