VERSION 5.00
Begin VB.UserControl CBill 
   Appearance      =   0  'Flat
   ClientHeight    =   5085
   ClientLeft      =   1125
   ClientTop       =   3180
   ClientWidth     =   9300
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5085
   ScaleWidth      =   9300
   Begin VB.TextBox edcWarningMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   555
      Left            =   2100
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Text            =   "cbill.ctx":0000
      Top             =   3210
      Visible         =   0   'False
      Width           =   4830
   End
   Begin VB.TextBox edcPromoMsg 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Height          =   330
      Left            =   1905
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   2475
      Visible         =   0   'False
      Width           =   6345
   End
   Begin VB.TextBox edcMerchMsg 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Height          =   330
      Left            =   1635
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   2130
      Visible         =   0   'False
      Width           =   6345
   End
   Begin VB.TextBox edcNTRMsg 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Height          =   330
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   1740
      Visible         =   0   'False
      Width           =   6345
   End
   Begin VB.CommandButton cmcClear 
      Appearance      =   0  'Flat
      Caption         =   "C&lear"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7155
      TabIndex        =   71
      Top             =   4740
      Width           =   945
   End
   Begin VB.TextBox edcInstallMsg 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Height          =   330
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   1380
      Visible         =   0   'False
      Width           =   6345
   End
   Begin VB.TextBox edcNoMonthsOff 
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
      HelpContextID   =   8
      Left            =   4065
      MaxLength       =   2
      TabIndex        =   15
      Top             =   1245
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox edcNoMonthsOn 
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
      HelpContextID   =   8
      Left            =   3675
      MaxLength       =   2
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.ListBox lbcPEDate 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "cbill.ctx":0028
      Left            =   4905
      List            =   "cbill.ctx":002A
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   2595
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcTax 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "cbill.ctx":002C
      Left            =   1365
      List            =   "cbill.ctx":002E
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   1140
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.TextBox edcAcqAmount 
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
      Left            =   6960
      MaxLength       =   12
      TabIndex        =   47
      Top             =   930
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox pbcLbcPSVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1710
      ScaleHeight     =   195
      ScaleWidth      =   840
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   4845
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox pbcLbcBVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1470
      ScaleHeight     =   195
      ScaleWidth      =   840
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   4635
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox pbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   390
      Picture         =   "cbill.ctx":0030
      ScaleHeight     =   930
      ScaleWidth      =   3285
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   3315
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
      Left            =   6750
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2445
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "cbill.ctx":A04A
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   40
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
            TabIndex        =   41
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
         TabIndex        =   37
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
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.TextBox edcSalesComm 
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
      Left            =   7155
      MaxLength       =   12
      TabIndex        =   45
      Top             =   1515
      Visible         =   0   'False
      Width           =   1275
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
      Left            =   15
      Picture         =   "cbill.ctx":CE64
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.ListBox lbcPESDate 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4020
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3885
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcPSSDate 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4290
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3075
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcMSSDate 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2310
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3735
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox edcPPercent 
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
      HelpContextID   =   8
      Left            =   2355
      MaxLength       =   2
      TabIndex        =   17
      Top             =   2970
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox edcMPercent 
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
      HelpContextID   =   8
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   16
      Top             =   2715
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox lbcMESDate 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   855
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   1410
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
      Left            =   5130
      ScaleHeight     =   210
      ScaleWidth      =   180
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1140
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox edcAmount 
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
      Left            =   5940
      MaxLength       =   12
      TabIndex        =   46
      Top             =   1170
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox edcDescription 
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
      Left            =   2070
      MaxLength       =   80
      TabIndex        =   43
      Top             =   705
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.ListBox lbcBVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4380
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3390
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.ListBox lbcBDate 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4170
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2685
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.ListBox lbcBItem 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   750
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2580
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lbcPSVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1950
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
      Width           =   2325
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
      Left            =   1980
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
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
      Left            =   1635
      Picture         =   "cbill.ctx":D16E
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1635
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcPSDropDown 
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
      Left            =   810
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmcPSDropDown 
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
      Left            =   1845
      Picture         =   "cbill.ctx":D268
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcUnits 
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
      HelpContextID   =   8
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   49
      Top             =   1215
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox edcNoItems 
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
      HelpContextID   =   8
      Left            =   7140
      MaxLength       =   5
      TabIndex        =   51
      Top             =   1230
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox edcPSAmount 
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
      Left            =   5790
      MaxLength       =   12
      TabIndex        =   23
      Top             =   1590
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox edcNoPeriods 
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
      HelpContextID   =   8
      Left            =   4920
      MaxLength       =   2
      TabIndex        =   13
      Top             =   825
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox pbcMPSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2430
      Picture         =   "cbill.ctx":D362
      ScaleHeight     =   375
      ScaleWidth      =   4485
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2205
      Visible         =   0   'False
      Width           =   4485
   End
   Begin VB.PictureBox pbcFixSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1065
      Picture         =   "cbill.ctx":12F0C
      ScaleHeight     =   375
      ScaleWidth      =   6150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1905
      Visible         =   0   'False
      Width           =   6150
   End
   Begin VB.PictureBox plcFixSpec 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   1035
      ScaleHeight     =   405
      ScaleWidth      =   7260
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1875
      Visible         =   0   'False
      Width           =   7320
      Begin VB.CommandButton cmcGen 
         Appearance      =   0  'Flat
         Caption         =   "&Generate"
         Height          =   285
         Left            =   6240
         TabIndex        =   25
         Top             =   60
         Width           =   945
      End
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   795
      Top             =   4380
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   300
      Top             =   4395
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5820
      TabIndex        =   59
      Top             =   4740
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcFBTab 
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
      Height          =   75
      Left            =   435
      ScaleHeight     =   75
      ScaleWidth      =   90
      TabIndex        =   54
      Top             =   4200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox pbcFSTab 
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
      Left            =   225
      ScaleHeight     =   90
      ScaleWidth      =   90
      TabIndex        =   24
      Top             =   4185
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox pbcIBTab 
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
      Left            =   -15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   53
      Top             =   4200
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox pbcFBSTab 
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
      Left            =   1635
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   27
      Top             =   390
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.VScrollBar vbcItemBill 
      Height          =   3480
      LargeChange     =   15
      Left            =   8910
      TabIndex        =   55
      Top             =   615
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4320
      TabIndex        =   58
      Top             =   4740
      Visible         =   0   'False
      Width           =   1050
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
      Height          =   420
      Left            =   1890
      ScaleHeight     =   360
      ScaleWidth      =   5595
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   5655
      Begin VB.OptionButton rbcOption 
         Caption         =   "Promotion"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   2805
         TabIndex        =   4
         Top             =   105
         Width           =   1215
      End
      Begin VB.OptionButton rbcOption 
         Caption         =   "Merchandising"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   990
         TabIndex        =   3
         Top             =   105
         Width           =   1545
      End
      Begin VB.OptionButton rbcOption 
         Caption         =   "NTR"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   105
         Width           =   765
      End
      Begin VB.OptionButton rbcOption 
         Caption         =   "Installment"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   4260
         TabIndex        =   5
         Top             =   105
         Width           =   1335
      End
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   825
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2820
      TabIndex        =   57
      Top             =   4740
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   645
      Sorted          =   -1  'True
      TabIndex        =   60
      Top             =   660
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.ListBox lbcPSDate 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4905
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2325
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.VScrollBar vbcFix 
      Height          =   2715
      LargeChange     =   11
      Left            =   8655
      TabIndex        =   56
      Top             =   1485
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pbcMP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2715
      Left            =   1275
      Picture         =   "cbill.ctx":1AC6E
      ScaleHeight     =   2715
      ScaleWidth      =   4455
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1710
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.PictureBox pbcFix 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2715
      Left            =   675
      Picture         =   "cbill.ctx":28248
      ScaleHeight     =   2715
      ScaleWidth      =   7995
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   7995
   End
   Begin VB.PictureBox plcFix 
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
      Height          =   2835
      Left            =   1695
      ScaleHeight     =   2775
      ScaleWidth      =   4710
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1395
      Visible         =   0   'False
      Width           =   4770
   End
   Begin VB.PictureBox pbcItemBill 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   255
      Picture         =   "cbill.ctx":6F6E2
      ScaleHeight     =   3495
      ScaleWidth      =   8655
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   705
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Label lacIBFrame 
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
         Left            =   30
         TabIndex        =   63
         Top             =   720
         Visible         =   0   'False
         Width           =   8670
      End
   End
   Begin VB.PictureBox plcItemBill 
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
      Height          =   3615
      Left            =   165
      ScaleHeight     =   3555
      ScaleWidth      =   8985
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   570
      Visible         =   0   'False
      Width           =   9045
   End
   Begin VB.VScrollBar vbcMP 
      Height          =   2715
      LargeChange     =   11
      Left            =   6135
      TabIndex        =   69
      Top             =   1770
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox plcMP 
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
      Height          =   2835
      Left            =   1635
      ScaleHeight     =   2775
      ScaleWidth      =   4710
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   1665
      Visible         =   0   'False
      Width           =   4770
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
      Left            =   -30
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4620
      Width           =   75
   End
   Begin VB.PictureBox pbcIBSTab 
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
      Left            =   -45
      ScaleHeight     =   120
      ScaleWidth      =   105
      TabIndex        =   26
      Top             =   450
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcFSSTab 
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
      Left            =   -45
      ScaleHeight     =   135
      ScaleWidth      =   90
      TabIndex        =   6
      Top             =   270
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   960
      Picture         =   "cbill.ctx":D29F4
      Top             =   45
      Width           =   480
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8715
      Picture         =   "cbill.ctx":D2CFE
      Top             =   4575
      Width           =   480
   End
   Begin VB.Label lacTotals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7905
      TabIndex        =   62
      Top             =   4275
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "CBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of cbill.ctl on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Constants (Removed)                                                            *
'*  SBFPKGFACIL                   SBFPKGDATE                    SBFPKGNOPER               *
'*  SBFPKGTOT                     SBFBODYBILL                   SBFBODYAMT                *
'*  SBFBODYAVG                    SBFBODYZERO                   SBFITEMTX                 *
'*  SBFBODYDONTALT                                                                        *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mInitNewFB                                                                            *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CBill.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract item/package billing input screen code

'  Method            Vehicles allowed                      Option        item names
'  Merchandizing     Contract Vehicle only                 rbcOption(2)  --MB----
'  Promotion         Contract Vehicles only                rbcOption(3)  --PB----
'  NTR               NTR vehicles and all other vehicles   rbcOption(1)  --IB----
'  Installment       Contract vehicles only                rbcOption(0)  --FB----
'
'  lbcVehicle has sort code|sort code| vehicle Name\vehicle code
'  lbcPSVehicle same vehicles as lbcVehicle except only the names
'  lbcBVehicle has NTR vehicles and all other vehicles
'
'  Installment only shows if user selects "Fixed broadcast" or "Fixed Calendar" for contract billing cycle.
'  Currently, these options are not available.
'
'  NTR option only avail ig site set to using NTR
'
'  Merchandizing only allowed if site set to using merchandizing and merchandizing comment defined
'
'  Promotion only allowed if site set to using promotion and promotion comment defined
'
Option Explicit
Option Compare Text

Private lmOpenPreviouslyCompleted As Long

Public Event NTRDollars(slNTRTotal As String)
Public Event SetSave(ilStatus As Integer)

Dim imSetButtons As Integer
'Package billing
'Dim tmFBCtrls(1 To 4) As FIELDAREA
'Dim tmFBCtrls(1 To 6) As FIELDAREA
'Dim tmFSCtrls(1 To 6) As FIELDAREA
Dim tmFBCtrls(0 To 6) As FIELDAREA  'Index zero ignored
Dim tmFSCtrls(0 To 6) As FIELDAREA  'Index zero ignored
Dim imLBCtrls As Integer
Dim smFBBTotal As String
Dim smGross As String   'Contract Gross
Dim smNet As String     'Contract Net
Dim smMPercent As String    'Merchandising Percent from contract or comment
Dim smPPercent As String    'Promotion Percent from contract or comment
Dim smAgyRate As String
Dim smLnTGross As String    'Total Gross from Lines
Dim smLnTNet As String      'Total Net from Lines
Dim smNTRTGross As String
Dim smFixTGross As String
'Merchandising or Promotion
'Dim tmMPBCtrls(1 To 3) As FIELDAREA
'Dim tmMSCtrls(1 To 3) As FIELDAREA
'Dim tmPSCtrls(1 To 3) As FIELDAREA
Dim tmMPBCtrls(0 To 3) As FIELDAREA
Dim tmMSCtrls(0 To 3) As FIELDAREA
Dim tmPSCtrls(0 To 3) As FIELDAREA
Dim smMBTotal As String
Dim smPBTotal As String
'Billing Items
'Dim tmIBCtrls(1 To 12) As FIELDAREA
Dim tmIBCtrls(0 To 12) As FIELDAREA
Dim smIBBTotal As String
Dim smIBPTotal As String
Dim hmSbf As Integer        'Special billing
Dim tmFBSbf() As SBFLIST    'SBF record image of package bill
Dim tmIBSbf() As SBFLIST    'SBF record image of billing items
Dim tmMBSbf() As SBFLIST    'SBF record image of merchandising
Dim tmPBSbf() As SBFLIST    'SBF record image of promotion
Dim lmFBSbfCode() As Long
Dim lmIBSbfCode() As Long
Dim imRecLen As Integer     'SBF record length
Dim tmItemCode() As SORTCODE
Dim smItemCodeTag As String
Dim tmTaxSortCode() As SORTCODE
Dim smTaxSortCodeTag As String
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imIBBoxNo As Integer
Dim imFSBoxNo As Integer
Dim imFBBoxNo As Integer
Dim imMSBoxNo As Integer
Dim imMBBoxNo As Integer
Dim imPSBoxNo As Integer
Dim imPBBoxNo As Integer
Dim imIBRowNo As Integer
Dim imFBRowNo As Integer
Dim imMBRowNo As Integer
Dim imPBRowNo As Integer
'Dim smFBSave() As String    'Values saved for Package Bill(1=Transaction; 2=Amount; 3=Billed)
'Dim imFBSave() As Integer   'Values saved for Package Bill (1=Vehicle; 2=Date)
'Dim smFBShow() As String    'Show values for Package Bill
Dim smIBSave() As String    'Values saved for Item bill(1=Transaction; 2=Description 3= Amount/item; 4= Units; 5=# Items; 6=Total Amount; 7=Billed; 8=Date; 9=PrintInvDate, 10=Acqusition Amount; 11=Game-Independent(Y or N) if ihfCode > 0
Dim imIBSave() As Integer   'Values saved for Item Bill (1=Vehicle; 2=Date(Not used); 3=Item Billing type; 4=Agy Comm; 5= Salesperson comm (-1=new,0=Yes;1=No;2=No comm defined); 6= sales tax, 7=ihfCode)
Dim smIBShow() As String    'Show values for Item Bill
Dim lmIBSave() As Long      'Values saved for Item Bill (1=Tax1; 2=Tax2)
Dim smMBSave() As String    'Values saved for Merchandising(1=Transaction Type(M); 2=Amount; 3=Date)
Dim imMBSave() As Integer   'Values saved for Merchandising (1=Vehicle)
Dim smMBShow() As String    'Show values for Merchandising
Dim smPBSave() As String    'Values saved for Promotion(1=Transaction Type (P); 2=Amount; 3= Date)
Dim imPBSave() As Integer   'Values saved for Promotion (1=Vehicle)
Dim smPBShow() As String    'Show values for Promotion
Dim imFixSort As Integer    '0=Date; 1=Vehicle
Dim imFBChg As Integer
Dim imIBChg As Integer
Dim imMBChg As Integer
Dim imPBChg As Integer
Dim imComboBoxIndex As Integer
Dim imSettingValue As Integer
Dim imShowPS As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUsingInstallBill As Integer
Dim lmCntrStartDate As Long
Dim lmCntrEndDate As Long
Dim imTaxDefined As Integer
Dim imBypassFocus As Integer
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imUpdateAllowed As Integer
Dim smLastClosingDate As String
Dim lmLastClosingDate As Long
Dim smLastBilledDate As String
Dim lmLastBilledDate As Long
Dim smBillCycle As String

Dim imPSNonPkgStartIndex As Integer
Dim imBPkgStartIndex As Integer
Dim imBNonPkgStartIndex As Integer
Dim imINPBCPaint As Integer

'Calendar
'Dim tmCDCtrls(1 To 7) As FIELDAREA
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer

'Agency
Dim tmAgf As AGF            'AGF record image
Dim tmAgfSrchKey As INTKEY0 'AGF key record image
Dim hmAgf As Integer        'AGF Handle
Dim imAgfRecLen As Integer     'AGF record length

Dim hmIhf As Integer
Dim tmIhf As IHF        'IHF record image
Dim tmIhfSrchKey0 As INTKEY0    'IHF key record image
Dim imIhfRecLen As Integer        'IHF record length

'Help
Dim tmSbfHelp() As HLF

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor
Dim lmOrigCBillWidth As Long

Const LBONE = 1

'Fixed period cost
Const FBVEHICLEINDEX = 1   'Vehicle control/field
Const FBDATEINDEX = 2       'Date control/field
Const FBORDEREDINDEX = 3 '4     'Amount control/field
Const FBREVENUETOTALINDEX = 4
Const FBBILLINGINDEX = 5
Const FBBILLTOTALINDEX = 6

'Fixed specifications
Const FSSTARTDATEINDEX = 1   'Vehicle control/field
Const FSNOMONTHSINDEX = 2       'Date control/field
Const FSENDDATEINDEX = 3   'Number of periods control/field
Const FSAMOUNTINDEX = 6     'Amount control/index
Const FSNOMONTHSONINDEX = 4
Const FSNOMONTHSOFFINDEX = 5
'Merchandising and Promotion period cost
Const MPBVEHICLEINDEX = 1   'Vehicle control/field
Const MPBDATEINDEX = 2       'Date control/field
Const MPBAMOUNTINDEX = 3     'Amount control/field
'Merchandising and Promotion specifications
Const MPSPERCENTINDEX = 1   'Percent control/field
Const MPSSTARTDATEINDEX = 2 'Start Date control/field
Const MPSENDDATEINDEX = 3   'End Date control/field
'NTR
Const IBVEHICLEINDEX = 1   'Vehicle control/field
Const IBDATEINDEX = 2       'Bill date control/index
Const IBDESCRIPTINDEX = 3   'Description control/field
Const IBITEMTYPEINDEX = 4   'Item billing type control/field
Const IBACINDEX = 5         'Agency commission type control/field
Const IBSCINDEX = 6         'Salesperson commission control/field
Const IBTXINDEX = 7         'Taxable control/field
Const IBAMOUNTINDEX = 8     'Amount per item control/field
Const IBUNITSINDEX = 9      'Units control/field
Const IBNOITEMSINDEX = 10    'Number of items control/field
Const IBTAMOUNTINDEX = 11    'Total amount control/field
Const IBACQCOSTINDEX = 12    'Acquisition Cost per item
'Help messages
Const SBFPKG = 1
Const SBFITEM = 2
Const SBFBTGEN = 7
Const SBFBODYFACIL = 8
Const SBFBODYDATE = 9
Const SBFBODYBT = 14
Const SBFBTPKGDONE = 16
Const SBFBTPKGCANCEL = 17
Const SBFBTPKGREVERT = 18
Const SBFITEMFACIL = 19
Const SBFITEMDATE = 20
Const SBFITEMDESC = 21
Const SBFITEMCOST = 22
Const SBFITEMUNITS = 23
Const SBFITEMNO = 24
Const SBFITEMAC = 25
Const SBFITEMTOT = 28
Const SBFITEMBILL = 29

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
    igPkgChgd = NO
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mAllSetShow
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    gCtrlGotFocus ActiveControl
    gShowHelpMess tmSbfHelp(), SBFBTPKGCANCEL
End Sub
Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcClear_Click()
    imFSBoxNo = -1
    imFBBoxNo = -1
    imFBRowNo = -1
    If UBound(tmInstallBillInfo) > LBound(tmInstallBillInfo) Then
        imFBChg = True
    End If
    ReDim tmInstallBillInfo(0 To 0) As INSTALLBILLINFO
    '12/18/17: Break out NTR separate from Air Time
    bgBreakoutNTR = True
    vbcFix.Min = LBound(tmInstallBillInfo) + 1 'LBound(smFBShow, 2)
    vbcFix.Max = UBound(tmInstallBillInfo) + 1  'LBound(smFBShow, 2)
    vbcFix.Value = vbcFix.Min
    pbcFix.Cls
    pbcFix_Paint
    mFBTotals True
    mSetCommands
End Sub

Private Sub cmcClear_GotFocus()
    mFSSetShow imFSBoxNo
    imFSBoxNo = -1
    mFBSetShow imFBBoxNo
    imFBBoxNo = -1
    imFBRowNo = -1
End Sub

Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        If rbcOption(0).Value Then
            mFBEnableBox imFBBoxNo
        ElseIf rbcOption(1).Value Then
            mIBEnableBox imIBBoxNo
        ElseIf rbcOption(2).Value Then
            mMBEnableBox imMBBoxNo
        ElseIf rbcOption(3).Value Then
            mPBEnableBox imPBBoxNo
        End If
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mAllSetShow
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    gCtrlGotFocus ActiveControl
    gShowHelpMess tmSbfHelp(), SBFBTPKGDONE
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_Click()
    If rbcOption(0).Value Then  'Package
        Select Case imFBBoxNo
            Case FBVEHICLEINDEX
                lbcPSVehicle.Visible = Not lbcPSVehicle.Visible
                imINPBCPaint = False
                If lbcPSVehicle.Visible Then
                    pbcLbcPSVehicle.Visible = True
                Else
                    pbcLbcPSVehicle.Visible = False
                End If
            Case FBDATEINDEX
                lbcBDate.Visible = Not lbcBDate.Visible
        End Select
    Else
        Select Case imIBBoxNo
            Case IBVEHICLEINDEX
                lbcBVehicle.Visible = Not lbcBVehicle.Visible
                imINPBCPaint = False
                If lbcBVehicle.Visible Then
                    pbcLbcBVehicle.Visible = True
                Else
                    pbcLbcBVehicle.Visible = False
                End If
            Case IBDATEINDEX
                'lbcBDate.Visible = Not lbcBDate.Visible
                plcCalendar.Visible = Not plcCalendar.Visible
            Case IBITEMTYPEINDEX
                lbcBItem.Visible = Not lbcBItem.Visible
            Case IBTXINDEX
                lbcTax.Visible = Not lbcTax.Visible
        End Select
    End If
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcGen_Click()
    Dim ilLoop As Integer
    Dim slATotal As String
    Dim slDate As String

    Screen.MousePointer = vbHourglass
    If rbcOption(2).Value Then
        slDate = lbcMSSDate.List(lbcMSSDate.ListIndex)
        If gDateValue(slDate) <= lmLastBilledDate Then
            Screen.MousePointer = vbDefault
            MsgBox "Start Date must be after last billed date " & smLastBilledDate, vbOKOnly + vbExclamation, "Incomplete"
            'cmcCancel.SetFocus
            Exit Sub
        End If
    ElseIf rbcOption(3).Value Then
        slDate = lbcPSSDate.List(lbcPSSDate.ListIndex)
        If gDateValue(slDate) <= lmLastBilledDate Then
            Screen.MousePointer = vbDefault
            MsgBox "Start Date must be after last billed date " & smLastBilledDate, vbOKOnly + vbExclamation, "Incomplete"
            'cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    mCreateDate True
    If rbcOption(0).Value Then
        'If UBound(tmFBSbf) <= LBound(tmFBSbf) Then
        '    mInitInstallBill
        '    mMoveFBRecToCtrl tmFBSbf()
        'End If
        mFBTotals True
        'For ilLoop = LBound(tmFSCtrls) To UBound(tmFSCtrls) - 1 Step 1
        '    tmFSCtrls(ilLoop).sShow = ""
        'Next ilLoop
        slATotal = "0"
        'For ilLoop = LBound(smFBSave, 2) To UBound(smFBSave, 2) - 1 Step 1
        '    slATotal = gAddStr(slATotal, smFBSave(2, ilLoop))
        'Next ilLoop
        For ilLoop = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
            slATotal = gAddStr(slATotal, gLongToStrDec(tmInstallBillInfo(ilLoop).lBillDollars, 2))
        Next ilLoop
        slATotal = gSubStr(smGross, slATotal)
        'edcPSAmount.Text = ""
        'If lbcPSVehicle.ListIndex <> lbcPSVehicle.ListCount - 1 Then
        '    lbcPSVehicle.ListIndex = lbcPSVehicle.ListIndex + 1
        'End If
        pbcFixSpec.Cls
        pbcFixSpec_Paint
        mSetCommands
        If Val(slATotal) = 0 Then
        '    pbcFBSTab.SetFocus
        Else
            imFBBoxNo = -1
            pbcFBSTab.SetFocus
        End If
    ElseIf rbcOption(2).Value Then

        mMTotals
        'For ilLoop = LBound(tmMSCtrls) To UBound(tmMSCtrls) Step 1
        '    tmMSCtrls(ilLoop).sShow = ""
        'Next ilLoop
        'pbcMPSpec.Cls
        mSetCommands
        imMSBoxNo = -1
        pbcFBSTab.SetFocus
    ElseIf rbcOption(3).Value Then
        mPTotals
        'For ilLoop = LBound(tmPSCtrls) To UBound(tmPSCtrls) Step 1
        '    tmPSCtrls(ilLoop).sShow = ""
        'Next ilLoop
        pbcMPSpec.Cls
        mSetCommands
        imPSBoxNo = -1
        pbcFBSTab.SetFocus
    End If
    mSetGenCommand
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmcGen_GotFocus()
    mAllSetShow
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    gCtrlGotFocus ActiveControl
    gShowHelpMess tmSbfHelp(), SBFBTGEN
End Sub
Private Sub cmcPSDropDown_Click()
    If rbcOption(0).Value Then
        Select Case imFSBoxNo
            Case FSSTARTDATEINDEX
                lbcPSDate.Visible = Not lbcPSDate.Visible
            Case FSENDDATEINDEX
                lbcPEDate.Visible = Not lbcPEDate.Visible
        End Select
    ElseIf rbcOption(2).Value Then
        Select Case imMSBoxNo
            Case MPSSTARTDATEINDEX
                lbcMSSDate.Visible = Not lbcMSSDate.Visible
            Case MPSENDDATEINDEX
                lbcMESDate.Visible = Not lbcMESDate.Visible
        End Select
    ElseIf rbcOption(3).Value Then
        Select Case imPSBoxNo
            Case MPSSTARTDATEINDEX
                lbcPSSDate.Visible = Not lbcPSSDate.Visible
            Case MPSENDDATEINDEX
                lbcPESDate.Visible = Not lbcPESDate.Visible
        End Select
    End If
    edcPSDropDown.SelStart = 0
    edcPSDropDown.SelLength = Len(edcPSDropDown.Text)
    edcPSDropDown.SetFocus
End Sub
Private Sub cmcPSDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUndo_Click()
    Dim ilLoop As Integer

    If rbcOption(0).Value Then 'Fix
        If Not mReadRec(1) Then
            imTerminate = True
            Exit Sub
        End If
        pbcFix.Cls
        mMoveFBRecToCtrl tmFBSbf()
        mInitFBShow
        cmcClear.Enabled = True
        For ilLoop = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
            If tmInstallBillInfo(ilLoop).sBilledFlag = "Y" Then
                cmcClear.Enabled = False
            End If
        Next ilLoop
        pbcFix_Paint
    ElseIf rbcOption(1).Value Then    'Item bill
        If Not mReadRec(2) Then
            imTerminate = True
            Exit Sub
        End If
        pbcItemBill.Cls
        mMoveIBRecToCtrl
        mInitIBShow
        pbcItemBill_Paint
    ElseIf rbcOption(2).Value Then 'Promotion
        If Not mReadRec(3) Then
            imTerminate = True
            Exit Sub
        End If
        pbcMP.Cls
        mMoveMBRecToCtrl
        mInitMBShow
        pbcMP_Paint
    ElseIf rbcOption(3).Value Then 'Promotion
        If Not mReadRec(4) Then
            imTerminate = True
            Exit Sub
        End If
        pbcMP.Cls
        mMovePBRecToCtrl
        mInitPBShow
        pbcMP_Paint
    End If
    If rbcOption(0).Value Then
        pbcFBSTab.SetFocus
    ElseIf rbcOption(1).Value Then
        pbcIBSTab.SetFocus
    ElseIf rbcOption(2).Value Then
        If UBound(tmMBSbf) > 0 Then
            pbcFBSTab.SetFocus
        Else
            'cmcCancel.SetFocus
        End If
    ElseIf rbcOption(3).Value Then
        If UBound(tmPBSbf) > 0 Then
            pbcFBSTab.SetFocus
        Else
            'cmcCancel.SetFocus
        End If
    End If
End Sub
Private Sub cmcUndo_GotFocus()
    mAllSetShow
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    gCtrlGotFocus ActiveControl
    gShowHelpMess tmSbfHelp(), SBFBTPKGREVERT
End Sub
Private Sub cmcUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub edcAcqAmount_GotFocus()
    Dim slNameCode As String
    Dim slAmount As String
    Dim ilRet As Integer
    If (edcAcqAmount.Text = "") And (rbcOption(1).Value) Then
        If imIBSave(3, imIBRowNo) > 0 Then
            slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
            ilRet = gParseItem(slNameCode, 7, "\", slAmount)
            If ilRet = CP_MSG_NONE Then
                edcAcqAmount.Text = slAmount
            End If
        End If
    End If
    gCtrlGotFocus ActiveControl

End Sub

Private Sub edcAcqAmount_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(ActiveControl.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(ActiveControl.Text, ".")    'Disallow multi-decimal points
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
    slStr = edcAcqAmount.Text
    slStr = Left$(slStr, edcAcqAmount.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcAcqAmount.SelStart - edcAcqAmount.SelLength)
    If gCompNumberStr(slStr, "9999999.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcAmount_GotFocus()
    Dim slNameCode As String
    Dim slAmount As String
    Dim ilRet As Integer
    If (edcAmount.Text = "") And (rbcOption(1).Value) Then
        If imIBSave(3, imIBRowNo) > 0 Then
            slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
            ilRet = gParseItem(slNameCode, 3, "\", slAmount)
            If ilRet = CP_MSG_NONE Then
                edcAmount.Text = slAmount
            End If
        End If
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcAmount_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(ActiveControl.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(ActiveControl.Text, ".")    'Disallow multi-decimal points
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
    slStr = edcAmount.Text
    slStr = Left$(slStr, edcAmount.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcAmount.SelStart - edcAmount.SelLength)
    If gCompNumberStr(slStr, "9999999.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDescription_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDescription_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim slDate As String
    Dim ilRet As Integer
    If rbcOption(0).Value Then  'Package
        Select Case imFBBoxNo
            Case FBVEHICLEINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcPSVehicle, imBSMode, imComboBoxIndex
            Case FBDATEINDEX
                imLbcArrowSetting = True
                ilRet = gOptionalLookAhead(edcDropDown, lbcBDate, imBSMode, slStr)
                If ilRet > 1 Then
                    'Reset dates
                    If gValidDate(slStr) Then
                        If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                            slStr = gObtainEndCal(slStr)
                            gEndCalDatePop slStr, 24, lbcBDate
                        ElseIf (smBillCycle = "W") Then
                            slStr = gObtainNextSunday(slStr)
                            gEndWkDatePop slStr, 96, lbcBDate
                        Else
                            slStr = gObtainEndStd(slStr)
                            gEndStdDatePop slStr, 24, lbcBDate
                        End If
                        slDate = Format$(slStr, "m/d/yy")
                        imChgMode = True
                        gFindMatch slDate, 0, lbcBDate
                        If gLastFound(lbcBDate) >= 0 Then
                            lbcBDate.ListIndex = gLastFound(lbcBDate)
                        Else
                            lbcBDate.ListIndex = -1
                            edcDropDown.Text = ""
                            imChgMode = False
                            Exit Sub
                        End If
                        imChgMode = False
                    Else
                        Exit Sub
                    End If
                ElseIf ilRet = 1 Then
                    Exit Sub
                End If
        End Select
    Else
        Select Case imIBBoxNo
            Case IBVEHICLEINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcBVehicle, imBSMode, imComboBoxIndex
            Case IBDATEINDEX
                'imLbcArrowSetting = True
                'ilRet = gOptionalLookAhead(edcDropDown, lbcBDate, imBSMode, slStr)
                'If ilRet > 1 Then
                '    'Reset dates
                '    If gValidDate(slStr) Then
                '        If (smBillCycle = "C") Or (smBillCycle = "D") Then
                '            slStr = gObtainEndCal(slStr)
                '            gEndCalDatePop slStr, 24, lbcBDate
                '        Else
                '            slStr = gObtainEndStd(slStr)
                '            gEndStdDatePop slStr, 24, lbcBDate
                '        End If
                '        slDate = Format$(slStr, "m/d/yy")
                '        imChgMode = True
                '        gFindMatch slDate, 0, lbcBDate
                '        If gLastFound(lbcBDate) >= 0 Then
                '            lbcBDate.ListIndex = gLastFound(lbcBDate)
                '        Else
                '            lbcBDate.ListIndex = -1
                '            edcDropDown.Text = ""
                '            imChgMode = False
                '            Exit Sub
                '        End If
                '        imChgMode = False
                '    Else
                '        Exit Sub
                '    End If
                'ElseIf ilRet = 1 Then
                '    Exit Sub
                'End If
                slStr = edcDropDown.Text
                If Not gValidDate(slStr) Then
                    lacDate.Visible = False
                    Exit Sub
                End If
                lacDate.Visible = True
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint   'mBoxCalDate called within paint
            Case IBITEMTYPEINDEX
                imLbcArrowSetting = True
                ilRet = gOptionalLookAhead(edcDropDown, lbcBItem, imBSMode, slStr)
                If ilRet = 1 Then
                    lbcBItem.ListIndex = 0
                End If
            Case IBTXINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcDropDown, lbcTax, imBSMode, imComboBoxIndex
        End Select
    End If
    imLbcArrowSetting = False
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcDropDown_GotFocus()
    If rbcOption(0).Value Then  'Package
        Select Case imFBBoxNo
            Case FBVEHICLEINDEX
                If lbcPSVehicle.ListCount = 1 Then
                    lbcPSVehicle.ListIndex = 0
                    'If imTabDirection = -1 Then  'Right To Left
                    '    pbcFBSTab.SetFocus
                    'Else
                    '    pbcFBTab.SetFocus
                    'End If
                    'Exit Sub
                End If
            Case FBDATEINDEX
                If lbcBDate.ListCount = 1 Then
                    lbcBDate.ListIndex = 0
                    'If imTabDirection = -1 Then  'Right To Left
                    '    pbcFBSTab.SetFocus
                    'Else
                    '    pbcFBTab.SetFocus
                    'End If
                    'Exit Sub
                End If
        End Select
    Else
        Select Case imIBBoxNo
            Case IBVEHICLEINDEX
                If lbcBVehicle.ListCount = 1 Then
                    lbcBVehicle.ListIndex = 0
                    'If imTabDirection = -1 Then  'Right To Left
                    '    pbcIBSTab.SetFocus
                    'Else
                    '    pbcIBTab.SetFocus
                    'End If
                    'Exit Sub
                End If
            Case IBDATEINDEX
                'If lbcBDate.ListCount = 1 Then
                '    lbcBDate.ListIndex = 0
                '    'If imTabDirection = -1 Then  'Right To Left
                '    '    pbcIBSTab.SetFocus
                '    'Else
                '    '    pbcIBTab.SetFocus
                '    'End If
                '    'Exit Sub
                'End If
            Case IBITEMTYPEINDEX
                If lbcBItem.ListCount = 1 Then
                    lbcBItem.ListIndex = 0
                    'If imTabDirection = -1 Then  'Right To Left
                    '    pbcIBSTab.SetFocus
                    'Else
                    '    pbcIBTab.SetFocus
                    'End If
                    'Exit Sub
                End If
            Case IBTXINDEX
                If lbcTax.ListCount = 1 Then
                    lbcTax.ListIndex = 0
                    'If imTabDirection = -1 Then  'Right To Left
                    '    pbcIBSTab.SetFocus
                    'Else
                    '    pbcIBTab.SetFocus
                    'End If
                    'Exit Sub
                End If
        End Select
    End If
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
    If rbcOption(0).Value Then  'Package
        ilKey = KeyAscii
        If Not gCheckKeyAscii(ilKey) Then
            KeyAscii = 0
            Exit Sub
        End If
    Else
        Select Case imIBBoxNo
            Case IBVEHICLEINDEX
                ilKey = KeyAscii
                If Not gCheckKeyAscii(ilKey) Then
                    KeyAscii = 0
                    Exit Sub
                End If
            Case IBDATEINDEX
                'Filter characters (allow only BackSpace, numbers 0 thru 9
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            Case IBITEMTYPEINDEX
                ilKey = KeyAscii
                If Not gCheckKeyAscii(ilKey) Then
                    KeyAscii = 0
                    Exit Sub
                End If
            Case IBTXINDEX
                ilKey = KeyAscii
                If Not gCheckKeyAscii(ilKey) Then
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
    End If
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String

    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        If rbcOption(0).Value Then  'Package
            Select Case imFBBoxNo
                Case FBVEHICLEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcPSVehicle, imLbcArrowSetting
                Case FBDATEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcBDate, imLbcArrowSetting
            End Select
        Else
            Select Case imIBBoxNo
                Case IBVEHICLEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcBVehicle, imLbcArrowSetting
                Case IBDATEINDEX
                    'gProcessArrowKey Shift, KeyCode, lbcBDate, imLbcArrowSetting
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
                Case IBITEMTYPEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcBItem, imLbcArrowSetting
                Case IBTXINDEX
                    gProcessArrowKey Shift, KeyCode, lbcTax, imLbcArrowSetting
            End Select
        End If
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imIBBoxNo
            Case IBITEMTYPEINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcIBSTab.SetFocus
                Else
                    pbcIBTab.SetFocus
                End If
                Exit Sub
        End Select
        imDoubleClickName = False
    End If
End Sub
Private Sub edcMPercent_Change()
    mSetGenCommand
End Sub

Private Sub edcMPercent_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcMPercent_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcMPercent.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcMPercent.Text, ".")    'Disallow multi-decimal points
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
    If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) <> MERCHPROMOBYDOLLAR Then
        slStr = edcMPercent.Text
        slStr = Left$(slStr, edcMPercent.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcMPercent.SelStart - edcMPercent.SelLength)
        If gCompNumberStr(slStr, "100.00") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Private Sub edcNoItems_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcNoItems_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcNoItems.Text
    slStr = Left$(slStr, edcNoItems.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcNoItems.SelStart - edcNoItems.SelLength)
    If gCompNumberStr(slStr, "30000") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcNoMonthsOff_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcNoMonthsOff_KeyPress(KeyAscii As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcNoMonthsOn_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcNoMonthsOn_KeyPress(KeyAscii As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcNoPeriods_Change()
    Dim ilLoop As Integer
    Dim slStr As String

    If rbcOption(0).Value Then
        Select Case imFSBoxNo
            Case FSNOMONTHSINDEX
                If edcNoPeriods.Text <> "" Then
                    If lbcPSDate.ListIndex >= 0 Then
                        slStr = lbcPSDate.List(lbcPSDate.ListIndex)
                        For ilLoop = 0 To lbcPEDate.ListCount - 1 Step 1
                            If slStr = lbcPEDate.List(ilLoop) Then
                                If ilLoop + Val(edcNoPeriods.Text) - 1 < lbcPEDate.ListCount Then
                                    lbcPEDate.ListIndex = ilLoop + Val(edcNoPeriods.Text) - 1
                                    gSetShow pbcFixSpec, lbcPEDate.List(lbcPEDate.ListIndex), tmFSCtrls(FSENDDATEINDEX)
                                    pbcFixSpec.Cls
                                    pbcFixSpec_Paint
                                    Exit For
                                End If
                            End If
                        Next ilLoop
                    End If
                End If
        End Select
    End If
End Sub

Private Sub edcNoPeriods_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcNoPeriods_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcNoPeriods.Text
    slStr = Left$(slStr, edcNoPeriods.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcNoPeriods.SelStart - edcNoPeriods.SelLength)
    If gCompNumberStr(slStr, "999") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcPPercent_Change()
    mSetGenCommand
End Sub
Private Sub edcPPercent_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcPPercent.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcPPercent.Text, ".")    'Disallow multi-decimal points
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
    If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) <> MERCHPROMOBYDOLLAR Then
        slStr = edcPPercent.Text
        slStr = Left$(slStr, edcPPercent.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcPPercent.SelStart - edcPPercent.SelLength)
        If gCompNumberStr(slStr, "100.00") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Private Sub edcPSAmount_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcPSAmount_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(ActiveControl.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(ActiveControl.Text, ".")    'Disallow multi-decimal points
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
    slStr = edcPSAmount.Text
    slStr = Left$(slStr, edcPSAmount.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcPSAmount.SelStart - edcPSAmount.SelLength)
    If gCompNumberStr(slStr, "999999999.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcPSDropDown_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slDate                                                                                *
'******************************************************************************************

    Dim slStr As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilCount As Integer

    If rbcOption(0).Value Then
        Select Case imFSBoxNo
            Case FSSTARTDATEINDEX
                imLbcArrowSetting = True
                ilRet = gOptionalLookAhead(edcPSDropDown, lbcPSDate, imBSMode, slStr)
                If ilRet = 0 Then
                    'Reset dates
                    'slStr = Left$(slStr, 3) & " 01" & Mid(slStr, 4)
                    'slStr = Format(slStr, "m/d/yy")
                    If gValidDate(slStr) Then
                        lbcPEDate.Clear
                        For ilLoop = lbcPSDate.ListIndex To lbcPSDate.ListCount - 1 Step 1
                            lbcPEDate.AddItem lbcPSDate.List(ilLoop)
                            lbcPEDate.ItemData(lbcPEDate.NewIndex) = lbcPSDate.ItemData(ilLoop)
                        Next ilLoop
                        If edcNoPeriods.Text <> "" Then
                            slStr = edcPSDropDown.Text
                            For ilLoop = 0 To lbcPEDate.ListCount - 1 Step 1
                                If slStr = lbcPEDate.List(ilLoop) Then
                                    If ilLoop + Val(edcNoPeriods.Text) - 1 < lbcPEDate.ListCount Then
                                        lbcPEDate.ListIndex = ilLoop + Val(edcNoPeriods.Text) - 1
                                        gSetShow pbcFixSpec, lbcPEDate.List(lbcPEDate.ListIndex), tmFSCtrls(FSENDDATEINDEX)
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                        End If
                        pbcFixSpec.Cls
                        pbcFixSpec_Paint
                    Else
                        Exit Sub
                    End If
                ElseIf ilRet = 1 Then
                    Exit Sub
                End If
            Case FSENDDATEINDEX
                imLbcArrowSetting = True
                ilRet = gOptionalLookAhead(edcPSDropDown, lbcPEDate, imBSMode, slStr)
                If ilRet = 0 Then
                    'Reset dates
                    slStr = edcPSDropDown.Text
                    'slStr = Left$(slStr, 3) & " 01" & Mid(slStr, 4)
                    'slStr = Format(slStr, "m/d/yy")
                    If gValidDate(slStr) Then
                        edcNoPeriods.Text = ""
                        If lbcPSDate.ListIndex >= 0 Then
                            ilCount = 0
                            slStr = edcPSDropDown.Text
                            For ilLoop = lbcPSDate.ListIndex To lbcPSDate.ListCount - 1 Step 1
                                If gDateValue(slStr) = lbcPSDate.ItemData(ilLoop) Then
                                    edcNoPeriods.Text = Trim$(str$(ilCount + 1))
                                    gSetShow pbcFixSpec, edcNoPeriods.Text, tmFSCtrls(FSNOMONTHSINDEX)
                                    Exit For
                                End If
                                ilCount = ilCount + 1
                            Next ilLoop
                        End If
                        pbcFixSpec.Cls
                        pbcFixSpec_Paint
                    Else
                        Exit Sub
                    End If
                ElseIf ilRet = 1 Then
                    Exit Sub
                End If
        End Select
    ElseIf rbcOption(2).Value Then
        Select Case imMSBoxNo
            Case MPSSTARTDATEINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcPSDropDown, lbcMSSDate, imBSMode, imComboBoxIndex
            Case MPSENDDATEINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcPSDropDown, lbcMESDate, imBSMode, imComboBoxIndex
        End Select
    ElseIf rbcOption(3).Value Then
        Select Case imPSBoxNo
            Case MPSSTARTDATEINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcPSDropDown, lbcPSSDate, imBSMode, imComboBoxIndex
            Case MPSENDDATEINDEX
                imLbcArrowSetting = True
                gMatchLookAhead edcPSDropDown, lbcPESDate, imBSMode, imComboBoxIndex
        End Select
    End If
    mSetGenCommand
End Sub
Private Sub edcPSDropDown_GotFocus()

    'Select Case imFSBoxNo
    '    Case FSVEHICLEINDEX
    '        If lbcPSVehicle.ListCount = 1 Then
    '            lbcPSVehicle.ListIndex = 0
    '            pbcFSTab.SetFocus
    '            Exit Sub
    '        End If
    '    Case FSDATEINDEX
    '        If lbcPSDate.ListCount = 1 Then
    '            lbcPSDate.ListIndex = 0
    '            pbcFSTab.SetFocus
    '            Exit Sub
    '        End If
    'End Select
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcPSDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcPSDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcPSDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        If rbcOption(0).Value Then
            Select Case imFSBoxNo
                Case FSSTARTDATEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcPSDate, imLbcArrowSetting
                Case FSENDDATEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcPSDate, imLbcArrowSetting
            End Select
        ElseIf rbcOption(2).Value Then
            Select Case imMSBoxNo
                Case MPSSTARTDATEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcMSSDate, imLbcArrowSetting
                Case MPSSTARTDATEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcMESDate, imLbcArrowSetting
            End Select
        ElseIf rbcOption(3).Value Then
            Select Case imPSBoxNo
                Case MPSSTARTDATEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcPSSDate, imLbcArrowSetting
                Case MPSSTARTDATEINDEX
                    gProcessArrowKey Shift, KeyCode, lbcPESDate, imLbcArrowSetting
            End Select
        End If
        edcPSDropDown.SelStart = 0
        edcPSDropDown.SelLength = Len(edcPSDropDown.Text)
    End If
End Sub

Private Sub edcSalesComm_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSalesComm_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(ActiveControl.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(ActiveControl.Text, ".")    'Disallow multi-decimal points
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
    slStr = edcSalesComm.Text
    slStr = Left$(slStr, edcSalesComm.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSalesComm.SelStart - edcSalesComm.SelLength)
    If gCompNumberStr(slStr, "99.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcUnits_GotFocus()
    Dim slNameCode As String
    Dim slUnits As String
    Dim ilRet As Integer
    If edcUnits.Text = "" Then
        If imIBSave(3, imIBRowNo) > 0 Then
            slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
            ilRet = gParseItem(slNameCode, 4, "\", slUnits)
            If ilRet = CP_MSG_NONE Then
                edcUnits.Text = slUnits
            End If
        End If
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Exit Sub
    End If
    imFirstActivate = False
    imUpdateAllowed = igUpdateAllowed
    If Contract.plcScreen = "Contracts" Then
        'If (igWinStatus(CONTRACTSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        If Not imUpdateAllowed Then
            pbcItemBill.Enabled = False
            pbcIBSTab.Enabled = False
            pbcIBTab.Enabled = False
            pbcFix.Enabled = False
            pbcFBSTab.Enabled = False
            pbcFBTab.Enabled = False
            pbcFixSpec.Enabled = False
            pbcFSSTab.Enabled = False
            pbcFSTab.Enabled = False
            pbcMP.Enabled = False
            pbcMPSpec.Enabled = False
            'imUpdateAllowed = False
        Else
            pbcItemBill.Enabled = True
            pbcIBSTab.Enabled = True
            pbcIBTab.Enabled = True
            pbcFix.Enabled = True
            pbcFBSTab.Enabled = True
            pbcFBTab.Enabled = True
            pbcFixSpec.Enabled = True
            pbcFSSTab.Enabled = True
            pbcFSTab.Enabled = True
            pbcMP.Enabled = True
            pbcMPSpec.Enabled = True
            'imUpdateAllowed = True
        End If
    Else
        'If (igWinStatus(PROPOSALSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        If Not imUpdateAllowed Then
            pbcItemBill.Enabled = False
            pbcIBSTab.Enabled = False
            pbcIBTab.Enabled = False
            pbcFix.Enabled = False
            pbcFBSTab.Enabled = False
            pbcFBTab.Enabled = False
            pbcFixSpec.Enabled = False
            pbcFSSTab.Enabled = False
            pbcFSTab.Enabled = False
            pbcMP.Enabled = False
            pbcMPSpec.Enabled = False
            'imUpdateAllowed = False
        Else
            pbcItemBill.Enabled = True
            pbcIBSTab.Enabled = True
            pbcIBTab.Enabled = True
            pbcFix.Enabled = True
            pbcFBSTab.Enabled = True
            pbcFBTab.Enabled = True
            pbcFixSpec.Enabled = True
            pbcFSSTab.Enabled = True
            pbcFSTab.Enabled = True
            pbcMP.Enabled = True
            pbcMPSpec.Enabled = True
            'imUpdateAllowed = True
        End If
    End If
    gShowBranner imUpdateAllowed

    If imUsingInstallBill And (UBound(tmInstallBillInfo) = LBound(tmInstallBillInfo)) Then
        rbcOption(0).Value = True
        'rbcOption(0).SetFocus
    Else

        If tgSpf.sUsingNTR = "Y" Then
            rbcOption(1).Value = True
        ElseIf imUsingInstallBill Then
            rbcOption(0).Value = True
            'rbcOption(0).SetFocus
        Else
            If imUpdateAllowed Then
            '    pbcIBSTab.SetFocus
                'If  lbcBVehicle.ListCount > 0 Then
                If (lbcBVehicle.ListCount > 0) And (rbcOption(1).Visible) Then
                    rbcOption(1).Value = True
                    If plcSelect.Enabled Then
                        'rbcOption(1).SetFocus
                    Else
                        'cmcDone.SetFocus
                    End If
                Else
                    If rbcOption(2).Enabled Then
                        rbcOption(2).Value = True
                        'rbcOption(2).SetFocus
                    Else
                        If rbcOption(3).Enabled Then
                            rbcOption(3).Value = True
                            'rbcOption(3).SetFocus
                        Else
                            'cmcCancel.SetFocus
                        End If
                    End If
                End If
            Else
                'cmcDone.SetFocus
            End If
        End If
    End If
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
        If imFBBoxNo > 0 Then
            mFBEnableBox imFBBoxNo
        ElseIf imFSBoxNo > 0 Then
            mFSEnableBox imFSBoxNo
        ElseIf imMSBoxNo > 0 Then
            mMSEnableBox imMSBoxNo
        ElseIf imPSBoxNo > 0 Then
            mPSEnableBox imPSBoxNo
        ElseIf imIBBoxNo > 0 Then
            mIBEnableBox imIBBoxNo
        End If
    End If
End Sub

Private Sub Form_Load()
    lmOrigCBillWidth = 9300 'Width
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Width))) / 100) / Width
        'Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.height) / (480 * 15 / height))) / 100) / height
        'Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Height))) / 100
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
    If Not igManUnload Then
        mFSSetShow imFSBoxNo
        imFSBoxNo = -1
        mFBSetShow imFBBoxNo
        imFBBoxNo = -1
        mIBSetShow imIBBoxNo
        imIBBoxNo = -1
        pbcArrow.Visible = False
        lacIBFrame.Visible = False
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            If imFBBoxNo <> -1 Then
                mFBEnableBox imFBBoxNo
            ElseIf imIBBoxNo <> -1 Then
                mIBEnableBox imIBBoxNo
            End If
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
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
    Dim slStr As String
    If (imIBRowNo < 1) Then
        Exit Sub
    End If
'    If (smIBSave(7, imIBRowNo) = "R") Or (smIBSave(7, imIBRowNo) = "B") Then
'        Exit Sub
'    End If
    If (smIBSave(7, imIBRowNo) = "Y") Then
        Exit Sub
    End If
    ilRowNo = imIBRowNo
    mIBSetShow imIBBoxNo
    imIBBoxNo = -1   '
    imIBRowNo = -1
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    gCtrlGotFocus ActiveControl
    ilUpperBound = UBound(smIBSave, 2)
    If ilRowNo = ilUpperBound Then
        For ilLoop = imLBCtrls To UBound(tmIBCtrls) Step 1
            slStr = ""
            gSetShow pbcItemBill, slStr, tmIBCtrls(ilLoop)
            smIBShow(ilLoop, ilRowNo) = tmIBCtrls(ilLoop).sShow
        Next ilLoop
        pbcItemBill_Paint
        mInitNewIB ilRowNo   'Set defaults for extra row
    ElseIf ilRowNo = ilUpperBound - 1 Then
        For ilLoop = imLBCtrls To UBound(tmIBCtrls) Step 1
            slStr = ""
            gSetShow pbcItemBill, slStr, tmIBCtrls(ilLoop)
            smIBShow(ilLoop, ilRowNo) = tmIBCtrls(ilLoop).sShow
        Next ilLoop
        pbcItemBill_Paint
        ilUpperBound = UBound(smIBSave, 2)
        'ReDim Preserve smIBSave(1 To 11, 1 To ilUpperBound - 1) As String
        'ReDim Preserve imIBSave(1 To 7, 1 To ilUpperBound - 1) As Integer
        'ReDim Preserve smIBShow(1 To IBACQCOSTINDEX, 1 To ilUpperBound - 1) As String
        'ReDim Preserve lmIBSave(1 To 2, 1 To ilUpperBound - 1) As Long
        ReDim Preserve smIBSave(0 To 11, 0 To ilUpperBound - 1) As String
        ReDim Preserve imIBSave(0 To 7, 0 To ilUpperBound - 1) As Integer
        ReDim Preserve smIBShow(0 To IBACQCOSTINDEX, 0 To ilUpperBound - 1) As String
        ReDim Preserve lmIBSave(0 To 2, 0 To ilUpperBound - 1) As Long
        mInitNewIB ilRowNo   'Set defaults for extra row
    Else
        For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
            For ilIndex = 1 To UBound(smIBSave, 1) Step 1
                smIBSave(ilIndex, ilLoop) = smIBSave(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(imIBSave, 1) Step 1
                imIBSave(ilIndex, ilLoop) = imIBSave(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(smIBShow, 1) Step 1
                smIBShow(ilIndex, ilLoop) = smIBShow(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(lmIBSave, 1) Step 1
                lmIBSave(ilIndex, ilLoop) = lmIBSave(ilIndex, ilLoop + 1)
            Next ilIndex
        Next ilLoop
        ilUpperBound = UBound(smIBSave, 2)
        'ReDim Preserve smIBSave(1 To 11, 1 To ilUpperBound - 1) As String
        'ReDim Preserve imIBSave(1 To 7, 1 To ilUpperBound - 1) As Integer
        'ReDim Preserve smIBShow(1 To IBACQCOSTINDEX, 1 To ilUpperBound - 1) As String
        'ReDim Preserve lmIBSave(1 To 2, 1 To ilUpperBound - 1) As Long
        ReDim Preserve smIBSave(0 To 11, 0 To ilUpperBound - 1) As String
        ReDim Preserve imIBSave(0 To 7, 0 To ilUpperBound - 1) As Integer
        ReDim Preserve smIBShow(0 To IBACQCOSTINDEX, 0 To ilUpperBound - 1) As String
        ReDim Preserve lmIBSave(0 To 2, 0 To ilUpperBound - 1) As Long
    End If
    imIBChg = True
    mSetCommands
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcItemBill.Cls
    pbcItemBill_Paint
    mIBTotals True
    pbcClickFocus.SetFocus
End Sub
Private Sub imcTrash_DragDrop(Source As control, X As Single, Y As Single)
'    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        lacIBFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
'        lacEvtFrame.DragIcon = IconTraf!imcIconTrash.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        lacIBFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub lacTotals_Click()
    If rbcOption(0).Value = 0 Then   'Package bill
        gShowHelpMess tmSbfHelp(), SBFBODYBT
    Else
        gShowHelpMess tmSbfHelp(), SBFITEMBILL
    End If
End Sub
Private Sub lacTotals_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub lbcBDate_Click()
    gProcessLbcClick lbcBDate, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcBDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcBItem_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcBItem, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcBItem_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcBItem_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcBItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcBItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcBItem, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcIBSTab.SetFocus
        Else
            pbcIBTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcBVehicle_Click()
    gProcessLbcClick lbcBVehicle, edcDropDown, imChgMode, imLbcArrowSetting
    If Not lbcBVehicle.Visible Then
        pbcLbcBVehicle.Visible = False
    Else
        pbcLbcBVehicle_Paint
    End If
End Sub
Private Sub lbcBVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcBVehicle_Scroll()
    pbcLbcBVehicle_Paint
End Sub

Private Sub lbcMESDate_Click()
    gProcessLbcClick lbcMESDate, edcPSDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcMESDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcMSSDate_Click()
    gProcessLbcClick lbcMSSDate, edcPSDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcMSSDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcPEDate_Click()
    gProcessLbcClick lbcPEDate, edcPSDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcPEDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcPESDate_Click()
    gProcessLbcClick lbcPESDate, edcPSDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcPESDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcPSDate_Click()
    gProcessLbcClick lbcPSDate, edcPSDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcPSDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcPSSDate_Click()
    gProcessLbcClick lbcPSSDate, edcPSDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcPSSDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcPSVehicle_Click()
    gProcessLbcClick lbcPSVehicle, edcPSDropDown, imChgMode, imLbcArrowSetting
    If Not lbcPSVehicle.Visible Then
        pbcLbcPSVehicle.Visible = False
    Else
        pbcLbcPSVehicle_Paint
    End If
End Sub
Private Sub lbcPSVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAllSetShow                     *
'*                                                     *
'*             Created:8/09/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove focus                   *
'*                                                     *
'*******************************************************
Private Sub mAllSetShow()
    mFSSetShow imFSBoxNo
    imFSBoxNo = -1
    mFBSetShow imFBBoxNo
    imFBBoxNo = -1
    imFBRowNo = -1
    mIBSetShow imIBBoxNo
    imIBBoxNo = -1
    imIBRowNo = -1
    mMSSetShow imMSBoxNo
    imMSBoxNo = -1
    mMBSetShow imMBBoxNo
    imMBBoxNo = -1
    imMBRowNo = -1
    mPSSetShow imPSBoxNo
    imPSBoxNo = -1
    mPBSetShow imPBBoxNo
    imPBBoxNo = -1
    imPBRowNo = -1
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateDate                     *
'*                                                     *
'*             Created:8/09/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create flights                 *
'*                                                     *
'*******************************************************
Private Sub mCreateDate(ilPaint As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slTotalAmount                 slRunAmount                   slPdAmount                *
'*  slLtPdAmount                  slNoPd                        slBilledDollars           *
'*                                                                                        *
'******************************************************************************************

'
'   mCreateDate ilPaint
'   Where:
'       ilPaint (I)- True=Paint week after creating
'                    False=Don't paint
'
    Dim ilUpperLimit As Integer
    Dim ilPdNo As Integer
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim slDate As String
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilClf As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim llWDate As Long
    Dim llPrice As Long
    Dim ilSpots As Integer
    Dim llTDollars As Long
    Dim slTDollars As String
    Dim llSDollars As Long
    Dim slSDollars As String
    Dim ilSZero As Integer
    Dim ilNoPds As Integer
    Dim llWDollars As Long
    Dim ilWSpots As Integer
    Dim slPercent As String
    Dim slWDollars As String
    Dim slLineType As String
    Dim ilPds As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim slATotal As String
    Dim slBTotal As String
    Dim ilFirstIndex As Integer
    Dim slDiffTotal As String
    Dim ilCff As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilRowNo As Integer
    Dim ilCol As Integer
    Dim ilLastBilledIndex As Integer
    Dim slTBilledDollars As String
    Dim llTest As Long
    Dim ilVef As Integer
    Dim ilMove As Integer
    Dim llDistDollars As Long
    Dim llMoDollars As Long
    Dim llTMoDollars As Long
    Dim ilLastMoIndex As Integer
    Dim ilUpper As Integer
    Dim ilModResult As Integer
    Dim ilMonthNo As Integer
    Dim ilNoMonthsOn As Integer
    Dim ilNoMonthsOff As Integer
    '12/18/17: Break out NTR separate from Air Time
    Dim slName As String
    Dim slCodeVal As String
    Dim ilIndex As Integer
    

'    ReDim smFBSave(1 To 3, 1 To 1) As String
'    ReDim imFBSave(1 To 2, 1 To 1) As Integer
'    ReDim smFBShow(1 To FBBILLINGINDEX, 1 To 1) As String
    If rbcOption(0).Value Then
'        slTotalAmount = edcPSAmount.Text
'        slNoPd = edcNoPeriods.Text & ".00"
'        slPdAmount = gDivStr(slTotalAmount, slNoPd)
'        ilPdNo = 1
'        slRunAmount = "0"
'        Do While ilPdNo <= Val(edcNoPeriods.Text) - 1
'            slRunAmount = gAddStr(slRunAmount, slPdAmount)
'            ilPdNo = ilPdNo + 1
'        Loop
'        slLtPdAmount = gSubStr(slTotalAmount, slRunAmount)
'        ilUpperLimit = UBound(smFBShow, 2)
'        ilPdNo = 1
'        slDate = lbcPSDate.List(lbcPSDate.ListIndex)
'        'slDate = Left$(slDate, 3) & " 01" & Mid(slDate, 4)
'        'slDate = Format(slDate, "m/d/yy")
'        If (smBillCycle = "C") Or (smBillCycle = "D") Then
'            slDate = gObtainEndCal(slDate)
'        Else
'            slDate = gObtainEndStd(slDate)
'        End If
'        llDate = gDateValue(slDate)
'        Do While ilPdNo <= Val(edcNoPeriods.Text)
'            'Vehicle
'            imFBSave(1, ilUpperLimit) = lbcPSVehicle.ListIndex
'            slStr = lbcPSVehicle.List(lbcPSVehicle.ListIndex)
'            gSetShow pbcFix, slStr, tmFBCtrls(FBVEHICLEINDEX)
'            smFBShow(FBVEHICLEINDEX, ilUpperLimit) = tmFBCtrls(FBVEHICLEINDEX).sShow
'            'Transaction
'            smFBSave(1, ilUpperLimit) = "F"
'            'slStr = " "
'            'gSetShow pbcFix, slStr, tmFBCtrls(FBBILLINDEX)
'            'smFBShow(FBBILLINDEX, ilUpperLimit) = tmFBCtrls(FBBILLINDEX).sShow
'            'Date
'            slStr = Format$(llDate, "m/d/yy")
'            gFindMatch slStr, 0, lbcBDate
'            If gLastFound(lbcBDate) >= 0 Then
'                imFBSave(2, ilUpperLimit) = gLastFound(lbcBDate)
'            Else
'            End If
'            gSetShow pbcFix, slStr, tmFBCtrls(FBDATEINDEX)
'            smFBShow(FBDATEINDEX, ilUpperLimit) = tmFBCtrls(FBDATEINDEX).sShow
'            'Amount
'            If ilPdNo <> Val(edcNoPeriods.Text) Then
'                smFBSave(2, ilUpperLimit) = slPdAmount
'                slStr = slPdAmount
'            Else
'                smFBSave(2, ilUpperLimit) = slLtPdAmount
'                slStr = slLtPdAmount
'            End If
'            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
'            gSetShow pbcFix, slStr, tmFBCtrls(FBBILLINGINDEX)
'            smFBShow(FBBILLINGINDEX, ilUpperLimit) = tmFBCtrls(FBBILLINGINDEX).sShow
'            'Billed
'            smFBSave(3, ilUpperLimit) = "R"
'            'slStr = "R"
'            'gSetShow pbcFix, slStr, tmFBCtrls(FBBILLINDEX)
'            'smFBShow(FBBILLINDEX, ilUpperLimit) = tmFBCtrls(FBBILLINDEX).sShow
'            ilPdNo = ilPdNo + 1
'            llDate = llDate + 1
'            slDate = Format$(llDate, "m/d/yy")
'            If (smBillCycle = "C") Or (smBillCycle = "D") Then
'                slDate = gObtainEndCal(slDate)
'            Else
'                slDate = gObtainEndStd(slDate)
'            End If
'            llDate = gDateValue(slDate)
'            ilUpperLimit = ilUpperLimit + 1
'            ReDim Preserve smFBSave(1 To 3, 1 To ilUpperLimit) As String
'            ReDim Preserve imFBSave(1 To 2, 1 To ilUpperLimit) As Integer
'            ReDim Preserve smFBShow(1 To FBBILLTOTALINDEX, 1 To ilUpperLimit) As String
'        Loop
'        mInitNewFB UBound(smFBSave, 2)
        If lbcPSDate.ListIndex < 0 Then
            Exit Sub
        End If
        If lbcPEDate.ListIndex < 0 Then
            Exit Sub
        End If
        If UBound(tmInstallBillInfo) <= LBound(tmInstallBillInfo) Then
            mInitInstallBill
            mNTRAddedToInstallment
            mInitInstallVeh
        End If
        ilLoop = LBound(tmInstallBillInfo)
        Do While ilLoop <= UBound(tmInstallBillInfo) - 1
            If (tmInstallBillInfo(ilLoop).sType <> "B") And (tmInstallBillInfo(ilLoop).sType <> "I") And (tmInstallBillInfo(ilLoop).sType <> "O") Then
                For ilMove = ilLoop To UBound(tmInstallBillInfo) - 1 Step 1
                    LSet tmInstallBillInfo(ilMove) = tmInstallBillInfo(ilMove + 1)
                Next ilMove
                ReDim Preserve tmInstallBillInfo(0 To UBound(tmInstallBillInfo) - 1) As INSTALLBILLINFO
            Else
                If (tmInstallBillInfo(ilLoop).sType <> "B") And (tmInstallBillInfo(ilLoop).sType <> "I") Then
                    tmInstallBillInfo(ilLoop).lBillDollars = 0
                End If
                ilLoop = ilLoop + 1
            End If
        Loop
        llSDate = lbcPSDate.ItemData(lbcPSDate.ListIndex)
        llEDate = lbcPEDate.ItemData(lbcPEDate.ListIndex)
        ilPdNo = edcNoPeriods.Text
        If Trim$(edcNoMonthsOn.Text) = "" Then
            edcNoMonthsOn.Text = "1"
        Else
            If Val(Trim$(edcNoMonthsOn.Text)) <= 0 Then
                edcNoMonthsOn.Text = "1"
            End If
        End If
        ilNoMonthsOn = edcNoMonthsOn.Text
        If Trim$(edcNoMonthsOff.Text) = "" Then
            edcNoMonthsOff.Text = "0"
        End If
        ilNoMonthsOff = edcNoMonthsOff.Text
        If ilNoMonthsOff > 0 Then
            ilModResult = ilPdNo Mod (ilNoMonthsOn + ilNoMonthsOff)
            ilPdNo = ilNoMonthsOn * (ilPdNo \ (ilNoMonthsOn + ilNoMonthsOff))
            If ilModResult <= ilNoMonthsOn Then
                ilPdNo = ilPdNo + ilModResult
            ElseIf ilModResult > ilNoMonthsOn Then
                ilPdNo = ilPdNo + ilNoMonthsOn
            End If
        End If
        For ilVef = LBound(tmInstallVehInfo) To UBound(tmInstallVehInfo) - 1 Step 1
            llDistDollars = tmInstallVehInfo(ilVef).lOrderedDollars - tmInstallVehInfo(ilVef).lBilledDollars
            llMoDollars = llDistDollars / ilPdNo
            llTMoDollars = 0
            ilLastMoIndex = -1
            llDate = llSDate
            ilMonthNo = 1
            Do
                ilModResult = ilMonthNo Mod (ilNoMonthsOn + ilNoMonthsOff)
                If (ilModResult >= 1) And (ilModResult <= ilNoMonthsOn) Or (ilNoMonthsOff = 0) Then
                    ilFound = False
                    For ilTest = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
                        If (tmInstallBillInfo(ilTest).iVefCode = tmInstallVehInfo(ilVef).iVefCode) And (tmInstallBillInfo(ilTest).lBillDate = llDate) Then
                            '12/18/17: Break out NTR separate from Air Time
                            If (Not bgBreakoutNTR) Or (bgBreakoutNTR And ((tmInstallBillInfo(ilTest).iMnfItem = tmInstallVehInfo(ilVef).iMnfItem))) Then
                                tmInstallBillInfo(ilTest).lBillDollars = llMoDollars
                                llTMoDollars = llTMoDollars + llMoDollars
                                ilLastMoIndex = ilTest
                                ilFound = True
                                Exit For
                            End If
                        End If
                    Next ilTest
                    If Not ilFound Then
                        ilUpper = UBound(tmInstallBillInfo)
                        tmInstallBillInfo(ilUpper).iVefCode = tmInstallVehInfo(ilVef).iVefCode
                        ilTest = gBinarySearchVef(tmInstallBillInfo(ilUpper).iVefCode)
                        If ilTest <> -1 Then
                            tmInstallBillInfo(ilUpper).sVehName = tgMVef(ilTest).sName
                            tmInstallBillInfo(ilUpper).lAirOrderedDollars = 0
                            tmInstallBillInfo(ilUpper).lNTROrderedDollars = 0
                            tmInstallBillInfo(ilUpper).lBillDate = llDate
                            tmInstallBillInfo(ilUpper).lBillDollars = llMoDollars
                            llTMoDollars = llTMoDollars + llMoDollars
                            ilLastMoIndex = ilUpper
                            tmInstallBillInfo(ilUpper).sBilledFlag = "N"
                            tmInstallBillInfo(ilUpper).sType = "N"
                            tmInstallBillInfo(ilUpper).iMnfItem = 0
                            tmInstallBillInfo(ilUpper).sMnfItem = ""
                            '12/18/17: Break out NTR separate from Air Time
                            If bgBreakoutNTR And tmInstallVehInfo(ilVef).iMnfItem > 0 Then
                                tmInstallBillInfo(ilUpper).iMnfItem = tmInstallVehInfo(ilVef).iMnfItem
                                For ilIndex = 0 To UBound(tmItemCode) - 1 Step 1 'lbcItemCode.ListCount - 1 Step 1
                                    slNameCode = tmItemCode(ilIndex).sKey  'lbcItemCode.List(ilIndex)
                                    ilRet = gParseItem(slNameCode, 2, "\", slCodeVal)
                                    If Val(slCodeVal) = tmInstallVehInfo(ilVef).iMnfItem Then
                                        ilRet = gParseItem(slNameCode, 1, "\", slName)
                                        tmInstallBillInfo(ilUpper).sMnfItem = slName
                                        Exit For
                                    End If
                                Next ilIndex
                            End If
                            ReDim Preserve tmInstallBillInfo(0 To ilUpper + 1) As INSTALLBILLINFO
                        End If
                    End If
                End If
                ilMonthNo = ilMonthNo + 1
                llDate = llDate + 1
                slDate = Format$(llDate, "m/d/yy")
                If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                    slDate = gObtainEndCal(slDate)
                ElseIf (smBillCycle = "W") Then
                    slDate = gObtainNextSunday(slDate)
                Else
                    slDate = gObtainEndStd(slDate)
                End If
                llDate = gDateValue(slDate)
            Loop While llDate <= llEDate
            If ilLastMoIndex <> -1 Then
                tmInstallBillInfo(ilLastMoIndex).lBillDollars = tmInstallBillInfo(ilLastMoIndex).lBillDollars + llDistDollars - llTMoDollars
            End If
        Next ilVef
        mInitInstallVeh
        vbcFix.Min = LBound(tmInstallBillInfo) + 1
        If UBound(tmInstallBillInfo) + 1 <= vbcFix.LargeChange + 1 Then
            vbcFix.Max = LBound(tmInstallBillInfo) + 1
        Else
            vbcFix.Max = UBound(tmInstallBillInfo) + 1 - vbcFix.LargeChange
        End If
        vbcFix.Value = vbcFix.Min
        imFBChg = True
        mFBSort
    ElseIf rbcOption(2).Value Then  'Merchandising
        'Retain previous billed dollars and weeks
        'ReDim slMBSave(1 To 3, 1 To 1) As String
        'ReDim ilMBSave(1 To 1, 1 To 1) As Integer
        'ReDim slMBShow(1 To MPBAMOUNTINDEX, 1 To 1) As String
        ReDim slMBSave(0 To 3, 0 To 1) As String
        'ReDim ilMBSave(0 To 1) As Integer
        ReDim ilMBSave(0 To 1, 0 To 1) As Integer
        ReDim slMBShow(0 To MPBAMOUNTINDEX, 0 To 1) As String
        For ilRowNo = LBONE To UBound(smMBSave, 2) - 1 Step 1
            slDate = smMBSave(3, ilRowNo)   'lbcBDate.List(imMBSave(2, ilRowNo))
            If gDateValue(slDate) <= lmLastBilledDate Then
                slMBSave(1, UBound(slMBSave, 2)) = smMBSave(1, ilRowNo)
                slMBSave(2, UBound(slMBSave, 2)) = smMBSave(2, ilRowNo)
                slMBSave(3, UBound(slMBSave, 2)) = smMBSave(3, ilRowNo)
                ilMBSave(1, UBound(slMBSave, 2)) = imMBSave(1, ilRowNo)
                'ilMBSave(2, UBound(slMBSave, 2)) = imMBSave(2, ilRowNo)
                'ReDim Preserve slMBSave(1 To 3, 1 To UBound(slMBSave, 2) + 1) As String
                'ReDim Preserve ilMBSave(1 To 1, 1 To UBound(ilMBSave, 2) + 1) As Integer
                ReDim Preserve slMBSave(0 To 3, 0 To UBound(slMBSave, 2) + 1) As String
                'ReDim Preserve ilMBSave(0 To UBound(ilMBSave, 2) + 1) As Integer
                ReDim Preserve ilMBSave(0 To 1, 0 To UBound(ilMBSave, 2) + 1) As Integer
                For ilCol = 1 To MPBAMOUNTINDEX Step 1
                    slMBShow(ilCol, UBound(slMBShow, 2)) = smMBShow(ilCol, ilRowNo)
                Next ilCol
                'ReDim Preserve slMBShow(1 To MPBAMOUNTINDEX, 1 To UBound(slMBShow, 2) + 1) As String
                ReDim Preserve slMBShow(0 To MPBAMOUNTINDEX, 0 To UBound(slMBShow, 2) + 1) As String
            End If
        Next ilRowNo
        slTBilledDollars = "0.00"
        'ReDim smMBSave(1 To 3, 1 To UBound(slMBSave, 2)) As String
        'ReDim imMBSave(1 To 1, 1 To UBound(ilMBSave, 2)) As Integer
        'ReDim smMBShow(1 To MPBAMOUNTINDEX, 1 To UBound(slMBShow, 2)) As String
        ReDim smMBSave(0 To 3, 0 To UBound(slMBSave, 2)) As String
        'ReDim imMBSave(0 To UBound(ilMBSave, 2)) As Integer
        ReDim imMBSave(0 To 1, 0 To UBound(ilMBSave, 2)) As Integer
        ReDim smMBShow(0 To MPBAMOUNTINDEX, 0 To UBound(slMBShow, 2)) As String
        For ilRowNo = 1 To UBound(slMBSave, 2) - 1 Step 1
            slTBilledDollars = gAddStr(slTBilledDollars, slMBSave(2, ilRowNo))
            smMBSave(1, ilRowNo) = slMBSave(1, ilRowNo)
            smMBSave(2, ilRowNo) = slMBSave(2, ilRowNo)
            smMBSave(3, ilRowNo) = slMBSave(3, ilRowNo)
            imMBSave(1, ilRowNo) = ilMBSave(1, ilRowNo)
            'imMBSave(2, ilRowNo) = ilMBSave(2, ilRowNo)
            For ilCol = 1 To MPBAMOUNTINDEX Step 1
                smMBShow(ilCol, ilRowNo) = slMBShow(ilCol, ilRowNo)
            Next ilCol
        Next ilRowNo
        ilLastBilledIndex = UBound(slMBSave, 2) - 1
        ilUpperLimit = UBound(smMBShow, 2)
        If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
            slPercent = edcMPercent.Text
            slPercent = gDivStr(gMulStr(slPercent, 100#), smLnTNet)
        Else
            slPercent = edcMPercent.Text
        End If
            'Vehicle
        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            slLineType = mGetLineType(ilClf + 1)
            'Ignore package lines
            If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                If ((tgClfCntr(ilClf).iStatus = 0) Or (tgClfCntr(ilClf).iStatus = 1)) And (Not tgClfCntr(ilClf).iCancel) Then
                    slRecCode = Trim$(str$(tgClfCntr(ilClf).ClfRec.iVefCode))
                    For ilTest = 0 To lbcVehicle.ListCount - 1 Step 1  'Contract!lbcVehicle.ListCount - 1 Step 1
                        slNameCode = lbcVehicle.List(ilTest)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If slRecCode = slCode Then
                            'Build array of dollars by month
                            slATotal = "0"
                            llSDate = 0
                            llEDate = 0
                            ilFirstIndex = -1
                            ilCff = tgClfCntr(ilClf).iFirstCff
                            Do While ilCff <> -1
                                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                    gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate    'Week Start date
                                    gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate    'Week Start date
                                    If gDateValue(slStartDate) <= gDateValue(slEndDate) Then
                                        If gDateValue(slEndDate) > lmLastBilledDate Then
                                            If gDateValue(slStartDate) <= lmLastBilledDate Then
                                                slStartDate = Format$(lmLastBilledDate + 1, "m/d/yy")
                                            End If
                                            If llSDate = 0 Then
                                                llSDate = gDateValue(slStartDate)
                                                llEDate = gDateValue(slEndDate)
                                            Else
                                                If gDateValue(slStartDate) < llSDate Then
                                                    llSDate = gDateValue(slStartDate)
                                                End If
                                                If gDateValue(slEndDate) > llEDate Then
                                                    llEDate = gDateValue(slEndDate)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                ilCff = tgCffCntr(ilCff).iNextCff
                            Loop
                            llTDollars = 0
                            For llWDate = llSDate To llEDate Step 7
                                ilSpots = mGetFlightSpots(ilClf + 1, llWDate, llPrice)
                                llTDollars = llTDollars + ilSpots * llPrice
                            Next llWDate
                            slTDollars = gLongToStrDec(llTDollars, 2)
                            ''Remove pasted billed dollars
                            'For ilRowNo = 1 To ilLastBilledIndex Step 1
                            '    If (imMBSave(1, ilRowNo) = ilTest) Then
                            '        slTDollars = gSubStr(slTDollars, smMBSave(2, ilRowNo))
                            '    End If
                            'Next ilRowNo
                            slDate = lbcMESDate.List(lbcMESDate.ListIndex)
                            llEDate = gDateValue(slDate)
                            ilNoPds = 0
                            For ilPds = lbcMSSDate.ListIndex To lbcMSSDate.ListCount - 1 Step 1
                                ilNoPds = ilNoPds + 1
                                slStr = lbcMSSDate.List(ilPds)
                                If llEDate = gDateValue(slStr) Then
                                    Exit For
                                End If
                            Next ilPds
                            slDate = lbcMSSDate.List(lbcMSSDate.ListIndex)
                            llDate = gDateValue(slDate)
                            llSDate = llDate - 1
                            slDate = Format$(llSDate, "m/d/yy")
                            If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                                slDate = gObtainStartCal(slDate)
                            ElseIf (smBillCycle = "W") Then
                                slDate = gObtainPrevMonday(slDate)
                            Else
                                slDate = gObtainStartStd(slDate)
                            End If
                            llSDate = gDateValue(slDate)
                            llSDollars = 0
                            For llWDate = llSDate To llEDate Step 7
                                ilSpots = mGetFlightSpots(ilClf + 1, llWDate, llPrice)
                                llSDollars = llSDollars + ilSpots * llPrice
                            Next llWDate
                            If llSDollars > 0 Then
                                slSDollars = gLongToStrDec(llSDollars, 2)
                                ilSZero = False
                            Else
                                slSDollars = slTDollars
                                ilSZero = True
                            End If
                            slDate = lbcMESDate.List(lbcMESDate.ListIndex)
                            llEDate = gDateValue(slDate)
                            slDate = lbcMSSDate.List(lbcMSSDate.ListIndex)
                            llDate = gDateValue(slDate)
                            Do While llDate <= llEDate
                                llSDate = llDate - 1
                                slDate = Format$(llSDate, "m/d/yy")
                                If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                                    slDate = gObtainStartCal(slDate)
                                ElseIf (smBillCycle = "W") Then
                                    slDate = gObtainPrevMonday(slDate)
                                Else
                                    slDate = gObtainStartStd(slDate)
                                End If
                                llSDate = gDateValue(slDate)
                                'Test if air week is within billing period
                                'To compute billing for the month, take the
                                'dollars to be billed for the month and multiple by %
                                ilWSpots = 0
                                llWDollars = 0
                                For llWDate = llSDate To llDate Step 7
                                    ilSpots = mGetFlightSpots(ilClf + 1, llWDate, llPrice)
                                    llWDollars = llWDollars + ilSpots * llPrice
                                Next llWDate
                                'Vehicle
                                imMBSave(1, ilUpperLimit) = ilTest
                                slStr = lbcPSVehicle.List(ilTest)
                                gSetShow pbcMP, slStr, tmMPBCtrls(MPBVEHICLEINDEX)
                                smMBShow(MPBVEHICLEINDEX, ilUpperLimit) = tmMPBCtrls(MPBVEHICLEINDEX).sShow
                                'Transaction
                                smMBSave(1, ilUpperLimit) = "M"
                                'Date
                                slStr = Format$(llDate, "m/d/yy")
                                'gFindMatch slStr, 0, lbcBDate
                                'If gLastFound(lbcBDate) >= 0 Then
                                '    imMBSave(2, ilUpperLimit) = gLastFound(lbcBDate)
                                'Else
                                'End If
                                smMBSave(3, ilUpperLimit) = slStr
                                gSetShow pbcMP, slStr, tmMPBCtrls(MPBDATEINDEX)
                                smMBShow(MPBDATEINDEX, ilUpperLimit) = tmMPBCtrls(MPBDATEINDEX).sShow
                                'Amount
                                If (llWDollars = 0) And (ilSZero) And (ilNoPds > 0) Then
                                    llWDollars = llTDollars / ilNoPds
                                End If
                                slWDollars = gLongToStrDec(llWDollars, 2)
                                If smAgyRate <> "" Then
                                    slWDollars = gDivStr(gMulStr(slWDollars, gSubStr("100.00", smAgyRate)), "100.00")
                                End If
                                slWDollars = gDivStr(gMulStr(slWDollars, slTDollars), slSDollars)
                                slStr = gDivStr(gMulStr(slWDollars, slPercent), "100")
                                smMBSave(2, ilUpperLimit) = gRoundStr(slStr, ".01", 2)
                                slATotal = gAddStr(slATotal, smMBSave(2, ilUpperLimit))
                                ilFound = False
                                For ilLoop = ilLastBilledIndex + 1 To ilUpperLimit - 1 Step 1
                                    'If (imMBSave(1, ilLoop) = imMBSave(1, ilUpperLimit)) And (imMBSave(2, ilLoop) = imMBSave(2, ilUpperLimit)) Then
                                    If (imMBSave(1, ilLoop) = imMBSave(1, ilUpperLimit)) And (gDateValue(smMBSave(3, ilLoop)) = gDateValue(smMBSave(3, ilUpperLimit))) Then
                                        ilFound = True
                                        If ilFirstIndex = -1 Then
                                            ilFirstIndex = ilLoop
                                        End If
                                        smMBSave(2, ilLoop) = gAddStr(smMBSave(2, ilLoop), smMBSave(2, ilUpperLimit))
                                        gFormatStr smMBSave(2, ilLoop), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                                        gSetShow pbcMP, slStr, tmMPBCtrls(MPBAMOUNTINDEX)
                                        smMBShow(MPBAMOUNTINDEX, ilLoop) = tmMPBCtrls(MPBAMOUNTINDEX).sShow
                                        Exit For
                                    End If
                                Next ilLoop
                                If Not ilFound Then
                                    If ilFirstIndex = -1 Then
                                        ilFirstIndex = ilUpperLimit
                                    End If
                                    gFormatStr smMBSave(2, ilUpperLimit), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                                    gSetShow pbcMP, slStr, tmMPBCtrls(MPBAMOUNTINDEX)
                                    smMBShow(MPBAMOUNTINDEX, ilUpperLimit) = tmMPBCtrls(MPBAMOUNTINDEX).sShow
                                    ilUpperLimit = ilUpperLimit + 1
                                    'ReDim Preserve smMBSave(1 To 3, 1 To ilUpperLimit) As String
                                    'ReDim Preserve imMBSave(1 To 1, 1 To ilUpperLimit) As Integer
                                    'ReDim Preserve smMBShow(1 To MPBAMOUNTINDEX, 1 To ilUpperLimit) As String
                                    ReDim Preserve smMBSave(0 To 3, 0 To ilUpperLimit) As String
                                    'ReDim Preserve imMBSave(0 To ilUpperLimit) As Integer
                                    ReDim Preserve imMBSave(0 To 1, 0 To ilUpperLimit) As Integer
                                    ReDim Preserve smMBShow(0 To MPBAMOUNTINDEX, 0 To ilUpperLimit) As String
                                End If
                                llDate = llDate + 1
                                slDate = Format$(llDate, "m/d/yy")
                                If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                                    slDate = gObtainEndCal(slDate)
                                ElseIf (smBillCycle = "W") Then
                                    slDate = gObtainNextSunday(slDate)
                                Else
                                    slDate = gObtainEndStd(slDate)
                                End If
                                llDate = gDateValue(slDate)
                            Loop
                            'Adjust dollars
                            If smAgyRate <> "" Then
                                slTDollars = gDivStr(gMulStr(slTDollars, gSubStr("100.00", smAgyRate)), "100.00")
                            End If
                            slBTotal = gDivStr(gMulStr(slTDollars, slPercent), "100")
                            slBTotal = gRoundStr(slBTotal, ".01", 2)
                            slDiffTotal = gSubStr(slBTotal, slATotal)
                            If ilFirstIndex <> -1 Then
                        If gStrDecToLong(slDiffTotal, 2) <> 0 Then
                        ilRet = ilRet
                        End If
                                smMBSave(2, ilFirstIndex) = gAddStr(smMBSave(2, ilFirstIndex), slDiffTotal)
                                gFormatStr smMBSave(2, ilFirstIndex), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                                gSetShow pbcMP, slStr, tmMPBCtrls(MPBAMOUNTINDEX)
                                smMBShow(MPBAMOUNTINDEX, ilFirstIndex) = tmMPBCtrls(MPBAMOUNTINDEX).sShow
                            End If
                            Exit For
                        End If
                    Next ilTest
                End If
            End If
        Next ilClf
        If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
            slBTotal = edcMPercent.Text
        Else
            slBTotal = gDivStr(gMulStr(smLnTNet, slPercent), "100")
            slBTotal = gRoundStr(slBTotal, ".01", 2)
        End If
        slATotal = ".00"
        For ilLoop = LBONE To UBound(smMBSave, 2) - 1 Step 1
            slATotal = gAddStr(slATotal, smMBSave(2, ilLoop))
        Next ilLoop
        slATotal = gRoundStr(slATotal, ".01", 2)
        slDiffTotal = gSubStr(slBTotal, slATotal)
        llTest = 0
        Do
            For ilLoop = LBONE To UBound(smMBSave, 2) - 1 Step 1
                If gCompNumberStr(slDiffTotal, ".00") = 0 Then
                    Exit For
                End If
                If ilLoop > ilLastBilledIndex Then
                    If gCompNumberStr(smMBSave(2, ilLoop), ".00") > 0 Then
                        If gCompNumberStr(slDiffTotal, ".00") > 0 Then
                            smMBSave(2, ilLoop) = gAddStr(smMBSave(2, ilLoop), ".01")
                            slDiffTotal = gSubStr(slDiffTotal, ".01")
                        Else
                            smMBSave(2, ilLoop) = gSubStr(smMBSave(2, ilLoop), ".01")
                            slDiffTotal = gAddStr(slDiffTotal, ".01")
                        End If
                        gFormatStr smMBSave(2, ilLoop), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                        gSetShow pbcMP, slStr, tmMPBCtrls(MPBAMOUNTINDEX)
                        smMBShow(MPBAMOUNTINDEX, ilLoop) = tmMPBCtrls(MPBAMOUNTINDEX).sShow
                    End If
                End If
            Next ilLoop
            llTest = llTest + 1
            If llTest > 100000 Then
                Exit Do
            End If
        Loop While gCompNumberStr(slDiffTotal, ".00") <> 0
        'mInitNewMB UBound(smMBSave, 2)
        vbcMP.Min = LBONE   'LBound(smMBShow, 2)
        If UBound(smMBShow, 2) - 1 <= vbcMP.LargeChange + 1 Then
            vbcMP.Max = LBONE   'LBound(smMBShow, 2)
        Else
            vbcMP.Max = UBound(smMBShow, 2) - vbcMP.LargeChange - 1
        End If
        vbcMP.Value = vbcMP.Min
        imMBChg = True
    ElseIf rbcOption(3).Value Then  'Promotion
        'Retain previous billed dollars and weeks
        'ReDim slPBSave(1 To 3, 1 To 1) As String
        'ReDim ilPBSave(1 To 1, 1 To 1) As Integer
        'ReDim slPBShow(1 To MPBAMOUNTINDEX, 1 To 1) As String
        ReDim slPBSave(0 To 3, 0 To 1) As String
        'ReDim ilPBSave(0 To 1) As Integer
        ReDim ilPBSave(0 To 1, 0 To 1) As Integer
        ReDim slPBShow(0 To MPBAMOUNTINDEX, 0 To 1) As String
        For ilRowNo = LBONE To UBound(smPBSave, 2) - 1 Step 1
            slDate = smPBSave(3, ilRowNo)   'lbcBDate.List(imPBSave(2, ilRowNo))
            If gDateValue(slDate) <= lmLastBilledDate Then
                slPBSave(1, UBound(slPBSave, 2)) = smPBSave(1, ilRowNo)
                slPBSave(2, UBound(slPBSave, 2)) = smPBSave(2, ilRowNo)
                slPBSave(3, UBound(slPBSave, 2)) = smPBSave(3, ilRowNo)
                ilPBSave(1, UBound(slPBSave, 2)) = imPBSave(1, ilRowNo)
                'ilPBSave(2, UBound(slPBSave, 2)) = imPBSave(2, ilRowNo)
                'ReDim Preserve slPBSave(1 To 3, 1 To UBound(slPBSave, 2) + 1) As String
                'ReDim Preserve ilPBSave(1 To 1, 1 To UBound(ilPBSave, 2) + 1) As Integer
                ReDim Preserve slPBSave(0 To 3, 0 To UBound(slPBSave, 2) + 1) As String
                'ReDim Preserve ilPBSave(0 To UBound(ilPBSave, 2) + 1) As Integer
                ReDim Preserve ilPBSave(0 To 1, 0 To UBound(ilPBSave, 2) + 1) As Integer
                For ilCol = 1 To MPBAMOUNTINDEX Step 1
                    slPBShow(ilCol, UBound(slPBShow, 2)) = smPBShow(ilCol, ilRowNo)
                Next ilCol
                'ReDim Preserve slPBShow(1 To MPBAMOUNTINDEX, 1 To UBound(slPBShow, 2) + 1) As String
                ReDim Preserve slPBShow(0 To MPBAMOUNTINDEX, 0 To UBound(slPBShow, 2) + 1) As String
            End If
        Next ilRowNo
        slTBilledDollars = "0.00"
        'ReDim smPBSave(1 To 3, 1 To UBound(slPBSave, 2)) As String
        'ReDim imPBSave(1 To 1, 1 To UBound(ilPBSave, 2)) As Integer
        'ReDim smPBShow(1 To MPBAMOUNTINDEX, 1 To UBound(slPBShow, 2)) As String
        ReDim smPBSave(0 To 3, 0 To UBound(slPBSave, 2)) As String
        'ReDim imPBSave(0 To UBound(ilPBSave, 2)) As Integer
        ReDim imPBSave(0 To 1, 0 To UBound(ilPBSave, 2)) As Integer
        ReDim smPBShow(0 To MPBAMOUNTINDEX, 0 To UBound(slPBShow, 2)) As String
        For ilRowNo = 1 To UBound(slPBSave, 2) - 1 Step 1
            slTBilledDollars = gAddStr(slTBilledDollars, slPBSave(2, ilRowNo))
            smPBSave(1, ilRowNo) = slPBSave(1, ilRowNo)
            smPBSave(2, ilRowNo) = slPBSave(2, ilRowNo)
            smPBSave(3, ilRowNo) = slPBSave(3, ilRowNo)
            imPBSave(1, ilRowNo) = ilPBSave(1, ilRowNo)
            'imPBSave(2, ilRowNo) = ilPBSave(2, ilRowNo)
            For ilCol = 1 To MPBAMOUNTINDEX Step 1
                smPBShow(ilCol, ilRowNo) = slPBShow(ilCol, ilRowNo)
            Next ilCol
        Next ilRowNo
        ilLastBilledIndex = UBound(slPBSave, 2) - 1
        ilUpperLimit = UBound(smPBShow, 2)
        If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
            slPercent = edcPPercent.Text
            slPercent = gDivStr(gMulStr(slPercent, 100#), smLnTNet)
        Else
            slPercent = edcPPercent.Text
        End If
            'Vehicle
        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            slLineType = mGetLineType(ilClf + 1)
            'Ignore package lines
            If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                If ((tgClfCntr(ilClf).iStatus = 0) Or (tgClfCntr(ilClf).iStatus = 1)) And (Not tgClfCntr(ilClf).iCancel) Then
                    slRecCode = Trim$(str$(tgClfCntr(ilClf).ClfRec.iVefCode))
                    For ilTest = 0 To lbcVehicle.ListCount - 1 Step 1  'Contract!lbcVehicle.ListCount - 1 Step 1
                        slNameCode = lbcVehicle.List(ilTest)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If slRecCode = slCode Then
                            slATotal = "0"
                            llSDate = 0
                            llEDate = 0
                            ilFirstIndex = -1
                            ilCff = tgClfCntr(ilClf).iFirstCff
                            Do While ilCff <> -1
                                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                    gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate    'Week Start date
                                    gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate    'Week Start date
                                    If gDateValue(slStartDate) <= gDateValue(slEndDate) Then
                                        If gDateValue(slEndDate) > lmLastBilledDate Then
                                            If gDateValue(slStartDate) <= lmLastBilledDate Then
                                                slStartDate = Format$(lmLastBilledDate + 1, "m/d/yy")
                                            End If
                                            If llSDate = 0 Then
                                                llSDate = gDateValue(slStartDate)
                                                llEDate = gDateValue(slEndDate)
                                            Else
                                                If gDateValue(slStartDate) < llSDate Then
                                                    llSDate = gDateValue(slStartDate)
                                                End If
                                                If gDateValue(slEndDate) > llEDate Then
                                                    llEDate = gDateValue(slEndDate)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                ilCff = tgCffCntr(ilCff).iNextCff
                            Loop
                            llTDollars = 0
                            For llWDate = llSDate To llEDate Step 7
                                ilSpots = mGetFlightSpots(ilClf + 1, llWDate, llPrice)
                                llTDollars = llTDollars + ilSpots * llPrice
                            Next llWDate
                            slTDollars = gLongToStrDec(llTDollars, 2)
                            ''Remove pasted billed dollars
                            'For ilRowNo = 1 To ilLastBilledIndex Step 1
                            '    If (imPBSave(1, ilRowNo) = ilTest) Then
                            '        slTDollars = gSubStr(slTDollars, smMBSave(2, ilRowNo))
                            '    End If
                            'Next ilRowNo

                            slDate = lbcPESDate.List(lbcPESDate.ListIndex)
                            llEDate = gDateValue(slDate)
                            ilNoPds = 0
                            For ilPds = lbcPSSDate.ListIndex To lbcPSSDate.ListCount - 1 Step 1
                                ilNoPds = ilNoPds + 1
                                slStr = lbcPSSDate.List(ilPds)
                                If llEDate = gDateValue(slStr) Then
                                    Exit For
                                End If
                            Next ilPds
                            slDate = lbcPSSDate.List(lbcPSSDate.ListIndex)
                            llDate = gDateValue(slDate)
                            llSDate = llDate - 1
                            slDate = Format$(llSDate, "m/d/yy")
                            If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                                slDate = gObtainStartCal(slDate)
                            ElseIf (smBillCycle = "W") Then
                                slDate = gObtainPrevMonday(slDate)
                            Else
                                slDate = gObtainStartStd(slDate)
                            End If
                            llSDate = gDateValue(slDate)
                            llSDollars = 0
                            For llWDate = llSDate To llEDate Step 7
                                ilSpots = mGetFlightSpots(ilClf + 1, llWDate, llPrice)
                                llSDollars = llSDollars + ilSpots * llPrice
                            Next llWDate
                            If llSDollars > 0 Then
                                slSDollars = gLongToStrDec(llSDollars, 2)
                                ilSZero = False
                            Else
                                slSDollars = slTDollars
                                ilSZero = True
                            End If
                            slDate = lbcPESDate.List(lbcPESDate.ListIndex)
                            llEDate = gDateValue(slDate)
                            slDate = lbcPSSDate.List(lbcPSSDate.ListIndex)
                            llDate = gDateValue(slDate)
                            Do While llDate <= llEDate
                                llSDate = llDate - 1
                                slDate = Format$(llSDate, "m/d/yy")
                                If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                                    slDate = gObtainStartCal(slDate)
                                ElseIf (smBillCycle = "W") Then
                                    slDate = gObtainPrevMonday(slDate)
                                Else
                                    slDate = gObtainStartStd(slDate)
                                End If
                                llSDate = gDateValue(slDate)
                                'Test if air week is within billing period
                                'To compute billing for the month, take the
                                'dollars to be billed for the month and multiple by %
                                ilWSpots = 0
                                llWDollars = 0
                                For llWDate = llSDate To llDate Step 7
                                    ilSpots = mGetFlightSpots(ilClf + 1, llWDate, llPrice)
                                    llWDollars = llWDollars + ilSpots * llPrice
                                Next llWDate
                                'Vehicle
                                imPBSave(1, ilUpperLimit) = ilTest
                                slStr = lbcPSVehicle.List(ilTest)
                                gSetShow pbcMP, slStr, tmMPBCtrls(MPBVEHICLEINDEX)
                                smPBShow(MPBVEHICLEINDEX, ilUpperLimit) = tmMPBCtrls(MPBVEHICLEINDEX).sShow
                                'Transaction
                                smPBSave(1, ilUpperLimit) = "M"
                                'Date
                                slStr = Format$(llDate, "m/d/yy")
                                'gFindMatch slStr, 0, lbcBDate
                                'If gLastFound(lbcBDate) >= 0 Then
                                '    imPBSave(2, ilUpperLimit) = gLastFound(lbcBDate)
                                'Else
                                'End If
                                smPBSave(3, ilUpperLimit) = slStr
                                gSetShow pbcMP, slStr, tmMPBCtrls(MPBDATEINDEX)
                                smPBShow(MPBDATEINDEX, ilUpperLimit) = tmMPBCtrls(MPBDATEINDEX).sShow
                                'Amount
                                If (llWDollars = 0) And (ilSZero) And (ilNoPds > 0) Then
                                    llWDollars = llTDollars / ilNoPds
                                End If
                                slWDollars = gLongToStrDec(llWDollars, 2)
                                If smAgyRate <> "" Then
                                    slWDollars = gDivStr(gMulStr(slWDollars, gSubStr("100.00", smAgyRate)), "100.00")
                                End If
                                slWDollars = gDivStr(gMulStr(slWDollars, slTDollars), slSDollars)
                                slStr = gDivStr(gMulStr(slWDollars, slPercent), "100")
                                smPBSave(2, ilUpperLimit) = gRoundStr(slStr, ".01", 2)
                                slATotal = gAddStr(slATotal, smPBSave(2, ilUpperLimit))
                                ilFound = False
                                For ilLoop = ilLastBilledIndex + 1 To ilUpperLimit - 1 Step 1
                                    'If (imPBSave(1, ilLoop) = imPBSave(1, ilUpperLimit)) And (imPBSave(2, ilLoop) = imPBSave(2, ilUpperLimit)) Then
                                    If (imPBSave(1, ilLoop) = imPBSave(1, ilUpperLimit)) And (gDateValue(smPBSave(3, ilLoop)) = gDateValue(smPBSave(3, ilUpperLimit))) Then
                                        ilFound = True
                                        If ilFirstIndex = -1 Then
                                            ilFirstIndex = ilLoop
                                        End If
                                        smPBSave(2, ilLoop) = gAddStr(smPBSave(2, ilLoop), smPBSave(2, ilUpperLimit))
                                        gFormatStr smPBSave(2, ilLoop), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                                        gSetShow pbcMP, slStr, tmMPBCtrls(MPBAMOUNTINDEX)
                                        smPBShow(MPBAMOUNTINDEX, ilLoop) = tmMPBCtrls(MPBAMOUNTINDEX).sShow
                                        Exit For
                                    End If
                                Next ilLoop
                                If Not ilFound Then
                                    If ilFirstIndex = -1 Then
                                        ilFirstIndex = ilUpperLimit
                                    End If
                                    gFormatStr smPBSave(2, ilUpperLimit), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                                    gSetShow pbcMP, slStr, tmMPBCtrls(MPBAMOUNTINDEX)
                                    smPBShow(MPBAMOUNTINDEX, ilUpperLimit) = tmMPBCtrls(MPBAMOUNTINDEX).sShow
                                    ilUpperLimit = ilUpperLimit + 1
                                    'ReDim Preserve smPBSave(1 To 3, 1 To ilUpperLimit) As String
                                    'ReDim Preserve imPBSave(1 To 1, 1 To ilUpperLimit) As Integer
                                    'ReDim Preserve smPBShow(1 To MPBAMOUNTINDEX, 1 To ilUpperLimit) As String
                                    ReDim Preserve smPBSave(0 To 3, 0 To ilUpperLimit) As String
                                    'ReDim Preserve imPBSave(0 To ilUpperLimit) As Integer
                                    ReDim Preserve imPBSave(0 To 1, 0 To ilUpperLimit) As Integer
                                    ReDim Preserve smPBShow(0 To MPBAMOUNTINDEX, 0 To ilUpperLimit) As String
                                End If
                                llDate = llDate + 1
                                slDate = Format$(llDate, "m/d/yy")
                                If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                                    slDate = gObtainEndCal(slDate)
                                ElseIf (smBillCycle = "W") Then
                                    slDate = gObtainNextSunday(slDate)
                                Else
                                    slDate = gObtainEndStd(slDate)
                                End If
                                llDate = gDateValue(slDate)
                            Loop
                            'Adjust dollars
                            If smAgyRate <> "" Then
                                slTDollars = gDivStr(gMulStr(slTDollars, gSubStr("100.00", smAgyRate)), "100.00")
                            End If
                            slBTotal = gDivStr(gMulStr(slTDollars, slPercent), "100")
                            slBTotal = gRoundStr(slBTotal, ".01", 2)
                            slDiffTotal = gSubStr(slBTotal, slATotal)
                            If ilFirstIndex <> -1 Then
                                smPBSave(2, ilFirstIndex) = gAddStr(smPBSave(2, ilFirstIndex), slDiffTotal)
                                gFormatStr smPBSave(2, ilFirstIndex), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                                gSetShow pbcMP, slStr, tmMPBCtrls(MPBAMOUNTINDEX)
                                smPBShow(MPBAMOUNTINDEX, ilFirstIndex) = tmMPBCtrls(MPBAMOUNTINDEX).sShow
                            End If
                            Exit For
                        End If
                    Next ilTest
                End If
            End If
        Next ilClf
        If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
            slBTotal = edcPPercent.Text
        Else
            slBTotal = gDivStr(gMulStr(smLnTNet, slPercent), "100")
            slBTotal = gRoundStr(slBTotal, ".01", 2)
        End If
        slATotal = ".00"
        For ilLoop = LBONE To UBound(smPBSave, 2) - 1 Step 1
            slATotal = gAddStr(slATotal, smPBSave(2, ilLoop))
        Next ilLoop
        slATotal = gRoundStr(slATotal, ".01", 2)
        slDiffTotal = gSubStr(slBTotal, slATotal)
        llTest = 0
        Do
            For ilLoop = LBONE To UBound(smPBSave, 2) - 1 Step 1
                If gCompNumberStr(slDiffTotal, ".00") = 0 Then
                    Exit For
                End If
                If ilLoop > ilLastBilledIndex Then
                    If gCompNumberStr(smPBSave(2, ilLoop), ".00") > 0 Then
                        If gCompNumberStr(slDiffTotal, ".00") > 0 Then
                            smPBSave(2, ilLoop) = gAddStr(smPBSave(2, ilLoop), ".01")
                            slDiffTotal = gSubStr(slDiffTotal, ".01")
                        Else
                            smPBSave(2, ilLoop) = gSubStr(smPBSave(2, ilLoop), ".01")
                            slDiffTotal = gAddStr(slDiffTotal, ".01")
                        End If
                        gFormatStr smPBSave(2, ilLoop), FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                        gSetShow pbcMP, slStr, tmMPBCtrls(MPBAMOUNTINDEX)
                        smPBShow(MPBAMOUNTINDEX, ilLoop) = tmMPBCtrls(MPBAMOUNTINDEX).sShow
                    End If
                End If
            Next ilLoop
            llTest = llTest + 1
            If llTest > 100000 Then
                Exit Do
            End If
        Loop While gCompNumberStr(slDiffTotal, ".00") <> 0
        'mInitNewMB UBound(smPBSave, 2)
        vbcMP.Min = LBONE   'LBound(smPBShow, 2)
        If UBound(smPBShow, 2) - 1 <= vbcMP.LargeChange + 1 Then
            vbcMP.Max = LBONE   'LBound(smPBShow, 2)
        Else
            vbcMP.Max = UBound(smPBShow, 2) - vbcMP.LargeChange - 1
        End If
        vbcMP.Value = vbcMP.Min
        imPBChg = True
    End If
    If ilPaint Then
        If rbcOption(0).Value Then
            imFBRowNo = 1
            pbcFix.Cls
            pbcFix_Paint
        ElseIf rbcOption(2).Value Then
            pbcMP.Cls
            pbcMP_Paint
        ElseIf rbcOption(3).Value Then
            pbcMP.Cls
            pbcMP_Paint
        End If
    End If
    mSetCommands
    Exit Sub

    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFBEnableBox                    *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mFBEnableBox(ilBoxNo As Integer)
'
'   mFBEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmFBCtrls)) Then
        Exit Sub
    End If

    If (imFBRowNo < vbcFix.Value) Or (imFBRowNo >= vbcFix.Value + vbcFix.LargeChange + 1) Then
        mFBSetShow ilBoxNo
        pbcArrow.Visible = False
        Exit Sub
    End If
    pbcArrow.Move plcFix.Left - pbcArrow.Width - 15, plcFix.Top + tmFBCtrls(FBVEHICLEINDEX).fBoxY + (imFBRowNo - vbcFix.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case FBVEHICLEINDEX 'Vehicle
            gShowHelpMess tmSbfHelp(), SBFBODYFACIL
'            lbcBVehicle.Height = gListBoxHeight(lbcBVehicle.ListCount, 15)
'            edcDropDown.Width = tmFSCtrls(FBVEHICLEINDEX).fBoxW
'            edcDropDown.MaxLength = 10
'            gMoveTableCtrl pbcFix, edcDropDown, tmFBCtrls(FBVEHICLEINDEX).fBoxX, tmFBCtrls(FBVEHICLEINDEX).fBoxY + (imFBRowNo - vbcFix.Value) * (fgBoxGridH + 15)
'            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
'            If imFBSave(1, imFBRowNo) >= 0 Then
'                edcDropDown.Text = lbcBVehicle.List(imFBSave(1, imFBRowNo))
'                lbcBVehicle.ListIndex = imFBSave(1, imFBRowNo)
'            Else
'                If imFBRowNo > 1 Then
'                    edcDropDown.Text = lbcBVehicle.List(imFBSave(1, imFBRowNo - 1))
'                    lbcBVehicle.ListIndex = imFBSave(1, imFBRowNo - 1)
'                Else
'                    slNameCode = Contract!lbcRateCard.List(Contract!lbcRateCard.ListIndex)
'                    ilRet = gParseItem(slNameCode, 2, "/", slName)
'                    If ilRet <> CP_MSG_NONE Then
'                        slName = sgUserDefVehicleName
'                    Else
'                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                        If InStr(slCode, "-") <> 0 Then
'                            slName = sgUserDefVehicleName
'                        End If
'                    End If
'                    gFndFirst lbcBVehicle,  slName
'                    If gLastFound(lbcBVehicle) >= 0 Then
'                        lbcBVehicle.ListIndex = gLastFound(lbcBVehicle)
'                        edcDropDown.Text = lbcBVehicle.List(lbcBVehicle.ListIndex)
'                    Else
'                        lbcBVehicle.ListIndex = 0
'                        edcDropDown.Text = lbcBVehicle.List(0)
'                    End If
'                End If
'            End If
'            If imFBRowNo - vbcFix.Value <= vbcFix.LargeChange \ 2 Then
'                lbcBVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
'            Else
'                lbcBVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcBVehicle.Height
'            End If
'            imComboBoxIndex = lbcBVehicle.ListIndex
'            edcDropDown.SelStart = 0
'            edcDropDown.SelLength = Len(edcDropDown.Text)
'            edcDropDown.Visible = True
'            cmcDropDown.Visible = True
'            edcDropDown.SetFocus
        Case FBDATEINDEX 'Program index
            gShowHelpMess tmSbfHelp(), SBFBODYDATE
'            lbcBDate.Height = gListBoxHeight(lbcBDate.ListCount, 15)
'            edcDropDown.Width = tmFBCtrls(FBDATEINDEX).fBoxW - cmcDropDown.Width
'            edcDropDown.MaxLength = 10
'            gMoveTableCtrl pbcFix, edcDropDown, tmFBCtrls(FBDATEINDEX).fBoxX, tmFBCtrls(FBDATEINDEX).fBoxY + (imFBRowNo - vbcFix.Value) * (fgBoxGridH + 15)
'            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
'            lbcBDate.ListIndex = imFBSave(2, imFBRowNo)
'            If lbcBDate.ListIndex < 0 Then
'                If imFBRowNo <= 1 Then
'                    lbcBDate.ListIndex = 0
'                    edcDropDown.Text = lbcBDate.List(0)
'                Else
'                    lbcBDate.ListIndex = imFBSave(2, imFBRowNo - 1) + 1
'                    edcDropDown.Text = lbcBDate.List(lbcBDate.ListIndex)
'                End If
'            Else
'                edcDropDown.Text = lbcBDate.List(lbcBDate.ListIndex)
'            End If
'            If imFBRowNo - vbcFix.Value <= vbcFix.LargeChange \ 2 Then
'                lbcBDate.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
'            Else
'                lbcBDate.Move edcDropDown.Left, edcDropDown.Top - lbcBDate.Height
'            End If
'            edcDropDown.SelStart = 0
'            edcDropDown.SelLength = Len(edcDropDown.Text)
'            edcDropDown.Visible = True
'            cmcDropDown.Visible = True
'            edcDropDown.SetFocus
        'Case FBBILLINDEX
        '    gShowHelpMess tmSbfHelp(), SBFBODYBILL
        Case FBBILLINGINDEX
            edcAmount.Width = tmFBCtrls(FBBILLINGINDEX).fBoxW
            gMoveTableCtrl pbcFix, edcAmount, tmFBCtrls(FBBILLINGINDEX).fBoxX, tmFBCtrls(FBBILLINGINDEX).fBoxY + (imFBRowNo - vbcFix.Value) * (fgBoxGridH + 15)
            edcAmount.Text = gLongToStrDec(tmInstallBillInfo(imFBRowNo - 1).lBillDollars, 2)    'smFBSave(2, imFBRowNo)
            edcAmount.Visible = True  'Set visibility
            edcAmount.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFBSetShow                      *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mFBSetShow(ilBoxNo As Integer)
'
'   mFBSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim llDollars As Long

    pbcArrow.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmFBCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case FBVEHICLEINDEX 'Vehicle
'            lbcBVehicle.Visible = False
'            edcDropDown.Visible = False
'            cmcDropDown.Visible = False
'            slStr = edcDropDown.Text
'            gSetShow pbcFix, slStr, tmFBCtrls(ilBoxNo)
'            smPBShow(FBVEHICLEINDEX, imFBRowNo) = tmFBCtrls(ilBoxNo).sShow
'            If imPBSave(1, imFBRowNo) <> lbcBVehicle.ListIndex Then
'                imPBSave(1, imFBRowNo) = lbcBVehicle.ListIndex
'                If imFBRowNo < UBound(imPBSave, 2) Then  'New lines set after all fields entered
'                    imPBChg = True
'                Else
'                    If smPBSave(1, imFBRowNo) = "" Then
'                        smPBSave(1, imFBRowNo) = "C"
'                        smPBShow(PBTRANINDEX, imFBRowNo) = "Contract"
'                        gPaintArea pbcFix, tmFBCtrls(PBTRANINDEX).fBoxX, tmFBCtrls(PBTRANINDEX).fBoxY + (imFBRowNo - vbcFix.Value) * (fgBoxGridH + 15), tmFBCtrls(PBTRANINDEX).fBoxW - 15, tmFBCtrls(PBTRANINDEX).fBoxH - 15, WHITE
'                        pbcFix.CurrentX = tmFBCtrls(PBTRANINDEX).fBoxX + fgBoxInsetX
'                        pbcFix.CurrentY = tmFBCtrls(PBTRANINDEX).fBoxY + (imFBRowNo - vbcFix.Value) * (fgBoxGridH + 15) - 30'+ fgBoxInsetY
'                        pbcFix.Print smPBShow(PBTRANINDEX, imFBRowNo)
'                    End If
'                End If
'            End If
        Case FBDATEINDEX 'Program index
'            lbcBDate.Visible = False
'            edcDropDown.Visible = False
'            cmcDropDown.Visible = False
'            slStr = edcDropDown.Text
'            gSetShow pbcFix, slStr, tmFBCtrls(ilBoxNo)
'            smPBShow(FBDATEINDEX, imFBRowNo) = tmFBCtrls(ilBoxNo).sShow
'            If imPBSave(2, imFBRowNo) <> lbcBDate.ListIndex Then
'                imPBSave(2, imFBRowNo) = lbcBDate.ListIndex
'                If imFBRowNo < UBound(tmFBSbf) + 1 Then   'New lines set after all fields entered
'                    imPBChg = True
'                End If
'            End If
        'Case FBBILLINDEX
        Case FBBILLINGINDEX
            edcAmount.Visible = False
            slStr = edcAmount.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            gSetShow pbcFix, slStr, tmFBCtrls(ilBoxNo)
            'smFBShow(FBBILLINGINDEX, imFBRowNo) = tmFBCtrls(ilBoxNo).sShow

            llDollars = gStrDecToLong(edcAmount.Text, 2)
            'If gCompNumberStr(smFBSave(2, imFBRowNo), slStr) <> 0 Then
            If tmInstallBillInfo(imFBRowNo - 1).lBillDollars <> llDollars Then
                'If imFBRowNo < UBound(tmFBSbf) + 1 Then   'New lines set after all fields entered
                    imFBChg = True
                'End If
                'smFBSave(2, imFBRowNo) = edcAmount.Text
                tmInstallBillInfo(imFBRowNo - 1).lBillDollars = llDollars
                mFBTotals True
            End If
    End Select
    mSetCommands
    pbcFix.Cls
    pbcFix_Paint
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFBTestFields                   *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mFBTestFields() As Integer
'
'   iRet = mFBTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim slStr As String
    For llSbf = LBound(tmFBSbf) To UBound(tmFBSbf) - 1 Step 1
        If (tmFBSbf(llSbf).iStatus = 0) Or (tmFBSbf(llSbf).iStatus = 1) Then
            If tmFBSbf(llSbf).SbfRec.iBillVefCode <= 0 Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imFBRowNo = llSbf + 1
                imFBBoxNo = FBVEHICLEINDEX
                mFBTestFields = NO
                Exit Function
            End If
            If (tmFBSbf(llSbf).SbfRec.iDate(0) = 0) And (tmFBSbf(llSbf).SbfRec.iDate(1) = 0) Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("Date must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imFBRowNo = llSbf + 1
                imFBBoxNo = FBDATEINDEX
                mFBTestFields = NO
                Exit Function
            End If
            gStrToPDN "", 2, 5, slStr
            'If StrComp(tmFBSbf(llSbf).SbfRec.sItemAmount, slStr, 0) = 0 Then
            '    Screen.MousePointer = vbDefault
            '    ilRes = MsgBox("Price must be specified", vbOkOnly + vbExclamation, "Incomplete")
            '    imFBRowNo = llSbf + 1
            '    imFBBoxNo = FBBILLINGINDEX
            '    mFBTestFields = NO
            '    Exit Function
            'End If
        End If
    Next llSbf
    mFBTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFBTestSaveFields               *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mFBTestSaveFields() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRes                                                                                 *
'******************************************************************************************

'
'   iRet = mFBTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
'    If imFBSave(1, imFBRowNo) < 0 Then
'        ilRes = MsgBox("Vehicle must be specified", vbOkOnly + vbExclamation, "Incomplete")
'        imFBBoxNo = FBVEHICLEINDEX
'        mFBTestSaveFields = NO
'        Exit Function
'    End If
'    If imFBSave(2, imFBRowNo) < 0 Then
'        ilRes = MsgBox("Date must be specified", vbOkOnly + vbExclamation, "Incomplete")
'        imFBBoxNo = FBDATEINDEX
'        mFBTestSaveFields = NO
'        Exit Function
'    End If
'    If smFBSave(2, imFBRowNo) = "" Then
'        ilRes = MsgBox("Amount must be specified", vbOkOnly + vbExclamation, "Incomplete")
'        imFBBoxNo = FBBILLINGINDEX
'        mFBTestSaveFields = NO
'        Exit Function
'    End If
    mFBTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFBTotals                       *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute totals                  *
'*                                                     *
'*******************************************************
Private Sub mFBTotals(ilShowTotals As Integer)
    Dim slATotal As String
    Dim slBTotal As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slRvfTypeI_Adj As String          'Sum of AN and W- for Install contracts (rvfType = "I")
    Dim slRvfTypeA_Adj As String          'Sum of AN and W- for Install contracts (rvfType = "A")

    slRvfTypeI_Adj = smRvfTypeI_Adj
    slRvfTypeA_Adj = smRvfTypeA_Adj
    '5/8/08: Ignore the above adjustment values
    slRvfTypeI_Adj = "0.00"
    slRvfTypeA_Adj = "0.00"

    mInitInstallVeh

    slATotal = "0"
    slBTotal = gAddStr(smLnTGross, smNTRTGross)
    slBTotal = gAddStr(slBTotal, slRvfTypeA_Adj)
    'For ilLoop = LBound(smFBSave, 2) To UBound(smFBSave, 2) - 1 Step 1
    '    slATotal = gAddStr(slATotal, smFBSave(2, ilLoop))
    'Next ilLoop
    For ilLoop = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
        slATotal = gAddStr(slATotal, gLongToStrDec(tmInstallBillInfo(ilLoop).lBillDollars, 2))
    Next ilLoop
    smFixTGross = slATotal
    slATotal = gAddStr(slATotal, slRvfTypeI_Adj)
    If (gStrDecToLong(slATotal, 2) = 0) And (gStrDecToLong(slBTotal, 2) = 0) Then
        lacTotals.Visible = False
        Exit Sub
    End If
    If ilShowTotals Then
        lacTotals.ForeColor = BLACK
        If gStrDecToLong(slBTotal, 2) = gStrDecToLong(slATotal, 2) Then
            lacTotals.BackColor = GREEN
        Else
            lacTotals.BackColor = Red
            lacTotals.ForeColor = WHITE
        End If
        slBTotal = gAddStr(smLnTGross, smNTRTGross)
        slATotal = smFixTGross
        lacTotals.Visible = True
        gFormatStr slBTotal, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        slStr = "Totals: Gross $" & slStr
        If gStrDecToLong(slRvfTypeA_Adj, 2) <> 0 Then
            slStr = slStr & slRvfTypeA_Adj
            slBTotal = gAddStr(slBTotal, slRvfTypeA_Adj)
            slStr = slStr & "= " & slBTotal
        End If
        If slATotal <> "0" Then
            gFormatStr slATotal, FMTLEAVEBLANK + FMTCOMMA, 2, slATotal
            slStr = slStr & "  Billed $" & slATotal
        End If
        If gStrDecToLong(slRvfTypeI_Adj, 2) <> 0 Then
            slStr = slStr & slRvfTypeI_Adj
            slATotal = gAddStr(smFixTGross, slRvfTypeI_Adj)
            slStr = slStr & "= " & slATotal
        End If
        lacTotals.Caption = slStr
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFSEnableBox                    *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mFSEnableBox(ilBoxNo As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slNameCode                    slName                        slCode                    *
'*  ilRet                         slATotal                                                *
'******************************************************************************************

'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilCount As Integer
    Dim slStr As String

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmFSCtrls) - 1) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case FSSTARTDATEINDEX 'Vehicle
            lbcPSDate.height = gListBoxHeight(lbcPSDate.ListCount, 8)
            edcPSDropDown.Width = tmFSCtrls(FSSTARTDATEINDEX).fBoxW - cmcPSDropDown.Width
            edcPSDropDown.MaxLength = 10
            gMoveFormCtrl pbcFixSpec, edcPSDropDown, tmFSCtrls(FSSTARTDATEINDEX).fBoxX, tmFSCtrls(FSSTARTDATEINDEX).fBoxY
            cmcPSDropDown.Move edcPSDropDown.Left + edcPSDropDown.Width, edcPSDropDown.Top
            imChgMode = True
            If lbcPSDate.ListIndex < 0 Then
                slStartDate = Format$(lmCntrStartDate, "m/d/yy")
                If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                    slStartDate = gObtainEndCal(slStartDate)
                ElseIf (smBillCycle = "W") Then
                    slStartDate = gObtainNextSunday(slStartDate)
                Else
                    slStartDate = gObtainEndStd(slStartDate)
                End If
                'slStartDate = Format$(slStartDate, "mmm, yy")
                For ilLoop = 0 To lbcPSDate.ListCount - 1 Step 1
                    If gDateValue(slStartDate) = gDateValue(lbcPSDate.List(ilLoop)) Then
                        lbcPSDate.ListIndex = ilLoop 'Start at first end date of period after contract start
                        edcPSDropDown.Text = lbcPSDate.List(ilLoop)
                        Exit For
                    End If
                Next ilLoop
                If lbcPSDate.ListIndex < 0 Then
                    slStartDate = Format$(gNow(), "m/d/yy")
                    If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                        slStartDate = gObtainEndCal(slStartDate)
                    ElseIf (smBillCycle = "W") Then
                        slStartDate = gObtainNextSunday(slStartDate)
                    Else
                        slStartDate = gObtainEndStd(slStartDate)
                    End If
                    'slStartDate = Format$(slStartDate, "mmm, yy")
                    For ilLoop = 0 To lbcPSDate.ListCount - 1 Step 1
                        If gDateValue(slStartDate) = gDateValue(lbcPSDate.List(ilLoop)) Then
                            lbcPSDate.ListIndex = ilLoop 'Start at first end date of period after contract start
                            edcPSDropDown.Text = lbcPSDate.List(ilLoop)
                            Exit For
                        End If
                    Next ilLoop
                End If
                If lbcPSDate.ListIndex < 0 Then
                    lbcPSDate.ListIndex = 0 'Start at first end date of period after contract start
                    edcPSDropDown.Text = lbcPSDate.List(0)
                End If
            Else
                edcPSDropDown.Text = lbcPSDate.List(lbcPSDate.ListIndex)
            End If
            imChgMode = False
            lbcPSDate.Move edcPSDropDown.Left, edcPSDropDown.Top + edcPSDropDown.height
            edcPSDropDown.SelStart = 0
            edcPSDropDown.SelLength = Len(edcPSDropDown.Text)
            edcPSDropDown.Visible = True
            cmcPSDropDown.Visible = True
            edcPSDropDown.SetFocus
        Case FSNOMONTHSINDEX 'No Periods
            edcNoPeriods.Width = tmFSCtrls(FSNOMONTHSINDEX).fBoxW
            gMoveFormCtrl pbcFixSpec, edcNoPeriods, tmFSCtrls(FSNOMONTHSINDEX).fBoxX, tmFSCtrls(FSNOMONTHSINDEX).fBoxY
            If edcNoPeriods.Text = "" Then
                If lbcPSDate.ListIndex >= 0 Then
                    ilCount = 0
                    slEndDate = Format$(lmCntrEndDate, "m/d/yy")
                    If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                        slEndDate = gObtainEndCal(slEndDate)
                    ElseIf (smBillCycle = "W") Then
                        slEndDate = gObtainNextSunday(slEndDate)
                    Else
                        slEndDate = gObtainEndStd(slEndDate)
                    End If
                    'slEndDate = Format$(slEndDate, "mmm, yy")
                    For ilLoop = 0 To lbcPEDate.ListCount - 1 Step 1
                        If gDateValue(slEndDate) = gDateValue(lbcPEDate.List(ilLoop)) Then
                            edcNoPeriods.Text = Trim$(str$(ilCount + 1))
                            Exit For
                        End If
                        ilCount = ilCount + 1
                    Next ilLoop
                End If
            End If
            edcNoPeriods.Visible = True  'Set visibility
            edcNoPeriods.SetFocus
        Case FSENDDATEINDEX 'Date
            lbcPEDate.height = gListBoxHeight(lbcPEDate.ListCount, 8)
            edcPSDropDown.Width = tmFSCtrls(FSENDDATEINDEX).fBoxW - cmcPSDropDown.Width
            edcPSDropDown.MaxLength = 10
            gMoveFormCtrl pbcFixSpec, edcPSDropDown, tmFSCtrls(FSENDDATEINDEX).fBoxX, tmFSCtrls(FSENDDATEINDEX).fBoxY
            cmcPSDropDown.Move edcPSDropDown.Left + edcPSDropDown.Width, edcPSDropDown.Top
            imChgMode = True
            If lbcPEDate.ListIndex < 0 Then
                If lbcPSDate.ListIndex >= 0 Then
                    If edcNoPeriods.Text <> "" Then
                        slStr = lbcPSDate.List(lbcPSDate.ListIndex)
                        For ilLoop = 0 To lbcPEDate.ListCount - 1 Step 1
                            If gDateValue(slStr) = gDateValue(lbcPEDate.List(ilLoop)) Then
                                lbcPEDate.ListIndex = ilLoop + Val(edcNoPeriods.Text) - 1
                                Exit For
                            End If
                        Next ilLoop
                    Else
                        slEndDate = Format$(lmCntrEndDate, "m/d/yy")
                        If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
                            slEndDate = gObtainEndCal(slEndDate)
                        ElseIf (smBillCycle = "W") Then
                            slEndDate = gObtainNextSunday(slEndDate)
                        Else
                            slEndDate = gObtainEndStd(slEndDate)
                        End If
                        'slEndDate = Format$(slEndDate, "mmm, yy")
                        For ilLoop = 0 To lbcPEDate.ListCount - 1 Step 1
                            If gDateValue(slEndDate) = gDateValue(lbcPEDate.List(ilLoop)) Then
                                lbcPEDate.ListIndex = ilLoop
                                Exit For
                            End If
                        Next ilLoop
                    End If
                Else
                    lbcPEDate.ListIndex = 0 'Start at first end date of period after contract start
                    edcPSDropDown.Text = lbcPEDate.List(0)
                End If
            Else
                edcPSDropDown.Text = lbcPEDate.List(lbcPEDate.ListIndex)
            End If
            imChgMode = False
            lbcPEDate.Move edcPSDropDown.Left, edcPSDropDown.Top + edcPSDropDown.height
            edcPSDropDown.SelStart = 0
            edcPSDropDown.SelLength = Len(edcPSDropDown.Text)
            edcPSDropDown.Visible = True
            cmcPSDropDown.Visible = True
            edcPSDropDown.SetFocus
        Case FSNOMONTHSONINDEX 'No Periods
            edcNoMonthsOn.Width = tmFSCtrls(FSNOMONTHSONINDEX).fBoxW
            gMoveFormCtrl pbcFixSpec, edcNoMonthsOn, tmFSCtrls(FSNOMONTHSONINDEX).fBoxX, tmFSCtrls(FSNOMONTHSONINDEX).fBoxY
            If edcNoMonthsOn.Text = "" Then
                edcNoMonthsOn.Text = "1"
            End If
            edcNoMonthsOn.Visible = True  'Set visibility
            edcNoMonthsOn.SetFocus
        Case FSNOMONTHSOFFINDEX 'No Periods
            edcNoMonthsOff.Width = tmFSCtrls(FSNOMONTHSOFFINDEX).fBoxW
            gMoveFormCtrl pbcFixSpec, edcNoMonthsOff, tmFSCtrls(FSNOMONTHSOFFINDEX).fBoxX, tmFSCtrls(FSNOMONTHSOFFINDEX).fBoxY
            If edcNoMonthsOff.Text = "" Then
                edcNoMonthsOff.Text = "0"
            End If
            edcNoMonthsOff.Visible = True  'Set visibility
            edcNoMonthsOff.SetFocus
        'Case FSAMOUNTINDEX   'Amount
        '    gShowHelpMess tmSbfHelp(), SBFPKGTOT
        '    If edcPSAmount.Text = "" Then
        '        slATotal = "0"
        '        For ilLoop = LBound(smFBSave, 2) To UBound(smFBSave, 2) - 1 Step 1
        '            slATotal = gAddStr(slATotal, smFBSave(2, ilLoop))
        '        Next ilLoop
        '        slATotal = gSubStr(smGross, slATotal)
        '        edcPSAmount.Text = slATotal
        '    End If
        '    edcPSAmount.Width = tmFSCtrls(FSAMOUNTINDEX).fBoxW
        '    gMoveFormCtrl pbcFixSpec, edcPSAmount, tmFSCtrls(FSAMOUNTINDEX).fBoxX, tmFSCtrls(FSAMOUNTINDEX).fBoxY
        '    edcPSAmount.Visible = True  'Set visibility
        '    edcPSAmount.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFBSetShow                      *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mFSSetShow(ilBoxNo As Integer)
'
'   mFBSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmFSCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case FSSTARTDATEINDEX 'Vehicle
            lbcPSDate.Visible = False
            edcPSDropDown.Visible = False
            cmcPSDropDown.Visible = False
            slStr = edcPSDropDown.Text
            'slStr = Left$(slStr, 3) & " 01" & Mid(slStr, 4)
            'slStr = Format(slStr, "m/d/yy")
            If gValidDate(slStr) Then
                slStr = edcPSDropDown.Text
                gSetShow pbcFixSpec, slStr, tmFSCtrls(ilBoxNo)
            Else
                Beep
            End If
        Case FSNOMONTHSINDEX
            edcNoPeriods.Visible = False
            slStr = edcNoPeriods.Text
            gSetShow pbcFixSpec, slStr, tmFSCtrls(ilBoxNo)
        Case FSENDDATEINDEX 'Program index
            lbcPEDate.Visible = False
            edcPSDropDown.Visible = False
            cmcPSDropDown.Visible = False
            slStr = edcPSDropDown.Text
            'slStr = Left$(slStr, 3) & " 01" & Mid(slStr, 4)
            'slStr = Format(slStr, "m/d/yy")
            If gValidDate(slStr) Then
                slStr = edcPSDropDown.Text
                gSetShow pbcFixSpec, slStr, tmFSCtrls(ilBoxNo)
            Else
                Beep
            End If
        Case FSNOMONTHSONINDEX
            edcNoMonthsOn.Visible = False
            slStr = edcNoMonthsOn.Text
            gSetShow pbcFixSpec, slStr, tmFSCtrls(ilBoxNo)
        Case FSNOMONTHSOFFINDEX
            edcNoMonthsOff.Visible = False
            slStr = edcNoMonthsOff.Text
            gSetShow pbcFixSpec, slStr, tmFSCtrls(ilBoxNo)
        Case FSAMOUNTINDEX
            edcPSAmount.Visible = False
            slStr = edcPSAmount.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            gSetShow pbcFixSpec, slStr, tmFSCtrls(ilBoxNo)
    End Select
    mSetGenCommand
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mIBEnableBox                    *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mIBEnableBox(ilBoxNo As Integer)
'
'   mIBEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slSlspComm As String
    Dim slAmount As String
    Dim slStr As String
    Dim ilTax As Integer
    Dim ilVefCode As Integer
    Dim slTax As String
    Dim ilCopyRow As Integer

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmIBCtrls)) Then
        Exit Sub
    End If

    If (imIBRowNo < vbcItemBill.Value) Or (imIBRowNo >= vbcItemBill.Value + vbcItemBill.LargeChange + 1) Then
        mIBSetShow ilBoxNo
        pbcArrow.Visible = False
        lacIBFrame.Visible = False
        Exit Sub
    End If
    lacIBFrame.Move 0, tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15) - 30
    lacIBFrame.Visible = True
    pbcArrow.Move plcItemBill.Left - pbcArrow.Width - 15, plcItemBill.Top + tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    ilCopyRow = False
    If imIBRowNo > 1 Then
        If imIBSave(7, imIBRowNo - 1) <= 0 Then
            ilCopyRow = True
        End If
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case IBVEHICLEINDEX 'Vehicle
            gShowHelpMess tmSbfHelp(), SBFITEMFACIL
            lbcBVehicle.height = gListBoxHeight(lbcBVehicle.ListCount, 8)
            'JW - 8/10/22 - Bonus improvement: fix NTR Tab - vehicle dropdown size
            'edcDropDown.Width = (3 * tmIBCtrls(IBVEHICLEINDEX).fBoxW) \ 2
            edcDropDown.Width = tmIBCtrls(IBVEHICLEINDEX).fBoxW
            lbcBVehicle.Width = tmIBCtrls(IBVEHICLEINDEX).fBoxW + 200
            If tgSpf.iVehLen <= 40 Then
                edcDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveTableCtrl pbcItemBill, edcDropDown, tmIBCtrls(IBVEHICLEINDEX).fBoxX, tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            imChgMode = True
            If imIBSave(1, imIBRowNo) >= 0 Then
                lbcBVehicle.ListIndex = imIBSave(1, imIBRowNo)
                imComboBoxIndex = lbcBVehicle.ListIndex
                edcDropDown.Text = lbcBVehicle.List(imIBSave(1, imIBRowNo))
            Else
                If ilCopyRow Then
                    lbcBVehicle.ListIndex = imIBSave(1, imIBRowNo - 1)
                    imComboBoxIndex = lbcBVehicle.ListIndex
                    edcDropDown.Text = lbcBVehicle.List(imIBSave(1, imIBRowNo - 1))
                Else
                    slNameCode = Contract.lbcRateCard.List(Contract.lbcRateCard.ListIndex)
                    ilRet = gParseItem(slNameCode, 2, "/", slName)
                    If ilRet <> CP_MSG_NONE Then
                        slName = sgUserDefVehicleName
                    Else
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If InStr(slCode, "-") <> 0 Then
                            slName = sgUserDefVehicleName
                        End If
                    End If
                    If (slName = "[All Vehicles]") And (sgUserDefVehicleName <> "") Then
                        slName = sgUserDefVehicleName
                    End If
                    gFindMatch slName, 0, lbcBVehicle
                    If gLastFound(lbcBVehicle) >= 0 Then
                        lbcBVehicle.ListIndex = gLastFound(lbcBVehicle)
                        imComboBoxIndex = lbcBVehicle.ListIndex
                        edcDropDown.Text = lbcBVehicle.List(lbcBVehicle.ListIndex)
                    Else
                        If lbcBVehicle.ListCount > 0 Then
                            lbcBVehicle.ListIndex = 0
                            imComboBoxIndex = lbcBVehicle.ListIndex
                            edcDropDown.Text = lbcBVehicle.List(0)
                        End If
                    End If
                End If
            End If
            imChgMode = False
            If imIBRowNo - vbcItemBill.Value <= vbcItemBill.LargeChange \ 2 Then
                lbcBVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                lbcBVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcBVehicle.height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IBDATEINDEX 'Date
            gShowHelpMess tmSbfHelp(), SBFITEMDATE
            'lbcBDate.Height = gListBoxHeight(lbcBDate.ListCount, 8)
            edcDropDown.Width = tmIBCtrls(IBDATEINDEX).fBoxW
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcItemBill, edcDropDown, tmIBCtrls(IBDATEINDEX).fBoxX, tmIBCtrls(IBDATEINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            'imChgMode = True
            'If lbcBDate.ListIndex < 0 Then
            '    If imIBRowNo <= 1 Then
            '        lbcBDate.ListIndex = 0
            '        edcDropDown.Text = lbcBDate.List(0)
            '    Else
            '        lbcBDate.ListIndex = imIBSave(2, imIBRowNo - 1)
            '        edcDropDown.Text = lbcBDate.List(lbcBDate.ListIndex)
            '    End If
            'Else
            '    edcDropDown.Text = lbcBDate.List(lbcBDate.ListIndex)
            'End If
            'imChgMode = False
            If Trim$(smIBSave(8, imIBRowNo)) = "" Then
                If imIBRowNo <= 1 Then
                    slStr = lbcBDate.List(0)
                Else
                    slStr = smIBSave(8, imIBRowNo - 1)
                End If
            Else
                slStr = smIBSave(8, imIBRowNo)
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcDropDown.Text = slStr

            If imIBRowNo - vbcItemBill.Value <= vbcItemBill.LargeChange \ 2 Then
            ''    lbcBDate.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
            ''    lbcBDate.Move edcDropDown.Left, edcDropDown.Top - lbcBDate.Height
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IBDESCRIPTINDEX
            gShowHelpMess tmSbfHelp(), SBFITEMDESC
            edcDescription.Width = tmIBCtrls(IBDESCRIPTINDEX).fBoxW
            gMoveTableCtrl pbcItemBill, edcDescription, tmIBCtrls(IBDESCRIPTINDEX).fBoxX, tmIBCtrls(IBDESCRIPTINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15)
            edcDescription.Text = smIBSave(2, imIBRowNo)
            If (Trim$(smIBSave(2, imIBRowNo)) = "") And (ilCopyRow) Then
                edcDescription.Text = smIBSave(2, imIBRowNo - 1)
            End If
            edcDescription.Visible = True  'Set visibility
            edcDescription.SetFocus
        Case IBITEMTYPEINDEX 'Item bill type
            lbcBItem.height = gListBoxHeight(lbcBItem.ListCount, 8)
            edcDropDown.Width = tmIBCtrls(IBITEMTYPEINDEX).fBoxW
            edcDropDown.MaxLength = 20
            gMoveTableCtrl pbcItemBill, edcDropDown, tmIBCtrls(IBITEMTYPEINDEX).fBoxX, tmIBCtrls(IBITEMTYPEINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcBItem.ListIndex = imIBSave(3, imIBRowNo)
            imChgMode = True
            If imIBSave(3, imIBRowNo) < 0 Then
                If ilCopyRow Then
                    lbcBItem.ListIndex = imIBSave(3, imIBRowNo - 1)
                    If lbcBItem.ListIndex < 0 Then
                        lbcBItem.ListIndex = 0
                        edcDropDown.Text = lbcBItem.List(0)
                    Else
                        edcDropDown.Text = lbcBItem.List(lbcBItem.ListIndex)
                    End If
                Else
                    If lbcBItem.ListCount > 1 Then
                        lbcBItem.ListIndex = 1
                        edcDropDown.Text = lbcBItem.List(1)
                    Else
                        If imIBSave(3, imIBRowNo) < 0 Then
                            lbcBItem.ListIndex = 0
                            edcDropDown.Text = lbcBItem.List(0)
                        Else
                            lbcBItem.ListIndex = -1
                            edcDropDown.Text = ""
                        End If
                    End If
                End If
            Else
                edcDropDown.Text = lbcBItem.List(lbcBItem.ListIndex)
            End If
            imChgMode = False
            If imIBRowNo - vbcItemBill.Value <= vbcItemBill.LargeChange \ 2 Then
                lbcBItem.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                lbcBItem.Move edcDropDown.Left, edcDropDown.Top - lbcBItem.height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            lbcBItem.Visible = True
            edcDropDown.SetFocus
        Case IBACINDEX
            If imIBSave(4, imIBRowNo) = 2 Then
                Exit Sub
            End If
            gShowHelpMess tmSbfHelp(), SBFITEMAC
            gMoveTableCtrl pbcItemBill, pbcYN, tmIBCtrls(ilBoxNo).fBoxX, tmIBCtrls(ilBoxNo).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            If imIBSave(4, imIBRowNo) = -1 Then
                If igDirAdvt = False Then
                    If ilCopyRow Then
                        imIBSave(4, imIBRowNo) = imIBSave(4, imIBRowNo - 1)
                    Else
                        imIBSave(4, imIBRowNo) = 0  'Changed 3/24/03 to No  it was: 1  'Default to No
                    End If
                Else
                    imIBSave(4, imIBRowNo) = 2  'Default to No
                End If
            End If
            pbcYN.Width = tmIBCtrls(IBACINDEX).fBoxW
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case IBSCINDEX
            'If imIBSave(5, imIBRowNo) = 2 Then
            '    Exit Sub
            'End If
            'gShowHelpMess tmSbfHelp(), SBFITEMSC
            'gMoveTableCtrl pbcItemBill, pbcYN, tmIBCtrls(ilBoxNo).fBoxX, tmIBCtrls(ilBoxNo).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            'If imIBSave(5, imIBRowNo) = -1 Then
            '    imIBSave(5, imIBRowNo) = 1  'Default to No
            'End If
            'pbcYN_Paint
            'pbcYN.Visible = True
            'pbcYN.SetFocus
            edcSalesComm.Width = tmIBCtrls(IBSCINDEX).fBoxW
            gMoveTableCtrl pbcItemBill, edcSalesComm, tmIBCtrls(IBSCINDEX).fBoxX, tmIBCtrls(IBSCINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15)
            If imIBSave(5, imIBRowNo) = -1 Then
                If (ilCopyRow) Then
                    If (imIBSave(3, imIBRowNo) = imIBSave(3, imIBRowNo - 1)) Then
                        edcSalesComm.Text = gIntToStrDec(imIBSave(5, imIBRowNo - 1), 2)
                    Else
                        If imIBSave(3, imIBRowNo) > 0 Then
                            slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                            ilRet = gParseItem(slNameCode, 5, "\", slSlspComm)
                            If ilRet <> CP_MSG_NONE Then
                                edcSalesComm.Text = 0
                            Else
                                edcSalesComm.Text = gIntToStrDec(gStrDecToInt(slSlspComm, 2), 2)
                            End If
                        Else
                            edcSalesComm.Text = 0
                        End If
                    End If
                Else
                    If imIBSave(3, imIBRowNo) > 0 Then
                        slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                        ilRet = gParseItem(slNameCode, 5, "\", slSlspComm)
                        If ilRet <> CP_MSG_NONE Then
                            edcSalesComm.Text = 0
                        Else
                            edcSalesComm.Text = gIntToStrDec(gStrDecToInt(slSlspComm, 2), 2)
                        End If
                    Else
                        edcSalesComm.Text = 0
                    End If
                End If
            Else
                edcSalesComm.Text = gIntToStrDec(imIBSave(5, imIBRowNo), 2)
            End If
            edcSalesComm.Visible = True  'Set visibility
            edcSalesComm.SetFocus
        Case IBTXINDEX

            If Not imTaxDefined Then
                Exit Sub
            End If
            lbcTax.height = gListBoxHeight(lbcTax.ListCount, 8)
            edcDropDown.Width = lbcTax.Width - cmcDropDown.Width
            edcDropDown.MaxLength = 0
            gMoveTableCtrl pbcItemBill, edcDropDown, tmIBCtrls(IBTXINDEX).fBoxX, tmIBCtrls(IBTXINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcTax.ListIndex = imIBSave(6, imIBRowNo)
            imChgMode = True
            If imIBSave(6, imIBRowNo) < 0 Then
                'Set default from vehicle
                slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                ilRet = gParseItem(slNameCode, 6, "\", slTax)
                If ilRet <> CP_MSG_NONE Then
                    slTax = "Y"
                End If


                lbcTax.ListIndex = 0
                edcDropDown.Text = lbcTax.List(lbcTax.ListIndex)
                If slTax = "Y" Then
                    ilVefCode = lbcBVehicle.ItemData(imIBSave(1, imIBRowNo))
                    ilRet = gBinarySearchVef(ilVefCode)
                    If ilRet <> -1 Then
                        For ilTax = 1 To lbcTax.ListCount - 1 Step 1
                            If tgMVef(ilRet).iTrfCode = lbcTax.ItemData(ilTax) Then
                                lbcTax.ListIndex = ilTax
                                edcDropDown.Text = lbcTax.List(lbcTax.ListIndex)
                                Exit For
                            End If
                        Next ilTax
                    End If
                End If
            Else
                edcDropDown.Text = lbcTax.List(lbcTax.ListIndex)
            End If
            imChgMode = False
            If imIBRowNo - vbcItemBill.Value <= vbcItemBill.LargeChange \ 2 Then
                lbcTax.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                lbcTax.Move edcDropDown.Left, edcDropDown.Top - lbcTax.height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            lbcTax.Visible = True
            edcDropDown.SetFocus
        Case IBAMOUNTINDEX
            gShowHelpMess tmSbfHelp(), SBFITEMCOST
            edcAmount.Width = tmIBCtrls(IBAMOUNTINDEX).fBoxW
            gMoveTableCtrl pbcItemBill, edcAmount, tmIBCtrls(IBAMOUNTINDEX).fBoxX, tmIBCtrls(IBAMOUNTINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15)
            If Trim$(smIBSave(3, imIBRowNo)) = "" Then
                If ilCopyRow Then
                    If (imIBSave(3, imIBRowNo) = imIBSave(3, imIBRowNo - 1)) Then
                        edcAmount.Text = smIBSave(3, imIBRowNo - 1)
                    Else
                        If imIBSave(3, imIBRowNo) > 0 Then
                            slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                            ilRet = gParseItem(slNameCode, 3, "\", slAmount)
                            If ilRet <> CP_MSG_NONE Then
                                edcAmount.Text = ""
                            Else
                                edcAmount.Text = slAmount
                            End If
                        Else
                            edcAmount.Text = ""
                        End If
                    End If
                Else
                    If imIBSave(3, imIBRowNo) > 0 Then
                        slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                        ilRet = gParseItem(slNameCode, 3, "\", slAmount)
                        If ilRet <> CP_MSG_NONE Then
                            edcAmount.Text = ""
                        Else
                            edcAmount.Text = slAmount
                        End If
                    Else
                        edcAmount.Text = ""
                    End If
                End If
            Else
                edcAmount.Text = smIBSave(3, imIBRowNo)
            End If
            edcAmount.Visible = True  'Set visibility
            edcAmount.SetFocus
        Case IBUNITSINDEX
            gShowHelpMess tmSbfHelp(), SBFITEMUNITS
'            edcUnits.Width = tmIBCtrls(IBUNITSINDEX).fBoxW
'            gMoveTableCtrl pbcItemBill, edcUnits, tmIBCtrls(IBUNITSINDEX).fBoxX, tmIBCtrls(IBUNITSINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15)
'            edcUnits.Text = smIBSave(4, imIBRowNo)
'            edcUnits.Visible = True  'Set visibility
'            edcUnits.SetFocus
        Case IBNOITEMSINDEX
            gShowHelpMess tmSbfHelp(), SBFITEMNO
            edcNoItems.Width = tmIBCtrls(IBNOITEMSINDEX).fBoxW
            gMoveTableCtrl pbcItemBill, edcNoItems, tmIBCtrls(IBNOITEMSINDEX).fBoxX, tmIBCtrls(IBNOITEMSINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15)
            If Trim$(smIBSave(5, imIBRowNo)) = "" Then
                If ilCopyRow Then
                    edcNoItems.Text = smIBSave(5, imIBRowNo - 1)
                Else
                    edcNoItems.Text = ""
                End If
            Else
                edcNoItems.Text = smIBSave(5, imIBRowNo)
            End If
            edcNoItems.Visible = True  'Set visibility
            edcNoItems.SetFocus
        Case IBTAMOUNTINDEX
            gShowHelpMess tmSbfHelp(), SBFITEMTOT
        Case IBACQCOSTINDEX
            'gShowHelpMess tmSbfHelp(), SBFITEMCOST
            edcAcqAmount.Width = tmIBCtrls(IBACQCOSTINDEX).fBoxW
            gMoveTableCtrl pbcItemBill, edcAcqAmount, tmIBCtrls(IBACQCOSTINDEX).fBoxX, tmIBCtrls(IBACQCOSTINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15)
            If Trim$(smIBSave(10, imIBRowNo)) = "" Then
                If ilCopyRow Then
                    If (imIBSave(3, imIBRowNo) = imIBSave(3, imIBRowNo - 1)) Then
                        edcAcqAmount.Text = smIBSave(10, imIBRowNo - 1)
                    Else
                        If imIBSave(3, imIBRowNo) > 0 Then
                            slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                            ilRet = gParseItem(slNameCode, 7, "\", slAmount)
                            If ilRet <> CP_MSG_NONE Then
                                edcAcqAmount.Text = ""
                            Else
                                edcAcqAmount.Text = slAmount
                            End If
                        Else
                            edcAcqAmount.Text = ""
                        End If
                    End If
                Else
                    If imIBSave(3, imIBRowNo) > 0 Then
                        slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                        ilRet = gParseItem(slNameCode, 7, "\", slAmount)
                        If ilRet <> CP_MSG_NONE Then
                            edcAcqAmount.Text = ""
                        Else
                            edcAcqAmount.Text = slAmount
                        End If
                    Else
                        edcAcqAmount.Text = ""
                    End If
                End If
            Else
                edcAcqAmount.Text = smIBSave(10, imIBRowNo)
            End If
            edcAcqAmount.Visible = True  'Set visibility
            edcAcqAmount.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mIBSetShow                      *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mIBSetShow(ilBoxNo As Integer)
'
'   mIBSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim slNameCode As String
    Dim slAmount As String
    Dim slUnits As String
    Dim ilRet As Integer
    Dim slSlspComm As String
    Dim slTax As String
    lacIBFrame.Visible = False
    pbcArrow.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmIBCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case IBVEHICLEINDEX 'Vehicle
            pbcLbcBVehicle.Visible = False
            lbcBVehicle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBVEHICLEINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If imIBSave(1, imIBRowNo) <> lbcBVehicle.ListIndex Then
                imIBSave(1, imIBRowNo) = lbcBVehicle.ListIndex
                If imIBRowNo < UBound(imIBSave, 2) Then  'New lines set after all fields entered
                    imIBChg = True
                End If
                If smIBSave(1, imIBRowNo) = "" Then
                    smIBSave(1, imIBRowNo) = "I"    '"C"
                End If
            End If
        Case IBDATEINDEX 'Date index
            'lbcBDate.Visible = False
            plcCalendar.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidDate(slStr) Then
                slStr = gFormatDate(slStr)
                gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
                smIBShow(IBDATEINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
                If gDateValue(smIBSave(8, imIBRowNo)) <> gDateValue(slStr) Then
                    smIBSave(8, imIBRowNo) = slStr
                    If imIBRowNo < UBound(imIBSave, 2) Then  'New lines set after all fields entered
                        imIBChg = True
                    End If
                End If
            Else
                Beep
            End If
        Case IBDESCRIPTINDEX
            edcDescription.Visible = False
            slStr = edcDescription.Text
            gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBDESCRIPTINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If smIBSave(2, imIBRowNo) <> edcDescription.Text Then
                If imIBRowNo < UBound(imIBSave, 2) Then  'New lines set after all fields entered
                    imIBChg = True
                End If
                smIBSave(2, imIBRowNo) = edcDescription.Text
            End If
        Case IBITEMTYPEINDEX 'Item type index
            lbcBItem.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcBItem.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcBItem.List(lbcBItem.ListIndex)
            End If
            gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBITEMTYPEINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If imIBSave(3, imIBRowNo) <> lbcBItem.ListIndex Then
                imIBSave(3, imIBRowNo) = lbcBItem.ListIndex
                If imIBRowNo < UBound(imIBSave, 2) Then  'New lines set after all fields entered
                    imIBChg = True
                End If
                If imIBSave(3, imIBRowNo) > 0 Then
                    slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                    ilRet = gParseItem(slNameCode, 3, "\", slAmount)
                    If ilRet <> CP_MSG_NONE Then
                        slAmount = ""
                    End If
                    If Val(smIBSave(3, imIBRowNo)) <> Val(slAmount) Then
                        smIBSave(3, imIBRowNo) = "" 'Amount/unit
                        slStr = ""
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBAMOUNTINDEX)
                        smIBShow(IBAMOUNTINDEX, imIBRowNo) = tmIBCtrls(IBAMOUNTINDEX).sShow
                        smIBSave(6, imIBRowNo) = "" 'Total
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBTAMOUNTINDEX)
                        smIBShow(IBTAMOUNTINDEX, imIBRowNo) = tmIBCtrls(IBTAMOUNTINDEX).sShow
                    End If
                    slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                    ilRet = gParseItem(slNameCode, 4, "\", slUnits)
                    If ilRet <> CP_MSG_NONE Then
                        slUnits = ""
                    End If
                    If slUnits <> smIBSave(4, imIBRowNo) Then
                        smIBSave(4, imIBRowNo) = slUnits ' "" 'Unit definition
                        slStr = slUnits ' ""
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBUNITSINDEX)
                        smIBShow(IBUNITSINDEX, imIBRowNo) = tmIBCtrls(IBUNITSINDEX).sShow
                    End If
                    slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                    ilRet = gParseItem(slNameCode, 5, "\", slSlspComm)
                    If tgSpf.sSubCompany = "Y" Then
                        imIBSave(5, imIBRowNo) = -1
                    Else
                        If ilRet <> CP_MSG_NONE Then
                            'imIBSave(5, imIBRowNo) = 2  'No
                            imIBSave(5, imIBRowNo) = 0
                        Else
                            'If Val(slSlspComm) = 0 Then
                            '    imIBSave(5, imIBRowNo) = 2  'No
                            'Else
                            '    imIBSave(5, imIBRowNo) = 0  'Yes
                            'End If
                            imIBSave(5, imIBRowNo) = gStrDecToInt(slSlspComm, 2)
                        End If
                    End If
                    If igDirAdvt = True Then
                        imIBSave(4, imIBRowNo) = 2
                        slStr = "No"
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBACINDEX)
                        smIBShow(IBACINDEX, imIBRowNo) = tmIBCtrls(IBACINDEX).sShow
                    End If
                    If Not imTaxDefined Then
                        imIBSave(6, imIBRowNo) = 0
                        slStr = "N"
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBTXINDEX)
                        smIBShow(IBTXINDEX, imIBRowNo) = tmIBCtrls(IBTXINDEX).sShow
                    Else
                        slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                        ilRet = gParseItem(slNameCode, 6, "\", slTax)
                        If ilRet <> CP_MSG_NONE Then
                            imIBSave(6, imIBRowNo) = 0
                            slStr = "N"
                            gSetShow pbcItemBill, slStr, tmIBCtrls(IBTXINDEX)
                            smIBShow(IBTXINDEX, imIBRowNo) = tmIBCtrls(IBTXINDEX).sShow
                        Else
                            If slTax = "N" Then
                                imIBSave(6, imIBRowNo) = 0
                                slStr = "N"
                                gSetShow pbcItemBill, slStr, tmIBCtrls(IBTXINDEX)
                                smIBShow(IBTXINDEX, imIBRowNo) = tmIBCtrls(IBTXINDEX).sShow
                            'Else
                            '    imIBSave(6, imIBRowNo) = 0  'Yes
                            End If
                        End If
                    End If
                End If
                mIBTotals False
            End If
        Case IBACINDEX
            pbcYN.Visible = False
            If imIBSave(4, imIBRowNo) = 0 Then
                slStr = "Yes"
            ElseIf imIBSave(4, imIBRowNo) = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(ilBoxNo, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
        Case IBSCINDEX
            'pbcYN.Visible = False
            'If imIBSave(5, imIBRowNo) = 0 Then
            '    slStr = "Yes"
            'ElseIf imIBSave(5, imIBRowNo) = 1 Then
            '    slStr = "No"
            'Else
            '    slStr = ""
            'End If
            'gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
            'smIBShow(ilBoxNo, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If tgSpf.sSubCompany = "Y" Then
                imIBSave(5, imIBRowNo) = -1
                slStr = ""
                gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
            Else
                edcSalesComm.Visible = False
                slStr = edcSalesComm.Text
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
                smIBShow(IBSCINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
                slStr = edcSalesComm.Text
                If imIBSave(5, imIBRowNo) <> gStrDecToInt(slStr, 2) Then
                    If imIBRowNo < UBound(imIBSave, 2) Then  'New lines set after all fields entered
                        imIBChg = True
                    End If
                    imIBSave(5, imIBRowNo) = gStrDecToInt(slStr, 2)
                End If
            End If
        Case IBTXINDEX
            lbcTax.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcTax.ListIndex < 0 Then
                slStr = ""
            ElseIf lbcTax.ListIndex = 0 Then
                slStr = "N"
            Else
                slStr = "Y"
            End If
            gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(ilBoxNo, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If imIBSave(6, imIBRowNo) <> lbcTax.ListIndex Then
                imIBSave(6, imIBRowNo) = lbcTax.ListIndex
                If imIBRowNo < UBound(imIBSave, 2) Then  'New lines set after all fields entered
                    imIBChg = True
                End If
            End If
        Case IBAMOUNTINDEX
            edcAmount.Visible = False
            slStr = edcAmount.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBAMOUNTINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If smIBSave(3, imIBRowNo) <> edcAmount.Text Then
                If imIBRowNo < UBound(imIBSave, 2) Then  'New lines set after all fields entered
                    imIBChg = True
                End If
                smIBSave(3, imIBRowNo) = edcAmount.Text
                If (smIBSave(3, imIBRowNo) = "") Or (smIBSave(5, imIBRowNo) = "") Then
                    smIBSave(6, imIBRowNo) = ""
                    slStr = ""
                Else
                    smIBSave(6, imIBRowNo) = gMulStr(smIBSave(3, imIBRowNo), smIBSave(5, imIBRowNo))
                    slStr = smIBSave(6, imIBRowNo)
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                End If
                gSetShow pbcItemBill, slStr, tmIBCtrls(IBTAMOUNTINDEX)
                smIBShow(IBTAMOUNTINDEX, imIBRowNo) = tmIBCtrls(IBTAMOUNTINDEX).sShow
                mIBTotals False
            End If
        Case IBUNITSINDEX
'            edcUnits.Visible = False  'Set visibility
'            slstr = edcUnits.Text
'            gSetShow pbcItemBill, slstr, tmIBCtrls(ilBoxNo)
'            smIBShow(IBUNITSINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
'            If smIBSave(4, imIBRowNo) <> edcUnits.Text Then
'                If imIBRowNo < UBound(tmIBSbf) + 1 Then   'New lines set after all fields entered
'                    imIBChg = True
'                End If
'                smIBSave(4, imIBRowNo) = edcUnits.Text
'            End If
        Case IBNOITEMSINDEX
            edcNoItems.Visible = False  'Set visibility
            slStr = edcNoItems.Text
            gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBNOITEMSINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If smIBSave(5, imIBRowNo) <> edcNoItems.Text Then
                If imIBRowNo < UBound(imIBSave, 2) Then  'New lines set after all fields entered
                    imIBChg = True
                End If
                smIBSave(5, imIBRowNo) = edcNoItems.Text
                If (smIBSave(3, imIBRowNo) = "") Or (smIBSave(5, imIBRowNo) = "") Then
                    smIBSave(6, imIBRowNo) = ""
                    slStr = ""
                Else
                    smIBSave(6, imIBRowNo) = gMulStr(smIBSave(3, imIBRowNo), smIBSave(5, imIBRowNo))
                    slStr = smIBSave(6, imIBRowNo)
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                End If
                gSetShow pbcItemBill, slStr, tmIBCtrls(IBTAMOUNTINDEX)
                smIBShow(IBTAMOUNTINDEX, imIBRowNo) = tmIBCtrls(IBTAMOUNTINDEX).sShow
                mIBTotals False
            End If
        Case IBACQCOSTINDEX
            edcAcqAmount.Visible = False
            slStr = edcAcqAmount.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBACQCOSTINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If smIBSave(10, imIBRowNo) <> edcAcqAmount.Text Then
                If imIBRowNo < UBound(imIBSave, 2) Then  'New lines set after all fields entered
                    imIBChg = True
                End If
                smIBSave(10, imIBRowNo) = edcAcqAmount.Text
                gSetShow pbcItemBill, slStr, tmIBCtrls(IBACQCOSTINDEX)
                smIBShow(IBACQCOSTINDEX, imIBRowNo) = tmIBCtrls(IBACQCOSTINDEX).sShow
                mIBTotals False
            End If
    End Select
    RaiseEvent NTRDollars(smNTRTGross)
    mSetCommands
    pbcItemBill.Cls
    pbcItemBill_Paint
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mIBTestFields                   *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mIBTestFields() As Integer
'
'   iRet = mIBTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    'TTP 10855 - fix potential NTR Overflow errors
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim llDate As Long
    Dim llStdDate As Long

    For llSbf = LBound(tmIBSbf) To UBound(tmIBSbf) - 1 Step 1
        If ((tmIBSbf(llSbf).iStatus = 0) Or (tmIBSbf(llSbf).iStatus = 1)) And (tmIBSbf(llSbf).SbfRec.sBilled <> "Y") Then
            If tmIBSbf(llSbf).SbfRec.iBillVefCode <= 0 Then
                Screen.MousePointer = vbDefault
                '6469
'                ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
                ilRes = MsgBox(" NTR vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imIBRowNo = llSbf + 1
                imIBBoxNo = IBVEHICLEINDEX
                mIBTestFields = NO
                Exit Function
            End If
            If (tmIBSbf(llSbf).SbfRec.iDate(0) = 0) And (tmIBSbf(llSbf).SbfRec.iDate(1) = 0) Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("Date must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imIBRowNo = llSbf + 1
                imIBBoxNo = IBDATEINDEX
                mIBTestFields = NO
                Exit Function
            End If
            gUnpackDateLong tmIBSbf(llSbf).SbfRec.iDate(0), tmIBSbf(llSbf).SbfRec.iDate(1), llDate
            gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llStdDate
            If llDate <= llStdDate Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("Date must be after last standard broadcast invoice date: " & Format$(llStdDate, "m/d/yy"), vbOKOnly + vbExclamation, "Incomplete")
                imIBRowNo = llSbf + 1
                imIBBoxNo = IBDATEINDEX
                mIBTestFields = NO
                Exit Function
            End If
            gUnpackDateLong tmIBSbf(llSbf).SbfRec.iDate(0), tmIBSbf(llSbf).SbfRec.iDate(1), llDate
            If llDate <= lmLastClosingDate Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("Date must be after last reconcile date of " & smLastClosingDate, vbOKOnly + vbExclamation, "Incomplete")
                imIBRowNo = llSbf + 1
                imIBBoxNo = IBDATEINDEX
                mIBTestFields = NO
                Exit Function
            End If
            If Trim$(tmIBSbf(llSbf).SbfRec.sDescr) = "" Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("Description must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imIBRowNo = llSbf + 1
                imIBBoxNo = IBDESCRIPTINDEX
                mIBTestFields = NO
                Exit Function
            End If
            'gStrToPDN "", 2, 5, slStr
            'If StrComp(tmIBSbf(llSbf).SbfRec.sItemAmount, slStr, 0) = 0 Then
            '    Screen.MousePointer = vbDefault
            '    ilRes = MsgBox("Price must be specified", vbOkOnly + vbExclamation, "Incomplete")
            '    imIBRowNo = llSbf + 1
            '    imIBBoxNo = IBAMOUNTINDEX
            '    mIBTestFields = NO
            '    Exit Function
            'End If
            If tmIBSbf(llSbf).SbfRec.iMnfItem <= 0 Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("NTR Type must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imIBRowNo = llSbf + 1
                imIBBoxNo = IBITEMTYPEINDEX 'IBUNITSINDEX
                mIBTestFields = NO
                Exit Function
            End If
            'If tmIBSbf(llSbf).SbfRec.sUnitName = "" Then
            '    Screen.MousePointer = vbDefault
            '    ilRes = MsgBox("Units must be specified", vbOkOnly + vbExclamation, "Incomplete")
            '    imIBRowNo = llSbf + 1
            '    imIBBoxNo = IBUNITSINDEX
            '    mIBTestFields = NO
            '    Exit Function
            'End If
            If tmIBSbf(llSbf).SbfRec.iNoItems <= 0 Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("Number of Items must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imIBRowNo = llSbf + 1
                imIBBoxNo = IBNOITEMSINDEX
                mIBTestFields = NO
                Exit Function
            End If
        End If
    Next llSbf
    mIBTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mIBTestSaveFields               *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mIBTestSaveFields() As Integer
'
'   iRet = mIBTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    If imIBSave(1, imIBRowNo) < 0 Then
        ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBVEHICLEINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    'If imIBSave(2, imIBRowNo) < 0 Then
    If smIBSave(8, imIBRowNo) = "" Then
        ilRes = MsgBox("Date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBDATEINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    If smIBSave(2, imIBRowNo) = "" Then
        ilRes = MsgBox("Description must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBDESCRIPTINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    If imIBSave(3, imIBRowNo) <= 0 Then
        ilRes = MsgBox("NTR Type must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBITEMTYPEINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    If smIBSave(3, imIBRowNo) = "" Then
        ilRes = MsgBox("Amount/Item must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBAMOUNTINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    'If smIBSave(4, imIBRowNo) = "" Then
    '    ilRes = MsgBox("Units must be specified", vbOkOnly + vbExclamation, "Incomplete")
    '    imIBBoxNo = IBUNITSINDEX
    '    mIBTestSaveFields = NO
    '    Exit Function
    'End If
    If smIBSave(5, imIBRowNo) = "" Then
        ilRes = MsgBox("# Items must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBNOITEMSINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    mIBTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mIBTotals                       *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute totals                  *
'*                                                     *
'*******************************************************
Private Sub mIBTotals(ilForceSettingTotals As Integer)
    Dim slCTotal As String
    Dim slPTotal As String
    Dim slBTotal As String
    Dim slAcqTotal As String
    Dim ilLoop As Integer
    Dim slStr As String

    slCTotal = "0"
    slPTotal = sgIBPTotal
    slBTotal = sgIBBTotal
    slAcqTotal = "0.00"
    For ilLoop = LBONE To UBound(smIBSave, 2) - 1 Step 1
        If smIBSave(1, ilLoop) = "I" Then   '"C" Then
            slCTotal = gAddStr(slCTotal, smIBSave(6, ilLoop))
            slAcqTotal = gAddStr(slAcqTotal, gMulStr(smIBSave(10, ilLoop), smIBSave(5, ilLoop)))
        End If
    Next ilLoop
    smNTRTGross = slCTotal
'    If (slCTotal = "0") And (slPTotal = "0") And (slBTotal = "0") Then
'        lacTotals.Visible = False
'    End If
    lacTotals.BackColor = WHITE
    If pbcItemBill.Visible Or ilForceSettingTotals Then
        lacTotals.Visible = True
        lacTotals.Caption = ""
        slStr = "Totals:"
        If slCTotal <> "0" Then
            gFormatStr slCTotal, FMTLEAVEBLANK + FMTCOMMA, 2, slCTotal
            slStr = slStr & " Contracted $" & slCTotal
        End If
        If slPTotal <> "0" Then
            gFormatStr slPTotal, FMTLEAVEBLANK + FMTCOMMA, 2, slPTotal
            slStr = slStr & "  Posted $" & slPTotal
        End If
        If slBTotal <> "0" Then
            gFormatStr slBTotal, FMTLEAVEBLANK + FMTCOMMA, 2, slBTotal
            slStr = slStr & "  Billed $" & slBTotal
        End If
        '6/7/15: replaced acquisition from site override with Barter in system options
        If (Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) = SPNTRACQUISITION Then
            If slAcqTotal <> "0.00" Then
                gFormatStr slAcqTotal, FMTLEAVEBLANK + FMTCOMMA, 2, slAcqTotal
                slStr = slStr & "  Acquisition $" & slAcqTotal
            End If
        End If
        lacTotals.ForeColor = BLACK
        lacTotals.Caption = slStr
    End If
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilDiff                                                                                *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mInitErr                                                                              *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer    'Return Status
    Dim ilLoop As Integer
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim slDate As String
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim llDate As Long
    Dim ilNoPds As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilLoc As Integer
    Dim slChar As String
    Dim llPrice As Long
    Dim ilSpots As Integer
    Dim llTDollars As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim llWDate As Long
    Dim slBDate As String
    Dim slLineType As String
    Dim ilOptionSet As Integer
    Dim tlSbf As SBF    'Only required to obtain record length
    'ReDim smFBSave(1 To 3, 1 To 1) As String
    'ReDim imFBSave(1 To 2, 1 To 1) As Integer
    'ReDim smFBShow(1 To FBBILLTOTALINDEX, 1 To 1) As String
    ReDim tmInstallBillInfo(0 To 0) As INSTALLBILLINFO
    '12/18/17: Break out NTR separate from Air Time
    bgBreakoutNTR = True
    'ReDim smIBSave(1 To 11, 1 To 1) As String
    'ReDim imIBSave(1 To 7, 1 To 1) As Integer
    'ReDim smIBShow(1 To IBACQCOSTINDEX, 1 To 1) As String
    'ReDim lmIBSave(1 To 2, 1 To 1) As Long
    'ReDim smMBSave(1 To 3, 1 To 1) As String
    'ReDim imMBSave(1 To 1, 1 To 1) As Integer
    'ReDim smMBShow(1 To MPBAMOUNTINDEX, 1 To 1) As String
    'ReDim smPBSave(1 To 3, 1 To 1) As String
    'ReDim imPBSave(1 To 1, 1 To 1) As Integer
    'ReDim smPBShow(1 To MPBAMOUNTINDEX, 1 To 1) As String
    'Index zero ignored in the arrays below
    ReDim smIBSave(0 To 11, 0 To 1) As String
    ReDim imIBSave(0 To 7, 0 To 1) As Integer
    ReDim smIBShow(0 To IBACQCOSTINDEX, 0 To 1) As String
    ReDim lmIBSave(0 To 2, 0 To 1) As Long
    ReDim smMBSave(0 To 3, 0 To 1) As String
    'ReDim imMBSave(0 To 1) As Integer
    ReDim imMBSave(0 To 1, 0 To 1) As Integer
    ReDim smMBShow(0 To MPBAMOUNTINDEX, 0 To 1) As String
    ReDim smPBSave(0 To 3, 0 To 1) As String
    'ReDim imPBSave(0 To 1) As Integer
    ReDim imPBSave(0 To 1, 0 To 1) As Integer
    ReDim smPBShow(0 To MPBAMOUNTINDEX, 0 To 1) As String
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imTerminate = False
    imFirstActivate = True
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.height = 165
    vbcFix.Min = LBound(tmInstallBillInfo) + 1 'LBound(smFBShow, 2)
    vbcFix.Max = UBound(tmInstallBillInfo) + 1  'LBound(smFBShow, 2)
    vbcFix.Value = vbcFix.Min
    vbcItemBill.Min = LBONE 'LBound(smIBShow, 2)
    vbcItemBill.Max = LBONE 'LBound(smIBShow, 2)
    vbcItemBill.Value = vbcItemBill.Min
    vbcMP.Min = LBONE   'LBound(smMBShow, 2)
    vbcMP.Max = LBONE   'LBound(smMBShow, 2)
    vbcMP.Value = vbcMP.Min
    gHlfRead "SBF", tmSbfHelp()
    cmcGen.Enabled = False
    'gPDNToStr tgSpf.sBTax(0), 2, slStr1
    'gPDNToStr tgSpf.sBTax(1), 2, slStr2
    'If (Val(slStr1) = 0) And (Val(slStr2) = 0) Then
    If (Asc(tgSpf.sUsingFeatures3) And TAXONNTR) <> TAXONNTR Then
        imTaxDefined = False
    Else
        imTaxDefined = True
        ilRet = gPopTaxRateBox(True, lbcTax, tmTaxSortCode(), smTaxSortCodeTag)
        If ilRet <> CP_MSG_NOPOPREQ Then
            lbcTax.AddItem "[None]", 0
        End If
    End If
    If Not imTaxDefined Then
        lbcTax.AddItem "[None]", 0
        ReDim tmTaxSortCode(0 To 0) As SORTCODE
    End If
    imFirstActivate = True
    imBypassFocus = False
    imIBBoxNo = -1 'Initialize current Box to N/A
    imFSBoxNo = -1
    imFBBoxNo = -1
    imMSBoxNo = -1
    imMBBoxNo = -1
    imPSBoxNo = -1
    imPBBoxNo = -1
    imIBRowNo = -1
    imFBRowNo = -1
    imMBRowNo = -1
    imPBRowNo = -1
    imFixSort = 0
    imFBChg = False
    imIBChg = False
    imMBChg = False
    imPBChg = False
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imSettingValue = False
    imChgMode = False
    imBSMode = False
    imShowPS = True
    mSetBillCycle
    imCalType = 0   'Standard
    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), smLastClosingDate
    lmLastClosingDate = gDateValue(smLastClosingDate)
    If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
        gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), smLastBilledDate
        smLastBilledDate = gObtainEndCal(smLastBilledDate)
    ElseIf (smBillCycle = "W") Then
        gUnpackDate tgSaf(0).iBLastWeeklyDate(0), tgSaf(0).iBLastWeeklyDate(1), smLastBilledDate
        smLastBilledDate = gObtainNextSunday(smLastBilledDate)
    Else
        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), smLastBilledDate
        smLastBilledDate = gObtainEndStd(smLastBilledDate)
    End If
    lmLastBilledDate = gDateValue(smLastBilledDate)

    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    'gPDNToStr tgChfCntr.sInputGross, 2, smGross
    smGross = gLongToStrDec(tgChfCntr.lInputGross, 2)
    smNet = smGross
    smMPercent = ""
    smLnTGross = "0"
    smLnTNet = ""
    smPPercent = ""
    If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) <> MERCHPROMOBYDOLLAR Then
        If tgChfCntr.iMerchPct <= 0 Then
            ilPos = 1
            Do
                ilPos = InStr(ilPos, smComment(0), "%", 1)
                If ilPos <= 0 Then
                    Exit Do
                End If
                ilLoc = ilPos - 1
                slChar = Mid$(smComment(0), ilLoc, 1)
                If ((slChar >= "0") And (slChar <= "9")) Or (slChar = ".") Then
                    ilLoc = ilPos - 1
                    slChar = Mid$(smComment(0), ilLoc, 1)
                    Do While (slChar = ".") Or ((slChar >= "0") And (slChar <= "9"))
                        ilLoc = ilLoc - 1
                        If ilLoc <= 1 Then
                            ilLoc = 0
                            Exit Do
                        End If
                        slChar = Mid$(smComment(0), ilLoc, 1)
                    Loop
                    smMPercent = Mid$(smComment(0), ilLoc + 1, ilPos - ilLoc - 1)
                    Exit Do
                End If
                ilPos = ilPos + 1
            Loop
            If smMPercent <> "" Then
                If InStr(smMPercent, ".") <= 0 Then
                    smMPercent = smMPercent & "."
                End If
            Else
                ilPos = 1
                Do
                    ilPos = InStr(ilPos, smComment(0), ".", 1)
                    If ilPos <= 0 Then
                        Exit Do
                    End If
                    ilLoc = ilPos + 1
                    slChar = Mid$(smComment(0), ilLoc, 1)
                    If (slChar >= "0") And (slChar <= "9") Then
                        If ilPos > 1 Then
                            ilLoc = ilPos - 1
                            slChar = Mid$(smComment(0), ilLoc, 1)
                            If (slChar >= "0") And (slChar <= "9") Then
                                Exit Do
                            End If
                            ilLoc = ilPos
                        Else
                            ilLoc = ilPos
                        End If
                        ilPos = ilPos + 1
                        slChar = Mid$(smComment(0), ilPos, 1)
                        Do While (slChar = ".") Or ((slChar >= "0") And (slChar <= "9"))
                            ilPos = ilPos + 1
                            If ilPos > Len(smComment(0)) Then
                                ilPos = Len(smComment(0)) + 1
                                Exit Do
                            End If
                            slChar = Mid$(smComment(0), ilPos, 1)
                        Loop
                        smMPercent = Mid$(smComment(0), ilLoc, ilPos - ilLoc)
                        Exit Do
                    End If
                    ilPos = ilPos + 1
                Loop
                If smMPercent <> "" Then
                    If InStr(smMPercent, ".") <= 0 Then
                        smMPercent = smMPercent & "."
                    End If
                    ilPos = InStr(ilPos, smComment(0), "%", 1)
                    If ilPos <= 0 Then
                        smMPercent = gMulStr(smMPercent, "100")
                    End If
                Else
                    If UBound(tgMBSbf) > LBound(tgMBSbf) Then
                        smMPercent = "0"
                    End If
                End If
            End If
        Else
            'smMPercent = gIntToStrDec(tgChfCntr.iMerchPct, 3)
            'Jim: change to two places as abc requested
            smMPercent = gIntToStrDec(tgChfCntr.iMerchPct, 2)
        End If
        If tgChfCntr.iPromoPct <= 0 Then
            ilPos = 1
            Do
                ilPos = InStr(ilPos, smComment(1), "%", 1)
                If ilPos <= 0 Then
                    Exit Do
                End If
                ilLoc = ilPos - 1
                slChar = Mid$(smComment(1), ilLoc, 1)
                If ((slChar >= "0") And (slChar <= "9")) Or (slChar = ".") Then
                    ilLoc = ilPos - 1
                    slChar = Mid$(smComment(1), ilLoc, 1)
                    Do While (slChar = ".") Or ((slChar >= "0") And (slChar <= "9"))
                        ilLoc = ilLoc - 1
                        If ilLoc <= 1 Then
                            ilLoc = 0
                            Exit Do
                        End If
                        slChar = Mid$(smComment(1), ilLoc, 1)
                    Loop
                    smPPercent = Mid$(smComment(1), ilLoc + 1, ilPos - ilLoc - 1)
                    Exit Do
                End If
                ilPos = ilPos + 1
            Loop
            If smPPercent <> "" Then
                If InStr(smPPercent, ".") <= 0 Then
                    smPPercent = smPPercent & "."
                End If
            Else
                ilPos = 1
                Do
                    ilPos = InStr(ilPos, smComment(1), ".", 1)
                    If ilPos <= 0 Then
                        Exit Do
                    End If
                    ilLoc = ilPos + 1
                    slChar = Mid$(smComment(1), ilLoc, 1)
                    If (slChar >= "0") And (slChar <= "9") Then
                        If ilPos > 1 Then
                            ilLoc = ilPos - 1
                            slChar = Mid$(smComment(1), ilLoc, 1)
                            If (slChar >= "0") And (slChar <= "9") Then
                                Exit Do
                            End If
                            ilLoc = ilPos
                        Else
                            ilLoc = ilPos
                        End If
                        ilPos = ilPos + 1
                        slChar = Mid$(smComment(1), ilPos, 1)
                        Do While (slChar = ".") Or ((slChar >= "0") And (slChar <= "9"))
                            ilPos = ilPos + 1
                            If ilPos > Len(smComment(1)) Then
                                ilPos = Len(smComment(1)) + 1
                                Exit Do
                            End If
                            slChar = Mid$(smComment(1), ilPos, 1)
                        Loop
                        smPPercent = Mid$(smComment(1), ilLoc, ilPos - ilLoc)
                        Exit Do
                    End If
                    ilPos = ilPos + 1
                Loop
                If smPPercent <> "" Then
                    If InStr(smPPercent, ".") <= 0 Then
                        smPPercent = smPPercent & "."
                    End If
                    ilPos = InStr(ilPos, smComment(1), "%", 1)
                    If ilPos <= 0 Then
                        smPPercent = gMulStr(smPPercent, "100")
                    End If
                End If
            End If
        Else
            'smPPercent = gIntToStrDec(tgChfCntr.iPromoPct, 3)
            smPPercent = gIntToStrDec(tgChfCntr.iPromoPct, 2)
        End If
    End If
    llStartDate = 0
    llEndDate = 0
    llTDollars = 0
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        If ((tgClfCntr(ilClf).iStatus = 0) Or (tgClfCntr(ilClf).iStatus = 1)) And (Not tgClfCntr(ilClf).iCancel) Then
            slLineType = mGetLineType(ilClf + 1)
            ilCff = tgClfCntr(ilClf).iFirstCff
            Do While ilCff <> -1
                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                    gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate    'Week Start date
                    gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate    'Week Start date
                    If gDateValue(slStartDate) <= gDateValue(slEndDate) Then
                        If llStartDate = 0 Then
                            llStartDate = gDateValue(slStartDate)
                            llEndDate = gDateValue(slEndDate)
                        Else
                            If gDateValue(slStartDate) < llStartDate Then
                                llStartDate = gDateValue(slStartDate)
                            End If
                            If gDateValue(slEndDate) > llEndDate Then
                                llEndDate = gDateValue(slEndDate)
                            End If
                        End If
                    End If
                    'Ignore package lines
                    'If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                    '5/5/07:  Use package lines as that is what is used when comparing merch and prom with contract amount
                    If slLineType <> "H" Then
                        llSDate = gDateValue(slStartDate)
                        llEDate = gDateValue(slEndDate)
                        For llWDate = llSDate To llEDate Step 7
                            ilSpots = mGetFlightSpots(ilClf + 1, llWDate, llPrice)
                            llTDollars = llTDollars + ilSpots * llPrice
                        Next llWDate
                    End If
                End If
                ilCff = tgCffCntr(ilCff).iNextCff
            Loop
        End If
    Next ilClf

    If Not mReadRec(0) Then
        imTerminate = True
        Exit Sub
    End If

    For ilLoop = 0 To UBound(tmIBSbf) - 1 Step 1
        gUnpackDateLong tmIBSbf(ilLoop).SbfRec.iDate(0), tmIBSbf(ilLoop).SbfRec.iDate(1), llDate
        If llStartDate = 0 Then
            llStartDate = llDate
            llEndDate = llDate
        Else
            If llDate < llStartDate Then
                llStartDate = llDate
            End If
            If llDate > llEndDate Then
                llEndDate = llDate
            End If
        End If
    Next ilLoop
    smLnTGross = gLongToStrDec(llTDollars, 2)
    lmCntrStartDate = llStartDate
    lmCntrEndDate = llEndDate
    If llStartDate = 0 Then
        llStartDate = gDateValue(gNow())
        llEndDate = llStartDate
    End If
    If llEndDate < llStartDate Then
        llEndDate = llStartDate
    End If
    If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
        gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), slBDate
        If slBDate <> "" Then
            slBDate = gObtainEndCal(slBDate)
            slBDate = gObtainEndCal(gIncOneDay(slBDate))
        Else
            slBDate = Format$(gNow(), "m/d/yy")
            slBDate = gObtainEndCal(slBDate)
        End If
        If gDateValue(slBDate) < llStartDate - 182 Then     '26*7, 6 months
            slBDate = Format$(llStartDate - 182, "m/d/yy")
            slBDate = gObtainEndCal(slBDate)
        'Else
        '    slBDate = Format$(llStartDate, "m/d/yy")
        '    slBDate = gObtainEndCal(slBDate)
        End If
        gEndCalDatePop slBDate, 36, lbcBDate
    ElseIf (smBillCycle = "W") Then
        gUnpackDate tgSaf(0).iBLastWeeklyDate(0), tgSaf(0).iBLastWeeklyDate(1), slBDate
        If (slBDate <> "") And (gDateValue(slBDate) <> gDateValue("1/1/1990")) Then
            slBDate = gObtainNextSunday(slBDate)
            slBDate = gObtainNextSunday(gIncOneDay(slBDate))
        Else
            slBDate = Format$(gNow(), "m/d/yy")
            slBDate = gObtainNextSunday(slBDate)
        End If
        If gDateValue(slBDate) < llStartDate - 182 Then     '26*7, 6 months
            slBDate = Format$(llStartDate - 182, "m/d/yy")
            slBDate = gObtainNextSunday(slBDate)
        'Else
        '    slBDate = Format$(llStartDate, "m/d/yy")
        '    slBDate = gObtainEndCal(slBDate)
        End If
        gEndWkDatePop slBDate, 132, lbcBDate
   Else
        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slBDate
        If slBDate <> "" Then
            slBDate = gObtainEndStd(slBDate)
            slBDate = gObtainEndStd(gIncOneDay(slBDate))
        Else
            slBDate = Format$(gNow(), "m/d/yy")
            slBDate = gObtainEndStd(slBDate)
        End If
        If gDateValue(slBDate) < llStartDate - 182 Then     '26*7, 6 months
            slBDate = Format$(llStartDate - 182, "m/d/yy")
            slBDate = gObtainEndStd(slBDate)
        'Else
        '    slBDate = Format$(llStartDate, "m/d/yy")
        '    slBDate = gObtainEndStd(slBDate)
        End If
        gEndStdDatePop slBDate, 36, lbcBDate
    End If
    lbcPSDate.Clear
    For ilLoop = 0 To lbcBDate.ListCount - 1 Step 1
        'slDate = Format$(lbcBDate.List(ilLoop), "mmm, yy")
        'lbcPSDate.AddItem slDate
        slDate = lbcBDate.List(ilLoop)
        lbcPSDate.AddItem slDate
        lbcPSDate.ItemData(lbcPSDate.NewIndex) = gDateValue(slDate)
    Next ilLoop
    slDate = Format$(llStartDate, "m/d/yy")
    If (smBillCycle = "C") Then    'Or (smBillCycle = "D") Then
        slStartDate = gObtainEndCal(slDate)
        slDate = Format$(llEndDate, "m/d/yy")
        slEndDate = gObtainEndCal(slDate)
        llDate = gDateValue(slStartDate)
        ilNoPds = 0
        Do
            ilNoPds = ilNoPds + 1
            llDate = llDate + 20
            slDate = Format$(llDate, "m/d/yy")
            slDate = gObtainEndCal(slDate)
            llDate = gDateValue(slDate)
        Loop While (llDate <= gDateValue(slEndDate)) And (ilNoPds < 32000)
        gEndCalDatePop slStartDate, ilNoPds, lbcMSSDate
    ElseIf (smBillCycle = "W") Then
        slStartDate = gObtainNextSunday(slDate)
        slDate = Format$(llEndDate, "m/d/yy")
        slEndDate = gObtainNextSunday(slDate)
        llDate = gDateValue(slStartDate)
        ilNoPds = 0
        Do
            ilNoPds = ilNoPds + 1
            llDate = llDate + 6
            slDate = Format$(llDate, "m/d/yy")
            slDate = gObtainNextSunday(slDate)
            llDate = gDateValue(slDate)
        Loop While (llDate <= gDateValue(slEndDate)) And (ilNoPds < 32000)
        gEndWkDatePop slStartDate, ilNoPds, lbcMSSDate
    Else
        slStartDate = gObtainEndStd(slDate)
        slDate = Format$(llEndDate, "m/d/yy")
        slEndDate = gObtainEndStd(slDate)
        llDate = gDateValue(slStartDate)
        ilNoPds = 0
        Do
            ilNoPds = ilNoPds + 1
            llDate = llDate + 20
            slDate = Format$(llDate, "m/d/yy")
            slDate = gObtainEndStd(slDate)
            llDate = gDateValue(slDate)
        Loop While (llDate <= gDateValue(slEndDate)) And (ilNoPds < 32000)
        gEndStdDatePop slStartDate, ilNoPds, lbcMSSDate
    End If
    lbcPSSDate.Clear
    For ilLoop = 0 To lbcMSSDate.ListCount - 1 Step 1
        slDate = lbcMSSDate.List(ilLoop)
        lbcPSSDate.AddItem slDate
    Next ilLoop
    mVehPop
    If imTerminate Then
        Exit Sub
    End If
    mItemPop
    If imTerminate Then
        Exit Sub
    End If
    If lmOpenPreviouslyCompleted <> 123456789 Then
        hmSbf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen: Sbf.Btr)", CBill
        'On Error GoTo 0
        hmAgf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen: Agf.Btr)", CBill
        'On Error GoTo 0
    End If
    imRecLen = Len(tlSbf) 'btrRecordLength(hmSbf)    'Get Sbf size
    imAgfRecLen = Len(tmAgf) 'btrRecordLength(hmAgf)    'Get Cff size
    smAgyRate = ""
    If tgChfCntr.iAgfCode > 0 Then
        tmAgfSrchKey.iCode = tgChfCntr.iAgfCode
        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrGetEqual: Agf.Btr)", CBill
        'On Error GoTo 0
        If tgChfCntr.iPctTrade <= 0 Then
            smAgyRate = gIntToStrDec(tmAgf.iComm, 2)
        Else
            If tgChfCntr.sAgyCTrade = "Y" Then
                smAgyRate = gIntToStrDec(tmAgf.iComm, 2)
            Else
                smAgyRate = ""
            End If
        End If
        If smAgyRate <> "" Then
            smNet = gDivStr(gMulStr(smGross, gSubStr("100.00", smAgyRate)), "100.00")
            smLnTNet = gDivStr(gMulStr(smLnTGross, gSubStr("100.00", smAgyRate)), "100.00")
        Else
            smNet = smGross
            smLnTNet = smLnTGross
        End If
    Else
        smNet = smGross
        smLnTNet = smLnTGross
    End If

    If lmOpenPreviouslyCompleted <> 123456789 Then
        hmIhf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmIhf, "", sgDBPath & "Ihf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        lmOpenPreviouslyCompleted = 123456789
    End If
    imIhfRecLen = Len(tmIhf) 'btrRecordLength(hmAgf)    'Get Cff size

    'CBill.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'gCenterModalForm CBill
    'Traffic!plcHelp.Cls
    mInitBox
    'gCenterModalForm CBill
    'Move up to tmIBSbf would be set
    'If Not mReadRec(0) Then
    '    imTerminate = True
    '    Exit Sub
    'End If
    lacTotals.Visible = False
    'Only create the BillInfo at this time if Installment exist
    If UBound(tmFBSbf) > LBound(tmFBSbf) Then
        mInitInstallBill
    End If
    'mInitInstallVeh
    'mInitFBShow
    mMoveIBRecToCtrl
    mInitIBShow
    'Only create the BillInfo at this time if Installment exist
    If UBound(tmFBSbf) > LBound(tmFBSbf) Then
        mNTRAddedToInstallment
    End If
    For ilLoop = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
        If (tmInstallBillInfo(ilLoop).lBillDate <= lmLastBilledDate) And (tmInstallBillInfo(ilLoop).lBillDollars = 0) And (tmInstallBillInfo(ilLoop).sType = "O") Then
            tmInstallBillInfo(ilLoop).sBilledFlag = "Y"
        End If
    Next ilLoop
    mFBTotals False
    mFBSort
    If UBound(tmFBSbf) > LBound(tmFBSbf) Then
        mMoveFBRecToCtrl tmFBSbf()
    End If
    cmcClear.Enabled = True
    For ilLoop = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
        If tmInstallBillInfo(ilLoop).sBilledFlag = "Y" Then
            cmcClear.Enabled = False
        End If
    Next ilLoop
    mMoveMBRecToCtrl
    mInitMBShow
    mMovePBRecToCtrl
    mInitPBShow
    'T = Fixed Broadcast month; D=Fixed Calendar month
    'If (smBillCycle = "T") Or (smBillCycle = "D") Or (sgSpecialPassword <> "") Then
    If ((Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) = INSTALLMENT) And ((imAllowCashTradeChgs) Or (UBound(tmFBSbf) > LBound(tmFBSbf))) Then
        imUsingInstallBill = True
        'plcSelect.Enabled = True
        'rbcOption(0).Visible = True
        rbcOption(0).Enabled = True
        If tgSpf.sUsingNTR = "Y" Then
            'rbcOption(1).Value = True
            rbcOption(1).Enabled = True
        Else
            'rbcOption(0).Value = True
            rbcOption(1).Enabled = False
        End If
        'rbcOption_Click 0
    Else
        imUsingInstallBill = False
        'plcSelect.Enabled = False
        rbcOption(0).Enabled = False
        If tgSpf.sUsingNTR = "Y" Then
        '    ilDiff = rbcOption(1).Left - rbcOption(0).Left
        '    rbcOption(1).Left = rbcOption(1).Left - ilDiff
        '    rbcOption(2).Left = rbcOption(2).Left - ilDiff
        '    rbcOption(3).Left = rbcOption(3).Left - ilDiff
        '    plcSelect.Width = plcSelect.Width - ilDiff
        '    'rbcOption(1).Value = True
            rbcOption(1).Enabled = True
        Else
        '    rbcOption(1).Visible = False
        '    ilDiff = rbcOption(2).Left - rbcOption(0).Left
        '    rbcOption(2).Left = rbcOption(2).Left - ilDiff
        '    rbcOption(3).Left = rbcOption(3).Left - ilDiff
        '    plcSelect.Width = plcSelect.Width - ilDiff
            rbcOption(1).Enabled = False
        End If
    End If
    'Contract must be an Order or Hold to define Merchandising or Promotion
    'This way vales can be stored directly into RVF
    If Contract.lbcStatus.ListIndex >= 0 Then
        slStr = Contract.lbcStatus.List(Contract.lbcStatus.ListIndex)
        If (InStr(1, slStr, "Order", 1) > 0) Or (InStr(1, slStr, "Hold", 1) > 0) Then
            If (tgSpf.sRUseMerch = "Y") And (Trim$(smComment(0)) <> "") Then  'Merchandising
                rbcOption(2).Enabled = True
                plcSelect.Enabled = True
            Else
                rbcOption(2).Enabled = False
            End If
            If (tgSpf.sRUsePromo = "Y") And (Trim$(smComment(1)) <> "") Then  'Promotion
                rbcOption(3).Enabled = True
                plcSelect.Enabled = True
            Else
                rbcOption(3).Enabled = False
            End If
        Else
            rbcOption(2).Enabled = False
            rbcOption(3).Enabled = False
        End If
    End If
    ilOptionSet = False
    For ilLoop = 1 To 3 Step 1
        If (rbcOption(ilLoop).Visible) And (rbcOption(ilLoop).Enabled) Then
            If rbcOption(ilLoop).Value = True Then
                rbcOption_Click ilLoop
            Else
                rbcOption(ilLoop).Value = True
            End If
            ilOptionSet = True
            Exit For
        End If
    Next ilLoop
    If Not ilOptionSet Then
        If imUsingInstallBill Then
            rbcOption(0).Value = True
        End If
    End If
    pbcFix_Paint
    vbcItemBill.Value = vbcItemBill.Min
    imSettingValue = False
    If UBound(smIBSave, 2) <= vbcItemBill.LargeChange + 1 Then 'was <=
        vbcItemBill.Max = LBONE 'LBound(smIBSave, 2)
    Else
        vbcItemBill.Max = UBound(smIBSave, 2) - vbcItemBill.LargeChange
    End If
    pbcItemBill_Paint
'    If imUsingInstallBill And (UBound(smFBSave, 2) = 1) Then
'        rbcOption(0).SetFocus
'    Else
'        rbcOption(1).Value = True
'        If imUsingInstallBill Then
'            rbcOption(1).SetFocus
'        Else
'            pbcIBSTab.SetFocus
'        End If
'    End If
    slStr = lbcBDate.List(0)
    'If (smBillCycle = "C") Or (smBillCycle = "D") Then
    '    slStr = gObtainEndCal(slStr)
    'Else
    '    slStr = gObtainEndStd(slStr)
    'End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
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
    Dim llMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long

    flTextHeight = pbcItemBill.TextHeight("1") - 35
    plcSelect.Move 1890, 90
    'Position panel and picture areas with panel
    'plcItemBill.Move 180, plcSelect.Top + plcFixSpec.Height + 60, pbcItemBill.Width + vbcItemBill.Width + fgPanelAdj - 15, pbcItemBill.Height + fgPanelAdj
    plcItemBill.Move 180, 120, pbcItemBill.Width + vbcItemBill.Width + fgPanelAdj - 15, pbcItemBill.height + fgPanelAdj
    pbcItemBill.Move plcItemBill.Left + fgBevelX, plcItemBill.Top + fgBevelY
    vbcItemBill.Move pbcItemBill.Left + pbcItemBill.Width, pbcItemBill.Top + 15
    pbcKey.Move plcItemBill.Left, plcItemBill.Top
    gSetCtrl tmIBCtrls(IBVEHICLEINDEX), 30, 375, 690, fgBoxGridH
    'Date
    gSetCtrl tmIBCtrls(IBDATEINDEX), 735, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 720, fgBoxGridH
    'Description
    gSetCtrl tmIBCtrls(IBDESCRIPTINDEX), 1470, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 1665, fgBoxGridH
    'Item Billing
    gSetCtrl tmIBCtrls(IBITEMTYPEINDEX), 3150, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 990, fgBoxGridH
    'Agency Commission
    gSetCtrl tmIBCtrls(IBACINDEX), 4155, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 180, fgBoxGridH
    'Salesperson Commission
    gSetCtrl tmIBCtrls(IBSCINDEX), 4350, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 435, fgBoxGridH
    If tgSpf.sSubCompany = "Y" Then
        tmIBCtrls(IBSCINDEX).iReq = False
    End If
    'Taxable
    gSetCtrl tmIBCtrls(IBTXINDEX), 4800, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 180, fgBoxGridH
    'Amount/Item
    gSetCtrl tmIBCtrls(IBAMOUNTINDEX), 4995, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 840, fgBoxGridH
    'Units
    gSetCtrl tmIBCtrls(IBUNITSINDEX), 5850, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 600, fgBoxGridH
    '# Items
    gSetCtrl tmIBCtrls(IBNOITEMSINDEX), 6465, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 375, fgBoxGridH
    'Total Amount
    gSetCtrl tmIBCtrls(IBTAMOUNTINDEX), 6855, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 930, fgBoxGridH
    'Amount/Item
    gSetCtrl tmIBCtrls(IBACQCOSTINDEX), 7800, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 840, fgBoxGridH

    'Package specification
    plcFixSpec.Move 1515, plcSelect.Top, cmcGen.Left + cmcGen.Width + 60 + fgPanelAdj, pbcFixSpec.height + fgPanelAdj
    pbcFixSpec.Move plcFixSpec.Left + fgBevelX, plcFixSpec.Top + fgBevelY
    pbcMPSpec.Move plcFixSpec.Left + fgBevelX, plcFixSpec.Top + fgBevelY
    'Start date
    gSetCtrl tmFSCtrls(FSSTARTDATEINDEX), 30, 30, 1185, fgBoxStH
    '# Periods
    gSetCtrl tmFSCtrls(FSNOMONTHSINDEX), 1230, 30, 645, fgBoxStH
    'End Date
    gSetCtrl tmFSCtrls(FSENDDATEINDEX), 1890, 30, 1185, fgBoxStH
    'Amount
    gSetCtrl tmFSCtrls(FSNOMONTHSONINDEX), 3090, 30, 915, fgBoxStH
    'Amount
    gSetCtrl tmFSCtrls(FSNOMONTHSOFFINDEX), 4020, 30, 915, fgBoxStH
    'Amount
    gSetCtrl tmFSCtrls(FSAMOUNTINDEX), 4950, 30, 1170, fgBoxStH
    'Package billing
    'plcFix.Move 2410, 1090, pbcFix.Width + vbcFix.Width + fgPanelAdj - 15, pbcFix.Height + fgPanelAdj
    plcFix.Move plcFixSpec.Left, plcFixSpec.Top + plcFixSpec.height + 60, pbcFix.Width + vbcFix.Width + fgPanelAdj - 15, pbcFix.height + fgPanelAdj
    pbcFix.Move plcFix.Left + fgBevelX, plcFix.Top + fgBevelY
    vbcFix.Move pbcFix.Left + pbcFix.Width, pbcFix.Top + 15
    plcMP.Move plcFix.Left, plcFix.Top, pbcMP.Width + vbcMP.Width + fgPanelAdj - 15, pbcMP.height + fgPanelAdj
    pbcMP.Move plcMP.Left + fgBevelX, plcMP.Top + fgBevelY
    vbcMP.Move pbcMP.Left + pbcMP.Width, pbcMP.Top + 15
    'Vehicle
    gSetCtrl tmFBCtrls(FBVEHICLEINDEX), 30, 375, 1980, fgBoxGridH
    'Date
    gSetCtrl tmFBCtrls(FBDATEINDEX), 2025, tmFBCtrls(FBVEHICLEINDEX).fBoxY, 1215, fgBoxGridH
    'Ordered
    gSetCtrl tmFBCtrls(FBORDEREDINDEX), 3255, tmFBCtrls(FBVEHICLEINDEX).fBoxY, 1170, fgBoxGridH
    'Revenue
    gSetCtrl tmFBCtrls(FBREVENUETOTALINDEX), 4440, tmFBCtrls(FBVEHICLEINDEX).fBoxY, 1170, fgBoxGridH
    'Billing
    gSetCtrl tmFBCtrls(FBBILLINGINDEX), 5625, tmFBCtrls(FBVEHICLEINDEX).fBoxY, 1170, fgBoxGridH
    'Billing/month ot billing/vehicle
    gSetCtrl tmFBCtrls(FBBILLTOTALINDEX), 6810, tmFBCtrls(FBVEHICLEINDEX).fBoxY, 1170, fgBoxGridH

    'Merchandising
    'Percent
    gSetCtrl tmMSCtrls(MPSPERCENTINDEX), 30, 30, 1140, fgBoxStH
    'Start Invoice Date
    gSetCtrl tmMSCtrls(MPSSTARTDATEINDEX), 1185, 30, 1635, fgBoxStH
    'End Invoice Date
    gSetCtrl tmMSCtrls(MPSENDDATEINDEX), 2835, 30, 1635, fgBoxStH
    'Promotion
    'Percent
    gSetCtrl tmPSCtrls(MPSPERCENTINDEX), 30, 30, 1140, fgBoxStH
    'Start Invoice Date
    gSetCtrl tmPSCtrls(MPSSTARTDATEINDEX), 1185, 30, 1635, fgBoxStH
    'End Invoice Date
    gSetCtrl tmPSCtrls(MPSENDDATEINDEX), 2835, 30, 1635, fgBoxStH
    'Vehicle
    gSetCtrl tmMPBCtrls(MPBVEHICLEINDEX), 30, 375, 1980, fgBoxGridH
    'Date
    gSetCtrl tmMPBCtrls(MPBDATEINDEX), 2025, tmMPBCtrls(MPBVEHICLEINDEX).fBoxY, 1215, fgBoxGridH
    'Amount
    gSetCtrl tmMPBCtrls(MPBAMOUNTINDEX), 3255, tmMPBCtrls(MPBVEHICLEINDEX).fBoxY, 1170, fgBoxGridH

    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop

    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmIBCtrls) Step 1
        If ilLoop = IBVEHICLEINDEX Then
            If fmAdjFactorW >= 1.9 Then
                tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + CLng(7 * (Width - lmOrigCBillWidth) / 18)
            ElseIf fmAdjFactorW >= 1.2 Then
                tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + CLng((Width - lmOrigCBillWidth) / 2)
            Else
                tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + CLng((Width - lmOrigCBillWidth))
            End If
            Do While (tmIBCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + 1
            Loop
        Else
            If fmAdjFactorW >= 1.9 Then
                If ilLoop = IBDESCRIPTINDEX Then
                    tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + CLng(5 * (Width - lmOrigCBillWidth) / 18)
                    Do While (tmIBCtrls(ilLoop).fBoxW Mod 15) <> 0
                        tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + 1
                    Loop
                End If
                If ilLoop = IBITEMTYPEINDEX Then
                    tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + CLng((Width - lmOrigCBillWidth) / 6)
                    Do While (tmIBCtrls(ilLoop).fBoxW Mod 15) <> 0
                        tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + 1
                    Loop
                End If
            ElseIf fmAdjFactorW >= 1.2 Then
                If ilLoop = IBDESCRIPTINDEX Then
                    tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + CLng((Width - lmOrigCBillWidth) / 8)
                    Do While (tmIBCtrls(ilLoop).fBoxW Mod 15) <> 0
                        tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + 1
                    Loop
                End If
                If ilLoop = IBITEMTYPEINDEX Then
                    tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + CLng((Width - lmOrigCBillWidth) / 6)
                    Do While (tmIBCtrls(ilLoop).fBoxW Mod 15) <> 0
                        tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + 1
                    Loop
                End If
            End If
            Do
                If tmIBCtrls(ilLoop).fBoxX < tmIBCtrls(ilLoop - 1).fBoxX + tmIBCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmIBCtrls(ilLoop).fBoxX = tmIBCtrls(ilLoop).fBoxX + 15
                ElseIf tmIBCtrls(ilLoop).fBoxX > tmIBCtrls(ilLoop - 1).fBoxX + tmIBCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmIBCtrls(ilLoop).fBoxX = tmIBCtrls(ilLoop).fBoxX - 15
                Else
                    Exit Do
                End If
            Loop
        End If
        If tmIBCtrls(ilLoop).fBoxX + tmIBCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmIBCtrls(ilLoop).fBoxX + tmIBCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    Do While llMax < Width - 3 * plcItemBill.Left - 2 * fgBevelX - vbcItemBill.Width
        For ilLoop = imLBCtrls To UBound(tmIBCtrls) Step 1
            If (ilLoop <> IBVEHICLEINDEX) And (ilLoop <> IBDESCRIPTINDEX) And (ilLoop <> IBDESCRIPTINDEX) Then
                tmIBCtrls(ilLoop).fBoxW = tmIBCtrls(ilLoop).fBoxW + 15
                llMax = llMax + 15
            End If
        Next ilLoop
    Loop
    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmIBCtrls) Step 1
        If ilLoop <> IBVEHICLEINDEX Then
            Do
                If tmIBCtrls(ilLoop).fBoxX < tmIBCtrls(ilLoop - 1).fBoxX + tmIBCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmIBCtrls(ilLoop).fBoxX = tmIBCtrls(ilLoop).fBoxX + 15
                ElseIf tmIBCtrls(ilLoop).fBoxX > tmIBCtrls(ilLoop - 1).fBoxX + tmIBCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmIBCtrls(ilLoop).fBoxX = tmIBCtrls(ilLoop).fBoxX - 15
                Else
                    Exit Do
                End If
            Loop
        End If
        If tmIBCtrls(ilLoop).fBoxX + tmIBCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmIBCtrls(ilLoop).fBoxX + tmIBCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    pbcItemBill.Picture = LoadPicture("")
    pbcItemBill.Width = llMax
    plcItemBill.Width = llMax + vbcItemBill.Width + 2 * fgBevelX + 15
    lacIBFrame.Width = llMax - 15

    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmMPBCtrls) Step 1
        If ilLoop = MPBVEHICLEINDEX Then
            tmMPBCtrls(ilLoop).fBoxW = 2 * tmMPBCtrls(ilLoop).fBoxW 'plcFixSpec.Width - tmMPBCtrls(2).fBoxW - tmMPBCtrls(3).fBoxW 'CLng(fmAdjFactorW * tmMPBCtrls(ilLoop).fBoxW)
            Do While (tmMPBCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmMPBCtrls(ilLoop).fBoxW = tmMPBCtrls(ilLoop).fBoxW + 1
            Loop
        Else
            Do
                If tmMPBCtrls(ilLoop).fBoxX < tmMPBCtrls(ilLoop - 1).fBoxX + tmMPBCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmMPBCtrls(ilLoop).fBoxX = tmMPBCtrls(ilLoop).fBoxX + 15
                ElseIf tmMPBCtrls(ilLoop).fBoxX > tmMPBCtrls(ilLoop - 1).fBoxX + tmMPBCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmMPBCtrls(ilLoop).fBoxX = tmMPBCtrls(ilLoop).fBoxX - 15
                Else
                    Exit Do
                End If
            Loop
        End If
        If tmMPBCtrls(ilLoop).fBoxX + tmMPBCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmMPBCtrls(ilLoop).fBoxX + tmMPBCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    pbcMP.Picture = LoadPicture("")
    pbcMP.Width = llMax
    plcMP.Width = llMax + vbcMP.Width + 2 * fgBevelX + 15

    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmFBCtrls) Step 1
        If ilLoop = FBVEHICLEINDEX Then
            'tmFBCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmFBCtrls(ilLoop).fBoxW)
            tmFBCtrls(ilLoop).fBoxW = tmMPBCtrls(MPBVEHICLEINDEX).fBoxW  'pbcItemBill.Width - tmFBCtrls(FBDATEINDEX).fBoxW - 4 * tmFBCtrls(FBORDEREDINDEX).fBoxW
            Do While (tmFBCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmFBCtrls(ilLoop).fBoxW = tmFBCtrls(ilLoop).fBoxW + 1
            Loop
        Else
            Do
                If tmFBCtrls(ilLoop).fBoxX < tmFBCtrls(ilLoop - 1).fBoxX + tmFBCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmFBCtrls(ilLoop).fBoxX = tmFBCtrls(ilLoop).fBoxX + 15
                ElseIf tmFBCtrls(ilLoop).fBoxX > tmFBCtrls(ilLoop - 1).fBoxX + tmFBCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmFBCtrls(ilLoop).fBoxX = tmFBCtrls(ilLoop).fBoxX - 15
                Else
                    Exit Do
                End If
            Loop
        End If
        If tmFBCtrls(ilLoop).fBoxX + tmFBCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmFBCtrls(ilLoop).fBoxX + tmFBCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop

    pbcFix.Picture = LoadPicture("")
    pbcFix.Width = llMax
    plcFix.Width = llMax + vbcFix.Width + 2 * fgBevelX + 15
    pbcFixSpec.Picture = LoadPicture("")

    If imSetButtons Then
        ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
        Do While ilSpaceBetweenButtons Mod 15 <> 0
            ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
        Loop
        cmcDone.Left = (Width - 3 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
        cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
        cmcUndo.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
        cmcDone.Top = height - (3 * cmcDone.height) / 2 - 60
        cmcCancel.Top = cmcDone.Top
        cmcUndo.Top = cmcDone.Top
    End If
    imSetButtons = False

    imcTrash.Top = height - imcTrash.height - 30
    imcTrash.Left = Width - (3 * imcTrash.Width) / 2
    imcKey.Top = imcTrash.Top
    imcKey.Left = plcItemBill.Left
    pbcKey.Move plcItemBill.Left, imcKey.Top - pbcKey.height - 60

    lacTotals.Top = height - lacTotals.height - 60
    lacTotals.Left = imcTrash.Left - lacTotals.Width - 120
    'llAdjTop = imcTrash.Top - lacTotals.Height - plcItemBill.Top - 2 * fgBevelY - 120
    llAdjTop = imcTrash.Top - plcItemBill.Top - 2 * fgBevelY - 60
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    llAdjTop = llAdjTop
    'Do While plcItemBill.Top + llAdjTop + 2 * fgBevelY + 240 < imcTrash.Top - lacTotals.Height - 60
    Do While plcItemBill.Top + llAdjTop + 2 * fgBevelY + 120 < imcTrash.Top - 60
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    llAdjTop = llAdjTop
    plcItemBill.height = llAdjTop + 2 * fgBevelY
    pbcItemBill.Left = plcItemBill.Left + fgBevelX
    pbcItemBill.Top = plcItemBill.Top + fgBevelY
    pbcItemBill.height = plcItemBill.height - 2 * fgBevelY

    vbcItemBill.Left = pbcItemBill.Left + pbcItemBill.Width + 15
    vbcItemBill.Top = pbcItemBill.Top
    vbcItemBill.height = pbcItemBill.height


    'llAdjTop = imcTrash.Top - lacTotals.Height - plcFix.Top - 2 * fgBevelY - 120
    llAdjTop = lacTotals.Top - plcFix.Top - 2 * fgBevelY - 60
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    llAdjTop = llAdjTop
    'Do While plcFix.Top + llAdjTop + 2 * fgBevelY + 240 < imcTrash.Top - lacTotals.Height - 60
    Do While plcFix.Top + llAdjTop + 2 * fgBevelY + 120 < lacTotals.Top - 60
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    llAdjTop = llAdjTop
    plcFix.height = llAdjTop + 2 * fgBevelY
    plcFix.Left = Width / 2 - plcFix.Width / 2
    plcFixSpec.Left = Width / 2 - plcFixSpec.Width / 2
    pbcFix.Left = plcFix.Left + fgBevelX
    pbcFix.Top = plcFix.Top + fgBevelY
    pbcFix.height = plcFix.height - 2 * fgBevelY
    pbcFixSpec.Left = plcFixSpec.Left + fgBevelX
    pbcFixSpec.Top = plcFixSpec.Top + fgBevelY


    'llAdjTop = imcTrash.Top - plcMP.Top - 2 * fgBevelY - 120
    llAdjTop = height - lacTotals.height - plcMP.Top - 2 * fgBevelY - 60
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    llAdjTop = llAdjTop
    Do While plcMP.Top + llAdjTop + 2 * fgBevelY + 150 < height - 60 - lacTotals.height
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    llAdjTop = llAdjTop
    plcMP.height = llAdjTop + 2 * fgBevelY
    plcMP.Left = Width / 2 - plcMP.Width / 2
    pbcMP.Left = plcMP.Left + fgBevelX
    pbcMP.Top = plcMP.Top + fgBevelY
    pbcMP.height = plcMP.height - 2 * fgBevelY

    pbcMPSpec.Left = pbcFixSpec.Left
    pbcMPSpec.Top = pbcFixSpec.Top

    vbcFix.Left = pbcFix.Left + pbcFix.Width + 15
    vbcFix.Top = pbcFix.Top
    vbcFix.height = pbcFix.height

    vbcMP.Left = pbcMP.Left + pbcMP.Width + 15
    vbcMP.Top = pbcMP.Top
    vbcMP.height = pbcMP.height



    pbcIBTab.Top = height
    pbcFBTab.Top = height
    pbcClickFocus.Top = height

    'pbcComingSoon.Left = Width / 2 - pbcComingSoon.Width / 2
    ''pbcComingSoon.Top = Height / 3 - pbcComingSoon.Height / 2
    'pbcComingSoon.Top = plcFix.Top + pbcComingSoon.Height / 2
    edcInstallMsg.Left = Width / 2 - edcInstallMsg.Width / 2
    edcInstallMsg.Top = plcFix.Top + edcInstallMsg.height / 2

    edcNTRMsg.Top = edcInstallMsg.Top
    edcNTRMsg.Left = edcInstallMsg.Left
    edcMerchMsg.Top = edcInstallMsg.Top
    edcMerchMsg.Left = edcInstallMsg.Left
    edcPromoMsg.Top = edcInstallMsg.Top
    edcPromoMsg.Left = edcInstallMsg.Left
    edcWarningMsg.Left = Width / 2 - edcWarningMsg.Width \ 2
    edcWarningMsg.Top = edcInstallMsg.Top
   
    cmcClear.Left = plcFix.Left
    cmcClear.Top = plcFix.Top + plcFix.height + 30

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitFBShow                     *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show values         *
'*                                                     *
'*******************************************************
Private Sub mInitFBShow()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRowNo                       ilBoxNo                       slStr                     *
'*                                                                                        *
'******************************************************************************************

    Dim ilSvFBRowNo
    ilSvFBRowNo = imFBRowNo
'    For ilRowNo = LBound(smFBSave, 2) To UBound(smFBSave, 2) - 1 Step 1
'        For ilBoxNo = FBVEHICLEINDEX To FBBILLTOTALINDEX Step 1
'            Select Case ilBoxNo 'Branch on box type (control)
'                Case FBVEHICLEINDEX 'Vehicle
'                    slStr = lbcPSVehicle.List(imFBSave(1, ilRowNo))
'                    gSetShow pbcFix, slStr, tmFBCtrls(FBVEHICLEINDEX)
'                    smFBShow(FBVEHICLEINDEX, ilRowNo) = tmFBCtrls(FBVEHICLEINDEX).sShow
'                Case FBDATEINDEX 'Date
'                    slStr = lbcBDate.List(imFBSave(2, ilRowNo))
'                    slStr = gFormatDate(slStr)
'                    gSetShow pbcFix, slStr, tmFBCtrls(FBDATEINDEX)
'                    smFBShow(FBDATEINDEX, ilRowNo) = tmFBCtrls(FBDATEINDEX).sShow
'                'Case FBBILLINDEX
'                '    slStr = smFBSave(3, ilRowNo)
'                '    gSetShow pbcFix, slStr, tmFBCtrls(FBBILLINDEX)
'                '    smFBShow(FBBILLINDEX, ilRowNo) = tmFBCtrls(FBBILLINDEX).sShow
'                Case FBBILLINGINDEX
'                    slStr = smFBSave(2, ilRowNo)
'                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
'                    gSetShow pbcFix, slStr, tmFBCtrls(FBBILLINGINDEX)
'                    smFBShow(FBBILLINGINDEX, ilRowNo) = tmFBCtrls(FBBILLINGINDEX).sShow
'            End Select
'        Next ilBoxNo
'    Next ilRowNo
    imFBRowNo = ilSvFBRowNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitIBShow                     *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show values         *
'*                                                     *
'*******************************************************
Private Sub mInitIBShow()
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    Dim slStr As String
    Dim ilSvIBRowNo As Integer
    Dim ilSvIBBoxNo As Integer
    ilSvIBRowNo = imIBRowNo
    ilSvIBBoxNo = imIBBoxNo
    For ilRowNo = LBONE To UBound(smIBSave, 2) - 1 Step 1
        If imIBSave(3, ilRowNo) <> -1 Then
            For ilBoxNo = IBVEHICLEINDEX To IBACQCOSTINDEX Step 1
                Select Case ilBoxNo 'Branch on box type (control)
                    Case IBVEHICLEINDEX 'Vehicle
                        slStr = lbcBVehicle.List(imIBSave(1, ilRowNo))
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBVEHICLEINDEX)
                        smIBShow(IBVEHICLEINDEX, ilRowNo) = tmIBCtrls(IBVEHICLEINDEX).sShow
                    Case IBDATEINDEX 'Date
                        slStr = smIBSave(8, ilRowNo)    'lbcBDate.List(imIBSave(2, ilRowNo))
                        slStr = gFormatDate(slStr)
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBDATEINDEX)
                        smIBShow(IBDATEINDEX, ilRowNo) = tmIBCtrls(IBDATEINDEX).sShow
                    Case IBDESCRIPTINDEX
                        slStr = smIBSave(2, ilRowNo)
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBDESCRIPTINDEX)
                        smIBShow(IBDESCRIPTINDEX, ilRowNo) = tmIBCtrls(IBDESCRIPTINDEX).sShow
                    Case IBITEMTYPEINDEX 'Date
                        If imIBSave(3, ilRowNo) > 0 Then
                            slStr = lbcBItem.List(imIBSave(3, ilRowNo))
                        Else
                            slStr = ""
                        End If
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBITEMTYPEINDEX)
                        smIBShow(IBITEMTYPEINDEX, ilRowNo) = tmIBCtrls(IBITEMTYPEINDEX).sShow
                    Case IBACINDEX
                        If imIBSave(4, ilRowNo) = 0 Then
                            slStr = "Yes"
                        ElseIf imIBSave(4, ilRowNo) = 1 Then
                            slStr = "No"
                        Else
                            slStr = ""
                        End If
                        gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
                        smIBShow(ilBoxNo, ilRowNo) = tmIBCtrls(ilBoxNo).sShow
                    Case IBSCINDEX
                        'If imIBSave(5, ilRowNo) = 0 Then
                        '    slStr = "Yes"
                        'ElseIf imIBSave(5, ilRowNo) = 1 Then
                        '    slStr = "No"
                        'Else
                        '    slStr = ""
                        'End If
                        If tgSpf.sSubCompany = "Y" Then
                            slStr = ""
                        Else
                            If imIBSave(5, ilRowNo) >= 0 Then
                                slStr = gIntToStrDec(imIBSave(5, ilRowNo), 2)
                            Else
                                slStr = ""
                            End If
                        End If
                        gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
                        smIBShow(ilBoxNo, ilRowNo) = tmIBCtrls(ilBoxNo).sShow
                    Case IBTXINDEX
                        If imIBSave(6, ilRowNo) < 0 Then
                            slStr = " "
                        ElseIf imIBSave(6, ilRowNo) = 0 Then
                            slStr = "N"
                        ElseIf imIBSave(6, ilRowNo) > 0 Then
                            slStr = "Y"
                        Else
                            slStr = ""
                        End If
                        gSetShow pbcItemBill, slStr, tmIBCtrls(ilBoxNo)
                        smIBShow(ilBoxNo, ilRowNo) = tmIBCtrls(ilBoxNo).sShow
                    Case IBAMOUNTINDEX
                        slStr = smIBSave(3, ilRowNo)
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBAMOUNTINDEX)
                        smIBShow(IBAMOUNTINDEX, ilRowNo) = tmIBCtrls(IBAMOUNTINDEX).sShow
                    Case IBUNITSINDEX
                        slStr = smIBSave(4, ilRowNo)
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBUNITSINDEX)
                        smIBShow(IBUNITSINDEX, ilRowNo) = tmIBCtrls(IBUNITSINDEX).sShow
                    Case IBNOITEMSINDEX
                        slStr = smIBSave(5, ilRowNo)
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBNOITEMSINDEX)
                        smIBShow(IBNOITEMSINDEX, ilRowNo) = tmIBCtrls(IBNOITEMSINDEX).sShow
                    Case IBTAMOUNTINDEX
                        If (smIBSave(3, ilRowNo) = "") And (smIBSave(5, ilRowNo) = "") Then
                            smIBSave(6, ilRowNo) = ""
                            slStr = ""
                        Else
                            smIBSave(6, ilRowNo) = gMulStr(smIBSave(3, ilRowNo), smIBSave(5, ilRowNo))
                            slStr = smIBSave(6, ilRowNo)
                            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                        End If
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBTAMOUNTINDEX)
                        smIBShow(IBTAMOUNTINDEX, ilRowNo) = tmIBCtrls(IBTAMOUNTINDEX).sShow
                    Case IBACQCOSTINDEX
                        slStr = smIBSave(10, ilRowNo)
                        '6/7/15: replaced acquisition from site override with Barter in system options
                        If (Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) = SPNTRACQUISITION Then
                            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                        End If
                        gSetShow pbcItemBill, slStr, tmIBCtrls(IBACQCOSTINDEX)
                        smIBShow(IBACQCOSTINDEX, ilRowNo) = tmIBCtrls(IBACQCOSTINDEX).sShow
                End Select
            Next ilBoxNo
        End If
    Next ilRowNo
    imIBBoxNo = ilSvIBBoxNo
    imIBRowNo = ilSvIBRowNo
    mIBTotals False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitMBShow                     *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show values         *
'*                                                     *
'*******************************************************
Private Sub mInitMBShow()
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    Dim slStr As String
    Dim ilSvMBRowNo
    ilSvMBRowNo = imMBRowNo
    For ilRowNo = LBONE To UBound(smMBSave, 2) - 1 Step 1
        For ilBoxNo = MPBVEHICLEINDEX To MPBAMOUNTINDEX Step 1
            Select Case ilBoxNo 'Branch on box type (control)
                Case MPBVEHICLEINDEX 'Vehicle
                    slStr = lbcPSVehicle.List(imMBSave(1, ilRowNo))
                    gSetShow pbcMP, slStr, tmMPBCtrls(MPBVEHICLEINDEX)
                    smMBShow(MPBVEHICLEINDEX, ilRowNo) = tmMPBCtrls(MPBVEHICLEINDEX).sShow
                Case MPBDATEINDEX 'Date
                    slStr = smMBSave(3, ilRowNo)    'lbcBDate.List(imMBSave(2, ilRowNo))
                    slStr = gFormatDate(slStr)
                    gSetShow pbcMP, slStr, tmMPBCtrls(MPBDATEINDEX)
                    smMBShow(MPBDATEINDEX, ilRowNo) = tmMPBCtrls(MPBDATEINDEX).sShow
                Case MPBAMOUNTINDEX
                    slStr = smMBSave(2, ilRowNo)
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    gSetShow pbcMP, slStr, tmMPBCtrls(MPBAMOUNTINDEX)
                    smMBShow(MPBAMOUNTINDEX, ilRowNo) = tmMPBCtrls(MPBAMOUNTINDEX).sShow
            End Select
        Next ilBoxNo
    Next ilRowNo
    imMBRowNo = ilSvMBRowNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNewIB                      *
'*                                                     *
'*             Created:9/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize Item billing        *
'*                                                     *
'*******************************************************
Private Sub mInitNewIB(ilRowNo As Integer)
    Dim ilLoop As Integer
    smIBSave(1, ilRowNo) = ""   'Transaction
    smIBSave(2, ilRowNo) = ""   'Description
    smIBSave(3, ilRowNo) = ""   'Amount/item
    smIBSave(4, ilRowNo) = ""   'Units
    smIBSave(5, ilRowNo) = ""   '# items
    smIBSave(6, ilRowNo) = ""   'Total
    smIBSave(7, ilRowNo) = "N"  'Billed (N or Y)
    smIBSave(8, ilRowNo) = ""   'Date
    smIBSave(9, ilRowNo) = ""   'Invoice Printed date
    smIBSave(10, ilRowNo) = ""   'Invoice Printed date
    smIBSave(11, ilRowNo) = ""   'Game Independent
    imIBSave(1, ilRowNo) = -1   'Vehicle
    imIBSave(2, ilRowNo) = -1   'Date (Not used)
    imIBSave(3, ilRowNo) = -1   'Item billing
    imIBSave(4, ilRowNo) = -1    'Agency commission-default = N
    lmIBSave(1, ilRowNo) = 0    'Tax 1
    lmIBSave(2, ilRowNo) = 0    'Tax 2
    If igDirAdvt = True Then
        imIBSave(4, ilRowNo) = 2
    End If
    imIBSave(5, ilRowNo) = -1   'Salesperson commission- Test item bill
    imIBSave(6, ilRowNo) = -1
    If imTaxDefined Then
        imIBSave(6, ilRowNo) = -1   'Taxable- test item bill if taxes defined
    Else
        imIBSave(6, ilRowNo) = 0    'Set to None
    End If
    imIBSave(7, ilRowNo) = 0    'IhfCode
    For ilLoop = IBVEHICLEINDEX To IBACQCOSTINDEX Step 1
        smIBShow(ilLoop, ilRowNo) = ""
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitPBShow                     *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show values         *
'*                                                     *
'*******************************************************
Private Sub mInitPBShow()
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    Dim slStr As String
    Dim ilSvPBRowNo
    ilSvPBRowNo = imPBRowNo
    For ilRowNo = LBONE To UBound(smPBSave, 2) - 1 Step 1
        For ilBoxNo = MPBVEHICLEINDEX To MPBAMOUNTINDEX Step 1
            Select Case ilBoxNo 'Branch on box type (control)
                Case MPBVEHICLEINDEX 'Vehicle
                    slStr = lbcPSVehicle.List(imPBSave(1, ilRowNo))
                    gSetShow pbcMP, slStr, tmMPBCtrls(MPBVEHICLEINDEX)
                    smPBShow(MPBVEHICLEINDEX, ilRowNo) = tmMPBCtrls(MPBVEHICLEINDEX).sShow
                Case MPBDATEINDEX 'Date
                    slStr = smPBSave(3, ilRowNo)    'lbcBDate.List(imPBSave(2, ilRowNo))
                    slStr = gFormatDate(slStr)
                    gSetShow pbcMP, slStr, tmMPBCtrls(MPBDATEINDEX)
                    smPBShow(MPBDATEINDEX, ilRowNo) = tmMPBCtrls(MPBDATEINDEX).sShow
                Case MPBAMOUNTINDEX
                    slStr = smPBSave(2, ilRowNo)
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    gSetShow pbcMP, slStr, tmMPBCtrls(MPBAMOUNTINDEX)
                    smPBShow(MPBAMOUNTINDEX, ilRowNo) = tmMPBCtrls(MPBAMOUNTINDEX).sShow
            End Select
        Next ilBoxNo
    Next ilRowNo
    imPBRowNo = ilSvPBRowNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mItemPop                       *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Item Billing Types    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mItemPop()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mItemPopErr                                                                           *
'******************************************************************************************

'
'   mAgyDPPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCodeVal As String

    ilIndex = lbcBItem.ListIndex
    If ilIndex > 0 Then
        slName = lbcBItem.List(ilIndex)
    End If
    'Save Code so that index can be reset
    'ReDim slCode(1 To UBound(imIBSave, 2)) As String
    ReDim slCode(0 To UBound(imIBSave, 2)) As String    'Index zero ignored
    For ilLoop = 1 To UBound(imIBSave, 2) - 1 Step 1
        If imIBSave(3, ilLoop) > 0 Then
            slNameCode = tmItemCode(imIBSave(3, ilLoop) - 1).sKey  'lbcItemCode.List(imIBSave(3, ilLoop) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode(ilLoop))
        Else
            slCode(ilLoop) = ""
        End If
    Next ilLoop
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopMnfPlusFieldsBox(CBill, lbcBItem, lbcItemCode, "I")
    'If trade contract, don't include Hard Cost NTR Items
    If tgChfCntr.iPctTrade <> 0 Then
        ilRet = gPopMnfPlusFieldsBoxNoForm(lbcBItem, tmItemCode(), smItemCodeTag, "INH")
    Else
        ilRet = gPopMnfPlusFieldsBoxNoForm(lbcBItem, tmItemCode(), smItemCodeTag, "I")
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        'On Error GoTo mItemPopErr
        'gCPErrorMsg ilRet, "mItemPop (gPopMnfPlusFieldsBox)", CBill
        'On Error GoTo 0
        lbcBItem.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcBItem
            If gLastFound(lbcBItem) > 0 Then
                lbcBItem.ListIndex = gLastFound(lbcBItem)
            Else
                lbcBItem.ListIndex = -1
            End If
        Else
            lbcBItem.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    For ilLoop = 1 To UBound(imIBSave, 2) - 1 Step 1
        imIBSave(3, ilLoop) = -1
        If slCode(ilLoop) <> "" Then
            For ilIndex = 0 To UBound(tmItemCode) - 1 Step 1 'lbcItemCode.ListCount - 1 Step 1
                slNameCode = tmItemCode(ilIndex).sKey  'lbcItemCode.List(ilIndex)
                ilRet = gParseItem(slNameCode, 2, "\", slCodeVal)
                If Val(slCodeVal) = slCode(ilLoop) Then
                    imIBSave(3, ilLoop) = ilIndex + 1
                    Exit For
                End If
            Next ilIndex
        End If
    Next ilLoop
    Exit Sub
mItemPopErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mItemTypeBranch                 *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Item   *
'*                      Billing Type and process       *
'*                      communication back from item   *
'*                      Billing Type                   *
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
Private Function mItemTypeBranch()
'
'   ilRet = mItemTypeBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcBItem, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mItemTypeBranch = False
        Exit Function
    End If
    If igWinStatus(ITEMBILLINGTYPESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mItemTypeBranch = True
        mIBSetFocus imIBBoxNo
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(ITEMBILLINGTYPESLIST)) Then
    '    mItemTypeBranch = True
    '    mIBEnableBox imIBBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourglass  'Wait
    sgMnfCallType = "I"
    igMNmCallSource = CALLSOURCECBILL
    If edcDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Traffic!edcLinkSrceHelpMsg.Text = ""
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
    End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'CBill.Enabled = False
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
    'CBill.Enabled = True
    'Traffic!edcLinkSrceHelpMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mItemTypeBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcBItem.Clear
        smItemCodeTag = ""
        mItemPop
        If imTerminate Then
            mItemTypeBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcBItem
'        mSetChg AGYDPINDEX
        sgMNmName = ""
        If gLastFound(lbcBItem) > 0 Then
            imChgMode = True
            lbcBItem.ListIndex = gLastFound(lbcBItem)
            edcDropDown.Text = lbcBItem.List(lbcBItem.ListIndex)
            imChgMode = False
            mItemTypeBranch = False
        Else
            imChgMode = True
            lbcBItem.ListIndex = 0
            edcDropDown.Text = lbcBItem.List(0)
            imChgMode = False
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mIBEnableBox imIBBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mIBEnableBox imIBBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMBEnableBox                    *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mMBEnableBox(ilBoxNo As Integer)
'
'   mMBEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmMPBCtrls)) Then
        Exit Sub
    End If

    If (imMBRowNo < vbcMP.Value) Or (imMBRowNo >= vbcMP.Value + vbcMP.LargeChange + 1) Then
        mMBSetShow ilBoxNo
        pbcArrow.Visible = False
        Exit Sub
    End If
    pbcArrow.Move plcMP.Left - pbcArrow.Width - 15, plcMP.Top + tmMPBCtrls(MPBVEHICLEINDEX).fBoxY + (imMBRowNo - vbcMP.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case MPBVEHICLEINDEX 'Vehicle
            'gShowHelpMess tmSbfHelp(), SBFBODYFACIL
        Case MPBDATEINDEX 'Program index
            'gShowHelpMess tmSbfHelp(), SBFBODYDATE
'            lbcBDate.Height = gListBoxHeight(lbcBDate.ListCount, 15)
        Case MPBAMOUNTINDEX
            'gShowHelpMess tmSbfHelp(), SBFBODYAMT
            edcAmount.Width = tmMPBCtrls(MPBAMOUNTINDEX).fBoxW
            gMoveTableCtrl pbcMP, edcAmount, tmMPBCtrls(MPBAMOUNTINDEX).fBoxX, tmMPBCtrls(MPBAMOUNTINDEX).fBoxY + (imMBRowNo - vbcMP.Value) * (fgBoxGridH + 15)
            edcAmount.Text = smMBSave(2, imMBRowNo)
            edcAmount.Visible = True  'Set visibility
            edcAmount.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMBSetShow                      *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mMBSetShow(ilBoxNo As Integer)
'
'   mMBSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String

    pbcArrow.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmMPBCtrls)) Then
        Exit Sub
    End If
    If (imMBRowNo < vbcMP.Value) Or (imMBRowNo >= vbcMP.Value + vbcMP.LargeChange + 1) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case MPBVEHICLEINDEX 'Vehicle
        Case MPBDATEINDEX 'Program index
        Case MPBAMOUNTINDEX
            edcAmount.Visible = False
            slStr = edcAmount.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            gSetShow pbcMP, slStr, tmMPBCtrls(ilBoxNo)
            smMBShow(MPBAMOUNTINDEX, imMBRowNo) = tmMPBCtrls(ilBoxNo).sShow
            slStr = edcAmount.Text
            If gCompNumberStr(smMBSave(2, imMBRowNo), slStr) <> 0 Then
                If imMBRowNo < UBound(tmMBSbf) + 1 Then   'New lines set after all fields entered
                    imMBChg = True
                End If
                smMBSave(2, imMBRowNo) = edcAmount.Text
                mMTotals
            End If
    End Select
    mSetCommands
    pbcMP.Cls
    pbcMP_Paint
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMBTestFields                   *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mMBTestFields() As Integer
'
'   iRet = mMBTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    'TTP 10855 - fix potential NTR Overflow errors
    'Dim ilSbf As Integer
    Dim llSbf As Long
    For llSbf = LBound(tmMBSbf) To UBound(tmMBSbf) - 1 Step 1
        If (tmMBSbf(llSbf).iStatus = 0) Or (tmMBSbf(llSbf).iStatus = 1) Then
            If tmMBSbf(llSbf).SbfRec.iBillVefCode <= 0 Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imMBRowNo = llSbf + 1
                imMBBoxNo = MPBVEHICLEINDEX
                mMBTestFields = NO
                Exit Function
            End If
            If (tmMBSbf(llSbf).SbfRec.iDate(0) = 0) And (tmMBSbf(llSbf).SbfRec.iDate(1) = 0) Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("Date must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imMBRowNo = llSbf + 1
                imMBBoxNo = MPBDATEINDEX
                mMBTestFields = NO
                Exit Function
            End If
            'gStrToPDN "", 2, 5, slStr
            'If StrComp(tmMBSbf(llSbf).SbfRec.sItemAmount, slStr, 0) = 0 Then
            '    Screen.MousePointer = vbDefault
            '    ilRes = MsgBox("Price must be specified", vbOkOnly + vbExclamation, "Incomplete")
            '    imMBRowNo = llSbf + 1
            '    imMBBoxNo = MPBAMOUNTINDEX
            '    mMBTestFields = No
            '    Exit Function
            'End If
        End If
    Next llSbf
    mMBTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMBTestSaveFields               *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mMBTestSaveFields() As Integer
'
'   iRet = mMBTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    If imMBSave(1, imMBRowNo) < 0 Then
        ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imMBBoxNo = MPBVEHICLEINDEX
        mMBTestSaveFields = NO
        Exit Function
    End If
    'If imMBSave(2, imMBRowNo) < 0 Then
    If Trim$(smMBSave(3, imMBRowNo)) = "" Then
        ilRes = MsgBox("Date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imMBBoxNo = MPBDATEINDEX
        mMBTestSaveFields = NO
        Exit Function
    End If
    'If smMBSave(2, imMBRowNo) = "" Then
    '    ilRes = MsgBox("Amount must be specified", vbOkOnly + vbExclamation, "Incomplete")
    '    imMBBoxNo = MPBAMOUNTINDEX
    '    mMBTestSaveFields = No
    '    Exit Function
    'End If
    mMBTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveFBCtrlToRec                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move controls values to record *
'*                                                     *
'*******************************************************
Private Sub mMoveFBCtrlToRec()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         slNameCode                    slCode                    *
'*  ilRet                                                                                 *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mMoveFBCtrlToRecErr                                                                   *
'******************************************************************************************

'
'   mMoveFBCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    ilIndex = LBound(tmFBSbf)
    'For ilLoop = LBound(smFBSave, 2) To UBound(smFBSave, 2) - 1 Step 1
    For ilLoop = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
        '12/18/17: Break out NTR separate from Air Time
        'If tmInstallBillInfo(ilLoop).lBillDollars > 0 Then
        If tmInstallBillInfo(ilLoop).lBillDollars <> 0 Then
            tmFBSbf(ilIndex).SbfRec.lChfCode = tgChfCntr.lCode
            tmFBSbf(ilIndex).SbfRec.sTranType = "F" 'smFBSave(1, ilLoop)
            'slNameCode = lbcVehicle.List(imFBSave(1, ilLoop))
            'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            'On Error GoTo mMoveFBCtrlToRecErr
            'gCPErrorMsg ilRet, "mMoveFBCtrlToRec (gParseItem field 2)", CBill
            tmFBSbf(ilIndex).SbfRec.iBillVefCode = tmInstallBillInfo(ilLoop).iVefCode 'CInt(slCode)
            tmFBSbf(ilIndex).SbfRec.iAirVefCode = tmInstallBillInfo(ilLoop).iVefCode 'CInt(slCode)
            'slStr = lbcBDate.List(imFBSave(2, ilLoop))
            'gPackDate slStr, tmFBSbf(ilIndex).SbfRec.iDate(0), tmFBSbf(ilIndex).SbfRec.iDate(1)
            gPackDateLong tmInstallBillInfo(ilLoop).lBillDate, tmFBSbf(ilIndex).SbfRec.iDate(0), tmFBSbf(ilIndex).SbfRec.iDate(1)
            'gStrToPDN smFBSave(2, ilLoop), 2, 5, tmFBSbf(ilIndex).SbfRec.sItemAmount
            tmFBSbf(ilIndex).SbfRec.lGross = tmInstallBillInfo(ilLoop).lBillDollars  'gStrDecToLong(smFBSave(2, ilLoop), 2)
            If tmInstallBillInfo(ilLoop).sBilledFlag = "Y" Then
                tmFBSbf(ilIndex).SbfRec.sBilled = tmInstallBillInfo(ilLoop).sBilledFlag  'smFBSave(3, ilLoop)
            Else
                tmFBSbf(ilIndex).SbfRec.sBilled = "N"
            End If
            tmFBSbf(ilIndex).SbfRec.iTrfCode = 0
            tmFBSbf(ilIndex).SbfRec.iMnfItem = 0
            tmFBSbf(ilIndex).SbfRec.iNoItems = 0
            tmFBSbf(ilIndex).SbfRec.sAgyComm = "Y"
            tmFBSbf(ilIndex).SbfRec.iCommPct = 0
            tmFBSbf(ilIndex).SbfRec.lAcquisitionCost = 0
            tmFBSbf(ilIndex).SbfRec.iIhfCode = 0
            tmFBSbf(ilIndex).SbfRec.iLineNo = 0
            'tmFBSbf(ilIndex).SbfRec.sUnitName = ""
            tmFBSbf(ilIndex).SbfRec.sDescr = ""
            If tmFBSbf(ilIndex).iStatus = -1 Then
                tmFBSbf(ilIndex).iStatus = 0
            ElseIf tmFBSbf(ilIndex).iStatus = 2 Then
                tmFBSbf(ilIndex).iStatus = 1
            End If
            '12/18/17: Break out NTR separate from Air Time
            If bgBreakoutNTR And (tmInstallBillInfo(ilLoop).iMnfItem > 0) Then
                tmFBSbf(ilIndex).SbfRec.iMnfItem = tmInstallBillInfo(ilLoop).iMnfItem
            End If
            ilIndex = ilIndex + 1
            If ilIndex > UBound(tmFBSbf) Then
                ReDim Preserve tmFBSbf(0 To ilIndex)
            End If
        End If
    Next ilLoop
    For ilLoop = ilIndex To UBound(tmFBSbf) - 1 Step 1
        If tmFBSbf(ilLoop).iStatus = 0 Then
            tmFBSbf(ilLoop).iStatus = -1
        ElseIf tmFBSbf(ilLoop).iStatus = 1 Then
            tmFBSbf(ilLoop).iStatus = 2
        End If
        tmFBSbf(ilLoop).SbfRec.sTranType = ""
        tmFBSbf(ilLoop).SbfRec.iBillVefCode = 0
        tmFBSbf(ilLoop).SbfRec.iDate(0) = 0
        tmFBSbf(ilLoop).SbfRec.iDate(1) = 0
        'gStrToPDN "", 2, 5, tmFBSbf(ilIndex).SbfRec.sItemAmount
        tmFBSbf(ilLoop).SbfRec.lGross = 0
        tmFBSbf(ilLoop).SbfRec.iMnfItem = 0
        tmFBSbf(ilLoop).SbfRec.iNoItems = 0
        'tmFBSbf(ilLoop).SbfRec.sUnitName = ""
        tmFBSbf(ilLoop).SbfRec.sDescr = ""
        tmFBSbf(ilLoop).SbfRec.sBilled = "N"
        tmFBSbf(ilLoop).SbfRec.iPrintInvDate(0) = 0
        tmFBSbf(ilLoop).SbfRec.iPrintInvDate(1) = 0
        tmFBSbf(ilLoop).SbfRec.iIhfCode = 0
        tmFBSbf(ilLoop).SbfRec.iLineNo = 0
    Next ilLoop
    Exit Sub
mMoveFBCtrlToRecErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveIBCtrlToRec                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move controls values to record *
'*                                                     *
'*******************************************************
Private Sub mMoveIBCtrlToRec()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mMoveIBCtrlToRecErr                                                                   *
'******************************************************************************************

'
'   mMoveIBCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    ilIndex = LBound(tmIBSbf)
    For ilLoop = LBONE To UBound(smIBSave, 2) - 1 Step 1
        If imIBSave(3, ilLoop) <> -1 Then
            tmIBSbf(ilIndex).SbfRec.lChfCode = tgChfCntr.lCode
            tmIBSbf(ilIndex).SbfRec.sTranType = smIBSave(1, ilLoop)
            'slNameCode = lbcBVehicle.List(imIBSave(1, ilLoop))
            'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            'On Error GoTo mMoveIBCtrlToRecErr
            'gCPErrorMsg ilRet, "mMoveIBCtrlToRec (gParseItem field 2)", CBill
            'tmIBSbf(ilIndex).SbfRec.iBillVefCode = CInt(slCode)
            '6469
            If imIBSave(1, ilLoop) <> -1 Then
                tmIBSbf(ilIndex).SbfRec.iBillVefCode = lbcBVehicle.ItemData(imIBSave(1, ilLoop))
                tmIBSbf(ilIndex).SbfRec.iAirVefCode = lbcBVehicle.ItemData(imIBSave(1, ilLoop))
            Else
                tmIBSbf(ilIndex).SbfRec.iBillVefCode = 0
                tmIBSbf(ilIndex).SbfRec.iAirVefCode = 0
            End If
'            tmIBSbf(ilIndex).SbfRec.iBillVefCode = lbcBVehicle.ItemData(imIBSave(1, ilLoop))
'            tmIBSbf(ilIndex).SbfRec.iAirVefCode = lbcBVehicle.ItemData(imIBSave(1, ilLoop)
            slStr = smIBSave(8, ilLoop) 'lbcBDate.List(imIBSave(2, ilLoop))
            gPackDate slStr, tmIBSbf(ilIndex).SbfRec.iDate(0), tmIBSbf(ilIndex).SbfRec.iDate(1)
            slStr = smIBSave(9, ilLoop)
            gPackDate slStr, tmIBSbf(ilIndex).SbfRec.iPrintInvDate(0), tmIBSbf(ilIndex).SbfRec.iPrintInvDate(1)
            tmIBSbf(ilIndex).SbfRec.sDescr = smIBSave(2, ilLoop)
            If imIBSave(3, ilLoop) > 0 Then
                slNameCode = tmItemCode(imIBSave(3, ilLoop) - 1).sKey  'lbcItemCode.List(imIBSave(3, ilLoop) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                'On Error GoTo mMoveIBCtrlToRecErr
                'gCPErrorMsg ilRet, "mMoveIBCtrlToRec (gParseItem field 2)", CBill
                tmIBSbf(ilIndex).SbfRec.iMnfItem = CInt(slCode)
            Else
                tmIBSbf(ilIndex).SbfRec.iMnfItem = 0
            End If
            'gStrToPDN smIBSave(3, ilLoop), 2, 5, tmIBSbf(ilIndex).SbfRec.sItemAmount
            tmIBSbf(ilIndex).SbfRec.lGross = gStrDecToLong(smIBSave(3, ilLoop), 2)
            'Get Units per from mnfItem (12/10/02)
            'slStr = smIBSave(4, ilLoop)
            'ilPos = InStr(1, slStr, "per ", 1)
            'If ilPos = 1 Then
            '    slStr = right$(slStr, Len(slStr) - 4)
            'End If
            'If Len(slStr) > 6 Then
            '    slStr = Left$(slStr, 6)
            'End If
            'tmIBSbf(ilIndex).SbfRec.sUnitName = slStr
            tmIBSbf(ilIndex).SbfRec.iNoItems = Val(smIBSave(5, ilLoop))
            If (imIBSave(4, ilLoop) = 0) And (Not igDirAdvt) Then
                tmIBSbf(ilIndex).SbfRec.sAgyComm = "Y"
            Else
                tmIBSbf(ilIndex).SbfRec.sAgyComm = "N"
            End If
            '2 was used to indicate that sales comm was set to zero, therefore could not change it to Yes
            'If imIBSave(5, ilLoop) = 0 Then
            '    tmIBSbf(ilIndex).SbfRec.sSlsComm = "Y"
            'Else    'Map -1 or 1 or 2 as No
            '    tmIBSbf(ilIndex).SbfRec.sSlsComm = "N"
            'End If
            If tgSpf.sSubCompany = "Y" Then
                tmIBSbf(ilIndex).SbfRec.iCommPct = 0
            Else
                tmIBSbf(ilIndex).SbfRec.iCommPct = imIBSave(5, ilLoop)
            End If
            '12/17/06-Change to tax by agency or vehicle
            'If imIBSave(6, ilLoop) = 0 Then
            '    tmIBSbf(ilIndex).SbfRec.sSlsTax = "Y"
            'Else    'Map -1 or 1 or 2 as No
            '    tmIBSbf(ilIndex).SbfRec.sSlsTax = "N"
            'End If
            If imIBSave(6, ilLoop) <= 0 Then
                tmIBSbf(ilIndex).SbfRec.iTrfCode = 0
            Else
                tmIBSbf(ilIndex).SbfRec.iTrfCode = lbcTax.ItemData(imIBSave(6, ilLoop))
            End If
            If smIBSave(7, ilLoop) = "Y" Then
                tmIBSbf(ilIndex).SbfRec.sBilled = smIBSave(7, ilLoop)   'Should always be "N" N/A
            Else
                tmIBSbf(ilIndex).SbfRec.sBilled = "N"
            End If
            '12/17/06-Change to tax by agency or vehicle
            'tmIBSbf(ilIndex).SbfRec.lTax1 = lmIBSave(1, ilLoop)
            'tmIBSbf(ilIndex).SbfRec.lTax2 = lmIBSave(2, ilLoop)
            If tmIBSbf(ilIndex).iStatus = -1 Then
                tmIBSbf(ilIndex).iStatus = 0
            ElseIf tmIBSbf(ilIndex).iStatus = 2 Then
                tmIBSbf(ilIndex).iStatus = 1
            End If
            '6/7/15: replaced acquisition from site override with Barter in system options
            If (Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) = SPNTRACQUISITION Then
                tmIBSbf(ilIndex).SbfRec.lAcquisitionCost = gStrDecToLong(smIBSave(10, ilLoop), 2)
            Else
                tmIBSbf(ilIndex).SbfRec.lAcquisitionCost = 0
            End If
            tmIBSbf(ilIndex).SbfRec.iIhfCode = imIBSave(7, ilLoop)
            ilIndex = ilIndex + 1
            If ilIndex > UBound(tmIBSbf) Then
                ReDim Preserve tmIBSbf(0 To ilIndex)
            End If
        End If
    Next ilLoop
    For ilLoop = ilIndex To UBound(tmIBSbf) - 1 Step 1
        If tmIBSbf(ilLoop).iStatus = 0 Then
            tmIBSbf(ilLoop).iStatus = -1
        ElseIf tmIBSbf(ilLoop).iStatus = 1 Then
            tmIBSbf(ilLoop).iStatus = 2
        End If
        tmIBSbf(ilLoop).SbfRec.sTranType = ""
        tmIBSbf(ilLoop).SbfRec.iBillVefCode = 0
        tmIBSbf(ilLoop).SbfRec.iDate(0) = 0
        tmIBSbf(ilLoop).SbfRec.iDate(1) = 0
        'gStrToPDN "", 2, 5, tmIBSbf(ilIndex).SbfRec.sItemAmount
        tmIBSbf(ilLoop).SbfRec.lGross = 0
        tmIBSbf(ilLoop).SbfRec.iMnfItem = 0
        tmIBSbf(ilLoop).SbfRec.iNoItems = 0
        'tmIBSbf(ilLoop).SbfRec.sUnitName = ""
        tmIBSbf(ilLoop).SbfRec.sDescr = ""
        tmIBSbf(ilLoop).SbfRec.sBilled = "N"
        tmIBSbf(ilLoop).SbfRec.iPrintInvDate(0) = 0
        tmIBSbf(ilLoop).SbfRec.iPrintInvDate(1) = 0
        tmIBSbf(ilLoop).SbfRec.iIhfCode = 0
        tmIBSbf(ilLoop).SbfRec.iLineNo = 0
    Next ilLoop
    Exit Sub
mMoveIBCtrlToRecErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveIBRecToCtrl                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveIBRecToCtrl()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mMoveIBRecToCtrlErr                                                                   *
'******************************************************************************************

'
'   mMoveIBRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilIndex As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slUnits As String
    Dim ilTax As Integer

    For ilLoop = 0 To UBound(tmIBSbf) - 1 Step 1
        imIBSave(1, ilLoop + 1) = -1
        'slRecCode = Trim$(Str$(tmIBSbf(ilLoop).SbfRec.iBillVefCode))
        For ilTest = 0 To lbcBVehicle.ListCount - 1 Step 1
            'slNameCode = lbcBVehicle.List(ilTest)
            'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            'On Error GoTo mMoveIBRecToCtrlErr
            'gCPErrorMsg ilRet, "mMoveIBRecToCtrl (gParseItem field 2)", CBill
            'On Error GoTo 0
            'If slRecCode = slCode Then
            If tmIBSbf(ilLoop).SbfRec.iBillVefCode = lbcBVehicle.ItemData(ilTest) Then
                imIBSave(1, ilLoop + 1) = ilTest
                Exit For
            End If
        Next ilTest
        If imIBSave(1, ilLoop + 1) = -1 Then
            gFindMatch sgUserDefVehicleName, 0, lbcBVehicle
            If gLastFound(lbcBVehicle) >= 0 Then
                imIBSave(1, ilLoop + 1) = gLastFound(lbcBVehicle)
            Else
                imIBSave(1, ilLoop + 1) = 0
            End If
        End If
        'If tmIBSbf(ilLoop).SbfRec.sTranType = "C" Then
        '    smIBSave(1, ilLoop + 1) = "C"
        'Else
            smIBSave(1, ilLoop + 1) = "I"
        'End If
        gUnpackDate tmIBSbf(ilLoop).SbfRec.iDate(0), tmIBSbf(ilLoop).SbfRec.iDate(1), slDate
        'gFindMatch slDate, 0, lbcBDate
        'If gLastFound(lbcBDate) >= 0 Then
        '    imIBSave(2, ilLoop + 1) = gLastFound(lbcBDate)
        'Else
        'End If
        smIBSave(8, ilLoop + 1) = slDate
        gUnpackDate tmIBSbf(ilLoop).SbfRec.iPrintInvDate(0), tmIBSbf(ilLoop).SbfRec.iPrintInvDate(1), slDate
        smIBSave(9, ilLoop + 1) = slDate
        smIBSave(2, ilLoop + 1) = Trim$(tmIBSbf(ilLoop).SbfRec.sDescr)
        slRecCode = Trim$(str$(tmIBSbf(ilLoop).SbfRec.iMnfItem))
        For ilIndex = 0 To UBound(tmItemCode) - 1 Step 1 'lbcItemCode.ListCount - 1 Step 1
            slNameCode = tmItemCode(ilIndex).sKey  'lbcItemCode.List(ilIndex)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            'On Error GoTo mMoveIBRecToCtrlErr
            'gCPErrorMsg ilRet, "mMoveIBRecToCtrl (gParseItem field 2)", CBill
            'On Error GoTo 0
            If slRecCode = slCode Then
                imIBSave(3, ilLoop + 1) = ilIndex + 1
                ilRet = gParseItem(slNameCode, 5, "\", slCode)
                'On Error GoTo mMoveIBRecToCtrlErr
                'gCPErrorMsg ilRet, "mMoveIBRecToCtrl (gParseItem field 2)", CBill
                'On Error GoTo 0
                'No value indicated No commission, therefore Yes no allowed
                'If Val(slCode) = 0 Then
                '    imIBSave(5, ilLoop + 1) = 2
                'Else
                '    If tmIBSbf(ilLoop).SbfRec.sSlsComm = "Y" Then
                '        imIBSave(5, ilLoop + 1) = 0
                '    Else
                '        imIBSave(5, ilLoop + 1) = 1
                '    End If
                'End If
                ilRet = gParseItem(slNameCode, 4, "\", slUnits)
                smIBSave(4, ilLoop + 1) = Trim$(slUnits)
                Exit For
            End If
        Next ilIndex
        'Show those rows which lsot the NTR Type
        If imIBSave(3, ilLoop + 1) < 0 Then
            imIBSave(3, ilLoop + 1) = 0
        End If
        If tgSpf.sSubCompany = "Y" Then
            imIBSave(5, ilLoop + 1) = -1
        Else
            imIBSave(5, ilLoop + 1) = tmIBSbf(ilLoop).SbfRec.iCommPct
        End If
        'gPDNToStr tmIBSbf(ilLoop).SbfRec.sItemAmount, 2, smIBSave(3, ilLoop + 1)
        smIBSave(3, ilLoop + 1) = gLongToStrDec(tmIBSbf(ilLoop).SbfRec.lGross, 2)
        'smIBSave(4, ilLoop + 1) = tmIBSbf(ilLoop).SbfRec.sUnitName
        smIBSave(5, ilLoop + 1) = Trim$(str$(tmIBSbf(ilLoop).SbfRec.iNoItems))
        If igDirAdvt = True Then
            imIBSave(4, ilLoop + 1) = 2
        Else
            If tmIBSbf(ilLoop).SbfRec.sAgyComm = "Y" Then
                imIBSave(4, ilLoop + 1) = 0
            Else
                imIBSave(4, ilLoop + 1) = 1
            End If
        End If
        '12/17/06-Change to tax by agency or vehicle
        If Not imTaxDefined Then
            imIBSave(6, ilLoop + 1) = 0     'None
        Else
        '    If tmIBSbf(ilLoop).SbfRec.sSlsTax = "Y" Then
        '        imIBSave(6, ilLoop + 1) = 0
        '    Else
        '        imIBSave(6, ilLoop + 1) = 1
        '    End If
            imIBSave(6, ilLoop + 1) = 0
            If tmIBSbf(ilLoop).SbfRec.iTrfCode > 0 Then
                slRecCode = Trim$(str$(tmIBSbf(ilLoop).SbfRec.iTrfCode))
                For ilTax = 0 To UBound(tmTaxSortCode) - 1 Step 1
                    slNameCode = tmTaxSortCode(ilTax).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    'On Error GoTo mMoveIBRecToCtrlErr
                    'gCPErrorMsg ilRet, "mMoveIBRecToCtrl (gParseItem field 2)", CBill
                    'On Error GoTo 0
                    If slRecCode = slCode Then
                        imIBSave(6, ilLoop + 1) = ilTax + 1
                        Exit For
                    End If
                Next ilTax
            End If
        End If
        smIBSave(6, ilLoop + 1) = gMulStr(smIBSave(3, ilLoop + 1), smIBSave(5, ilLoop + 1))
        smIBSave(7, ilLoop + 1) = tmIBSbf(ilLoop).SbfRec.sBilled
        '12/17/06-Change to tax by agency or vehicle
        'lmIBSave(1, ilLoop + 1) = tmIBSbf(ilLoop).SbfRec.lTax1
        'lmIBSave(2, ilLoop + 1) = tmIBSbf(ilLoop).SbfRec.lTax2
        smIBSave(10, ilLoop + 1) = ""
        '6/7/15: replaced acquisition from site override with Barter in system options
        If (Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) = SPNTRACQUISITION Then
            smIBSave(10, ilLoop + 1) = gLongToStrDec(tmIBSbf(ilLoop).SbfRec.lAcquisitionCost, 2)
        End If
        imIBSave(7, ilLoop + 1) = tmIBSbf(ilLoop).SbfRec.iIhfCode
        If tmIBSbf(ilLoop).SbfRec.iIhfCode > 0 Then
            tmIhfSrchKey0.iCode = tmIBSbf(ilLoop).SbfRec.iIhfCode
            ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            smIBSave(11, ilLoop + 1) = tmIhf.sGameIndependent
        Else
            smIBSave(11, ilLoop + 1) = ""
        End If
    Next ilLoop
    mInitNewIB UBound(smIBSave, 2)
    Exit Sub
mMoveIBRecToCtrlErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveMBCtrlToRec                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move controls values to record *
'*                                                     *
'*******************************************************
Private Sub mMoveMBCtrlToRec()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mMoveMBCtrlToRecErr                                                                   *
'******************************************************************************************

'
'   mMoveMBCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    ilIndex = LBound(tmMBSbf)
    For ilLoop = LBONE To UBound(smMBSave, 2) - 1 Step 1
        tmMBSbf(ilIndex).SbfRec.lChfCode = tgChfCntr.lCode
        tmMBSbf(ilIndex).SbfRec.sTranType = smMBSave(1, ilLoop)
        slNameCode = lbcVehicle.List(imMBSave(1, ilLoop))
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        'On Error GoTo mMoveMBCtrlToRecErr
        'gCPErrorMsg ilRet, "mMoveMBCtrlToRec (gParseItem field 2)", CBill
        tmMBSbf(ilIndex).SbfRec.iBillVefCode = CInt(slCode)
        slStr = smMBSave(3, ilLoop) 'lbcBDate.List(imMBSave(2, ilLoop))
        gPackDate slStr, tmMBSbf(ilIndex).SbfRec.iDate(0), tmMBSbf(ilIndex).SbfRec.iDate(1)
        'gStrToPDN smMBSave(2, ilLoop), 2, 5, tmMBSbf(ilIndex).SbfRec.sItemAmount
        tmMBSbf(ilIndex).SbfRec.lGross = gStrDecToLong(smMBSave(2, ilLoop), 2)
        tmMBSbf(ilIndex).SbfRec.sBilled = "N"
        tmMBSbf(ilIndex).SbfRec.iMnfItem = 0
        tmMBSbf(ilIndex).SbfRec.iNoItems = 0
        'tmMBSbf(ilIndex).SbfRec.sUnitName = ""
        tmMBSbf(ilIndex).SbfRec.sDescr = ""
        If tmMBSbf(ilIndex).iStatus = -1 Then
            tmMBSbf(ilIndex).iStatus = 0
        ElseIf tmMBSbf(ilIndex).iStatus = 2 Then
            tmMBSbf(ilIndex).iStatus = 1
        End If
        tmMBSbf(ilIndex).SbfRec.iPrintInvDate(0) = 0
        tmMBSbf(ilIndex).SbfRec.iPrintInvDate(1) = 0
        tmMBSbf(ilIndex).SbfRec.iIhfCode = 0
        tmMBSbf(ilIndex).SbfRec.iLineNo = 0
        ilIndex = ilIndex + 1
        If ilIndex > UBound(tmMBSbf) Then
            ReDim Preserve tmMBSbf(0 To ilIndex) As SBFLIST
        End If
    Next ilLoop
    For ilLoop = ilIndex To UBound(tmMBSbf) - 1 Step 1
        If tmMBSbf(ilLoop).iStatus = 0 Then
            tmMBSbf(ilLoop).iStatus = -1
        ElseIf tmMBSbf(ilLoop).iStatus = 1 Then
            tmMBSbf(ilLoop).iStatus = 2
        End If
        tmMBSbf(ilLoop).SbfRec.sTranType = ""
        tmMBSbf(ilLoop).SbfRec.iBillVefCode = 0
        tmMBSbf(ilLoop).SbfRec.iDate(0) = 0
        tmMBSbf(ilLoop).SbfRec.iDate(1) = 0
        'gStrToPDN "", 2, 5, tmMBSbf(ilLoop).SbfRec.sItemAmount
        tmMBSbf(ilLoop).SbfRec.lGross = 0
        tmMBSbf(ilLoop).SbfRec.iMnfItem = 0
        tmMBSbf(ilLoop).SbfRec.iNoItems = 0
        'tmMBSbf(ilLoop).SbfRec.sUnitName = ""
        tmMBSbf(ilLoop).SbfRec.sDescr = ""
        tmMBSbf(ilLoop).SbfRec.iPrintInvDate(0) = 0
        tmMBSbf(ilLoop).SbfRec.iPrintInvDate(1) = 0
        tmMBSbf(ilLoop).SbfRec.sBilled = "N"
        tmMBSbf(ilLoop).SbfRec.iIhfCode = 0
        tmMBSbf(ilLoop).SbfRec.iLineNo = 0
    Next ilLoop
    Exit Sub
mMoveMBCtrlToRecErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveMBRecToCtrl                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveMBRecToCtrl()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mMoveMBRecToCtrlErr                                                                   *
'******************************************************************************************

'
'   mMoveMBRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim ilUpper As Integer
    'ReDim smMBSave(1 To 3, 1 To 1) As String
    'ReDim imMBSave(1 To 1, 1 To 1) As Integer
    'ReDim smMBShow(1 To MPBAMOUNTINDEX, 1 To 1) As String
    ReDim smMBSave(0 To 3, 0 To 1) As String
    'ReDim imMBSave(0 To 1) As Integer
    ReDim imMBSave(0 To 1, 0 To 1) As Integer
    ReDim smMBShow(0 To MPBAMOUNTINDEX, 0 To 1) As String
    ilUpper = 1
    For ilLoop = 0 To UBound(tmMBSbf) - 1 Step 1
        If (tmMBSbf(ilLoop).iStatus = 0) Or (tmMBSbf(ilLoop).iStatus = 1) Then
            imMBSave(1, ilUpper) = -1
            slRecCode = Trim$(str$(tmMBSbf(ilLoop).SbfRec.iBillVefCode))
            For ilTest = 0 To lbcVehicle.ListCount - 1 Step 1
                slNameCode = lbcVehicle.List(ilTest)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                'On Error GoTo mMoveMBRecToCtrlErr
                'gCPErrorMsg ilRet, "mMoveMBRecToCtrl (gParseItem field 2)", CBill
                On Error GoTo 0
                If slRecCode = slCode Then
                    imMBSave(1, ilUpper) = ilTest
                    Exit For
                End If
            Next ilTest
            'If imMBSave(1, ilLoop + 1) = -1 Then
            '    gFindMatch sgUserDefVehicleName, 0, lbcPSVehicle
            '    If gLastFound(lbcPSVehicle) >= 0 Then
            '        imMBSave(1, ilLoop + 1) = gLastFound(lbcPSVehicle)
            '    Else
            '        imMBSave(1, ilLoop + 1) = 0
            '    End If
            'End If
            If imMBSave(1, ilUpper) <> -1 Then
                smMBSave(1, ilUpper) = "M"
                gUnpackDate tmMBSbf(ilLoop).SbfRec.iDate(0), tmMBSbf(ilLoop).SbfRec.iDate(1), slDate
                'gFindMatch slDate, 0, lbcBDate
                'If gLastFound(lbcBDate) >= 0 Then
                '    imMBSave(2, ilUpper) = gLastFound(lbcBDate)
                'Else
                'End If
                smMBSave(3, ilUpper) = slDate
                'gPDNToStr tmMBSbf(ilLoop).SbfRec.sItemAmount, 2, smMBSave(2, ilUpper)
                smMBSave(2, ilUpper) = gLongToStrDec(tmMBSbf(ilLoop).SbfRec.lGross, 2)
                ilUpper = ilUpper + 1
                'ReDim Preserve smMBSave(1 To 3, 1 To ilUpper) As String
                'ReDim Preserve imMBSave(1 To 1, 1 To ilUpper) As Integer
                'ReDim Preserve smMBShow(1 To MPBAMOUNTINDEX, 1 To ilUpper) As String
                ReDim Preserve smMBSave(0 To 3, 0 To ilUpper) As String
                'ReDim Preserve imMBSave(0 To ilUpper) As Integer
                ReDim Preserve imMBSave(0 To 1, 0 To ilUpper) As Integer
                ReDim Preserve smMBShow(0 To MPBAMOUNTINDEX, 0 To ilUpper) As String
            Else
                If tmMBSbf(ilLoop).iStatus = 0 Then
                    tmMBSbf(ilLoop).iStatus = -1
                ElseIf tmMBSbf(ilLoop).iStatus = 1 Then
                    tmMBSbf(ilLoop).iStatus = 2
                End If
                tmMBSbf(ilLoop).SbfRec.sTranType = ""
                tmMBSbf(ilLoop).SbfRec.iBillVefCode = 0
                tmMBSbf(ilLoop).SbfRec.iDate(0) = 0
                tmMBSbf(ilLoop).SbfRec.iDate(1) = 0
                'gStrToPDN "", 2, 5, tmMBSbf(ilLoop).SbfRec.sItemAmount
                tmMBSbf(ilLoop).SbfRec.lGross = 0
                tmMBSbf(ilLoop).SbfRec.iMnfItem = 0
                tmMBSbf(ilLoop).SbfRec.iNoItems = 0
                'tmMBSbf(ilLoop).SbfRec.sUnitName = ""
                tmMBSbf(ilLoop).SbfRec.sDescr = ""
                tmMBSbf(ilLoop).SbfRec.iPrintInvDate(0) = 0
                tmMBSbf(ilLoop).SbfRec.iPrintInvDate(1) = 0
                tmMBSbf(ilLoop).SbfRec.sBilled = "N"
                tmMBSbf(ilLoop).SbfRec.iIhfCode = 0
                tmMBSbf(ilLoop).SbfRec.iLineNo = 0
            End If
        End If
    Next ilLoop
    'mInitNewMB UBound(smMBSave, 2)
    Exit Sub
mMoveMBRecToCtrlErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMovePBCtrlToRec                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move controls values to record *
'*                                                     *
'*******************************************************
Private Sub mMovePBCtrlToRec()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mMovePBCtrlToRecErr                                                                   *
'******************************************************************************************

'
'   mMovePBCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    ilIndex = LBound(tmPBSbf)
    For ilLoop = LBONE To UBound(smPBSave, 2) - 1 Step 1
        tmPBSbf(ilIndex).SbfRec.lChfCode = tgChfCntr.lCode
        tmPBSbf(ilIndex).SbfRec.sTranType = smPBSave(1, ilLoop)
        slNameCode = lbcVehicle.List(imPBSave(1, ilLoop))
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        'On Error GoTo mMovePBCtrlToRecErr
        'gCPErrorMsg ilRet, "mMovePBCtrlToRec (gParseItem field 2)", CBill
        tmPBSbf(ilIndex).SbfRec.iBillVefCode = CInt(slCode)
        slStr = smPBSave(3, ilLoop) 'lbcBDate.List(imPBSave(2, ilLoop))
        gPackDate slStr, tmPBSbf(ilIndex).SbfRec.iDate(0), tmPBSbf(ilIndex).SbfRec.iDate(1)
        'gStrToPDN smPBSave(2, ilLoop), 2, 5, tmPBSbf(ilIndex).SbfRec.sItemAmount
        tmPBSbf(ilIndex).SbfRec.lGross = gStrDecToLong(smPBSave(2, ilLoop), 2)
        tmPBSbf(ilIndex).SbfRec.sBilled = "N"
        tmPBSbf(ilIndex).SbfRec.iMnfItem = 0
        tmPBSbf(ilIndex).SbfRec.iNoItems = 0
        'tmPBSbf(ilIndex).SbfRec.sUnitName = ""
        tmPBSbf(ilIndex).SbfRec.sDescr = ""
        tmPBSbf(ilIndex).SbfRec.iPrintInvDate(0) = 0
        tmPBSbf(ilIndex).SbfRec.iPrintInvDate(1) = 0
        tmPBSbf(ilIndex).SbfRec.iIhfCode = 0
        tmPBSbf(ilIndex).SbfRec.iLineNo = 0
        If tmPBSbf(ilIndex).iStatus = -1 Then
            tmPBSbf(ilIndex).iStatus = 0
        ElseIf tmPBSbf(ilIndex).iStatus = 2 Then
            tmPBSbf(ilIndex).iStatus = 1
        End If
        ilIndex = ilIndex + 1
        If ilIndex > UBound(tmPBSbf) Then
            ReDim Preserve tmPBSbf(0 To ilIndex) As SBFLIST
        End If
    Next ilLoop
    For ilLoop = ilIndex To UBound(tmPBSbf) - 1 Step 1
        If tmPBSbf(ilLoop).iStatus = 0 Then
            tmPBSbf(ilLoop).iStatus = -1
        ElseIf tmPBSbf(ilLoop).iStatus = 1 Then
            tmPBSbf(ilLoop).iStatus = 2
        End If
        tmPBSbf(ilLoop).SbfRec.sTranType = ""
        tmPBSbf(ilLoop).SbfRec.iBillVefCode = 0
        tmPBSbf(ilLoop).SbfRec.iDate(0) = 0
        tmPBSbf(ilLoop).SbfRec.iDate(1) = 0
        'gStrToPDN "", 2, 5, tmPBSbf(ilLoop).SbfRec.sItemAmount
        tmPBSbf(ilLoop).SbfRec.lGross = 0
        tmPBSbf(ilLoop).SbfRec.iMnfItem = 0
        tmPBSbf(ilLoop).SbfRec.iNoItems = 0
        'tmPBSbf(ilLoop).SbfRec.sUnitName = ""
        tmPBSbf(ilLoop).SbfRec.sDescr = ""
        tmPBSbf(ilLoop).SbfRec.iPrintInvDate(0) = 0
        tmPBSbf(ilLoop).SbfRec.iPrintInvDate(1) = 0
        tmPBSbf(ilLoop).SbfRec.sBilled = "N"
        tmPBSbf(ilLoop).SbfRec.iIhfCode = 0
        tmPBSbf(ilLoop).SbfRec.iLineNo = 0
    Next ilLoop
    Exit Sub
mMovePBCtrlToRecErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMovePBRecToCtrl                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMovePBRecToCtrl()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mMovePBRecToCtrlErr                                                                   *
'******************************************************************************************

'
'   mMovePBRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim ilUpper As Integer
    'ReDim smPBSave(1 To 3, 1 To 1) As String
    'ReDim imPBSave(1 To 1, 1 To 1) As Integer
    'ReDim smPBShow(1 To MPBAMOUNTINDEX, 1 To 1) As String
    ReDim smPBSave(0 To 3, 0 To 1) As String
    'ReDim imPBSave(0 To 1) As Integer
    ReDim imPBSave(0 To 1, 0 To 1) As Integer
    'ReDim smPBShow(0 To MPBAMOUNTINDEX, 1 To 1) As String
    ReDim smPBShow(0 To MPBAMOUNTINDEX, 0 To 1) As String
    ilUpper = 1
    For ilLoop = 0 To UBound(tmPBSbf) - 1 Step 1
        If (tmPBSbf(ilLoop).iStatus = 0) Or (tmPBSbf(ilLoop).iStatus = 1) Then
            imPBSave(1, ilUpper) = -1
            slRecCode = Trim$(str$(tmPBSbf(ilLoop).SbfRec.iBillVefCode))
            For ilTest = 0 To lbcVehicle.ListCount - 1 Step 1
                slNameCode = lbcVehicle.List(ilTest)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                'On Error GoTo mMovePBRecToCtrlErr
                'gCPErrorMsg ilRet, "mMovePBRecToCtrl (gParseItem field 2)", CBill
                On Error GoTo 0
                If slRecCode = slCode Then
                    imPBSave(1, ilUpper) = ilTest
                    Exit For
                End If
            Next ilTest
            'If imPBSave(1, ilLoop + 1) = -1 Then
            '    gFindMatch sgUserDefVehicleName, 0, lbcPSVehicle
            '    If gLastFound(lbcPSVehicle) >= 0 Then
            '        imPBSave(1, ilLoop + 1) = gLastFound(lbcPSVehicle)
            '    Else
            '        imPBSave(1, ilLoop + 1) = 0
            '    End If
            'End If
            If imPBSave(1, ilUpper) <> -1 Then
                smPBSave(1, ilUpper) = "P"
                gUnpackDate tmPBSbf(ilLoop).SbfRec.iDate(0), tmPBSbf(ilLoop).SbfRec.iDate(1), slDate
                'gFindMatch slDate, 0, lbcBDate
                'If gLastFound(lbcBDate) >= 0 Then
                '    imPBSave(2, ilUpper) = gLastFound(lbcBDate)
                'Else
                'End If
                smPBSave(3, ilUpper) = slDate
                'gPDNToStr tmPBSbf(ilLoop).SbfRec.sItemAmount, 2, smPBSave(2, ilUpper)
                smPBSave(2, ilUpper) = gLongToStrDec(tmPBSbf(ilLoop).SbfRec.lGross, 2)
                ilUpper = ilUpper + 1
                'ReDim Preserve smPBSave(1 To 3, 1 To ilUpper) As String
                'ReDim Preserve imPBSave(1 To 1, 1 To ilUpper) As Integer
                'ReDim Preserve smPBShow(1 To MPBAMOUNTINDEX, 1 To ilUpper) As String
                ReDim Preserve smPBSave(0 To 3, 0 To ilUpper) As String
                'ReDim Preserve imPBSave(0 To ilUpper) As Integer
                ReDim Preserve imPBSave(0 To 1, 0 To ilUpper) As Integer
                ReDim Preserve smPBShow(0 To MPBAMOUNTINDEX, 0 To ilUpper) As String
            Else
                If tmPBSbf(ilLoop).iStatus = 0 Then
                    tmPBSbf(ilLoop).iStatus = -1
                ElseIf tmPBSbf(ilLoop).iStatus = 1 Then
                    tmPBSbf(ilLoop).iStatus = 2
                End If
                tmPBSbf(ilLoop).SbfRec.sTranType = ""
                tmPBSbf(ilLoop).SbfRec.iBillVefCode = 0
                tmPBSbf(ilLoop).SbfRec.iDate(0) = 0
                tmPBSbf(ilLoop).SbfRec.iDate(1) = 0
                'gStrToPDN "", 2, 5, tmPBSbf(ilLoop).SbfRec.sItemAmount
                tmPBSbf(ilLoop).SbfRec.lGross = 0
                tmPBSbf(ilLoop).SbfRec.iMnfItem = 0
                tmPBSbf(ilLoop).SbfRec.iNoItems = 0
                'tmPBSbf(ilLoop).SbfRec.sUnitName = ""
                tmPBSbf(ilLoop).SbfRec.sDescr = ""
                tmPBSbf(ilLoop).SbfRec.iPrintInvDate(0) = 0
                tmPBSbf(ilLoop).SbfRec.iPrintInvDate(1) = 0
                tmPBSbf(ilLoop).SbfRec.sBilled = "N"
                tmPBSbf(ilLoop).SbfRec.iIhfCode = 0
                tmPBSbf(ilLoop).SbfRec.iLineNo = 0
            End If
        End If
    Next ilLoop
    'mInitNewPB UBound(smPBSave, 2)
    Exit Sub
mMovePBRecToCtrlErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMSEnableBox                    *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mMSEnableBox(ilBoxNo As Integer)
'
'   mMSEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slDate As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmMSCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case MPSPERCENTINDEX 'Percent
            If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
                edcMPercent.MaxLength = 0
            Else
                edcMPercent.MaxLength = 6
            End If
            If edcMPercent.Text = "" Then
                edcMPercent.Text = smMPercent
            End If
            edcMPercent.Width = tmMSCtrls(MPSPERCENTINDEX).fBoxW
            gMoveFormCtrl pbcMPSpec, edcMPercent, tmMSCtrls(MPSPERCENTINDEX).fBoxX, tmMSCtrls(MPSPERCENTINDEX).fBoxY
            edcMPercent.Visible = True  'Set visibility
            edcMPercent.SetFocus
        Case MPSSTARTDATEINDEX 'Date
            'gShowHelpMess tmSbfHelp(), SBFPKGDATE
            lbcMSSDate.height = gListBoxHeight(lbcMSSDate.ListCount, 8)
            edcPSDropDown.Width = tmMSCtrls(MPSSTARTDATEINDEX).fBoxW - cmcPSDropDown.Width
            edcPSDropDown.MaxLength = 10
            gMoveFormCtrl pbcMPSpec, edcPSDropDown, tmMSCtrls(MPSSTARTDATEINDEX).fBoxX, tmMSCtrls(MPSSTARTDATEINDEX).fBoxY
            cmcPSDropDown.Move edcPSDropDown.Left + edcPSDropDown.Width, edcPSDropDown.Top
            imChgMode = True
            If lbcMSSDate.ListIndex < 0 Then
                For ilLoop = 0 To lbcMSSDate.ListCount - 1 Step 1
                    slDate = lbcMSSDate.List(ilLoop)
                    If gDateValue(slDate) > lmLastBilledDate Then
                        lbcMSSDate.ListIndex = ilLoop
                        edcPSDropDown.Text = lbcMSSDate.List(ilLoop)
                        Exit For
                    End If
                Next ilLoop
            Else
                edcPSDropDown.Text = lbcMSSDate.List(lbcMSSDate.ListIndex)
            End If
            imChgMode = False
            lbcMSSDate.Move edcPSDropDown.Left, edcPSDropDown.Top + edcPSDropDown.height
            edcPSDropDown.SelStart = 0
            edcPSDropDown.SelLength = Len(edcPSDropDown.Text)
            edcPSDropDown.Visible = True
            cmcPSDropDown.Visible = True
            edcPSDropDown.SetFocus
        Case MPSENDDATEINDEX 'Date
            'gShowHelpMess tmSbfHelp(), SBFPKGDATE
            If lbcMESDate.ListIndex >= 0 Then
                slDate = lbcMESDate.List(lbcMESDate.ListIndex)
            Else
                slDate = lbcMSSDate.List(lbcMSSDate.ListCount - 1)
            End If
            lbcMESDate.Clear
            If lbcMSSDate.ListIndex >= 0 Then
                For ilLoop = lbcMSSDate.ListIndex To lbcMSSDate.ListCount - 1 Step 1
                    lbcMESDate.AddItem lbcMSSDate.List(ilLoop)
                Next ilLoop
            Else
                For ilLoop = 0 To lbcMSSDate.ListCount - 1 Step 1
                    lbcMESDate.AddItem lbcMSSDate.List(ilLoop)
                Next ilLoop
            End If
            lbcMESDate.height = gListBoxHeight(lbcMESDate.ListCount, 8)
            edcPSDropDown.Width = tmMSCtrls(MPSENDDATEINDEX).fBoxW - cmcPSDropDown.Width
            edcPSDropDown.MaxLength = 10
            gMoveFormCtrl pbcMPSpec, edcPSDropDown, tmMSCtrls(MPSENDDATEINDEX).fBoxX, tmMSCtrls(MPSENDDATEINDEX).fBoxY
            cmcPSDropDown.Move edcPSDropDown.Left + edcPSDropDown.Width, edcPSDropDown.Top
            imChgMode = True
            gFindMatch slDate, 0, lbcMESDate
            If gLastFound(lbcMESDate) < 0 Then
                lbcMESDate.ListIndex = 0 'Start at first end date of period after contract start
                edcPSDropDown.Text = lbcMESDate.List(0)
            Else
                lbcMESDate.ListIndex = gLastFound(lbcMESDate)
                edcPSDropDown.Text = lbcMESDate.List(gLastFound(lbcMESDate))
            End If
            imChgMode = False
            lbcMESDate.Move edcPSDropDown.Left, edcPSDropDown.Top + edcPSDropDown.height
            edcPSDropDown.SelStart = 0
            edcPSDropDown.SelLength = Len(edcPSDropDown.Text)
            edcPSDropDown.Visible = True
            cmcPSDropDown.Visible = True
            edcPSDropDown.SetFocus
    End Select
    mSetGenCommand
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMBSetShow                      *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mMSSetShow(ilBoxNo As Integer)
'
'   mMBSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmMSCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case MPSPERCENTINDEX
            edcMPercent.Visible = False
            slStr = edcMPercent.Text
            smMPercent = slStr
            If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
                gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
            Else
                gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            End If
            gSetShow pbcMPSpec, slStr, tmMSCtrls(ilBoxNo)
        Case MPSSTARTDATEINDEX 'Program index
            lbcMSSDate.Visible = False
            edcPSDropDown.Visible = False
            cmcPSDropDown.Visible = False
            slStr = edcPSDropDown.Text
            If gValidDate(slStr) Then
                slStr = gFormatDate(slStr)
                gSetShow pbcMPSpec, slStr, tmMSCtrls(ilBoxNo)
            Else
                Beep
            End If
        Case MPSENDDATEINDEX 'Program index
            lbcMESDate.Visible = False
            edcPSDropDown.Visible = False
            cmcPSDropDown.Visible = False
            slStr = edcPSDropDown.Text
            If gValidDate(slStr) Then
                slStr = gFormatDate(slStr)
                gSetShow pbcMPSpec, slStr, tmMSCtrls(ilBoxNo)
            Else
                Beep
            End If
    End Select
    mSetGenCommand
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMTotals                        *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute totals                  *
'*                                                     *
'*******************************************************
Private Sub mMTotals()
    Dim slATotal As String
    Dim slBTotal As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slPercent As String

    slATotal = "0"
    slPercent = edcMPercent.Text
    If slPercent = "" Then
        slPercent = smMPercent
    End If
    If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) <> MERCHPROMOBYDOLLAR Then
        If slPercent = "" Then
            lacTotals.Visible = False
            Exit Sub
        End If
    Else
        If (slPercent = "") And (UBound(smMBSave, 2) <= LBONE) Then
            lacTotals.Visible = False
            Exit Sub
        End If
    End If
    'If gStrDecToLong(slPercent, 2) <> 0 Then
    If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) <> MERCHPROMOBYDOLLAR Then
        slBTotal = gDivStr(gMulStr(smLnTNet, slPercent), "100")
        slBTotal = gRoundStr(slBTotal, ".01", 2)
        For ilLoop = LBONE To UBound(smMBSave, 2) - 1 Step 1
            slATotal = gAddStr(slATotal, smMBSave(2, ilLoop))
        Next ilLoop
        slATotal = gRoundStr(slATotal, ".01", 2)
    '    If (slATotal = "0") And (slBTotal = "0") Then
    '        lacTotals.Visible = False
    '        Exit Sub
    '    End If
        lacTotals.ForeColor = BLACK
        If Val(slBTotal) = Val(slATotal) Then
            lacTotals.BackColor = GREEN
        Else
            lacTotals.BackColor = Red
            lacTotals.ForeColor = WHITE
        End If
        lacTotals.Visible = True
        gFormatStr slBTotal, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        slStr = "Totals: " & slPercent & "%*Net $" & slStr
        If slATotal <> "0" Then
            gFormatStr slATotal, FMTLEAVEBLANK + FMTCOMMA, 2, slATotal
            slStr = slStr & "  Allocated $" & slATotal
        End If
    Else
        slBTotal = slPercent
        slBTotal = gRoundStr(slBTotal, ".01", 2)
        For ilLoop = LBONE To UBound(smMBSave, 2) - 1 Step 1
            slATotal = gAddStr(slATotal, smMBSave(2, ilLoop))
        Next ilLoop
        slATotal = gRoundStr(slATotal, ".01", 2)
        lacTotals.ForeColor = BLACK
        If (Val(slBTotal) = Val(slATotal)) Or (slPercent = "") Then
            lacTotals.BackColor = GREEN
        Else
            lacTotals.BackColor = Red
            lacTotals.ForeColor = WHITE
        End If
        lacTotals.Visible = True
        gFormatStr slATotal, FMTLEAVEBLANK + FMTCOMMA, 2, slATotal
        slStr = "Total $" & slATotal
    End If
    lacTotals.Caption = slStr
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPBEnableBox                    *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mPBEnableBox(ilBoxNo As Integer)
'
'   mPBEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmMPBCtrls)) Then
        Exit Sub
    End If

    If (imPBRowNo < vbcMP.Value) Or (imPBRowNo >= vbcMP.Value + vbcMP.LargeChange + 1) Then
        mPBSetShow ilBoxNo
        pbcArrow.Visible = False
        Exit Sub
    End If
    pbcArrow.Move plcMP.Left - pbcArrow.Width - 15, plcMP.Top + tmMPBCtrls(MPBVEHICLEINDEX).fBoxY + (imPBRowNo - vbcMP.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case MPBVEHICLEINDEX 'Vehicle
            'gShowHelpMess tmSbfHelp(), SBFBODYFACIL
        Case MPBDATEINDEX 'Program index
            'gShowHelpMess tmSbfHelp(), SBFBODYDATE
        Case MPBAMOUNTINDEX
            'gShowHelpMess tmSbfHelp(), SBFBODYAMT
            edcAmount.Width = tmMPBCtrls(MPBAMOUNTINDEX).fBoxW
            gMoveTableCtrl pbcMP, edcAmount, tmMPBCtrls(MPBAMOUNTINDEX).fBoxX, tmMPBCtrls(MPBAMOUNTINDEX).fBoxY + (imPBRowNo - vbcMP.Value) * (fgBoxGridH + 15)
            edcAmount.Text = smPBSave(2, imPBRowNo)
            edcAmount.Visible = True  'Set visibility
            edcAmount.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPBSetShow                      *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mPBSetShow(ilBoxNo As Integer)
'
'   mPBSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String

    pbcArrow.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmMPBCtrls)) Then
        Exit Sub
    End If
    If (imPBRowNo < vbcMP.Value) Or (imPBRowNo >= vbcMP.Value + vbcMP.LargeChange + 1) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case MPBVEHICLEINDEX 'Vehicle
        Case MPBDATEINDEX 'Program index
        Case MPBAMOUNTINDEX
            edcAmount.Visible = False
            slStr = edcAmount.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            gSetShow pbcMP, slStr, tmMPBCtrls(ilBoxNo)
            smPBShow(MPBAMOUNTINDEX, imPBRowNo) = tmMPBCtrls(ilBoxNo).sShow
            slStr = edcAmount.Text
            If gCompNumberStr(smPBSave(2, imPBRowNo), slStr) <> 0 Then
                If imPBRowNo < UBound(tmPBSbf) + 1 Then   'New lines set after all fields entered
                    imPBChg = True
                End If
                smPBSave(2, imPBRowNo) = edcAmount.Text
                mPTotals
            End If
    End Select
    mSetCommands
    pbcMP.Cls
    pbcMP_Paint
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPBTestFields                   *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mPBTestFields() As Integer
'
'   iRet = mPBTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    'TTP 10855 - fix potential NTR Overflow errors
    'Dim ilSbf As Integer
    Dim llSbf As Long
    For llSbf = LBound(tmPBSbf) To UBound(tmPBSbf) - 1 Step 1
        If (tmPBSbf(llSbf).iStatus = 0) Or (tmPBSbf(llSbf).iStatus = 1) Then
            If tmPBSbf(llSbf).SbfRec.iBillVefCode <= 0 Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imPBRowNo = llSbf + 1
                imPBBoxNo = MPBVEHICLEINDEX
                mPBTestFields = NO
                Exit Function
            End If
            If (tmPBSbf(llSbf).SbfRec.iDate(0) = 0) And (tmPBSbf(llSbf).SbfRec.iDate(1) = 0) Then
                Screen.MousePointer = vbDefault
                ilRes = MsgBox("Date must be specified", vbOKOnly + vbExclamation, "Incomplete")
                imPBRowNo = llSbf + 1
                imPBBoxNo = MPBDATEINDEX
                mPBTestFields = NO
                Exit Function
            End If
            'gStrToPDN "", 2, 5, slStr
            'If StrComp(tmPBSbf(llSbf).SbfRec.sItemAmount, slStr, 0) = 0 Then
            '    Screen.MousePointer = vbDefault
            '    ilRes = MsgBox("Price must be specified", vbOkOnly + vbExclamation, "Incomplete")
            '    imPBRowNo = llSbf + 1
            '    imPBBoxNo = MPBAMOUNTINDEX
            '    mPBTestFields = No
            '    Exit Function
            'End If
        End If
    Next llSbf
    mPBTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPBTestSaveFields               *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mPBTestSaveFields() As Integer
'
'   iRet = mPBTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    If imPBSave(1, imPBRowNo) < 0 Then
        ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imPBBoxNo = MPBVEHICLEINDEX
        mPBTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smPBSave(3, imPBRowNo)) = "" Then
        ilRes = MsgBox("Date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imPBBoxNo = MPBDATEINDEX
        mPBTestSaveFields = NO
        Exit Function
    End If
    'If smPBSave(2, imPBRowNo) = "" Then
    '    ilRes = MsgBox("Amount must be specified", vbOkOnly + vbExclamation, "Incomplete")
    '    imPBBoxNo = MPBAMOUNTINDEX
    '    mPBTestSaveFields = No
    '    Exit Function
    'End If
    mPBTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPSEnableBox                    *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mPSEnableBox(ilBoxNo As Integer)
'
'   mPSEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slDate As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmPSCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case MPSPERCENTINDEX 'Percent
            If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
                edcPPercent.MaxLength = 0
            Else
                edcPPercent.MaxLength = 6
            End If
            If edcPPercent.Text = "" Then
                edcPPercent.Text = smPPercent
            End If
            edcPPercent.Width = tmPSCtrls(MPSPERCENTINDEX).fBoxW
            gMoveFormCtrl pbcMPSpec, edcPPercent, tmPSCtrls(MPSPERCENTINDEX).fBoxX, tmPSCtrls(MPSPERCENTINDEX).fBoxY
            edcPPercent.Visible = True  'Set visibility
            edcPPercent.SetFocus
        Case MPSSTARTDATEINDEX 'Date
            'gShowHelpMess tmSbfHelp(), SBFPKGDATE
            lbcPSSDate.height = gListBoxHeight(lbcPSSDate.ListCount, 8)
            edcPSDropDown.Width = tmPSCtrls(MPSSTARTDATEINDEX).fBoxW - cmcPSDropDown.Width
            edcPSDropDown.MaxLength = 10
            gMoveFormCtrl pbcMPSpec, edcPSDropDown, tmPSCtrls(MPSSTARTDATEINDEX).fBoxX, tmPSCtrls(MPSSTARTDATEINDEX).fBoxY
            cmcPSDropDown.Move edcPSDropDown.Left + edcPSDropDown.Width, edcPSDropDown.Top
            imChgMode = True
            If lbcPSSDate.ListIndex < 0 Then
                For ilLoop = 0 To lbcPSSDate.ListCount - 1 Step 1
                    slDate = lbcPSSDate.List(ilLoop)
                    If gDateValue(slDate) > lmLastBilledDate Then
                        lbcPSSDate.ListIndex = ilLoop
                        edcPSDropDown.Text = lbcPSSDate.List(ilLoop)
                        Exit For
                    End If
                Next ilLoop
            Else
                edcPSDropDown.Text = lbcPSSDate.List(lbcPSSDate.ListIndex)
            End If
            imChgMode = False
            lbcPSSDate.Move edcPSDropDown.Left, edcPSDropDown.Top + edcPSDropDown.height
            edcPSDropDown.SelStart = 0
            edcPSDropDown.SelLength = Len(edcPSDropDown.Text)
            edcPSDropDown.Visible = True
            cmcPSDropDown.Visible = True
            edcPSDropDown.SetFocus
        Case MPSENDDATEINDEX 'Date
            'gShowHelpMess tmSbfHelp(), SBFPKGDATE
            If lbcPESDate.ListIndex >= 0 Then
                slDate = lbcPESDate.List(lbcPESDate.ListIndex)
            Else
                slDate = lbcPSSDate.List(lbcPSSDate.ListCount - 1)
            End If
            lbcPESDate.Clear
            If lbcPSSDate.ListIndex >= 0 Then
                For ilLoop = lbcPSSDate.ListIndex To lbcPSSDate.ListCount - 1 Step 1
                    lbcPESDate.AddItem lbcPSSDate.List(ilLoop)
                Next ilLoop
            Else
                For ilLoop = 0 To lbcPSSDate.ListCount - 1 Step 1
                    lbcPESDate.AddItem lbcPSSDate.List(ilLoop)
                Next ilLoop
            End If
            lbcPESDate.height = gListBoxHeight(lbcPESDate.ListCount, 8)
            edcPSDropDown.Width = tmPSCtrls(MPSENDDATEINDEX).fBoxW - cmcPSDropDown.Width
            edcPSDropDown.MaxLength = 10
            gMoveFormCtrl pbcMPSpec, edcPSDropDown, tmPSCtrls(MPSENDDATEINDEX).fBoxX, tmPSCtrls(MPSENDDATEINDEX).fBoxY
            cmcPSDropDown.Move edcPSDropDown.Left + edcPSDropDown.Width, edcPSDropDown.Top
            imChgMode = True
            gFindMatch slDate, 0, lbcPESDate
            If gLastFound(lbcPESDate) < 0 Then
                lbcPESDate.ListIndex = 0 'Start at first end date of period after contract start
                edcPSDropDown.Text = lbcPESDate.List(0)
            Else
                lbcPESDate.ListIndex = gLastFound(lbcPESDate)
                edcPSDropDown.Text = lbcPESDate.List(gLastFound(lbcPESDate))
            End If
            imChgMode = False
            lbcPESDate.Move edcPSDropDown.Left, edcPSDropDown.Top + edcPSDropDown.height
            edcPSDropDown.SelStart = 0
            edcPSDropDown.SelLength = Len(edcPSDropDown.Text)
            edcPSDropDown.Visible = True
            cmcPSDropDown.Visible = True
            edcPSDropDown.SetFocus
    End Select
    mSetGenCommand
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPSSetShow                      *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mPSSetShow(ilBoxNo As Integer)
'
'   mPSSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmPSCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case MPSPERCENTINDEX
            edcPPercent.Visible = False
            slStr = edcPPercent.Text
            smPPercent = slStr
            If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
                gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
            Else
                gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            End If
            gSetShow pbcMPSpec, slStr, tmPSCtrls(ilBoxNo)
        Case MPSSTARTDATEINDEX 'Program index
            lbcPSSDate.Visible = False
            edcPSDropDown.Visible = False
            cmcPSDropDown.Visible = False
            slStr = edcPSDropDown.Text
            If gValidDate(slStr) Then
                slStr = gFormatDate(slStr)
                gSetShow pbcMPSpec, slStr, tmPSCtrls(ilBoxNo)
            Else
                Beep
            End If
        Case MPSENDDATEINDEX 'Program index
            lbcPESDate.Visible = False
            edcPSDropDown.Visible = False
            cmcPSDropDown.Visible = False
            slStr = edcPSDropDown.Text
            If gValidDate(slStr) Then
                slStr = gFormatDate(slStr)
                gSetShow pbcMPSpec, slStr, tmPSCtrls(ilBoxNo)
            Else
                Beep
            End If
    End Select
    mSetGenCommand
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPTotals                        *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute totals                  *
'*                                                     *
'*******************************************************
Private Sub mPTotals()
    Dim slATotal As String
    Dim slBTotal As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slPercent As String

    slATotal = "0"
    slPercent = edcPPercent.Text
    If slPercent = "" Then
        slPercent = smPPercent
    End If
    If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) <> MERCHPROMOBYDOLLAR Then
        If slPercent = "" Then
            lacTotals.Visible = False
            Exit Sub
        End If
    Else
        If (slPercent = "") And (UBound(smPBSave, 2) <= LBONE) Then
            lacTotals.Visible = False
            Exit Sub
        End If
    End If
    If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) <> MERCHPROMOBYDOLLAR Then
        slBTotal = gDivStr(gMulStr(smLnTNet, slPercent), "100")
        slBTotal = gRoundStr(slBTotal, ".01", 2)
        For ilLoop = LBONE To UBound(smPBSave, 2) - 1 Step 1
            slATotal = gAddStr(slATotal, smPBSave(2, ilLoop))
        Next ilLoop
        slATotal = gRoundStr(slATotal, ".01", 2)
    '    If (slATotal = "0") And (slBTotal = "0") Then
    '        lacTotals.Visible = False
    '        Exit Sub
    '    End If
        lacTotals.ForeColor = BLACK
        If Val(slBTotal) = Val(slATotal) Then
            lacTotals.BackColor = GREEN
        Else
            lacTotals.BackColor = Red
            lacTotals.ForeColor = WHITE
        End If
        lacTotals.Visible = True
        gFormatStr slBTotal, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
        slStr = "Totals: " & slPercent & "%*Net $" & slStr
        If slATotal <> "0" Then
            gFormatStr slATotal, FMTLEAVEBLANK + FMTCOMMA, 2, slATotal
            slStr = slStr & "  Allocated $" & slATotal
        End If
    Else
        slBTotal = slPercent
        slBTotal = gRoundStr(slBTotal, ".01", 2)
        For ilLoop = LBONE To UBound(smPBSave, 2) - 1 Step 1
            slATotal = gAddStr(slATotal, smPBSave(2, ilLoop))
        Next ilLoop
        slATotal = gRoundStr(slATotal, ".01", 2)
        lacTotals.ForeColor = BLACK
        If (Val(slBTotal) = Val(slATotal)) Or (slPercent = "") Then
            lacTotals.BackColor = GREEN
        Else
            lacTotals.BackColor = Red
            lacTotals.ForeColor = WHITE
        End If
        lacTotals.Visible = True
        gFormatStr slATotal, FMTLEAVEBLANK + FMTCOMMA, 2, slATotal
        slStr = "Total $" & slATotal
    End If
    lacTotals.Caption = slStr
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read package/item bill records *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilReSet As Integer) As Integer
'
'   iRet = mReadRec(ilReSet)
'   Where:
'       ilReSet (I)_ 0=All; 1=Fix only; 2=Item bill only; 3=Merchandising; 4= Promotion
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilFBUpperBound As Integer
    Dim ilIBUpperBound As Integer
    Dim ilMBUpperBound As Integer
    Dim ilPBUpperBound As Integer
    Dim ilLoop As Integer
    Dim ilCount As Integer
    If (ilReSet = 0) Or (ilReSet = 1) Then
        smFBBTotal = "0"
        imFBChg = False
        ilFBUpperBound = UBound(tgFBSbf)
        ReDim tmFBSbf(0 To ilFBUpperBound) As SBFLIST
        ReDim lmFBSbfCode(0 To ilFBUpperBound) As Long
        ilCount = 0
        If UBound(tgFBSbf) > LBound(tgFBSbf) Then
            For ilLoop = LBound(tgFBSbf) To UBound(tgFBSbf) - 1 Step 1
                If (tgFBSbf(ilLoop).iStatus = 0) Or (tgFBSbf(ilLoop).iStatus = 1) Then
                    LSet tmFBSbf(ilLoop) = tgFBSbf(ilLoop)
                    ilCount = ilCount + 1
                End If
                If (tgFBSbf(ilLoop).iStatus = 1) Or (tgFBSbf(ilLoop).iStatus = 2) Then
                    lmFBSbfCode(ilLoop) = tgFBSbf(ilLoop).lRecPos
                Else
                    lmFBSbfCode(ilLoop) = 0
                End If
            Next ilLoop
            If ilCount < ilFBUpperBound Then
            ReDim Preserve tmFBSbf(0 To ilCount) As SBFLIST
            End If
        End If
        'ReDim smFBSave(1 To 3, 1 To UBound(tmFBSbf) + 1) As String
        'ReDim imFBSave(1 To 2, 1 To UBound(tmFBSbf) + 1) As Integer
        'ReDim smFBShow(1 To FBBILLTOTALINDEX, 1 To UBound(tmFBSbf) + 1) As String
    End If
    If (ilReSet = 0) Or (ilReSet = 2) Then
        imIBChg = False
        ilIBUpperBound = UBound(tgIBSbf)
        ReDim tmIBSbf(0 To ilIBUpperBound) As SBFLIST
        ReDim lmIBSbfCode(0 To ilIBUpperBound) As Long
        smIBBTotal = "0"
        smIBPTotal = "0"
        ilCount = 0
        If UBound(tgIBSbf) > LBound(tgIBSbf) Then
            For ilLoop = LBound(tgIBSbf) To UBound(tgIBSbf) - 1 Step 1
                If (tgIBSbf(ilLoop).iStatus = 0) Or (tgIBSbf(ilLoop).iStatus = 1) Then
                    LSet tmIBSbf(ilCount) = tgIBSbf(ilLoop)
                    ilCount = ilCount + 1
                End If
                If (tgIBSbf(ilLoop).iStatus = 1) Or (tgIBSbf(ilLoop).iStatus = 2) Then
                    lmIBSbfCode(ilLoop) = tgIBSbf(ilLoop).lRecPos
                Else
                    lmIBSbfCode(ilLoop) = 0
                End If
            Next ilLoop
            If ilCount < ilIBUpperBound Then
                ReDim Preserve tmIBSbf(0 To ilCount) As SBFLIST
            End If
        End If
        'ReDim smIBSave(1 To 11, 1 To UBound(tmIBSbf) + 1) As String
        'ReDim imIBSave(1 To 7, 1 To UBound(tmIBSbf) + 1) As Integer
        'ReDim smIBShow(1 To IBACQCOSTINDEX, 1 To UBound(tmIBSbf) + 1) As String
        'ReDim lmIBSave(1 To 2, 1 To UBound(tmIBSbf) + 1) As Long
        ReDim smIBSave(0 To 11, 0 To UBound(tmIBSbf) + 1) As String
        ReDim imIBSave(0 To 7, 0 To UBound(tmIBSbf) + 1) As Integer
        ReDim smIBShow(0 To IBACQCOSTINDEX, 0 To UBound(tmIBSbf) + 1) As String
        ReDim lmIBSave(0 To 2, 0 To UBound(tmIBSbf) + 1) As Long
    End If
    If (ilReSet = 0) Or (ilReSet = 3) Then
        smMBTotal = "0"
        imMBChg = False
        ilMBUpperBound = UBound(tgMBSbf)
        ReDim tmMBSbf(0 To ilMBUpperBound) As SBFLIST
        ilCount = 0
        For ilLoop = LBound(tgMBSbf) To UBound(tgMBSbf) - 1 Step 1
            LSet tmMBSbf(ilLoop) = tgMBSbf(ilLoop)
            If (tmMBSbf(ilLoop).iStatus = 0) Or (tmMBSbf(ilLoop).iStatus = 1) Then
                ilCount = ilCount + 1
            End If
        Next ilLoop
        ''ReDim smMBSave(1 To 2, 1 To ilCount + 1) As String
        ''ReDim imMBSave(1 To 2, 1 To ilCount + 1) As Integer
        ''ReDim smMBShow(1 To MPBAMOUNTINDEX, 1 To ilCount + 1) As String
        'ReDim smMBSave(1 To 3, 1 To 1) As String
        'ReDim imMBSave(1 To 1, 1 To 1) As Integer
        'ReDim smMBShow(1 To MPBAMOUNTINDEX, 1 To 1) As String
        ReDim smMBSave(0 To 3, 0 To 1) As String
        'ReDim imMBSave(0 To 1) As Integer
        ReDim imMBSave(0 To 1, 0 To 1) As Integer
        ReDim smMBShow(0 To MPBAMOUNTINDEX, 0 To 1) As String
    End If
    If (ilReSet = 0) Or (ilReSet = 4) Then
        smPBTotal = "0"
        imPBChg = False
        ilPBUpperBound = UBound(tgPBSbf)
        ReDim tmPBSbf(0 To ilPBUpperBound) As SBFLIST
        ilCount = 0
        For ilLoop = LBound(tgPBSbf) To UBound(tgPBSbf) - 1 Step 1
            LSet tmPBSbf(ilLoop) = tgPBSbf(ilLoop)
            If (tmPBSbf(ilLoop).iStatus = 0) Or (tmPBSbf(ilLoop).iStatus = 1) Then
                ilCount = ilCount + 1
            End If
        Next ilLoop
        ''ReDim smPBSave(1 To 2, 1 To ilCount + 1) As String
        ''ReDim imPBSave(1 To 2, 1 To ilCount + 1) As Integer
        ''ReDim smPBShow(1 To MPBAMOUNTINDEX, 1 To ilCount + 1) As String
        'ReDim smPBSave(1 To 3, 1 To 1) As String
        'ReDim imPBSave(1 To 1, 1 To 1) As Integer
        'ReDim smPBShow(1 To MPBAMOUNTINDEX, 1 To 1) As String
        ReDim smPBSave(0 To 3, 0 To 1) As String
        'ReDim imPBSave(0 To 1) As Integer
        ReDim imPBSave(0 To 1, 0 To 1) As Integer
        ReDim smPBShow(0 To MPBAMOUNTINDEX, 0 To 1) As String
    End If
    mReadRec = True
    Exit Function

    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
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
    'TTP 10855 - fix potential NTR Overflow errors
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim ilFound As Integer

    Dim ilFBUpperBound As Integer
    Dim ilIBUpperBound As Integer
    Dim ilMBUpperBound As Integer
    Dim ilPBUpperBound As Integer
    Screen.MousePointer = vbHourglass
    mFBSetShow imFBBoxNo
    mIBSetShow imIBBoxNo
    mMoveFBCtrlToRec
    If mFBTestFields() = NO Then
        rbcOption(0).Value = True
        mSaveRec = False
        Exit Function
    End If
    mMoveIBCtrlToRec
    If mIBTestFields() = NO Then
        rbcOption(1).Value = True
        mSaveRec = False
        Exit Function
    End If
    mMoveMBCtrlToRec
    If mMBTestFields() = NO Then
        rbcOption(2).Value = True
        mSaveRec = False
        Exit Function
    End If
    mMovePBCtrlToRec
    If mPBTestFields() = NO Then
        rbcOption(3).Value = True
        mSaveRec = False
        Exit Function
    End If
    If imFBChg Or imIBChg Or imMBChg Or imPBChg Then
        igPkgChgd = YES
    End If
    ilFBUpperBound = UBound(tmFBSbf)
    ReDim tgFBSbf(0 To ilFBUpperBound)
    ilIBUpperBound = UBound(tmIBSbf)
    ReDim tgIBSbf(0 To ilIBUpperBound)
    ilMBUpperBound = UBound(tmMBSbf)
    ReDim tgMBSbf(0 To ilMBUpperBound)
    ilPBUpperBound = UBound(tmPBSbf)
    ReDim tgPBSbf(0 To ilPBUpperBound)
    For ilLoop = LBound(tmFBSbf) To UBound(tmFBSbf) Step 1
        LSet tgFBSbf(ilLoop) = tmFBSbf(ilLoop)
    Next ilLoop
    For ilLoop = LBound(lmFBSbfCode) To UBound(lmFBSbfCode) - 1 Step 1
        If lmFBSbfCode(ilLoop) <> 0 Then
            ilFound = False
            For llSbf = LBound(tgFBSbf) To UBound(tgFBSbf) - 1 Step 1
                If (tgFBSbf(llSbf).iStatus <> -1) Then
                    If tgFBSbf(llSbf).lRecPos = lmFBSbfCode(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next llSbf
            If Not ilFound Then
                ilFBUpperBound = UBound(tgFBSbf)
                tgFBSbf(ilFBUpperBound).iStatus = 2
                tgFBSbf(ilFBUpperBound).lRecPos = lmFBSbfCode(ilLoop)
                ReDim Preserve tgFBSbf(0 To ilFBUpperBound + 1)
            End If
        End If
    Next ilLoop
    For ilLoop = LBound(tmIBSbf) To UBound(tmIBSbf) Step 1
        LSet tgIBSbf(ilLoop) = tmIBSbf(ilLoop)
    Next ilLoop
    For ilLoop = LBound(lmIBSbfCode) To UBound(lmIBSbfCode) - 1 Step 1
        If lmIBSbfCode(ilLoop) <> 0 Then
            ilFound = False
            For llSbf = LBound(tgIBSbf) To UBound(tgIBSbf) - 1 Step 1
                If (tgIBSbf(llSbf).iStatus <> -1) Then
                    If tgIBSbf(llSbf).lRecPos = lmIBSbfCode(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next llSbf
            If Not ilFound Then
                ilIBUpperBound = UBound(tgIBSbf)
                tgIBSbf(ilIBUpperBound).iStatus = 2
                tgIBSbf(ilIBUpperBound).lRecPos = lmIBSbfCode(ilLoop)
                ReDim Preserve tgIBSbf(0 To ilIBUpperBound + 1)
            End If
        End If
    Next ilLoop
    'tgChfCntr.iMerchPct = gStrDecToInt(smMPercent, 3)
    If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
        tgChfCntr.iMerchPct = 0
    Else
        tgChfCntr.iMerchPct = gStrDecToInt(smMPercent, 2)
    End If
    For ilLoop = LBound(tmMBSbf) To UBound(tmMBSbf) Step 1
        LSet tgMBSbf(ilLoop) = tmMBSbf(ilLoop)
    Next ilLoop
    For ilLoop = LBound(tmPBSbf) To UBound(tmPBSbf) Step 1
        LSet tgPBSbf(ilLoop) = tmPBSbf(ilLoop)
    Next ilLoop
    'tgChfCntr.iPromoPct = gStrDecToInt(smPPercent, 3)
    If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
        tgChfCntr.iPromoPct = 0
    Else
        tgChfCntr.iPromoPct = gStrDecToInt(smPPercent, 2)
    End If
    'mInitNewFB UBound(smFBSave, 2)
    mInitNewIB UBound(smIBSave, 2)
    mSaveRec = True
    Screen.MousePointer = vbDefault
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
'*             Created:9/05/93       By:D. LeVine      *
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
    Dim slATotal As String
    Dim slBTotal As String
    Dim ilLoop As Integer
    If imFBChg Or imIBChg Or imMBChg Or imPBChg Then
        If ilAsk Then
            slMess = "Save Changes"
            If (smMPercent <> "") And imMBChg Then
                slATotal = ".00"
                If ((Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR) Then
                    slBTotal = smMPercent
                Else
                    slBTotal = gDivStr(gMulStr(smLnTNet, smMPercent), "100")
                End If
                slBTotal = gRoundStr(slBTotal, ".01", 2)
                For ilLoop = LBONE To UBound(smMBSave, 2) - 1 Step 1
                    slATotal = gAddStr(slATotal, smMBSave(2, ilLoop))
                Next ilLoop
                slATotal = gRoundStr(slATotal, ".01", 2)
                If (Val(slBTotal) <> Val(slATotal)) And (gStrDecToLong(smMPercent, 2) <> 0) Then
                    slMess = "Merchandising Out of Balance, Save Changes"
                End If
            End If
            If (smPPercent <> "") And imPBChg Then
                slATotal = ".00"
                If ((Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR) Then
                    slBTotal = smPPercent
                Else
                    slBTotal = gDivStr(gMulStr(smLnTNet, smPPercent), "100")
                End If
                slBTotal = gRoundStr(slBTotal, ".01", 2)
                For ilLoop = LBONE To UBound(smPBSave, 2) - 1 Step 1
                    slATotal = gAddStr(slATotal, smPBSave(2, ilLoop))
                Next ilLoop
                slATotal = gRoundStr(slATotal, ".01", 2)
                If Val(slBTotal) <> Val(slATotal) Then
                    slATotal = ".00"
                    If slMess <> "Save Changes" Then
                        slMess = "Merchandising and Promotion Out of Balance, Save Changes"
                    Else
                        slMess = "Promotion Out of Balance, Save Changes"
                    End If
                End If
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
            End If
        Else
            slMess = ""
            If (smMPercent <> "") And imMBChg Then
                slATotal = ".00"
                If ((Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR) Then
                    slBTotal = smMPercent
                Else
                    slBTotal = gDivStr(gMulStr(smLnTNet, smMPercent), "100")
                End If
                slBTotal = gRoundStr(slBTotal, ".01", 2)
                For ilLoop = LBONE To UBound(smMBSave, 2) - 1 Step 1
                    slATotal = gAddStr(slATotal, smMBSave(2, ilLoop))
                Next ilLoop
                slATotal = gRoundStr(slATotal, ".01", 2)
                If (Val(slBTotal) <> Val(slATotal)) And (gStrDecToLong(smMPercent, 2) <> 0) Then
                    slMess = "Merchandising Out of Balance, Retain Changes"
                End If
            End If
            If (smPPercent <> "") And imPBChg Then
                slATotal = ".00"
                If ((Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR) Then
                    slBTotal = smPPercent
                Else
                    slBTotal = gDivStr(gMulStr(smLnTNet, smPPercent), "100")
                End If
                slBTotal = gRoundStr(slBTotal, ".01", 2)
                For ilLoop = LBONE To UBound(smPBSave, 2) - 1 Step 1
                    slATotal = gAddStr(slATotal, smPBSave(2, ilLoop))
                Next ilLoop
                slATotal = gRoundStr(slATotal, ".01", 2)
                If Val(slBTotal) <> Val(slATotal) Then
                    If slMess <> "" Then
                        slMess = "Merchandising and Promotion Out of Balance, Retain Changes"
                    Else
                        slMess = "Promotion Out of Balance, Retain Changes"
                    End If
                End If
            End If
            If slMess <> "" Then
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
'*             Created:4/20/93       By:D. LeVine      *
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
    'Revert button set if any field changed
    If imFBChg Or imIBChg Or imMBChg Or imPBChg Then
        cmcUndo.Enabled = True
        RaiseEvent SetSave(True)
    Else
        RaiseEvent SetSave(False)
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:9/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetGenCommand()
    If rbcOption(0).Value Then
        If (lbcPSDate.ListIndex >= 0) And (edcNoPeriods.Text <> "") And (lbcPEDate.ListIndex >= 0) Then
            cmcGen.Enabled = True
        Else
            cmcGen.Enabled = False
        End If
    ElseIf rbcOption(2).Value Then
        If (edcMPercent.Text <> "") And (lbcMSSDate.ListIndex >= 0) And (lbcMESDate.ListIndex >= 0) Then
            cmcGen.Enabled = True
        Else
            cmcGen.Enabled = False
        End If
    ElseIf rbcOption(3).Value Then
        If (edcPPercent.Text <> "") And (lbcPSSDate.ListIndex >= 0) And (lbcPESDate.ListIndex >= 0) Then
            cmcGen.Enabled = True
        Else
            cmcGen.Enabled = False
        End If
    End If
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
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    ilRet = btrClose(hmIhf)
    btrDestroy hmIhf
    'Deallocate memory
    Erase lmFBSbfCode
    Erase lmIBSbfCode
    Erase tmFBSbf
    Erase tmIBSbf
    Erase tmItemCode
    Erase tmInstallBillInfo
    Erase tmInstallVehInfo
    'Erase smFBSave
    'Erase imFBSave
    'Erase smFBShow
    Erase smIBSave
    Erase imIBSave
    Erase smIBShow
    Erase lmIBSave
    Erase smMBSave
    Erase imMBSave
    Erase smMBShow
    Erase smPBSave
    Erase imPBSave
    Erase smPBShow
    Erase tmSbfHelp
    Erase tmTaxSortCode
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload CBill
    'Set CBill = Nothing
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the Vehicle box      *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilIndex As Integer
    Dim tlVsf As VSF
    Dim hlVsf As Integer
    Dim ilRecLen As Integer     'Vsf record length
    Dim tlSrchKey As INTKEY0
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilFound As Integer
    Dim ilMkt As Integer
    'TTP 10855 - fix potential NTR Overflow errors
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim ilVef As Integer
    Dim blIncludeVehicle As Boolean

    imPSNonPkgStartIndex = -1
    imBPkgStartIndex = -1
    imBNonPkgStartIndex = -1

    hlVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'On Error GoTo mVehPopErr
    'gBtrvErrorMsg ilRet, "mVehPop (btrOpen)" & "Vsf.Btr", CBill
    On Error GoTo 0
    ilRecLen = Len(tlVsf)  'btrRecordLength(hlVpf)  'Get and save record length
    lbcVehicle.Clear
    lbcPSVehicle.Clear
    lbcBVehicle.Clear
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        If ((tgClfCntr(ilClf).iStatus = 0) Or (tgClfCntr(ilClf).iStatus = 1)) And (Not tgClfCntr(ilClf).iCancel) Then
            If tgClfCntr(ilClf).ClfRec.iVefCode > 0 Then
                slRecCode = Trim$(str$(tgClfCntr(ilClf).ClfRec.iVefCode))
                For ilTest = 0 To UBound(tmVehicleCode) - 1 Step 1  'Contract!lbcVehicle.ListCount - 1 Step 1
                    slNameCode = tmVehicleCode(ilTest).sKey    'Contract!lbcVehicle.List(ilTest)
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    'On Error GoTo mVehPopErr
                    'gCPErrorMsg ilRet, "mVehPop (gParseItem field 1)", CBill
                    'On Error GoTo 0
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    'On Error GoTo mVehPopErr
                    'gCPErrorMsg ilRet, "mVehPop (gParseItem field 2)", CBill
                    'On Error GoTo 0
                    If slRecCode = slCode Then
                        gFindMatch slNameCode, 0, lbcVehicle
                        If gLastFound(lbcVehicle) < 0 Then
                            lbcVehicle.AddItem slNameCode
                        End If
                        Exit For
                    End If
                Next ilTest
            ElseIf tgClfCntr(ilClf).ClfRec.iVefCode < 0 Then
                tlSrchKey.iCode = -tgClfCntr(ilClf).ClfRec.iVefCode
                ilRet = btrGetEqual(hlVsf, tlVsf, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                'On Error GoTo mVehPopErr
                'gBtrvErrorMsg ilRet, "mVehPop (btrGetEqual)", CBill
                'On Error GoTo 0
                For ilIndex = LBound(tlVsf.iFSCode) To UBound(tlVsf.iFSCode) Step 1
                    If tlVsf.iFSCode(ilIndex) <= 0 Then
                        Exit For
                    End If
                    slRecCode = Trim$(str$(tlVsf.iFSCode(ilIndex)))
                    For ilTest = 0 To UBound(tmVehicleCode) - 1 Step 1  'Contract!lbcVehicle.ListCount - 1 Step 1
                        slNameCode = tmVehicleCode(ilTest).sKey    'Contract!lbcVehicle.List(ilTest)
                        ilRet = gParseItem(slNameCode, 1, "\", slName)
                        'On Error GoTo mVehPopErr
                        'gCPErrorMsg ilRet, "mVehPop (gParseItem field 1)", CBill
                        'On Error GoTo 0
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        'On Error GoTo mVehPopErr
                        'gCPErrorMsg ilRet, "mVehPop (gParseItem field 2)", CBill
                        On Error GoTo 0
                        If slRecCode = slCode Then
                            gFindMatch slNameCode, 0, lbcVehicle
                            If gLastFound(lbcVehicle) < 0 Then
                                lbcVehicle.AddItem slNameCode
                            End If
                            Exit For
                        End If
                    Next ilTest
                Next ilIndex
            End If
        End If
    Next ilClf
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        slNameCode = lbcVehicle.List(ilLoop)
        If Left$(slNameCode, 1) <> "A" Then
            If imPSNonPkgStartIndex = -1 Then
                imPSNonPkgStartIndex = ilLoop
            End If
        End If
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        On Error GoTo mVehPopErr
        'gCPErrorMsg ilRet, "mVehPop (gParseItem field 1)", CBill
        'On Error GoTo 0
        ilRet = gParseItem(slName, 3, "|", slName)
        'On Error GoTo mVehPopErr
        'gCPErrorMsg ilRet, "mVehPop (gParseItem field 1)", CBill
        On Error GoTo 0
        lbcPSVehicle.AddItem Trim$(slName)  'Add ID to list box
        'lbcBVehicle.AddItem Trim$(slName)
    Next ilLoop
    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '8/24/18: Ignore Dormant vehicles
        
        If tgMVef(ilLoop).sState = "D" Then
            blIncludeVehicle = False
            For llSbf = 0 To UBound(tmIBSbf) - 1 Step 1
                If tmIBSbf(llSbf).SbfRec.iBillVefCode = tgMVef(ilLoop).iCode Then
                    blIncludeVehicle = True
                    Exit For
                End If
            Next llSbf
        Else
            blIncludeVehicle = True
        End If
        If (tgMVef(ilLoop).sType = "N") And (blIncludeVehicle) Then
            slName = Trim$(tgMVef(ilLoop).sName) '& "\" & tgMVef(ilLoop).iCode
            'To ignore market of NTR remove this test.
            If (tgSpf.sMktBase = "Y") And (tgMVef(ilLoop).iMnfVehGp3Mkt > 0) Then
                ilFound = False
                For ilMkt = 0 To UBound(igCntrMktCode) - 1 Step 1
                    If tgMVef(ilLoop).iMnfVehGp3Mkt = igCntrMktCode(ilMkt) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilMkt
            Else
                ilFound = True
            End If
            If ilFound Then
                'Place in sorted order
                ilFound = False
                For ilTest = 0 To lbcBVehicle.ListCount - 1 Step 1
                    slNameCode = lbcBVehicle.List(ilTest)
                    If StrComp(slName, slNameCode, vbTextCompare) < 0 Then
                        lbcBVehicle.AddItem slName, ilTest
                        lbcBVehicle.ItemData(ilTest) = tgMVef(ilLoop).iCode
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    lbcBVehicle.AddItem slName
                    lbcBVehicle.ItemData(lbcBVehicle.NewIndex) = tgMVef(ilLoop).iCode
                End If
            End If
        End If
    Next ilLoop
    For ilTest = 0 To UBound(tmVehicleCode) - 1 Step 1  'Contract!lbcVehicle.ListCount - 1 Step 1
        slNameCode = tmVehicleCode(ilTest).sKey    'Contract!lbcVehicle.List(ilTest)
        If Left$(slNameCode, 1) = "A" Then
            If imBPkgStartIndex = -1 Then
                imBPkgStartIndex = lbcBVehicle.ListCount
            End If
        Else
            If imBNonPkgStartIndex = -1 Then
                imBNonPkgStartIndex = lbcBVehicle.ListCount
            End If
        End If
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        'On Error GoTo mVehPopErr
        'gCPErrorMsg ilRet, "mVehPop (gParseItem field 1)", CBill
        On Error GoTo 0
        ilRet = gParseItem(slName, 3, "|", slName)
        'On Error GoTo mVehPopErr
        'gCPErrorMsg ilRet, "mVehPop (gParseItem field 1)", CBill
        On Error GoTo 0
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        
        '8/24/18: Check if vehicle is dormant
        ilVef = gBinarySearchVef(Val(slCode))
        If ilVef <> -1 Then
            If tgMVef(ilVef).sState = "D" Then
                blIncludeVehicle = False
                For llSbf = 0 To UBound(tmIBSbf) - 1 Step 1
                    If tmIBSbf(llSbf).SbfRec.iBillVefCode = tgMVef(ilVef).iCode Then
                        blIncludeVehicle = True
                        Exit For
                    End If
                Next llSbf
            Else
                blIncludeVehicle = True
            End If
        Else
            blIncludeVehicle = False
        End If
        If (blIncludeVehicle) Then
            lbcBVehicle.AddItem Trim$(slName)
            lbcBVehicle.ItemData(lbcBVehicle.NewIndex) = Val(slCode)
        End If
    Next ilTest
    
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    imTerminate = True
    Exit Sub
End Sub

Private Sub lbcPSVehicle_Scroll()
    pbcLbcPSVehicle_Paint
End Sub

Private Sub lbcTax_Click()
    gProcessLbcClick lbcTax, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcTax_GotFocus()
    gCtrlGotFocus ActiveControl
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
    mAllSetShow
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcFBStab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    Dim ilFound As Integer
    Dim slDate As String

    If GetFocus() <> pbcFBSTab.HWnd Then
        Exit Sub
    End If
    mFSSetShow imFSBoxNo
    imFSBoxNo = -1
    mMSSetShow imMSBoxNo
    imMSBoxNo = -1
    mPSSetShow imPSBoxNo
    imPSBoxNo = -1
    If rbcOption(0).Value Then
        Select Case imFBBoxNo
            Case -1 'Tab from control prior to form area
                'If UBound(smFBSave, 2) <= vbcFix.LargeChange + 1 Then 'was <=
                '    vbcFix.Max = LBound(smFBSave, 2)
                'Else
                '    vbcFix.Max = UBound(smFBSave, 2) - vbcFix.LargeChange ' - 1
                'End If
                If UBound(tmInstallBillInfo) + 1 <= vbcFix.LargeChange + 1 Then 'was <=
                    vbcFix.Max = 1  'LBound(smFBSave, 2)
                Else
                    vbcFix.Max = UBound(tmInstallBillInfo) + 1 - vbcFix.LargeChange ' - 1
                End If
                imFBRowNo = 1
                imSettingValue = True
                vbcFix.Value = vbcFix.Min
                imSettingValue = False
                'If (imFBRowNo = UBound(smFBSave, 2)) And (imFBSave(1, 1) = -1) Then
                If (imFBRowNo = UBound(tmInstallBillInfo) + 1) Then
                    'pbcFSSTab.SetFocus
                    Exit Sub
                End If
                ilBox = FBBILLINGINDEX
                imFBBoxNo = ilBox
                mFBEnableBox ilBox
                Exit Sub
            Case FBBILLINGINDEX 'Name (first control within header)
                mFBSetShow imFBBoxNo
                ilBox = FBBILLINGINDEX
                Do
                    If imFBRowNo <= 1 Then
                        imFBBoxNo = -1
                        'cmcDone.SetFocus
                        Exit Sub
                    End If
                    imFBRowNo = imFBRowNo - 1
                    If imFBRowNo < vbcFix.Value Then
                        imSettingValue = True
                        vbcFix.Value = vbcFix.Value - 1
                        imSettingValue = False
                    End If
                Loop While (tmInstallBillInfo(imFBRowNo - 1).sBilledFlag) = "Y"
                imFBBoxNo = ilBox
                mFBEnableBox ilBox
                Exit Sub
            Case FBDATEINDEX
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
                ilBox = FBVEHICLEINDEX
            Case Else
                ilBox = imFBBoxNo - 1
        End Select
        mFBSetShow imFBBoxNo
        imFBBoxNo = ilBox
        mFBEnableBox ilBox
    ElseIf rbcOption(2).Value Then
        Select Case imMBBoxNo
            Case -1 'Tab from control prior to form area
                If UBound(smMBSave, 2) - 1 <= vbcMP.LargeChange + 1 Then 'was <=
                    vbcMP.Max = LBONE  'LBound(smMBSave, 2)
                Else
                    vbcMP.Max = UBound(smMBSave, 2) - vbcMP.LargeChange - 1
                End If
                imMBRowNo = 0
                imSettingValue = True
                vbcMP.Value = vbcMP.Min
                imSettingValue = False
                Do
                    If imMBRowNo >= UBound(smMBSave, 2) - 1 Then
                        imMBBoxNo = -1
                        If lacTotals.BackColor = GREEN Then
                            cmcDone.Enabled = True
                            'cmcDone.SetFocus
                            Exit Sub
                        Else
                            cmcCancel.Enabled = True
                            'cmcCancel.SetFocus
                            Exit Sub
                        End If
                    End If
                    imMBRowNo = imMBRowNo + 1
                    If imMBRowNo > vbcMP.Value + vbcMP.LargeChange Then
                        imSettingValue = True
                        vbcMP.Value = vbcMP.Value + 1
                        imSettingValue = False
                    End If
                    slDate = smMBSave(3, imMBRowNo) 'lbcBDate.List(imMBSave(2, imMBRowNo))
                    If gDateValue(slDate) <= lmLastBilledDate Then
                        ilFound = False
                    Else
                        ilFound = True
                        ilBox = MPBAMOUNTINDEX
                        imMBBoxNo = ilBox
                        mMBEnableBox ilBox
                    End If
                Loop While Not ilFound
                Exit Sub
            Case MPBAMOUNTINDEX 'Name (first control within header)
                mMBSetShow imMBBoxNo
                Do
                    ilBox = MPBAMOUNTINDEX
                    If imMBRowNo <= 1 Then
                        imMBBoxNo = -1
                        'cmcDone.SetFocus
                        Exit Sub
                    End If
                    imMBRowNo = imMBRowNo - 1
                    If imMBRowNo < vbcMP.Value Then
                        imSettingValue = True
                        vbcMP.Value = vbcMP.Value - 1
                        imSettingValue = False
                    End If
                    slDate = smMBSave(3, imMBRowNo) 'lbcBDate.List(imMBSave(2, imMBRowNo))
                    If gDateValue(slDate) <= lmLastBilledDate Then
                        ilFound = False
                    Else
                        ilFound = True
                        imMBBoxNo = ilBox
                        mMBEnableBox ilBox
                    End If
                Loop While Not ilFound
                Exit Sub
            Case Else
                ilBox = imMBBoxNo - 1
        End Select
        mMBSetShow imMBBoxNo
        imMBBoxNo = ilBox
        mMBEnableBox ilBox
    ElseIf rbcOption(3).Value Then
        Select Case imPBBoxNo
            Case -1 'Tab from control prior to form area
                If UBound(smPBSave, 2) - 1 <= vbcMP.LargeChange + 1 Then 'was <=
                    vbcMP.Max = LBONE   'LBound(smPBSave, 2)
                Else
                    vbcMP.Max = UBound(smPBSave, 2) - vbcMP.LargeChange - 1
                End If
                imPBRowNo = 0
                imSettingValue = True
                vbcMP.Value = vbcMP.Min
                imSettingValue = False
                Do
                    If imPBRowNo >= UBound(smPBSave, 2) - 1 Then
                        imPBBoxNo = -1
                        If lacTotals.BackColor = GREEN Then
                            cmcDone.Enabled = True
                            'cmcDone.SetFocus
                            Exit Sub
                        Else
                            cmcCancel.Enabled = True
                            'cmcCancel.SetFocus
                            Exit Sub
                        End If
                    End If
                    imPBRowNo = imPBRowNo + 1
                    If imPBRowNo > vbcMP.Value + vbcMP.LargeChange Then
                        imSettingValue = True
                        vbcMP.Value = vbcMP.Value + 1
                        imSettingValue = False
                    End If
                    slDate = smPBSave(3, imPBRowNo) 'lbcBDate.List(imPBSave(2, imPBRowNo))
                    If gDateValue(slDate) <= lmLastBilledDate Then
                        ilFound = False
                    Else
                        ilFound = True
                        ilBox = MPBAMOUNTINDEX
                        imPBBoxNo = ilBox
                        mPBEnableBox ilBox
                    End If
                Loop While Not ilFound
                Exit Sub
            Case MPBAMOUNTINDEX 'Name (first control within header)
                mPBSetShow imPBBoxNo
                Do
                    ilBox = MPBAMOUNTINDEX
                    If imPBRowNo <= 1 Then
                        imPBBoxNo = -1
                        'cmcDone.SetFocus
                        Exit Sub
                    End If
                    imPBRowNo = imPBRowNo - 1
                    If imPBRowNo < vbcMP.Value Then
                        imSettingValue = True
                        vbcMP.Value = vbcMP.Value - 1
                        imSettingValue = False
                    End If
                    slDate = smPBSave(3, imPBRowNo) 'lbcBDate.List(imPBSave(2, imPBRowNo))
                    If gDateValue(slDate) <= lmLastBilledDate Then
                        ilFound = False
                    Else
                        ilFound = True
                        imPBBoxNo = ilBox
                        mPBEnableBox ilBox
                    End If
                Loop While Not ilFound
                Exit Sub
            Case Else
                ilBox = imPBBoxNo - 1
        End Select
        mPBSetShow imPBBoxNo
        imPBBoxNo = ilBox
        mPBEnableBox ilBox
    End If
End Sub
Private Sub pbcFBTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim slATotal As String
    Dim ilFound As Integer
    Dim slDate As String

    If GetFocus() <> pbcFBTab.HWnd Then
        Exit Sub
    End If
    mFSSetShow imFSBoxNo
    imFSBoxNo = -1
    mMSSetShow imMSBoxNo
    imMSBoxNo = -1
    mPSSetShow imPSBoxNo
    imPSBoxNo = -1
    If rbcOption(0).Value Then
        Select Case imFBBoxNo
            Case -1 'Tab from control prior to form area
                imFBRowNo = UBound(tmInstallBillInfo) + 1 'UBound(smFBSave, 2)
                imSettingValue = True
                If imFBRowNo <= vbcFix.LargeChange + 1 Then
                    vbcFix.Value = vbcFix.Min
                Else
                    vbcFix.Value = imFBRowNo - vbcFix.LargeChange
                End If
                imSettingValue = False
                ilBox = FBBILLINGINDEX
'            Case FBVEHICLEINDEX
'                If (imFBRowNo >= UBound(smFBSave, 2)) And (lbcPSVehicle.ListIndex < 0) Then
'                    mFBSetShow imFBBoxNo
'                    For ilLoop = FBVEHICLEINDEX To FBBILLTOTALINDEX Step 1
'                        slStr = ""
'                        gSetShow pbcFix, slStr, tmFBCtrls(ilLoop)
'                        smFBShow(ilLoop, imFBRowNo) = tmFBCtrls(ilLoop).sShow
'                    Next ilLoop
'                    imFBSave(1, imFBRowNo) = -1
'                    imFBBoxNo = -1
'                    pbcFix_Paint
'                    If cmcDone.Enabled Then
'                        cmcDone.SetFocus
'                    Else
'                        cmcCancel.SetFocus
'                    End If
'                    Exit Sub
'                End If
'                ilBox = FBDATEINDEX
'            Case FBDATEINDEX
'                slStr = edcDropDown.Text
'                If slStr <> "" Then
'                    If Not gValidDate(slStr) Then
'                        Beep
'                        edcDropDown.SetFocus
'                        Exit Sub
'                    End If
'                Else
'                    Beep
'                    edcDropDown.SetFocus
'                    Exit Sub
'                End If
'                ilBox = imFBBoxNo + 1
            Case FBBILLINGINDEX 'Last control
                mFBSetShow imFBBoxNo
                If mFBTestSaveFields() = NO Then
                    mFBEnableBox imFBBoxNo
                    Exit Sub
                End If
                If imFBRowNo >= UBound(tmInstallBillInfo) + 1 Then  'UBound(smFBSave, 2) Then
                    imFBBoxNo = -1
                    slATotal = "0"
                    'For ilLoop = LBound(smFBSave, 2) To UBound(smFBSave, 2) - 1 Step 1
                    '    slATotal = gAddStr(slATotal, smFBSave(2, ilLoop))
                    'Next ilLoop
                    For ilLoop = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
                        slATotal = gAddStr(slATotal, gLongToStrDec(tmInstallBillInfo(ilLoop).lBillDollars, 2))
                    Next ilLoop
                    slATotal = gSubStr(smGross, slATotal)
                    If Val(slATotal) = 0 Then
                        cmcDone.Enabled = True
                        'cmcDone.SetFocus
                        Exit Sub
                    Else
                        imFBBoxNo = -1
                        pbcFSSTab.SetFocus
                        Exit Sub
                    End If
                End If
                Do
                    imFBRowNo = imFBRowNo + 1
                    If imFBRowNo >= UBound(tmInstallBillInfo) + 1 Then  'UBound(smFBSave, 2) Then
                        imFBBoxNo = -1
                        slATotal = "0"
                        'For ilLoop = LBound(smFBSave, 2) To UBound(smFBSave, 2) - 1 Step 1
                        '    slATotal = gAddStr(slATotal, smFBSave(2, ilLoop))
                        'Next ilLoop
                        For ilLoop = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
                            slATotal = gAddStr(slATotal, gLongToStrDec(tmInstallBillInfo(ilLoop).lBillDollars, 2))
                        Next ilLoop
                        slATotal = gSubStr(smGross, slATotal)
                        If Val(slATotal) = 0 Then
                            cmcDone.Enabled = True
                            'cmcDone.SetFocus
                            Exit Sub
                        Else
                            imFBBoxNo = -1
                            pbcFSSTab.SetFocus
                            Exit Sub
                        End If
                    End If
                    If imFBRowNo > vbcFix.Value + vbcFix.LargeChange Then
                        imSettingValue = True
                        vbcFix.Value = vbcFix.Value + 1
                        imSettingValue = False
                    End If
                'Loop While smFBSave(3, imFBRowNo) = "B"
                Loop While tmInstallBillInfo(imFBRowNo - 1).sBilledFlag = "Y"
                ilBox = FBBILLINGINDEX
                imFBBoxNo = ilBox
                mFBEnableBox ilBox
                Exit Sub
            Case Else
                ilBox = imFBBoxNo + 1
        End Select
        mFBSetShow imFBBoxNo
        imFBBoxNo = ilBox
        mFBEnableBox ilBox
    ElseIf rbcOption(2).Value Then
        Select Case imMBBoxNo
            Case -1 'Tab from control prior to form area
                imMBRowNo = UBound(smMBSave, 2) - 1
                imSettingValue = True
                If imMBRowNo <= vbcMP.LargeChange + 1 Then
                    vbcMP.Value = vbcMP.Min
                Else
                    vbcMP.Value = imMBRowNo - vbcMP.LargeChange
                End If
                imSettingValue = False
                ilBox = MPBAMOUNTINDEX
            Case MPBAMOUNTINDEX 'Last control
                mMBSetShow imMBBoxNo
                If mMBTestSaveFields() = NO Then
                    mMBEnableBox imMBBoxNo
                    Exit Sub
                End If
                Do
                    If imMBRowNo >= UBound(smMBSave, 2) - 1 Then
                        imMBBoxNo = -1
                        If lacTotals.BackColor = GREEN Then
                            cmcDone.Enabled = True
                            'cmcDone.SetFocus
                            Exit Sub
                        Else
                            cmcCancel.Enabled = True
                            'cmcCancel.SetFocus
                            Exit Sub
                        End If
                    End If
                    imMBRowNo = imMBRowNo + 1
                    If imMBRowNo > vbcMP.Value + vbcMP.LargeChange Then
                        imSettingValue = True
                        vbcMP.Value = vbcMP.Value + 1
                        imSettingValue = False
                    End If
                    slDate = smMBSave(3, imMBRowNo) 'lbcBDate.List(imMBSave(2, imMBRowNo))
                    If gDateValue(slDate) <= lmLastBilledDate Then
                        ilFound = False
                    Else
                        ilFound = True
                        ilBox = MPBAMOUNTINDEX
                        imMBBoxNo = ilBox
                        mMBEnableBox ilBox
                    End If
                Loop While Not ilFound
                Exit Sub
            Case Else
                ilBox = imMBBoxNo + 1
        End Select
        mMBSetShow imMBBoxNo
        imMBBoxNo = ilBox
        mMBEnableBox ilBox
    ElseIf rbcOption(3).Value Then
        Select Case imPBBoxNo
            Case -1 'Tab from control prior to form area
                imPBRowNo = UBound(smPBSave, 2) - 1
                imSettingValue = True
                If imPBRowNo <= vbcMP.LargeChange + 1 Then
                    vbcMP.Value = vbcMP.Min
                Else
                    vbcMP.Value = imPBRowNo - vbcMP.LargeChange
                End If
                imSettingValue = False
                ilBox = MPBAMOUNTINDEX
            Case MPBAMOUNTINDEX 'Last control
                mPBSetShow imPBBoxNo
                If mPBTestSaveFields() = NO Then
                    mPBEnableBox imPBBoxNo
                    Exit Sub
                End If
                Do
                    If imPBRowNo >= UBound(smPBSave, 2) - 1 Then
                        imPBBoxNo = -1
                        If lacTotals.BackColor = GREEN Then
                            cmcDone.Enabled = True
                            'cmcDone.SetFocus
                            Exit Sub
                        Else
                            cmcCancel.Enabled = True
                            'cmcCancel.SetFocus
                            Exit Sub
                        End If
                    End If
                    imPBRowNo = imPBRowNo + 1
                    If imPBRowNo > vbcMP.Value + vbcMP.LargeChange Then
                        imSettingValue = True
                        vbcMP.Value = vbcMP.Value + 1
                        imSettingValue = False
                    End If
                    slDate = smPBSave(3, imPBRowNo) 'lbcBDate.List(imPBSave(2, imPBRowNo))
                    If gDateValue(slDate) <= lmLastBilledDate Then
                        ilFound = False
                    Else
                        ilFound = True
                        ilBox = MPBAMOUNTINDEX
                        imPBBoxNo = ilBox
                        mPBEnableBox ilBox
                    End If
                Loop While Not ilFound
                Exit Sub
            Case Else
                ilBox = imPBBoxNo + 1
        End Select
        mPBSetShow imPBBoxNo
        imPBBoxNo = ilBox
        mPBEnableBox ilBox
    End If
End Sub
Private Sub pbcFix_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer

    For ilBox = imLBCtrls To UBound(tmFBCtrls) Step 1
        If (X >= tmFBCtrls(ilBox).fBoxX) And (X <= (tmFBCtrls(ilBox).fBoxX + tmFBCtrls(ilBox).fBoxW)) Then
            If (Y >= 30) And (Y <= tmFBCtrls(ilBox).fBoxY - 30) Then
                mFSSetShow imFSBoxNo
                imFSBoxNo = -1
                mFBSetShow imFBBoxNo
                imFBBoxNo = -1
                If ilBox = FBVEHICLEINDEX Then
                    imFixSort = 1
                    mFBSort
                    Exit Sub
                End If
                If ilBox = FBDATEINDEX Then
                    imFixSort = 0
                    mFBSort
                    Exit Sub
                End If
            End If
        End If
    Next ilBox
    ilCompRow = vbcFix.LargeChange + 1
    'If UBound(smFBSave, 2) > ilCompRow Then
    If UBound(tmInstallBillInfo) + 1 > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tmInstallBillInfo) + 1    'UBound(smFBSave, 2)
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmFBCtrls) Step 1
            If (X >= tmFBCtrls(ilBox).fBoxX) And (X <= (tmFBCtrls(ilBox).fBoxX + tmFBCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmFBCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmFBCtrls(ilBox).fBoxY + tmFBCtrls(ilBox).fBoxH)) Then
                    If ilBox <> FBBILLINGINDEX Then
                        Beep
                        Exit Sub
                    End If
                    ilRowNo = ilRow + vbcFix.Value - 1
                    mFSSetShow imFSBoxNo
                    imFSBoxNo = -1
                    mFBSetShow imFBBoxNo
                    'If smFBSave(3, ilRowNo) = "B" Then
                    If tmInstallBillInfo(ilRowNo - 1).sBilledFlag = "Y" Then
                        Beep
                        Exit Sub
                    End If
                    imFBRowNo = ilRowNo
                    imFBBoxNo = ilBox
                    mFBEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcFix_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim llColor As Long
    Dim llRevenue As Long
    Dim llBill As Long
    Dim slStr As String
    Dim ilTest As Integer
    Dim ilVef As Integer

    mPaintFixTitle
    ilStartRow = vbcFix.Value 'Top location
    ilEndRow = vbcFix.Value + vbcFix.LargeChange
    'If ilEndRow > UBound(smFBSave, 2) Then
    If ilEndRow > UBound(tmInstallBillInfo) Then
        ilEndRow = UBound(tmInstallBillInfo)    'UBound(smFBSave, 2) 'include blank row as it might have data
    End If
    llColor = pbcFix.ForeColor
    llRevenue = 0
    llBill = 0
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBCtrls To UBound(tmFBCtrls) Step 1
            'If (ilBox = FBVEHICLEINDEX) Or (ilBox = FBDATEINDEX) Or (ilBox = FBBILLINDEX) Then
            If (ilBox = FBVEHICLEINDEX) Or (ilBox = FBDATEINDEX) Then
                'gPaintArea pbcFix, tmFBCtrls(ilBox).fBoxX, tmFBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmFBCtrls(ilBox).fBoxW - 15, tmFBCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
            Else
                'gPaintArea pbcFix, tmFBCtrls(ilBox).fBoxX, tmFBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmFBCtrls(ilBox).fBoxW - 15, tmFBCtrls(ilBox).fBoxH - 15, WHITE
            End If
            'If ilRow = UBound(smFBSave, 2) Then
            '    pbcFix.ForeColor = DARKPURPLE
            'Else
            '    pbcFix.ForeColor = llColor
            'End If
            'If (smFBSave(3, ilRow) = "B") Then
            If (tmInstallBillInfo(ilRow - 1).sBilledFlag = "Y") Then
                pbcFix.ForeColor = DARKGREEN
            End If
            pbcFix.CurrentX = tmFBCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcFix.CurrentY = tmFBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            Select Case ilBox
                Case FBVEHICLEINDEX
                    slStr = Trim$(tmInstallBillInfo(ilRow - 1).sVehName)
                    '12/18/17: Break out NTR separate from Air Time
                    If bgBreakoutNTR And (tmInstallBillInfo(ilRow - 1).iMnfItem > 0) Then
                        slStr = slStr + "(" + Trim$(tmInstallBillInfo(ilRow - 1).sMnfItem) + ")"
                    End If
                Case FBDATEINDEX
                    slStr = Format$(tmInstallBillInfo(ilRow - 1).lBillDate, "m/d/yy")
                Case FBORDEREDINDEX
                    If (tmInstallBillInfo(ilRow - 1).lAirOrderedDollars > 0) Or (tmInstallBillInfo(ilRow - 1).lNTROrderedDollars > 0) Then
                        slStr = gLongToStrDec(tmInstallBillInfo(ilRow - 1).lAirOrderedDollars + tmInstallBillInfo(ilRow - 1).lNTROrderedDollars, 2)
                    Else
                        slStr = ""
                    End If
                Case FBREVENUETOTALINDEX
                    If ilRow = ilStartRow Then
                        If imFixSort = 0 Then   'Date
                            For ilTest = ilRow To LBound(tmInstallBillInfo) + 1 Step -1
                                If tmInstallBillInfo(ilTest - 1).lBillDate = tmInstallBillInfo(ilRow - 1).lBillDate Then
                                    llRevenue = llRevenue + tmInstallBillInfo(ilTest - 1).lAirOrderedDollars + tmInstallBillInfo(ilTest - 1).lNTROrderedDollars
                                Else
                                    Exit For
                                End If
                            Next ilTest
                        Else
                            For ilTest = ilRow To LBound(tmInstallBillInfo) + 1 Step -1
                                If tmInstallBillInfo(ilTest - 1).iVefCode = tmInstallBillInfo(ilRow - 1).iVefCode Then
                                    llRevenue = llRevenue + tmInstallBillInfo(ilTest - 1).lAirOrderedDollars + tmInstallBillInfo(ilTest - 1).lNTROrderedDollars
                                Else
                                    Exit For
                                End If
                            Next ilTest
                        End If
                    ElseIf ilRow = UBound(tmInstallBillInfo) Then
                        If imFixSort = 0 Then   'Date
                            For ilTest = ilRow To UBound(tmInstallBillInfo) Step 1
                                If tmInstallBillInfo(ilTest - 1).lBillDate = tmInstallBillInfo(ilRow - 1).lBillDate Then
                                    llRevenue = llRevenue + tmInstallBillInfo(ilTest - 1).lAirOrderedDollars + tmInstallBillInfo(ilTest - 1).lNTROrderedDollars
                                Else
                                    Exit For
                                End If
                            Next ilTest
                        Else
                            For ilTest = ilRow To UBound(tmInstallBillInfo) Step 1
                                If tmInstallBillInfo(ilTest - 1).iVefCode = tmInstallBillInfo(ilRow - 1).iVefCode Then
                                    llRevenue = llRevenue + tmInstallBillInfo(ilTest - 1).lAirOrderedDollars + tmInstallBillInfo(ilTest - 1).lNTROrderedDollars
                                Else
                                    Exit For
                                End If
                            Next ilTest
                        End If
                    Else
                        llRevenue = llRevenue + tmInstallBillInfo(ilRow - 1).lAirOrderedDollars + tmInstallBillInfo(ilRow - 1).lNTROrderedDollars
                    End If
                    slStr = ""
                    If ilRow = UBound(tmInstallBillInfo) + 1 Then
                        slStr = gLongToStrDec(llRevenue, 2)
                        llRevenue = 0
                    Else
                        If imFixSort = 0 Then   'Date
                            If tmInstallBillInfo(ilRow).lBillDate <> tmInstallBillInfo(ilRow - 1).lBillDate Then
                                slStr = gLongToStrDec(llRevenue, 2)
                                llRevenue = 0
                            End If
                        Else
                            If tmInstallBillInfo(ilRow).iVefCode <> tmInstallBillInfo(ilRow - 1).iVefCode Then
                                slStr = gLongToStrDec(llRevenue, 2)
                                llRevenue = 0
                            End If
                        End If
                    End If
                Case FBBILLINGINDEX
                    'If tmInstallBillInfo(ilRow - 1).lBillDollars > 0 Then
                        For ilVef = 0 To UBound(tmInstallVehInfo) - 1 Step 1
                            If tmInstallVehInfo(ilVef).iVefCode = tmInstallBillInfo(ilRow - 1).iVefCode Then
                                '12/18/17: Break out NTR separate from Air Time
                                If (Not bgBreakoutNTR) Or (bgBreakoutNTR And (tmInstallVehInfo(ilVef).iMnfItem = tmInstallBillInfo(ilRow - 1).iMnfItem)) Then
                                    If tmInstallVehInfo(ilVef).lOrderedDollars <> tmInstallVehInfo(ilVef).lTotalBillDollars Then
                                        If tmInstallBillInfo(ilRow - 1).sBilledFlag <> "Y" Then
                                            pbcFix.ForeColor = vbRed
                                        End If
                                        Exit For
                                    End If
                                End If
                            End If
                        Next ilVef
                        slStr = gLongToStrDec(tmInstallBillInfo(ilRow - 1).lBillDollars, 2)
                        If Left(slStr, 1) = "-" Then
                            slStr = Mid$(slStr, 2) + "-"
                        End If
                    'Else
                    '    slStr = ""
                    'End If
                Case FBBILLTOTALINDEX
                    If ilRow = ilStartRow Then
                        If imFixSort = 0 Then   'Date
                            For ilTest = ilRow To LBound(tmInstallBillInfo) + 1 Step -1
                                If tmInstallBillInfo(ilTest - 1).lBillDate = tmInstallBillInfo(ilRow - 1).lBillDate Then
                                    llBill = llBill + tmInstallBillInfo(ilTest - 1).lBillDollars
                                Else
                                    Exit For
                                End If
                            Next ilTest
                        Else
                            For ilTest = ilRow To LBound(tmInstallBillInfo) + 1 Step -1
                                If tmInstallBillInfo(ilTest - 1).iVefCode = tmInstallBillInfo(ilRow - 1).iVefCode Then
                                    llBill = llBill + tmInstallBillInfo(ilTest - 1).lBillDollars
                                Else
                                    Exit For
                                End If
                            Next ilTest
                        End If
                    ElseIf ilRow = UBound(tmInstallBillInfo) Then
                        If imFixSort = 0 Then   'Date
                            For ilTest = ilRow To UBound(tmInstallBillInfo) Step 1
                                If tmInstallBillInfo(ilTest - 1).lBillDate = tmInstallBillInfo(ilRow - 1).lBillDate Then
                                    llBill = llBill + tmInstallBillInfo(ilTest - 1).lBillDollars
                                Else
                                    Exit For
                                End If
                            Next ilTest
                        Else
                            For ilTest = ilRow To UBound(tmInstallBillInfo) Step 1
                                If tmInstallBillInfo(ilTest - 1).iVefCode = tmInstallBillInfo(ilRow - 1).iVefCode Then
                                    llBill = llBill + tmInstallBillInfo(ilTest - 1).lBillDollars
                                Else
                                    Exit For
                                End If
                            Next ilTest
                        End If
                    Else
                        llBill = llBill + tmInstallBillInfo(ilRow - 1).lBillDollars
                    End If
                    slStr = ""
                    If ilRow = UBound(tmInstallBillInfo) + 1 Then
                        slStr = gLongToStrDec(llBill, 2)
                        llBill = 0
                    Else
                        If imFixSort = 0 Then   'Date
                            If tmInstallBillInfo(ilRow).lBillDate <> tmInstallBillInfo(ilRow - 1).lBillDate Then
                                slStr = gLongToStrDec(llBill, 2)
                                llBill = 0
                            End If
                        Else
                            If tmInstallBillInfo(ilRow).iVefCode <> tmInstallBillInfo(ilRow - 1).iVefCode Then
                                slStr = gLongToStrDec(llBill, 2)
                                llBill = 0
                            End If
                        End If
                    End If
                    If (imFixSort <> 0) And (slStr <> "") Then
                        For ilVef = 0 To UBound(tmInstallVehInfo) - 1 Step 1
                            If tmInstallVehInfo(ilVef).iVefCode = tmInstallBillInfo(ilRow - 1).iVefCode Then
                                If tmInstallVehInfo(ilVef).lOrderedDollars <> tmInstallVehInfo(ilVef).lTotalBillDollars Then
                                    If tmInstallBillInfo(ilRow - 1).sBilledFlag <> "Y" Then
                                        pbcFix.ForeColor = vbRed
                                    End If
                                    Exit For
                                End If
                            End If
                        Next ilVef
                    End If
                    If Left(slStr, 1) = "-" Then
                        slStr = Mid$(slStr, 2) + "-"
                    End If
            End Select
            gSetShow pbcFix, slStr, tmFBCtrls(ilBox)
            slStr = tmFBCtrls(ilBox).sShow
            pbcFix.Print slStr
            'If ((ilBox = FBBILLINGINDEX) Or (ilBox = FBBILLTOTALINDEX)) And (tmInstallBillInfo(ilRow - 1).sBilledFlag <> "Y") Then
                pbcFix.ForeColor = llColor
            'End If
        Next ilBox
    Next ilRow
    pbcFix.ForeColor = llColor
End Sub
Private Sub pbcFixSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBCtrls To UBound(tmFSCtrls) - 1 Step 1
        If (X >= tmFSCtrls(ilBox).fBoxX) And (X <= tmFSCtrls(ilBox).fBoxX + tmFSCtrls(ilBox).fBoxW) Then
            If (Y >= tmFSCtrls(ilBox).fBoxY) And (Y <= tmFSCtrls(ilBox).fBoxY + tmFSCtrls(ilBox).fBoxH) Then
                mFBSetShow imFBBoxNo
                imFBBoxNo = -1
                lacIBFrame.Visible = False
                pbcArrow.Visible = False
                mFSSetShow imFSBoxNo
                imFSBoxNo = ilBox
                mFSEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcFixSpec_Paint()
    Dim ilBox As Integer
    mPaintFixSpecTitle
    For ilBox = imLBCtrls To UBound(tmFSCtrls) Step 1
        pbcFixSpec.CurrentX = tmFSCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcFixSpec.CurrentY = tmFSCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcFixSpec.Print tmFSCtrls(ilBox).sShow
    Next ilBox
End Sub


Private Sub pbcFSSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcFSSTab.HWnd Then
        Exit Sub
    End If
    mFBSetShow imFBBoxNo
    imFBBoxNo = -1
    imFBRowNo = -1
    mMBSetShow imMBBoxNo
    imMBBoxNo = -1
    imMBRowNo = -1
    mPBSetShow imPBBoxNo
    imPBBoxNo = -1
    imPBRowNo = -1
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    If rbcOption(0).Value Then
        Select Case imFSBoxNo
            Case -1 'Tab from control prior to form area
                ilBox = FSSTARTDATEINDEX
            Case FSSTARTDATEINDEX 'Type (first control within header)
                slStr = edcPSDropDown.Text
                If slStr <> "" Then
                    slStr = edcPSDropDown.Text
                    'slStr = Left$(slStr, 3) & " 01" & Mid(slStr, 4)
                    'slStr = Format(slStr, "m/d/yy")
                    If Not gValidDate(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcPSDropDown.SetFocus
                    Exit Sub
                End If
                mFSSetShow imFSBoxNo
                imFSBoxNo = -1
                If cmcGen.Enabled Then
                    cmcGen.SetFocus
                    Exit Sub
                End If
                'cmcDone.SetFocus
                Exit Sub
            Case FSENDDATEINDEX
                slStr = edcPSDropDown.Text
                If slStr <> "" Then
                    slStr = edcPSDropDown.Text
                    'slStr = Left$(slStr, 3) & " 01" & Mid(slStr, 4)
                    'slStr = Format(slStr, "m/d/yy")
                    If Not gValidDate(slStr) Then
                        Beep
                        edcDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcPSDropDown.SetFocus
                    Exit Sub
                End If
                ilBox = imFSBoxNo - 1
            Case Else
                ilBox = imFSBoxNo - 1
        End Select
        mFSSetShow imFSBoxNo
        imFSBoxNo = ilBox
        mFSEnableBox ilBox
    ElseIf rbcOption(2).Value Then
        Select Case imMSBoxNo
            Case -1 'Tab from control prior to form area
                ilBox = MPSPERCENTINDEX
            Case MPSPERCENTINDEX 'Type (first control within header)
                mMSSetShow imMSBoxNo
                imMSBoxNo = -1
                If cmcGen.Enabled Then
                    cmcGen.SetFocus
                    Exit Sub
                End If
                'cmcDone.SetFocus
                Exit Sub
            Case Else
                ilBox = imMSBoxNo - 1
        End Select
        mMSSetShow imMSBoxNo
        imMSBoxNo = ilBox
        mMSEnableBox ilBox
    ElseIf rbcOption(3).Value Then
        Select Case imPSBoxNo
            Case -1 'Tab from control prior to form area
                ilBox = MPSPERCENTINDEX
            Case MPSPERCENTINDEX 'Type (first control within header)
                mPSSetShow imPSBoxNo
                imPSBoxNo = -1
                If cmcGen.Enabled Then
                    cmcGen.SetFocus
                    Exit Sub
                End If
                'cmcDone.SetFocus
                Exit Sub
            Case Else
                ilBox = imPSBoxNo - 1
        End Select
        mPSSetShow imPSBoxNo
        imPSBoxNo = ilBox
        mPSEnableBox ilBox
    End If
End Sub
Private Sub pbcFSTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPos                                                                                 *
'******************************************************************************************

    Dim ilBox As Integer
    Dim slStr As String

    If GetFocus() <> pbcFSTab.HWnd Then
        Exit Sub
    End If
    mFBSetShow imFBBoxNo
    imFBBoxNo = -1
    imFBRowNo = -1
    mMBSetShow imMBBoxNo
    imMBBoxNo = -1
    imMBRowNo = -1
    mPBSetShow imPBBoxNo
    imPBBoxNo = -1
    imPBRowNo = -1
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    If rbcOption(0).Value Then
        Select Case imFSBoxNo
            Case -1 'Shift tab from button
                ilBox = FSNOMONTHSOFFINDEX   'FSAMOUNTINDEX
            Case FSSTARTDATEINDEX
                slStr = edcPSDropDown.Text
                If slStr <> "" Then
                    'slStr = Left$(slStr, 3) & " 01" & Mid(slStr, 4)
                    'slStr = Format(slStr, "m/d/yy")
                    If Not gValidDate(slStr) Then
                        Beep
                        edcPSDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcPSDropDown.SetFocus
                    Exit Sub
                End If
                ilBox = imFSBoxNo + 1
            Case FSENDDATEINDEX
                slStr = edcPSDropDown.Text
                If slStr <> "" Then
                    'slStr = Left$(slStr, 3) & " 01" & Mid(slStr, 4)
                    'slStr = Format(slStr, "m/d/yy")
                    If Not gValidDate(slStr) Then
                        Beep
                        edcPSDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcPSDropDown.SetFocus
                    Exit Sub
                End If
                ilBox = imFSBoxNo + 1
            Case FSNOMONTHSOFFINDEX 'Type (first control within header)
                mFSSetShow imFSBoxNo
                imFSBoxNo = -1
                If cmcGen.Enabled Then
                    cmcGen.SetFocus
                    Exit Sub
                End If
                'cmcDone.SetFocus
                Exit Sub
            Case Else
                ilBox = imFSBoxNo + 1
        End Select
        mFSSetShow imFSBoxNo
        imFSBoxNo = ilBox
        mFSEnableBox ilBox
    ElseIf rbcOption(2).Value Then
        Select Case imMSBoxNo
            Case -1 'Shift tab from button
                ilBox = MPSENDDATEINDEX
            Case MPSSTARTDATEINDEX
                slStr = edcPSDropDown.Text
                If slStr <> "" Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcPSDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcPSDropDown.SetFocus
                    Exit Sub
                End If
                ilBox = imMSBoxNo + 1
            Case MPSENDDATEINDEX 'Type (first control within header)
                slStr = edcPSDropDown.Text
                If slStr <> "" Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcPSDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcPSDropDown.SetFocus
                    Exit Sub
                End If
                mMSSetShow imMSBoxNo
                imMSBoxNo = -1
                If cmcGen.Enabled Then
                    cmcGen.SetFocus
                    Exit Sub
                End If
                'cmcDone.SetFocus
                Exit Sub
            Case Else
                ilBox = imMSBoxNo + 1
        End Select
        mMSSetShow imMSBoxNo
        imMSBoxNo = ilBox
        mMSEnableBox ilBox
    ElseIf rbcOption(3).Value Then
        Select Case imPSBoxNo
            Case -1 'Shift tab from button
                ilBox = MPSENDDATEINDEX
            Case MPSSTARTDATEINDEX
                slStr = edcPSDropDown.Text
                If slStr <> "" Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcPSDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcPSDropDown.SetFocus
                    Exit Sub
                End If
                ilBox = imPSBoxNo + 1
            Case MPSENDDATEINDEX 'Type (first control within header)
                slStr = edcPSDropDown.Text
                If slStr <> "" Then
                    If Not gValidDate(slStr) Then
                        Beep
                        edcPSDropDown.SetFocus
                        Exit Sub
                    End If
                Else
                    Beep
                    edcPSDropDown.SetFocus
                    Exit Sub
                End If
                mPSSetShow imPSBoxNo
                imPSBoxNo = -1
                If cmcGen.Enabled Then
                    cmcGen.SetFocus
                    Exit Sub
                End If
                'cmcDone.SetFocus
                Exit Sub
            Case Else
                ilBox = imPSBoxNo + 1
        End Select
        mPSSetShow imPSBoxNo
        imPSBoxNo = ilBox
        mPSEnableBox ilBox
    End If
End Sub
Private Sub pbcIBSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slTax As String
    Dim ilNext As Integer

    If Not imUpdateAllowed Then
        'cmcDone.SetFocus
        Exit Sub
    End If
    If GetFocus() <> pbcIBSTab.HWnd Then
        Exit Sub
    End If
    If imIBBoxNo = IBITEMTYPEINDEX Then
        If mItemTypeBranch() Then
            Exit Sub
        End If
    End If
    imTabDirection = -1  'Set-right to left
    ilBox = imIBBoxNo
    Do
        ilNext = False
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imTabDirection = 0  'Set-Left to right
                imIBRowNo = 1
                imSettingValue = True
                vbcItemBill.Value = vbcItemBill.Min
                imSettingValue = False
                If UBound(smIBSave, 2) <= vbcItemBill.LargeChange + 1 Then 'was <=
                    vbcItemBill.Max = LBONE 'LBound(smIBSave, 2)
                Else
                    vbcItemBill.Max = UBound(smIBSave, 2) - vbcItemBill.LargeChange ' - 1
                End If
    '            Do While (imIBRowNo < UBound(smIBSave, 2)) And ((smIBSave(1, imIBRowNo) = "I") Or (smIBSave(7, imIBRowNo) = "R") Or (smIBSave(7, imIBRowNo) = "B"))
    '                imIBRowNo = imIBRowNo + 1
    '                If imIBRowNo > vbcItemBill.Value + vbcItemBill.LargeChange Then
    '                    imSettingValue = True
    '                    vbcItemBill.Value = vbcItemBill.Value + 1
    '                End If
    '            Loop
                If (imIBRowNo = UBound(smIBSave, 2)) And (imIBSave(1, 1) = -1) Then
                    mInitNewIB imIBRowNo
                Else
                    imIBRowNo = 0
                    Do
                        imIBRowNo = imIBRowNo + 1
                        If imIBRowNo > vbcItemBill.Value + vbcItemBill.LargeChange Then
                            imSettingValue = True
                            vbcItemBill.Value = vbcItemBill.Value + 1
                            imSettingValue = False
                        End If
        '            Loop While (smIBSave(1, imIBRowNo) = "I") Or (smIBSave(7, imIBRowNo) = "R") Or (smIBSave(7, imIBRowNo) = "B")
                    Loop While (smIBSave(7, imIBRowNo) = "Y")
                End If
                ilBox = 1
                If mIBColOk(imIBRowNo, ilBox) Then
                    imIBBoxNo = ilBox
                    mIBEnableBox ilBox
                    Exit Sub
                End If
            Case IBVEHICLEINDEX 'Name (first control within header)
                mIBSetShow imIBBoxNo
                ilBox = IBNOITEMSINDEX
                Do
                    If imIBRowNo <= 1 Then
                        imIBBoxNo = -1
                        'cmcDone.SetFocus
                        Exit Sub
                    End If
                    imIBRowNo = imIBRowNo - 1
                    If imIBRowNo < vbcItemBill.Value Then
                        imSettingValue = True
                        vbcItemBill.Value = vbcItemBill.Value - 1
                        imSettingValue = False
                    End If
    '            Loop While (smIBSave(1, imIBRowNo) = "I") Or (smIBSave(7, imIBRowNo) = "R") Or (smIBSave(7, imIBRowNo) = "B")
                Loop While (smIBSave(7, imIBRowNo) = "Y")
                If mIBColOk(imIBRowNo, ilBox) Then
                    imIBBoxNo = ilBox
                    mIBEnableBox ilBox
                    Exit Sub
                End If
            Case IBNOITEMSINDEX
                ilBox = IBAMOUNTINDEX
            Case IBAMOUNTINDEX
                If Not imTaxDefined Then
                    'If imIBSave(5, imIBRowNo) = 2 Then
                    '    ilBox = IBACINDEX
                    'Else
    '                    ilBox = IBSCINDEX
                        If tgSpf.sSubCompany = "Y" Then
                            ilBox = IBACINDEX
                        Else
                            ilBox = IBSCINDEX
                        End If
                    'End If
                Else
                    ilBox = IBTXINDEX
                    If imIBSave(3, imIBRowNo) > 0 Then
                        slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                        ilRet = gParseItem(slNameCode, 6, "\", slTax)
                        If ilRet = CP_MSG_NONE Then
                            If slTax <> "Y" Then
                                If tgSpf.sSubCompany = "Y" Then
                                    ilBox = IBACINDEX
                                Else
                                    ilBox = IBSCINDEX
                                End If
                            End If
                        End If
                    End If
                End If
            Case IBTXINDEX
                'If imIBSave(5, imIBRowNo) = 2 Then
                '    ilBox = IBACINDEX
                'Else
    '                ilBox = IBSCINDEX
                    If tgSpf.sSubCompany = "Y" Then
                        ilBox = IBACINDEX
                    Else
                        ilBox = IBSCINDEX
                    End If
                'End If
            Case IBSCINDEX
                'If imIBSave(4, imIBRowNo) = 2 Then  '2=direct
                If igDirAdvt Then
                    ilBox = IBITEMTYPEINDEX
                Else
                    ilBox = IBACINDEX
                End If
            Case IBDATEINDEX
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
                ilBox = IBVEHICLEINDEX
            Case Else
                ilBox = ilBox - 1
        End Select
        If mIBColOk(imIBRowNo, ilBox) Then
            Exit Do
        Else
            ilNext = True
        End If
    Loop While ilNext
    mIBSetShow imIBBoxNo
    imIBBoxNo = ilBox
    mIBEnableBox ilBox
End Sub
Private Sub pbcIBSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcIBTab_GotFocus()
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slTax As String
    Dim ilNext As Integer

    If Not imUpdateAllowed Then
        'cmcDone.SetFocus
        Exit Sub
    End If
    If GetFocus() <> pbcIBTab.HWnd Then
        Exit Sub
    End If
    If imIBBoxNo = IBITEMTYPEINDEX Then
        If mItemTypeBranch() Then
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    ilBox = imIBBoxNo
    Do
        ilNext = False
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imTabDirection = -1  'Set-Right to left
                imIBRowNo = UBound(smIBSave, 2)
                imSettingValue = True
                If imIBRowNo <= vbcItemBill.LargeChange + 1 Then
                    vbcItemBill.Value = vbcItemBill.Min
                Else
                    vbcItemBill.Value = imIBRowNo - vbcItemBill.LargeChange
                End If
                imSettingValue = False
                ilBox = 1
            Case IBVEHICLEINDEX
                If (imIBRowNo >= UBound(smIBSave, 2)) And (lbcBVehicle.ListIndex < 0) Then
                    mIBSetShow imIBBoxNo
                    For ilLoop = IBVEHICLEINDEX To IBACQCOSTINDEX Step 1
                        slStr = ""
                        gSetShow pbcItemBill, slStr, tmIBCtrls(ilLoop)
                        smIBShow(ilLoop, imIBRowNo) = tmIBCtrls(ilLoop).sShow
                    Next ilLoop
                    imIBSave(1, imIBRowNo) = -1
                    imIBBoxNo = -1
                    pbcItemBill_Paint
                    If cmcDone.Enabled Then
                        'cmcDone.SetFocus
                    Else
                        'cmcCancel.SetFocus
                    End If
                    Exit Sub
                End If
                ilBox = IBDATEINDEX
            Case IBDATEINDEX
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
                ilBox = ilBox + 1
            Case IBITEMTYPEINDEX
                'If imIBSave(4, imIBRowNo) = 2 Then  '2=direct
                If igDirAdvt Then
                    ilBox = IBSCINDEX
                Else
                    ilBox = IBACINDEX
                End If
            Case IBACINDEX
                'If imIBSave(5, imIBRowNo) = 2 Then  '2=no salesperson commission defined
                '    If imIBSave(6, imIBRowNo) = 2 Then  '2=no tax defined
                '        ilBox = IBAMOUNTINDEX
                '    Else
                '        ilBox = IBTXINDEX
                '    End If
                'Else
    '                ilBox = IBSCINDEX
                'End If
                If tgSpf.sSubCompany = "Y" Then
                    If Not imTaxDefined Then
                        ilBox = IBAMOUNTINDEX
                    Else
                        ilBox = IBTXINDEX
                        If imIBSave(3, imIBRowNo) > 0 Then
                            slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                            ilRet = gParseItem(slNameCode, 6, "\", slTax)
                            If ilRet = CP_MSG_NONE Then
                                If slTax <> "Y" Then
                                    ilBox = IBAMOUNTINDEX
                                End If
                            End If
                        End If
                    End If
                Else
                    ilBox = IBSCINDEX
                End If
            Case IBSCINDEX
                If Not imTaxDefined Then
                    ilBox = IBAMOUNTINDEX
                Else
                    ilBox = IBTXINDEX
                    If imIBSave(3, imIBRowNo) > 0 Then
                        slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                        ilRet = gParseItem(slNameCode, 6, "\", slTax)
                        If ilRet = CP_MSG_NONE Then
                            If slTax <> "Y" Then
                                ilBox = IBAMOUNTINDEX
                            End If
                        End If
                    End If
                End If
            Case IBAMOUNTINDEX
                ilBox = IBNOITEMSINDEX
            Case IBNOITEMSINDEX, IBACQCOSTINDEX 'Last control
                '6/7/15: replaced acquisition from site override with Barter in system options
                If ((ilBox = IBNOITEMSINDEX) And ((Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) <> SPNTRACQUISITION)) Or (ilBox = IBACQCOSTINDEX) Then
                    mIBSetShow imIBBoxNo
                    If mIBTestSaveFields() = NO Then
                        mIBEnableBox imIBBoxNo
                        Exit Sub
                    End If
                    If imIBRowNo >= UBound(smIBSave, 2) Then
                        imIBChg = True
                        'ReDim Preserve smIBSave(1 To 11, 1 To imIBRowNo + 1) As String
                        'ReDim Preserve imIBSave(1 To 7, 1 To imIBRowNo + 1) As Integer
                        'ReDim Preserve smIBShow(1 To IBACQCOSTINDEX, 1 To imIBRowNo + 1) As String
                        'ReDim Preserve lmIBSave(1 To 2, 1 To imIBRowNo + 1) As Long
                        ReDim Preserve smIBSave(0 To 11, 0 To imIBRowNo + 1) As String
                        ReDim Preserve imIBSave(0 To 7, 0 To imIBRowNo + 1) As Integer
                        ReDim Preserve smIBShow(0 To IBACQCOSTINDEX, 0 To imIBRowNo + 1) As String
                        ReDim Preserve lmIBSave(0 To 2, 0 To imIBRowNo + 1) As Long
                        mInitNewIB imIBRowNo + 1
                        If UBound(smIBSave, 2) <= vbcItemBill.LargeChange + 1 Then 'was <=
                            vbcItemBill.Max = LBONE 'LBound(smIBSave, 2)
                        Else
                            vbcItemBill.Max = UBound(smIBSave, 2) - vbcItemBill.LargeChange ' - 1
                        End If
                        mIBTotals False
                    End If
                    Do
                        If imIBRowNo + 1 > UBound(smIBSave, 2) Then
                            imIBBoxNo = -1
                            imIBRowNo = -1
                            pbcItemBill_Paint
                            If cmcDone.Enabled Then
                                'cmcDone.SetFocus
                            Else
                                'cmcCancel.SetFocus
                            End If
                            Exit Sub
                        End If
                        imIBRowNo = imIBRowNo + 1
                        If imIBRowNo > vbcItemBill.Value + vbcItemBill.LargeChange Then
                            imSettingValue = True
                            If vbcItemBill.Value + 1 > vbcItemBill.Max Then
                                vbcItemBill.Max = vbcItemBill.Value + 1
                            End If
                            vbcItemBill.Value = vbcItemBill.Value + 1
                            imSettingValue = False
                        End If
        '            Loop While (smIBSave(1, imIBRowNo) = "I") Or (smIBSave(7, imIBRowNo) = "R") Or (smIBSave(7, imIBRowNo) = "B")
                    Loop While (smIBSave(7, imIBRowNo) = "Y")
                    If imIBRowNo >= UBound(smIBSave, 2) Then
                        mSetCommands
                        imIBBoxNo = 0
                        lacIBFrame.Move 0, tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15) - 30
                        lacIBFrame.Visible = True
                        pbcArrow.Move plcItemBill.Left - pbcArrow.Width - 15, plcItemBill.Top + tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15) + 45
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                        Exit Sub
                    Else
                        ilBox = 1
                        If mIBColOk(imIBRowNo, ilBox) Then
                            imIBBoxNo = ilBox
                            mIBEnableBox ilBox
                            Exit Sub
                        End If
                    End If
                    'Exit Sub
                Else
                    ilBox = IBACQCOSTINDEX
                End If
            Case Else
                ilBox = ilBox + 1
        End Select
        If mIBColOk(imIBRowNo, ilBox) Then
            Exit Do
        Else
            ilNext = True
        End If
    Loop While ilNext
    mIBSetShow imIBBoxNo
    imIBBoxNo = ilBox
    mIBEnableBox ilBox
End Sub
Private Sub pbcIBTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcItemBill_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcItemBill_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slTax As String

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    ilCompRow = vbcItemBill.LargeChange + 1
    If UBound(smIBSave, 2) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smIBSave, 2)
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmIBCtrls) Step 1
            If (X >= tmIBCtrls(ilBox).fBoxX) And (X <= (tmIBCtrls(ilBox).fBoxX + tmIBCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmIBCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmIBCtrls(ilBox).fBoxY + tmIBCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcItemBill.Value - 1
                    mFSSetShow imFBBoxNo
                    imFBBoxNo = -1
                    mFBSetShow imFBBoxNo
                    imFBBoxNo = -1
                    mIBSetShow imIBBoxNo
'                    If (smIBSave(1, ilRowNo) = "I") Or (smIBSave(7, ilRowNo) = "R") Or (smIBSave(7, ilRowNo) = "B") Then
'                        Beep
'                        Exit Sub
'                    End If
                    If ilBox = IBUNITSINDEX Then
                        Beep
'                        Exit Sub
                    End If
                    If ilBox = IBTAMOUNTINDEX Then
                        Beep
                        Exit Sub
                    End If
                    If smIBSave(7, ilRowNo) = "Y" Then
                        Beep
                        Exit Sub
                    End If
                    'If (ilBox = IBACINDEX) And (imIBSave(4, ilRowNo) = 2) Then
                    If (ilBox = IBACINDEX) And (igDirAdvt) Then
                        Beep
                        Exit Sub
                    End If
                    If (ilBox = IBSCINDEX) And (tgSpf.sSubCompany = "Y") Then
                        Beep
                        Exit Sub
                    End If
                    If (ilBox = IBTXINDEX) Then
                        If Not imTaxDefined Then
                            Beep
                            Exit Sub
                        End If
                        If imIBSave(3, ilRowNo) > 0 Then
                            slNameCode = tmItemCode(imIBSave(3, ilRowNo) - 1).sKey   'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                            ilRet = gParseItem(slNameCode, 6, "\", slTax)
                            If ilRet = CP_MSG_NONE Then
                                If slTax <> "Y" Then
                                    Beep
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                    '6/7/15: replaced acquisition from site override with Barter in system options
                    If (Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) <> SPNTRACQUISITION Then
                        If ilBox = IBACQCOSTINDEX Then
                            Beep
                            Exit Sub
                        End If
                    End If
                    If imIBSave(7, ilRowNo) > 0 Then
                        If smIBSave(11, ilRowNo) = "Y" Then
                            If (ilBox = IBVEHICLEINDEX) Or (ilBox = IBITEMTYPEINDEX) Or (ilBox = IBAMOUNTINDEX) Or (ilBox = IBNOITEMSINDEX) Or (ilBox = IBUNITSINDEX) Or (ilBox = IBTAMOUNTINDEX) Or (ilBox = IBACQCOSTINDEX) Then
                                Beep
                                Exit Sub
                            End If
                        Else
                            If (ilBox = IBVEHICLEINDEX) Or (ilBox = IBDATEINDEX) Or (ilBox = IBITEMTYPEINDEX) Or (ilBox = IBAMOUNTINDEX) Or (ilBox = IBNOITEMSINDEX) Or (ilBox = IBUNITSINDEX) Or (ilBox = IBTAMOUNTINDEX) Or (ilBox = IBACQCOSTINDEX) Then
                                Beep
                                Exit Sub
                            End If
                        End If
                    End If
                    imIBRowNo = ilRowNo
                    imIBBoxNo = ilBox
                    mIBEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcItemBill_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim llColor As Long
    Dim slStr As String

    mPaintItemBillTitle
    ilStartRow = vbcItemBill.Value  'Top location
    ilEndRow = vbcItemBill.Value + vbcItemBill.LargeChange
    If ilEndRow > UBound(smIBSave, 2) Then
        ilEndRow = UBound(smIBSave, 2) 'include blank row as it might have data
    End If
    llColor = pbcItemBill.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBCtrls To UBound(tmIBCtrls) Step 1
            If imIBRowNo = ilRow Then
                If (ilBox = IBUNITSINDEX) Or (ilBox = IBTAMOUNTINDEX) Then
                    gPaintArea pbcItemBill, tmIBCtrls(ilBox).fBoxX, tmIBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmIBCtrls(ilBox).fBoxW - 15, tmIBCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                Else
                    gPaintArea pbcItemBill, tmIBCtrls(ilBox).fBoxX, tmIBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmIBCtrls(ilBox).fBoxW - 15, tmIBCtrls(ilBox).fBoxH - 15, WHITE
                End If
            End If
            If imIBSave(7, ilRow) > 0 Then
                If smIBSave(11, ilRow) = "Y" Then
                    If (ilBox = IBVEHICLEINDEX) Or (ilBox = IBITEMTYPEINDEX) Or (ilBox = IBAMOUNTINDEX) Or (ilBox = IBNOITEMSINDEX) Or (ilBox = IBUNITSINDEX) Or (ilBox = IBTAMOUNTINDEX) Or (ilBox = IBACQCOSTINDEX) Then
                        gPaintArea pbcItemBill, tmIBCtrls(ilBox).fBoxX, tmIBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmIBCtrls(ilBox).fBoxW - 15, tmIBCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                    Else
                        gPaintArea pbcItemBill, tmIBCtrls(ilBox).fBoxX, tmIBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmIBCtrls(ilBox).fBoxW - 15, tmIBCtrls(ilBox).fBoxH - 15, WHITE
                    End If
                Else
                    If (ilBox = IBVEHICLEINDEX) Or (ilBox = IBDATEINDEX) Or (ilBox = IBITEMTYPEINDEX) Or (ilBox = IBAMOUNTINDEX) Or (ilBox = IBNOITEMSINDEX) Or (ilBox = IBUNITSINDEX) Or (ilBox = IBTAMOUNTINDEX) Or (ilBox = IBACQCOSTINDEX) Then
                        gPaintArea pbcItemBill, tmIBCtrls(ilBox).fBoxX, tmIBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmIBCtrls(ilBox).fBoxW - 15, tmIBCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                    Else
                        gPaintArea pbcItemBill, tmIBCtrls(ilBox).fBoxX, tmIBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmIBCtrls(ilBox).fBoxW - 15, tmIBCtrls(ilBox).fBoxH - 15, WHITE
                    End If
                End If
            End If
            If (ilBox = IBITEMTYPEINDEX) And (imIBSave(3, ilRow) >= 0) Then
                slStr = lbcBItem.List(imIBSave(3, ilRow))
                If InStr(1, slStr, "(Hard Cost)", vbTextCompare) > 0 Then
                    pbcItemBill.ForeColor = vbRed
                Else
                    If ilRow = UBound(smIBSave, 2) Then
                        pbcItemBill.ForeColor = DARKPURPLE
                    Else
                        pbcItemBill.ForeColor = llColor
                    End If
                End If
            Else
                If ilRow = UBound(smIBSave, 2) Then
                    pbcItemBill.ForeColor = DARKPURPLE
                Else
                    pbcItemBill.ForeColor = llColor
                End If
            End If
            If (ilBox = imLBCtrls) And (smIBSave(7, ilRow) = "Y") Then
              pbcItemBill.ForeColor = DARKGREEN
            End If
            pbcItemBill.CurrentX = tmIBCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcItemBill.CurrentY = tmIBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            pbcItemBill.Print smIBShow(ilBox, ilRow)
        Next ilBox
        pbcItemBill.ForeColor = llColor
    Next ilRow
End Sub

Private Sub pbcLbcBVehicle_Paint()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilField                       slFields                                                *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLinesEnd As Integer
    Dim llWidth As Long
    Dim llFgColor As Long

    If imINPBCPaint Then
        Exit Sub
    End If
    imINPBCPaint = True
    pbcLbcBVehicle.Move lbcBVehicle.Left + 15, lbcBVehicle.Top + 15, pbcLbcBVehicle.Width - 30, lbcBVehicle.height - 30 '2115, 1560
    ilLinesEnd = lbcBVehicle.TopIndex + lbcBVehicle.height \ fgListHtArial825
    If ilLinesEnd > lbcBVehicle.ListCount Then
        ilLinesEnd = lbcBVehicle.ListCount
    End If
    If lbcBVehicle.ListCount <= lbcBVehicle.height \ fgListHtArial825 Then
        llWidth = lbcBVehicle.Width - 30
    Else
        llWidth = lbcBVehicle.Width - igScrollBarWidth - 30
    End If
    pbcLbcBVehicle.Width = llWidth
    pbcLbcBVehicle.Cls
    llFgColor = pbcLbcBVehicle.ForeColor
    For ilLoop = lbcBVehicle.TopIndex To ilLinesEnd - 1 Step 1
        pbcLbcBVehicle.ForeColor = llFgColor
        If lbcBVehicle.Selected(ilLoop) Then
            gPaintArea pbcLbcBVehicle, CSng(0), CSng((ilLoop - lbcBVehicle.TopIndex) * fgListHtArial825), CSng(pbcLbcBVehicle.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
            pbcLbcBVehicle.ForeColor = vbWhite
        Else
            If ilLoop < imBPkgStartIndex Then
                pbcLbcBVehicle.ForeColor = DARKGREEN
            ElseIf (ilLoop >= imBPkgStartIndex) And (ilLoop < imBNonPkgStartIndex) Then
                pbcLbcBVehicle.ForeColor = vbBlue
            Else
                pbcLbcBVehicle.ForeColor = vbBlack
            End If
        End If
        pbcLbcBVehicle.CurrentX = 15
        pbcLbcBVehicle.CurrentY = (ilLoop - lbcBVehicle.TopIndex) * fgListHtArial825 + 15
        slStr = lbcBVehicle.List(ilLoop)
        gAdjShowLen pbcLbcBVehicle, slStr, pbcLbcBVehicle.Width
        pbcLbcBVehicle.Print slStr
        pbcLbcBVehicle.ForeColor = llFgColor
    Next ilLoop
    imINPBCPaint = False
End Sub

Private Sub pbcLbcPSVehicle_Paint()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilField                       slFields                                                *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLinesEnd As Integer
    Dim llWidth As Long
    Dim llFgColor As Long

    If imINPBCPaint Then
        Exit Sub
    End If
    imINPBCPaint = True
    pbcLbcPSVehicle.Move lbcPSVehicle.Left + 15, lbcPSVehicle.Top + 15, pbcLbcPSVehicle.Width - 30, lbcPSVehicle.height - 30 '2115, 1560
    ilLinesEnd = lbcPSVehicle.TopIndex + lbcPSVehicle.height \ fgListHtArial825
    If ilLinesEnd > lbcPSVehicle.ListCount Then
        ilLinesEnd = lbcPSVehicle.ListCount
    End If
    If lbcPSVehicle.ListCount <= lbcPSVehicle.height \ fgListHtArial825 Then
        llWidth = lbcPSVehicle.Width - 30
    Else
        llWidth = lbcPSVehicle.Width - igScrollBarWidth - 30
    End If
    pbcLbcPSVehicle.Width = llWidth
    pbcLbcPSVehicle.Cls
    llFgColor = pbcLbcPSVehicle.ForeColor
    For ilLoop = lbcPSVehicle.TopIndex To ilLinesEnd - 1 Step 1
        pbcLbcPSVehicle.ForeColor = llFgColor
        If lbcPSVehicle.Selected(ilLoop) Then
            gPaintArea pbcLbcPSVehicle, CSng(0), CSng((ilLoop - lbcPSVehicle.TopIndex) * fgListHtArial825), CSng(pbcLbcPSVehicle.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
            pbcLbcPSVehicle.ForeColor = vbWhite
        Else
            If ilLoop < imPSNonPkgStartIndex Then
                pbcLbcPSVehicle.ForeColor = vbBlue
            Else
                pbcLbcPSVehicle.ForeColor = vbBlack
            End If
        End If
        pbcLbcPSVehicle.CurrentX = 15
        pbcLbcPSVehicle.CurrentY = (ilLoop - lbcPSVehicle.TopIndex) * fgListHtArial825 + 15
        slStr = lbcPSVehicle.List(ilLoop)
        gAdjShowLen pbcLbcPSVehicle, slStr, pbcLbcPSVehicle.Width
        pbcLbcPSVehicle.Print slStr
        pbcLbcPSVehicle.ForeColor = llFgColor
    Next ilLoop
    imINPBCPaint = False
End Sub

Private Sub pbcMP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim slDate As String

    If rbcOption(2).Value Then
        ilCompRow = vbcMP.LargeChange + 1
        If UBound(smMBSave, 2) >= ilCompRow Then
            ilMaxRow = ilCompRow
        Else
            ilMaxRow = UBound(smMBSave, 2) - 1
        End If
        For ilRow = 1 To ilMaxRow Step 1
            For ilBox = imLBCtrls To UBound(tmMPBCtrls) Step 1
                If (X >= tmMPBCtrls(ilBox).fBoxX) And (X <= (tmMPBCtrls(ilBox).fBoxX + tmMPBCtrls(ilBox).fBoxW)) Then
                    If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmMPBCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmMPBCtrls(ilBox).fBoxY + tmMPBCtrls(ilBox).fBoxH)) Then
                        ilRowNo = ilRow + vbcMP.Value - 1
                        slDate = smMBSave(3, ilRowNo)   'lbcBDate.List(imMBSave(2, ilRowNo))
                        If gDateValue(slDate) <= lmLastBilledDate Then
                            Beep
                            Exit Sub
                        End If
                        mMSSetShow imMSBoxNo
                        imMSBoxNo = -1
                        mMBSetShow imMBBoxNo
                        imMBRowNo = ilRowNo
                        imMBBoxNo = ilBox
                        mMBEnableBox ilBox
                        Exit Sub
                    End If
                End If
            Next ilBox
        Next ilRow
    ElseIf rbcOption(3).Value Then
        ilCompRow = vbcMP.LargeChange + 1
        If UBound(smPBSave, 2) >= ilCompRow Then
            ilMaxRow = ilCompRow
        Else
            ilMaxRow = UBound(smPBSave, 2) - 1
        End If
        For ilRow = 1 To ilMaxRow Step 1
            For ilBox = imLBCtrls To UBound(tmMPBCtrls) Step 1
                If (X >= tmMPBCtrls(ilBox).fBoxX) And (X <= (tmMPBCtrls(ilBox).fBoxX + tmMPBCtrls(ilBox).fBoxW)) Then
                    If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmMPBCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmMPBCtrls(ilBox).fBoxY + tmMPBCtrls(ilBox).fBoxH)) Then
                        ilRowNo = ilRow + vbcMP.Value - 1
                        slDate = smPBSave(3, ilRowNo)   'lbcBDate.List(imPBSave(2, ilRowNo))
                        If gDateValue(slDate) <= lmLastBilledDate Then
                            Beep
                            Exit Sub
                        End If
                        mPSSetShow imPSBoxNo
                        imPSBoxNo = -1
                        mPBSetShow imPBBoxNo
                        imPBRowNo = ilRowNo
                        imPBBoxNo = ilBox
                        mPBEnableBox ilBox
                        Exit Sub
                    End If
                End If
            Next ilBox
        Next ilRow
    End If
End Sub
Private Sub pbcMP_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slDate As String

    mPaintMPTitle
    If rbcOption(2).Value Then
        ilStartRow = vbcMP.Value 'Top location
        ilEndRow = vbcMP.Value + vbcMP.LargeChange
        If ilEndRow >= UBound(smMBSave, 2) Then
            ilEndRow = UBound(smMBSave, 2) - 1
        End If
        For ilRow = ilStartRow To ilEndRow Step 1
            For ilBox = imLBCtrls To UBound(tmMPBCtrls) Step 1
                If ilBox = MPBAMOUNTINDEX Then
                    slDate = smMBSave(3, ilRow) 'lbcBDate.List(imMBSave(2, ilRow))
                    If gDateValue(slDate) <= lmLastBilledDate Then
                        gPaintArea pbcMP, tmMPBCtrls(ilBox).fBoxX, tmMPBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmMPBCtrls(ilBox).fBoxW - 15, tmMPBCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                    Else
                        gPaintArea pbcMP, tmMPBCtrls(ilBox).fBoxX, tmMPBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmMPBCtrls(ilBox).fBoxW - 15, tmMPBCtrls(ilBox).fBoxH - 15, WHITE
                    End If
                End If
                pbcMP.CurrentX = tmMPBCtrls(ilBox).fBoxX + fgBoxInsetX
                pbcMP.CurrentY = tmMPBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                pbcMP.Print smMBShow(ilBox, ilRow)
            Next ilBox
        Next ilRow
    ElseIf rbcOption(3).Value Then
        ilStartRow = vbcMP.Value 'Top location
        ilEndRow = vbcMP.Value + vbcMP.LargeChange
        If ilEndRow >= UBound(smPBSave, 2) Then
            ilEndRow = UBound(smPBSave, 2) - 1
        End If
        For ilRow = ilStartRow To ilEndRow Step 1
            For ilBox = imLBCtrls To UBound(tmMPBCtrls) Step 1
                If ilBox = MPBAMOUNTINDEX Then
                    slDate = smPBSave(3, ilRow) 'lbcBDate.List(imPBSave(ilRow, 2))
                    If gDateValue(slDate) <= lmLastBilledDate Then
                        gPaintArea pbcMP, tmMPBCtrls(ilBox).fBoxX, tmMPBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmMPBCtrls(ilBox).fBoxW - 15, tmMPBCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                    Else
                        gPaintArea pbcMP, tmMPBCtrls(ilBox).fBoxX, tmMPBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmMPBCtrls(ilBox).fBoxW - 15, tmMPBCtrls(ilBox).fBoxH - 15, WHITE
                    End If
                End If
                pbcMP.CurrentX = tmMPBCtrls(ilBox).fBoxX + fgBoxInsetX
                pbcMP.CurrentY = tmMPBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                pbcMP.Print smPBShow(ilBox, ilRow)
            Next ilBox
        Next ilRow
   End If
End Sub
Private Sub pbcMPSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    If rbcOption(2).Value Then
        For ilBox = imLBCtrls To UBound(tmMSCtrls) Step 1
            If (X >= tmMSCtrls(ilBox).fBoxX) And (X <= tmMSCtrls(ilBox).fBoxX + tmMSCtrls(ilBox).fBoxW) Then
                If (Y >= tmMSCtrls(ilBox).fBoxY) And (Y <= tmMSCtrls(ilBox).fBoxY + tmMSCtrls(ilBox).fBoxH) Then
                    mMBSetShow imMBBoxNo
                    imMBBoxNo = -1
                    imMBRowNo = -1
                    lacIBFrame.Visible = False
                    pbcArrow.Visible = False
                    mMSSetShow imMSBoxNo
                    imMSBoxNo = ilBox
                    mMSEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    ElseIf rbcOption(3).Value Then
        For ilBox = imLBCtrls To UBound(tmPSCtrls) Step 1
            If (X >= tmPSCtrls(ilBox).fBoxX) And (X <= tmPSCtrls(ilBox).fBoxX + tmPSCtrls(ilBox).fBoxW) Then
                If (Y >= tmPSCtrls(ilBox).fBoxY) And (Y <= tmPSCtrls(ilBox).fBoxY + tmPSCtrls(ilBox).fBoxH) Then
                    mPBSetShow imPBBoxNo
                    imPBBoxNo = -1
                    imPBRowNo = -1
                    lacIBFrame.Visible = False
                    pbcArrow.Visible = False
                    mPSSetShow imPSBoxNo
                    imPSBoxNo = ilBox
                    mPSEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
   End If
End Sub
Private Sub pbcMPSpec_Paint()
    Dim ilBox As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    llColor = pbcMPSpec.ForeColor
    slFontName = pbcMPSpec.FontName
    flFontSize = pbcMPSpec.FontSize
    pbcMPSpec.ForeColor = BLUE
    pbcMPSpec.FontBold = False
    pbcMPSpec.FontSize = 7
    pbcMPSpec.FontName = "Arial"
    pbcMPSpec.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    If (Asc(tgSpf.sUsingFeatures2) And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
        If rbcOption(2).Value Then
            pbcMPSpec.CurrentX = tmMSCtrls(MPSPERCENTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcMPSpec.CurrentY = tmMSCtrls(MPSPERCENTINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcMPSpec.Print "Net Dollars ($)"
        Else
            pbcMPSpec.CurrentX = tmPSCtrls(MPSPERCENTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcMPSpec.CurrentY = tmPSCtrls(MPSPERCENTINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcMPSpec.Print "Net Dollars ($)"
        End If
    Else
        If rbcOption(2).Value Then
            pbcMPSpec.CurrentX = tmMSCtrls(MPSPERCENTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcMPSpec.CurrentY = tmMSCtrls(MPSPERCENTINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcMPSpec.Print "Percent (%)"
        Else
            pbcMPSpec.CurrentX = tmPSCtrls(MPSPERCENTINDEX).fBoxX + 15  'fgBoxInsetX
            pbcMPSpec.CurrentY = tmPSCtrls(MPSPERCENTINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
            pbcMPSpec.Print "Percent (%)"
        End If
    End If
    pbcMPSpec.FontSize = flFontSize
    pbcMPSpec.FontName = slFontName
    pbcMPSpec.FontSize = flFontSize
    pbcMPSpec.ForeColor = llColor
    pbcMPSpec.FontBold = True

    If rbcOption(2).Value Then
        For ilBox = imLBCtrls To UBound(tmMSCtrls) Step 1
            pbcMPSpec.CurrentX = tmMSCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcMPSpec.CurrentY = tmMSCtrls(ilBox).fBoxY + fgBoxInsetY
            pbcMPSpec.Print tmMSCtrls(ilBox).sShow
        Next ilBox
    ElseIf rbcOption(3).Value Then
        For ilBox = imLBCtrls To UBound(tmPSCtrls) Step 1
            pbcMPSpec.CurrentX = tmPSCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcMPSpec.CurrentY = tmPSCtrls(ilBox).fBoxY + fgBoxInsetY
            pbcMPSpec.Print tmPSCtrls(ilBox).sShow
        Next ilBox
    End If
End Sub


Private Sub pbcYN_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    If imIBBoxNo = IBACINDEX Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            If imIBSave(4, imIBRowNo) <> 0 Then
                imIBChg = True
            End If
            imIBSave(4, imIBRowNo) = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imIBSave(4, imIBRowNo) <> 1 Then
                imIBChg = True
            End If
            imIBSave(4, imIBRowNo) = 1
            pbcYN_Paint
        End If
    'ElseIf imIBBoxNo = IBSCINDEX Then
    '    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
    '        If imIBSave(5, imIBRowNo) <> 0 Then
    '            imIBChg = True
    '        End If
    '        imIBSave(5, imIBRowNo) = 0
    '        pbcYN_Paint
    '    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
    '        If imIBSave(5, imIBRowNo) <> 1 Then
    '            imIBChg = True
    '        End If
    '        imIBSave(5, imIBRowNo) = 1
    '        pbcYN_Paint
    '    End If
    End If
    If KeyAscii = Asc(" ") Then
        If imIBBoxNo = IBACINDEX Then
            If imIBSave(4, imIBRowNo) = 0 Then
                imIBChg = True
                imIBSave(4, imIBRowNo) = 1
            Else
                imIBChg = True
                imIBSave(4, imIBRowNo) = 0
            End If
        'ElseIf imIBBoxNo = IBSCINDEX Then
        '    If imIBSave(5, imIBRowNo) = 0 Then
        '        imIBChg = True
        '        imIBSave(5, imIBRowNo) = 1
        '    Else
        '        imIBChg = True
        '        imIBSave(5, imIBRowNo) = 0
        '    End If
        End If
        pbcYN_Paint
    End If
End Sub
Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imIBBoxNo = IBACINDEX Then
        If imIBSave(4, imIBRowNo) = 0 Then
            imIBChg = True
            imIBSave(4, imIBRowNo) = 1
        Else
            imIBChg = True
            imIBSave(4, imIBRowNo) = 0
        End If
    'ElseIf imIBBoxNo = IBSCINDEX Then
    '    If imIBSave(5, imIBRowNo) = 0 Then
    '        imIBChg = True
    '        imIBSave(5, imIBRowNo) = 1
    '    Else
    '        imIBChg = True
    '        imIBSave(5, imIBRowNo) = 0
    '    End If
    End If
    pbcYN_Paint
End Sub
Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    If imIBBoxNo = IBACINDEX Then
        If imIBSave(4, imIBRowNo) = 0 Then
            pbcYN.Print "Yes"
        ElseIf imIBSave(4, imIBRowNo) = 1 Then
            pbcYN.Print "No"
        End If
    'ElseIf imIBBoxNo = IBSCINDEX Then
    '    If imIBSave(5, imIBRowNo) = 0 Then
    '        pbcYN.Print "Yes"
    '    ElseIf imIBSave(5, imIBRowNo) = 1 Then
    '        pbcYN.Print "No"
    '    End If
    End If
End Sub
Private Sub plcFix_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcFixSpec_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcItemBill_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcItemBill_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
Private Sub rbcOption_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Dim ilAirTimeNTRConflict As Integer
    Dim ilRepDef As Integer

    Value = rbcOption(Index).Value
    'End of coded added
    If Value Then
        'pbcComingSoon.Visible = False
        edcInstallMsg.Visible = False
        edcNTRMsg.Visible = False
        edcMerchMsg.Visible = False
        edcPromoMsg.Visible = False
        edcWarningMsg.Visible = False
        If Index = 0 Then   'Fixed bill
            'If (Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) <> INSTALLMENT Then
            '    pbcComingSoon.Visible = True
            'End If
            imcKey.Visible = False
            imcTrash.Visible = False
            plcItemBill.Visible = False
            pbcItemBill.Visible = False
            vbcItemBill.Visible = False
            plcMP.Visible = False
            pbcMP.Visible = False
            vbcMP.Visible = False
            pbcMPSpec.Visible = False

            If imShowPS Then
                plcFixSpec.Width = pbcFixSpec.Width + 2 * fgBevelX + cmcGen.Width + 60
                plcFixSpec.Left = Width / 2 - plcFixSpec.Width / 2
                pbcFixSpec.Left = plcFixSpec.Left + fgBevelX
                cmcGen.Left = plcFixSpec.Width - cmcGen.Width - fgBevelX - 60
                plcFixSpec.Visible = True
                pbcFixSpec.Visible = True
            End If
            plcFix.Visible = True
            pbcFix.Visible = True
            vbcFix.Visible = True
            cmcClear.Visible = True
            DoEvents
            pbcIBSTab.Visible = False
            pbcIBTab.Visible = False
            If imShowPS Then
                pbcFSSTab.Visible = True
                pbcFSTab.Visible = True
            End If
            pbcFBSTab.Visible = True
            pbcFBTab.Visible = True
            ilAirTimeNTRConflict = mAirTimeNTRConflict()
            ilRepDef = mAnyRepWithCntr()
            'Recheck if terminate set between test at top and at this point
            If (ilAirTimeNTRConflict = 0) And (Not ilRepDef) And ((Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) = INSTALLMENT) And ((imAllowCashTradeChgs) Or (UBound(tgFBSbf) > LBound(tgFBSbf))) And ((UBound(tgClfCntr) >= 1) Or (UBound(tgIBSbf) > LBound(tgIBSbf))) And (Contract.lbcBillCycle.ListIndex >= 0) Then
                'Only create the BillInfo at this time if Installment exist
                If UBound(tmFBSbf) > LBound(tmFBSbf) Then
                    mNTRAddedToInstallment
                End If
                mIBTotals False
                edcPSAmount.Text = gAddStr(smLnTGross, smNTRTGross)
                mFSSetShow FSAMOUNTINDEX
                lacTotals.Move pbcFix.Left + pbcFix.Width - lacTotals.Width, height - lacTotals.height - 60
                lacTotals.Caption = ""
                mFBTotals True
                mFBSort
                vbcFix.Min = LBound(tmFBSbf) + 1
                'If UBound(tmFBSbf) <= vbcFix.LargeChange + 1 Then
                '    vbcFix.Max = LBound(tmFBSbf) + 1
                'Else
                '    vbcFix.Max = UBound(tmFBSbf) - vbcFix.LargeChange + 1
                'End If
                If UBound(tmInstallBillInfo) + 1 <= vbcFix.LargeChange + 1 Then
                    vbcFix.Max = vbcFix.Min
                Else
                    vbcFix.Max = UBound(tmInstallBillInfo) + 1 - vbcFix.LargeChange
                End If
                vbcFix.Value = vbcFix.Min
            Else
                If (Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) = INSTALLMENT Then
                    If (Contract.lbcBillCycle.ListIndex = 0) Then
                        If (UBound(tgClfCntr) >= 1) Or (UBound(tgIBSbf) > LBound(tgIBSbf)) Then
                            If imAllowCashTradeChgs Then
                                If ilRepDef Then
                                    edcInstallMsg.Text = "Installment not allowed with Rep Vehicles"
                                    edcInstallMsg.Visible = True
                                Else
                                    If ilAirTimeNTRConflict = 1 Then
                                        edcInstallMsg.Text = "Air Time and NTR Agency Commission in Conflict"
                                        edcInstallMsg.Visible = True
                                    ElseIf ilAirTimeNTRConflict = 2 Then
                                        edcInstallMsg.Text = "Agency or Direct Advertiser and NTR Tax in Conflict"
                                        edcInstallMsg.Visible = True
                                    Else
                                        edcInstallMsg.Text = "Both Air Time/NTR Agency Commission and Taxes in Conflict"
                                        edcInstallMsg.Visible = True
                                    End If
                                End If
                            Else
                                edcInstallMsg.Text = "Invoice Generated- Too late to use Installments"
                                edcInstallMsg.Visible = True
                            End If
                        Else
                            edcInstallMsg.Text = "Contract Air Time and/or NTR must be defined prior to defining Installment"
                            edcInstallMsg.Visible = True
                        End If
                    Else
                        edcInstallMsg.Text = "Billing Cycle must be defined as Standard Broadcast prior to defining Installment"
                        edcInstallMsg.Visible = True
                    End If
                Else
                    edcInstallMsg.Text = "Contact Sales@Counterpoint.net to activate this feature"
                    edcInstallMsg.Visible = True
                End If
            End If
        ElseIf Index = 1 Then   'Item Billing
            imcKey.Visible = True
            imcTrash.Visible = True
            plcItemBill.Visible = True
            pbcItemBill.Visible = True
            vbcItemBill.Visible = True
            plcMP.Visible = False
            pbcMP.Visible = False
            vbcMP.Visible = False
            pbcMPSpec.Visible = False
            plcFixSpec.Visible = False
            pbcFixSpec.Visible = False
            plcFix.Visible = False
            pbcFix.Visible = False
            vbcFix.Visible = False
            cmcClear.Visible = False
            pbcIBSTab.Visible = True
            pbcIBTab.Visible = True
            pbcFSSTab.Visible = False
            pbcFSTab.Visible = False
            pbcFBSTab.Visible = False
            pbcFBTab.Visible = False
            lacTotals.Move imcTrash.Left - lacTotals.Width - 60, height - lacTotals.height - 60
            lacTotals.Caption = ""
            mIBTotals True
            vbcItemBill.Min = LBound(tmIBSbf) + 1
            If UBound(tmIBSbf) <= vbcItemBill.LargeChange Then
                vbcItemBill.Max = LBound(tmIBSbf) + 1
            Else
                vbcItemBill.Max = UBound(tmIBSbf) - vbcItemBill.LargeChange + 1
            End If
            vbcItemBill.Value = vbcItemBill.Min
            If tgSpf.sUsingNTR <> "Y" Then
                edcNTRMsg.Text = "Contact Sales@Counterpoint.net to activate this feature"
                edcNTRMsg.Visible = True
            End If
        ElseIf Index = 2 Then   'Merchandising
            imcKey.Visible = False
            imcTrash.Visible = False
            plcItemBill.Visible = False
            pbcItemBill.Visible = False
            vbcItemBill.Visible = False
            pbcFixSpec.Visible = False
            plcFix.Visible = False
            pbcFix.Visible = False
            vbcFix.Visible = False
            cmcClear.Visible = False
            plcFixSpec.Width = pbcMPSpec.Width + 2 * fgBevelX + cmcGen.Width + 60
            plcFixSpec.Left = Width / 2 - plcFixSpec.Width / 2
            pbcMPSpec.Left = plcFixSpec.Left + fgBevelX
            cmcGen.Left = plcFixSpec.Width - cmcGen.Width - fgBevelX - 60
            plcFixSpec.Visible = True
            pbcMPSpec.Visible = True
            plcMP.Visible = True
            pbcMP.Visible = True
            vbcMP.Visible = True
            pbcIBSTab.Visible = False
            pbcIBTab.Visible = False
            pbcFSSTab.Visible = True
            pbcFSTab.Visible = True
            pbcFBSTab.Visible = True
            pbcFBTab.Visible = True
            lacTotals.Move pbcMP.Left + pbcMP.Width - lacTotals.Width, plcMP.Top + plcMP.height + 30
            lacTotals.Caption = ""
            pbcMPSpec.Cls
            pbcMP.Cls
            mMTotals
            vbcMP.Min = LBONE   'LBound(smMBSave, 2)   'LBound(tmMBSbf) + 1
            'If UBound(tmMBSbf) - 1 <= vbcFix.LargeChange + 1 Then
            If UBound(smMBSave, 2) <= vbcMP.LargeChange + 1 Then
                vbcMP.Max = LBONE   'LBound(smMBSave, 2)    'LBound(tmMBSbf)
            Else
                vbcMP.Max = UBound(smMBSave, 2) - vbcMP.LargeChange   'UBound(tmMBSbf) - vbcFix.LargeChange - 1
            End If
            vbcMP.Value = vbcMP.Min
            pbcMPSpec_Paint
            pbcMP_Paint
            If tgSpf.sRUseMerch <> "Y" Then
                edcMerchMsg.Text = "Contact Sales@Counterpoint.net to activate this feature"
                edcMerchMsg.Visible = True
            Else
                If Not rbcOption(2).Enabled Then
                    edcWarningMsg.Text = "Merchandising can only be added or changed for Orders or Holds"
                    edcWarningMsg.Visible = True
                End If
            End If
        ElseIf Index = 3 Then   'Promotion
            imcKey.Visible = False
            imcTrash.Visible = False
            plcItemBill.Visible = False
            pbcItemBill.Visible = False
            vbcItemBill.Visible = False
            pbcFixSpec.Visible = False
            plcFix.Visible = False
            pbcFix.Visible = False
            vbcFix.Visible = False
            cmcClear.Visible = False
            plcFixSpec.Width = pbcMPSpec.Width + 2 * fgBevelX + cmcGen.Width + 60
            plcFixSpec.Left = Width / 2 - plcFixSpec.Width / 2
            pbcMPSpec.Left = plcFixSpec.Left + fgBevelX
            cmcGen.Left = plcFixSpec.Width - cmcGen.Width - fgBevelX - 60
            plcFixSpec.Visible = True
            pbcMPSpec.Visible = True
            plcMP.Visible = True
            pbcMP.Visible = True
            vbcMP.Visible = True
            pbcIBSTab.Visible = False
            pbcIBTab.Visible = False
            pbcFSSTab.Visible = True
            pbcFSTab.Visible = True
            pbcFBSTab.Visible = True
            pbcFBTab.Visible = True
            lacTotals.Move pbcMP.Left + pbcMP.Width - lacTotals.Width, plcMP.Top + plcMP.height + 30
            lacTotals.Caption = ""
            pbcMPSpec.Cls
            pbcMP.Cls
            mPTotals
            vbcMP.Min = LBONE   'LBound(smPBSave, 2)    'LBound(tmPBSbf) + 1
            'If UBound(tmPBSbf) - 1 <= vbcFix.LargeChange + 1 Then
            If UBound(smPBSave, 2) <= vbcMP.LargeChange + 1 Then
                vbcMP.Max = LBONE   'LBound(smPBSave, 2)    'LBound(tmPBSbf)
            Else
                vbcMP.Max = UBound(smPBSave, 2) - vbcMP.LargeChange   'UBound(tmPBSbf) - vbcFix.LargeChange - 1
            End If
            vbcMP.Value = vbcMP.Min
            pbcMPSpec_Paint
            pbcMP_Paint
            If tgSpf.sRUsePromo <> "Y" Then
                edcPromoMsg.Text = "Contact Sales@Counterpoint.net to activate this feature"
                edcPromoMsg.Visible = True
            Else
                If Not rbcOption(3).Enabled Then
                    edcWarningMsg.Text = "Promotional can only be added or changed for Orders or Holds"
                    edcWarningMsg.Visible = True
                End If
            End If
        End If
    End If
    mSetGenCommand
    Exit Sub
End Sub
Private Sub rbcOption_GotFocus(Index As Integer)
    mAllSetShow
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    gCtrlGotFocus ActiveControl
    If Index = 0 Then   'Package bill
        gShowHelpMess tmSbfHelp(), SBFPKG
    ElseIf Index = 1 Then
        gShowHelpMess tmSbfHelp(), SBFITEM
    End If
End Sub
Private Sub rbcOption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imIBBoxNo
        Case IBITEMTYPEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcBItem, edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub
Private Sub tmcDrag_Timer()
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            ilCompRow = vbcItemBill.LargeChange + 1
            If UBound(smIBSave, 2) > ilCompRow Then
                ilMaxRow = ilCompRow
            Else
                ilMaxRow = UBound(smIBSave, 2)
            End If
            For ilRow = 1 To ilMaxRow Step 1
                If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmIBCtrls(IBVEHICLEINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmIBCtrls(IBVEHICLEINDEX).fBoxY + tmIBCtrls(IBVEHICLEINDEX).fBoxH)) Then
                    'Only allow deletion of new- might want to be able to delete unbilled
'                    If (smIBSave(7, ilRow + vbcItemBill.Value - 1) = "R") Or (smIBSave(7, ilRow + vbcItemBill.Value - 1) = "B") Then
'                        Beep
'                        Exit Sub
'                    End If
                    If (smIBSave(7, ilRow + vbcItemBill.Value - 1) = "Y") Then
                        Beep
                        Exit Sub
                    End If
                    mIBSetShow imIBBoxNo
                    imIBBoxNo = -1
                    imIBRowNo = -1
                    imIBRowNo = ilRow + vbcItemBill.Value - 1
                    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                    lacIBFrame.Move 0, tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15) - 30
                    'If gInvertArea call then remove visible setting
                    lacIBFrame.Visible = True
                    pbcArrow.Move plcItemBill.Left - pbcArrow.Width - 15, plcItemBill.Top + tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcItemBill.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    lacIBFrame.Drag vbBeginDrag
                    lacIBFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                    Exit Sub
                End If
            Next ilRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub

Private Sub UserControl_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub UserControl_Initialize()
    imSetButtons = True
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub vbcFix_Change()
    If rbcOption(0).Value Then
        If imSettingValue Then
            pbcFix.Cls
            pbcFix_Paint
            imSettingValue = False
        Else
            mFBSetShow imFBBoxNo
            pbcFix.Cls
            pbcFix_Paint
            'mFBEnableBox imFBBoxNo
        End If
    End If
End Sub
Private Sub vbcFix_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub vbcFix_Scroll()
    If rbcOption(0).Value Then
        If imSettingValue Then
            pbcFix.Cls
            pbcFix_Paint
            imSettingValue = False
        Else
            mFBSetShow imFBBoxNo
            pbcFix.Cls
            pbcFix_Paint
            'mFBEnableBox imFBBoxNo
        End If
    End If
End Sub

Private Sub vbcItemBill_Change()
    If imSettingValue Then
        pbcItemBill.Cls
        pbcItemBill_Paint
        imSettingValue = False
    Else
        mIBSetShow imIBBoxNo
        pbcItemBill.Cls
        pbcItemBill_Paint
        mIBEnableBox imIBBoxNo
    End If
End Sub
Private Sub vbcItemBill_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Special Billing"
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
Private Sub mPaintItemBillTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer

    llColor = pbcItemBill.ForeColor
    slFontName = pbcItemBill.FontName
    flFontSize = pbcItemBill.FontSize
    ilFillStyle = pbcItemBill.FillStyle
    llFillColor = pbcItemBill.FillColor
    pbcItemBill.ForeColor = BLUE
    pbcItemBill.FontBold = False
    pbcItemBill.FontSize = 7
    pbcItemBill.FontName = "Arial"
    pbcItemBill.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    ilHalfY = tmIBCtrls(IBVEHICLEINDEX).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop
    pbcItemBill.Line (tmIBCtrls(IBVEHICLEINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBVEHICLEINDEX).fBoxW + 15, tmIBCtrls(IBVEHICLEINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.CurrentX = tmIBCtrls(IBVEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcItemBill.Print "Vehicle"
    pbcItemBill.Line (tmIBCtrls(IBDATEINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBDATEINDEX).fBoxW + 15, tmIBCtrls(IBDATEINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.CurrentX = tmIBCtrls(IBDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcItemBill.Print "Billing Date"
    pbcItemBill.Line (tmIBCtrls(IBDESCRIPTINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBDESCRIPTINDEX).fBoxW + 15, tmIBCtrls(IBDESCRIPTINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.CurrentX = tmIBCtrls(IBDESCRIPTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcItemBill.Print "Description"
    pbcItemBill.Line (tmIBCtrls(IBITEMTYPEINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBITEMTYPEINDEX).fBoxW + 15, tmIBCtrls(IBITEMTYPEINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.CurrentX = tmIBCtrls(IBITEMTYPEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15
    pbcItemBill.Print "NTR Type"
    pbcItemBill.Line (tmIBCtrls(IBACINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBACINDEX).fBoxW + 15, tmIBCtrls(IBACINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.CurrentX = tmIBCtrls(IBACINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15
    pbcItemBill.Print "A"
    pbcItemBill.CurrentX = tmIBCtrls(IBACINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = ilHalfY + 15
    pbcItemBill.Print "C"
    pbcItemBill.Line (tmIBCtrls(IBSCINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBSCINDEX).fBoxW + 15, tmIBCtrls(IBSCINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.CurrentX = tmIBCtrls(IBSCINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15
    pbcItemBill.Print "Sales"
    pbcItemBill.CurrentX = tmIBCtrls(IBSCINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = ilHalfY + 15
    pbcItemBill.Print "Comm"
    pbcItemBill.Line (tmIBCtrls(IBTXINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBTXINDEX).fBoxW + 15, tmIBCtrls(IBTXINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.CurrentX = tmIBCtrls(IBTXINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15
    pbcItemBill.Print "T"
    pbcItemBill.CurrentX = tmIBCtrls(IBTXINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = ilHalfY + 15
    pbcItemBill.Print "X"
    pbcItemBill.Line (tmIBCtrls(IBAMOUNTINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBAMOUNTINDEX).fBoxW + 15, tmIBCtrls(IBAMOUNTINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.CurrentX = tmIBCtrls(IBAMOUNTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15
    pbcItemBill.Print "Amount/Item"
    pbcItemBill.Line (tmIBCtrls(IBUNITSINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBUNITSINDEX).fBoxW + 15, tmIBCtrls(IBUNITSINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.Line (tmIBCtrls(IBUNITSINDEX).fBoxX, 30)-Step(tmIBCtrls(IBUNITSINDEX).fBoxW - 15, tmIBCtrls(IBUNITSINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcItemBill.CurrentX = tmIBCtrls(IBUNITSINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15
    pbcItemBill.Print "Units Per"
    pbcItemBill.Line (tmIBCtrls(IBNOITEMSINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBNOITEMSINDEX).fBoxW + 15, tmIBCtrls(IBNOITEMSINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.CurrentX = tmIBCtrls(IBNOITEMSINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15
    pbcItemBill.Print "#"
    pbcItemBill.CurrentX = tmIBCtrls(IBNOITEMSINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = ilHalfY + 15
    pbcItemBill.Print "Items"
    pbcItemBill.Line (tmIBCtrls(IBTAMOUNTINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBTAMOUNTINDEX).fBoxW + 15, tmIBCtrls(IBTAMOUNTINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.Line (tmIBCtrls(IBTAMOUNTINDEX).fBoxX, 30)-Step(tmIBCtrls(IBTAMOUNTINDEX).fBoxW - 15, tmIBCtrls(IBTAMOUNTINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcItemBill.CurrentX = tmIBCtrls(IBTAMOUNTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15
    pbcItemBill.Print "Total Bill"
    pbcItemBill.CurrentX = tmIBCtrls(IBTAMOUNTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = ilHalfY + 15
    pbcItemBill.Print "Amount"
    pbcItemBill.Line (tmIBCtrls(IBACQCOSTINDEX).fBoxX - 15, 15)-Step(tmIBCtrls(IBACQCOSTINDEX).fBoxW + 15, tmIBCtrls(IBACQCOSTINDEX).fBoxY - 30), BLUE, B
    pbcItemBill.CurrentX = tmIBCtrls(IBACQCOSTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = 15
    pbcItemBill.Print "Acq Cost/"
    pbcItemBill.CurrentX = tmIBCtrls(IBACQCOSTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcItemBill.CurrentY = ilHalfY + 15
    pbcItemBill.Print "Item"

    ilLineCount = 0
    llTop = tmIBCtrls(1).fBoxY
    Do
        For ilLoop = imLBCtrls To UBound(tmIBCtrls) Step 1
            pbcItemBill.Line (tmIBCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmIBCtrls(ilLoop).fBoxW + 15, tmIBCtrls(ilLoop).fBoxH + 15), BLUE, B
            If (ilLoop = IBUNITSINDEX) Or (ilLoop = IBTAMOUNTINDEX) Then
                pbcItemBill.Line (tmIBCtrls(ilLoop).fBoxX, llTop + 15)-Step(tmIBCtrls(ilLoop).fBoxW - 15, tmIBCtrls(ilLoop).fBoxH - 30), LIGHTYELLOW, BF
            End If
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmIBCtrls(1).fBoxH + 15
    Loop While llTop + tmIBCtrls(1).fBoxH < pbcItemBill.height
    vbcItemBill.LargeChange = ilLineCount - 1
    pbcItemBill.FontSize = flFontSize
    pbcItemBill.FontName = slFontName
    pbcItemBill.FontSize = flFontSize
    pbcItemBill.ForeColor = llColor
    pbcItemBill.FontBold = True
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
Private Sub mPaintFixTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer

    llColor = pbcFix.ForeColor
    slFontName = pbcFix.FontName
    flFontSize = pbcFix.FontSize
    ilFillStyle = pbcFix.FillStyle
    llFillColor = pbcFix.FillColor
    pbcFix.ForeColor = BLUE
    pbcFix.FontBold = False
    pbcFix.FontSize = 7
    pbcFix.FontName = "Arial"
    pbcFix.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    ilHalfY = tmFBCtrls(FBVEHICLEINDEX).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop
    pbcFix.Line (tmFBCtrls(FBVEHICLEINDEX).fBoxX - 15, 15)-Step(tmFBCtrls(FBVEHICLEINDEX).fBoxW + 15, tmFBCtrls(FBVEHICLEINDEX).fBoxY - 30), BLUE, B
    pbcFix.Line (tmFBCtrls(FBVEHICLEINDEX).fBoxX, 30)-Step(tmFBCtrls(FBVEHICLEINDEX).fBoxW - 15, tmFBCtrls(FBVEHICLEINDEX).fBoxY - 60), LIGHTBLUE, BF
    pbcFix.CurrentX = tmFBCtrls(FBVEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFix.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcFix.Print "Vehicle"
    pbcFix.Line (tmFBCtrls(FBDATEINDEX).fBoxX - 15, 15)-Step(tmFBCtrls(FBDATEINDEX).fBoxW + 15, tmFBCtrls(FBDATEINDEX).fBoxY - 30), BLUE, B
    pbcFix.Line (tmFBCtrls(FBDATEINDEX).fBoxX, 30)-Step(tmFBCtrls(FBDATEINDEX).fBoxW - 15, tmFBCtrls(FBDATEINDEX).fBoxY - 60), LIGHTBLUE, BF
    pbcFix.CurrentX = tmFBCtrls(FBDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFix.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcFix.Print "Date"
    pbcFix.Line (tmFBCtrls(FBORDEREDINDEX).fBoxX - 15, 15)-Step(tmFBCtrls(FBORDEREDINDEX).fBoxW + 15, tmFBCtrls(FBORDEREDINDEX).fBoxY - 30), BLUE, B
    pbcFix.Line (tmFBCtrls(FBORDEREDINDEX).fBoxX, 30)-Step(tmFBCtrls(FBORDEREDINDEX).fBoxW - 15, tmFBCtrls(FBORDEREDINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcFix.CurrentX = tmFBCtrls(FBORDEREDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFix.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcFix.Print "Ordered"
    pbcFix.Line (tmFBCtrls(FBREVENUETOTALINDEX).fBoxX - 15, 15)-Step(tmFBCtrls(FBREVENUETOTALINDEX).fBoxW + 15, tmFBCtrls(FBREVENUETOTALINDEX).fBoxY - 30), BLUE, B
    pbcFix.Line (tmFBCtrls(FBREVENUETOTALINDEX).fBoxX, 30)-Step(tmFBCtrls(FBREVENUETOTALINDEX).fBoxW - 15, tmFBCtrls(FBREVENUETOTALINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcFix.CurrentX = tmFBCtrls(FBREVENUETOTALINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFix.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcFix.Print "Revenue"
    pbcFix.Line (tmFBCtrls(FBBILLINGINDEX).fBoxX - 15, 15)-Step(tmFBCtrls(FBBILLINGINDEX).fBoxW + 15, tmFBCtrls(FBBILLINGINDEX).fBoxY - 30), BLUE, B
    pbcFix.CurrentX = tmFBCtrls(FBBILLINGINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFix.CurrentY = 15
    pbcFix.Print "To Invoice"
    pbcFix.Line (tmFBCtrls(FBBILLTOTALINDEX).fBoxX - 15, 15)-Step(tmFBCtrls(FBBILLTOTALINDEX).fBoxW + 15, tmFBCtrls(FBBILLTOTALINDEX).fBoxY - 30), BLUE, B
    pbcFix.Line (tmFBCtrls(FBBILLTOTALINDEX).fBoxX, 30)-Step(tmFBCtrls(FBBILLTOTALINDEX).fBoxW - 15, tmFBCtrls(FBBILLTOTALINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcFix.CurrentX = tmFBCtrls(FBBILLTOTALINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFix.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    If imFixSort = 0 Then
        pbcFix.Print "Invoice/Month"
    Else
        pbcFix.Print "Invoice/Vehicle"
    End If

    ilLineCount = 0
    llTop = tmFBCtrls(1).fBoxY
    Do
        For ilLoop = imLBCtrls To UBound(tmFBCtrls) Step 1
            pbcFix.Line (tmFBCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmFBCtrls(ilLoop).fBoxW + 15, tmFBCtrls(ilLoop).fBoxH + 15), BLUE, B
            If (ilLoop <> FBBILLINGINDEX) Then
                pbcFix.Line (tmFBCtrls(ilLoop).fBoxX, llTop + 15)-Step(tmFBCtrls(ilLoop).fBoxW - 15, tmFBCtrls(ilLoop).fBoxH - 30), LIGHTYELLOW, BF
            End If
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmFBCtrls(1).fBoxH + 15
    Loop While llTop + tmFBCtrls(1).fBoxH < pbcFix.height
    vbcFix.LargeChange = ilLineCount - 1
    pbcFix.FontSize = flFontSize
    pbcFix.FontName = slFontName
    pbcFix.FontSize = flFontSize
    pbcFix.ForeColor = llColor
    pbcFix.FontBold = True
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
Private Sub mPaintMPTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer

    llColor = pbcMP.ForeColor
    slFontName = pbcMP.FontName
    flFontSize = pbcMP.FontSize
    ilFillStyle = pbcMP.FillStyle
    llFillColor = pbcMP.FillColor
    pbcMP.ForeColor = BLUE
    pbcMP.FontBold = False
    pbcMP.FontSize = 7
    pbcMP.FontName = "Arial"
    pbcMP.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    ilHalfY = tmMPBCtrls(MPBVEHICLEINDEX).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop
    pbcMP.Line (tmMPBCtrls(MPBVEHICLEINDEX).fBoxX - 15, 15)-Step(tmMPBCtrls(MPBVEHICLEINDEX).fBoxW + 15, tmMPBCtrls(MPBVEHICLEINDEX).fBoxY - 30), BLUE, B
    pbcMP.Line (tmMPBCtrls(MPBVEHICLEINDEX).fBoxX, 30)-Step(tmMPBCtrls(MPBVEHICLEINDEX).fBoxW - 15, tmMPBCtrls(MPBVEHICLEINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcMP.CurrentX = tmMPBCtrls(MPBVEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcMP.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcMP.Print "Vehicle"
    pbcMP.Line (tmMPBCtrls(MPBDATEINDEX).fBoxX - 15, 15)-Step(tmMPBCtrls(MPBDATEINDEX).fBoxW + 15, tmMPBCtrls(MPBDATEINDEX).fBoxY - 30), BLUE, B
    pbcMP.Line (tmMPBCtrls(MPBDATEINDEX).fBoxX, 30)-Step(tmMPBCtrls(MPBDATEINDEX).fBoxW - 15, tmMPBCtrls(MPBDATEINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcMP.CurrentX = tmMPBCtrls(MPBDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcMP.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcMP.Print "Date"
    pbcMP.Line (tmMPBCtrls(MPBAMOUNTINDEX).fBoxX - 15, 15)-Step(tmMPBCtrls(MPBAMOUNTINDEX).fBoxW + 15, tmMPBCtrls(MPBAMOUNTINDEX).fBoxY - 30), BLUE, B
    pbcMP.CurrentX = tmMPBCtrls(MPBAMOUNTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcMP.CurrentY = 15
    pbcMP.Print "Amount"

    ilLineCount = 0
    llTop = tmMPBCtrls(1).fBoxY
    Do
        For ilLoop = imLBCtrls To UBound(tmMPBCtrls) Step 1
            pbcMP.Line (tmMPBCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmMPBCtrls(ilLoop).fBoxW + 15, tmMPBCtrls(ilLoop).fBoxH + 15), BLUE, B
            If (ilLoop = MPBVEHICLEINDEX) Or (ilLoop = MPBDATEINDEX) Then
                pbcMP.Line (tmMPBCtrls(ilLoop).fBoxX, llTop + 15)-Step(tmMPBCtrls(ilLoop).fBoxW - 15, tmMPBCtrls(ilLoop).fBoxH - 30), LIGHTYELLOW, BF
            End If
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmMPBCtrls(1).fBoxH + 15
    Loop While llTop + tmMPBCtrls(1).fBoxH < pbcMP.height
    vbcMP.LargeChange = ilLineCount - 1
    pbcMP.FontSize = flFontSize
    pbcMP.FontName = slFontName
    pbcMP.FontSize = flFontSize
    pbcMP.ForeColor = llColor
    pbcMP.FontBold = True
End Sub

Private Sub vbcMP_Change()
    If rbcOption(2).Value Then
        If imSettingValue Then
            pbcMP.Cls
            pbcMP_Paint
            imSettingValue = False
        Else
            mMBSetShow imMBBoxNo
            pbcMP.Cls
            pbcMP_Paint
            'mMBEnableBox imMBBoxNo
        End If
    ElseIf rbcOption(3).Value Then
        If imSettingValue Then
            pbcMP.Cls
            pbcMP_Paint
            imSettingValue = False
        Else
            mFBSetShow imFBBoxNo
            pbcMP.Cls
            pbcMP_Paint
            'mFBEnableBox imFBBoxNo
        End If
    End If
End Sub

Private Sub vbcMP_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub vbcMP_Scroll()
    If rbcOption(2).Value Then
        If imSettingValue Then
            pbcMP.Cls
            pbcMP_Paint
            imSettingValue = False
        Else
            mMBSetShow imMBBoxNo
            pbcMP.Cls
            pbcMP_Paint
            'mMBEnableBox imMBBoxNo
        End If
    ElseIf rbcOption(3).Value Then
        If imSettingValue Then
            pbcMP.Cls
            pbcMP_Paint
            imSettingValue = False
        Else
            mFBSetShow imFBBoxNo
            pbcMP.Cls
            pbcMP_Paint
            'mFBEnableBox imFBBoxNo
        End If
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
Private Sub mPaintFixSpecTitle()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        llTop                         ilLineCount               *
'*  ilHalfY                                                                               *
'******************************************************************************************

    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilFillStyle As Integer
    Dim llFillColor As Long

    llColor = pbcFixSpec.ForeColor
    slFontName = pbcFixSpec.FontName
    flFontSize = pbcFixSpec.FontSize
    ilFillStyle = pbcFixSpec.FillStyle
    llFillColor = pbcFixSpec.FillColor
    pbcFixSpec.ForeColor = BLUE
    pbcFixSpec.FontBold = False
    pbcFixSpec.FontSize = 7
    pbcFixSpec.FontName = "Arial"
    pbcFixSpec.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    pbcFixSpec.Line (tmFSCtrls(FSSTARTDATEINDEX).fBoxX - 15, 15)-Step(tmFSCtrls(FSSTARTDATEINDEX).fBoxW + 15, tmFSCtrls(FSSTARTDATEINDEX).fBoxH + 15), BLUE, B
    pbcFixSpec.CurrentX = tmFSCtrls(FSSTARTDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFixSpec.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcFixSpec.Print "Start Invoice Date"
    pbcFixSpec.Line (tmFSCtrls(FSNOMONTHSINDEX).fBoxX - 15, 15)-Step(tmFSCtrls(FSNOMONTHSINDEX).fBoxW + 15, tmFSCtrls(FSNOMONTHSINDEX).fBoxH + 15), BLUE, B
    pbcFixSpec.CurrentX = tmFSCtrls(FSNOMONTHSINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFixSpec.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcFixSpec.Print "# Months"
    pbcFixSpec.Line (tmFSCtrls(FSENDDATEINDEX).fBoxX - 15, 15)-Step(tmFSCtrls(FSENDDATEINDEX).fBoxW + 15, tmFSCtrls(FSENDDATEINDEX).fBoxH + 15), BLUE, B
    pbcFixSpec.CurrentX = tmFSCtrls(FSENDDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFixSpec.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcFixSpec.Print "End Invoice Date"
    pbcFixSpec.Line (tmFSCtrls(FSNOMONTHSONINDEX).fBoxX - 15, 15)-Step(tmFSCtrls(FSNOMONTHSONINDEX).fBoxW + 15, tmFSCtrls(FSNOMONTHSONINDEX).fBoxH + 15), BLUE, B
    pbcFixSpec.CurrentX = tmFSCtrls(FSNOMONTHSONINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFixSpec.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcFixSpec.Print "# Months On"
    pbcFixSpec.Line (tmFSCtrls(FSNOMONTHSOFFINDEX).fBoxX - 15, 15)-Step(tmFSCtrls(FSNOMONTHSOFFINDEX).fBoxW + 15, tmFSCtrls(FSNOMONTHSOFFINDEX).fBoxH + 15), BLUE, B
    pbcFixSpec.CurrentX = tmFSCtrls(FSNOMONTHSOFFINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFixSpec.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcFixSpec.Print "# Months Off"
    pbcFixSpec.Line (tmFSCtrls(FSAMOUNTINDEX).fBoxX - 15, 15)-Step(tmFSCtrls(FSAMOUNTINDEX).fBoxW + 15, tmFSCtrls(FSAMOUNTINDEX).fBoxH + 15), BLUE, B
    pbcFixSpec.Line (tmFSCtrls(FSAMOUNTINDEX).fBoxX, 30)-Step(tmFSCtrls(FSAMOUNTINDEX).fBoxW - 15, tmFSCtrls(FSAMOUNTINDEX).fBoxH - 60), LIGHTYELLOW, BF
    pbcFixSpec.CurrentX = tmFSCtrls(FSAMOUNTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcFixSpec.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcFixSpec.Print "Total Gross"
    pbcFixSpec.FontSize = flFontSize
    pbcFixSpec.FontName = slFontName
    pbcFixSpec.FontSize = flFontSize
    pbcFixSpec.ForeColor = llColor
    pbcFixSpec.FontBold = True
End Sub




Private Sub mFBSort()
    Dim ilLoop As Integer
    Dim slDateSort As String


    pbcFix.Cls
    For ilLoop = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
        slDateSort = Trim$(str$(tmInstallBillInfo(ilLoop).lBillDate))
        Do While Len(slDateSort) < 6
            slDateSort = "0" & slDateSort
        Loop
        If imFixSort = 0 Then
            tmInstallBillInfo(ilLoop).sKey = slDateSort & tmInstallBillInfo(ilLoop).sVehName
        Else
            tmInstallBillInfo(ilLoop).sKey = tmInstallBillInfo(ilLoop).sVehName & slDateSort
        End If
    Next ilLoop
    If UBound(tmInstallBillInfo) - 1 > 0 Then
        ArraySortTyp fnAV(tmInstallBillInfo(), 0), UBound(tmInstallBillInfo), 0, LenB(tmInstallBillInfo(0)), 0, LenB(tmInstallBillInfo(0).sKey), 0
    End If
    pbcFix_Paint
End Sub

Private Sub mNTRAddedToInstallment()
    Dim ilLoop As Integer
    Dim ilInstall As Integer
    Dim ilFound As Integer
    Dim llDate As Long
    Dim llDollars As Long
    Dim ilUpper As Integer
    '12/18/17: Add separation of NTR from AirTime
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilMnfItem As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer

    For ilLoop = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
        tmInstallBillInfo(ilLoop).lNTROrderedDollars = 0
    Next ilLoop
    For ilLoop = LBONE To UBound(smIBSave, 2) - 1 Step 1
        If imIBSave(3, ilLoop) <> -1 Then
            '12/18/17: Break out NTR separate from Air Time
            ilMnfItem = 0
            If bgBreakoutNTR Then
                slNameCode = tmItemCode(imIBSave(3, ilLoop) - 1).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilMnfItem = CInt(slCode)
            End If
            ilFound = False
            llDate = gDateValue(smIBSave(8, ilLoop))
            llDollars = gStrDecToLong(smIBSave(6, ilLoop), 2)
            For ilInstall = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
                If tmInstallBillInfo(ilInstall).iVefCode = lbcBVehicle.ItemData(imIBSave(1, ilLoop)) Then
                    '12/18/17: Break out NTR separate from Air Time
                    If (Not bgBreakoutNTR) Or ((bgBreakoutNTR) And (tmInstallBillInfo(ilInstall).iMnfItem = ilMnfItem)) Then
                        If tmInstallBillInfo(ilInstall).lBillDate = llDate Then
                            tmInstallBillInfo(ilInstall).lNTROrderedDollars = tmInstallBillInfo(ilInstall).lNTROrderedDollars + llDollars
                            ilFound = True
                        End If
                    End If
                End If
            Next ilInstall
            If Not ilFound Then
                ilUpper = UBound(tmInstallBillInfo)
                tmInstallBillInfo(ilUpper).iVefCode = lbcBVehicle.ItemData(imIBSave(1, ilLoop))
                tmInstallBillInfo(ilUpper).sVehName = lbcBVehicle.List(imIBSave(1, ilLoop))
                tmInstallBillInfo(ilUpper).lBillDollars = 0
                tmInstallBillInfo(ilUpper).lAirOrderedDollars = 0
                tmInstallBillInfo(ilUpper).lBillDate = llDate
                tmInstallBillInfo(ilUpper).lNTROrderedDollars = llDollars
                tmInstallBillInfo(ilUpper).sBilledFlag = smIBSave(7, ilLoop)
                tmInstallBillInfo(ilUpper).sType = "O"
                If smIBSave(7, ilLoop) = "Y" Then
                    tmInstallBillInfo(ilUpper).sType = "I"
                End If
                '12/18/17: Break out NTR separate from Air Time
                tmInstallBillInfo(ilUpper).iMnfItem = 0
                tmInstallBillInfo(ilUpper).sMnfItem = ""
                If bgBreakoutNTR Then
                    tmInstallBillInfo(ilUpper).iMnfItem = ilMnfItem
                    For ilIndex = 0 To UBound(tmItemCode) - 1 Step 1 'lbcItemCode.ListCount - 1 Step 1
                        slNameCode = tmItemCode(ilIndex).sKey  'lbcItemCode.List(ilIndex)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If ilMnfItem = Val(slCode) Then
                            ilRet = gParseItem(slNameCode, 1, "\", slName)
                            tmInstallBillInfo(ilUpper).sMnfItem = slName
                            Exit For
                        End If
                    Next ilIndex
                End If
                ReDim Preserve tmInstallBillInfo(0 To ilUpper + 1) As INSTALLBILLINFO
            End If
        End If
    Next ilLoop
End Sub

Public Sub Action(ilType As Integer)
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Select Case ilType
        Case 1  'Clear Focus
            mAllSetShow
            pbcArrow.Visible = False
            lacIBFrame.Visible = False
        Case 2  'Init function
            'Test if unloading control
            ilRet = 0
            On Error GoTo UserControlErr:
            ilIndex = Contract.tscLine.SelectedItem.Index
            If ilRet = 0 Then
                Form_Load
                Form_Activate
                Select Case Contract.tscLine.SelectedItem.Index
                    Case imTabMap(MULTIMEDIA)   '1  'Multi-Media
                    'Case 2  'Digital
                    Case imTabMap(TABNTR)   '3  'NTR
                        If rbcOption(1).Value Then
                            rbcOption_Click 1
                        Else
                            rbcOption(1).Value = True
                        End If
                    Case imTabMap(TABAIRTIME)   '4  'Air Time
                    Case imTabMap(TABPODCASTCPM)   'Podcast CPM
                    Case imTabMap(TABMERCH)   '5  'Merchandising
                        If rbcOption(2).Value Then
                            rbcOption_Click 2
                        Else
                            rbcOption(2).Value = True
                        End If
                    Case imTabMap(TABPROMO)   '6  'Promotional
                        If rbcOption(3).Value Then
                            rbcOption_Click 3
                        Else
                            rbcOption(3).Value = True
                        End If
                    Case imTabMap(TABINSTALL)   '7  'Installment
                        If rbcOption(0).Value = True Then
                            rbcOption_Click 0
                        Else
                            rbcOption(0).Value = True
                        End If
                End Select
            End If
        Case 3  'terminate function
            mAllSetShow
            pbcArrow.Visible = False
            lacIBFrame.Visible = False
            cmcCancel_Click
        Case 4  'Clear
            '1/23/18: Clear flags TTP 8776
            smMPercent = ""
            smPPercent = ""
            imFBChg = False
            imIBChg = False
            imMBChg = False
            imPBChg = False
            
            ReDim tmFBSbf(0 To 0) As SBFLIST
            ReDim lmFBSbfCode(0 To 0) As Long
            ReDim tmIBSbf(0 To 0) As SBFLIST
            ReDim lmIBSbfCode(0 To 0) As Long
            ReDim tmInstallBillInfo(0 To 0) As INSTALLBILLINFO
            '12/18/17: Break out NTR separate from Air Time
            bgBreakoutNTR = True
            ReDim tmMBSbf(0 To 0) As SBFLIST
            ReDim tmPBSbf(0 To 0) As SBFLIST
            'ReDim smIBSave(1 To 11, 1 To 1) As String
            'ReDim imIBSave(1 To 7, 1 To 1) As Integer
            'ReDim smIBShow(1 To IBACQCOSTINDEX, 1 To 1) As String
            'ReDim lmIBSave(1 To 2, 1 To 1) As Long
            'ReDim smMBSave(1 To 3, 1 To 1) As String
            'ReDim imMBSave(1 To 1, 1 To 1) As Integer
            'ReDim smMBShow(1 To MPBAMOUNTINDEX, 1 To 1) As String
            'ReDim smPBSave(1 To 3, 1 To 1) As String
            'ReDim imPBSave(1 To 1, 1 To 1) As Integer
            'ReDim smPBShow(1 To MPBAMOUNTINDEX, 1 To 1) As String
            
            ReDim smIBSave(0 To 11, 0 To 1) As String
            ReDim imIBSave(0 To 7, 0 To 1) As Integer
            ReDim smIBShow(0 To IBACQCOSTINDEX, 0 To 1) As String
            ReDim lmIBSave(0 To 2, 0 To 1) As Long
            ReDim smMBSave(0 To 3, 0 To 1) As String
            'ReDim imMBSave(0 To 1) As Integer
            ReDim imMBSave(0 To 1, 0 To 1) As Integer
            ReDim smMBShow(0 To MPBAMOUNTINDEX, 0 To 1) As String
            ReDim smPBSave(0 To 3, 0 To 1) As String
            'ReDim imPBSave(0 To 1) As Integer
            ReDim imPBSave(0 To 1, 0 To 1) As Integer
            ReDim smPBShow(0 To MPBAMOUNTINDEX, 0 To 1) As String
            For ilLoop = imLBCtrls To UBound(tmFSCtrls) Step 1
                tmFSCtrls(ilLoop).sShow = ""
            Next ilLoop
            vbcFix.Min = LBound(tmInstallBillInfo) + 1 'LBound(smFBShow, 2)
            vbcFix.Max = UBound(tmInstallBillInfo) + 1  'LBound(smFBShow, 2)
            vbcFix.Value = vbcFix.Min
            vbcItemBill.Min = LBONE 'LBound(smIBShow, 2)
            vbcItemBill.Max = LBONE 'LBound(smIBShow, 2)
            vbcItemBill.Value = vbcItemBill.Min
            vbcMP.Min = LBONE   'LBound(smMBShow, 2)
            vbcMP.Max = LBONE   'LBound(smMBShow, 2)
            vbcMP.Value = vbcMP.Min
            imIBBoxNo = -1 'Initialize current Box to N/A
            imFSBoxNo = -1
            imFBBoxNo = -1
            imMSBoxNo = -1
            imMBBoxNo = -1
            imPSBoxNo = -1
            imPBBoxNo = -1
            imIBRowNo = -1
            imFBRowNo = -1
            imMBRowNo = -1
            imPBRowNo = -1
            edcPPercent.Text = ""
            lbcPSSDate.ListIndex = -1
            lbcPESDate.ListIndex = -1
            edcMPercent.Text = ""
            lbcMSSDate.ListIndex = -1
            lbcMESDate.ListIndex = -1
            lbcPSDate.ListIndex = -1
            edcNoPeriods.Text = ""
            lbcPEDate.ListIndex = -1
            edcNoMonthsOn.Text = ""
            edcNoMonthsOff.Text = ""
            edcPSAmount.Text = ""
            pbcItemBill.Cls
            pbcItemBill_Paint
            pbcMP.Cls
            pbcMP_Paint
            pbcMPSpec.Cls
            pbcMPSpec_Paint
            pbcFix.Cls
            pbcFix_Paint
            pbcFixSpec.Cls
            pbcFixSpec_Paint
            lacTotals.Caption = ""
        Case 5  'Save
            mAllSetShow
            pbcArrow.Visible = False
            lacIBFrame.Visible = False
            mSaveRec
        Case 6  '12/5/14: Hide Installment
            plcFix.Visible = False
            pbcFix.Visible = False
            vbcFix.Visible = False
            plcFixSpec.Visible = False
            pbcFixSpec.Visible = False
            edcInstallMsg.Visible = False
    End Select
    Exit Sub
UserControlErr:
    ilRet = 1
    Resume Next
End Sub

Public Property Let Enabled(ilState As Integer)
    UserControl.Enabled = ilState
    PropertyChanged "Enabled"
End Property

Public Property Get Verify() As Integer
    mAllSetShow
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    If (imUpdateAllowed) Then
        Select Case Contract.tscLine.SelectedItem.Index
            Case imTabMap(TABMULTIMEDIA)    '1  'Multi-Media
                If mIBTestFields() = NO Then
                    Verify = False
                Else
                    Verify = True
                End If
            'Case 2  'Digital
            Case imTabMap(TABNTR)    '3  'NTR
                mMoveIBCtrlToRec
                If mIBTestFields() = NO Then
                    mIBEnableBox imIBBoxNo
                    Verify = False
                Else
                    Verify = True
                End If

            Case imTabMap(TABAIRTIME)    '4  'Air Time
            Case imTabMap(TABPODCASTCPM)    'Podcast CPM
            Case imTabMap(TABMERCH)    '5  'Merchandising
                mMoveMBCtrlToRec
                If mMBTestFields() = NO Then
                    mMBEnableBox imMBBoxNo
                    Verify = False
                Else
                    Verify = True
                End If
            Case imTabMap(TABPROMO)    '6  'Promotional
                mMovePBCtrlToRec
                If mPBTestFields() = NO Then
                    mPBEnableBox imPBBoxNo
                    Verify = False
                Else
                    Verify = True
                End If
            Case imTabMap(TABINSTALL)    '7  'Installment
                Verify = True
        End Select
    Else
        Verify = True
    End If
End Property
Public Property Get CheckVehicles(ilType As Integer) As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilRet As Integer

    CheckVehicles = True
    mVehPop
    'ilType (I): 0=Merch; 1=Promo
    Select Case ilType
        Case 0  'Merch
            For ilLoop = 0 To UBound(tmMBSbf) - 1 Step 1
                If (tmMBSbf(ilLoop).iStatus = 0) Or (tmMBSbf(ilLoop).iStatus = 1) Then
                    ilFound = False
                    slRecCode = Trim$(str$(tmMBSbf(ilLoop).SbfRec.iBillVefCode))
                    For ilTest = 0 To lbcVehicle.ListCount - 1 Step 1
                        slNameCode = lbcVehicle.List(ilTest)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        On Error GoTo 0
                        If slRecCode = slCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilTest
                    If Not ilFound Then
                        CheckVehicles = False
                        Exit For
                    End If
                End If
            Next ilLoop
        Case 1  'Promo
            For ilLoop = 0 To UBound(tmPBSbf) - 1 Step 1
                If (tmPBSbf(ilLoop).iStatus = 0) Or (tmPBSbf(ilLoop).iStatus = 1) Then
                    ilFound = False
                    slRecCode = Trim$(str$(tmPBSbf(ilLoop).SbfRec.iBillVefCode))
                    For ilTest = 0 To lbcVehicle.ListCount - 1 Step 1
                        slNameCode = lbcVehicle.List(ilTest)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        On Error GoTo 0
                        If slRecCode = slCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilTest
                    If Not ilFound Then
                        CheckVehicles = False
                        Exit For
                    End If
                End If
            Next ilLoop
    End Select
End Property

Public Property Get NTRDataExist() As Integer
    If UBound(smIBSave, 2) > LBONE Then
        NTRDataExist = True
    Else
        NTRDataExist = False
    End If
End Property
Public Property Get MerchandisingDataExist() As Integer
    If UBound(smMBSave, 2) > LBONE Then
        MerchandisingDataExist = True
    Else
        MerchandisingDataExist = False
    End If
End Property
Public Property Get PromotionalDataExist() As Integer
    If UBound(smPBSave, 2) > LBONE Then
        PromotionalDataExist = True
    Else
        PromotionalDataExist = False
    End If
End Property
Public Property Get InstallmentDataExist() As Integer
    Dim ilLoop As Integer
    For ilLoop = LBound(tmInstallBillInfo) To UBound(tmInstallBillInfo) - 1 Step 1
        '12/18/17: Break out NTR separate from Air Time
        'If tmInstallBillInfo(ilLoop).lBillDollars > 0 Then
        If tmInstallBillInfo(ilLoop).lBillDollars <> 0 Then
            InstallmentDataExist = True
            Exit Property
        End If
    Next ilLoop
    InstallmentDataExist = False
End Property


Private Function mIBColOk(ilRowNo As Integer, ilBox As Integer) As Integer
    mIBColOk = True
    If imIBSave(7, ilRowNo) > 0 Then
        If smIBSave(11, ilRowNo) = "Y" Then
            If (ilBox = IBVEHICLEINDEX) Or (ilBox = IBITEMTYPEINDEX) Or (ilBox = IBAMOUNTINDEX) Or (ilBox = IBNOITEMSINDEX) Or (ilBox = IBUNITSINDEX) Or (ilBox = IBTAMOUNTINDEX) Or (ilBox = IBACQCOSTINDEX) Then
                mIBColOk = False
                Exit Function
            End If
        Else
            If (ilBox = IBVEHICLEINDEX) Or (ilBox = IBDATEINDEX) Or (ilBox = IBITEMTYPEINDEX) Or (ilBox = IBAMOUNTINDEX) Or (ilBox = IBNOITEMSINDEX) Or (ilBox = IBUNITSINDEX) Or (ilBox = IBTAMOUNTINDEX) Or (ilBox = IBACQCOSTINDEX) Then
                mIBColOk = False
                Exit Function
            End If
        End If
    End If
End Function

Private Sub mIBSetFocus(ilBoxNo As Integer)
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmIBCtrls)) Then
        Exit Sub
    End If
    If (imIBRowNo < vbcItemBill.Value) Or (imIBRowNo >= vbcItemBill.Value + vbcItemBill.LargeChange + 1) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case IBVEHICLEINDEX 'Vehicle
            edcDropDown.SetFocus
        Case IBDATEINDEX 'Date
            edcDropDown.SetFocus
        Case IBITEMTYPEINDEX 'Item bill type
            edcDropDown.SetFocus
        Case IBACINDEX
            pbcYN.SetFocus
        Case IBSCINDEX
            edcSalesComm.SetFocus
        Case IBTXINDEX
            edcDropDown.SetFocus
        Case IBAMOUNTINDEX
            edcAmount.SetFocus
        Case IBNOITEMSINDEX
            edcNoItems.SetFocus
        Case IBACQCOSTINDEX
            edcAcqAmount.SetFocus
    End Select

End Sub

Private Sub mSetBillCycle()
    Dim slStr As String
    
    smBillCycle = "S"   '"B"
    If (Contract.lbcBillCycle.ListIndex >= 0) And (Contract.lbcBillCycle.ListIndex <= 2) Then
        smBillCycle = Left(Contract.lbcBillCycle.List(Contract.lbcBillCycle.ListIndex), 1)
    End If
End Sub

