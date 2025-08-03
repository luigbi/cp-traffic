VERSION 5.00
Begin VB.UserControl SiteTabs 
   Appearance      =   0  'Flat
   ClientHeight    =   5895
   ClientLeft      =   840
   ClientTop       =   2190
   ClientWidth     =   10080
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5895
   ScaleWidth      =   10080
   Begin VB.Frame frcRepNet 
      Height          =   4680
      Left            =   8445
      TabIndex        =   83
      Top             =   5340
      Width           =   9075
      Begin VB.Frame frcNet 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3690
         Left            =   15
         TabIndex        =   84
         Top             =   720
         Width           =   8910
         Begin VB.TextBox edcFTPUserID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   300
            Left            =   1395
            MaxLength       =   20
            TabIndex        =   88
            Top             =   0
            Width           =   2535
         End
         Begin VB.TextBox edcFTPUserPW 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2085
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   90
            Top             =   420
            Width           =   2550
         End
         Begin VB.TextBox edcFTPPort 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   300
            Left            =   1395
            MaxLength       =   5
            TabIndex        =   92
            Top             =   840
            Width           =   1080
         End
         Begin VB.TextBox edcFTPAddress 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   300
            Left            =   1395
            MaxLength       =   20
            TabIndex        =   94
            Top             =   1245
            Width           =   3000
         End
         Begin VB.TextBox edcFTPImport 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   300
            Left            =   2085
            MaxLength       =   120
            TabIndex        =   96
            Top             =   1680
            Width           =   6795
         End
         Begin VB.TextBox edcFTPExport 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   300
            Left            =   2085
            MaxLength       =   120
            TabIndex        =   98
            Top             =   2100
            Width           =   6795
         End
         Begin VB.TextBox edcIISRootURL 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   300
            Left            =   1395
            MaxLength       =   120
            TabIndex        =   100
            Top             =   2520
            Width           =   7485
         End
         Begin VB.TextBox edcIISRegSection 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   300
            Left            =   2085
            MaxLength       =   80
            TabIndex        =   102
            Top             =   2940
            Width           =   6795
         End
         Begin VB.Label lacFTPUserID 
            Appearance      =   0  'Flat
            Caption         =   "FTP User ID                                                                       Address"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   75
            TabIndex        =   87
            Top             =   30
            Width           =   1275
         End
         Begin VB.Label lacFTPUserPW 
            Appearance      =   0  'Flat
            Caption         =   "FTP User Password                                                                      Address"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   75
            TabIndex        =   89
            Top             =   450
            Width           =   1830
         End
         Begin VB.Label lacFTPPort 
            Appearance      =   0  'Flat
            Caption         =   "FTP Port                                                                       Address"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   75
            TabIndex        =   91
            Top             =   870
            Width           =   915
         End
         Begin VB.Label lacFTPAddress 
            Appearance      =   0  'Flat
            Caption         =   "FTP Address                                                                       Address"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   75
            TabIndex        =   93
            Top             =   1275
            Width           =   1275
         End
         Begin VB.Label lacFPTImport 
            Appearance      =   0  'Flat
            Caption         =   "FTP Import Directory                                                                       Address"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   75
            TabIndex        =   95
            Top             =   1710
            Width           =   1875
         End
         Begin VB.Label lacFTPExport 
            Appearance      =   0  'Flat
            Caption         =   "FTP Export Directory                                                                     Address"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   75
            TabIndex        =   97
            Top             =   2130
            Width           =   1935
         End
         Begin VB.Label lacIISRootURL 
            Appearance      =   0  'Flat
            Caption         =   "IIS Root URL                                                                       Address"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   75
            TabIndex        =   99
            Top             =   2550
            Width           =   1275
         End
         Begin VB.Label lacIISRegSection 
            Appearance      =   0  'Flat
            Caption         =   "IIS Register Section                                                                        Address"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   75
            TabIndex        =   101
            Top             =   2985
            Width           =   1875
         End
      End
      Begin VB.TextBox edcDBID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   1380
         MaxLength       =   5
         TabIndex        =   86
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label lacDBID 
         Appearance      =   0  'Flat
         Caption         =   "Database ID                                                                           Address"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         TabIndex        =   85
         Top             =   330
         Width           =   1275
      End
   End
   Begin VB.Frame frcEMail 
      Height          =   4995
      Left            =   8760
      TabIndex        =   57
      Top             =   4380
      Width           =   8460
      Begin VB.CommandButton cmdVerifyByCSI 
         Caption         =   "Verify with Counterpoint Values"
         Height          =   390
         Left            =   4560
         TabIndex        =   103
         Top             =   4440
         Width           =   3045
      End
      Begin VB.TextBox edcVerifyTo 
         BackColor       =   &H00FFFF80&
         Height          =   360
         Left            =   6195
         TabIndex        =   67
         Top             =   2340
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox chkTLS 
         Caption         =   "Yes"
         Height          =   405
         Left            =   5340
         TabIndex        =   77
         Top             =   2295
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox edcVerifyEmail 
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Please verify before saving information"
         Top             =   3960
         Width           =   6390
      End
      Begin VB.CommandButton cmdVerifyEmail 
         Caption         =   "&Verify Client Information"
         Height          =   390
         Left            =   1320
         TabIndex        =   81
         Top             =   4440
         Width           =   2925
      End
      Begin VB.TextBox edcHost 
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   330
         MaxLength       =   80
         TabIndex        =   69
         Top             =   435
         Width           =   7695
      End
      Begin VB.TextBox edcPort 
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   5370
         MaxLength       =   5
         TabIndex        =   75
         Top             =   1935
         Width           =   1095
      End
      Begin VB.TextBox edcAcctName 
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   360
         MaxLength       =   80
         TabIndex        =   71
         Top             =   1170
         Width           =   7695
      End
      Begin VB.TextBox edcPassword 
         BackColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   330
         MaxLength       =   80
         PasswordChar    =   "*"
         TabIndex        =   73
         Top             =   1920
         Width           =   4095
      End
      Begin VB.TextBox edcFromName 
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   330
         MaxLength       =   80
         TabIndex        =   78
         Top             =   2760
         Width           =   7935
      End
      Begin VB.TextBox edcFromAddress 
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   360
         MaxLength       =   80
         TabIndex        =   80
         Top             =   3480
         Width           =   7935
      End
      Begin VB.Label lblFrom 
         Caption         =   "From Name, e.g., XYZ Radio Networks IT dept."
         Height          =   255
         Left            =   330
         TabIndex        =   76
         Top             =   2415
         Width           =   5415
      End
      Begin VB.Label lblPort 
         Caption         =   "Port Number, e.g., 25"
         Height          =   255
         Left            =   5370
         TabIndex        =   74
         Top             =   1680
         Width           =   2130
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   330
         TabIndex        =   72
         Top             =   1665
         Width           =   975
      End
      Begin VB.Label lblHost 
         Caption         =   "SMTP Server, e.g., smtp.att.yahoo.com"
         Height          =   255
         Left            =   330
         TabIndex        =   68
         Top             =   180
         Width           =   4815
      End
      Begin VB.Label lblUserName 
         Caption         =   "Account Name, e.g., abcd@xyz.com"
         Height          =   255
         Left            =   330
         TabIndex        =   70
         Top             =   915
         Width           =   4815
      End
      Begin VB.Label lblFromAddress 
         Caption         =   "From Address, e.g., admin@xyzradionetworks.net"
         Height          =   255
         Left            =   330
         TabIndex        =   79
         Top             =   3240
         Width           =   5415
      End
   End
   Begin VB.PictureBox plcResearch 
      Height          =   2610
      Left            =   9015
      ScaleHeight     =   2550
      ScaleWidth      =   8490
      TabIndex        =   13
      Top             =   4185
      Width           =   8550
      Begin VB.CheckBox ckcAudByPackage 
         Alignment       =   1  'Right Justify
         Caption         =   "Audience Data by Package"
         Height          =   240
         Left            =   120
         TabIndex        =   66
         Top             =   1905
         Width           =   2880
      End
      Begin VB.CheckBox ckcHideDemo 
         Alignment       =   1  'Right Justify
         Caption         =   "Show ""Impressions"" and ""Download""  on Proposals/Orders"
         Height          =   240
         Left            =   105
         TabIndex        =   65
         Top             =   1590
         Width           =   5385
      End
      Begin VB.CheckBox ckcRschCustDemo 
         Alignment       =   1  'Right Justify
         Caption         =   "Single Custom Demo Only"
         Height          =   240
         Left            =   105
         TabIndex        =   64
         Top             =   1275
         Width           =   2565
      End
      Begin VB.PictureBox plcSGRPCPPCal 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   6480
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   360
         Width           =   6480
         Begin VB.OptionButton rbcSGRPCPPCal 
            Caption         =   "Audience"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   5340
            TabIndex        =   58
            Top             =   0
            Width           =   1140
         End
         Begin VB.OptionButton rbcSGRPCPPCal 
            Caption         =   "Aud/GRP(2places)"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3330
            TabIndex        =   104
            Top             =   0
            Width           =   2040
         End
         Begin VB.OptionButton rbcSGRPCPPCal 
            Caption         =   "Rating"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2385
            TabIndex        =   56
            Top             =   0
            Width           =   945
         End
      End
      Begin VB.PictureBox plcSAudData 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   6840
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   45
         Width           =   6840
         Begin VB.OptionButton rbcSAudData 
            Caption         =   "Units"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   5055
            TabIndex        =   54
            Top             =   0
            Width           =   810
         End
         Begin VB.OptionButton rbcSAudData 
            Caption         =   "Tens"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   4215
            TabIndex        =   53
            Top             =   0
            Width           =   795
         End
         Begin VB.OptionButton rbcSAudData 
            Caption         =   "Hundreds"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2985
            TabIndex        =   52
            Top             =   0
            Width           =   1155
         End
         Begin VB.OptionButton rbcSAudData 
            Caption         =   "Thousands"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1650
            TabIndex        =   51
            Top             =   0
            Width           =   1410
         End
      End
      Begin VB.PictureBox plcWeight 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   6585
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   660
         Width           =   6585
         Begin VB.OptionButton rbcWeight 
            Caption         =   "Times"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   3915
            TabIndex        =   60
            Top             =   0
            Width           =   870
         End
         Begin VB.OptionButton rbcWeight 
            Caption         =   "Days"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   4845
            TabIndex        =   61
            Top             =   0
            Width           =   795
         End
         Begin VB.OptionButton rbcWeight 
            Caption         =   "None"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   5640
            TabIndex        =   62
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.CheckBox ckcHiddenOverride 
         Alignment       =   1  'Right Justify
         Caption         =   "For Hidden Line Research, ignore Overrides"
         Height          =   240
         Left            =   120
         TabIndex        =   63
         Top             =   960
         Width           =   4065
      End
   End
   Begin VB.PictureBox plcSports 
      Height          =   3705
      Left            =   9330
      ScaleHeight     =   3645
      ScaleWidth      =   8295
      TabIndex        =   1
      Top             =   4065
      Visible         =   0   'False
      Width           =   8355
      Begin VB.TextBox edcEventSubtotal 
         Height          =   285
         Index           =   1
         Left            =   3345
         MaxLength       =   15
         TabIndex        =   12
         Top             =   2295
         Width           =   3000
      End
      Begin VB.TextBox edcEventSubtotal 
         Height          =   285
         Index           =   0
         Left            =   3345
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1905
         Width           =   3000
      End
      Begin VB.TextBox edcEventTitle 
         Height          =   285
         Index           =   1
         Left            =   3345
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1485
         Width           =   3000
      End
      Begin VB.TextBox edcEventTitle 
         Height          =   285
         Index           =   0
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1095
         Width           =   3000
      End
      Begin VB.CheckBox ckcSportsInfo 
         Caption         =   "Pre-empt Regular Programming"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   750
         Width           =   3285
      End
      Begin VB.CheckBox ckcSportsInfo 
         Caption         =   "Using Feed Source"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2430
      End
      Begin VB.CheckBox ckcSportsInfo 
         Caption         =   "Using Multiple Language Feeds"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   435
         Width           =   3135
      End
      Begin VB.Label lacEventSubtotal 
         Caption         =   "Event Subtotal Title 2"
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2340
         Width           =   2655
      End
      Begin VB.Label lacEventSubtotal 
         Caption         =   "Event Subtotal Title 1"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1935
         Width           =   3000
      End
      Begin VB.Label lacEventTitle 
         Caption         =   "@ Event Title 2 (Ex: Home Team)"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1530
         Width           =   3030
      End
      Begin VB.Label lacEventTitle 
         Caption         =   "Event Title 1 (Ex: Visiting Team)"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1125
         Width           =   3030
      End
   End
   Begin VB.PictureBox plcAutomation 
      Height          =   5370
      Left            =   210
      ScaleHeight     =   5310
      ScaleWidth      =   9810
      TabIndex        =   0
      Top             =   270
      Width           =   9870
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Replace COM with media code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   29
         Left            =   2355
         TabIndex        =   105
         Top             =   1065
         Width           =   4260
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "RadioMan"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   135
         TabIndex        =   30
         Top             =   2085
         Width           =   1335
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Station Playlist"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   27
         Left            =   120
         TabIndex        =   38
         Top             =   3360
         Width           =   1785
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Zetta"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   120
         TabIndex        =   47
         Top             =   4635
         Width           =   1335
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Scott V5"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   25
         Left            =   1380
         TabIndex        =   36
         Top             =   2850
         Width           =   1275
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Include the Generation of Auto File w/o Spots"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   48
         Top             =   5070
         Width           =   4545
      End
      Begin VB.TextBox edcWegenerIPump 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   3930
         MaxLength       =   1
         TabIndex        =   44
         Top             =   3855
         Width           =   375
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Wegener-iPump"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   42
         Top             =   3870
         Width           =   1830
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Linkup"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   1380
         TabIndex        =   22
         Top             =   810
         Width           =   1965
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Jelli"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   8370
         TabIndex        =   49
         Top             =   2085
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Wide Orbit"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   22
         Left            =   120
         TabIndex        =   45
         Top             =   4125
         Width           =   1440
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Air"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   1380
         TabIndex        =   15
         Top             =   45
         Width           =   735
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "RPS"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   2190
         TabIndex        =   16
         Top             =   45
         Width           =   855
      End
      Begin VB.TextBox edcWegener 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   7140
         MaxLength       =   1
         TabIndex        =   41
         Top             =   3570
         Width           =   375
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "ISCI Export"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1380
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Rivendell"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   17
         Left            =   120
         TabIndex        =   34
         Top             =   2595
         Width           =   1350
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "5 Digit Cart #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   3015
         TabIndex        =   33
         Top             =   2340
         Width           =   2070
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "OLA"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   25
         Top             =   1575
         Width           =   1335
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Wegener-Compel"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   39
         Top             =   3615
         Width           =   1830
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Simian"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   37
         Top             =   3105
         Width           =   1065
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "WireReady"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   46
         Top             =   4365
         Width           =   1335
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Scott"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   35
         Top             =   2850
         Width           =   795
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "4 Digit Cart #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1380
         TabIndex        =   32
         Top             =   2340
         Width           =   2025
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "NexGen"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3015
         TabIndex        =   28
         Top             =   1830
         Width           =   1080
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Drake"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   555
         Width           =   840
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Dalet"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   795
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Wizard"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   4545
         TabIndex        =   29
         Top             =   1830
         Width           =   1110
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "MediaStar"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   1380
         TabIndex        =   27
         Top             =   1830
         Width           =   1245
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "iMediaTouch"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   1065
         Width           =   1515
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Sat"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   3165
         TabIndex        =   17
         Top             =   45
         Width           =   735
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Enco"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   21
         Top             =   810
         Width           =   1335
      End
      Begin VB.CheckBox ckcGAuto 
         Caption         =   "Include Media Definition with Audio Vault Sat"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   4050
         TabIndex        =   18
         Top             =   45
         Width           =   4260
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7605
         Y1              =   4950
         Y2              =   4950
      End
      Begin VB.Label lacWegenerIPump 
         Caption         =   "Export Time Zone"
         Height          =   210
         Left            =   2340
         TabIndex        =   43
         Top             =   3885
         Width           =   1740
      End
      Begin VB.Label lacTitle 
         Caption         =   "RCS:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   2340
         Width           =   1125
      End
      Begin VB.Label lacTitle 
         Caption         =   "Prophet:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   1830
         Width           =   1125
      End
      Begin VB.Label lacTitle 
         Caption         =   "Audio Vault:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   45
         Width           =   1125
      End
      Begin VB.Label lacWegener 
         Caption         =   "Prefix to Dynamically generated Wegener Group Names"
         Height          =   210
         Left            =   2355
         TabIndex        =   40
         Top             =   3630
         Width           =   4845
      End
   End
End
Attribute VB_Name = "SiteTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SiteTabs.ctl on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  imPopReqd                     imSelectedIndex               imComboBoxIndex           *
'*  imLastColSorted               imLastSort                    imBypassSetting           *
'*  imTypeRowNo                   lmFirstAllowedChgDate         lmEnableRow               *
'*  lmEnableCol                   lmTopRow                      imInitNoRows              *
'*                                                                                        *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mTerminate                    mSetCommands                  pbcClickFocus_Click       *
'*                                                                                        *
'*                                                                                        *
'* Public Property Procedures (Marked)                                                    *
'*  Verify(Get)                                                                           *
'*                                                                                        *
'* Public User-Defined Events (Marked)                                                    *
'*  SetSave                                                                               *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SiteTabs.ctl
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text

Public Event SetSave(ilStatus As Integer) 'VBC NR

'Program library dates Field Areas
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imLbcArrowSetting As Integer
Dim imBypassFocus As Integer
Dim imDoubleClickName As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imStartMode As Integer
Dim imIgnoreClickEvent As Integer

Dim smNowDate As String
Dim lmNowDate As Long

Dim imCtrlVisible As Integer
Public Enum ProtectChanges
Start = -1
Done = 0
End Enum

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Exit Sub
    End If
    imFirstActivate = False
    imUpdateAllowed = igUpdateAllowed
    'If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
    If Not imUpdateAllowed Then
    Else
    End If
    gShowBranner imUpdateAllowed
End Sub

Private Sub Form_Click()
End Sub

Private Sub Form_Deactivate()
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         slNameCode                    slName                    *
'*  slCode                        ilLoop                        slDaypart                 *
'*  slLineNo                      slStr                                                   *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mInitErr                                                                              *
'******************************************************************************************

'
'   mInit
'   Where:
'

    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    imBypassFocus = False
    imSettingValue = False
    imStartMode = True
    imChgMode = False
    imBSMode = False
    imLbcArrowSetting = False
    imLbcMouseDown = False
    imTabDirection = 0  'Left to right movement
    imDoubleClickName = False
    imCtrlVisible = False
    imCtrlVisible = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    mInitBox
    plcSports.BorderStyle = 0
    plcSports.Visible = False
    plcSports.Move 0, 0
    plcAutomation.BorderStyle = 0
    plcAutomation.Visible = False
    plcAutomation.Move 0, 0
    plcResearch.BorderStyle = 0
    plcResearch.Visible = False
    plcResearch.Move 0, 0
    frcEMail.BorderStyle = 0
    frcEMail.Visible = False
    frcEMail.Move 0, 0
    'dan 6/28/11 no longer using from name or address; add tls
    edcFromName.Visible = False
    'lblFromAddress.Visible = False
    edcFromAddress.Visible = False
    lblFrom.Caption = "Transport Layer Security (start with false and test with verify)"
    chkTLS.Visible = True
    chkTLS.Left = edcFromName.Left
    chkTLS.Top = edcFromName.Top
    edcVerifyTo.Left = edcFromAddress.Left
    edcVerifyTo.Top = edcFromAddress.Top
    edcVerifyTo.Width = edcFromAddress.Width
    edcVerifyTo.height = edcFromAddress.height
    edcVerifyTo.Visible = True
    lblFromAddress.Caption = "To address for verification."
    frcRepNet.BorderStyle = 0
    frcRepNet.Visible = False
    frcRepNet.Move 0, 0

    'Parent/Child: "iMediaTouch" must be checked for "Replace COM with Media Code" option
    If ckcGAuto(7).Value = vbUnchecked Then
        ckcGAuto(29).Value = vbUnchecked
        ckcGAuto(29).Enabled = False 'iMediaTouch - Replace COM with Media Code
    Else
        ckcGAuto(29).Enabled = True 'iMediaTouch - Replace COM with Media Code
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr: 'VBC NR
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  flTextHeight                  ilLoop                        llRow                     *
'*  ilCol                                                                                 *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    'flTextHeight = pbcDates.TextHeight("1") - 35

End Sub









Public Sub Action(ilType As Integer, ilState As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilIndex As Integer
    Select Case ilType
        Case 1  'Clear Focus
        Case 2  'Init function
            'Test if unloading control
            ilRet = 0
            On Error GoTo UserControlErr:
            ilIndex = SiteOpt!tbcSelection.SelectedItem.Index
            If ilRet = 0 Then
                Form_Load
                Form_Activate
                Select Case ilIndex
                    Case 1  'General
                    Case 2  'Sales
                    Case 3  'Commission
                    Case 4  'Reserach
                        plcResearch.Visible = True
                        plcResearch.Enabled = ilState
                    Case 5  'Schedule
                    Case 6  'Agency/Advertiser
                    Case 7  'Contract
                    Case 8  'Copy
                    Case 9  'Log
                    Case 10  'Invoice
                    Case 11  'Accounting
                    Case 12  'Backup
                    Case 13 'Options
                    Case 14 'Sports
                        plcSports.Visible = True
                        plcSports.Enabled = ilState
                    Case 15 'Automation
                        plcAutomation.Visible = True
                        plcAutomation.Enabled = ilState
                    Case 16 'Comments
                    'Dan 11/04/09 email removed from site options. Rep-Net now case 16
'                    Case 16 'E-Mail
''                        frcEMail.Visible = True
''                        frcEMail.Enabled = ilState
''                    Case 17
'                        frcRepNet.Visible = True
'                        frcRepNet.Enabled = ilState
                    Case 17
                        frcRepNet.Visible = True
                        frcRepNet.Enabled = ilState
                    'Dan 2/07/2011 re-enable email in site options
                    Case 18
                        frcEMail.Visible = True
                        frcEMail.Enabled = ilState
                End Select
            End If
        Case 3  'terminate function
        Case 4  'Clear
            Screen.MousePointer = vbDefault
        Case 5  'Save
            plcAutomation.Visible = False
            plcSports.Visible = False
            plcResearch.Visible = False
            frcEMail.Visible = False
            frcRepNet.Visible = False
        Case 6
            If ilState = 0 Then
                frcRepNet.Visible = False
            ElseIf ilState = 1 Then
                frcRepNet.Visible = True
                lacDBID.Visible = True
                edcDBID.Visible = True
                frcNet.Visible = False
            Else
                frcRepNet.Visible = True
                lacDBID.Visible = True
                edcDBID.Visible = True
                frcNet.Visible = True
            End If
        Case 7
            If ilState = 0 Then
                imIgnoreClickEvent = True
            Else
                imIgnoreClickEvent = False
            End If
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




Private Sub chkTLS_Click()
    If Not SiteOpt.bmIgnoreChange Then
        mChangeOccured
    End If
End Sub

Private Sub chkTLS_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcGAuto_Click(Index As Integer)
    Dim ilPasswordOk As Integer
    If Index = 7 Then
        If ckcGAuto(7).Value = vbUnchecked Then
            ckcGAuto(29).Value = vbUnchecked
            ckcGAuto(29).Enabled = False 'iMediaTouch - Replace COM with Media Code
        Else
            ckcGAuto(29).Enabled = True 'iMediaTouch - Replace COM with Media Code
        End If
    End If
    If Index = 29 Then 'iMediaTouch - Replace COM with Media Code
        'Dont Prompt for PW
    Else
        'If (Index = 13) Or (Index = 14) Or (Index = 15) Then
        If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
            If igPasswordOk Then
                If ckcGAuto(Index).Value = vbChecked Then
                    ilPasswordOk = igPasswordOk
                    'If Index = 13 Then
                    '    sgPasswordAddition = "XD-"
                    'ElseIf Index = 14 Then
                    '    sgPasswordAddition = "WE-"
                    'Else
                    '    sgPasswordAddition = "OE-"
                    'End If
                    sgPasswordAddition = "AE-"
                    mProtectChangesAllowed Start
                    CSPWord.Show vbModal
                    mProtectChangesAllowed Done
                    If Not igPasswordOk Then
                        ckcGAuto(Index).Value = vbUnchecked
                    End If
                    sgPasswordAddition = ""
                    igPasswordOk = ilPasswordOk
                End If
            Else
                ckcGAuto(Index).Value = vbUnchecked
            End If
        End If
        'End If
    End If
    
    If (imIgnoreClickEvent = False) And ((Index = 14) Or (Index = 15)) Then
        If (ckcGAuto(14).Value = vbChecked) Or (ckcGAuto(15).Value = vbChecked) Then
            'Disallow Mix length because Wegener and OLA are a one to one replacement
            SiteOpt!ckcRegionMixLen.Value = vbUnchecked
            SiteOpt!ckcRegionMixLen.Enabled = False
        Else
            If SiteOpt!ckcUsingSplitNetworks.Value = vbChecked Then
                SiteOpt!ckcRegionMixLen.Enabled = True
            End If
        End If
    End If
    If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
    End If
    
End Sub

Private Sub ckcHiddenOverride_Click()
    If SiteOpt!ckcUsingSpecialResearch.Value = vbChecked Then
        SiteOpt!edcComment(3).Visible = True
        SiteOpt!lacComment(3).Caption = "Research Estimate Comment"
    ElseIf ckcHiddenOverride.Value = vbChecked Then
        SiteOpt!edcComment(3).Visible = True
        SiteOpt!lacComment(3).Caption = "Research Override Comment"
    Else
        SiteOpt!edcComment(3).Visible = False
        SiteOpt!lacComment(3).Caption = ""
    End If
    If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
    End If

End Sub
Private Sub cmdVerifyByCSI_Click()
    '8272
    Set ogEmailer = New CEmail
    
    edcVerifyEmail.ForeColor = vbBlack
    edcVerifyEmail.Text = "attempting to send test message with CSI values...please wait."
    Screen.MousePointer = vbHourglass
    With ogEmailer
        .Message = "Email is working correctly."
        .Subject = "verify email"
        .ToAddress = Trim$(edcVerifyTo.Text)
        .FromAddress = CSIUSERNAME '"testVerify@csi.net"  ' TTP 10837 jjb
        .FromName = "csi site test"
        '10564 also 10564, changed tls to constant
'        .SetHost "smtpauth.hosting.earthlink.net", 587, "emailSend@counterpoint.net", "TestMyEmail1", False
        '.SetHost CSISITE, CSIPORT, CSIUSERNAME, CSIPASSWORD, True
        .SetHost CSISITE, CSIPORT, CSIUSERNAME, CSIPASSWORD, CSITLS
    End With
    If ogEmailer.Send(edcVerifyEmail) Then
        If InStr(1, edcVerifyEmail.Text, "sent", vbTextCompare) > 0 Then
            edcVerifyEmail.Text = "Verified"
        End If
    Else
        If InStr(1, edcVerifyEmail.Text, "11004", vbTextCompare) > 0 Then
            edcVerifyEmail.Text = "Host name not recognized."
        ElseIf InStr(1, edcVerifyEmail.Text, "535") > 0 Then
            edcVerifyEmail.Text = "Password not correct."
        ElseIf InStr(1, edcVerifyEmail.Text, "454") > 0 Then
            edcVerifyEmail.Text = "acount name not recognized."
        ElseIf InStr(1, edcVerifyEmail.Text, "Time", vbTextCompare) > 0 Then
            edcVerifyEmail.Text = "Session timed out.  Port setting may not be correct."
        ElseIf InStr(1, edcVerifyEmail.Text, "18") > 0 Then
            edcVerifyEmail.Text = "Transcript Layer Security must be set to false."
        End If
    End If
   Set ogEmailer = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdVerifyEmail_Click()
   ' Dim myMail  As MailSender
    'Dim tlEmailInfo As EmailInformation
    'Const MAILREGKEY = "16374-29451-54460"
    Set ogEmailer = New CEmail
    Dim blTls As Boolean
    
    edcVerifyEmail.ForeColor = vbBlack
    edcVerifyEmail.Text = "attempting to send test message...please wait."
    Screen.MousePointer = vbHourglass
   ' Set myMail = New MailSender
   ' myMail.RegKey = MAILREGKEY
    
   ' If (LenB(edcHost.Text) = 0) Or (LenB(edcFromAddress.Text) = 0) Or (LenB(edcPassword.Text) = 0) Or (LenB(edcAcctName.Text) = 0) Or (LenB(edcFromName.Text) = 0) Or (LenB(edcPort.Text) = 0) Then
    If (LenB(edcHost.Text) = 0) Or (LenB(edcPassword.Text) = 0) Or (LenB(edcAcctName.Text) = 0) Or (LenB(edcPort.Text) = 0) Or (LenB(edcVerifyTo.Text) = 0) Then
        edcVerifyEmail.Text = "A field is blank.  Cannot verify."
        Screen.MousePointer = vbDefault
       ' Set myMail = Nothing
        Set ogEmailer = Nothing
        Exit Sub
    End If
    If chkTLS.Value = vbChecked Then
        blTls = True
    Else
        blTls = False
    End If
    With ogEmailer
        .Message = "Email is working correctly."
        .Subject = "verify email"
        .ToAddress = Trim$(edcVerifyTo.Text)
        .FromAddress = edcAcctName.Text '"testVerify@csi.net" ' TTP 10837 jjb
        .FromName = "csi site test"
        .SetHost Trim$(edcHost.Text), Trim$(edcPort.Text), Trim$(edcAcctName.Text), Trim$(edcPassword.Text), blTls
    End With
'    tlEmailInfo.bTLSSet = True
'    tlEmailInfo.sMessage = "Email is working correctly."
'    tlEmailInfo.sSubject = "verify email"
'    tlEmailInfo.sToAddress = Trim$(edcVerifyTo.Text)
'    With myMail
''        .From = edcFromAddress.Text
''        .FromName = edcFromName.Text
'        .From = "testVerify@csi.net"
'        .FromName = "csi site test"
'       'TODO allow to send to themselves?
'       ' .AddAddress "danmichaelson@counterpoint.net"
'       ' .Subject = "verify email"
'       ' .Body = "Email is working correctly."
'        .Host = Trim$(edcHost.Text)
'        .Password = Trim$(edcPassword.Text)
'        .Username = Trim$(edcAcctName.Text)
'        .Port = Trim$(edcPort.Text)
'        If chkTLS.Value = 1 Then
'            .TLS = True
'        Else
'            .TLS = False
'        End If
'       ' todo is this temporary?  google must have this set to true. most do not require it, and some will fail if it is set to true (earthlink)
'        If StrComp(.Host, "smtp.gmail.com", vbTextCompare) = 0 Then
'            .TLS = True
'        Else
'            .TLS = False
'        End If
'    End With
    ' I use imUpdateAllowed to not allow saving these values unless verified.
'    If gSendEmail(tlEmailInfo, edcVerifyEmail, , myMail) Then
    If ogEmailer.Send(edcVerifyEmail) Then
        If InStr(1, edcVerifyEmail.Text, "sent", vbTextCompare) > 0 Then
            edcVerifyEmail.Text = "Verified"
        End If
    Else
        If InStr(1, edcVerifyEmail.Text, "11004", vbTextCompare) > 0 Then
            edcVerifyEmail.Text = "Host name not recognized."
        ElseIf InStr(1, edcVerifyEmail.Text, "535") > 0 Then
            edcVerifyEmail.Text = "Password not correct."
        ElseIf InStr(1, edcVerifyEmail.Text, "454") > 0 Then
            edcVerifyEmail.Text = "acount name not recognized."
        ElseIf InStr(1, edcVerifyEmail.Text, "Time", vbTextCompare) > 0 Then
            edcVerifyEmail.Text = "Session timed out.  Port setting may not be correct."
        ElseIf InStr(1, edcVerifyEmail.Text, "18") > 0 Then
            edcVerifyEmail.Text = "Transcript Layer Security must be set to false."
        End If
    End If
   ' cmdExit.Caption = "Done"
   ' Set myMail = Nothing
   Set ogEmailer = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub edcWegener_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcWegenerIPump_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Form_MouseUp Button, Shift, X, Y
End Sub



Public Property Get Verify() As Integer 'VBC NR
    If imUpdateAllowed Then 'VBC NR
        Verify = True 'VBC NR
    Else 'VBC NR
        Verify = True 'VBC NR
    End If 'VBC NR
End Property 'VBC NR



Public Property Let Automation(ilIndex As Integer, ilValue As Integer)
    ckcGAuto(ilIndex).Value = ilValue
End Property
Public Property Get Automation(ilIndex As Integer) As Integer
    Automation = ckcGAuto(ilIndex).Value
End Property
Public Property Get Sports(ilIndex As Integer) As Integer
    Sports = ckcSportsInfo(ilIndex).Value
End Property
Public Property Let Sports(ilIndex As Integer, ilValue As Integer)
    ckcSportsInfo(ilIndex).Value = ilValue
End Property
Public Property Get Research(ilIndex As Integer) As Integer
    Select Case ilIndex
        Case 1  'rbcSAudData(0)
            Research = rbcSAudData(0).Value
        Case 2  'rbcSAudData(1)
            Research = rbcSAudData(1).Value
        Case 3  'rbcSAudData(2)
            Research = rbcSAudData(2).Value
        Case 4
            Research = rbcSAudData(3).Value
        Case 11
            Research = rbcSGRPCPPCal(0).Value
        Case 12
            Research = rbcSGRPCPPCal(1).Value
        Case 13
            Research = rbcSGRPCPPCal(2).Value
        Case 21
            Research = rbcWeight(0).Value
        Case 22
            Research = rbcWeight(1).Value
        Case 23
            Research = rbcWeight(2).Value
        Case 31
            Research = ckcHiddenOverride.Value
        Case 41
            Research = ckcRschCustDemo.Value
        Case 42
            Research = ckcHideDemo.Value
        Case 43
            Research = ckcAudByPackage.Value
    End Select
End Property
Public Property Let Research(ilIndex As Integer, ilValue As Integer)
    Select Case ilIndex
        Case 1  'rbcSAudData(0)
            rbcSAudData(0).Value = ilValue
        Case 2  'rbcSAudData(1)
            rbcSAudData(1).Value = ilValue
        Case 3  'rbcSAudData(2)
            rbcSAudData(2).Value = ilValue
        Case 4
            rbcSAudData(3).Value = ilValue
        Case 11
            rbcSGRPCPPCal(0).Value = ilValue
        Case 12
            rbcSGRPCPPCal(1).Value = ilValue
        Case 13
            rbcSGRPCPPCal(2).Value = ilValue
        Case 21
            rbcWeight(0).Value = ilValue
        Case 22
            rbcWeight(1).Value = ilValue
        Case 23
            rbcWeight(2).Value = ilValue
        Case 31
            ckcHiddenOverride.Value = ilValue
        Case 41
            ckcRschCustDemo.Value = ilValue
        Case 42
            ckcHideDemo.Value = ilValue
        Case 43
            ckcAudByPackage.Value = ilValue
    End Select
End Property

Public Property Get Email(ilIndex As Integer) As String
    Select Case ilIndex
        Case 1
            Email = edcHost.Text
        Case 2
            Email = edcAcctName.Text
        Case 3
            Email = edcPassword.Text
        Case 4
            Email = edcPort.Text
        Case 5
'            Email = edcFromName.Text
            Email = chkTLS.Value
'        Case 6
'            Email = edcFromAddress.Text
    End Select
End Property

Public Property Let Email(ilIndex As Integer, slvalue As String)
    Select Case ilIndex
        Case 1
            edcHost.Text = slvalue
        Case 2
            edcAcctName.Text = slvalue
        Case 3
            edcPassword.Text = slvalue
        Case 4
            edcPort.Text = slvalue
        Case 5
            chkTLS.Value = slvalue
            'edcFromName.Text = slValue
'            edcFromName.Text = slValue
'        Case 6
'            edcFromAddress.Text = slValue
    End Select
End Property

Public Property Get RepNet(ilIndex As Integer) As String
    Select Case ilIndex
        Case 1
            RepNet = edcDBID.Text
        Case 2
            RepNet = edcFTPUserID.Text
        Case 3
            RepNet = edcFTPUserPW.Text
        Case 4
            RepNet = edcFTPPort.Text
        Case 5
            RepNet = edcFTPAddress.Text
        Case 6
            RepNet = edcFTPImport.Text
        Case 7
            RepNet = edcFTPExport.Text
        Case 8
            RepNet = edcIISRootURL.Text
        Case 9
            RepNet = edcIISRegSection.Text
    End Select
End Property

Public Property Let RepNet(ilIndex As Integer, slvalue As String)
    Select Case ilIndex
        Case 1
            edcDBID.Text = slvalue
        Case 2
            edcFTPUserID.Text = slvalue
        Case 3
            edcFTPUserPW.Text = slvalue
        Case 4
            edcFTPPort.Text = slvalue
        Case 5
            edcFTPAddress.Text = slvalue
        Case 6
            edcFTPImport.Text = slvalue
        Case 7
            edcFTPExport.Text = slvalue
        Case 8
            edcIISRootURL.Text = slvalue
        Case 9
            edcIISRegSection.Text = slvalue
    End Select
End Property


Private Sub plcWeight_Paint()
    plcWeight.CurrentX = 0
    plcWeight.CurrentY = -30
    plcWeight.Print "For Non-Matching Dayparts Give Weight To"
End Sub

Private Sub plcSGRPCPPCal_Paint()
    plcSGRPCPPCal.CurrentX = 0
    plcSGRPCPPCal.CurrentY = -30
    plcSGRPCPPCal.Print "GRP/CPP Calculations by"
End Sub
Private Sub plcSAudData_Paint()
    plcSAudData.CurrentX = 0
    plcSAudData.CurrentY = -30
    plcSAudData.Print "Audience Data in"
End Sub
'auto generated code Dan M 4/23/09

Private Sub ckcRschCustDemo_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub ckcHideDemo_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub ckcAudByPackage_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcSGRPCPPCal_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcSAudData_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcWeight_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcHiddenOverride_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcSportsInfo_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcGAuto_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcFTPUserID_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcFTPUserPW_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcFTPPort_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcFTPAddress_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcFTPImport_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcFTPExport_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcIISRootURL_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcIISRegSection_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcDBID_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcHost_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcPort_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAcctName_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcPassword_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcFromName_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcFromAddress_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcRschCustDemo_Click()
    If Not SiteOpt.bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
Private Sub ckcHideDemo_Click()
    If Not SiteOpt.bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
Private Sub ckcAudByPackage_Click()
    If Not SiteOpt.bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
Private Sub rbcSGRPCPPCal_Click(Index As Integer)
    If Not SiteOpt.bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcSAudData_Click(Index As Integer)
    If Not SiteOpt.bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcWeight_Click(Index As Integer)
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub


Private Sub ckcSportsInfo_Click(Index As Integer)
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcFTPUserID_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcFTPUserPW_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcFTPPort_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcFTPAddress_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcFTPImport_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcFTPExport_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcIISRootURL_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcIISRegSection_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcDBID_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcHost_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcPort_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcAcctName_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcPassword_Change()
If Not SiteOpt.bmIgnoreChange Then
      mChangeOccured
End If
End Sub

'Private Sub edcFromName_Change()
'If Not SiteOpt.bmIgnoreChange Then
'      mChangeOccured
'End If
'End Sub
'
'Private Sub edcFromAddress_Change()
'If Not SiteOpt.bmIgnoreChange Then
'      mChangeOccured
'End If
'End Sub
Private Sub mCtrlGotFocusAndIgnoreChange(Ctrl As control)
'   Dan M 4/22/09 hijacked gCtrlGotFocus and added siteOpt.bmignoreChange to help control textboxes and counting changes
    SiteOpt.bmIgnoreChange = False
   ' gCtrlGotFocus Ctrl
End Sub
Public Sub mChangeOccured()
    SiteOpt.mChangeOccured
End Sub
Private Sub mProtectChangesAllowed(blStart As ProtectChanges)
    SiteOpt.mProtectChangesAllowed blStart
'Static isSaveChangesAllowed As Integer
'    If blStart Then
'        isSaveChangesAllowed = igChangesAllowed
'    Else
'        igChangesAllowed = isSaveChangesAllowed
'        If Not igPasswordOk Then   'And imChangesOccured > 1
'            imChangesOccured = imChangesOccured - 1
'        End If
'    End If
End Sub

Public Property Get Wegener() As String
    Wegener = edcWegener.Text
End Property
Public Property Let Wegener(slvalue As String)
    edcWegener.Text = slvalue
End Property
Public Property Get WegenerIPump() As String
    WegenerIPump = edcWegenerIPump.Text
End Property
Public Property Let WegenerIPump(slvalue As String)
    edcWegenerIPump.Text = slvalue
End Property
Public Property Get EventTitle(ilIndex As Integer) As String
    Select Case ilIndex
        Case 1
            EventTitle = edcEventTitle(ilIndex - 1).Text
        Case 2
            EventTitle = edcEventTitle(ilIndex - 1).Text
    End Select
End Property

Public Property Let EventTitle(ilIndex As Integer, slvalue As String)
    Select Case ilIndex
        Case 1
            edcEventTitle(ilIndex - 1).Text = slvalue
        Case 2
            edcEventTitle(ilIndex - 1).Text = slvalue
    End Select
End Property
Public Property Get EventSubtotalTitle(ilIndex As Integer) As String
    Select Case ilIndex
        Case 1
            EventSubtotalTitle = edcEventSubtotal(ilIndex - 1).Text
        Case 2
            EventSubtotalTitle = edcEventSubtotal(ilIndex - 1).Text
    End Select
End Property

Public Property Let EventSubtotalTitle(ilIndex As Integer, slvalue As String)
    Select Case ilIndex
        Case 1
            edcEventSubtotal(ilIndex - 1).Text = slvalue
        Case 2
            edcEventSubtotal(ilIndex - 1).Text = slvalue
    End Select
End Property
