VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSiteOptions 
   Caption         =   "Site Options"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   Icon            =   "frmSiteOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9510
   Visible         =   0   'False
   Begin VB.Frame frcTab 
      Caption         =   "Not Shown"
      Height          =   3960
      Index           =   5
      Left            =   7815
      TabIndex        =   103
      Top             =   6885
      Visible         =   0   'False
      Width           =   8460
      Begin VB.TextBox edcFromAddress 
         Height          =   285
         Left            =   360
         MaxLength       =   80
         TabIndex        =   12
         Top             =   3480
         Width           =   7935
      End
      Begin VB.TextBox edcFromName 
         Height          =   285
         Left            =   330
         MaxLength       =   80
         TabIndex        =   10
         Top             =   2760
         Width           =   7935
      End
      Begin VB.TextBox edcPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   330
         MaxLength       =   80
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1920
         Width           =   4095
      End
      Begin VB.TextBox edcAcctName 
         Height          =   285
         Left            =   360
         MaxLength       =   80
         TabIndex        =   4
         Top             =   1170
         Width           =   7695
      End
      Begin VB.TextBox edcPort 
         Height          =   285
         Left            =   5370
         MaxLength       =   5
         TabIndex        =   8
         Top             =   1935
         Width           =   1095
      End
      Begin VB.TextBox edcHost 
         Height          =   285
         Left            =   330
         MaxLength       =   80
         TabIndex        =   2
         Top             =   435
         Width           =   7695
      End
      Begin VB.Label lblFromAddress 
         Caption         =   "Optional - From Address, e.g., admin@xyzradionetworks.net"
         Height          =   255
         Left            =   330
         TabIndex        =   11
         Top             =   3240
         Width           =   5415
      End
      Begin VB.Label lblUserName 
         Caption         =   "Account Name, e.g., abcd@xyz.com"
         Height          =   255
         Left            =   330
         TabIndex        =   3
         Top             =   915
         Width           =   4815
      End
      Begin VB.Label lblHost 
         Caption         =   "SMTP Server, e.g., smtp.att.yahoo.com"
         Height          =   255
         Left            =   330
         TabIndex        =   1
         Top             =   180
         Width           =   4815
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   330
         TabIndex        =   5
         Top             =   1665
         Width           =   975
      End
      Begin VB.Label lblPort 
         Caption         =   "Port Number, e.g., 25"
         Height          =   255
         Left            =   5370
         TabIndex        =   7
         Top             =   1680
         Width           =   2130
      End
      Begin VB.Label lblFrom 
         Caption         =   "Optional - From Name, e.g., XYZ Radio Networks IT dept."
         Height          =   255
         Left            =   330
         TabIndex        =   9
         Top             =   2415
         Width           =   5415
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "Options"
      Height          =   4875
      Index           =   0
      Left            =   8880
      TabIndex        =   13
      Top             =   5400
      Width           =   8850
      Begin VB.CheckBox optUsingWeb_N 
         Caption         =   "Using Network Web Site"
         Height          =   240
         Left            =   6240
         TabIndex        =   157
         Top             =   4440
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.TextBox edcBand 
         Height          =   285
         Left            =   2625
         TabIndex        =   45
         Top             =   2160
         Width           =   4575
      End
      Begin VB.CheckBox ckcUsingServiceAgreement 
         Caption         =   "Using Service Agreements"
         Height          =   195
         Left            =   240
         TabIndex        =   59
         Top             =   4485
         Width           =   3180
      End
      Begin VB.Frame frcRADARMultiAir 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         TabIndex        =   50
         Top             =   2850
         Width           =   8205
         Begin VB.OptionButton rbcRADARMultiAir 
            Caption         =   "by Air Time"
            Height          =   195
            Index           =   2
            Left            =   5985
            TabIndex        =   54
            Top             =   0
            Width           =   1275
         End
         Begin VB.OptionButton rbcRADARMultiAir 
            Caption         =   "Multi-AirPlay by Program Code"
            Height          =   195
            Index           =   1
            Left            =   3420
            TabIndex        =   53
            Top             =   0
            Width           =   2535
         End
         Begin VB.OptionButton rbcRADARMultiAir 
            Caption         =   "Multi-AirPlay by Spot ID"
            Height          =   195
            Index           =   0
            Left            =   1305
            TabIndex        =   52
            Top             =   0
            Width           =   2025
         End
         Begin VB.Label lacRADARMultiAir 
            Caption         =   "RADAR Export:"
            Height          =   225
            Left            =   0
            TabIndex        =   51
            Top             =   -15
            Width           =   1260
         End
      End
      Begin VB.CheckBox ckcDefaultEstDay 
         Caption         =   "Default to Estimate Day (blank if unchecked)"
         Height          =   240
         Left            =   240
         TabIndex        =   58
         Top             =   4155
         Width           =   5070
      End
      Begin VB.CheckBox optXDS 
         Caption         =   "Support X-Digital Automatic Playback Feature"
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   56
         Top             =   3495
         Width           =   4695
      End
      Begin VB.CheckBox optXDS 
         Caption         =   "Vehicle Information"
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   49
         Top             =   2505
         Width           =   2355
      End
      Begin VB.CheckBox optXDS 
         Caption         =   "Generate the Transparent File (If checked, it will be created as part of XDS Spot Insertion)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   55
         Top             =   3165
         Width           =   8565
      End
      Begin VB.CheckBox optMissed 
         Caption         =   "Include Missed Dates/Time for Missed Spots on Reports"
         Height          =   240
         Left            =   240
         TabIndex        =   57
         Top             =   3825
         Width           =   5595
      End
      Begin VB.CheckBox optXDS 
         Caption         =   "Agreement Information"
         Height          =   255
         Index           =   1
         Left            =   3690
         TabIndex        =   48
         Top             =   2505
         Width           =   2310
      End
      Begin VB.CheckBox optXDS 
         Caption         =   "Station Information"
         Height          =   255
         Index           =   0
         Left            =   1635
         TabIndex        =   47
         Top             =   2505
         Width           =   1995
      End
      Begin VB.Frame frcUsingID 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   240
         TabIndex        =   39
         Top             =   1845
         Width           =   8265
         Begin VB.OptionButton rbcUsingID 
            Caption         =   "No: Import in Add mode"
            Height          =   210
            Index           =   2
            Left            =   6015
            TabIndex        =   43
            Top             =   0
            Width           =   2355
         End
         Begin VB.OptionButton rbcUsingID 
            Caption         =   "Yes"
            Height          =   210
            Index           =   0
            Left            =   2490
            TabIndex        =   41
            Top             =   0
            Width           =   885
         End
         Begin VB.OptionButton rbcUsingID 
            Caption         =   "No: Import in Update mode"
            Height          =   210
            Index           =   1
            Left            =   3405
            TabIndex        =   42
            Top             =   0
            Width           =   2595
         End
         Begin VB.Label lacUsingID 
            Caption         =   "Using Station ID"
            Height          =   210
            Left            =   0
            TabIndex        =   40
            Top             =   0
            Width           =   1755
         End
      End
      Begin VB.Frame frcUsingViero 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         TabIndex        =   35
         Top             =   1425
         Width           =   5280
         Begin VB.OptionButton rbcUsingViero 
            Caption         =   "No"
            Height          =   375
            Index           =   1
            Left            =   4080
            TabIndex        =   38
            Top             =   0
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton rbcUsingViero 
            Caption         =   "Yes"
            Height          =   375
            Index           =   0
            Left            =   3210
            TabIndex        =   37
            Top             =   0
            Width           =   885
         End
         Begin VB.Label lacUsingViero 
            Caption         =   "Using Viero Transact Enterprise"
            Height          =   375
            Left            =   0
            TabIndex        =   36
            Top             =   90
            Width           =   3225
         End
      End
      Begin VB.Frame frcShowAgreementDates 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   180
         Width           =   5790
         Begin VB.OptionButton rbcShowAgreementDates 
            Caption         =   "Yes"
            Height          =   375
            Index           =   0
            Left            =   3630
            TabIndex        =   17
            Top             =   0
            Width           =   885
         End
         Begin VB.OptionButton rbcShowAgreementDates 
            Caption         =   "No"
            Height          =   375
            Index           =   1
            Left            =   4530
            TabIndex        =   18
            Top             =   0
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.Label lacShowAgreementDates 
            Caption         =   "Allow Agreement Start, End and Drop Dates"
            Height          =   375
            Left            =   0
            TabIndex        =   16
            Top             =   90
            Width           =   3570
         End
      End
      Begin VB.CheckBox optMultiVehPosting 
         Caption         =   "Allow Multi-Vehicle Affidavit Posting"
         Height          =   255
         Left            =   5280
         TabIndex        =   60
         Top             =   600
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Frame frcRCS5 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   4440
         TabIndex        =   31
         Top             =   1095
         Width           =   5625
         Begin VB.OptionButton rbcRCS5 
            Caption         =   "No"
            Height          =   210
            Index           =   1
            Left            =   3405
            TabIndex        =   34
            Top             =   90
            Width           =   870
         End
         Begin VB.OptionButton rbcRCS5 
            Caption         =   "Yes"
            Height          =   210
            Index           =   0
            Left            =   2490
            TabIndex        =   33
            Top             =   90
            Width           =   870
         End
         Begin VB.Label lacRCS5 
            Caption         =   "Export to RCS 5 Digit Cart #"
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   90
            Width           =   2370
         End
      End
      Begin VB.Frame frcRCS4 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   240
         TabIndex        =   27
         Top             =   1095
         Width           =   5625
         Begin VB.OptionButton rbcRCS4 
            Caption         =   "Yes"
            Height          =   210
            Index           =   0
            Left            =   2490
            TabIndex        =   29
            Top             =   90
            Width           =   870
         End
         Begin VB.OptionButton rbcRCS4 
            Caption         =   "No"
            Height          =   210
            Index           =   1
            Left            =   3405
            TabIndex        =   30
            Top             =   90
            Width           =   870
         End
         Begin VB.Label lacRCS4 
            Caption         =   "Export to RCS 4 Digit Cart #"
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   90
            Width           =   2355
         End
      End
      Begin VB.Frame frcVehType 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   4905
         Begin VB.OptionButton rbcVehType 
            Caption         =   "Yes"
            Height          =   375
            Index           =   0
            Left            =   2490
            TabIndex        =   21
            Top             =   15
            Width           =   885
         End
         Begin VB.OptionButton rbcVehType 
            Caption         =   "No"
            Height          =   375
            Index           =   1
            Left            =   3405
            TabIndex        =   22
            Top             =   15
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.Label lacVehType 
            Caption         =   "Show by Vehicle Type"
            Height          =   375
            Left            =   0
            TabIndex        =   20
            Top             =   90
            Width           =   2865
         End
      End
      Begin VB.Frame frcVehStn 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Top             =   795
         Width           =   5625
         Begin VB.OptionButton rbcVehStn 
            Caption         =   "No"
            Height          =   210
            Index           =   1
            Left            =   3405
            TabIndex        =   26
            Top             =   90
            Width           =   870
         End
         Begin VB.OptionButton rbcVehStn 
            Caption         =   "Yes"
            Height          =   210
            Index           =   0
            Left            =   2490
            TabIndex        =   25
            Top             =   90
            Width           =   870
         End
         Begin VB.Label lbcVehStn 
            Caption         =   "Are Vehicles also Stations"
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   90
            Width           =   2415
         End
      End
      Begin VB.CheckBox optUsingUnivision 
         Caption         =   "Using Univision"
         Height          =   255
         Left            =   6960
         TabIndex        =   14
         Top             =   210
         Width           =   1860
      End
      Begin VB.Label lacBandNote 
         Caption         =   "(Separate Bands by Commas)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   7245
         TabIndex        =   155
         Top             =   2175
         Width           =   1500
      End
      Begin VB.Label lacBand 
         Caption         =   "Station Import Additional Bands"
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   2175
         Width           =   2760
      End
      Begin VB.Label lacXDS 
         Caption         =   "Send to X-Digital"
         Height          =   225
         Left            =   240
         TabIndex        =   46
         Top             =   2505
         Width           =   1725
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "Message"
      Height          =   4590
      Index           =   1
      Left            =   600
      TabIndex        =   61
      Top             =   960
      Width           =   8445
      Begin VB.ListBox lbcVehicles 
         Height          =   2310
         ItemData        =   "frmSiteOptions.frx":08CA
         Left            =   210
         List            =   "frmSiteOptions.frx":08CC
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   68
         Top             =   885
         Width           =   2205
      End
      Begin VB.TextBox edcNCRWks 
         Height          =   285
         Left            =   7215
         MaxLength       =   2
         TabIndex        =   73
         Top             =   3750
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox edcOMMindate 
         Height          =   285
         Left            =   6390
         MaxLength       =   10
         TabIndex        =   74
         Top             =   4140
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox edcOverdue 
         Height          =   285
         Left            =   7215
         MaxLength       =   2
         TabIndex        =   71
         Top             =   3345
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Frame frcMessage 
         Height          =   570
         Left            =   285
         TabIndex        =   62
         Top             =   195
         Width           =   8100
         Begin VB.OptionButton rbcMessage 
            Caption         =   "Missed"
            Height          =   300
            Index           =   5
            Left            =   6960
            TabIndex        =   156
            Top             =   120
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton rbcMessage 
            Caption         =   "Overdue"
            Height          =   300
            Index           =   3
            Left            =   5745
            TabIndex        =   67
            Top             =   150
            Width           =   1110
         End
         Begin VB.OptionButton rbcMessage 
            Caption         =   "Password"
            Height          =   300
            Index           =   2
            Left            =   4455
            TabIndex        =   66
            Top             =   150
            Width           =   1185
         End
         Begin VB.OptionButton rbcMessage 
            Caption         =   "Welcome by Vehicle"
            Height          =   300
            Index           =   4
            Left            =   2430
            TabIndex        =   65
            Top             =   150
            Width           =   2055
         End
         Begin VB.OptionButton rbcMessage 
            Caption         =   "Welcome"
            Height          =   300
            Index           =   1
            Left            =   1215
            TabIndex        =   64
            Top             =   150
            Width           =   1185
         End
         Begin VB.OptionButton rbcMessage 
            Caption         =   "Citation"
            Height          =   300
            Index           =   0
            Left            =   90
            TabIndex        =   63
            Top             =   120
            Width           =   1170
         End
      End
      Begin VB.TextBox edcMessage 
         Height          =   2310
         Left            =   2700
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   69
         Top             =   885
         Width           =   5415
      End
      Begin VB.Label lacNCRWks 
         Caption         =   "Number of Consecutive Weeks behind considered Critical"
         Height          =   315
         Left            =   2280
         TabIndex        =   72
         Top             =   3780
         Visible         =   0   'False
         Width           =   4320
      End
      Begin VB.Label lacOMMindate 
         Caption         =   "Don't Send Overdue Emails for Dates Prior to:"
         Height          =   315
         Left            =   2280
         TabIndex        =   102
         Top             =   4200
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label lacOverdue 
         Caption         =   "Number of Weeks behind considered Overdue"
         Height          =   315
         Left            =   2280
         TabIndex        =   70
         Top             =   3405
         Visible         =   0   'False
         Width           =   4830
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "Administrator"
      Height          =   3960
      Index           =   2
      Left            =   8505
      TabIndex        =   75
      Top             =   6255
      Visible         =   0   'False
      Width           =   8460
      Begin VB.TextBox txtCCEMail 
         Height          =   285
         Left            =   1770
         MaxLength       =   70
         TabIndex        =   101
         Top             =   3600
         Width           =   3135
      End
      Begin VB.TextBox edcAdminFName 
         Height          =   285
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   77
         Top             =   195
         Width           =   3135
      End
      Begin VB.TextBox edcAdminCity 
         Height          =   285
         Left            =   1770
         MaxLength       =   40
         TabIndex        =   84
         Top             =   1695
         Width           =   1695
      End
      Begin VB.TextBox edcAdminZip 
         Height          =   285
         Left            =   7185
         MaxLength       =   20
         TabIndex        =   88
         Top             =   1680
         Width           =   1035
      End
      Begin VB.TextBox edcAdminCountry 
         Height          =   285
         Left            =   1770
         MaxLength       =   40
         TabIndex        =   90
         Top             =   2070
         Width           =   1620
      End
      Begin VB.TextBox edcAdminState 
         Height          =   285
         Left            =   4530
         MaxLength       =   40
         TabIndex        =   86
         Top             =   1695
         Width           =   1620
      End
      Begin VB.TextBox edcAdminAddress 
         Height          =   285
         Index           =   1
         Left            =   1770
         MaxLength       =   40
         TabIndex        =   82
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox edcAdminAddress 
         Height          =   285
         Index           =   0
         Left            =   1770
         MaxLength       =   40
         TabIndex        =   81
         Top             =   945
         Width           =   2295
      End
      Begin VB.TextBox edcAdminLName 
         Height          =   285
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   79
         Top             =   570
         Width           =   3135
      End
      Begin VB.TextBox edcAdminPhone 
         Height          =   285
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   92
         Top             =   2445
         Width           =   2295
      End
      Begin VB.TextBox edcAdminFax 
         Height          =   285
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   94
         Top             =   2820
         Width           =   2295
      End
      Begin VB.TextBox edcAdminEMail 
         Height          =   285
         Left            =   1770
         MaxLength       =   70
         TabIndex        =   96
         Top             =   3195
         Width           =   3135
      End
      Begin VB.Label lacAdminFName 
         Caption         =   "First Name:"
         Height          =   255
         Left            =   150
         TabIndex        =   76
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label lblCCEMail 
         Caption         =   "Bcc E-mail Address:"
         Height          =   315
         Left            =   120
         TabIndex        =   100
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label lacAdminCity 
         Caption         =   "City:"
         Height          =   255
         Left            =   150
         TabIndex        =   83
         Top             =   1695
         Width           =   975
      End
      Begin VB.Label lacAdminZip 
         Caption         =   "Zip:"
         Height          =   255
         Left            =   6420
         TabIndex        =   87
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label lacAdminCounter 
         Caption         =   "Country:"
         Height          =   255
         Left            =   150
         TabIndex        =   89
         Top             =   2055
         Width           =   810
      End
      Begin VB.Label lacAdminState 
         Caption         =   "State:"
         Height          =   255
         Left            =   3705
         TabIndex        =   85
         Top             =   1695
         Width           =   795
      End
      Begin VB.Label ladAdminAddress 
         Caption         =   "Address:"
         Height          =   255
         Left            =   150
         TabIndex        =   80
         Top             =   945
         Width           =   975
      End
      Begin VB.Label lacAdminLName 
         Caption         =   "Last Name:"
         Height          =   255
         Left            =   150
         TabIndex        =   78
         Top             =   555
         Width           =   1335
      End
      Begin VB.Label lacAdminPhone 
         Caption         =   "Telephone:"
         Height          =   315
         Left            =   150
         TabIndex        =   91
         Top             =   2445
         Width           =   1215
      End
      Begin VB.Label lacAdminFax 
         Caption         =   "Fax:"
         Height          =   315
         Left            =   150
         TabIndex        =   93
         Top             =   2820
         Width           =   1095
      End
      Begin VB.Label lacAdminEMail 
         Caption         =   "E-mail Address:"
         Height          =   315
         Left            =   150
         TabIndex        =   95
         Top             =   3195
         Width           =   1815
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "Web Options"
      Height          =   4725
      Index           =   3
      Left            =   9000
      TabIndex        =   115
      Top             =   6240
      Visible         =   0   'False
      Width           =   8805
      Begin VB.CheckBox optUsingWeb 
         Caption         =   "Using Network Web Site"
         Height          =   240
         Left            =   240
         TabIndex        =   158
         Top             =   240
         Width           =   2250
      End
      Begin VB.CheckBox ckcSyncMulticast 
         Caption         =   "Keep Multi-cast Stations in Sync"
         Height          =   240
         Left            =   255
         TabIndex        =   132
         Top             =   3600
         Width           =   3615
      End
      Begin VB.CheckBox optAutoPost 
         Caption         =   "Station must approve auto-posted information"
         Height          =   255
         Left            =   4290
         TabIndex        =   151
         Top             =   4335
         Visible         =   0   'False
         Width           =   3825
      End
      Begin VB.CheckBox ckcAllowPostInFuture 
         Caption         =   "Allow Posting on Today and in Future Days"
         Height          =   240
         Left            =   240
         TabIndex        =   131
         Top             =   3255
         Width           =   4470
      End
      Begin VB.CheckBox ckcNoMissedReason 
         Caption         =   "Suppress asking Missed Reason"
         Height          =   240
         Left            =   240
         TabIndex        =   128
         Top             =   2265
         Width           =   3900
      End
      Begin VB.CheckBox optSuppressLogs 
         Caption         =   "Suppress Logs on Web Site"
         Height          =   240
         Left            =   4845
         TabIndex        =   133
         Top             =   4080
         Visible         =   0   'False
         Width           =   3090
      End
      Begin VB.CheckBox ckcAllowBonus 
         Caption         =   "Allow Web Bonus Spots"
         Height          =   240
         Left            =   5385
         TabIndex        =   150
         Top             =   3750
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.CheckBox ckcAllowReplacement 
         Caption         =   "Allow Web Replacement Spots"
         Enabled         =   0   'False
         Height          =   240
         Left            =   240
         TabIndex        =   130
         Top             =   2925
         Width           =   2955
      End
      Begin VB.CheckBox ckcAllowWebMG 
         Caption         =   "Allow Web MG Spots"
         Height          =   255
         Left            =   240
         TabIndex        =   129
         Top             =   2595
         Width           =   2280
      End
      Begin VB.TextBox edcDaysRetainPosted 
         Height          =   285
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   121
         Top             =   795
         Width           =   555
      End
      Begin VB.TextBox edcNoDaysRetainPostedB4Importing 
         Height          =   285
         Left            =   7200
         MaxLength       =   2
         TabIndex        =   123
         Top             =   1185
         Width           =   555
      End
      Begin VB.Frame frcChngPswd 
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   116
         Top             =   525
         Width           =   5775
         Begin VB.OptionButton rbcChngPswd 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   118
            Top             =   30
            Width           =   1215
         End
         Begin VB.OptionButton rbcChngPswd 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   119
            Top             =   30
            Width           =   1215
         End
         Begin VB.Label lacChngPswd 
            Caption         =   "Allow Web Password Changes"
            Height          =   315
            Left            =   120
            TabIndex        =   117
            Top             =   45
            Width           =   2535
         End
      End
      Begin VB.TextBox edcNoDaysRetainMissed 
         Height          =   285
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   125
         Top             =   1575
         Width           =   555
      End
      Begin VB.TextBox edcNoDaysViewPost 
         Height          =   285
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   127
         Top             =   1950
         Width           =   555
      End
      Begin VB.Label lacDaysRetainPosted 
         Caption         =   "Number of days to retain posted and exported spots on the web before erasing"
         Height          =   315
         Left            =   240
         TabIndex        =   120
         Top             =   855
         Width           =   7230
      End
      Begin VB.Label lacNoDaysRetainPostedB4Importing 
         Caption         =   "Number of days to delay the network from importing posted web spots (0-30)"
         Height          =   315
         Left            =   240
         TabIndex        =   122
         Top             =   1185
         Width           =   6915
      End
      Begin VB.Label lacNoDaysRetainMissed 
         Caption         =   "Number of days to retain missed spots on the web"
         Height          =   315
         Left            =   240
         TabIndex        =   124
         Top             =   1575
         Width           =   6915
      End
      Begin VB.Label lacNoDaysViewPost 
         Caption         =   "Number of days to view posted spots on the web"
         Height          =   315
         Left            =   240
         TabIndex        =   126
         Top             =   1950
         Width           =   6915
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "MG Policy"
      Height          =   4815
      Index           =   4
      Left            =   9000
      TabIndex        =   105
      Top             =   5850
      Visible         =   0   'False
      Width           =   8730
      Begin VB.Frame frcMG 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   225
         TabIndex        =   113
         Top             =   3615
         Width           =   8400
         Begin VB.OptionButton rbcMG 
            Caption         =   "Yes"
            Height          =   225
            Index           =   8
            Left            =   6465
            TabIndex        =   152
            Top             =   0
            Width           =   900
         End
         Begin VB.OptionButton rbcMG 
            Caption         =   "No"
            Height          =   465
            Index           =   9
            Left            =   7425
            TabIndex        =   153
            Top             =   -120
            Width           =   645
         End
         Begin VB.Label lacMG 
            Caption         =   "Allow Station the option to bypass Makegood scheduling when permission granted"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   114
            Top             =   0
            Width           =   6555
         End
      End
      Begin VB.Frame frcMG 
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   2
         Left            =   240
         TabIndex        =   106
         Top             =   315
         Width           =   6990
         Begin VB.OptionButton rbcMG 
            Caption         =   "can air on any date"
            Height          =   225
            Index           =   7
            Left            =   1440
            TabIndex        =   138
            Top             =   510
            Width           =   2070
         End
         Begin VB.TextBox edcMGDays 
            Height          =   285
            Left            =   3210
            MaxLength       =   3
            TabIndex        =   136
            Top             =   225
            Width           =   555
         End
         Begin VB.OptionButton rbcMG 
            Caption         =   "must air within missed Standard Broadcast Month"
            Height          =   225
            Index           =   5
            Left            =   1440
            TabIndex        =   134
            Top             =   0
            Width           =   4935
         End
         Begin VB.OptionButton rbcMG 
            Caption         =   "must air within"
            Height          =   225
            Index           =   6
            Left            =   1440
            TabIndex        =   135
            Top             =   240
            Width           =   1920
         End
         Begin VB.Label lacMG 
            Caption         =   "days of missed date"
            Height          =   255
            Index           =   2
            Left            =   3900
            TabIndex        =   137
            Top             =   255
            Width           =   1860
         End
         Begin VB.Label lacMG 
            Caption         =   "Makegood"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   107
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.TextBox edcCompetSepTime 
         Height          =   285
         Left            =   3645
         MaxLength       =   4
         TabIndex        =   149
         Top             =   3240
         Width           =   705
      End
      Begin VB.Frame frcMG 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   110
         Top             =   2520
         Width           =   8400
         Begin VB.OptionButton rbcMG 
            Caption         =   "Same Order, then Same Advertiser + New"
            Height          =   225
            Index           =   11
            Left            =   3525
            TabIndex        =   148
            Top             =   360
            Width           =   4245
         End
         Begin VB.OptionButton rbcMG 
            Caption         =   "Same Order + New"
            Height          =   225
            Index           =   12
            Left            =   1680
            TabIndex        =   154
            Top             =   360
            Width           =   2340
         End
         Begin VB.OptionButton rbcMG 
            Caption         =   "Same Order, then Same Advertiser"
            Height          =   225
            Index           =   4
            Left            =   3525
            TabIndex        =   147
            Top             =   0
            Width           =   3525
         End
         Begin VB.OptionButton rbcMG 
            Caption         =   "Same Order"
            Height          =   225
            Index           =   3
            Left            =   1680
            TabIndex        =   146
            Top             =   0
            Width           =   1995
         End
         Begin VB.Label lacMG 
            Caption         =   "ISCI Restrictions"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   111
            Top             =   0
            Width           =   1875
         End
      End
      Begin VB.Frame frcMG 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   108
         Top             =   2160
         Width           =   8500
         Begin VB.OptionButton rbcMG 
            Caption         =   "Any Time"
            Height          =   240
            Index           =   10
            Left            =   6600
            TabIndex        =   145
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton rbcMG 
            Caption         =   "Day Split"
            Height          =   240
            Index           =   2
            Left            =   5520
            TabIndex        =   144
            Top             =   0
            Width           =   1125
         End
         Begin VB.OptionButton rbcMG 
            Caption         =   "Station Pledge Times"
            Height          =   240
            Index           =   1
            Left            =   3540
            TabIndex        =   143
            Top             =   0
            Width           =   2190
         End
         Begin VB.OptionButton rbcMG 
            Caption         =   "Order Flight Times"
            Height          =   240
            Index           =   0
            Left            =   1680
            TabIndex        =   142
            Top             =   0
            Width           =   2145
         End
         Begin VB.Label lacMG 
            Caption         =   "Time Restrictions"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   109
            Top             =   0
            Width           =   2115
         End
      End
      Begin VB.CheckBox ckcMGSpec 
         Caption         =   "Book only into Order Flight Days"
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   141
         Top             =   1800
         Width           =   5850
      End
      Begin VB.CheckBox ckcMGSpec 
         Caption         =   "Disallow MGs in Hiatus Weeks"
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   140
         Top             =   1440
         Width           =   5850
      End
      Begin VB.CheckBox ckcMGSpec 
         Caption         =   "Miss in Last Week of Broadcast Month, Makegood OK in Next Week"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   139
         Top             =   1080
         Width           =   5850
      End
      Begin VB.Label lacMG 
         Caption         =   "Competitive Separation Time in Minutes"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   112
         Top             =   3225
         Width           =   3405
      End
   End
   Begin VB.Timer tmcLoadMessages 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1815
      Top             =   6555
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5580
      TabIndex        =   99
      Tag             =   "Cancel"
      Top             =   6540
      Width           =   1335
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   2550
      TabIndex        =   97
      Tag             =   "OK"
      Top             =   6540
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4065
      TabIndex        =   98
      Tag             =   "OK"
      Top             =   6540
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   825
      Top             =   5985
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7110
      FormDesignWidth =   9510
   End
   Begin ComctlLib.TabStrip tsSiteOptions 
      Height          =   6315
      Left            =   210
      TabIndex        =   0
      Top             =   135
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   11139
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Options"
            Key             =   """1"""
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Messages"
            Key             =   """2"""
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Administrator"
            Key             =   """3"""
            Object.Tag             =   "3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Web Options"
            Key             =   """4"""
            Object.Tag             =   "4"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Web MG Policy"
            Key             =   """5"""
            Object.Tag             =   "5"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lacChanges 
      AutoSize        =   -1  'True
      Caption         =   "change label"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   104
      Top             =   6495
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "frmSiteOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmAddMG - displays site option
'*
'*  Copyright Counterpoint Software, Inc.
'*******************************************************

Option Explicit

Private smAutoPost As String
Private imLastTabIndex As Integer
Private imIgnoreClickEvent As Boolean
Private imIgnoreChange As Boolean
Private imChangesOccured As Integer
Private sToFileHeader As String
Private smWebExports As String
Private imWebSiteNeedsUpdated As Boolean
Private imNoDaysRetainPostedB4Importing As Integer
Private imNoDaysViewPost As Integer
Private imCMCmtCode As Integer
Private smCitation As String
Private imWMCmtCode As Integer
Private smWelcome As String
Private imPMCmtCode As Integer
Private smPassword As String
Private imOMCmtCode As Integer
Private smOverdue As String
Private imBandCmtCode As Integer
Private smBand As String
Private imMessageIndex As Integer
Private commrst As ADODB.Recordset
Private lmAdminArttCode As Long
Private imISCIExpt As Integer
Private adrst As ADODB.Recordset
Private smCCEMail As String
Private smChngPswd As String
Private smEmailOK As Integer
Private smSuppress As String
Private smMultiVehWebPost As String
Private imNoDaysRetainMissed As Integer
Private smAllowBonusSpots As String
Private smAllowMGSpots As String
Private smAllowReplSpots As String
Private smAllowPostInFuture As String
Private smSyncMulticast As String
Private smNoMissedReason As String
Private smShowAgreementDates As String
Private smUsingViero As String
Private smUsingID As String
'Private smCompliantBy As String 'Not Used with v7.0; it was: A=Advertiser; P=Pledge
'D.S. 04/13/16
Private smWithinMissMonth As String
Private smLastWk1stWk As String
Private smSkipHiatusWk As String
Private smValidDaysOnly As String
Private smTimeRange As String
Private smISCIPolicy As String
Private smMissedMGBypass As String
Private imMGDays As Integer
Private imMGCompetitiveSepTime As Integer

Private Type WMVINFO
    iCmtCode As Integer
    iVefCode As Integer
    sComment As String * 1000
End Type
Private tmWMVInfo() As WMVINFO 'Private smCompliantBy As String 'Not Used with v7.0; it was: A=Advertiser; P=Pledge
Private imCurrentSelectedVehicle As Integer
Private bmInClick As Boolean
Private imShift As Integer
Private bmExceededChanges As Boolean
'9926
Const MISSED As Integer = 5
Dim smMissed As String
Dim imUMCmtCode As Integer
Const UNLIMITED = 999



Private Sub ckcAllowBonus_Click()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub ckcAllowBonus_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcAllowPostInFuture_Click()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub ckcAllowPostInFuture_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcAllowReplacement_Click()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub ckcAllowReplacement_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcAllowWebMG_Click()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub ckcAllowWebMG_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcDefaultEstDay_Click()
    If Not imIgnoreChange Then
        mChangeOccured
    End If
End Sub

Private Sub ckcDefaultEstDay_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcMGSpec_Click(Index As Integer)
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub ckcMGSpec_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcNoMissedReason_Click()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
    If ckcNoMissedReason.Value = vbChecked Then
        ckcAllowWebMG.Value = vbUnchecked
        ckcAllowWebMG.Enabled = False
    Else
        ckcAllowWebMG.Enabled = True
    End If
End Sub

Private Sub ckcNoMissedReason_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcSyncMulticast_Click()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub ckcSyncMulticast_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingServiceAgreement_Click()
    If Not imIgnoreChange Then
        mChangeOccured
    End If
End Sub

Private Sub ckcUsingServiceAgreement_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAcctName_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcAcctName_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAdminAddress_Change(Index As Integer)
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcAdminAddress_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAdminCity_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcAdminCity_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAdminCountry_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcAdminCountry_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAdminEMail_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcAdminEMail_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAdminFax_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcAdminFax_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAdminFName_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcAdminFName_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAdminLName_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcAdminLName_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAdminPhone_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcAdminPhone_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAdminState_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcAdminState_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcAdminZip_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcAdminZip_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcBand_Change()
    If Not imIgnoreChange Then
        mChangeOccured
    End If
End Sub

Private Sub edcBand_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcCompetSepTime_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcCompetSepTime_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcDaysRetainPosted_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcDaysRetainPosted_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcDaysRetainPosted_KeyPress(KeyAscii As Integer)
    'If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub edcDaysRetainPosted_LostFocus()
    Dim ilDaysRetainPosted As Integer
    
    ilDaysRetainPosted = Val(edcDaysRetainPosted.Text)
    If ilDaysRetainPosted < 180 Then
        MsgBox "The Number of days to view posted spots on the web must be a value equal or greater than 180"
        edcDaysRetainPosted.Text = "180"
        edcDaysRetainPosted.SetFocus
        Exit Sub
    End If
End Sub

Private Sub edcFromAddress_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcFromAddress_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcFromName_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcFromName_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcHost_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcHost_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcMGDays_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
        rbcMG(6).Value = True   'vbTrue
    End If
End Sub

Private Sub edcMGDays_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcNCRWks_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcNCRWks_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcNoDaysResend_Change()
    'If Not imIgnoreChange Then
    '    mChangeOccured
    '    imWebSiteNeedsUpdated = False
    'End If
End Sub

Private Sub edcNoDaysResend_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcNoDaysResend_KeyPress(KeyAscii As Integer)
    'If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub edcNoDaysRetainMissed_Change()
    'Doug- Send to Web
    If Not imIgnoreChange Then
        mChangeOccured
    '    imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcNoDaysRetainMissed_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcNoDaysRetainMissed_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub edcNoDaysRetainPostedB4Importing_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcNoDaysRetainPostedB4Importing_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub


Private Sub edcNoDaysViewPost_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcNoDaysViewPost_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcNoDaysViewPost_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub edcNoDaysViewPost_LostFocus()
    Dim ilNoDaysViewPost As Integer
    
    ilNoDaysViewPost = Val(edcNoDaysViewPost.Text)
    If ilNoDaysViewPost < 10 Then
        MsgBox "The Number of days to view posted spots on the web must be a value equal or greater than 10"
        edcNoDaysViewPost.Text = "10"
        edcNoDaysViewPost.SetFocus
        Exit Sub
    End If
    imNoDaysViewPost = edcNoDaysViewPost.Text
   
End Sub

Private Sub edcOMMindate_Change()
    If Not imIgnoreChange Then
        mChangeOccured
    End If
End Sub

Private Sub edcOverdue_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub


Private Sub edcOverdue_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcOverdue_KeyPress(KeyAscii As Integer)
    'If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub edcPassword_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcPassword_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcPort_Change()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub edcPort_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.4
    Me.Height = Screen.Height / 1.4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmSiteOptions
    gCenterForm frmSiteOptions


End Sub

Private Sub Form_Load()
    On Error GoTo ErrHndlr_Form_Load
    bgSiteVisible = True
    imChangesOccured = 0
    Screen.MousePointer = vbHourglass
    If igPasswordOk Then
        smEmailOK = True
        ' Dan M 4/20/09 limited changes
        mTestUnlimitedChanges
        mChangeLabel True
    Else
        smEmailOK = False
    End If
    If Not igPasswordOk Then
        ' Dan M  thse were missing
        edcNoDaysRetainPostedB4Importing.Enabled = False
        rbcChngPswd(0).Enabled = False
        rbcChngPswd(1).Enabled = False
        edcOMMindate.Enabled = False
        '
        optUsingUnivision.Enabled = False
        optUsingWeb.Enabled = False
        edcOverdue.Enabled = False
        edcNCRWks.Enabled = False
        edcDaysRetainPosted.Enabled = False
        'edcNoMnthRetain.Enabled = False
        'edcNoDaysResend.Enabled = False
        edcMessage.Enabled = False
        edcAdminFName.Enabled = False
        edcAdminLName.Enabled = False
        edcAdminAddress(0).Enabled = False
        edcAdminAddress(1).Enabled = False
        edcAdminCity.Enabled = False
        edcAdminState.Enabled = False
        edcAdminCountry.Enabled = False
        edcAdminZip.Enabled = False
        edcAdminPhone.Enabled = False
        edcAdminFax.Enabled = False
        edcAdminEMail.Enabled = False
        rbcVehStn(0).Enabled = False
        rbcVehStn(1).Enabled = False
        rbcVehType(0).Enabled = False
        rbcVehType(1).Enabled = False
        'rbcISCI(0).Enabled = False
        'rbcISCI(1).Enabled = False
        'rbcISCI(2).Enabled = False
        rbcRCS4(0).Enabled = False
        rbcRCS4(1).Enabled = False
        rbcRCS5(0).Enabled = False
        rbcRCS5(1).Enabled = False
        cmdSave.Enabled = False
        cmdDone.Enabled = False
        rbcRADARMultiAir(0).Enabled = False
        rbcRADARMultiAir(1).Enabled = False
        rbcRADARMultiAir(2).Enabled = False
    End If
    
    If gIsUsingNovelty Then
        ' Move the only check box we need to the first frame
        Set optUsingWeb.Container = frcTab(0)
        optUsingWeb.Top = optUsingWeb_N.Top
        optUsingWeb.Left = optUsingWeb_N.Left
        
        tsSiteOptions.Tabs.Remove 5 ' Remove the Web MG Policy page.
        tsSiteOptions.Tabs.Remove 4 ' Remove the Web Options page.
        tsSiteOptions.Tabs.Remove 3 ' Remove the Admin page.
        
        ' Hide these on the web options page
        rbcChngPswd(0).Visible = False
        rbcChngPswd(1).Visible = False
        lacChngPswd.Visible = False
        edcNoDaysRetainMissed.Visible = False
        edcNoDaysRetainPostedB4Importing.Visible = False
        edcNoDaysViewPost.Visible = False
        lacNoDaysRetainMissed.Visible = False
        lacNoDaysRetainPostedB4Importing.Visible = False
        lacNoDaysViewPost.Visible = False
        ckcNoMissedReason.Visible = False
        ckcAllowWebMG.Visible = False
        ckcAllowReplacement.Visible = False
        ckcAllowPostInFuture.Visible = False
        ckcSyncMulticast.Visible = False
        lacDaysRetainPosted.Visible = False
        edcDaysRetainPosted.Visible = False
    End If
    
    igPasswordOk = False
    frmSiteOptions.Caption = "Site Options - " & sgClientName
    imLastTabIndex = 1
    frcTab(imLastTabIndex - 1).Visible = True
    frcTab(imLastTabIndex).Visible = False
    edcMessage.Text = ""
    If Not mInit() Then
        gMsgBox "The site options form could not be initialized"
    End If
    
    imWebSiteNeedsUpdated = False
    Screen.MousePointer = vbArrow
    tmcLoadMessages.Enabled = True
    Exit Sub
ErrHndlr_Form_Load:
    Resume Next
End Sub

Private Function mInit()
    On Error GoTo ErrHndlr_Init
    Dim SQLQuery As String
    Dim ilVef As Integer
    
    mInit = False
    imIgnoreClickEvent = True
    imIgnoreChange = True
    '9960 align
    For ilVef = 1 To MISSED
        rbcMessage(ilVef).Top = rbcMessage(0).Top
    Next ilVef
    
    If gIsUsingNovelty Then
        rbcMessage(0).Visible = False
        rbcMessage(2).Visible = False
        rbcMessage(1).Left = 100
        rbcMessage(4).Left = 1000
        rbcMessage(3).Left = 2800
        rbcMessage(5).Left = 3800
    End If
    
    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    
    If Not rst.EOF Then
        'suppress
        'If IsNull(rst!siteWebSuppressLog) Then
        '    optSuppressLogs.Value = vbUnchecked
        'Else
        '    If rst!siteWebSuppressLog = "Y" Then
        '        optSuppressLogs.Value = vbChecked
        '    Else
        '        optSuppressLogs.Value = vbUnchecked
        '    End If
        'End If
        
        If IsNull(rst!siteMultiVehWebPost) Then
            optMultiVehPosting.Value = vbUnchecked
        Else
            If rst!siteMultiVehWebPost = "Y" Then
                optMultiVehPosting.Value = vbChecked
            Else
                optMultiVehPosting.Value = vbUnchecked
            End If
        End If
        
        
        'email options
        If IsNull(rst!siteEmailHost) Then
            edcHost.Text = ""
        Else
            edcHost.Text = Trim$(rst!siteEmailHost)
        End If
    
        If IsNull(rst!siteEmailPort) Then
            edcPort.Text = ""
        Else
            edcPort.Text = Trim$(rst!siteEmailPort)
        End If
    
        If IsNull(rst!siteEmailAcctName) Then
            edcAcctName.Text = ""
        Else
            edcAcctName.Text = Trim$(rst!siteEmailAcctName)
        End If
        
        If IsNull(rst!siteEmailPassword) Then
            edcPassword.Text = ""
        Else
            edcPassword.Text = Trim$(rst!siteEmailPassword)
        End If
        
        If IsNull(rst!siteEmailFromName) Then
            edcFromName.Text = ""
        Else
            edcFromName.Text = Trim$(rst!siteEmailFromName)
        End If
        
        If IsNull(rst!siteEmailFromAddress) Then
            edcFromAddress.Text = ""
        Else
            edcFromAddress.Text = Trim$(rst!siteEmailFromAddress)
        End If
        
        If IsNull(rst!siteOMMinDate) Then
            edcOMMindate.Text = ""
        Else
            edcOMMindate.Text = Trim$(rst!siteOMMinDate)
            edcOMMindate.Text = Format$(edcOMMindate.Text, sgShowDateForm)
        End If
        
        If Trim$(edcOMMindate.Text) = "1/1/1970" Then
            edcOMMindate.Text = ""
        End If
        edcNoDaysRetainPostedB4Importing.Text = rst!siteDayRetainPost
        txtCCEmail.Text = Trim$(rst!siteCCEMail)
        optUsingUnivision.Value = rst!siteMarketron
        optUsingWeb.Value = rst!siteWeb
        imCMCmtCode = 0
        smCitation = ""
        If rst!siteCMCmtCode > 0 Then
            SQLQuery = "SELECT * FROM CMT Where cmtCode = " & rst!siteCMCmtCode
            Set commrst = gSQLSelectCall(SQLQuery)
            If Not commrst.EOF Then
                imCMCmtCode = rst!siteCMCmtCode
                smCitation = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
            End If
        End If
        imWMCmtCode = 0
        smWelcome = ""
        If rst!siteWMCmtCode > 0 Then
            SQLQuery = "SELECT * FROM CMT Where cmtCode = " & rst!siteWMCmtCode
            Set commrst = gSQLSelectCall(SQLQuery)
            If Not commrst.EOF Then
                imWMCmtCode = rst!siteWMCmtCode
                smWelcome = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
            End If
        End If
        imPMCmtCode = 0
        smPassword = ""
        If rst!sitePMCmtCode > 0 Then
            SQLQuery = "SELECT * FROM CMT Where cmtCode = " & rst!sitePMCmtCode
            Set commrst = gSQLSelectCall(SQLQuery)
            If Not commrst.EOF Then
                imPMCmtCode = rst!sitePMCmtCode
                smPassword = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
            End If
        End If
        imOMCmtCode = 0
        smOverdue = ""
        If rst!siteOMCmtCode > 0 Then
            SQLQuery = "SELECT * FROM CMT Where cmtCode = " & rst!siteOMCmtCode
            Set commrst = gSQLSelectCall(SQLQuery)
            If Not commrst.EOF Then
                imOMCmtCode = rst!siteOMCmtCode
                smOverdue = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
            End If
        End If
                
        imBandCmtCode = 0
        smBand = ""
        If rst!siteBandCmtCode > 0 Then
            SQLQuery = "SELECT * FROM CMT Where cmtCode = " & rst!siteBandCmtCode
            Set commrst = gSQLSelectCall(SQLQuery)
            If Not commrst.EOF Then
                imBandCmtCode = rst!siteBandCmtCode
                smBand = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
            End If
        End If
        '9926
        smMissed = ""
        imUMCmtCode = 0
        If rst!siteUMCmtCode > 0 Then
            SQLQuery = "SELECT * FROM CMT Where cmtCode = " & rst!siteUMCmtCode
            Set commrst = gSQLSelectCall(SQLQuery)
            If Not commrst.EOF Then
                imUMCmtCode = rst!siteUMCmtCode
                smMissed = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
            End If
        End If
        
        smWelcome = Trim$(smWelcome)
        smPassword = Trim$(smPassword)
        smOverdue = Trim$(smOverdue)
        smCitation = Trim$(smCitation)
        smBand = Trim$(smBand)
        '9926
        smMissed = Trim$(smMissed)
        If rbcMessage(1).Value Then
            imMessageIndex = 1
            edcMessage.Text = smWelcome
        ElseIf rbcMessage(2).Value Then
            imMessageIndex = 2
            edcMessage.Text = smPassword
        ElseIf rbcMessage(3).Value Then
            imMessageIndex = 3
            edcMessage.Text = smOverdue
        ElseIf rbcMessage(4).Value Then
            lbcVehicles.Visible = True
            imMessageIndex = 4
            edcMessage.Text = ""
        '9926
        ElseIf rbcMessage(MISSED).Value Then
            imMessageIndex = MISSED
            edcMessage.Text = smMissed
        Else
            imMessageIndex = 0
            edcMessage.Text = smCitation
        End If
        edcBand.Text = smBand
        edcDaysRetainPosted.Text = rst!siteDaysRetainSpots
        '6/3/08:  Removed
        'edcNoMnthRetain.text = rst!siteNoMnthRetain
        'If IsNull(rst!siteLastDateArch) Then
        '    edcLastDateArch.text = ""
        'Else
        '    If (DateValue(gAdjYear(rst!siteLastDateArch)) = DateValue("1/1/1970")) Then  'Or (rst!attOnAir = "1/1/70") Then
        '        edcLastDateArch.text = ""
        '    Else
        '        edcLastDateArch.text = Format$(rst!siteLastDateArch, sgShowDateForm)
        '    End If
        'End If
        'edcNoDaysResend.text = ""   'rst!siteDaysResendISCI
        edcOverdue.Text = rst!siteOMNoWeeks
        edcNCRWks.Text = rst!siteNCRWks     '6-30-09
        edcAdminFName.Text = ""
        edcAdminLName.Text = ""
        edcAdminAddress(0).Text = ""
        edcAdminAddress(1).Text = ""
        edcAdminCity.Text = ""
        edcAdminState.Text = ""
        edcAdminCountry.Text = ""
        edcAdminZip.Text = ""
        edcAdminPhone.Text = ""
        edcAdminFax.Text = ""
        edcAdminEMail.Text = ""
        lmAdminArttCode = 0
        If rst!siteAdminArttCode > 0 Then
            SQLQuery = "SELECT * FROM ARTT Where arttCode = " & rst!siteAdminArttCode
            Set adrst = gSQLSelectCall(SQLQuery)
            If Not adrst.EOF Then
                lmAdminArttCode = rst!siteAdminArttCode
                edcAdminFName.Text = Trim$(adrst!arttFirstName)
                edcAdminLName.Text = Trim$(adrst!arttLastName)
                edcAdminAddress(0).Text = Trim$(adrst!arttAddress1)
                edcAdminAddress(1).Text = Trim$(adrst!arttAddress2)
                edcAdminCity.Text = Trim$(adrst!arttCity)
                edcAdminState.Text = Trim$(adrst!arttAddressState)
                edcAdminCountry.Text = Trim$(adrst!arttCountry)
                edcAdminZip.Text = Trim$(adrst!arttZip)
                edcAdminPhone.Text = Trim$(adrst!arttPhone)
                edcAdminFax.Text = Trim$(adrst!arttFax)
                edcAdminEMail.Text = Trim$(adrst!arttEmail)
            End If
        End If
        If rst!siteVehicleStn = "Y" Then
            rbcVehStn(0).Value = True
        Else
            rbcVehStn(1).Value = True
        End If
        If rst!siteShowVehType = "Y" Then
            rbcVehType(0).Value = True
        Else
            rbcVehType(1).Value = True
        End If
        'If rst!siteISCIExport = "N" Then
        '    rbcISCI(0).Value = True
        'ElseIf rst!siteISCIExport = "A" Then
        '    rbcISCI(1).Value = True
        'Else
        '    rbcISCI(2).Value = True
        'End If
        If rst!siteExportCart4 = "N" Then
            rbcRCS4(1).Value = True
        Else
            rbcRCS4(0).Value = True
        End If
        If rst!siteExportCart5 = "N" Then
            rbcRCS5(1).Value = True
        Else
            rbcRCS5(0).Value = True
        End If
        If rst!siteChngPswd = "Y" Then
            rbcChngPswd(0).Value = True
            smChngPswd = "Y"
        Else
            rbcChngPswd(1).Value = True
            smChngPswd = "N"
        End If
        edcNoDaysRetainMissed.Text = rst!siteWebNoDyKeepMiss
        edcNoDaysViewPost.Text = rst!siteWebNoDyViewPost
        If rst!siteAllowBonusSpots = "Y" Then
            'rbcAllowBonusSpots(0).Value = True
            ckcAllowBonus.Value = vbChecked
        Else
            'rbcAllowBonusSpots(1).Value = True
            ckcAllowBonus.Value = vbUnchecked
        End If
        If rst!siteAllowMGSpots = "N" Then
            ckcAllowWebMG.Value = vbUnchecked
        Else
            ckcAllowWebMG.Value = vbChecked
        End If
        'Replacement temporarily removed until we are ready to implement on the Web
        'If rst!siteAllowReplSpots = "N" Then
            ckcAllowReplacement.Value = vbUnchecked
        'Else
        '    ckcAllowReplacement.Value = vbChecked
        'End If
        If rst!siteWebPostInFuture = "N" Then
            ckcAllowPostInFuture.Value = vbUnchecked
        Else
            ckcAllowPostInFuture.Value = vbChecked
        End If
        If rst!siteNoMissedReason = "Y" Then
            ckcNoMissedReason.Value = vbChecked
        Else
            ckcNoMissedReason.Value = vbUnchecked
        End If
        If rst!siteShowContrDate = "Y" Then
            rbcShowAgreementDates(0).Value = True
        Else
            rbcShowAgreementDates(1).Value = True
        End If
        If rst!siteUsingViero = "Y" Then
            rbcUsingViero(0).Value = True
        Else
            rbcUsingViero(1).Value = True
        End If
        If rst!siteUsingStationID = "Y" Then
            rbcUsingID(0).Value = True
        ElseIf rst!siteUsingStationID = "A" Then
            rbcUsingID(2).Value = True
        Else
            rbcUsingID(1).Value = True
        End If
        If rst!siteStationToXDS = "Y" Then
            optXDS(0).Value = vbChecked
        Else
            optXDS(0).Value = vbUnchecked
        End If
        If rst!siteAgreementToXDS = "Y" Then
            optXDS(1).Value = vbChecked
        Else
            optXDS(1).Value = vbUnchecked
        End If
        If rst!siteGenTransparent = "Y" Then
            optXDS(2).Value = vbChecked
        Else
            optXDS(2).Value = vbUnchecked
        End If
        If rst!siteProgToXDS = "Y" Then
            optXDS(3).Value = vbChecked
        Else
            optXDS(3).Value = vbUnchecked
        End If
        If rst!siteSupportXDSDelay = "Y" Then
            optXDS(4).Value = vbChecked
        Else
            optXDS(4).Value = vbUnchecked
        End If
        If rst!siteMissedDateTime = "Y" Then
            optMissed.Value = vbChecked
        Else
            optMissed.Value = vbUnchecked
        End If
        '8274 reverse
        If rst!siteAllowAutoPost = "Y" Then
            optAutoPost.Value = vbUnchecked
        Else
            optAutoPost.Value = vbChecked
        End If
'        If rst!siteAllowAutoPost = "Y" Then
'            optAutoPost.Value = vbChecked
'        Else
'            optAutoPost.Value = vbUnchecked
'        End If
        If rst!siteDefaultEstDay = "Y" Then
            ckcDefaultEstDay.Value = vbChecked
        Else
            ckcDefaultEstDay.Value = vbUnchecked
        End If
        If rst!siteUsingServAgree = "Y" Then
            ckcUsingServiceAgreement.Value = vbChecked
        Else
            ckcUsingServiceAgreement.Value = vbUnchecked
        End If
        If rst!siteSyncMulticast = "Y" Then
            ckcSyncMulticast.Value = vbChecked
        Else
            ckcSyncMulticast.Value = vbUnchecked
        End If
        '8/1/14: Not Used with V7.0
        'If rst!siteCompliantBy = "A" Then
        '    rbcCompliantBy(0).Value = True
        'Else
        '    rbcCompliantBy(1).Value = True
        'End If
        edcMGDays.Text = ""
        If rst!siteWithinMissMonth = "Y" Then
            rbcMG(5).Value = True    'vbTrue
            If rst!siteLastWk1stWk = "Y" Then
                ckcMGSpec(1).Value = vbChecked
            Else
                ckcMGSpec(1).Value = vbUnchecked
            End If
            ckcMGSpec(1).Enabled = True
            edcMGDays.Text = ""
        ElseIf Val(rst!siteMGDays) > 0 Then
            rbcMG(6).Value = True    'vbTrue
            edcMGDays.Text = rst!siteMGDays
        Else
            rbcMG(7).Value = True    'vbTrue
            ckcMGSpec(1).Value = vbUnchecked
            ckcMGSpec(1).Enabled = False
        End If
        If rst!siteSkipHiatusWk = "Y" Then
            ckcMGSpec(2).Value = vbChecked
        Else
            ckcMGSpec(2).Value = vbUnchecked
        End If
        If rst!siteValidDaysOnly = "Y" Then
            ckcMGSpec(3).Value = vbChecked
        Else
            ckcMGSpec(3).Value = vbUnchecked
        End If
        
        
        If rst!siteTimeRange = "O" Then
            rbcMG(0).Value = True
        ElseIf rst!siteTimeRange = "P" Then
            rbcMG(1).Value = True
        ElseIf rst!siteTimeRange = "S" Then
            rbcMG(2).Value = True
        Else
            rbcMG(10).Value = True
        End If
        If rst!siteISCIPolicy = "O" Then
            rbcMG(3).Value = True
        ElseIf rst!siteISCIPolicy = "A" Then
            rbcMG(4).Value = True
        ElseIf rst!siteISCIPolicy = "P" Then
            rbcMG(11).Value = True
        'D.S. 02/11/19 start
        ElseIf rst!siteISCIPolicy = "N" Then
            rbcMG(12).Value = True
        'D.S. 02/11/19 end
        Else
            rbcMG(3).Value = True
        End If
        If rst!siteMissedMGBypass = "Y" Then
            rbcMG(8).Value = True
        Else
            rbcMG(9).Value = True
        End If
        If Val(rst!siteCompetSepTime) > 0 Then
            edcCompetSepTime.Text = rst!siteCompetSepTime
        Else
            edcCompetSepTime.Text = ""
        End If
        If ((Asc(sgSpfUsingFeatures5) And RADAR) <> RADAR) Then
            rbcRADARMultiAir(0).Value = True
            frcRADARMultiAir.Enabled = False
        Else
            If rst!siteRADARMultiAir = "P" Then
                rbcRADARMultiAir(1).Value = True
            ElseIf rst!siteRADARMultiAir = "A" Then
                rbcRADARMultiAir(2).Value = True
            Else
                rbcRADARMultiAir(0).Value = True
            End If
        End If
    Else
        ' Insert a new record for site options if one did not exist. There is always only 1 record in this table.
        SQLQuery = "Insert Into Site (siteCode, siteMarketron, siteWeb, siteNoDaysDelq, siteWMVCmtCode, siteCMCmtCode, "
        SQLQuery = SQLQuery & "siteWMCmtCode, sitePMCmtCode, siteOMCmtCode, siteOMNoWeeks, siteDaysRetainSpots, "
        'SQLQuery = SQLQuery & "siteAdminArttCode, siteDaysResendISCI, siteVehicleStn, siteISCIExport, "
        SQLQuery = SQLQuery & "siteAdminArttCode, siteVehicleStn, siteUsingServAgree, siteSyncMulticast, "
        SQLQuery = SQLQuery & "siteShowVehType, siteDayRetainPost, "
        SQLQuery = SQLQuery & "siteOMMinDate, siteEmailHost, siteEmailPort, siteEmailAcctName, siteEmailPassword, "
        SQLQuery = SQLQuery & "siteEmailFromAddress, siteEmailFromName, "
        SQLQuery = SQLQuery & "siteNCRWks, "
        SQLQuery = SQLQuery & "siteWebSuppressLog, "
        SQLQuery = SQLQuery & "siteMultiVehWebPost, "
        SQLQuery = SQLQuery & "siteWebNoDyKeepMiss, "
        SQLQuery = SQLQuery & "siteAllowBonusSpots, "
        SQLQuery = SQLQuery & "siteDDF092710, "
        SQLQuery = SQLQuery & "siteShowContrDate, "
        SQLQuery = SQLQuery & "siteUsingViero, "
        SQLQuery = SQLQuery & "siteUsingStationID, "
        SQLQuery = SQLQuery & "siteStationToXDS, "
        SQLQuery = SQLQuery & "siteAgreementToXDS, "
        SQLQuery = SQLQuery & "siteMissedDateTime, "
        'SQLQuery = SQLQuery & "siteCompliantBy, "
        SQLQuery = SQLQuery & "siteWebNoDyViewPost, "
        SQLQuery = SQLQuery & "siteRqtDate, "
        SQLQuery = SQLQuery & "siteRqtTime, "
        SQLQuery = SQLQuery & "siteGenTransparent"
        SQLQuery = SQLQuery & "siteProgToXDS, "
        SQLQuery = SQLQuery & "siteSSBDate, "
        SQLQuery = SQLQuery & "siteSSBTime, "
        SQLQuery = SQLQuery & "siteSupportXDSDelay, "
        SQLQuery = SQLQuery & "siteAllowAutoPost, "
        SQLQuery = SQLQuery & "siteWithinMissMonth, "
        SQLQuery = SQLQuery & "siteLastWk1stWk, "
        SQLQuery = SQLQuery & "siteSkipHiatusWk, "
        SQLQuery = SQLQuery & "siteValidDaysOnly, "
        SQLQuery = SQLQuery & "siteTimeRange, "
        SQLQuery = SQLQuery & "siteISCIPolicy, "
        SQLQuery = SQLQuery & "siteMGDays, "
        SQLQuery = SQLQuery & "siteCompetSepTime, "
        SQLQuery = SQLQuery & "siteAllowMGSpots, "
        SQLQuery = SQLQuery & "siteAllowReplSpots, "
        SQLQuery = SQLQuery & "siteNoMissedReason, "
        SQLQuery = SQLQuery & "siteDefaultEstDay, "
        SQLQuery = SQLQuery & "siteWebPostInFuture, "
        SQLQuery = SQLQuery & "siteMissedMGBypass, "
        SQLQuery = SQLQuery & "siteBandCmtCode, "
        'SQLQuery = SQLQuery & "siteUnused "
        SQLQuery = SQLQuery & "siteRADARMultiAir, "
        SQLQuery = SQLQuery & "siteAstMaxLastValue, "
        SQLQuery = SQLQuery & "siteAstMaxDate, "
        SQLQuery = SQLQuery & "siteUnused "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "Values (1, 0, 0, 0, 0, 0, "
        SQLQuery = SQLQuery & "0, 0, 0, 0, 0, "
        'SQLQuery = SQLQuery & "0, 0, 'N', '', "
        SQLQuery = SQLQuery & "0, 'N', 'N', 'N', "  'siteAdminArttCode, siteVehicleStn, siteUsingServAgree, siteSyncMulticast
        SQLQuery = SQLQuery & "'N', 0, "    'siteShowVehType, siteDayRetainPost
        SQLQuery = SQLQuery & "'1970-01-01','',25,'','', "  'siteOMMinDate, siteEmailHost, siteEmailPort, siteEmailAcctName, siteEmailPassword
        SQLQuery = SQLQuery & "'','',"  'siteEmailFromAddress, siteEmailFromName
        SQLQuery = SQLQuery & 60 & ", " 'siteNCRWks
        SQLQuery = SQLQuery & "'" & "N" & "', " 'siteWebSuppressLog
        SQLQuery = SQLQuery & "'" & "N" & "', " 'siteMultiVehWebPost
        SQLQuery = SQLQuery & 60 & ", "         'siteWebNoDyKeepMiss
        SQLQuery = SQLQuery & "'" & "N" & "', " 'siteAllowBonusSpots
        SQLQuery = SQLQuery & "'" & "Y" & "', " 'siteDDF092710
        SQLQuery = SQLQuery & "'" & "N" & "', " 'siteShowContrDate
        SQLQuery = SQLQuery & "'" & "N" & "', " 'siteUsingViero
        SQLQuery = SQLQuery & "'" & "N" & "', " 'siteUsingStationID
        SQLQuery = SQLQuery & "'" & "N" & "', " 'siteStationToXDS
        SQLQuery = SQLQuery & "'" & "N" & "', " 'siteAgreementToXDS
        SQLQuery = SQLQuery & "'" & "N" & "', " 'siteMissedDateTime
        '8/1/14: Not Used with V7.0
        'SQLQuery = SQLQuery & "'" & "P" & "', " 'Compliant by
        SQLQuery = SQLQuery & 14 & ", " 'siteWebNoDyViewPost
        SQLQuery = SQLQuery & "'" & Format("1/1/1970", sgSQLDateForm) & "', " 'RQT Date
        SQLQuery = SQLQuery & "'" & Format("12AM", sgSQLTimeForm) & "', " 'RQT Time
        SQLQuery = SQLQuery & "'" & "N" & "', "     'siteGenTransparent
        SQLQuery = SQLQuery & "'" & "N" & "', "     'siteProgToXDS
        SQLQuery = SQLQuery & "'" & Format("1/1/1970", sgSQLDateForm) & "', " '"siteSSBDate,
        SQLQuery = SQLQuery & "'" & Format("12AM", sgSQLTimeForm) & "', " 'siteSSBTime
        SQLQuery = SQLQuery & "'" & "N" & "', "         'siteSupportXDSDelay
        SQLQuery = SQLQuery & "'" & "N" & "', "         'siteAllowAutoPost
        SQLQuery = SQLQuery & "'" & "N" & "', "         'siteWithinMissMonth
        SQLQuery = SQLQuery & "'" & "N" & "', "         'siteLastWk1stWk
        SQLQuery = SQLQuery & "'" & "N" & "', "         'siteSkipHiatusWk
        SQLQuery = SQLQuery & "'" & "N" & "', "         'siteValidDaysOnly
        SQLQuery = SQLQuery & "'" & "N" & "', "         'siteTimeRange
        SQLQuery = SQLQuery & "'" & "O" & "', "         'siteISCIPolicy
        SQLQuery = SQLQuery & -1 & ", "                 'siteMGDays
        SQLQuery = SQLQuery & 0 & ", "                  'siteCompetSepTime
        SQLQuery = SQLQuery & "'" & "Y" & "', "         'siteAllowMGSpots
        SQLQuery = SQLQuery & "'" & "Y" & "', "         'siteAllowReplSpots
        SQLQuery = SQLQuery & "'" & "N" & "', "         'siteNoMissedReason
        SQLQuery = SQLQuery & "'" & "N" & "', "         'siteDefaultEstDay
        SQLQuery = SQLQuery & "'" & "Y" & "', "         'siteWebPostInFuture
        SQLQuery = SQLQuery & "'" & "N" & "', "         'siteMissedMGBypass
        SQLQuery = SQLQuery & 0 & ", "                  'siteBandCmtCode
        'SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & "'" & "S" & "', "          'siteRADARMultiAir
        SQLQuery = SQLQuery & 0 & ", " 'siteAstMaxLastValue
        SQLQuery = SQLQuery & "'" & Format("1/1/1970", sgSQLDateForm) & "', " 'siteAstMaxDate
        SQLQuery = SQLQuery & "'" & "" & "' "   'Unused
        SQLQuery = SQLQuery & ") "
        
        cnn.BeginTrans
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHndlr_Init:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "SiteOptions-mInit"
            cnn.RollbackTrans
            mInit = False
            Exit Function
        End If
        cnn.CommitTrans
        optUsingUnivision.Value = vbUnchecked
        optUsingWeb.Value = vbUnchecked
        imCMCmtCode = 0
        smCitation = ""
        imWMCmtCode = 0
        smWelcome = ""
        imPMCmtCode = 0
        smPassword = ""
        imOMCmtCode = 0
        smOverdue = ""
        '9926
        imUMCmtCode = 0
        smMissed = ""
        imBandCmtCode = 0
        smBand = ""
        edcMessage.Text = ""
        'edcNoDaysResend.text = ""
        edcDaysRetainPosted.Text = ""
        'edcNoMnthRetain.text = ""
        edcOverdue.Text = ""
        lmAdminArttCode = 0
        edcAdminFName.Text = ""
        edcAdminLName.Text = ""
        edcAdminAddress(0).Text = ""
        edcAdminAddress(1).Text = ""
        edcAdminCity.Text = ""
        edcAdminState.Text = ""
        edcAdminCountry.Text = ""
        edcAdminZip.Text = ""
        edcAdminPhone.Text = ""
        edcAdminFax.Text = ""
        edcAdminEMail.Text = ""
        edcNoDaysRetainPostedB4Importing.Text = ""
        
        rbcVehStn(1).Value = True
        rbcVehType(1).Value = True
        'rbcISCI(2).Value = True
        rbcRCS4(0).Value = True
        rbcRCS5(1).Value = False
        rbcUsingID(1).Value = True
        optXDS(0).Value = vbUnchecked
        optXDS(1).Value = vbUnchecked
        optXDS(2).Value = vbUnchecked
        optXDS(3).Value = vbUnchecked
        optXDS(4).Value = vbUnchecked
        optMissed.Value = vbUnchecked
        optAutoPost.Value = vbUnchecked
        ckcDefaultEstDay.Value = vbUnchecked
        ckcUsingServiceAgreement.Value = vbUnchecked
        rbcMG(7).Value = True    'vbTrue
        ckcMGSpec(1).Value = vbUnchecked
        ckcMGSpec(1).Enabled = False
        ckcMGSpec(2).Value = vbUnchecked
        ckcMGSpec(3).Value = vbUnchecked
        rbcMG(2).Value = True
        rbcMG(3).Value = True
        rbcMG(9).Value = True
        edcMGDays.Text = ""
        edcCompetSepTime.Text = ""
        rbcRADARMultiAir(0).Value = True
    End If
    'If (rbcISCI(0).Enabled = False) And (rbcISCI(1).Enabled = False) And (rbcISCI(2).Enabled = False) Then
    'Else
    '    If Not rbcISCI(0).Value Then
    '        edcNoDaysResend.Enabled = False
    '        lbcNoDaysResend.Enabled = False
    '    Else
    '        edcNoDaysResend.Enabled = True
    '        lbcNoDaysResend.Enabled = True
    '    End If
    'End If
    'If rbcISCI(0).Value Then
    '    imISCIExpt = 0
    'ElseIf rbcISCI(1).Value Then
    '    imISCIExpt = 1
    'Else
    '    imISCIExpt = 2
    'End If
    If ((Asc(sgSpfUsingFeatures5) And STATIONINTERFACE) <> STATIONINTERFACE) Then
        optUsingUnivision.Enabled = False
        optUsingWeb.Enabled = False
    End If
    
    imIgnoreChange = False
    imIgnoreClickEvent = False
    mInit = True
    Exit Function
ErrHndlr_Init:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "SiteOptions-mInit"
    Exit Function
End Function

Private Sub Form_Deactivate()
    gUsingUnivision = False
    gUsingWeb = False
    If ((Asc(sgSpfUsingFeatures5) And STATIONINTERFACE) = STATIONINTERFACE) Then
        If optUsingUnivision.Value = vbChecked Then
            gUsingUnivision = True
        End If
        If optUsingWeb.Value = vbChecked Then
            gUsingWeb = True
        End If
    End If
    sgPasswordAddition = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    bgSiteVisible = False
    igChangesAllowed = 0
    bmExceededChanges = False
    ' Dan M this code was in deactivate: not being called
    gUsingUnivision = False
    gUsingWeb = False
    If ((Asc(sgSpfUsingFeatures5) And STATIONINTERFACE) = STATIONINTERFACE) Then
        ' Dan M 1/14/10 1<>true for optvalue
       ' If optUsingUnivision.Value = True Then
        If optUsingUnivision.Value = vbChecked Then
            gUsingUnivision = True
        End If
       ' If optUsingWeb.Value = True Then
        If optUsingWeb.Value = vbChecked Then
            gUsingWeb = True
        End If
    End If
    sgPasswordAddition = ""
    Erase tmWMVInfo
    commrst.Close
    adrst.Close
    Set frmSiteOptions = Nothing

End Sub

Private Sub lacNoDaysRetainMissed_Click()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub lbcVehicles_Click()
    Dim ilItem As Integer
    ilItem = lbcVehicles.ListIndex
    lbcVehicles_ItemCheck ilItem
End Sub

Private Sub lbcVehicles_ItemCheck(Item As Integer)
    Dim ilVef As Integer
    
    If Not imIgnoreChange Then
        mChangeOccured
    End If
    
    If bmInClick Then
        Exit Sub
    End If
    bmInClick = True
    
    mRetainWMVInfo
    If lbcVehicles.ListIndex < 0 Then
        bmInClick = False
        Exit Sub
    End If
    imCurrentSelectedVehicle = lbcVehicles.ListIndex
    lbcVehicles.Selected(lbcVehicles.ListIndex) = True
    For ilVef = 0 To UBound(tmWMVInfo) - 1 Step 1
        If tmWMVInfo(ilVef).iVefCode = lbcVehicles.ItemData(imCurrentSelectedVehicle) Then
            If lbcVehicles.Selected(imCurrentSelectedVehicle) Then
                edcMessage.Text = Trim$(tmWMVInfo(ilVef).sComment)
            End If
            bmInClick = False
            Exit Sub
        End If
    Next ilVef
    edcMessage.Text = ""
    tmWMVInfo(UBound(tmWMVInfo)).iCmtCode = 0
    tmWMVInfo(UBound(tmWMVInfo)).iVefCode = lbcVehicles.ItemData(imCurrentSelectedVehicle)
    tmWMVInfo(UBound(tmWMVInfo)).sComment = ""
    ReDim Preserve tmWMVInfo(0 To UBound(tmWMVInfo) + 1) As WMVINFO
    bmInClick = False

End Sub

Private Sub optAutoPost_Click()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub optAutoPost_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub optMissed_Click()
    If Not imIgnoreChange Then
        mChangeOccured
    End If
End Sub

Private Sub optMissed_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub optMultiVehPosting_Click()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub optMultiVehPosting_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub optSuppressLogs_Click()
    If Not imIgnoreChange Then
        mChangeOccured
        imWebSiteNeedsUpdated = True
    End If
End Sub

Private Sub optSuppressLogs_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub optUsingUnivision_GotFocus()
imIgnoreChange = False
End Sub

Private Sub optUsingWeb_GotFocus()
imIgnoreChange = False
End Sub

Private Sub optXDS_Click(Index As Integer)
    If Not imIgnoreChange Then
        mChangeOccured
    End If
End Sub

Private Sub optXDS_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub



Private Sub rbcChngPswd_Click(Index As Integer)
    
    If rbcChngPswd(Index).Value Then
        If Not imIgnoreChange Then
            mChangeOccured
            imWebSiteNeedsUpdated = True
        End If
    End If


End Sub

Private Sub rbcChngPswd_GotFocus(Index As Integer)
    imIgnoreChange = False
End Sub

Private Sub rbcISCI_Click(Index As Integer)
'    If rbcISCI(Index).Value Then
'        On Error GoTo ErrHndlr_rbcISCI_Click
'        If imIgnoreClickEvent Then
'            Exit Sub
'        End If
'        imIgnoreClickEvent = True
'        sgPasswordAddition = "ISCI-"
'        mProtectChangesAllowed True
'        If Not igPasswordOk Then
'            CSPWord.Show vbModal
'        End If
'        mProtectChangesAllowed False
'        If Not igPasswordOk Then
'            rbcISCI(imISCIExpt).Value = True
'            If imISCIExpt <> 0 Then
'                edcNoDaysResend.Enabled = False
'                lbcNoDaysResend.Enabled = False
'            Else
'                edcNoDaysResend.Enabled = True
'                lbcNoDaysResend.Enabled = True
'            End If
'        Else
'            If Index <> 0 Then
'                edcNoDaysResend.Enabled = False
'                lbcNoDaysResend.Enabled = False
'            Else
'                edcNoDaysResend.Enabled = True
'                lbcNoDaysResend.Enabled = True
'            End If
'        End If
'        igPasswordOk = False  ' Reset this for next time.
'        imIgnoreClickEvent = False
'        If Not imIgnoreChange Then
'            mChangeOccured
'        End If
'    End If
    Exit Sub
ErrHndlr_rbcISCI_Click:
    Resume Next
End Sub

Private Sub rbcISCI_GotFocus(Index As Integer)
imIgnoreChange = False
End Sub


Private Sub rbcMessage_Click(Index As Integer)
    If rbcMessage(Index).Value Then
        imIgnoreChange = True
        lacOverdue.Visible = False
        lacOMMindate.Visible = False
        lacNCRWks.Visible = False
        edcOverdue.Visible = False
        edcOMMindate.Visible = False
        edcNCRWks.Visible = False
        lbcVehicles.Visible = False
        Select Case imMessageIndex
            Case 0  'Citation
                smCitation = Trim$(edcMessage.Text)
            Case 1  'Welcome
                smWelcome = Trim$(edcMessage.Text)
            Case 2  'Password
                smPassword = Trim$(edcMessage.Text)
            Case 3  'Overdue
                smOverdue = Trim$(edcMessage.Text)
            Case 4  'Welcome by vehicle
                mRetainWMVInfo
            '9926
            Case MISSED
                smMissed = Trim$(edcMessage.Text)
        End Select
        mResetMessage Index
        Select Case Index
            Case 0  'Citation
                edcMessage.Text = Trim$(smCitation)
            Case 1  'Welcome
                edcMessage.Text = Trim$(smWelcome)
            Case 2  'Password
                edcMessage.Text = Trim$(smPassword)
                'frcChngPswd.Visible = True
            Case 3  'Overdue
                edcMessage.Text = Trim$(smOverdue)
                lacOverdue.Visible = True
                lacOMMindate.Visible = True
                edcOverdue.Visible = True
                edcOMMindate.Visible = True
                lacNCRWks.Visible = True
                edcNCRWks.Visible = True
            Case 4  'Welcome by Vehicle
                lbcVehicles.Visible = True
                edcMessage.Text = ""
            '9926
            Case MISSED
                edcMessage.Text = Trim$(smMissed)
        End Select
        imIgnoreChange = False
        imMessageIndex = Index
    End If
End Sub

Private Sub rbcMG_Click(Index As Integer)
    If rbcMG(Index).Value Then
        If Not imIgnoreChange Then
            mChangeOccured
            imWebSiteNeedsUpdated = True
        End If
    End If
    If (Index = 5) And rbcMG(Index).Value Then
        ckcMGSpec(1).Enabled = True
        edcMGDays.Text = ""
    ElseIf (Index = 7) And rbcMG(Index).Value Then
        ckcMGSpec(1).Enabled = False
        ckcMGSpec(1).Value = vbUnchecked
        edcMGDays.Text = ""
    ElseIf (Index = 6) And rbcMG(Index).Value Then
        ckcMGSpec(1).Enabled = False
        ckcMGSpec(1).Value = vbUnchecked
    End If
   
End Sub

Private Sub rbcMG_GotFocus(Index As Integer)
    imIgnoreChange = False
End Sub

Private Sub rbcRADARMultiAir_Click(Index As Integer)
    If imIgnoreClickEvent Then
        Exit Sub
    End If
    mChangeOccured
    imIgnoreClickEvent = True
    If rbcRADARMultiAir(Index).Value Then
        sgPasswordAddition = "RMA-"
        mProtectChangesAllowed True
        If Not igPasswordOk Then
            CSPWord.Show vbModal
        End If
        mProtectChangesAllowed False
        If Not igPasswordOk Then
            rbcRADARMultiAir(Index).Value = False
        End If
    End If
    igPasswordOk = False  ' Reset this for next time.
    imIgnoreClickEvent = False
End Sub

Private Sub rbcRCS4_Click(Index As Integer)
    If rbcRCS4(Index).Value Then
        If Not imIgnoreChange Then
            mChangeOccured
        End If
    End If
End Sub

Private Sub rbcRCS4_GotFocus(Index As Integer)
    imIgnoreChange = False
End Sub

Private Sub rbcRCS5_Click(Index As Integer)
    If rbcRCS5(Index).Value Then
        If Not imIgnoreChange Then
            mChangeOccured
        End If
    End If
End Sub

Private Sub rbcRCS5_GotFocus(Index As Integer)
    imIgnoreChange = False
End Sub

Private Sub rbcShowAgreementDates_Click(Index As Integer)
    If rbcShowAgreementDates(Index).Value Then
        If Not imIgnoreChange Then
            mChangeOccured
        End If
    End If
End Sub

Private Sub rbcShowAgreementDates_GotFocus(Index As Integer)
    imIgnoreChange = False
End Sub

Private Sub rbcUsingID_Click(Index As Integer)
    If rbcUsingID(Index).Value Then
        If Not imIgnoreChange Then
            mChangeOccured
        End If
    End If
End Sub

Private Sub rbcUsingViero_Click(Index As Integer)
    If rbcUsingViero(Index).Value Then
        If Not imIgnoreChange Then
            mChangeOccured
        End If
    End If
End Sub

Private Sub rbcVehStn_Click(Index As Integer)
    If rbcVehStn(Index).Value Then
        If Not imIgnoreChange Then
            mChangeOccured
        End If
    End If
End Sub

Private Sub rbcVehStn_GotFocus(Index As Integer)
    imIgnoreChange = False
End Sub

Private Sub rbcVehType_Click(Index As Integer)
    If rbcVehType(Index).Value Then
        If Not imIgnoreChange Then
            mChangeOccured
        End If
    End If
End Sub

Private Sub rbcVehType_GotFocus(Index As Integer)
    imIgnoreChange = False
End Sub

Private Sub tmcLoadMessages_Timer()
    tmcLoadMessages.Enabled = False
    mLoadVehicleMessages
End Sub

Private Sub tsSiteOptions_Click()
    On Error GoTo ErrHndlr_tsSiteOptions_Click
    frcTab(imLastTabIndex - 1).Visible = False
    frcTab(tsSiteOptions.SelectedItem.Index - 1).Visible = True
    
    '6 was for the E-Mail server information.  The tab was removed, frame was retained
    'If smEmailOK And tsSiteOptions.SelectedItem.Index = 6 Then
    '    frcTab(tsSiteOptions.SelectedItem.Index - 1).Visible = True
    'End If
    'If Not smEmailOK And tsSiteOptions.SelectedItem.Index = 6 Then
    '    frcTab(tsSiteOptions.SelectedItem.Index - 1).Visible = False
    'End If
    
    imLastTabIndex = tsSiteOptions.SelectedItem.Index
    Exit Sub
ErrHndlr_tsSiteOptions_Click:
    Resume Next
End Sub

Private Sub cmdSave_Click()
    mSave (False)
   ' imChangesOccured = False
End Sub

Private Sub cmdDone_Click()
    If mSave(True) Then
        If (gUsingWeb = False) Then
            frmDirectory!cmdEMail.Enabled = False
        Else
            frmDirectory!cmdEMail.Enabled = True
        End If
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub edcMessage_Change()
    If Not imIgnoreChange Then
        mChangeOccured
'        If rbcMessage(0).Value Then
'            imWebSiteNeedsUpdated = True
'        End If
    End If
    'Dan M 1/15/10 moved out of if statement
    If rbcMessage(0).Value Then
        imWebSiteNeedsUpdated = True
    End If

End Sub

Private Function mSave(iAsk As Boolean) As Integer
    On Error GoTo ErrHndlr_mSave
    Dim ilNoDaysResend As Integer
    Dim ilDaysRetainPosted As Integer
    'Dim ilNoMnthRetain As Integer
    Dim ilOverdue As Integer
    Dim ilNCRWks As Integer '6-30-09
    Dim ilRet As Integer
    Dim slVehStn As String
    Dim slVehType As String
    Dim slISCI As String
    Dim slRCSExportCart4 As String
    Dim slRCSExportCart5 As String
    Dim slOverdueDate As String
    Dim blExcessiveChanges As Boolean
    Dim slXDSStation As String
    Dim slXDSAgreement As String
    Dim slGenTransparent As String
    Dim slProgToXDS As String
    Dim slSupportXDSDelay As String
    Dim slMissedDT As String
    Dim slDefaultEstDay As String
    Dim ilVef As Integer
    Dim slRADARMultiAir As String
    Dim slUsingServiceAgreement As String

    
    mSave = True
    If imChangesOccured = 0 Then
        Exit Function
    Else
        If imChangesOccured > igChangesAllowed Or bmExceededChanges Then
            If Not iAsk Then
                gMsgBox "You have made too many changes.  Current changes will not be saved."
                imChangesOccured = igChangesAllowed
                bmExceededChanges = True
                mChangeLabel
                ilRet = mInit()
            End If
            Exit Function
        End If
    End If
    
    If iAsk Then
        If gMsgBox("Save all changes?", vbYesNo) <> vbYes Then
            Exit Function
        End If
    End If
    
    smCCEMail = Trim$(txtCCEmail.Text)
    If Not gTestForMultipleEmail(smCCEMail, "BCC") Then
        If imLastTabIndex = 3 Then
            txtCCEmail.SetFocus
        End If
        Screen.MousePointer = vbDefault
        gMsgBox sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Bcc Email Address Before Continuing", vbExclamation
        gLogMsg sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Bcc Email Address Before Continuing", "WebEmailLog.Txt", False
        Exit Function
    End If
    
    If Trim$(edcMGDays.Text) = "0" Then
        Screen.MousePointer = vbDefault
        mSave = False
        gMsgBox "Zero (0) not allowed for MG Policy->Makegood within number of days. Blank is allowed"
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHndlr_mSave
    
    Select Case imMessageIndex
        Case 0  'Citation
            smCitation = Trim$(edcMessage.Text)
        Case 1  'Welcome
            smWelcome = Trim$(edcMessage.Text)
        Case 2  'Password
            smPassword = Trim$(edcMessage.Text)
        Case 3  'Overdue
            smOverdue = Trim$(edcMessage.Text)
        Case 4  'Welcome by Vehicle
            mRetainWMVInfo
        '9926
        Case MISSED
            smMissed = Trim$(edcMessage.Text)
    End Select
    smBand = Trim$(edcBand.Text)
    If rbcVehStn(0).Value Then
        slVehStn = "Y"
    Else
        slVehStn = "N"
    End If
    If rbcVehType(0).Value Then
        slVehType = "Y"
    Else
        slVehType = "N"
    End If
    slISCI = ""
    sgExportISCI = slISCI
    If rbcRCS4(1).Value Then
        slRCSExportCart4 = "N"
        frmMain.mnuExportRCS4.Enabled = False
    Else
        slRCSExportCart4 = "Y"
        frmMain.mnuExportRCS4.Enabled = True
    End If
    
    If rbcRCS5(1).Value Then
        slRCSExportCart5 = "N"
        frmMain.mnuExportRCS5.Enabled = False
    Else
        slRCSExportCart5 = "Y"
        frmMain.mnuExportRCS5.Enabled = True
    End If
    If optXDS(0).Value = vbChecked Then
        slXDSStation = "Y"
    Else
        slXDSStation = "N"
    End If
    If optXDS(1).Value = vbChecked Then
        slXDSAgreement = "Y"
    Else
        slXDSAgreement = "N"
    End If
    If optXDS(2).Value = vbChecked Then
        slGenTransparent = "Y"
    Else
        slGenTransparent = "N"
    End If
    If optXDS(3).Value = vbChecked Then
        slProgToXDS = "Y"
    Else
        slProgToXDS = "N"
    End If
    If optXDS(4).Value = vbChecked Then
        slSupportXDSDelay = "Y"
    Else
        slSupportXDSDelay = "N"
    End If
    If optMissed.Value = vbChecked Then
        slMissedDT = "Y"
    Else
        slMissedDT = "N"
    End If
    If rbcRADARMultiAir(1).Value Then
        slRADARMultiAir = "P"
    ElseIf rbcRADARMultiAir(2).Value Then
        slRADARMultiAir = "A"
    Else
        slRADARMultiAir = "S"
    End If
    '8274 reversed
    If optAutoPost.Value = vbChecked Then
        smAutoPost = "N"
    Else
        smAutoPost = "Y"
    End If
    If ckcDefaultEstDay.Value = vbChecked Then
        slDefaultEstDay = "Y"
    Else
        slDefaultEstDay = "N"
    End If
    If ckcUsingServiceAgreement.Value = vbChecked Then
        slUsingServiceAgreement = "Y"
    Else
        slUsingServiceAgreement = "N"
    End If
    smWithinMissMonth = "N"
    imMGDays = -1
    If rbcMG(5).Value = True Then   'vbTrue Then
        smWithinMissMonth = "Y"
    ElseIf rbcMG(6).Value = True Then   'vbTrue Then
        imMGDays = Val(edcMGDays.Text)
    End If
    smLastWk1stWk = "N"
    If (ckcMGSpec(1).Value = vbChecked) And (rbcMG(5).Value = True) Then    'vbTrue) Then
        smLastWk1stWk = "Y"
    End If
    smSkipHiatusWk = "N"
    If ckcMGSpec(2).Value = vbChecked Then
        smSkipHiatusWk = "Y"
    End If
    smValidDaysOnly = "N"
    If ckcMGSpec(3).Value = vbChecked Then
        smValidDaysOnly = "Y"
    End If
    
    smTimeRange = "N"
    If rbcMG(0).Value Then
        smTimeRange = "O"
    ElseIf rbcMG(1).Value Then
        smTimeRange = "P"
    ElseIf rbcMG(2).Value Then
        smTimeRange = "S"
    End If
    'D.S. 02/11/19 start
    smISCIPolicy = "O"
    If rbcMG(4).Value Then
        smISCIPolicy = "A"
    ElseIf rbcMG(11).Value Then
        smISCIPolicy = "P"
    ElseIf rbcMG(12).Value Then
        smISCIPolicy = "N"
    End If
    'D.S. 02/11/19 end
    If rbcMG(8).Value Then
        smMissedMGBypass = "Y"
    Else
        smMissedMGBypass = "N"
    End If
    
    imMGCompetitiveSepTime = Val(edcCompetSepTime.Text)
        
    mSave = False
    ilRet = mSaveMessage(smCitation, imCMCmtCode, "C", 0)
    If Not ilRet Then
        Exit Function
    End If
    ilRet = mSaveMessage(smWelcome, imWMCmtCode, "W", 0)
    If Not ilRet Then
        Exit Function
    End If
    ilRet = mSaveMessage(smPassword, imPMCmtCode, "P", 0)
    If Not ilRet Then
        Exit Function
    End If
    ilRet = mSaveMessage(smOverdue, imOMCmtCode, "O", 0)
    If Not ilRet Then
        Exit Function
    End If
    ilRet = mSaveMessage(smBand, imBandCmtCode, "B", 0)
    If Not ilRet Then
        Exit Function
    End If
    '9926
    ilRet = mSaveMessage(smMissed, imUMCmtCode, "M", 0)
    If Not ilRet Then
        Exit Function
    End If
    If imCurrentSelectedVehicle >= 0 Then
        For ilVef = 0 To UBound(tmWMVInfo) - 1 Step 1
            If tmWMVInfo(ilVef).iVefCode = lbcVehicles.ItemData(imCurrentSelectedVehicle) Then
                If lbcVehicles.Selected(imCurrentSelectedVehicle) Then
                    tmWMVInfo(ilVef).sComment = Trim$(edcMessage.Text)
                Else
                    tmWMVInfo(ilVef).sComment = ""
                End If
                Exit For
            End If
        Next ilVef
    End If
    imCurrentSelectedVehicle = -1
    For ilVef = 0 To UBound(tmWMVInfo) - 1 Step 1
        ilRet = mSaveMessage(tmWMVInfo(ilVef).sComment, tmWMVInfo(ilVef).iCmtCode, "V", tmWMVInfo(ilVef).iVefCode)
        If Not ilRet Then
            Exit Function
        End If
    Next ilVef
    
    ilRet = mSaveAdmin()
    If Not ilRet Then
        Exit Function
    End If
    
    ilRet = mSaveEmailConfig()
    If Not ilRet Then
        Exit Function
    End If
    
    ilDaysRetainPosted = 0
    If Len(Trim$(edcDaysRetainPosted.Text)) > 0 Then
        ilDaysRetainPosted = Val(Trim$(edcDaysRetainPosted.Text))
    End If
    
    imNoDaysRetainPostedB4Importing = 0
    If Len(Trim$(edcNoDaysRetainPostedB4Importing.Text)) > 0 Then
        imNoDaysRetainPostedB4Importing = Val(Trim$(edcNoDaysRetainPostedB4Importing.Text))
    End If
    
    If imNoDaysRetainPostedB4Importing < 0 Or imNoDaysRetainPostedB4Importing > 30 Then
        edcNoDaysRetainPostedB4Importing.Text = " "
        Screen.MousePointer = vbDefault
        If imLastTabIndex = 1 Then
            edcNoDaysRetainPostedB4Importing.SetFocus
        End If
        Exit Function
    End If
    
    'ilNoMnthRetain = 0
    'If Len(Trim$(edcNoMnthRetain.text)) > 0 Then
    '    ilNoMnthRetain = Val(Trim$(edcNoMnthRetain.text))
    'End If
    
    ilNoDaysResend = 0
    '2/27/10:  Question moved to ISCI screen
    'If Len(Trim$(edcNoDaysResend.text)) > 0 Then
    '    ilNoDaysResend = Val(Trim$(edcNoDaysResend.text))
    'End If
    
    ilOverdue = 0
    ilNCRWks = 0            '6-30-09
    If Len(Trim$(edcNCRWks.Text)) > 0 Then
        ilNCRWks = Val(Trim$(edcNCRWks.Text))
    End If
    
    If Len(Trim$(edcOverdue.Text)) > 0 Then
        ilOverdue = Val(Trim$(edcOverdue.Text))
    End If
    
    If Len(Trim$(edcOMMindate)) > 0 Then
        If gIsDate(Trim$(edcOMMindate.Text)) Then
            slOverdueDate = Format$(edcOMMindate.Text, sgSQLDateForm)
        Else
            gMsgBox "Please Enter a Valid Date", vbOKOnly
            mSave = False
            Exit Function
        End If
    Else
        slOverdueDate = "1970-01-01"
    End If
    
    If rbcChngPswd(0).Value Then
        smChngPswd = "Y"
    Else
        smChngPswd = "N"
    End If
    
    'If optSuppressLogs.Value = vbChecked Then
    '    smSuppress = "Y"
    'Else
        smSuppress = "N"
    'End If

    If optMultiVehPosting.Value = vbChecked Then
        smMultiVehWebPost = "Y"
    Else
        smMultiVehWebPost = "N"
    End If
    imNoDaysRetainMissed = edcNoDaysRetainMissed.Text
    imNoDaysViewPost = edcNoDaysViewPost.Text
    'If rbcAllowBonusSpots(0).Value Then
    If ckcAllowBonus.Value = vbChecked Then
        smAllowBonusSpots = "Y"
    Else
        smAllowBonusSpots = "N"
    End If
    If ckcAllowWebMG.Value = vbChecked Then
        smAllowMGSpots = "Y"
    Else
        smAllowMGSpots = "N"
    End If
    If ckcAllowReplacement.Value = vbChecked Then
        smAllowReplSpots = "Y"
    Else
        smAllowReplSpots = "N"
    End If
    If ckcAllowPostInFuture.Value = vbChecked Then
        smAllowPostInFuture = "Y"
    Else
        smAllowPostInFuture = "N"
    End If
    If ckcSyncMulticast.Value = vbChecked Then
        smSyncMulticast = "Y"
    Else
        smSyncMulticast = "N"
    End If
    If ckcNoMissedReason.Value = vbChecked Then
        smNoMissedReason = "Y"
    Else
        smNoMissedReason = "N"
    End If
    If rbcShowAgreementDates(0).Value Then
        smShowAgreementDates = "Y"
    Else
        smShowAgreementDates = "N"
    End If
    If rbcUsingViero(0).Value Then
        smUsingViero = "Y"
    Else
        smUsingViero = "N"
    End If
    If rbcUsingID(0).Value Then
        smUsingID = "Y"
    ElseIf rbcUsingID(2).Value Then
        smUsingID = "A"
    Else
        smUsingID = "N"
    End If
    '8/1/14: Not used with v7.0
    'If rbcCompliantBy(0).Value Then
    '    smCompliantBy = "A"
    'Else
    '    smCompliantBy = "P"
    'End If
    'sgMarketronCompliant = smCompliantBy
    '9926
    SQLQuery = "Update Site Set " & " " & _
                "siteMarketron = " & Str(optUsingUnivision.Value) & ", " & _
                "siteWeb = " & Str(optUsingWeb.Value) & ", " & _
                "siteCMCmtCode = " & imCMCmtCode & ", " & _
                "siteWMCmtCode = " & imWMCmtCode & ", " & _
                "sitePMCmtCode = " & imPMCmtCode & ", " & _
                "siteOMCmtCode = " & imOMCmtCode & ", " & "siteBandCmtCode = " & imBandCmtCode & ", " & _
                "siteOMNoWeeks = " & ilOverdue & ", " & _
                "siteNCRWks = " & ilNCRWks & ", " & _
                "siteAdminArttCode = " & lmAdminArttCode & ", " & _
                "siteDaysRetainSpots = " & ilDaysRetainPosted & ", " & _
                "siteDayRetainPost = " & imNoDaysRetainPostedB4Importing & ", " & _
                "siteVehicleStn = '" & slVehStn & "', " & "siteUsingServAgree = '" & slUsingServiceAgreement & "', " & _
                "siteShowVehType = '" & slVehType & "', " & _
                "siteExportCart4 = '" & slRCSExportCart4 & "', " & _
                "siteExportCart5 = '" & slRCSExportCart5 & "', " & _
                "siteOMMindate = '" & slOverdueDate & "', " & _
                "siteChngPswd = '" & smChngPswd & "', " & "siteStationToXDS = '" & slXDSStation & "', " & "siteAgreementToXDS = '" & slXDSAgreement & "', " & "siteGenTransparent = '" & slGenTransparent & "', " & "siteMissedDateTime = '" & slMissedDT & "', " & _
                "siteWebSuppressLog = '" & smSuppress & "', " & "siteProgToXDS = '" & slProgToXDS & "', " & _
                "siteMultiVehWebPost = '" & smMultiVehWebPost & "', " & "siteSupportXDSDelay = '" & slSupportXDSDelay & "', " & _
                "siteWebNoDyKeepMiss  = " & imNoDaysRetainMissed & ", " & "siteWebNoDyViewPost  = " & imNoDaysViewPost & ", " & _
                "siteAllowBonusSpots  = '" & smAllowBonusSpots & "', " & "siteShowContrDate  = '" & smShowAgreementDates & "', " & "siteAllowAutoPost = '" & smAutoPost & "', " & "siteDefaultEstDay = '" & slDefaultEstDay & "', " & _
                "siteWithinMissMonth = '" & smWithinMissMonth & "', " & "siteLastWk1stWk = '" & smLastWk1stWk & "', " & "siteSkipHiatusWk = '" & smSkipHiatusWk & "', " & "siteValidDaysOnly = '" & smValidDaysOnly & "', " & "siteTimeRange = '" & smTimeRange & "', " & "siteISCIPolicy = '" & smISCIPolicy & "', " & "siteMissedMGBypass = '" & smMissedMGBypass & "', " & "siteMGDays = " & imMGDays & ", " & "siteCompetSepTime = " & imMGCompetitiveSepTime & ", " & _
                "siteUsingViero  = '" & smUsingViero & "', " & "siteUsingStationID  = '" & smUsingID & "', " & "siteAllowMGSpots  = '" & smAllowMGSpots & "', " & "siteAllowReplSpots  = '" & smAllowReplSpots & "', " & "siteWebPostInFuture  = '" & smAllowPostInFuture & "', " & "siteNoMissedReason  = '" & smNoMissedReason & "', " & "siteRADARMultiAir  = '" & slRADARMultiAir & "', " & "siteSyncMulticast  = '" & smSyncMulticast & "'" & _
                ",siteUMCmtCode = " & imUMCmtCode & " Where siteCode = 1"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHndlr_mSave:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "SiteOptions-mSave"
        cnn.RollbackTrans
        mSave = False
        Exit Function
    End If
    cnn.CommitTrans
    If (imWebSiteNeedsUpdated) And (optUsingWeb.Value = vbChecked) Then
        Call UpdateWebSite
    End If
    sgShowByVehType = slVehType
    sgRCSExportCart4 = slRCSExportCart4
    sgRCSExportCart5 = slRCSExportCart5
    sgUsingStationID = smUsingID
    sgMissedMGBypass = smMissedMGBypass
    sgUsingServiceAgreement = slUsingServiceAgreement
    
    If sgUsingStationID = "Y" Then
        frmMain!mnuImportStation.Caption = "Update/Add Stations"
    ElseIf sgUsingStationID = "A" Then
        If UBound(tgStationInfo) > LBound(tgStationInfo) Then
            frmMain!mnuImportStation.Caption = "Continue Adding Stations"
        Else
            frmMain!mnuImportStation.Caption = "Add Initial Stations"
        End If
    Else
        If UBound(tgStationInfo) > LBound(tgStationInfo) Then
            frmMain!mnuImportStation.Caption = "Update Existing Stations"
        Else
            frmMain!mnuImportStation.Caption = "Add Initial Stations"
        End If
    End If
    imWebSiteNeedsUpdated = False
    mSave = True
    Screen.MousePointer = vbNormal
    Exit Function

ErrHndlr_mSave:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "SiteOptions-mSave"
End Function

Private Sub Form_Resize()
    On Error GoTo ErrHndlr_Form_Resize
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    tsSiteOptions.Height = (cmdCancel.Top - cmdCancel.Height)
    frcTab(0).BorderStyle = 0
    frcTab(1).BorderStyle = 0
    frcTab(2).BorderStyle = 0
    frcTab(3).BorderStyle = 0
    frcTab(4).BorderStyle = 0
    frcTab(5).BorderStyle = 0
    frcTab(0).Caption = ""
    frcTab(1).Caption = ""
    frcTab(2).Caption = ""
    frcTab(3).Caption = ""
    frcTab(4).Caption = ""
    frcTab(5).Caption = ""
    frcTab(0).Move tsSiteOptions.ClientLeft, tsSiteOptions.ClientTop, tsSiteOptions.ClientWidth, tsSiteOptions.ClientHeight
    frcTab(1).Move tsSiteOptions.ClientLeft, tsSiteOptions.ClientTop, tsSiteOptions.ClientWidth, tsSiteOptions.ClientHeight
    frcTab(2).Move tsSiteOptions.ClientLeft, tsSiteOptions.ClientTop, tsSiteOptions.ClientWidth, tsSiteOptions.ClientHeight
    frcTab(3).Move tsSiteOptions.ClientLeft, tsSiteOptions.ClientTop, tsSiteOptions.ClientWidth, tsSiteOptions.ClientHeight
    frcTab(4).Move tsSiteOptions.ClientLeft, tsSiteOptions.ClientTop, tsSiteOptions.ClientWidth, tsSiteOptions.ClientHeight
    frcTab(5).Move tsSiteOptions.ClientLeft, tsSiteOptions.ClientTop, tsSiteOptions.ClientWidth, tsSiteOptions.ClientHeight
    mResetMessage imMessageIndex
    If imMessageIndex <> 4 Then
        lbcVehicles.Visible = False
    Else
        lbcVehicles.Visible = True
    End If
    'edcMGDays.Left = ckcMGSpec(0).Left + ckcMGSpec(0).Width + 90
    'lacMG(2).Left = edcMGDays.Left + edcMGDays.Width + 90
    Exit Sub
ErrHndlr_Form_Resize:
    Resume Next
End Sub

Private Sub optUsingUnivision_Click()
    On Error GoTo ErrHndlr_optUsingUnivision_Click
    If imIgnoreClickEvent Then
        Exit Sub
    End If
    mChangeOccured
    imIgnoreClickEvent = True
    If optUsingUnivision.Value = vbChecked Then
        sgPasswordAddition = "MKT-"
        mProtectChangesAllowed True
        If Not igPasswordOk Then
            CSPWord.Show vbModal
        End If
        mProtectChangesAllowed False
        If Not igPasswordOk Then
            optUsingUnivision.Value = vbUnchecked
        End If
    End If
    igPasswordOk = False  ' Reset this for next time.
    imIgnoreClickEvent = False
    If optUsingUnivision.Value = vbChecked Then
        frmMain.mnuImportAiredStationSpots.Enabled = True
        frmMain.mnuExportSchdStationSpots.Enabled = True
    Else
        frmMain.mnuImportAiredStationSpots.Enabled = False
        frmMain.mnuExportSchdStationSpots.Enabled = False
    End If
    gUsingUnivision = optUsingUnivision.Value
    Exit Sub
ErrHndlr_optUsingUnivision_Click:
    Resume Next
End Sub

Private Sub optUsingWeb_Click()
    On Error GoTo ErrHndlr_optUsingWeb_Click
    If imIgnoreClickEvent Then
        Exit Sub
    End If
    mChangeOccured
    imWebSiteNeedsUpdated = True
    imIgnoreClickEvent = True
    If optUsingWeb.Value = vbChecked Then
        sgPasswordAddition = "WEB-"
        mProtectChangesAllowed True
        If Not igPasswordOk Then
            CSPWord.Show vbModal
        End If
        mProtectChangesAllowed False
        If Not igPasswordOk Then
            optUsingWeb.Value = vbUnchecked
        Else
            gUsingWeb = True ' Make sure this global variable is turned on now as well.
        End If
    Else
        imWebSiteNeedsUpdated = False
        gUsingWeb = False
    End If
    If optUsingWeb.Value = vbChecked Then
        ' Verify the web settings are ok if they are turning web on.
        While Not gVerifyWebIniSettings()
            frmWebIniOptions.Show vbModal
            If Not igWebIniOptionsOK Then
                imWebSiteNeedsUpdated = False
                optUsingWeb.Value = vbUnchecked
                gUsingWeb = optUsingWeb.Value
                igPasswordOk = False  ' Reset this for next time.
                imIgnoreClickEvent = False
                If optUsingWeb.Value = vbChecked Then
                    frmMain.mnuWebImportAiredStationSpot.Enabled = True
                    frmMain.mnuWebExportSchdStationSpots.Enabled = True
                Else
                    frmMain.mnuWebImportAiredStationSpot.Enabled = False
                    frmMain.mnuWebExportSchdStationSpots.Enabled = False
                End If
                Exit Sub
            End If
        Wend
        If Not gTestAccessToWebServer() Then
            gMsgBox "WARNING!" & vbCrLf & vbCrLf & _
            "Web Server Access Error: The Affiliate System does not have access to the web server or the web server is not responding." & vbCrLf & vbCrLf & _
            "No data will be exported to the web site." & vbCrLf & _
            "No data will be imported from the web site." & vbCrLf & _
            "Sign off system immediately and contact system administrator.", vbExclamation

            'gMsgBox "ALERT!" & vbCrLf & vbCrLf & _
            '"Web Server Access Error: This PC does not have access to the web server or the web server is not responding." & vbCrLf & vbCrLf & _
            '"No web site changes will be made."
        End If
    End If
    
    igPasswordOk = False  ' Reset this for next time.
    imIgnoreClickEvent = False
    'frmMain.mnuWebImportAiredStationSpot.Enabled = optUsingWeb.Value
    'frmMain.mnuWebExportSchdStationSpots.Enabled = optUsingWeb.Value
    If optUsingWeb.Value = vbChecked Then
        frmMain.mnuWebImportAiredStationSpot.Enabled = True
        frmMain.mnuWebExportSchdStationSpots.Enabled = True
    Else
        frmMain.mnuWebImportAiredStationSpot.Enabled = False
        frmMain.mnuWebExportSchdStationSpots.Enabled = False
    End If
    gUsingWeb = optUsingWeb.Value
    Exit Sub
ErrHndlr_optUsingWeb_Click:
    Resume Next
End Sub

Private Function UpdateWebSite() As Boolean
    On Error GoTo ErrHand
    Dim hmToHeader As Integer
    Dim iRet As Integer
    Dim iLoop As Integer
    Dim LastValidCharacter As Integer
    Dim slEUA, slOneChar, slStr As String
    Dim ilSuppress As Integer
    Dim ilMultiVehWebPost As Integer
    Dim slChngPswd As String
    Dim slAllowBonus As String

    If gIsUsingNovelty Then
        UpdateWebSite = True
        Exit Function
    End If
    
    'frmProgressMsg.Show vbModeless, Me
    frmProgressMsg.Show vbModeless
    frmProgressMsg.SetMessage 0, "Updating Web Site... "
    DoEvents
    Screen.MousePointer = vbHourglass
    UpdateWebSite = False
    
    Call gLoadOption(sgWebServerSection, "WebExports", smWebExports)
    smWebExports = gSetPathEndSlash(smWebExports, True)
    sToFileHeader = smWebExports & "WebSiteOptions.txt"
    'hmToHeader = FreeFile
    'iRet = 0
    'Open sToFileHeader For Output Lock Write As hmToHeader
    iRet = gFileOpen(sToFileHeader, "Output Lock Write", hmToHeader)
    If iRet <> 0 Then
        Screen.MousePointer = vbDefault
        frmProgressMsg.SetMessage 1, "Unable to open file " & sToFileHeader & ". Web site not updated."
        Exit Function
    End If

    ' The EUA contains four separate strings. The edit box puts carriage returns in the string.
    ' We want to trim the string before sending it to the web server. The following loop copies
    ' the EUA without the carriage returns and then trims it where the last non space character was found.
    ' slEUA = edcWEBEUA.Text
    slEUA = ""
    LastValidCharacter = 1
    For iLoop = 1 To Len(smCitation)
        slOneChar = Mid(smCitation, iLoop, 1)
        'D.S. 3/29/11 If the character is a double quote the concatenate another double quote onto it
        'so the web doesn't think it's a delimiter
        If slOneChar = """" Then
            slOneChar = slOneChar & """"
        End If
        If Asc(slOneChar) >= 20 Then
            slEUA = slEUA + slOneChar
        End If
    Next

    'Allow users to change their web passwords?
    If smChngPswd = "Y" Then
        slChngPswd = "1"
    Else
        slChngPswd = "0"
    End If
    
    'If smSuppress = "Y" Then
    '    ilSuppress = 1
    'Else
        ilSuppress = 0
    'End If
    
    If Not gIsUsingNovelty Then
        Print #hmToHeader, "AdminFirstName, AdminLastName, AdminAddress1, AdminAddress2, AdminCity, AdminState, AdminZip, AdminCountry, AdminPhone, AdminFax, AdminEmail, DaysRetainPosted, WeeksOverdue, EndUserAgreement,DaysToDelayExport, ChngPswd, SuppressLogs, CombineVehicles, DaysToRetainMissed, AllowBonusSpots, NoDaysViewPost, siteWithinMissMonth, siteLastWk1stWk, siteSkipHiatusWk, siteValidDaysOnly, siteTimeRange, siteISCIPolicy, siteAllowMGSpots, siteAllowReplSpots, siteNoMissedReason, siteMGDays, siteCompetSepTime, siteAllowAutoPost, siteAllowTodayAndFuturePosting, siteMissedMGBypass, siteSyncMulticast"
    Else
        Print #hmToHeader, "AdminFirstName, AdminLastName, AdminAddress1, AdminAddress2, AdminCity, AdminState, AdminZip, AdminCountry, AdminPhone, AdminFax, AdminEmail"
    End If
    
    slStr = ""
    slStr = slStr + """" & edcAdminFName.Text & """" & ", "
    slStr = slStr + """" & edcAdminLName.Text & """" & ", "
    slStr = slStr + """" & edcAdminAddress(0).Text & """" & ", "
    slStr = slStr + """" & edcAdminAddress(1).Text & """" & ", "
    slStr = slStr + """" & edcAdminCity.Text & """" & ", "
    slStr = slStr + """" & edcAdminState.Text & """" & ", "
    slStr = slStr + """" & edcAdminZip.Text & """" & ", "
    slStr = slStr + """" & edcAdminCountry.Text & """" & ", "
    slStr = slStr + """" & edcAdminPhone.Text & """" & ", "
    slStr = slStr + """" & edcAdminFax.Text & """" & ", "
    slStr = slStr + """" & edcAdminEMail.Text & """" & ", "
    
    slStr = slStr + """" & edcDaysRetainPosted.Text & """" & ", "
    slStr = slStr + """" & edcOverdue.Text & """" & ", "
    slStr = slStr + """" & slEUA & """" & ", "
    slStr = slStr + """" & imNoDaysRetainPostedB4Importing & """" & ", "
    slStr = slStr + """" & slChngPswd & """" & ", "
    slStr = slStr + """" & ilSuppress & """" & ", "
    slStr = slStr + """" & smMultiVehWebPost & """" & ","
    slStr = slStr & """" & Trim$(edcNoDaysRetainMissed.Text) & """" & ","
    'If rbcAllowBonusSpots(0).Value Then
    If ckcAllowBonus.Value = vbChecked Then
        slAllowBonus = "Y"
    Else
        slAllowBonus = "N"
    End If
    slStr = slStr & """" & Trim$(slAllowBonus) & """" & ","
    slStr = slStr & """" & Trim$(edcNoDaysViewPost.Text) & """" & ","
    
    slStr = slStr & """" & Trim$(smWithinMissMonth) & """" & ","
    slStr = slStr & """" & Trim$(smLastWk1stWk) & """" & ","
    slStr = slStr & """" & Trim$(smSkipHiatusWk) & """" & ","
    slStr = slStr & """" & Trim$(smValidDaysOnly) & """" & ","
    slStr = slStr & """" & Trim$(smTimeRange) & """" & ","
    slStr = slStr & """" & Trim$(smISCIPolicy) & """" & ","
    slStr = slStr & """" & Trim$(smAllowMGSpots) & """" & ","
    slStr = slStr & """" & Trim$(smAllowReplSpots) & """" & ","
    slStr = slStr & """" & Trim$(smNoMissedReason) & """" & ","
    slStr = slStr & Trim$(imMGDays) & ","
    slStr = slStr & imMGCompetitiveSepTime & ","
    slStr = slStr & """" & smAutoPost & """" & ","
    If ckcAllowPostInFuture = vbChecked Then
        slStr = slStr & """" & "Y" & """" & ","
    Else
        slStr = slStr & """" & "N" & """" & ","
    End If
    If smMissedMGBypass = "Y" Then
        slStr = slStr & """" & "Y" & """" & ","
    Else
        slStr = slStr & """" & "N" & """" & ","
    End If
    'D.S. 8/26/19 TTP 9466
    If ckcSyncMulticast.Value = vbChecked Then
        slStr = slStr & """" & "Y" & """"
    Else
        slStr = slStr & """" & "N" & """"
    End If
        
    Print #hmToHeader, slStr
    Close #hmToHeader
    
    If Not gFTPFileToWebServer(sToFileHeader, "WebSiteOptions.txt") Then
        Screen.MousePointer = vbDefault
        frmProgressMsg.SetMessage 1, "Unable to update the Web Server." & vbCrLf & "Web site not updated"
        Exit Function
    End If
    If Not gSendCmdToWebServer("ImportSiteOptions.dll", "WebSiteOptions.txt") Then
        Screen.MousePointer = vbDefault
        frmProgressMsg.SetMessage 1, "FAIL: Unable to instruct Web Server to Import..."
        Exit Function
    End If
    Unload frmProgressMsg
    Screen.MousePointer = vbDefault
    UpdateWebSite = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "SiteOptions-UpdateWebSite"
End Function


Private Function mSaveMessage(slMessage As String, ilCmtCode As Integer, slType As String, ilVefCode As Integer) As Integer
    Dim slPart1 As String
    Dim slPart2 As String
    Dim slPart3 As String
    Dim slPart4 As String
    On Error GoTo ErrHand
    
    slPart1 = gFixQuote(Mid(slMessage, 1, 255))
    slPart2 = gFixQuote(Mid(slMessage, 256, 255))
    slPart3 = gFixQuote(Mid(slMessage, 511, 255))
    slPart4 = gFixQuote(Mid(slMessage, 766, 255))
    If Len(Trim$(slMessage)) > 0 Then
        If ilCmtCode > 0 Then
            SQLQuery = "Update Cmt Set "
            SQLQuery = SQLQuery & "cmtPart1 = '" & slPart1 & "', "
            SQLQuery = SQLQuery & "cmtPart2 = '" & slPart2 & "', "
            SQLQuery = SQLQuery & "cmtPart3 = '" & slPart3 & "', "
            SQLQuery = SQLQuery & "cmtPart4 = '" & slPart4 & "' "
            SQLQuery = SQLQuery & "Where cmtCode = " & ilCmtCode
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub Hand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SiteOptions-mSaveMessage"
                cnn.RollbackTrans
                mSaveMessage = False
                Exit Function
            End If
            cnn.CommitTrans
        Else
            SQLQuery = "INSERT INTO cmt (cmtType, cmtVefCode, cmtPart1, cmtPart2, cmtPart3, cmtPart4)"
            SQLQuery = SQLQuery & " VALUES ('" & slType & "', " & ilVefCode & ", '" & slPart1 & "', '" & slPart2 & "', "
            SQLQuery = SQLQuery & "'" & slPart3 & "', '" & slPart4 & "')"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub Hand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SiteOptions-mSaveMessage"
                mSaveMessage = False
                Exit Function
            End If
            SQLQuery = "Select MAX(cmtCode) from cmt"
            Set commrst = gSQLSelectCall(SQLQuery)
            ilCmtCode = commrst(0).Value
        End If
    Else
        If ilCmtCode > 0 Then
            cnn.BeginTrans
            SQLQuery = "DELETE FROM Cmt WHERE (cmtCode = " & ilCmtCode & ")"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub Hand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SiteOptions-mSaveMessage"
                cnn.RollbackTrans
                mSaveMessage = False
                Exit Function
            End If
            cnn.CommitTrans
        End If
        ilCmtCode = 0
    End If
    mSaveMessage = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "SiteOptions-mSaveMessage"
    mSaveMessage = False
End Function

Private Function mSaveAdmin()
    Dim slFName As String
    Dim slLName As String
    Dim slAddress(0 To 1) As String
    Dim slCity As String
    Dim slState As String
    Dim slZip As String
    Dim slCountry As String
    Dim slPhone As String
    Dim slFax As String
    Dim slEMail As String
    
    On Error GoTo ErrHand
    smCCEMail = Trim$(txtCCEmail.Text)
    slFName = gFixQuote(Trim$(edcAdminFName.Text))
    slLName = gFixQuote(Trim$(edcAdminLName.Text))
    slAddress(0) = gFixQuote(Trim$(edcAdminAddress(0).Text))
    slAddress(1) = gFixQuote(Trim$(edcAdminAddress(1).Text))
    slCity = gFixQuote(Trim$(edcAdminCity.Text))
    slState = gFixQuote(Trim$(edcAdminState.Text))
    slZip = Trim$(edcAdminZip.Text)
    slCountry = gFixQuote(Trim$(edcAdminCountry.Text))
    slPhone = Trim$(edcAdminPhone.Text)
    slFax = Trim$(edcAdminFax.Text)
    slEMail = Trim$(edcAdminEMail.Text)
    
    If Not gTestForMultipleEmail(slEMail, "Reg") Then
        If imLastTabIndex = 3 Then
            edcAdminEMail.SetFocus
        End If
        Screen.MousePointer = vbDefault
        gMsgBox sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Email Address Before Continuing", vbExclamation
        gLogMsg sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Email Address Before Continuing", "WebEmailLog.Txt", False
        Exit Function
    End If
    
    If (Len(slFName) <> 0) Or (Len(slLName) <> 0) Then
        If lmAdminArttCode > 0 Then
            SQLQuery = "Update artt Set "
            SQLQuery = SQLQuery & "arttFirstName = '" & gFixQuote(slFName) & "', "
            SQLQuery = SQLQuery & "arttLastName = '" & gFixQuote(slLName) & "', "
            SQLQuery = SQLQuery & "arttAddress1 = '" & slAddress(0) & "', "
            SQLQuery = SQLQuery & "arttAddress2 = '" & slAddress(1) & "', "
            SQLQuery = SQLQuery & "arttCity = '" & slCity & "', "
            SQLQuery = SQLQuery & "arttAddressState = '" & slState & "', "
            SQLQuery = SQLQuery & "arttCountry = '" & slCountry & "', "
            SQLQuery = SQLQuery & "arttZip = '" & slZip & "', "
            SQLQuery = SQLQuery & "arttPhone = '" & slPhone & "', "
            SQLQuery = SQLQuery & "arttFax = '" & slFax & "', "
            SQLQuery = SQLQuery & "arttEMail = '" & slEMail & "' "
            SQLQuery = SQLQuery & "Where arttCode = " & lmAdminArttCode
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub Hand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SiteOptions-mSaveAdmin"
                cnn.RollbackTrans
                mSaveAdmin = False
                Exit Function
            End If
            cnn.CommitTrans
            SQLQuery = "Update site Set "
            SQLQuery = SQLQuery & "siteCCEMail = '" & smCCEMail & "' "
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub Hand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SiteOptions-mSaveAdmin"
                cnn.RollbackTrans
                mSaveAdmin = False
                Exit Function
            End If
            cnn.CommitTrans
        Else
            SQLQuery = "INSERT INTO artt (arttType, arttFirstName, arttLastName, arttAddress1, arttAddress2, "
            SQLQuery = SQLQuery & "arttCity, arttAddressState, arttCountry, "
            SQLQuery = SQLQuery & "arttZip, arttPhone, arttFax, arttEMail, arttEMailRights)"
            SQLQuery = SQLQuery & " VALUES ('" & "A" & "', '" & gFixQuote(slFName) & "', '" & gFixQuote(slLName) & "', '" & gFixQuote(slAddress(0)) & "', '" & gFixQuote(slAddress(1)) & "', "
            SQLQuery = SQLQuery & "'" & slCity & "', '" & slState & "', '" & slCountry & "', '" & slZip & "', "
            SQLQuery = SQLQuery & "'" & slPhone & "', '" & slFax & "', '" & slEMail & "', '" & "N" & "')"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub Hand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SiteOptions-mSaveAdmin"
                mSaveAdmin = False
                Exit Function
            End If
            SQLQuery = "Select MAX(arttCode) from artt"
            Set adrst = gSQLSelectCall(SQLQuery)
            lmAdminArttCode = adrst(0).Value
        End If
    Else
        If lmAdminArttCode > 0 Then
            cnn.BeginTrans
            SQLQuery = "DELETE FROM artt WHERE (arttCode = " & lmAdminArttCode & ")"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub Hand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SiteOptions-mSaveAdmin"
                cnn.RollbackTrans
                mSaveAdmin = False
                Exit Function
            End If
            cnn.CommitTrans
        End If
        lmAdminArttCode = 0
    End If
    mSaveAdmin = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "SiteOptions-mSaveAdmin"
    mSaveAdmin = False
End Function

Private Sub txtCCEMail_Change()
    If Not imIgnoreChange Then
        mChangeOccured
    End If
End Sub

Private Function mSaveEmailConfig()

    Dim slHost As String
    Dim ilPortNumber As Integer
    Dim slAccountName As String
    Dim slPassword As String
    Dim slFromName As String
    Dim slFromAddress As String
    Dim rstSite As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    slHost = gFixQuote(Trim$(edcHost.Text))
    ilPortNumber = CInt(Trim$(edcPort.Text))
    slAccountName = gFixQuote(Trim$(edcAcctName.Text))
    slPassword = gFixQuote(Trim$(edcPassword.Text))
    slFromName = gFixQuote(Trim$(edcFromName.Text))
    slFromAddress = gFixQuote(Trim$(edcFromAddress.Text))

    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
    Set rstSite = gSQLSelectCall(SQLQuery)
    If Not rstSite.EOF Then
        SQLQuery = "Update site Set "
        SQLQuery = SQLQuery & "siteEmailHost = '" & slHost & "', "
        SQLQuery = SQLQuery & "siteEmailPort = " & ilPortNumber & ","
        SQLQuery = SQLQuery & "siteEmailAcctName = '" & slAccountName & "', "
        SQLQuery = SQLQuery & "siteEmailPassword = '" & slPassword & "', "
        SQLQuery = SQLQuery & "siteEmailFromName = '" & slFromName & "', "
        SQLQuery = SQLQuery & "siteEmailFromAddress = '" & slFromAddress & "'"
        SQLQuery = SQLQuery & "Where siteCode = " & 1
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub Hand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "SiteOptions-mSaveEmailConfig"
            mSaveEmailConfig = False
            Exit Function
        End If
    Else
        SQLQuery = "INSERT INTO site (siteEmailHost, siteEmailPort, siteEmailAcctName, siteEmailPassword, siteEmailFromName, slFromAddress"
        SQLQuery = SQLQuery & " VALUES ('" & slHost & "', ilPortNumber, '" & slAccountName & "', '" & slPassword & "', '" & slFromName & "', '" & slFromAddress & "'"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub Hand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "SiteOptions-mSaveEmailConfig"
            mSaveEmailConfig = False
            Exit Function
        End If
    End If
    mSaveEmailConfig = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "SiteOptions-mSaveEmailConfig"
    mSaveEmailConfig = False
End Function
Private Sub mChangeOccured()
    imChangesOccured = imChangesOccured + 1
    mChangeLabel
    imIgnoreChange = True
End Sub
Private Sub mChangeLabel(Optional blFirstCall As Boolean = False)
Dim slChangesAllowed As String
    If Not bmExceededChanges Then
        If igChangesAllowed = 1 Then
            slChangesAllowed = " change made."
        Else
            slChangesAllowed = " changes made."
        End If
        lacChanges(0).Caption = imChangesOccured & " of " & igChangesAllowed & slChangesAllowed
    Else
        lacChanges(0).Caption = "No more changes may be made."
    End If
    If blFirstCall Then
        lacChanges(0).Visible = True
    End If
End Sub
Private Sub mTestUnlimitedChanges()
    If igChangesAllowed = -1 Then
        igChangesAllowed = UNLIMITED
    End If
End Sub
Private Sub mCtrlGotFocusAndIgnoreChange(Ctrl As control)
'   Dan M 4/22/09 copied gCtrlGotFocus and added bmIgnoreChange to help control textboxes and counting changes
'   gCtrlGotFocus Ctrl
'   Where:
'       Ctrl (I)- control for which text will be highlighted
'
    imIgnoreChange = False
    If TypeOf Ctrl Is TextBox Then
        Ctrl.SelStart = 0
        Ctrl.SelLength = Len(Ctrl.Text)
    End If
End Sub
Private Sub mProtectChangesAllowed(blStart As Boolean)
    Static isSaveChangesAllowed As Integer
    If blStart Then
        isSaveChangesAllowed = igChangesAllowed
    Else
        igChangesAllowed = isSaveChangesAllowed
    End If
End Sub

Private Sub mFillVehicle()
    Dim iLoop As Integer
    lbcVehicles.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
End Sub

Private Sub mRetainWMVInfo()
    Dim ilVef As Integer
    
    If imCurrentSelectedVehicle >= 0 Then
        For ilVef = 0 To UBound(tmWMVInfo) - 1 Step 1
            If tmWMVInfo(ilVef).iVefCode = lbcVehicles.ItemData(imCurrentSelectedVehicle) Then
                If lbcVehicles.Selected(imCurrentSelectedVehicle) Then
                    tmWMVInfo(ilVef).sComment = Trim$(edcMessage.Text)
                Else
                    tmWMVInfo(ilVef).sComment = ""
                End If
                Exit For
            End If
        Next ilVef
    End If
    imCurrentSelectedVehicle = -1

End Sub

Private Sub mResetMessage(ilIndex As Integer)

    If ilIndex <> 4 Then
        edcMessage.Left = frcMessage.Left
        edcMessage.Width = frcMessage.Width
    Else
        lbcVehicles.Left = frcMessage.Left
        edcMessage.Left = frcMessage.Left + lbcVehicles.Width + 240
        edcMessage.Width = frcMessage.Width + frcMessage.Left - edcMessage.Left
    End If

End Sub


Private Sub mLoadVehicleMessages()
    Dim ilVef As Integer
    bmInClick = True
    imCurrentSelectedVehicle = -1
    mFillVehicle
    ReDim tmWMVInfo(0 To 0) As WMVINFO
    SQLQuery = "SELECT * FROM CMT Where cmtType = " & "'V'"
    Set commrst = gSQLSelectCall(SQLQuery)
    Do While Not commrst.EOF
        bmInClick = True
        tmWMVInfo(UBound(tmWMVInfo)).iCmtCode = commrst!cmtCode
        tmWMVInfo(UBound(tmWMVInfo)).iVefCode = commrst!cmtVefCode
        tmWMVInfo(UBound(tmWMVInfo)).sComment = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
        ReDim Preserve tmWMVInfo(0 To UBound(tmWMVInfo) + 1) As WMVINFO
        For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
            If commrst!cmtVefCode = lbcVehicles.ItemData(ilVef) Then
                lbcVehicles.Selected(ilVef) = True
                Exit For
            End If
        Next ilVef
        commrst.MoveNext
    Loop
    bmInClick = False
    lbcVehicles.ListIndex = -1
End Sub
