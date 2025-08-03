VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form SiteOpt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7155
   ClientLeft      =   3465
   ClientTop       =   10365
   ClientWidth     =   12240
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
   Icon            =   "Siteopt.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7155
   ScaleWidth      =   12240
   Begin VB.PictureBox plcCopy 
      Height          =   5820
      Left            =   7650
      ScaleHeight     =   5760
      ScaleWidth      =   10935
      TabIndex        =   205
      Top             =   9240
      Visible         =   0   'False
      Width           =   10995
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Pre-Rec Promo (PP)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   6045
         TabIndex        =   254
         Top             =   5445
         Width           =   2085
      End
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Pre-Rec Cmml(PC)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   3915
         TabIndex        =   253
         Top             =   5445
         Width           =   1950
      End
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Rec Promo (RP)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   2085
         TabIndex        =   252
         Top             =   5445
         Width           =   1605
      End
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Live Promo (LP)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   6045
         TabIndex        =   251
         Top             =   5205
         Width           =   1665
      End
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Live Cmml (LC)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   3915
         TabIndex        =   250
         Top             =   5205
         Width           =   1605
      End
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Rec Cmml (RC)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   2085
         TabIndex        =   249
         Top             =   5205
         Width           =   1605
      End
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Synchronize Copy within same Rotation across Vehicles"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   240
         Top             =   4200
         Width           =   5745
      End
      Begin VB.Frame frcCopy 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   243
         Top             =   4890
         Width           =   7410
         Begin VB.OptionButton rbcSplitCopyState 
            Caption         =   "Mailing"
            Height          =   210
            Index           =   0
            Left            =   3795
            TabIndex        =   245
            Top             =   0
            Width           =   960
         End
         Begin VB.OptionButton rbcSplitCopyState 
            Caption         =   "License"
            Height          =   210
            Index           =   1
            Left            =   4830
            TabIndex        =   246
            Top             =   0
            Width           =   1005
         End
         Begin VB.OptionButton rbcSplitCopyState 
            Caption         =   "Physical"
            Height          =   210
            Index           =   2
            Left            =   5910
            TabIndex        =   247
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label lacCopy 
            Caption         =   "Split Network/Copy Station State Address by"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   244
            Top             =   -15
            Width           =   3780
         End
      End
      Begin VB.TextBox edcSchedule 
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
         Height          =   300
         Index           =   2
         Left            =   3075
         MaxLength       =   10
         TabIndex        =   242
         Top             =   4500
         Width           =   1080
      End
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Limit ISCI to 15 Characters"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   239
         Top             =   3900
         Width           =   3045
      End
      Begin VB.PictureBox plcBB 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   6105
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   720
         Width           =   6105
         Begin VB.OptionButton rbcBBType 
            Caption         =   "Closest Open or Close BB"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   2625
            TabIndex        =   220
            Top             =   0
            Width           =   2505
         End
         Begin VB.OptionButton rbcBBType 
            Caption         =   "Open/Close"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1215
            TabIndex        =   219
            Top             =   0
            Width           =   1380
         End
      End
      Begin VB.PictureBox plcBB 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   5370
         ScaleHeight     =   240
         ScaleWidth      =   4245
         TabIndex        =   215
         TabStop         =   0   'False
         Top             =   990
         Visible         =   0   'False
         Width           =   4245
         Begin VB.OptionButton rbcBBOnLine 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2595
            TabIndex        =   216
            Top             =   0
            Width           =   675
         End
         Begin VB.OptionButton rbcBBOnLine 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3315
            TabIndex        =   217
            Top             =   0
            Width           =   540
         End
      End
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Media Code by Vehicle"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   238
         Top             =   3600
         Width           =   2445
      End
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Using Promo Copy with Schedule Lines"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   237
         Top             =   3300
         Width           =   3645
      End
      Begin VB.Frame frcCopy 
         Caption         =   "Fill Copy Rotation Assignment by"
         Height          =   570
         Index           =   3
         Left            =   120
         TabIndex        =   233
         Top             =   2625
         Width           =   8205
         Begin VB.OptionButton rbcFillCopyAssign 
            Caption         =   "Scheduled Vehicle Only"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   2250
            TabIndex        =   235
            Top             =   240
            Width           =   2370
         End
         Begin VB.OptionButton rbcFillCopyAssign 
            Caption         =   "Original Vehicle Only"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   234
            Top             =   240
            Width           =   2070
         End
         Begin VB.OptionButton rbcFillCopyAssign 
            Caption         =   "Scheduled Vehicle or Original Vehicle"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   4695
            TabIndex        =   236
            Top             =   240
            Width           =   3405
         End
      End
      Begin VB.Frame frcCopy 
         Caption         =   "MG and Outside Copy Rotation Assignment by"
         Height          =   960
         Index           =   2
         Left            =   120
         TabIndex        =   226
         Top             =   1470
         Width           =   8205
         Begin VB.OptionButton rbcMGCopyAssign 
            Caption         =   "Scheduled Vehicle or Original Vehicle"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   4695
            TabIndex        =   229
            Top             =   240
            Width           =   3390
         End
         Begin VB.OptionButton rbcMGCopyAssign 
            Caption         =   "Original Vehicle Only"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   75
            TabIndex        =   227
            Top             =   240
            Width           =   2085
         End
         Begin VB.OptionButton rbcMGCopyAssign 
            Caption         =   "Scheduled Vehicle Only"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   2250
            TabIndex        =   228
            Top             =   240
            Width           =   2235
         End
         Begin VB.Frame frcMGAssignRule 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   15
            TabIndex        =   230
            Top             =   630
            Width           =   5175
            Begin VB.OptionButton rbcMGRules 
               Caption         =   "Use above Rules"
               Height          =   210
               Index           =   0
               Left            =   45
               TabIndex        =   231
               Top             =   0
               Width           =   1815
            End
            Begin VB.OptionButton rbcMGRules 
               Caption         =   "Ask Above Rule in Copy"
               Height          =   210
               Index           =   1
               Left            =   1875
               TabIndex        =   232
               Top             =   0
               Width           =   2310
            End
         End
         Begin VB.Line Line1 
            X1              =   60
            X2              =   8130
            Y1              =   540
            Y2              =   540
         End
      End
      Begin VB.PictureBox plcAISCI 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   6240
         TabIndex        =   221
         TabStop         =   0   'False
         Top             =   1005
         Width           =   6240
         Begin VB.OptionButton rbcAISCI 
            Caption         =   "Never"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   4260
            TabIndex        =   223
            Top             =   0
            Width           =   780
         End
         Begin VB.OptionButton rbcAISCI 
            Caption         =   "Always"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2520
            TabIndex        =   222
            Top             =   0
            Width           =   900
         End
         Begin VB.OptionButton rbcAISCI 
            Caption         =   "Ask- default Yes"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   224
            Top             =   225
            Width           =   1680
         End
         Begin VB.OptionButton rbcAISCI 
            Caption         =   "Ask- default No"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   4245
            TabIndex        =   225
            Top             =   225
            Width           =   1725
         End
      End
      Begin VB.Frame frcCopy 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   211
         Top             =   375
         Width           =   4725
         Begin VB.OptionButton rbcTapeShowForm 
            Caption         =   "Approved"
            Height          =   210
            Index           =   0
            Left            =   1500
            TabIndex        =   213
            Top             =   0
            Width           =   1110
         End
         Begin VB.OptionButton rbcTapeShowForm 
            Caption         =   "Carted"
            Height          =   210
            Index           =   1
            Left            =   2700
            TabIndex        =   214
            Top             =   -15
            Width           =   855
         End
         Begin VB.Label lacCopy 
            Caption         =   "Copy Tape Show"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   212
            Top             =   -15
            Width           =   1740
         End
      End
      Begin VB.Frame frcCopy 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   206
         Top             =   90
         Width           =   5145
         Begin VB.OptionButton rbcCUseCartNo 
            Caption         =   "Both"
            Height          =   210
            Index           =   2
            Left            =   3450
            TabIndex        =   210
            Top             =   0
            Width           =   750
         End
         Begin VB.OptionButton rbcCUseCartNo 
            Caption         =   "ISCI #'s"
            Height          =   210
            Index           =   1
            Left            =   2535
            TabIndex        =   209
            Top             =   0
            Width           =   945
         End
         Begin VB.OptionButton rbcCUseCartNo 
            Caption         =   "Cart #'s"
            Height          =   210
            Index           =   0
            Left            =   1575
            TabIndex        =   208
            Top             =   0
            Width           =   960
         End
         Begin VB.Label lacCopy 
            Caption         =   "Copy Inventory by"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   207
            Top             =   -15
            Width           =   1530
         End
      End
      Begin VB.Label lacCopy 
         Caption         =   "Audio Type to Exclude"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   248
         Top             =   5190
         Width           =   1935
      End
      Begin VB.Label lacCopy 
         Caption         =   "Last Retrieval Date from vCreative"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   241
         Top             =   4545
         Width           =   3000
      End
   End
   Begin VB.PictureBox plcOptions 
      Height          =   6030
      Left            =   13485
      ScaleHeight     =   5970
      ScaleWidth      =   11235
      TabIndex        =   425
      Top             =   315
      Width           =   11295
      Begin VB.Frame frcOption 
         Caption         =   "System Options"
         Height          =   1500
         Index           =   0
         Left            =   0
         TabIndex        =   426
         Top             =   0
         Width           =   11145
         Begin VB.CheckBox ckcUsingSplitCopy 
            Caption         =   "Split Copy"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   7320
            TabIndex        =   451
            Top             =   1200
            Width           =   1785
         End
         Begin VB.CheckBox ckcAffiliateCRM 
            Caption         =   "Affiliate Mgmt"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   430
            Top             =   960
            Width           =   1515
         End
         Begin VB.CheckBox ckcGUseAffSys 
            Caption         =   "Affiliate System"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   431
            Top             =   1200
            Width           =   1680
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Podcast Spots"
            Height          =   210
            Index           =   29
            Left            =   5400
            TabIndex        =   442
            Top             =   240
            Width           =   1785
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Ad Server"
            Height          =   210
            Index           =   30
            Left            =   120
            TabIndex        =   428
            Top             =   480
            Width           =   1785
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "AirTime"
            Height          =   210
            Index           =   28
            Left            =   1940
            TabIndex        =   432
            Top             =   240
            Width           =   1785
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Advanced Avails"
            Height          =   210
            Index           =   22
            Left            =   120
            TabIndex        =   429
            Top             =   720
            Width           =   1785
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "ACT1 Codes"
            Height          =   210
            Index           =   16
            Left            =   120
            TabIndex        =   427
            Top             =   240
            Width           =   1785
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "E-Mail Distribution"
            Height          =   210
            Index           =   14
            Left            =   1940
            TabIndex        =   435
            Top             =   960
            Width           =   1785
         End
         Begin VB.CheckBox ckcEventRevenue 
            Caption         =   "Event Revenue"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1940
            TabIndex        =   436
            Top             =   1200
            Width           =   1785
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Research"
            Height          =   210
            Index           =   7
            Left            =   7320
            TabIndex        =   447
            Top             =   240
            Width           =   1785
         End
         Begin VB.CheckBox ckcRN_Net 
            Caption         =   "R-N Net"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   7320
            TabIndex        =   448
            Top             =   465
            Width           =   1785
         End
         Begin VB.CheckBox ckcRN_Rep 
            Caption         =   "R-N Rep"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   7320
            TabIndex        =   449
            Top             =   705
            Width           =   1785
         End
         Begin VB.CheckBox ckcDigital 
            Caption         =   "Digital Content"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   10200
            TabIndex        =   456
            Top             =   1080
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.CheckBox ckcInstallment 
            Caption         =   "Installment"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3840
            TabIndex        =   437
            Top             =   240
            Width           =   1425
         End
         Begin VB.CheckBox ckcRemoteImport 
            Caption         =   "Remote Import"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5400
            TabIndex        =   445
            Top             =   945
            Width           =   1785
         End
         Begin VB.CheckBox ckcRemoteExport 
            Caption         =   "Remote Export"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5400
            TabIndex        =   444
            Top             =   720
            Width           =   1785
         End
         Begin VB.CheckBox ckcStrongPassword 
            Caption         =   "Strong Password"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   9240
            TabIndex        =   454
            Top             =   720
            Width           =   1785
         End
         Begin VB.CheckBox ckcRegionalCopy 
            Caption         =   "Regional Copy"
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   10200
            TabIndex        =   458
            Top             =   1320
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox ckcUsingBarter 
            Caption         =   "Barter"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1940
            TabIndex        =   433
            Top             =   480
            Width           =   1785
         End
         Begin VB.CheckBox ckcUsingSplitNetworks 
            Caption         =   "Split Net"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   9240
            TabIndex        =   452
            Top             =   240
            Width           =   1785
         End
         Begin VB.CheckBox ckcUsingRep 
            Caption         =   "Rep"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5400
            TabIndex        =   446
            Top             =   1185
            Width           =   1785
         End
         Begin VB.CheckBox ckcUsingMultiMedia 
            Caption         =   "MultiMedia"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3840
            TabIndex        =   440
            Top             =   960
            Width           =   1425
         End
         Begin VB.CheckBox ckcUsingLiveCopy 
            Caption         =   "Live Copy"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3840
            TabIndex        =   438
            Top             =   480
            Width           =   1425
         End
         Begin VB.CheckBox ckcUsingLiveLog 
            Caption         =   "Live Log"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3840
            TabIndex        =   439
            Top             =   720
            Width           =   1425
         End
         Begin VB.CheckBox ckcUsingSports 
            Caption         =   "Sports"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   9240
            TabIndex        =   453
            Top             =   480
            Width           =   1785
         End
         Begin VB.CheckBox ckcGUsePropSys 
            Caption         =   "Proposal System"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5400
            TabIndex        =   443
            Top             =   480
            Width           =   1785
         End
         Begin VB.CheckBox ckcUsingNTR 
            Caption         =   "NTR"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3840
            TabIndex        =   441
            Top             =   1200
            Width           =   1425
         End
         Begin VB.CheckBox ckcUsingSpecialResearch 
            Caption         =   "Special Research"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   7320
            TabIndex        =   450
            Top             =   960
            Width           =   1785
         End
         Begin VB.CheckBox ckcUsingTraffic 
            Caption         =   "Traffic"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   9240
            TabIndex        =   455
            Top             =   960
            Width           =   1785
         End
         Begin VB.CheckBox ckcUsingBBs 
            Caption         =   "Billboards"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1940
            TabIndex        =   434
            Top             =   720
            Width           =   1785
         End
      End
      Begin VB.Frame frcOption 
         Caption         =   "Import/Export Options"
         Height          =   1520
         Index           =   6
         Left            =   0
         TabIndex        =   604
         Top             =   1480
         Width           =   11145
         Begin VB.CheckBox ckcCntrLineExport 
            Caption         =   "Contract Line Export"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   461
            Top             =   720
            Width           =   2025
         End
         Begin VB.CheckBox ckcWOInvoiceExport 
            Caption         =   "WO Invoice Export"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   9240
            TabIndex        =   605
            Top             =   960
            Width           =   1785
         End
         Begin VB.CheckBox ckcEfficio 
            Caption         =   "Efficio"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2520
            TabIndex        =   465
            Top             =   480
            Width           =   1545
         End
         Begin VB.CheckBox ckcInvoiceExport 
            Caption         =   "Invoice Export"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4200
            TabIndex        =   469
            Top             =   240
            Width           =   1665
         End
         Begin VB.CheckBox ckcVCreative 
            Caption         =   "vCreative Export"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   9240
            TabIndex        =   486
            Top             =   720
            Width           =   1785
         End
         Begin VB.CheckBox ckcUsingMatrix 
            Caption         =   "Tableau-Standard"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   9240
            TabIndex        =   485
            Top             =   480
            Width           =   1785
         End
         Begin VB.CheckBox ckcUsingMatrix 
            Caption         =   "Tableau-Calendar"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   9240
            TabIndex        =   484
            Top             =   240
            Width           =   1785
         End
         Begin VB.CheckBox ckcRevenueExport 
            Caption         =   "Revenue Export"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   7560
            TabIndex        =   480
            Top             =   480
            Width           =   1665
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "RAB-Calendar"
            Height          =   210
            Index           =   23
            Left            =   6000
            TabIndex        =   477
            Top             =   960
            Width           =   1545
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "RAB-Standard"
            Height          =   210
            Index           =   26
            Left            =   6000
            TabIndex        =   478
            Top             =   1185
            Width           =   1545
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "RAB-Cal Spots"
            Height          =   210
            Index           =   27
            Left            =   6000
            TabIndex        =   476
            Top             =   705
            Width           =   1545
         End
         Begin VB.CheckBox ckcUsingMatrix 
            Caption         =   "Matrix-Standard"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   4200
            TabIndex        =   472
            Top             =   960
            Width           =   1665
         End
         Begin VB.CheckBox ckcUsingMatrix 
            Caption         =   "Matrix-Calendar"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   4200
            TabIndex        =   471
            Top             =   720
            Width           =   1665
         End
         Begin VB.CheckBox ckcJelli 
            Caption         =   "Jelli"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4200
            TabIndex        =   470
            Top             =   480
            Width           =   1665
         End
         Begin VB.CheckBox ckcPrefeed 
            Caption         =   "Pre-feed"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   6000
            TabIndex        =   474
            Top             =   240
            Width           =   1425
         End
         Begin VB.CheckBox ckcSalesForce 
            Caption         =   "Sales Force"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   7560
            TabIndex        =   481
            Top             =   720
            Width           =   1665
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Cust-Revenue"
            Height          =   210
            Index           =   31
            Left            =   2520
            TabIndex        =   464
            Top             =   240
            Width           =   1545
         End
         Begin VB.CheckBox ckcGetPaidExport 
            Caption         =   "GetPaid Export"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2520
            TabIndex        =   467
            Top             =   960
            Width           =   1665
         End
         Begin VB.CheckBox ckcGreatPlainGL 
            Caption         =   "Great Plain G/L"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2520
            TabIndex        =   468
            Top             =   1200
            Width           =   1665
         End
         Begin VB.CheckBox ckcUsingRevenue 
            Caption         =   "Corporate Export"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   462
            Top             =   960
            Width           =   2385
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Filemaker"
            Height          =   210
            Index           =   15
            Left            =   2520
            TabIndex        =   466
            Top             =   720
            Width           =   1545
         End
         Begin VB.CheckBox ckcProposalXML 
            Caption         =   "Proposal XML"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   6000
            TabIndex        =   475
            Top             =   480
            Width           =   1545
         End
         Begin VB.CheckBox ckcCompensation 
            Caption         =   "Compensation"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   460
            Top             =   480
            Width           =   2385
         End
         Begin VB.CheckBox ckcMetroSplitCopy 
            Caption         =   "Metro Split Copy"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4200
            TabIndex        =   473
            Top             =   1200
            Width           =   1665
         End
         Begin VB.CheckBox ckcGUseAffFeed 
            Caption         =   "Station Feed"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   7560
            TabIndex        =   482
            Top             =   960
            Width           =   1665
         End
         Begin VB.CheckBox ckcGUseAffSys 
            Caption         =   "RADAR"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   7560
            TabIndex        =   479
            Top             =   240
            Width           =   1545
         End
         Begin VB.CheckBox ckcGUseAffSys 
            Caption         =   "Station Interface"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   7560
            TabIndex        =   483
            Top             =   1200
            Width           =   1665
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "CSV Affidavit Import"
            Height          =   210
            Index           =   25
            Left            =   120
            TabIndex        =   463
            Top             =   1200
            Width           =   2385
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Affidavit Overdue Export"
            Height          =   210
            Index           =   24
            Left            =   120
            TabIndex        =   459
            Top             =   240
            Width           =   2385
         End
      End
      Begin VB.Frame frcOption 
         Caption         =   "Proposal/Order Options"
         Height          =   1040
         Index           =   1
         Left            =   0
         TabIndex        =   457
         Top             =   2980
         Width           =   11145
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Weekly Billing"
            Height          =   285
            Index           =   13
            Left            =   8070
            TabIndex        =   497
            Top             =   480
            Width           =   1515
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Ad Server Tab View Only"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   12
            Left            =   8070
            TabIndex        =   498
            Top             =   720
            Width           =   2370
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Delivery Guarantee GrImp"
            Height          =   285
            Index           =   11
            Left            =   5310
            TabIndex        =   493
            Top             =   240
            Width           =   2490
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Change Billed Prices"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   10
            Left            =   2505
            TabIndex        =   492
            Top             =   720
            Width           =   2160
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Agency Estimate Numbers"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   8
            Left            =   2505
            TabIndex        =   490
            Top             =   240
            Width           =   2550
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Program Exclusions"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   9
            Left            =   150
            TabIndex        =   488
            Top             =   480
            Width           =   2040
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Co-op Billing"
            Height          =   285
            Index           =   6
            Left            =   9330
            TabIndex        =   500
            Top             =   1050
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Calendar Billing"
            Height          =   285
            Index           =   5
            Left            =   8070
            TabIndex        =   496
            Top             =   240
            Width           =   1635
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Delivery Guarantee %"
            Height          =   285
            Index           =   4
            Left            =   2505
            TabIndex        =   491
            Top             =   480
            Width           =   2115
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Revenue Sets"
            Height          =   285
            Index           =   3
            Left            =   150
            TabIndex        =   489
            Top             =   720
            Width           =   1560
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Share"
            Height          =   285
            Index           =   2
            Left            =   5310
            TabIndex        =   495
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Business Category"
            Height          =   285
            Index           =   1
            Left            =   5310
            TabIndex        =   494
            Top             =   480
            Width           =   1950
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Projections"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   487
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame frcOption 
         Height          =   615
         Index           =   4
         Left            =   -15
         TabIndex        =   515
         Top             =   5340
         Width           =   11145
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Show Avail Count"
            Height          =   195
            Index           =   18
            Left            =   150
            TabIndex        =   508
            Top             =   270
            Width           =   1875
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Show CPP Tab"
            Height          =   195
            Index           =   19
            Left            =   2220
            TabIndex        =   509
            Top             =   270
            Width           =   1695
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Programmatic"
            Height          =   195
            Index           =   17
            Left            =   45
            TabIndex        =   507
            Top             =   0
            Width           =   1695
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Show Price Tab"
            Height          =   195
            Index           =   21
            Left            =   6360
            TabIndex        =   511
            Top             =   270
            Width           =   1695
         End
         Begin VB.CheckBox ckcOptionFields 
            Caption         =   "Show CPM Tab"
            Height          =   195
            Index           =   20
            Left            =   4290
            TabIndex        =   510
            Top             =   270
            Width           =   1620
         End
      End
      Begin VB.Frame frcOption 
         Caption         =   "Schedule Line Overrides"
         Height          =   1275
         Index           =   2
         Left            =   0
         TabIndex        =   512
         Top             =   4020
         Width           =   11145
         Begin VB.CheckBox ckcOverrideOptions 
            Caption         =   "Live/Recorded Mandatory"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   5
            Left            =   150
            TabIndex        =   503
            Top             =   960
            Width           =   2640
         End
         Begin VB.TextBox edcLnOverride 
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
            Height          =   285
            Index           =   1
            Left            =   7290
            MaxLength       =   3
            TabIndex        =   506
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox edcLnOverride 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
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
            Height          =   255
            Index           =   0
            Left            =   3750
            MaxLength       =   3
            TabIndex        =   504
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox ckcOverrideOptions 
            Caption         =   "Island Avail Booking"
            Height          =   285
            Index           =   4
            Left            =   150
            TabIndex        =   502
            Top             =   720
            Width           =   2025
         End
         Begin VB.CheckBox ckcOverrideOptions 
            Caption         =   "Preferred Days-Times"
            Height          =   285
            Index           =   3
            Left            =   150
            TabIndex        =   499
            Top             =   240
            Width           =   2235
         End
         Begin VB.CheckBox ckcOverrideOptions 
            Caption         =   "1st Position"
            Height          =   285
            Index           =   2
            Left            =   2910
            TabIndex        =   505
            Top             =   720
            Width           =   1395
         End
         Begin VB.CheckBox ckcOverrideOptions 
            Caption         =   "Allocation %"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   501
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lacLnOverride 
            Caption         =   "Last Week, Solo Avail and 1st Position Priority Index"
            Height          =   240
            Index           =   1
            Left            =   2895
            TabIndex        =   514
            Top             =   990
            Width           =   4545
         End
         Begin VB.Label lacLnOverride 
            Caption         =   "Schedule              % of the Line Spots using Preferred Values"
            Height          =   240
            Index           =   0
            Left            =   2910
            TabIndex        =   513
            Top             =   240
            Width           =   4980
         End
      End
   End
   Begin VB.PictureBox plcGeneral 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   26745
      ScaleHeight     =   5355
      ScaleWidth      =   9045
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   9720
      Width           =   9105
      Begin VB.TextBox edcGRetain 
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
         Height          =   300
         Index           =   9
         Left            =   2430
         MaxLength       =   3
         TabIndex        =   27
         Top             =   3780
         Width           =   450
      End
      Begin VB.TextBox edcGRetain 
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
         Height          =   300
         Index           =   8
         Left            =   3495
         MaxLength       =   3
         TabIndex        =   25
         Top             =   3465
         Width           =   450
      End
      Begin VB.TextBox edcGRetain 
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
         Height          =   300
         Index           =   7
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   23
         Top             =   3135
         Width           =   450
      End
      Begin VB.TextBox edcGRetain 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   7815
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   32
         Top             =   4125
         Width           =   990
      End
      Begin VB.TextBox edcGRetain 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   5790
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   31
         Top             =   4125
         Width           =   990
      End
      Begin VB.TextBox edcGRetain 
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
         Height          =   300
         Index           =   3
         Left            =   1980
         MaxLength       =   3
         TabIndex        =   19
         Top             =   2415
         Width           =   450
      End
      Begin VB.TextBox edcGRetain 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2085
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   29
         Top             =   4125
         Width           =   990
      End
      Begin VB.Frame frcSystemType 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   1455
         Width           =   6075
         Begin VB.OptionButton rbcSystemType 
            Caption         =   "Radio Stations"
            Height          =   210
            Index           =   1
            Left            =   3615
            TabIndex        =   13
            Top             =   0
            Width           =   1530
         End
         Begin VB.OptionButton rbcSystemType 
            Caption         =   "Network/Syndication"
            Height          =   210
            Index           =   0
            Left            =   1530
            TabIndex        =   12
            Top             =   0
            Width           =   2010
         End
         Begin VB.Label lacGen 
            Caption         =   "System Used For"
            Height          =   240
            Index           =   2
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.TextBox edcGClientAbbr 
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
         Height          =   300
         Left            =   2265
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1080
         Width           =   1380
      End
      Begin VB.CheckBox ckcSDelivery 
         Caption         =   "Using Delivery Vehicles"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2685
         TabIndex        =   41
         Top             =   5160
         Width           =   2340
      End
      Begin VB.CheckBox ckcSSelling 
         Caption         =   "Using Selling Vehicles"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   40
         Top             =   5160
         Width           =   2220
      End
      Begin VB.TextBox edcGAlertInterval 
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
         Height          =   300
         Left            =   6585
         MaxLength       =   3
         TabIndex        =   39
         Top             =   4770
         Width           =   615
      End
      Begin VB.TextBox edcGRetainPassword 
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
         Height          =   300
         Left            =   2430
         MaxLength       =   4
         TabIndex        =   37
         Top             =   4770
         Width           =   615
      End
      Begin VB.TextBox edcGRetain 
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
         Height          =   300
         Index           =   5
         Left            =   3225
         MaxLength       =   3
         TabIndex        =   21
         Top             =   2775
         Width           =   450
      End
      Begin VB.TextBox edcGRetain 
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
         Height          =   300
         Index           =   1
         Left            =   2265
         MaxLength       =   3
         TabIndex        =   15
         Top             =   1725
         Width           =   450
      End
      Begin VB.TextBox edcGRetain 
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
         Height          =   300
         Index           =   2
         Left            =   2100
         MaxLength       =   3
         TabIndex        =   17
         Top             =   2070
         Width           =   450
      End
      Begin VB.TextBox edcGAddr 
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
         Height          =   300
         Index           =   2
         Left            =   5610
         MaxLength       =   25
         TabIndex        =   7
         Top             =   765
         Width           =   3210
      End
      Begin VB.TextBox edcGAddr 
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
         Height          =   300
         Index           =   1
         Left            =   5610
         MaxLength       =   25
         TabIndex        =   6
         Top             =   435
         Width           =   3210
      End
      Begin VB.TextBox edcGClient 
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
         Height          =   300
         Left            =   1215
         MaxLength       =   25
         TabIndex        =   4
         Top             =   90
         Width           =   3210
      End
      Begin VB.TextBox edcGAddr 
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
         Height          =   300
         Index           =   0
         Left            =   5610
         MaxLength       =   25
         TabIndex        =   5
         Top             =   90
         Width           =   3210
      End
      Begin VB.PictureBox plcGTBar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   4845
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4485
         Width           =   4845
         Begin VB.OptionButton rbcGTBar 
            Caption         =   "Client Name"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2115
            TabIndex        =   34
            Top             =   0
            Width           =   1290
         End
         Begin VB.OptionButton rbcGTBar 
            Caption         =   "Signon Name"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3435
            TabIndex        =   35
            Top             =   0
            Width           =   1380
         End
      End
      Begin VB.Label lacGRetain 
         Appearance      =   0  'Flat
         Caption         =   "Retain User Activity Log for              Days"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   3840
         Width           =   7470
      End
      Begin VB.Label lacGRetain 
         Appearance      =   0  'Flat
         Caption         =   "Retain Payment and Revenue History for              Broadcast Months from End of Last Invoicing"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   3510
         Width           =   7470
      End
      Begin VB.Label lacGRetain 
         Appearance      =   0  'Flat
         Caption         =   "Retain Expired Contracts for              Broadcast Months from End of Last Invoicing"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   3180
         Width           =   7470
      End
      Begin VB.Label lacGRetain 
         Appearance      =   0  'Flat
         Caption         =   "Earliest Spot Date: Traffic                              Affiliate"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   3615
         TabIndex        =   30
         Top             =   4170
         Width           =   4230
      End
      Begin VB.Label lacGRetain 
         Appearance      =   0  'Flat
         Caption         =   "Retain Projections for              Broadcast Months from End of Last Invoicing"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   2460
         Width           =   6750
      End
      Begin VB.Label lacGRetain 
         Appearance      =   0  'Flat
         Caption         =   "Last Date Archive Run"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   4170
         Width           =   1905
      End
      Begin VB.Label lacGen 
         Appearance      =   0  'Flat
         Caption         =   "Client Name Abbreviation"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1125
         Width           =   5415
      End
      Begin VB.Label lacGen 
         Appearance      =   0  'Flat
         Caption         =   "Check for New Alerts every                  Minutes"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   4215
         TabIndex        =   38
         Top             =   4800
         Width           =   4110
      End
      Begin VB.Label lacGen 
         Appearance      =   0  'Flat
         Caption         =   "# Days to retain Password"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   36
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Label lacGRetain 
         Appearance      =   0  'Flat
         Caption         =   "Retain All Proposals Entered prior to              Broadcast Months from End of Last Invoicing"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   2820
         Width           =   7470
      End
      Begin VB.Label lacGRetain 
         Appearance      =   0  'Flat
         Caption         =   "Retain Affiliate Spots for               Broadcast Months from End of Last Invoicing"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1755
         Width           =   6435
      End
      Begin VB.Label lacGRetain 
         Appearance      =   0  'Flat
         Caption         =   "Retain Traffic Spots for              Broadcast Months from End of Last Invoicing"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   2115
         Width           =   6750
      End
      Begin VB.Label lacGen 
         Appearance      =   0  'Flat
         Caption         =   "Client Name                                                                               Address"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   135
         Width           =   5415
      End
   End
   Begin VB.PictureBox plcAccount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5715
      Left            =   15660
      ScaleHeight     =   5655
      ScaleWidth      =   10920
      TabIndex        =   359
      TabStop         =   0   'False
      Top             =   14655
      Width           =   10980
      Begin VB.CheckBox ckcRUseTMP 
         Caption         =   "Compress Transactions"
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
         Index           =   6
         Left            =   6540
         TabIndex        =   411
         Top             =   5445
         Width           =   2505
      End
      Begin VB.CheckBox ckcOverrideOptions 
         Caption         =   "NTR Acquisition Cost"
         Height          =   210
         Index           =   1
         Left            =   3855
         TabIndex        =   410
         Top             =   5445
         Width           =   2340
      End
      Begin VB.CheckBox ckcRUseTMP 
         Caption         =   "Station Payment on Collection"
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
         Index           =   5
         Left            =   120
         TabIndex        =   409
         Top             =   5445
         Width           =   3240
      End
      Begin VB.CheckBox ckcRUseTMP 
         Caption         =   "Acquisition Commissionable for Conventional and Selling Vehicles"
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
         Index           =   4
         Left            =   120
         TabIndex        =   408
         Top             =   5160
         Width           =   7350
      End
      Begin VB.CheckBox ckcRUseTMP 
         Caption         =   "Cutoff Proposals/Orders when Credit Limit reached (Unchecked-warning only)"
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
         Left            =   120
         TabIndex        =   382
         Top             =   2310
         Width           =   7320
      End
      Begin VB.PictureBox plcInstRev 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   8460
         TabIndex        =   398
         TabStop         =   0   'False
         Top             =   3825
         Width           =   8460
         Begin VB.OptionButton rbcInstRev 
            Caption         =   "Invoiced (Inv is Rev)"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2085
            TabIndex        =   399
            Top             =   0
            Width           =   2025
         End
         Begin VB.OptionButton rbcInstRev 
            Caption         =   "Aired (Separate Inv from Rev)"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   4275
            TabIndex        =   400
            Top             =   0
            Width           =   2790
         End
      End
      Begin VB.PictureBox plcMerchPromo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   5775
         TabIndex        =   395
         TabStop         =   0   'False
         Top             =   3570
         Width           =   5775
         Begin VB.OptionButton rbcRMerchPromo 
            Caption         =   "Percent"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   3720
            TabIndex        =   396
            Top             =   0
            Width           =   945
         End
         Begin VB.OptionButton rbcRMerchPromo 
            Caption         =   "Dollars"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   4785
            TabIndex        =   397
            Top             =   0
            Width           =   930
         End
      End
      Begin VB.TextBox edcBarterLPD 
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
         Height          =   300
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   381
         Top             =   1950
         Width           =   1260
      End
      Begin VB.ComboBox cbcReconGroup 
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
         ItemData        =   "Siteopt.frx":08CA
         Left            =   6840
         List            =   "Siteopt.frx":08E3
         TabIndex        =   550
         Top             =   1275
         Width           =   1500
      End
      Begin VB.CheckBox ckcRUseTMP 
         Caption         =   "Using Promotion Receivables"
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
         Left            =   6135
         TabIndex        =   394
         Top             =   3300
         Width           =   3075
      End
      Begin VB.CheckBox ckcRUseTMP 
         Caption         =   "Using Merchandising Receivables"
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
         Left            =   2760
         TabIndex        =   393
         Top             =   3300
         Width           =   3240
      End
      Begin VB.CheckBox ckcRUseTMP 
         Caption         =   "Using Trade Receivables"
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
         Left            =   120
         TabIndex        =   392
         Top             =   3300
         Width           =   2745
      End
      Begin VB.PictureBox plcRUnbilled 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   2340
         TabIndex        =   387
         TabStop         =   0   'False
         Top             =   3045
         Width           =   2340
         Begin VB.OptionButton rbcRUnbilled 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   1725
            TabIndex        =   389
            Top             =   0
            Width           =   540
         End
         Begin VB.OptionButton rbcRUnbilled 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1080
            TabIndex        =   388
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox plcRCurrAR 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   2340
         TabIndex        =   384
         TabStop         =   0   'False
         Top             =   2820
         Width           =   2340
         Begin VB.OptionButton rbcRCurrAmt 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   1725
            TabIndex        =   386
            Top             =   0
            Width           =   510
         End
         Begin VB.OptionButton rbcRCurrAmt 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1080
            TabIndex        =   385
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.TextBox edcRCreditDate 
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
         Height          =   300
         Left            =   5535
         MaxLength       =   10
         TabIndex        =   404
         Top             =   4410
         Width           =   1260
      End
      Begin VB.TextBox edcRNewCntr 
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
         Height          =   300
         Left            =   5505
         MaxLength       =   3
         TabIndex        =   391
         Top             =   2805
         Width           =   510
      End
      Begin VB.TextBox edcRCollectContact 
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
         Height          =   300
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   406
         Top             =   4755
         Width           =   3210
      End
      Begin MSMask.MaskEdBox mkcRCollectPhoneNo 
         Height          =   315
         Left            =   5910
         TabIndex        =   407
         Tag             =   "The number and extension of the buyer."
         Top             =   4755
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16776960
         ForeColor       =   0
         MaxLength       =   24
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(AAA) AAA-AAAA Ext(AAAA)"
         PromptChar      =   "_"
      End
      Begin VB.TextBox edcAPenny 
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
         Height          =   300
         Left            =   2190
         MaxLength       =   6
         TabIndex        =   379
         Top             =   1950
         Width           =   1065
      End
      Begin VB.TextBox edcRB 
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
         Height          =   300
         Left            =   4800
         MaxLength       =   12
         TabIndex        =   377
         Top             =   1605
         Width           =   1545
      End
      Begin VB.TextBox edcRNRP 
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
         Height          =   300
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   375
         Top             =   1275
         Width           =   1260
      End
      Begin VB.TextBox edcRCRP 
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
         Height          =   300
         Left            =   7530
         MaxLength       =   10
         TabIndex        =   373
         Top             =   930
         Width           =   1260
      End
      Begin VB.TextBox edcRPRP 
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
         Height          =   300
         Left            =   3075
         MaxLength       =   10
         TabIndex        =   372
         Top             =   930
         Width           =   1260
      End
      Begin VB.PictureBox plcRRP 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   7740
         TabIndex        =   367
         TabStop         =   0   'False
         Top             =   690
         Width           =   7740
         Begin VB.OptionButton rbcRRP 
            Caption         =   "Corporate Month"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   4980
            TabIndex        =   370
            Top             =   0
            Width           =   1710
         End
         Begin VB.OptionButton rbcRRP 
            Caption         =   "Broadcast Month"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1650
            TabIndex        =   368
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton rbcRRP 
            Caption         =   "Calendar Month"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3345
            TabIndex        =   369
            Top             =   0
            Width           =   1605
         End
      End
      Begin VB.PictureBox plcRCorpCal 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   135
         ScaleHeight     =   255
         ScaleWidth      =   2925
         TabIndex        =   360
         TabStop         =   0   'False
         Top             =   90
         Width           =   2925
         Begin VB.OptionButton rbcRCorpCal 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2205
            TabIndex        =   362
            Top             =   0
            Width           =   1485
         End
         Begin VB.OptionButton rbcRCorpCal 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1680
            TabIndex        =   361
            Top             =   0
            Width           =   510
         End
      End
      Begin VB.TextBox edcRPctCredit 
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
         Height          =   300
         Left            =   7230
         MaxLength       =   7
         TabIndex        =   366
         Top             =   345
         Width           =   885
      End
      Begin VB.Label lacAcct 
         Appearance      =   0  'Flat
         Caption         =   "Barter Last Paid Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   3420
         TabIndex        =   380
         Top             =   1980
         Width           =   1845
      End
      Begin VB.Label lacAcct 
         Appearance      =   0  'Flat
         Caption         =   "Reconciliation Vehicle Group"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   4320
         TabIndex        =   551
         Top             =   1320
         Width           =   2520
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   19
         Left            =   2625
         TabIndex        =   365
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label lacLastPurgedDate 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1725
         TabIndex        =   402
         Top             =   4095
         Width           =   1140
      End
      Begin VB.Label lacLastPurgedDate 
         Caption         =   "Last Purged Date"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   401
         Top             =   4125
         Width           =   1470
      End
      Begin VB.Label lacAcct 
         Appearance      =   0  'Flat
         Caption         =   "Highest Penny Variance"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   378
         Top             =   1980
         Width           =   2040
      End
      Begin VB.Label lacAcct 
         Appearance      =   0  'Flat
         Caption         =   "Date that the Advertisers and Agencies Unbilled Value Computed"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   403
         Top             =   4455
         Width           =   5430
      End
      Begin VB.Label lacAcct 
         Appearance      =   0  'Flat
         Caption         =   "New and Existing Contracts for                 Weeks"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   2895
         TabIndex        =   390
         Top             =   2835
         Width           =   4230
      End
      Begin VB.Label lacAcct 
         Appearance      =   0  'Flat
         Caption         =   "Collection Contact                                                                               Phone #"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   405
         Top             =   4785
         Width           =   5820
      End
      Begin VB.Label lacAcct 
         Appearance      =   0  'Flat
         Caption         =   "Receivables Balance, End of Previous Reconciling Period                         Highest Penny Variance"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   376
         Top             =   1650
         Width           =   4815
      End
      Begin VB.Label lacAcct 
         Appearance      =   0  'Flat
         Caption         =   "End of Next Reconciling Period"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   374
         Top             =   1305
         Width           =   2505
      End
      Begin VB.Label lacAcct 
         Appearance      =   0  'Flat
         Caption         =   "End of Previous Reconciling Period                                     End of Current Reconciling Period"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   371
         Top             =   960
         Width           =   7425
      End
      Begin VB.Label lacAcct 
         Appearance      =   0  'Flat
         Caption         =   "When computing Credit Limit, include:"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   383
         Top             =   2595
         Width           =   3285
      End
      Begin VB.Label lacAcct 
         Appearance      =   0  'Flat
         Caption         =   "Date of last payment received                                % of Credit limit to trigger attention"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   135
         TabIndex        =   364
         Top             =   375
         Width           =   7005
      End
   End
   Begin VB.PictureBox plcSales 
      Height          =   5490
      Left            =   570
      ScaleHeight     =   5430
      ScaleWidth      =   11190
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   15000
      Width           =   11250
      Begin VB.Frame frcOption 
         Caption         =   "Insertion Order"
         Height          =   765
         Index           =   5
         Left            =   120
         TabIndex        =   67
         Top             =   2595
         Width           =   9570
         Begin VB.CheckBox ckcSales 
            Caption         =   "Include Monthly Billed Summary"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   18
            Left            =   5295
            TabIndex        =   72
            Top             =   240
            Width           =   3270
         End
         Begin VB.PictureBox plcSSale 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   4
            Left            =   90
            ScaleHeight     =   240
            ScaleWidth      =   5385
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   225
            Width           =   5385
            Begin VB.OptionButton rbcInsertAddr 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3900
               TabIndex        =   71
               Top             =   0
               Width           =   1080
            End
            Begin VB.OptionButton rbcInsertAddr 
               Caption         =   "Site (Invoices)"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2340
               TabIndex        =   70
               Top             =   0
               Width           =   1515
            End
            Begin VB.OptionButton rbcInsertAddr 
               Caption         =   "Payee"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   1440
               TabIndex        =   69
               Top             =   0
               Width           =   870
            End
         End
         Begin VB.CheckBox ckcSales 
            Caption         =   "Show Spot Prices for $0 Acquisition Costs"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   73
            Top             =   480
            Width           =   3990
         End
         Begin VB.CheckBox ckcSales 
            Caption         =   "Suppress Acquisition Net and Commission"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   17
            Left            =   5295
            TabIndex        =   74
            Top             =   480
            Width           =   4080
         End
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Open Avail Stripe on Unsold Avails"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   5400
         TabIndex        =   592
         Top             =   615
         Width           =   3465
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Wordwrap Vehicle Name on Form Reports"
         Height          =   210
         Index           =   12
         Left            =   5400
         TabIndex        =   591
         Top             =   1185
         Width           =   3930
      End
      Begin VB.PictureBox plcSSale 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   6
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   5850
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   3675
         Width           =   5850
         Begin VB.OptionButton rbcEqualize 
            Caption         =   "None"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   4350
            TabIndex        =   81
            Top             =   0
            Width           =   780
         End
         Begin VB.OptionButton rbcEqualize 
            Caption         =   "60"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3675
            TabIndex        =   80
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton rbcEqualize 
            Caption         =   "30"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   3000
            TabIndex        =   79
            Top             =   0
            Width           =   645
         End
      End
      Begin VB.Frame frcOption 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   95
         Top             =   5190
         Width           =   5445
         Begin VB.OptionButton rbcNewBusYear 
            Caption         =   "Rolling"
            Height          =   210
            Index           =   1
            Left            =   3165
            TabIndex        =   98
            Top             =   0
            Width           =   945
         End
         Begin VB.OptionButton rbcNewBusYear 
            Caption         =   "Calendar"
            Height          =   210
            Index           =   0
            Left            =   2055
            TabIndex        =   97
            Top             =   0
            Width           =   1245
         End
         Begin VB.Label lacSales 
            Appearance      =   0  'Flat
            Caption         =   "Base New Business on                                                 Year"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   0
            TabIndex        =   96
            Top             =   0
            Width           =   4590
         End
      End
      Begin VB.TextBox edcSales 
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
         Height          =   300
         Index           =   8
         Left            =   7530
         MaxLength       =   2
         TabIndex        =   94
         Top             =   4845
         Width           =   600
      End
      Begin VB.TextBox edcSales 
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
         Height          =   300
         Index           =   7
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   92
         Top             =   4845
         Width           =   600
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Prohibit Rescheduling Across Calendar Months"
         Height          =   210
         Index           =   11
         Left            =   120
         TabIndex        =   90
         Top             =   4545
         Width           =   4665
      End
      Begin VB.TextBox edcSales 
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
         Height          =   300
         Index           =   6
         Left            =   7215
         MaxLength       =   3
         TabIndex        =   89
         Top             =   4185
         Width           =   600
      End
      Begin VB.TextBox edcSales 
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
         Height          =   300
         Index           =   5
         Left            =   6465
         MaxLength       =   3
         TabIndex        =   88
         Top             =   4185
         Width           =   600
      End
      Begin VB.TextBox edcSales 
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
         Height          =   300
         Index           =   4
         Left            =   5715
         MaxLength       =   3
         TabIndex        =   87
         Top             =   4185
         Width           =   600
      End
      Begin VB.TextBox edcSales 
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
         Height          =   300
         Index           =   3
         Left            =   4965
         MaxLength       =   3
         TabIndex        =   86
         Top             =   4185
         Width           =   600
      End
      Begin VB.TextBox edcSales 
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
         Height          =   300
         Index           =   2
         Left            =   4215
         MaxLength       =   3
         TabIndex        =   85
         Top             =   4185
         Width           =   600
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Allow Spot Moves on Today"
         Height          =   210
         Index           =   3
         Left            =   3825
         TabIndex        =   557
         Top             =   885
         Width           =   2655
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Allow Mixture of Air Time and Rep on Same Contract"
         Height          =   210
         Index           =   13
         Left            =   4140
         TabIndex        =   83
         Top             =   3930
         Visible         =   0   'False
         Width           =   4665
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Hub-based"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   2415
         TabIndex        =   61
         Top             =   2130
         Width           =   1305
      End
      Begin VB.PictureBox plcBPkageGenMeth 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6855
         ScaleHeight     =   240
         ScaleWidth      =   4245
         TabIndex        =   554
         TabStop         =   0   'False
         Top             =   1515
         Visible         =   0   'False
         Width           =   4245
         Begin VB.OptionButton rbcBPkageGenMeth 
            Caption         =   "Line"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3330
            TabIndex        =   556
            Top             =   0
            Width           =   705
         End
         Begin VB.OptionButton rbcBPkageGenMeth 
            Caption         =   "Week"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2535
            TabIndex        =   555
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.PictureBox plcSSale 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   3720
         ScaleHeight     =   240
         ScaleWidth      =   4410
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1815
         Width           =   4410
         Begin VB.CheckBox ckcSales 
            Caption         =   "Virtual"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   1680
            TabIndex        =   57
            Top             =   -15
            Width           =   900
         End
         Begin VB.CheckBox ckcSales 
            Caption         =   "Real"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   6
            Left            =   2550
            TabIndex        =   58
            Top             =   -15
            Width           =   750
         End
         Begin VB.CheckBox ckcSales 
            Caption         =   "Equal"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   7
            Left            =   3345
            TabIndex        =   59
            Top             =   -15
            Visible         =   0   'False
            Width           =   750
         End
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Ignore Selling to Airing Conflicts"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   553
         Top             =   885
         Width           =   3060
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Send Billboard Spots to Affiliate System"
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   82
         Top             =   3930
         Width           =   3705
      End
      Begin VB.PictureBox plcSSale 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   5235
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   3405
         Width           =   5235
         Begin VB.OptionButton rbcUnitOr3060 
            Caption         =   "Units"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2100
            TabIndex        =   76
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton rbcUnitOr3060 
            Caption         =   "30/60"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3000
            TabIndex        =   77
            Top             =   0
            Width           =   750
         End
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Allow Daily Buys"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   4410
         TabIndex        =   62
         Top             =   2130
         Width           =   2130
      End
      Begin VB.PictureBox plcSSale 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   5850
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   2400
         Width           =   5850
         Begin VB.OptionButton rbcPLMove 
            Caption         =   "MG's"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2520
            TabIndex        =   64
            Top             =   0
            Width           =   720
         End
         Begin VB.OptionButton rbcPLMove 
            Caption         =   "Outsides"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3270
            TabIndex        =   65
            Top             =   0
            Width           =   1050
         End
         Begin VB.OptionButton rbcPLMove 
            Caption         =   "Ask"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   4335
            TabIndex        =   66
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.CheckBox ckcSMktBase 
         Caption         =   "Market-based"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   2130
         Width           =   2130
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Retain Spot Screen Date when Vehicle Switched"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   615
         Width           =   4710
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "On Spot Screen, Allow Moves to Violate Contract Parameters"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   5790
      End
      Begin VB.TextBox edcSales 
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
         Height          =   300
         Index           =   1
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   55
         Top             =   1785
         Width           =   600
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Remoter Users Allowed"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   9060
         TabIndex        =   53
         Top             =   5130
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.PictureBox plcSSale 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   135
         ScaleHeight     =   240
         ScaleWidth      =   7170
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1515
         Width           =   7170
         Begin VB.OptionButton rbcSEnterAge 
            Caption         =   "Entered Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   3990
            TabIndex        =   51
            Top             =   0
            Width           =   1350
         End
         Begin VB.OptionButton rbcSEnterAge 
            Caption         =   "Aged Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   5415
            TabIndex        =   52
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.TextBox edcSales 
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
         Height          =   300
         Index           =   0
         Left            =   3660
         MaxLength       =   5
         TabIndex        =   49
         Top             =   1155
         Width           =   855
      End
      Begin VB.PictureBox plcSSale 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   4920
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   105
         Width           =   4920
         Begin VB.OptionButton rbcSUseProd 
            Caption         =   "Advertiser, Product"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   2940
            TabIndex        =   45
            Top             =   0
            Width           =   1950
         End
         Begin VB.OptionButton rbcSUseProd 
            Caption         =   "Short Title"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1740
            TabIndex        =   44
            Top             =   0
            Width           =   1155
         End
      End
      Begin VB.Label lacSales 
         Appearance      =   0  'Flat
         Caption         =   "# Months New Business remains New"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   4260
         TabIndex        =   93
         Top             =   4875
         Width           =   3345
      End
      Begin VB.Label lacSales 
         Appearance      =   0  'Flat
         Caption         =   "# Months off Before considered New Business"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   91
         Top             =   4875
         Width           =   3225
      End
      Begin VB.Label lacSales 
         Appearance      =   0  'Flat
         Caption         =   "Event (Sport) Avails Report Default Spot Length"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   84
         Top             =   4230
         Width           =   4050
      End
      Begin VB.Label lacSales 
         Appearance      =   0  'Flat
         Caption         =   "Vehicle Name Length"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   1815
         Width           =   1860
      End
      Begin VB.Label lacSales 
         Appearance      =   0  'Flat
         Caption         =   "On Reports Show Dollars as (1, 10, 100,..)"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   1185
         Width           =   3480
      End
   End
   Begin VB.PictureBox plcInv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   14970
      ScaleHeight     =   5835
      ScaleWidth      =   11700
      TabIndex        =   262
      TabStop         =   0   'False
      Top             =   6825
      Width           =   11760
      Begin VB.PictureBox plcFlatRate 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   9120
         ScaleHeight     =   495
         ScaleWidth      =   2415
         TabIndex        =   619
         Top             =   4755
         Width           =   2415
         Begin VB.OptionButton rbcFlatRateAverageFormula 
            Caption         =   "Daily"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   621
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton rbcFlatRateAverageFormula 
            Caption         =   "Monthly"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   620
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Flat Rate Average Formula"
            Height          =   255
            Left            =   0
            TabIndex        =   558
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Selective"
         Height          =   210
         Index           =   12
         Left            =   3840
         TabIndex        =   618
         Top             =   5610
         Width           =   1155
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Bill Over-Delivered CPM Impressions"
         Height          =   210
         Index           =   11
         Left            =   8355
         TabIndex        =   358
         Top             =   5370
         Width           =   3345
      End
      Begin VB.CheckBox ckcPodShowWk 
         Caption         =   "For Mixed Packages of AirTime and Podcast Spots, Show  Packages as ""Week of"".  If Unchecked, Date/Time will show"
         Height          =   255
         Left            =   120
         TabIndex        =   294
         Top             =   2120
         Width           =   9975
      End
      Begin VB.OptionButton rbcInvEmail 
         Caption         =   "NTR only"
         Height          =   210
         Index           =   2
         Left            =   8640
         TabIndex        =   603
         Top             =   5610
         Width           =   1095
      End
      Begin VB.OptionButton rbcInvEmail 
         Caption         =   "air time only"
         Height          =   210
         Index           =   1
         Left            =   7320
         TabIndex        =   602
         Top             =   5610
         Width           =   1335
      End
      Begin VB.OptionButton rbcInvEmail 
         Caption         =   "air time and NTR"
         Height          =   210
         Index           =   0
         Left            =   5640
         TabIndex        =   601
         Top             =   5610
         Width           =   1815
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Automatic"
         Height          =   210
         Index           =   10
         Left            =   2520
         TabIndex        =   357
         Top             =   5610
         Width           =   1395
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Activate"
         Height          =   210
         Index           =   9
         Left            =   1425
         TabIndex        =   356
         Top             =   5610
         Width           =   1275
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Require Complete Station Posting Prior to Agency Invoices"
         Height          =   210
         Index           =   8
         Left            =   3075
         TabIndex        =   354
         Top             =   5340
         Width           =   5265
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Bill Spots X-Mid in Aired Month"
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   353
         Top             =   5340
         Width           =   2940
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Print Final EDI Invoices"
         Height          =   210
         Index           =   6
         Left            =   6840
         TabIndex        =   352
         Top             =   5085
         Width           =   2415
      End
      Begin VB.TextBox edcBBillDate 
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
         Height          =   300
         Index           =   3
         Left            =   5955
         MaxLength       =   10
         TabIndex        =   287
         Top             =   1275
         Width           =   870
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Date/Time,"
         Height          =   210
         Index           =   5
         Left            =   5475
         TabIndex        =   333
         Top             =   4245
         Width           =   1185
      End
      Begin VB.PictureBox plcInvSpotTimeZone 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   135
         ScaleHeight     =   240
         ScaleWidth      =   6105
         TabIndex        =   346
         Top             =   5070
         Width           =   6105
         Begin VB.OptionButton rbcInvSpotTimeZone 
            Caption         =   "N/A"
            Height          =   210
            Index           =   4
            Left            =   5280
            TabIndex        =   351
            Top             =   0
            Width           =   570
         End
         Begin VB.OptionButton rbcInvSpotTimeZone 
            Caption         =   "PT"
            Height          =   210
            Index           =   3
            Left            =   4650
            TabIndex        =   350
            Top             =   0
            Width           =   570
         End
         Begin VB.OptionButton rbcInvSpotTimeZone 
            Caption         =   "MT"
            Height          =   210
            Index           =   2
            Left            =   4020
            TabIndex        =   349
            Top             =   0
            Width           =   570
         End
         Begin VB.OptionButton rbcInvSpotTimeZone 
            Caption         =   "CT"
            Height          =   210
            Index           =   1
            Left            =   3390
            TabIndex        =   348
            Top             =   0
            Width           =   570
         End
         Begin VB.OptionButton rbcInvSpotTimeZone 
            Caption         =   "ET"
            Height          =   210
            Index           =   0
            Left            =   2760
            TabIndex        =   347
            Top             =   0
            Width           =   570
         End
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Combine Commerical and NTR"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   4155
         TabIndex        =   312
         Top             =   2925
         Width           =   3015
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Invoice Selective Vehicles"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7200
         TabIndex        =   313
         Top             =   2925
         Width           =   2490
      End
      Begin VB.PictureBox plcInvSortBy 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   3045
         TabIndex        =   338
         TabStop         =   0   'False
         Top             =   4500
         Width           =   3045
         Begin VB.OptionButton rbcInvSortBy 
            Caption         =   "Payee"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   825
            TabIndex        =   339
            Top             =   0
            Width           =   825
         End
         Begin VB.OptionButton rbcInvSortBy 
            Caption         =   "Vehicle"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1830
            TabIndex        =   340
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.CheckBox ckcTaxOn 
         Caption         =   "NTR Tax"
         Height          =   210
         Index           =   1
         Left            =   7110
         TabIndex        =   325
         Top             =   3660
         Width           =   1095
      End
      Begin VB.CheckBox ckcTaxOn 
         Caption         =   "Commercial Tax"
         Height          =   210
         Index           =   0
         Left            =   5100
         TabIndex        =   324
         Top             =   3660
         Width           =   1875
      End
      Begin VB.PictureBox plcTaxOnAirTime 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   4470
         TabIndex        =   320
         TabStop         =   0   'False
         Top             =   3645
         Width           =   4470
         Begin VB.OptionButton rbcTaxOnAirTime 
            Caption         =   "None"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1410
            TabIndex        =   321
            Top             =   0
            Width           =   720
         End
         Begin VB.OptionButton rbcTaxOnAirTime 
            Caption         =   "USA"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2310
            TabIndex        =   322
            Top             =   0
            Width           =   690
         End
         Begin VB.OptionButton rbcTaxOnAirTime 
            Caption         =   "Canadian"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   3075
            TabIndex        =   323
            Top             =   -15
            Width           =   1110
         End
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Lock Box by Vehicle instead of Payee"
         Height          =   210
         Index           =   3
         Left            =   4215
         TabIndex        =   341
         Top             =   4500
         Width           =   3585
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "Day is Complete testing"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   342
         Top             =   4755
         Width           =   2340
      End
      Begin VB.PictureBox plcInvISCIForm 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2235
         ScaleHeight     =   195
         ScaleWidth      =   5325
         TabIndex        =   303
         TabStop         =   0   'False
         Top             =   2670
         Width           =   5325
         Begin VB.OptionButton rbcInvISCIForm 
            Caption         =   "Wrap Around"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   3795
            TabIndex        =   306
            Top             =   0
            Width           =   1425
         End
         Begin VB.OptionButton rbcInvISCIForm 
            Caption         =   "Truncate Left"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   2295
            TabIndex        =   305
            Top             =   0
            Width           =   1470
         End
         Begin VB.OptionButton rbcInvISCIForm 
            Caption         =   "Truncate Right"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   304
            Top             =   0
            Width           =   1515
         End
      End
      Begin VB.CheckBox ckcInv 
         Caption         =   "BBs on Same Line"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   302
         Top             =   2655
         Width           =   1935
      End
      Begin VB.TextBox edcBLogoSpaces 
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
         Height          =   300
         Index           =   1
         Left            =   8400
         MaxLength       =   1
         TabIndex        =   345
         Top             =   4755
         Width           =   465
      End
      Begin VB.TextBox edcBLogoSpaces 
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
         Height          =   300
         Index           =   0
         Left            =   5205
         MaxLength       =   1
         TabIndex        =   344
         Top             =   4755
         Width           =   465
      End
      Begin VB.PictureBox plcPostRepAffidavit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4350
         ScaleHeight     =   195
         ScaleWidth      =   6525
         TabIndex        =   332
         TabStop         =   0   'False
         Top             =   4275
         Width           =   6525
         Begin VB.OptionButton rbcPostRepAffidavit 
            Caption         =   "None"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   5760
            TabIndex        =   337
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton rbcPostRepAffidavit 
            Caption         =   "Wkly"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   5025
            TabIndex        =   336
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton rbcPostRepAffidavit 
            Caption         =   "Cal"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   4410
            TabIndex        =   335
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton rbcPostRepAffidavit 
            Caption         =   "Std"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   3765
            TabIndex        =   334
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox plcPrintRepInv 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   4005
         TabIndex        =   329
         TabStop         =   0   'False
         Top             =   4275
         Width           =   4005
         Begin VB.OptionButton rbcPrintRepInv 
            Caption         =   "Market"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1830
            TabIndex        =   330
            Top             =   0
            Width           =   900
         End
         Begin VB.OptionButton rbcPrintRepInv 
            Caption         =   "Vehicle"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2820
            TabIndex        =   331
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.TextBox edcBBillDate 
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
         Height          =   300
         Index           =   2
         Left            =   9015
         MaxLength       =   10
         TabIndex        =   288
         Top             =   1275
         Width           =   870
      End
      Begin VB.TextBox edcBTTax 
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
         Height          =   300
         Index           =   1
         Left            =   5550
         MaxLength       =   30
         TabIndex        =   328
         Top             =   3915
         Width           =   2505
      End
      Begin VB.TextBox edcBTTax 
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
         Height          =   300
         Index           =   0
         Left            =   1245
         MaxLength       =   30
         TabIndex        =   327
         Top             =   3915
         Width           =   2505
      End
      Begin VB.TextBox edcInvExportId 
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
         Height          =   300
         Left            =   8430
         MaxLength       =   4
         TabIndex        =   308
         Top             =   2595
         Width           =   450
      End
      Begin VB.PictureBox plcBMissedDT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   4455
         TabIndex        =   295
         TabStop         =   0   'False
         Top             =   2415
         Width           =   4455
         Begin VB.OptionButton rbcBMissedDT 
            Caption         =   "Random"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3480
            TabIndex        =   297
            Top             =   0
            Width           =   1005
         End
         Begin VB.OptionButton rbcBMissedDT 
            Caption         =   "Avail times"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2265
            TabIndex        =   296
            Top             =   0
            Width           =   1275
         End
      End
      Begin VB.PictureBox plcBLaserForm 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   120
         ScaleHeight     =   420
         ScaleWidth      =   10575
         TabIndex        =   314
         TabStop         =   0   'False
         Top             =   3180
         Width           =   10575
         Begin VB.CheckBox ckcSortSS 
            Caption         =   "Sort by Sales Source"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   6780
            TabIndex        =   317
            Top             =   15
            Width           =   2175
         End
         Begin VB.OptionButton rbcBLaserForm 
            Caption         =   "(4) 3-Col Aired"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   5910
            TabIndex        =   590
            Top             =   225
            Width           =   1695
         End
         Begin VB.CheckBox ckcSuppressTimeForm1 
            Caption         =   "Suppress Air Time,"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4740
            TabIndex        =   316
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton rbcBLaserForm 
            Caption         =   "(3) Aired Column"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   4095
            TabIndex        =   319
            Top             =   225
            Width           =   1695
         End
         Begin VB.OptionButton rbcBLaserForm 
            Caption         =   "(2) Separate Invoice/Affidavit"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1425
            TabIndex        =   318
            Top             =   225
            Width           =   2700
         End
         Begin VB.OptionButton rbcBLaserForm 
            Caption         =   "(1) Ordered/Aired/Reconc Columns, "
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1425
            TabIndex        =   315
            Top             =   15
            Width           =   3240
         End
      End
      Begin VB.PictureBox plcBOrderDPShow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4845
         ScaleHeight     =   240
         ScaleWidth      =   4095
         TabIndex        =   298
         TabStop         =   0   'False
         Top             =   2415
         Width           =   4095
         Begin VB.OptionButton rbcBOrderDPShow 
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1905
            TabIndex        =   299
            Top             =   0
            Width           =   780
         End
         Begin VB.OptionButton rbcBOrderDPShow 
            Caption         =   "Time"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   2655
            TabIndex        =   300
            Top             =   0
            Width           =   720
         End
         Begin VB.OptionButton rbcBOrderDPShow 
            Caption         =   "None"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   3360
            TabIndex        =   301
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox plcSInvCntr 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   0
         ScaleHeight     =   420
         ScaleWidth      =   11145
         TabIndex        =   289
         TabStop         =   0   'False
         Top             =   1600
         Width           =   11145
         Begin VB.OptionButton rbcSInvCntr 
            Caption         =   "(D) as Aired, Update Ordered Veh"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   8160
            TabIndex        =   293
            Top             =   240
            Visible         =   0   'False
            Width           =   3675
         End
         Begin VB.OptionButton rbcSInvCntr 
            Caption         =   "(C) as Aired"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   6840
            TabIndex        =   292
            Top             =   240
            Width           =   1410
         End
         Begin VB.OptionButton rbcSInvCntr 
            Caption         =   "(A) as Ordered, Update Ordered Veh"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   290
            Top             =   240
            Width           =   3345
         End
         Begin VB.OptionButton rbcSInvCntr 
            Caption         =   "(B) as Ordered, Update Aired Veh"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3720
            TabIndex        =   291
            Top             =   240
            Width           =   3105
         End
      End
      Begin VB.TextBox edcBPayAddr 
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
         Height          =   300
         Index           =   0
         Left            =   7980
         MaxLength       =   25
         TabIndex        =   265
         Top             =   45
         Width           =   3210
      End
      Begin VB.TextBox edcBPayName 
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
         Height          =   300
         Left            =   1425
         MaxLength       =   25
         TabIndex        =   264
         Top             =   60
         Width           =   3210
      End
      Begin VB.TextBox edcBPayAddr 
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
         Height          =   300
         Index           =   1
         Left            =   7980
         MaxLength       =   25
         TabIndex        =   266
         Top             =   375
         Width           =   3210
      End
      Begin VB.TextBox edcBPayAddr 
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
         Height          =   300
         Index           =   2
         Left            =   7980
         MaxLength       =   25
         TabIndex        =   267
         Top             =   705
         Width           =   3210
      End
      Begin VB.TextBox edcBBillDate 
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
         Height          =   300
         Index           =   1
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   286
         Top             =   1275
         Width           =   870
      End
      Begin VB.TextBox edcBBillDate 
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
         Height          =   300
         Index           =   0
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   285
         Top             =   1275
         Width           =   870
      End
      Begin VB.PictureBox plcBCombine 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   3840
         TabIndex        =   309
         TabStop         =   0   'False
         Top             =   2925
         Width           =   3840
         Begin VB.OptionButton rbcBCombine 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3240
            TabIndex        =   311
            Top             =   0
            Width           =   540
         End
         Begin VB.OptionButton rbcBCombine 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2520
            TabIndex        =   310
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.PictureBox plcBNCycle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7440
         ScaleHeight     =   225
         ScaleWidth      =   2760
         TabIndex        =   279
         TabStop         =   0   'False
         Top             =   1070
         Width           =   2760
         Begin VB.OptionButton rbcBNCycle 
            Caption         =   "Wk"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2100
            TabIndex        =   282
            Top             =   0
            Width           =   555
         End
         Begin VB.OptionButton rbcBNCycle 
            Caption         =   "B'cast"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   280
            Top             =   0
            Width           =   810
         End
         Begin VB.OptionButton rbcBNCycle 
            Caption         =   "Cal"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1530
            TabIndex        =   281
            Top             =   0
            Width           =   600
         End
      End
      Begin VB.PictureBox plcBRCycle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4440
         ScaleHeight     =   225
         ScaleWidth      =   2835
         TabIndex        =   275
         TabStop         =   0   'False
         Top             =   1070
         Width           =   2835
         Begin VB.OptionButton rbcBRCycle 
            Caption         =   "Wk"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2130
            TabIndex        =   278
            Top             =   15
            Width           =   555
         End
         Begin VB.OptionButton rbcBRCycle 
            Caption         =   "Cal"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   277
            Top             =   0
            Width           =   585
         End
         Begin VB.OptionButton rbcBRCycle 
            Caption         =   "B'cast"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   765
            TabIndex        =   276
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.PictureBox plcBLCycle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   4290
         TabIndex        =   271
         TabStop         =   0   'False
         Top             =   1070
         Width           =   4290
         Begin VB.OptionButton rbcBLCycle 
            Caption         =   "Wk"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   3630
            TabIndex        =   274
            Top             =   0
            Width           =   570
         End
         Begin VB.OptionButton rbcBLCycle 
            Caption         =   "Cal"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3060
            TabIndex        =   273
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton rbcBLCycle 
            Caption         =   "B'cast"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2280
            TabIndex        =   272
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.TextBox edcBNo 
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
         Height          =   300
         Index           =   0
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   268
         Top             =   705
         Width           =   1065
      End
      Begin VB.TextBox edcBNo 
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
         Height          =   300
         Index           =   1
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   269
         Top             =   705
         Width           =   1065
      End
      Begin VB.TextBox edcBNo 
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
         Height          =   300
         Index           =   2
         Left            =   5760
         MaxLength       =   8
         TabIndex        =   270
         Top             =   705
         Width           =   1065
      End
      Begin VB.Label lacInv 
         Appearance      =   0  'Flat
         Caption         =   "Invoice E-Mail"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   355
         Top             =   5595
         Width           =   1365
      End
      Begin VB.Label lacInv 
         Appearance      =   0  'Flat
         Caption         =   "# Lines to Skip:  Above Logo                  Between Logo and Address"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   5
         Left            =   2790
         TabIndex        =   343
         Top             =   4755
         Width           =   6135
      End
      Begin VB.Label lacInv 
         Appearance      =   0  'Flat
         Caption         =   "Export ID        Media Type              Band"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   7575
         TabIndex        =   307
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label lacInv 
         Appearance      =   0  'Flat
         Caption         =   "Tax 1 Name                                                                         Tax 2 Name"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   326
         Top             =   3915
         Width           =   5325
      End
      Begin VB.Label lacInv 
         Appearance      =   0  'Flat
         Caption         =   "Lowest Inv # To Assign                           Highest #                          Next # "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   283
         Top             =   705
         Width           =   6660
      End
      Begin VB.Label lacInv 
         Appearance      =   0  'Flat
         Caption         =   $"Siteopt.frx":092E
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   263
         Top             =   120
         Width           =   7740
      End
      Begin VB.Label lacInv 
         Appearance      =   0  'Flat
         Caption         =   $"Siteopt.frx":09C3
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   284
         Top             =   1380
         Width           =   9240
      End
   End
   Begin VB.PictureBox plcComments 
      Height          =   5955
      Left            =   12960
      ScaleHeight     =   5895
      ScaleWidth      =   10995
      TabIndex        =   539
      Top             =   315
      Width           =   11055
      Begin VB.TextBox edcComment 
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
         Height          =   825
         Index           =   5
         Left            =   1575
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   599
         Top             =   4020
         Width           =   9000
      End
      Begin VB.TextBox edcComment 
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
         Height          =   825
         Index           =   4
         Left            =   1560
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   547
         Top             =   3045
         Width           =   9000
      End
      Begin VB.TextBox edcComment 
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
         Height          =   825
         Index           =   3
         Left            =   1545
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   549
         Top             =   4995
         Width           =   9000
      End
      Begin VB.TextBox edcComment 
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
         Height          =   825
         Index           =   1
         Left            =   1590
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   543
         Top             =   1095
         Width           =   9000
      End
      Begin VB.TextBox edcComment 
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
         Height          =   825
         Index           =   0
         Left            =   1605
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   541
         Top             =   120
         Width           =   9000
      End
      Begin VB.TextBox edcComment 
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
         Height          =   825
         Index           =   2
         Left            =   1575
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   545
         Top             =   2070
         Width           =   9000
      End
      Begin VB.Label lacComment 
         Appearance      =   0  'Flat
         Caption         =   "Station Posting Citation"
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   5
         Left            =   195
         TabIndex        =   598
         Top             =   3990
         Width           =   1170
      End
      Begin VB.Label lacComment 
         Appearance      =   0  'Flat
         Caption         =   "Statement of Account Comment"
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   4
         Left            =   240
         TabIndex        =   546
         Top             =   3015
         Width           =   1335
      End
      Begin VB.Label lacComment 
         Appearance      =   0  'Flat
         Caption         =   "Research Estimate Comment"
         ForeColor       =   &H80000008&
         Height          =   690
         Index           =   3
         Left            =   210
         TabIndex        =   548
         Top             =   4965
         Width           =   1155
      End
      Begin VB.Label lacComment 
         Appearance      =   0  'Flat
         Caption         =   "Insertion Comment"
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   542
         Top             =   1065
         Width           =   1125
      End
      Begin VB.Label lacComment 
         Appearance      =   0  'Flat
         Caption         =   "Contract Comment"
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   240
         TabIndex        =   540
         Top             =   90
         Width           =   1155
      End
      Begin VB.Label lacComment 
         Appearance      =   0  'Flat
         Caption         =   "Invoice Disclaimer"
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   2
         Left            =   240
         TabIndex        =   544
         Top             =   2040
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmcRCorpCal 
      Appearance      =   0  'Flat
      Caption         =   "D&efine"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8235
      TabIndex        =   363
      Top             =   6375
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmcCommand 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Height          =   285
      Index           =   4
      Left            =   7050
      TabIndex        =   537
      Top             =   6810
      Width           =   1050
   End
   Begin VB.CommandButton cmcCommand 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
      Height          =   285
      Index           =   3
      Left            =   5913
      TabIndex        =   536
      Top             =   6810
      Width           =   1050
   End
   Begin VB.CommandButton cmcCommand 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Index           =   2
      Left            =   4777
      TabIndex        =   535
      Top             =   6810
      Width           =   1050
   End
   Begin VB.CommandButton cmcCommand 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Index           =   1
      Left            =   3641
      TabIndex        =   534
      Top             =   6810
      Width           =   1050
   End
   Begin VB.CommandButton cmcCommand 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Index           =   0
      Left            =   2505
      TabIndex        =   533
      Top             =   6810
      Width           =   1050
   End
   Begin VB.PictureBox plcAgyAdv 
      Height          =   6105
      Left            =   240
      ScaleHeight     =   6045
      ScaleWidth      =   11460
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   7575
      Width           =   11520
      Begin VB.CheckBox ckcDaylightSavings 
         Caption         =   "Interface Honors Daylight Saving time"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   105
         TabIndex        =   147
         ToolTipText     =   "Honor Daylight Saving/Standard Time changes for interface purposes"
         Top             =   5790
         Width           =   3915
      End
      Begin VB.TextBox edcSageIE 
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
         Height          =   300
         Index           =   3
         Left            =   10110
         MaxLength       =   20
         TabIndex        =   127
         Top             =   2595
         Width           =   1200
      End
      Begin VB.TextBox edcSageIE 
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
         Height          =   300
         Index           =   2
         Left            =   8190
         MaxLength       =   20
         TabIndex        =   126
         Top             =   2595
         Width           =   1200
      End
      Begin VB.TextBox edcSageIE 
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
         Height          =   300
         Index           =   1
         Left            =   5925
         MaxLength       =   24
         TabIndex        =   125
         Top             =   2595
         Width           =   1245
      End
      Begin VB.TextBox edcSageIE 
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
         Height          =   300
         Index           =   0
         Left            =   2805
         MaxLength       =   40
         TabIndex        =   124
         Top             =   2595
         Width           =   2175
      End
      Begin VB.TextBox edcIE 
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
         Height          =   300
         Index           =   0
         Left            =   2805
         MaxLength       =   12
         TabIndex        =   606
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox edcIE 
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
         Height          =   300
         Index           =   1
         Left            =   4965
         MaxLength       =   6
         TabIndex        =   607
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox edcIE 
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
         Height          =   300
         Index           =   2
         Left            =   7365
         MaxLength       =   12
         TabIndex        =   608
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox edcExpt 
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
         Height          =   300
         Index           =   3
         Left            =   9090
         MaxLength       =   2
         TabIndex        =   107
         Top             =   690
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox ckcAEDI 
         Caption         =   "Using EDI Client and Product Codes"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   3360
         TabIndex        =   105
         Top             =   735
         Width           =   3270
      End
      Begin VB.CheckBox ckcXDSBy 
         Caption         =   "Unit ID by Affiliate Spot ID"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   5
         Left            =   2280
         TabIndex        =   600
         Top             =   4725
         Width           =   2490
      End
      Begin VB.CheckBox ckcACodes 
         Caption         =   "For Invoice Export: Suppress Zero-$ Spots"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   105
         TabIndex        =   146
         Top             =   5550
         Visible         =   0   'False
         Width           =   3915
      End
      Begin VB.TextBox edcExpt 
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
         Height          =   300
         Index           =   2
         Left            =   7380
         MaxLength       =   10
         TabIndex        =   144
         Top             =   4965
         Width           =   1470
      End
      Begin VB.TextBox edcExpt 
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
         Height          =   300
         Index           =   1
         Left            =   4155
         MaxLength       =   1
         TabIndex        =   142
         Top             =   4965
         Width           =   495
      End
      Begin VB.CheckBox ckcACodes 
         Caption         =   "Engineering/Cart Export- Suppress Media Code Except ""L"" (Live)"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   105
         TabIndex        =   145
         Top             =   5295
         Width           =   6075
      End
      Begin VB.CheckBox ckcXDSBy 
         Caption         =   "Unit ID by Affiliate Spot ID"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   8520
         TabIndex        =   139
         Top             =   4485
         Width           =   2490
      End
      Begin VB.CheckBox ckcXDSBy 
         Caption         =   "Midnight-Based Hours"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   6210
         TabIndex        =   138
         Top             =   4485
         Width           =   2475
      End
      Begin VB.CheckBox ckcXDSBy 
         Caption         =   "Include Advertiser abbreviation with ISCI"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   2280
         TabIndex        =   137
         Top             =   4485
         Width           =   3900
      End
      Begin VB.CheckBox ckcXDSBy 
         Caption         =   "by ISCI"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   1035
         TabIndex        =   140
         Top             =   4725
         Width           =   1155
      End
      Begin VB.CheckBox ckcXDSBy 
         Caption         =   "by Break"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   1035
         TabIndex        =   136
         Top             =   4485
         Width           =   1245
      End
      Begin VB.Frame frcGP 
         Caption         =   "Great Plains G/L"
         Height          =   1035
         Left            =   120
         TabIndex        =   128
         Top             =   3330
         Width           =   6795
         Begin VB.TextBox edcGP 
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
            Height          =   300
            Index           =   2
            Left            =   5160
            MaxLength       =   10
            TabIndex        =   134
            Top             =   600
            Width           =   1425
         End
         Begin VB.TextBox edcGP 
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
            Height          =   300
            Index           =   1
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   132
            Top             =   600
            Width           =   645
         End
         Begin VB.TextBox edcGP 
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
            Height          =   300
            Index           =   0
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   130
            Top             =   225
            Width           =   1425
         End
         Begin VB.Label lacAG 
            Appearance      =   0  'Flat
            Caption         =   "Customer # Prefix"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   131
            Top             =   645
            Width           =   1620
         End
         Begin VB.Label lacAG 
            Appearance      =   0  'Flat
            Caption         =   "Next Customer # to Assign"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   2745
            TabIndex        =   133
            Top             =   645
            Width           =   2220
         End
         Begin VB.Label lacAG 
            Appearance      =   0  'Flat
            Caption         =   "Batch #"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   129
            Top             =   270
            Width           =   1290
         End
      End
      Begin VB.TextBox edcTerms 
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
         Height          =   300
         Left            =   1410
         MaxLength       =   20
         TabIndex        =   123
         Top             =   2175
         Width           =   2310
      End
      Begin VB.PictureBox plcDefFillInv 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   4605
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   1905
         Width           =   4605
         Begin VB.OptionButton rbcDefFillInv 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3225
            TabIndex        =   121
            Top             =   0
            Width           =   540
         End
         Begin VB.OptionButton rbcDefFillInv 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2520
            TabIndex        =   120
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.TextBox edcEDI 
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
         Height          =   300
         Index           =   0
         Left            =   1905
         MaxLength       =   4
         TabIndex        =   109
         Top             =   1020
         Width           =   840
      End
      Begin VB.TextBox edcEDI 
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
         Height          =   300
         Index           =   1
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   110
         Top             =   1020
         Width           =   480
      End
      Begin VB.TextBox edcEDI 
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
         Height          =   300
         Index           =   2
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   111
         Top             =   1020
         Width           =   480
      End
      Begin VB.CheckBox ckcACodes 
         Caption         =   "Using Agency Codes"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   4170
         TabIndex        =   118
         Top             =   1620
         Width           =   2445
      End
      Begin VB.CheckBox ckcACodes 
         Caption         =   "Using Station Codes"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   1995
         TabIndex        =   117
         Top             =   1620
         Width           =   2595
      End
      Begin VB.CheckBox ckcACodes 
         Caption         =   "Using G/L Codes"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   116
         Top             =   1620
         Width           =   2655
      End
      Begin VB.CheckBox ckcAEDI 
         Caption         =   "Using EDI Service for Invoices"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   105
         TabIndex        =   104
         Top             =   735
         Width           =   2970
      End
      Begin VB.CheckBox ckcAEDI 
         Caption         =   "Using EDI Service for Contracts"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   103
         Top             =   480
         Width           =   3090
      End
      Begin VB.PictureBox plcAPrtStyle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   105
         ScaleHeight     =   255
         ScaleWidth      =   4980
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   1380
         Width           =   4980
         Begin VB.OptionButton rbcAPrtStyle 
            Caption         =   "Ask by Client"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2970
            TabIndex        =   115
            Top             =   0
            Width           =   1440
         End
         Begin VB.OptionButton rbcAPrtStyle 
            Caption         =   "Narrow"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2040
            TabIndex        =   114
            Top             =   0
            Width           =   900
         End
         Begin VB.OptionButton rbcAPrtStyle 
            Caption         =   "Wide"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1275
            TabIndex        =   113
            Top             =   0
            Width           =   780
         End
      End
      Begin VB.TextBox edcExpt 
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
         Height          =   300
         Index           =   0
         Left            =   1845
         MaxLength       =   1
         TabIndex        =   102
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "Dept ID"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   18
         Left            =   9465
         TabIndex        =   617
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "Location ID"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   17
         Left            =   7215
         TabIndex        =   616
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "Account #"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   16
         Left            =   5070
         TabIndex        =   615
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "Terms"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   15
         Left            =   1950
         TabIndex        =   614
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "SAGE Intacct Export:"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   14
         Left            =   120
         TabIndex        =   613
         Top             =   2640
         Width           =   1635
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "WO Invoice Export:"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   10
         Left            =   120
         TabIndex        =   612
         Top             =   2985
         Width           =   1530
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "Property"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   11
         Left            =   1950
         TabIndex        =   611
         Top             =   3000
         Width           =   1290
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "Prefix"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   12
         Left            =   4365
         TabIndex        =   610
         Top             =   3000
         Width           =   1290
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "Billing Group"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   13
         Left            =   6165
         TabIndex        =   609
         Top             =   3000
         Width           =   1290
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "Invoice Export Delimiter"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   7005
         TabIndex        =   106
         Top             =   735
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "Last Import Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   5655
         TabIndex        =   143
         Top             =   5010
         Width           =   1710
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "Head End Time Zone (E, C, M or P)"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   1005
         TabIndex        =   141
         Top             =   5010
         Width           =   2880
      End
      Begin VB.Label lacAG 
         Caption         =   "X-Digital:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   135
         Top             =   4485
         Width           =   1125
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "Default Terms"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   122
         Top             =   2205
         Width           =   1290
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "EDI\XML: Call Letters                       Media Type              Band"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   108
         Top             =   1065
         Width           =   5490
      End
      Begin VB.Label lacAG 
         Appearance      =   0  'Flat
         Caption         =   "# Target Demos"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   101
         Top             =   150
         Width           =   1710
      End
   End
   Begin VB.PictureBox plcLog 
      Height          =   1440
      Left            =   27600
      ScaleHeight     =   1380
      ScaleWidth      =   4290
      TabIndex        =   255
      Top             =   7800
      Visible         =   0   'False
      Width           =   4350
      Begin VB.CheckBox ckcAllowFinalLogDisplay 
         Caption         =   "Allow Display of Final Logs"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   261
         Top             =   1095
         Width           =   2820
      End
      Begin VB.CheckBox ckcAllowPrelLog 
         Caption         =   "Allow Preliminary Logs"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   260
         Top             =   735
         Width           =   2250
      End
      Begin VB.PictureBox plcDefLogCopy 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   4605
         TabIndex        =   257
         TabStop         =   0   'False
         Top             =   390
         Width           =   4605
         Begin VB.OptionButton rbcDefLogCopy 
            Caption         =   "On"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2265
            TabIndex        =   258
            Top             =   0
            Width           =   540
         End
         Begin VB.OptionButton rbcDefLogCopy 
            Caption         =   "Off"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   2865
            TabIndex        =   259
            Top             =   0
            Width           =   540
         End
      End
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Use Blackouts on Logs"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   256
         Top             =   45
         Width           =   2295
      End
   End
   Begin VB.PictureBox plcComm 
      Height          =   2265
      Left            =   26295
      ScaleHeight     =   2205
      ScaleWidth      =   10740
      TabIndex        =   575
      Top             =   630
      Visible         =   0   'False
      Width           =   10800
      Begin VB.Frame frcComm 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   585
         Top             =   735
         Width           =   5775
         Begin VB.OptionButton rbcCommBy 
            Caption         =   "Std"
            Height          =   210
            Index           =   6
            Left            =   3570
            TabIndex        =   587
            Top             =   0
            Width           =   630
         End
         Begin VB.OptionButton rbcCommBy 
            Caption         =   "Fiscal"
            Height          =   210
            Index           =   2
            Left            =   4320
            TabIndex        =   586
            Top             =   0
            Width           =   915
         End
         Begin VB.Label lacComm 
            Caption         =   "Commission on Std B'Cast Year or Fiscal Year"
            Height          =   210
            Index           =   2
            Left            =   0
            TabIndex        =   588
            Top             =   0
            Width           =   3450
         End
      End
      Begin VB.Frame frcComm 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   580
         Top             =   120
         Width           =   6870
         Begin VB.OptionButton rbcCommBy 
            Caption         =   "A/E"
            Height          =   210
            Index           =   5
            Left            =   1950
            TabIndex        =   583
            Top             =   0
            Width           =   630
         End
         Begin VB.OptionButton rbcCommBy 
            Caption         =   "Sub-Company"
            Height          =   210
            Index           =   4
            Left            =   2625
            TabIndex        =   582
            Top             =   0
            Width           =   1485
         End
         Begin VB.OptionButton rbcCommBy 
            Caption         =   "Contract"
            Height          =   210
            Index           =   3
            Left            =   4215
            TabIndex        =   581
            Top             =   15
            Width           =   1125
         End
         Begin VB.Label lacComm 
            Caption         =   "A/E Commissions by"
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   584
            Top             =   0
            Width           =   1875
         End
      End
      Begin VB.Frame frcComm 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   576
         Top             =   435
         Width           =   4830
         Begin VB.OptionButton rbcCommBy 
            Caption         =   "No"
            Height          =   210
            Index           =   1
            Left            =   3900
            TabIndex        =   579
            Top             =   0
            Width           =   585
         End
         Begin VB.OptionButton rbcCommBy 
            Caption         =   "Yes"
            Height          =   210
            Index           =   0
            Left            =   3195
            TabIndex        =   578
            Top             =   0
            Width           =   630
         End
         Begin VB.Label lacComm 
            Caption         =   "Allow Bonus % for New or Increases"
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   577
            Top             =   0
            Width           =   3135
         End
      End
   End
   Begin V81SiteOpt.SiteTabs udcSiteTabs 
      Height          =   255
      Left            =   11400
      TabIndex        =   622
      Top             =   -90
      Width           =   180
      _extentx        =   318
      _extenty        =   450
   End
   Begin VB.PictureBox plcBackup 
      Height          =   4500
      Left            =   26445
      ScaleHeight     =   4440
      ScaleWidth      =   8475
      TabIndex        =   412
      Top             =   2970
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CheckBox ckcCSIBackup 
         Caption         =   "Using CSI Backup"
         Height          =   255
         Left            =   720
         TabIndex        =   574
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox edcLastBkupLoc 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   330
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   420
         Top             =   3870
         Width           =   3105
      End
      Begin VB.TextBox edcLastBkupName 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   360
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   418
         Top             =   2880
         Width           =   7395
      End
      Begin VB.TextBox edcLastBkupSize 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   330
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   424
         Top             =   3870
         Width           =   1455
      End
      Begin VB.TextBox edcLastBkupDtTime 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   330
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   422
         Top             =   3870
         Width           =   2295
      End
      Begin VB.ComboBox cbcBkupTime 
         BackColor       =   &H00FFFF00&
         Height          =   330
         ItemData        =   "Siteopt.frx":0A55
         Left            =   660
         List            =   "Siteopt.frx":0AB3
         TabIndex        =   415
         Top             =   1350
         Width           =   2250
      End
      Begin VB.Frame frameDaysToRun 
         Caption         =   "Days to run"
         Height          =   1095
         Left            =   3360
         TabIndex        =   559
         Top             =   600
         Width           =   2775
         Begin VB.CheckBox chkDOW 
            Caption         =   "Check1"
            Height          =   255
            Index           =   6
            Left            =   2280
            TabIndex        =   573
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox chkDOW 
            Caption         =   "Check1"
            Height          =   255
            Index           =   5
            Left            =   1920
            TabIndex        =   572
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox chkDOW 
            Caption         =   "Check1"
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   571
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox chkDOW 
            Caption         =   "Check1"
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   570
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox chkDOW 
            Caption         =   "Check1"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   569
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox chkDOW 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   568
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox chkDOW 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   567
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Su"
            Height          =   255
            Index           =   6
            Left            =   2280
            TabIndex        =   566
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Sa"
            Height          =   255
            Index           =   5
            Left            =   1920
            TabIndex        =   565
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Fr"
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   564
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Th"
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   563
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "We"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   562
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Tu"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   561
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Mo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   560
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Label lacBackup 
         Caption         =   "File Location"
         Height          =   330
         Index           =   4
         Left            =   660
         TabIndex        =   419
         Top             =   3510
         Width           =   1770
      End
      Begin VB.Label lacBackup 
         Caption         =   "File Name"
         Height          =   330
         Index           =   3
         Left            =   660
         TabIndex        =   417
         Top             =   2535
         Width           =   3030
      End
      Begin VB.Label lacBackup 
         Caption         =   "Last Backup Detail Information:"
         Height          =   315
         Index           =   2
         Left            =   600
         TabIndex        =   416
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label lacBackup 
         Caption         =   "Backup Options:"
         Height          =   360
         Index           =   7
         Left            =   360
         TabIndex        =   413
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label lacBackup 
         Caption         =   "File Size"
         Height          =   330
         Index           =   6
         Left            =   6540
         TabIndex        =   423
         Top             =   3510
         Width           =   1770
      End
      Begin VB.Label lacBackup 
         Caption         =   "Date and Time"
         Height          =   330
         Index           =   5
         Left            =   4005
         TabIndex        =   421
         Top             =   3510
         Width           =   2805
      End
      Begin VB.Label lacBackup 
         Caption         =   "Choose Daily Backup Time"
         Height          =   270
         Index           =   0
         Left            =   660
         TabIndex        =   414
         Top             =   1080
         Width           =   2895
      End
   End
   Begin VB.PictureBox plcCntr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      Left            =   390
      ScaleHeight     =   5850
      ScaleWidth      =   11235
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   615
      Width           =   11295
      Begin VB.CheckBox ckcDisallowAuthorScheduling 
         Caption         =   "Disallow user who edited a contract from scheduling it"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   199
         Top             =   5265
         Width           =   4845
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Calculate Research Totals on Line Change"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   180
         Top             =   3375
         Width           =   4170
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Print Signature Line on Proposals"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   7800
         TabIndex        =   198
         Top             =   4950
         Width           =   3420
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Show Day Dropdown if Flight Button Allowed"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   168
         Top             =   2400
         Width           =   4365
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Include Audience Percentages for Podcast"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   163
         Top             =   1920
         Width           =   4305
      End
      Begin VB.CheckBox ckcCAvails 
         Caption         =   "Per Inquiry"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   9885
         TabIndex        =   597
         Top             =   1425
         Width           =   1275
      End
      Begin VB.CheckBox ckcCAvails 
         Caption         =   "Direct Response"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   8145
         TabIndex        =   596
         Top             =   1425
         Width           =   1740
      End
      Begin VB.CheckBox ckcCAvails 
         Caption         =   "Remnant"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   7005
         TabIndex        =   595
         Top             =   1425
         Width           =   1095
      End
      Begin VB.CheckBox ckcCAvails 
         Caption         =   "Reservation"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   5625
         TabIndex        =   594
         Top             =   1425
         Width           =   1425
      End
      Begin VB.CheckBox ckcSales 
         Caption         =   "Print less than 13 Weeks by Standard Quarters"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   3360
         TabIndex        =   197
         Top             =   4950
         Width           =   4410
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Show Station Market Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   510
         TabIndex        =   196
         Top             =   4950
         Width           =   2775
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Cancel Clause Mandatory"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   4755
         TabIndex        =   184
         Top             =   3630
         Width           =   2955
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "'Freeze Calculation' As Default on Toggle"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   4755
         TabIndex        =   186
         Top             =   3885
         Width           =   3840
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Package Lines- Hide Hidden Lines"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   185
         Top             =   3885
         Width           =   3570
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Show Audio Type"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   3825
         TabIndex        =   194
         Top             =   4680
         Width           =   1965
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Contract Verification"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   7305
         TabIndex        =   182
         Top             =   3375
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Change spots in deleted weeks behind last log date to Fills"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   4515
         TabIndex        =   154
         Top             =   465
         Width           =   5460
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Package Lines- Show Flight Rates"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   5865
         TabIndex        =   195
         Top             =   4680
         Width           =   3480
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Product Protection Mandatory"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   183
         Top             =   3630
         Width           =   2955
      End
      Begin VB.CheckBox ckcCntr 
         Caption         =   "Show Comments on Detail Page"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   193
         Top             =   4680
         Width           =   3510
      End
      Begin VB.PictureBox plcCSchd 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   510
         ScaleHeight     =   255
         ScaleWidth      =   5085
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   4380
         Width           =   5085
         Begin VB.OptionButton rbcCSortBy 
            Caption         =   "Line #"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   3660
            TabIndex        =   192
            Top             =   15
            Width           =   885
         End
         Begin VB.OptionButton rbcCSortBy 
            Caption         =   "Rate Card"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1455
            TabIndex        =   190
            Top             =   15
            Width           =   1185
         End
         Begin VB.OptionButton rbcCSortBy 
            Caption         =   "Daypart"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2670
            TabIndex        =   191
            Top             =   15
            Width           =   975
         End
         Begin VB.Label lacCntr 
            Caption         =   "Sort Lines by-"
            Height          =   210
            Index           =   5
            Left            =   0
            TabIndex        =   189
            Top             =   15
            Width           =   2865
         End
      End
      Begin VB.CheckBox ckcBookInto 
         Caption         =   "Book PSA into Contract Avail"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   4995
         TabIndex        =   179
         Top             =   3105
         Width           =   3105
      End
      Begin VB.CheckBox ckcBookInto 
         Caption         =   "Book Promo into Contract Avail"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   4995
         TabIndex        =   175
         Top             =   2880
         Width           =   3105
      End
      Begin VB.PictureBox pbcImptCntr 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   8340
         TabIndex        =   200
         TabStop         =   0   'False
         Top             =   5640
         Visible         =   0   'False
         Width           =   8340
         Begin VB.OptionButton rbcImptCntr 
            Caption         =   "Prohibit"
            Height          =   210
            Index           =   2
            Left            =   7005
            TabIndex        =   204
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton rbcImptCntr 
            Caption         =   "Without Move Separation Rules"
            Height          =   210
            Index           =   1
            Left            =   3990
            TabIndex        =   203
            Top             =   0
            Width           =   2865
         End
         Begin VB.OptionButton rbcImptCntr 
            Caption         =   "With Move Separation Rules"
            Height          =   210
            Index           =   0
            Left            =   1320
            TabIndex        =   202
            Top             =   0
            Width           =   2625
         End
         Begin VB.Label lacCntr 
            Caption         =   "Import Orders"
            Height          =   210
            Index           =   4
            Left            =   0
            TabIndex        =   201
            Top             =   0
            Width           =   1260
         End
      End
      Begin VB.PictureBox plcCSchd 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   4785
         TabIndex        =   176
         TabStop         =   0   'False
         Top             =   3090
         Width           =   4785
         Begin VB.OptionButton rbcCSchdPSA 
            Caption         =   "Manual"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2625
            TabIndex        =   177
            Top             =   0
            Width           =   930
         End
         Begin VB.OptionButton rbcCSchdPSA 
            Caption         =   "Automatic"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3555
            TabIndex        =   178
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.PictureBox plcCSchd 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   4815
         TabIndex        =   172
         TabStop         =   0   'False
         Top             =   2865
         Width           =   4815
         Begin VB.OptionButton rbcCSchdPromo 
            Caption         =   "Manual"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2625
            TabIndex        =   173
            Top             =   0
            Width           =   930
         End
         Begin VB.OptionButton rbcCSchdPromo 
            Caption         =   "Automatic"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3555
            TabIndex        =   174
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.CheckBox ckcCUseSegments 
         Caption         =   "Using Segments"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4755
         TabIndex        =   181
         Top             =   3375
         Width           =   2400
      End
      Begin VB.PictureBox plcVirtPkg 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   9750
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   2130
         Width           =   9750
         Begin VB.OptionButton rbcVirtPkg 
            Caption         =   "Alter Virtual Spot Retain Total $; or Real N/A"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   5265
            TabIndex        =   167
            Top             =   60
            Width           =   4110
         End
         Begin VB.OptionButton rbcVirtPkg 
            Caption         =   "Alter Virtual/Real Spot/$ Adjust Hidden"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1485
            TabIndex        =   166
            Top             =   45
            Width           =   3705
         End
      End
      Begin VB.PictureBox plcCSchd 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   5790
         TabIndex        =   169
         TabStop         =   0   'False
         Top             =   2640
         Width           =   5790
         Begin VB.OptionButton rbcCSchdRemnant 
            Caption         =   "Automatic"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3555
            TabIndex        =   171
            Top             =   0
            Width           =   1140
         End
         Begin VB.OptionButton rbcCSchdRemnant 
            Caption         =   "Manual"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2625
            TabIndex        =   170
            Top             =   0
            Width           =   930
         End
      End
      Begin VB.CheckBox ckcCAudPkg 
         Caption         =   "Allow Audience Packages"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6525
         TabIndex        =   164
         Top             =   1920
         Width           =   2835
      End
      Begin VB.CheckBox ckcCWarnMsg 
         Caption         =   "Show Contract Warning Messages"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   162
         Top             =   1680
         Width           =   3375
      End
      Begin VB.CheckBox ckcCAvails 
         Caption         =   "Missed"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   4635
         TabIndex        =   161
         Top             =   1425
         Width           =   1095
      End
      Begin VB.CheckBox ckcCBump 
         Caption         =   "For New Schedule Lines, Bump Spots in Past"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   153
         Top             =   465
         Width           =   4290
      End
      Begin VB.ComboBox cbcReallDemo 
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
         Height          =   285
         Left            =   2595
         TabIndex        =   156
         Top             =   705
         Width           =   1260
      End
      Begin VB.PictureBox plcReallDate 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Height          =   300
         Left            =   4530
         ScaleHeight     =   270
         ScaleWidth      =   1065
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   705
         Width           =   1095
      End
      Begin VB.PictureBox plcDiscDate 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Height          =   300
         Left            =   6390
         ScaleHeight     =   270
         ScaleWidth      =   990
         TabIndex        =   160
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1020
      End
      Begin VB.TextBox edcDiscNo 
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
         Height          =   300
         Left            =   3585
         MaxLength       =   8
         TabIndex        =   159
         Top             =   1050
         Width           =   1065
      End
      Begin VB.TextBox edcCNo 
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
         Height          =   300
         Index           =   2
         Left            =   8175
         MaxLength       =   8
         TabIndex        =   152
         Top             =   75
         Width           =   1065
      End
      Begin VB.TextBox edcCNo 
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
         Height          =   300
         Index           =   1
         Left            =   4590
         MaxLength       =   8
         TabIndex        =   151
         Top             =   75
         Width           =   1065
      End
      Begin VB.TextBox edcCNo 
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
         Height          =   300
         Index           =   0
         Left            =   3315
         MaxLength       =   8
         TabIndex        =   150
         Top             =   75
         Width           =   1065
      End
      Begin VB.Label lacCntr 
         Appearance      =   0  'Flat
         Caption         =   "Contract Avail Counts && 6 Month Avail Report, Include"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   593
         Top             =   1410
         Width           =   4380
      End
      Begin VB.Label lacCntr 
         Appearance      =   0  'Flat
         Caption         =   "On Proposals/Orders:"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   187
         Top             =   4155
         Width           =   2400
      End
      Begin VB.Label lacCntr 
         Appearance      =   0  'Flat
         Caption         =   "Audience Reallocation: Demo                                   Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   155
         Top             =   735
         Width           =   4425
      End
      Begin VB.Label lacCntr 
         Appearance      =   0  'Flat
         Caption         =   "Contract Discrepancy: Current Contract #                                     Last Run Date"
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   3
         Left            =   120
         TabIndex        =   158
         Top             =   1095
         Width           =   6180
      End
      Begin VB.Label lacCntr 
         Appearance      =   0  'Flat
         Caption         =   $"Siteopt.frx":0B42
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   149
         Top             =   120
         Width           =   8205
      End
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
      Left            =   -30
      ScaleHeight     =   165
      ScaleWidth      =   105
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   6090
      Width           =   105
   End
   Begin VB.PictureBox plcSchedule 
      Height          =   4410
      Left            =   26790
      ScaleHeight     =   4350
      ScaleWidth      =   8745
      TabIndex        =   516
      Top             =   15540
      Visible         =   0   'False
      Width           =   8805
      Begin VB.CheckBox ckcRegionMixLen 
         Caption         =   "Allow Mixed Split Network Spot Lengths to Combine"
         Height          =   210
         Left            =   120
         TabIndex        =   518
         Top             =   345
         Width           =   5070
      End
      Begin VB.Frame frcPollUsers 
         Caption         =   "User for Feed Import"
         Height          =   1185
         Left            =   120
         TabIndex        =   528
         Top             =   2700
         Width           =   4065
         Begin VB.ComboBox cbcUser 
            BackColor       =   &H00FFFF00&
            Height          =   330
            Index           =   1
            ItemData        =   "Siteopt.frx":0BCA
            Left            =   1125
            List            =   "Siteopt.frx":0BCC
            Sorted          =   -1  'True
            TabIndex        =   532
            Top             =   735
            Width           =   2700
         End
         Begin VB.ComboBox cbcUser 
            BackColor       =   &H00FFFF00&
            Height          =   330
            Index           =   0
            ItemData        =   "Siteopt.frx":0BCE
            Left            =   1125
            List            =   "Siteopt.frx":0BD0
            Sorted          =   -1  'True
            TabIndex        =   530
            Top             =   255
            Width           =   2700
         End
         Begin VB.Label lacSchedule 
            Caption         =   "Secondary"
            Height          =   210
            Index           =   4
            Left            =   180
            TabIndex        =   531
            Top             =   780
            Width           =   900
         End
         Begin VB.Label lacSchedule 
            Caption         =   "Primary"
            Height          =   210
            Index           =   3
            Left            =   180
            TabIndex        =   529
            Top             =   300
            Width           =   825
         End
      End
      Begin VB.CheckBox ckcCmmlSchStatus 
         Caption         =   "Ask Disposition of Commercial Changes when Generating Logs"
         Height          =   210
         Left            =   120
         TabIndex        =   517
         Top             =   45
         Width           =   5760
      End
      Begin VB.TextBox edcLevelPrice 
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
         Left            =   615
         MaxLength       =   5
         TabIndex        =   526
         TabStop         =   0   'False
         Top             =   1755
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.PictureBox pbcSchedule 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   105
         Picture         =   "Siteopt.frx":0BD2
         ScaleHeight     =   630
         ScaleWidth      =   8520
         TabIndex        =   525
         Top             =   1605
         Width           =   8520
      End
      Begin VB.TextBox edcSchedule 
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
         Height          =   300
         Index           =   0
         Left            =   3900
         MaxLength       =   5
         TabIndex        =   520
         Top             =   690
         Width           =   1080
      End
      Begin VB.TextBox edcSchedule 
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
         Height          =   300
         Index           =   1
         Left            =   3900
         MaxLength       =   5
         TabIndex        =   522
         Top             =   1095
         Width           =   1080
      End
      Begin VB.PictureBox pbcLevelSTab 
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
         Left            =   -195
         ScaleHeight     =   60
         ScaleWidth      =   120
         TabIndex        =   524
         Top             =   1125
         Width           =   120
      End
      Begin VB.PictureBox pbcLevelTab 
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
         ScaleWidth      =   120
         TabIndex        =   527
         Top             =   2160
         Width           =   120
      End
      Begin VB.CommandButton cmcLevelPrice 
         Appearance      =   0  'Flat
         Caption         =   "Generate Price Levels"
         Height          =   255
         Left            =   5385
         TabIndex        =   523
         Top             =   915
         Width           =   2175
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "Lowest Spot Price to Generate Price Levels"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   519
         Top             =   705
         Width           =   3705
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "Highest Spot Price to Generate Price Levels"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   105
         TabIndex        =   521
         Top             =   1110
         Width           =   3765
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "Level 0: Fill Spot; Level 1: $0 and N/C Spot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   552
         Top             =   2280
         Width           =   5790
      End
   End
   Begin ComctlLib.TabStrip tbcSelection 
      Height          =   6480
      Left            =   -15
      TabIndex        =   1
      Top             =   240
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   11430
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   18
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Gen"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Sale"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Comm"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "A&ud"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sch&d"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Expt"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Cntr"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cop&y"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Log"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Inv"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Acc&t"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Bkup"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "O&pt"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Spo&rt"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Aut&o"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "C&mnt"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab17 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Lin&k"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab18 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "E-M"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lacChanges 
      Caption         =   "No more changes may be made"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   75
      TabIndex        =   589
      Top             =   6810
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Label lacDBLocation 
      Alignment       =   1  'Right Justify
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
      Height          =   195
      Left            =   1770
      TabIndex        =   538
      Top             =   0
      Width           =   7515
   End
   Begin VB.Label plcScreen 
      Caption         =   "Site Options"
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   1080
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1830
      Top             =   6750
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "SiteOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SiteOpt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Site Option input screen code
Option Explicit
Option Compare Text
Private smOrigClientName As String
Dim imFirstActivate As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim hmCxf As Integer    'Comment file handle
Dim tmCxf As CXF        'CXF record image
Dim tmCxfSrchKey As LONGKEY0    'CXF key record image
Dim imCxfRecLen As Integer        'CXF record length
Dim tmMnf As MNF        'Mnf record image
Dim tmMnfSrchKey As INTKEY0    'Mnf key record image
Dim hmMnf As Integer    'Multi-Name file handle
Dim imMnfRecLen As Integer        'MNF record length
Dim hmSaf As Integer
Dim tmSaf As SAF            'Schedule Attributes record image
Dim tmSafSrchKey1 As SAFKEY1    'Vef key record image
Dim imSafRecLen As Integer
Dim hmSite As Integer
Dim tmSiteSrchKey As LONGKEY0    'Vef key record image
Dim imSiteRecLen As Integer
Dim hmNrf As Integer 'Avail name file handle
Dim tmNrfSrchKey As INTKEY0    'ANF key record image
Dim imNrfRecLen As Integer        'ANF record length
Dim hmSdf As Integer
'Dim hmAtt As Integer
'Dim hmAst As Integer
Dim smStdTerms As String
Dim smSyncDate As String
Dim smSyncTime As String
'Delete
'Dim tmDsf As DSF            'DSF record image
'Dim tmDsfSrchKey As LONGKEY0  'DSF key record image
'Dim hmDsf As Integer        'DSF Handle
'Dim imDsfRecLen As Integer      'DSF record length
Dim imTerminate As Integer
Dim smReallDateCaption As String
Dim smDiscDateCaption As String
Dim smRLastPayCaption As String
Dim imPWStatus As Integer   '0=Not Asked; 1=Ok; -1=Not Ok
Dim imUpdateAllowed As Integer    'User can update records
Dim imInvCommentLen As Integer  '2-12-03
Dim smInvComment As String      '2-12-03

Dim imContrCommentLen As Integer  '2-12-03
Dim smContrComment As String      '2-12-03

Dim imInsertCommentLen As Integer  '2-12-03
Dim smInsertComment As String      '2-12-03

Dim imEstCommentLen As Integer  '2-12-03
Dim smEstComment As String      '2-12-03

Dim imStatementCommentLen As Integer
Dim smStatementComment As String

Dim imCitationCommentLen As Integer
Dim smCitationComment As String

Dim lmBCxfDisclaimer As Long        '2-20-03
Dim lmCxfContrComment As Long       '2-20-03
Dim lmCxfInsertComment As Long      '2-20-03
Dim lmCxfEstComment As Long       '2-20-03
Dim lmCxfStatementComment As Long      '2-20-03
Dim lmCxfCitationComment As Long      '2-20-03

Dim imIgnoreClickEvent As Integer
Private imUpdateError As Integer

'Schedule Price
Dim tmSCtrls(0 To 14) As FIELDAREA
Dim imLBSCtrls As Integer
Dim lmSSave(0 To 14) As Long    'Index zero ignored
Dim imSBoxNo As Integer
Dim imSRowNo As Integer

Dim imPWPrefSpotPct As Integer
Dim imWk1stSolo As Integer
Dim smCSIServerINIFile As String
Dim smBUWeekDays As String
' Dan M limit #of changes 4/21/09
Private imChangesOccured As Integer
Private bmExceededChanges As Boolean
Public bmIgnoreChange As Boolean
'7942
Private smHeadEndZoneChange As String
Const UNLIMITED = 999

Const LBONE = 1

Const LEVEL2INDEX = 1
Const LEVEL3INDEX = 2
Const LEVEL4INDEX = 3
Const LEVEL5INDEX = 4
Const LEVEL6INDEX = 5
Const LEVEL7INDEX = 6
Const LEVEL8INDEX = 7
Const LEVEL9INDEX = 8
Const LEVEL10INDEX = 9
Const LEVEL11INDEX = 10
Const LEVEL12INDEX = 11
Const LEVEL13INDEX = 12
Const LEVEL14INDEX = 13
Const LEVEL15INDEX = 14
'9114
Const ASTBREAK = 4
Const ASTISCI = 5
Const BYISCI = 0
Const BYBREAK = 1
'10016
Const AIRANDNTR = 0
Const AIRONLY = 1
Const NTRONLY = 2
Const INVAUTO = 10
Const INVSELECTIVE = 12 'ckcInv(12) - TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
Const INVCOMBINE = 1
'10048
Const PODAIRTIMECKC = 28
Const PODSPOTSCKC = 29
Const ADSERVERCKC = 30
Const PODMIXCKC = 12

Const COMMAND_DONE = 0
Const COMMAND_CANCEL = 1
Const COMMAND_SAVE = 2
Const COMMAND_UNDO = 3
Const COMMAND_REPORT = 4

'TTP 10626 JJB 2023-01-10
Const SAGE_IE_TERMS = 0
Const SAGE_IE_ACCOUNT = 1
Const SAGE_IE_LOCATIONID = 2
Const SAGE_IE_DEPTID = 3
''''''''''''''''''''''''

Private Sub cbcBkupTime_Click()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub cbcBkupTime_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub


Private Sub cbcReallDemo_Click()
If Not bmIgnoreChange Then
      mChangeOccured
End If

End Sub

Private Sub cbcReallDemo_gotFocus()
mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub cbcReconGroup_Click()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub cbcReconGroup_gotFocus()
mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub cbcUser_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub chkDOW_Click(Index As Integer)
    If Not bmIgnoreChange Then
        mChangeOccured
    End If

End Sub

Private Sub chkDOW_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcAffiliateCRM_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcAffiliateCRM.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "AC-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcAffiliateCRM.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcAffiliateCRM.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcAffiliateCRM_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcAllowFinalLogDisplay_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcAllowFinalLogDisplay_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub


Private Sub ckcCntr_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcCntr_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcCompensation_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcCompensation.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "CP-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcCompensation.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcCompensation.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcCompensation_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcCopy_Click(Index As Integer)
    If Index = 1 Then
        'If (ckcUsingLiveCopy.Value = vbUnchecked) And (ckcCopy(1).Value = vbUnchecked) Then
        '    ckcOverrideOptions(5).Value = vbUnchecked
        '    ckcOverrideOptions(5).Enabled = False
        '    ckcCopy(5).Enabled = True
        '    ckcCopy(6).Enabled = False
        '    ckcCopy(6).Value = vbChecked
        '    ckcCopy(7).Enabled = False
        '    ckcCopy(7).Value = vbChecked
        '    ckcCopy(8).Enabled = False
        '    ckcCopy(8).Value = vbChecked
        '    ckcCopy(9).Enabled = False
        '    ckcCopy(9).Value = vbChecked
        '    ckcCopy(10).Enabled = False
        '    ckcCopy(10).Value = vbChecked
        'ElseIf (ckcUsingLiveCopy.Value = vbUnchecked) Then
        '    ckcCopy(6).Enabled = False
        '    ckcCopy(6).Value = vbChecked
        '    ckcCopy(7).Enabled = False
        '    ckcCopy(7).Value = vbChecked
        '    ckcCopy(9).Enabled = False
        '    ckcCopy(9).Value = vbChecked
        '    ckcCopy(10).Enabled = False
        '    ckcCopy(10).Value = vbChecked
        'ElseIf (ckcCopy(1).Value = vbUnchecked) Then
        '    ckcCopy(7).Enabled = False
        '    ckcCopy(7).Value = vbChecked
        '    ckcCopy(8).Enabled = False
        '    ckcCopy(8).Value = vbChecked
        '    ckcCopy(10).Enabled = False
        '    ckcCopy(10).Value = vbChecked
        'Else
        '    ckcOverrideOptions(5).Enabled = True
        '    ckcCopy(5).Enabled = True
        '    ckcCopy(6).Enabled = True
        '    ckcCopy(7).Enabled = True
        '    ckcCopy(8).Enabled = True
        '    ckcCopy(9).Enabled = True
        '    ckcCopy(10).Enabled = True
        'End If
        mSetAudioTypes
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcCopy_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcCSIBackup_Click()
    Dim ilLoop As Integer

    If ckcCSIBackup.Value = 1 Then
        edcLastBkupName.Enabled = True
        edcLastBkupLoc.Enabled = True
        edcLastBkupDtTime.Enabled = True
        edcLastBkupSize.Enabled = True
        cbcBkupTime.Enabled = True
        For ilLoop = 0 To 6
            chkDOW(ilLoop).Enabled = True
        Next
    Else
        edcLastBkupName.Enabled = False
        edcLastBkupLoc.Enabled = False
        edcLastBkupDtTime.Enabled = False
        edcLastBkupSize.Enabled = False
        cbcBkupTime.Enabled = False
        For ilLoop = 0 To 6
            chkDOW(ilLoop).Enabled = False
        Next
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcACodes_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcDaylightSavings_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcAEDI_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub


Private Sub ckcAllowPrelLog_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub



Private Sub ckcBookInto_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcCAudPkg_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub ckcCBump_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
'10843 removed ckc
'Private Sub ckcCLnStdQt_GotFocus()
'    mCtrlGotFocusAndIgnoreChange ActiveControl
'End Sub
Private Sub ckcCAvails_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcCmmlSchStatus_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub ckcCUseSegments_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub ckcCWarnMsg_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcDisallowAuthorScheduling_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcDigital_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcDigital.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "DC-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcDigital.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcDigital.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcDigital_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub




Private Sub ckcEfficio_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcEfficio.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "EE-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcEfficio.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcEfficio.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcEfficio_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcEventRevenue_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcEventRevenue.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "ER-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcEventRevenue.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcEventRevenue.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
        mChangeOccured
    End If
End Sub

Private Sub ckcEventRevenue_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcGetPaidExport_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcGetPaidExport.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "GP-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcGetPaidExport.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcGetPaidExport.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcGetPaidExport_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcGreatPlainGL_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcGreatPlainGL.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "GP-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcGreatPlainGL.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcGreatPlainGL.Value = vbUnchecked
        End If
    End If
    If ckcGreatPlainGL.Value = vbChecked Then
        frcGP.Enabled = True
        ckcACodes(0).Value = vbChecked
    Else
        frcGP.Enabled = False
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcGreatPlainGL_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcGUseAffFeed_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If (ckcUsingBBs.Value = vbChecked) And (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) <= 0) And (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) <= 0) Then
                'No support code placed into Station Feed; Exports : NY, Dallas and Phonxie; CnC
                MsgBox "Station Feed not allowed in conjunction with Billboards at this time."
                ckcGUseAffFeed.Value = vbUnchecked
            Else
                If ckcGUseAffFeed.Value = vbChecked Then
                    ilPasswordOk = igPasswordOk
                    sgPasswordAddition = "USF-"
                    mProtectChangesAllowed Start
                    CSPWord.Show vbModal
                    mProtectChangesAllowed Done
                    If Not igPasswordOk Then
                        ckcGUseAffFeed.Value = vbUnchecked
                    End If
                    igPasswordOk = ilPasswordOk
                    sgPasswordAddition = ""
                End If
            End If
        Else
            ckcGUseAffFeed.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcGUseAffFeed_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
'    If imPWStatus = 0 Then
'        imPWStatus = -2
'        If (Trim$(tgUrf(0).sName) <> sgCPName) Then
'            ilPasswordOk = igPasswordOk
'            CSPWord.Show vbModal
'            DoEvents
'            If Not igPasswordOk Then
'                imPWStatus = -1
'                igPasswordOk = ilPasswordOk
'                If tgSpf.sGUseAffFeed = "Y" Then 'Using Proposal System
'                    ckcGUseAffFeed.Value = vbChecked
'                Else
'                    ckcGUseAffFeed.Value = vbUnchecked
'                End If
'                cmcCancel.SetFocus
'                Exit Sub
'            Else
'                imPWStatus = 1
'                igPasswordOk = ilPasswordOk
'            End If
'        Else
'            imPWStatus = 1
'        End If
'    Else
'        If imPWStatus = -1 Then
'            If tgSpf.sGUseAffFeed = "Y" Then 'Using Proposal System
'                ckcGUseAffFeed.Value = vbChecked
'            Else
'                ckcGUseAffFeed.Value = vbUnchecked
'            End If
'            cmcCancel.SetFocus
'        End If
'    End If
End Sub
Private Sub ckcGUseAffFeed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If imPWStatus = 0 Then
'        If (Trim$(tgUrf(0).sName) <> sgCPName) Then
'            ilPasswordOk = igPasswordOk
'            CSPWord.Show vbModal
'            DoEvents
'            If Not igPasswordOk Then
'                imPWStatus = -1
'                igPasswordOk = ilPasswordOk
'                If tgSpf.sGUseAffFeed = "Y" Then 'Using Proposal System
'                    ckcGUseAffFeed.Value = vbChecked
'                Else
'                    ckcGUseAffFeed.Value = vbUnchecked
'                End If
'                cmcCancel.SetFocus
'                Exit Sub
'            Else
'                imPWStatus = 1
'                igPasswordOk = ilPasswordOk
'            End If
'        Else
'            imPWStatus = 1
'        End If
'    Else
'        If imPWStatus = -1 Then
'            If tgSpf.sGUseAffFeed = "Y" Then 'Using Proposal System
'                ckcGUseAffFeed.Value = vbChecked
'            Else
'                ckcGUseAffFeed.Value = vbUnchecked
'            End If
'            cmcCancel.SetFocus
'        End If
'    End If
End Sub

Private Sub ckcGUseAffSys_Click(Index As Integer)
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If Index = 0 Then
                If (ckcSales(10).Value = vbChecked) And (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) <= 0) And (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) <= 0) Then
                    'No support code placed into Station Feed; Exports : NY, Dallas and Phonxie; CnC
                    MsgBox "Affiliate System not allowed in conjunction with Billboards at this time."
                    ckcGUseAffSys(0).Value = vbUnchecked
                Else
                    If ckcGUseAffSys(0).Value = vbChecked Then
                        ilPasswordOk = igPasswordOk
                        sgPasswordAddition = "UAS-"
                        mProtectChangesAllowed Start
                        CSPWord.Show vbModal
                        mProtectChangesAllowed Done
                        If Not igPasswordOk Then
                            ckcGUseAffSys(0).Value = vbUnchecked
                        End If
                        igPasswordOk = ilPasswordOk
                        sgPasswordAddition = ""
                    End If
                End If
            ElseIf Index = 1 Then
                If ckcGUseAffSys(1).Value = vbChecked Then
                    ilPasswordOk = igPasswordOk
                    sgPasswordAddition = "USI-"
                    mProtectChangesAllowed Start
                    CSPWord.Show vbModal
                    mProtectChangesAllowed Done
                    If Not igPasswordOk Then
                        ckcGUseAffSys(1).Value = vbUnchecked
                    End If
                    igPasswordOk = ilPasswordOk
                    sgPasswordAddition = ""
                End If
            ElseIf Index = 2 Then
                If ckcGUseAffSys(2).Value = vbChecked Then
                    ilPasswordOk = igPasswordOk
                    sgPasswordAddition = "UR-"
                    mProtectChangesAllowed Start
                    CSPWord.Show vbModal
                    mProtectChangesAllowed Done
                    If Not igPasswordOk Then
                        ckcGUseAffSys(2).Value = vbUnchecked
                    End If
                    igPasswordOk = ilPasswordOk
                    sgPasswordAddition = ""
                End If
            End If
        Else
            ckcGUseAffSys(Index).Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcGUseAffSys_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub


Private Sub ckcGUsePropSys_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcGUsePropSys.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "UPS-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcGUsePropSys.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcGUsePropSys.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcGUsePropSys_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
'    If imPWStatus = 0 Then
'        imPWStatus = -2
'        If (Trim$(tgUrf(0).sName) <> sgCPName) Then
'            ilPasswordOk = igPasswordOk
'            CSPWord.Show vbModal
'            DoEvents
'            If Not igPasswordOk Then
'                imPWStatus = -1
'                igPasswordOk = ilPasswordOk
'                If tgSpf.sGUsePropSys = "Y" Then 'Using Proposal System
'                    ckcGUsePropSys.Value = vbChecked
'                Else
'                    ckcGUsePropSys.Value = vbUnchecked
'                End If
'                cmcCancel.SetFocus
'                Exit Sub
'            Else
'                imPWStatus = 1
'                igPasswordOk = ilPasswordOk
'            End If
'        Else
'            imPWStatus = 1
'        End If
'    Else
'        If imPWStatus = -1 Then
'            If tgSpf.sGUsePropSys = "Y" Then 'Using Proposal System
'                ckcGUsePropSys.Value = vbChecked
'            Else
'                ckcGUsePropSys.Value = vbUnchecked
'            End If
'            cmcCancel.SetFocus
'        End If
'    End If
End Sub
Private Sub ckcGUsePropSys_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If imPWStatus = 0 Then
'        If (Trim$(tgUrf(0).sName) <> sgCPName) Then
'            ilPasswordOk = igPasswordOk
'            CSPWord.Show vbModal
'            DoEvents
'            If Not igPasswordOk Then
'                imPWStatus = -1
'                igPasswordOk = ilPasswordOk
'                If tgSpf.sGUsePropSys = "Y" Then 'Using Proposal System
'                    ckcGUsePropSys.Value = vbChecked
'                Else
'                    ckcGUsePropSys.Value = vbUnchecked
'                End If
'                cmcCancel.SetFocus
'                Exit Sub
'            Else
'                imPWStatus = 1
'                igPasswordOk = ilPasswordOk
'            End If
'        Else
'            imPWStatus = 1
'        End If
'    Else
'        If imPWStatus = -1 Then
'            If tgSpf.sGUsePropSys = "Y" Then 'Using Proposal System
'                ckcGUsePropSys.Value = vbChecked
'            Else
'                ckcGUsePropSys.Value = vbUnchecked
'            End If
'            cmcCancel.SetFocus
'        End If
'    End If
End Sub

Private Sub ckcInstallment_Click()
    Dim ilPasswordOk As Integer
    Dim ilRet As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcInstallment.Value = vbChecked Then
                ilRet = MsgBox("Air Time and NTR will be Combined into One Invoice, Continue", vbYesNo + vbDefaultButton2 + vbQuestion, "Override")
                If ilRet = vbYes Then
                    ilPasswordOk = igPasswordOk
                    sgPasswordAddition = "IN-"
                    mProtectChangesAllowed Start
                    CSPWord.Show vbModal
                    mProtectChangesAllowed Done
                    If Not igPasswordOk Then
                        ckcInstallment.Value = vbUnchecked
                    End If
                    sgPasswordAddition = ""
                    igPasswordOk = ilPasswordOk
                Else
                    ckcInstallment.Value = vbUnchecked
                End If
            End If
        Else
            ckcInstallment.Value = vbUnchecked
        End If
    End If
    If ckcInstallment.Value = vbUnchecked Then
        rbcInstRev(1).Value = True
        rbcInstRev(0).Enabled = False
        ckcInv(1).Enabled = True
    Else
        rbcInstRev(0).Enabled = True
        ckcInv(1).Value = vbChecked
        ckcInv(1).Enabled = False
        rbcBCombine(0).Value = True
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcInstallment_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub



Private Sub ckcInv_Click(Index As Integer)
    If Index = 1 Then
        If ckcInv(1).Value = vbUnchecked Then
            ckcInstallment.Value = vbUnchecked
        End If
    End If
    If Index = 5 Then
        If ckcInv(5).Value = vbUnchecked Then
            rbcPostRepAffidavit(1).Enabled = True
            rbcPostRepAffidavit(2).Enabled = True
        Else
            rbcPostRepAffidavit(1).Enabled = False
            rbcPostRepAffidavit(2).Enabled = False
            rbcPostRepAffidavit(1).Value = False
            rbcPostRepAffidavit(2).Value = False
        End If
    End If
    If Index = 9 Then
        If ckcInv(9).Value = vbUnchecked Then
            ckcInv(10).Value = vbUnchecked
            ckcInv(10).Enabled = False
        Else
            ckcInv(10).Enabled = True
        End If
    End If
    '10016
    If Index = INVAUTO Or Index = INVCOMBINE Then
        mInvoiceEmailOptions
    End If
    If Index = INVSELECTIVE Or Index = INVCOMBINE Then
        mInvoiceEmailOptions
    End If
    
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcInv_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub


Private Sub ckcInvoiceExport_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcInvoiceExport.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "RE-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcInvoiceExport.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcInvoiceExport.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
    If ckcInvoiceExport.Value = vbUnchecked Then
        ckcACodes(4).Enabled = False
        ckcACodes(4).Value = vbUnchecked
        lacAG(9).Visible = False
        edcExpt(3).Visible = False
    Else
        ckcACodes(4).Enabled = True
        lacAG(9).Visible = True
        edcExpt(3).Visible = True
    End If
End Sub

Private Sub ckcJelli_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcJelli.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "JE-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcJelli.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcJelli.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcJelli_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcMetroSplitCopy_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcMetroSplitCopy.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "MSC-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcMetroSplitCopy.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcMetroSplitCopy.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcMetroSplitCopy_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcOptionFields_Click(Index As Integer)
    Dim ilPasswordOk As Integer
    '10048 podcast
    If (Index = 14) Or (Index = 10) Or (Index = 7) Or (Index = 15) Or (Index = 16) Or (Index = 17) Or (Index = 22) Or (Index = 23) Or (Index = 24) Or (Index = 25) Or (Index = 26) Or (Index = 27) Or (Index = 31) Or (Index = PODAIRTIMECKC) Or (Index = PODSPOTSCKC) Or (Index = ADSERVERCKC) Then
        If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
            If igPasswordOk Then
                If ckcOptionFields(Index).Value = vbChecked Then
                    ilPasswordOk = igPasswordOk
                    If Index = 10 Then
                        sgPasswordAddition = "CBP-"
                    ElseIf Index = 14 Then
                        sgPasswordAddition = "EMD-"
                    ElseIf Index = 15 Then
                        sgPasswordAddition = "FM-"
                    ElseIf Index = 16 Then
                        sgPasswordAddition = "AC-"
                    ElseIf Index = 17 Then
                        sgPasswordAddition = "PB-"
                    ElseIf Index = 22 Then
                        sgPasswordAddition = "AA-"
                    ElseIf Index = 23 Then
                        sgPasswordAddition = "RC-"
                    ElseIf Index = 24 Then
                        sgPasswordAddition = "AOE-"
                    ElseIf Index = 25 Then
                        sgPasswordAddition = "CAI-"
                    ElseIf Index = 26 Then
                        sgPasswordAddition = "RS-"
                    ElseIf Index = 27 Then
                        sgPasswordAddition = "RCS-"
                    ElseIf Index = 31 Then
                        sgPasswordAddition = "CRE-"
                    ElseIf Index = PODAIRTIMECKC Then '28
                        sgPasswordAddition = "AIR-"
                    ElseIf Index = PODSPOTSCKC Then '29
                        sgPasswordAddition = "PSP-"
                    ElseIf Index = ADSERVERCKC Then '31
                        sgPasswordAddition = "AS-"
                    Else
                        sgPasswordAddition = "RS-"
                    End If
                    mProtectChangesAllowed Start
                    CSPWord.Show vbModal
                    mProtectChangesAllowed Done
                    If Not igPasswordOk Then
                        ckcOptionFields(Index).Value = vbUnchecked
                    End If
                    sgPasswordAddition = ""
                    igPasswordOk = ilPasswordOk
                End If
            Else
                ckcOptionFields(Index).Value = vbUnchecked
            End If
        End If
    End If
    If Index = 4 Then
        If ckcOptionFields(4).Value = vbChecked Then
            ckcOptionFields(11).Value = vbUnchecked
        End If
    End If
    If Index = 11 Then
        If ckcOptionFields(11).Value = vbChecked Then
            ckcOptionFields(4).Value = vbUnchecked
        End If
    End If
    If Index = 17 Then
        If ckcOptionFields(17).Value = vbChecked Then
            ckcOptionFields(18).Enabled = True
            ckcOptionFields(19).Enabled = True
            ckcOptionFields(20).Enabled = True
            ckcOptionFields(21).Enabled = True
        Else
            ckcOptionFields(17).Value = vbUnchecked
            ckcOptionFields(18).Value = vbUnchecked
            ckcOptionFields(19).Value = vbUnchecked
            ckcOptionFields(20).Value = vbUnchecked
            ckcOptionFields(21).Value = vbUnchecked
            ckcOptionFields(18).Enabled = False
            ckcOptionFields(19).Enabled = False
            ckcOptionFields(20).Enabled = False
            ckcOptionFields(21).Enabled = False
        End If
    End If
    '10048
    mPodcastOptions Index
    If Not bmIgnoreChange Then
      mChangeOccured
    End If

End Sub

Private Sub ckcOptionFields_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcOverrideOptions_Click(Index As Integer)
    Dim slStr As String

    If Index = 3 Then
        If ckcOverrideOptions(3).Value = vbChecked Then
            slStr = Trim$(edcLnOverride(0).Text)
            If (slStr = "") Or (Val(slStr) = 0) Then
                edcLnOverride(0).Text = "25"
            End If
            edcLnOverride(0).Enabled = True
        Else
            edcLnOverride(0).Enabled = False
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If

End Sub

Private Sub ckcOverrideOptions_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcPodShowWk_Click()
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcPrefeed_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcPrefeed.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "PF-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcPrefeed.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcPrefeed.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcPrefeed_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub


Private Sub ckcProposalXML_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcProposalXML.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "PX-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcProposalXML.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcProposalXML.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcProposalXML_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcRegionalCopy_Click()
'12/9/14: Removed Regional Copy as not used any longer

'    Dim ilPasswordOk As Integer
'
'    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
'        If igPasswordOk Then
'            If ckcRegionalCopy.Value = vbChecked Then
'                ilPasswordOk = igPasswordOk
'                sgPasswordAddition = "RC-"
'                mProtectChangesAllowed Start
'                CSPWord.Show vbModal
'                mProtectChangesAllowed Done
'                If Not igPasswordOk Then
'                    ckcRegionalCopy.Value = vbUnchecked
'                Else
'                    ckcUsingSplitCopy.Value = vbUnchecked
'                End If
'                sgPasswordAddition = ""
'                igPasswordOk = ilPasswordOk
'            End If
'        Else
'            ckcRegionalCopy.Value = vbUnchecked
'        End If
'    End If
'    If Not bmIgnoreChange Then
'      mChangeOccured
'    End If
End Sub

Private Sub ckcRegionalCopy_GotFocus()
'12/9/14: Removed Regional Copy as not used any longer
'    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcRemoteExport_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcRemoteExport.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "RE-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcRemoteExport.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcRemoteExport.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcRemoteExport_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcRemoteImport_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcRemoteImport.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "RI-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcRemoteImport.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcRemoteImport.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcRemoteImport_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcRevenueExport_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcRevenueExport.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "RE-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcRevenueExport.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcRevenueExport.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcRevenueExport_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcRN_Net_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcRN_Net.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "RNN-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcRN_Net.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcRN_Net.Value = vbUnchecked
        End If
    End If
    If ckcRN_Net.Value = vbChecked Then
        ckcRN_Rep.Value = vbUnchecked
        udcSiteTabs.Action 6, 2
    Else
        If ckcRN_Rep.Value = vbUnchecked Then
            udcSiteTabs.Action 6, 0
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcRN_Net_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcRN_Rep_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcRN_Rep.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "RNR-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcRN_Rep.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcRN_Rep.Value = vbUnchecked
        End If
    End If
    If ckcRN_Rep.Value = vbChecked Then
        ckcRN_Net.Value = vbUnchecked
        udcSiteTabs.Action 6, 1
    Else
        If ckcRN_Net.Value = vbUnchecked Then
            udcSiteTabs.Action 6, 0
        End If
        '1/8/10:  Allow Time posting if Using Rep or using RN_Rep
        'If (rbcPostRepAffidavit(3).Value) Then
        If (ckcUsingRep.Value = vbUnchecked) And (rbcPostRepAffidavit(3).Value) Then
            rbcPostRepAffidavit(0).Value = True
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcRN_Rep_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcRUseTMP_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcSales_Click(Index As Integer)
    Dim Value As Integer
    
    If Index = 10 Then  'BBToAff
        If ckcSales(10).Value = vbChecked Then
            If (ckcGUseAffSys(0).Value = vbChecked) And (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) <= 0) And (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) <= 0) Then
                'No support code placed into Station Feed; Exports : NY, Dallas and Phonxie; CnC
                MsgBox "Billboards not allowed in conjunction with Affiliate System at this time."
                ckcSales(10).Value = vbUnchecked
            End If
        End If
    End If
    If (Index = 5) Or (Index = 6) Or (Index = 7) Then   'CPackage
        Value = False
        If ckcSales(Index).Value = vbChecked Then
            Value = True
        End If
        'End of coded added
        If Index = 5 Then
            If Value Then
                rbcBPkageGenMeth(0).Enabled = True
                rbcBPkageGenMeth(1).Enabled = True
            Else
                rbcBPkageGenMeth(0).Enabled = False
                rbcBPkageGenMeth(1).Enabled = False
            End If
        End If
    End If
    If Not bmIgnoreChange Then
        mChangeOccured
    End If
End Sub

Private Sub ckcSales_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcSalesForce_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcSalesForce.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "SF-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcSalesForce.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcSalesForce.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcSalesForce_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcSDelivery_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub ckcSMktBase_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub ckcSSelling_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcStrongPassword_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcStrongPassword.Value = vbUnchecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "SP-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcStrongPassword.Value = vbChecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcStrongPassword.Value = vbChecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcStrongPassword_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcSuppressTimeForm1_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcTaxOn_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingBarter_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingBarter.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "UB-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingBarter.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingBarter.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingBarter_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingBBs_Click()
    Dim ilPasswordOk As Integer


    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingBBs.Value = vbChecked Then
                If (ckcGUseAffFeed.Value = vbChecked) And (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) <= 0) And (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) <= 0) Then
                    'No support code placed into Station Feed; Exports : NY, Dallas and Phonxie; CnC
                    MsgBox "Billboards not allowed in conjunction with Station Feed at this time."
                    ckcUsingBBs.Value = vbUnchecked
                Else
                    ilPasswordOk = igPasswordOk
                    sgPasswordAddition = "UBB-"
                    mProtectChangesAllowed Start
                    CSPWord.Show vbModal
                    mProtectChangesAllowed Done
                    If Not igPasswordOk Then
                        ckcUsingBBs.Value = vbUnchecked
                    End If
                    sgPasswordAddition = ""
                    igPasswordOk = ilPasswordOk
                End If
            End If
        Else
            ckcUsingBBs.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingBBs_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingLiveCopy_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingLiveCopy.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "ULC-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingLiveCopy.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingLiveCopy.Value = vbUnchecked
        End If
    End If
    'If (ckcUsingLiveCopy.Value = vbUnchecked) And (ckcCopy(1).Value = vbUnchecked) Then
    '    ckcOverrideOptions(5).Enabled = False
    '    ckcOverrideOptions(5).Value = vbUnchecked
    '    ckcCopy(5).Enabled = True
    '    ckcCopy(6).Enabled = False
    '    ckcCopy(6).Value = vbUnchecked
    '    ckcCopy(7).Enabled = False
    '    ckcCopy(7).Value = vbUnchecked
    '    ckcCopy(8).Enabled = False
    '    ckcCopy(8).Value = vbUnchecked
    '    ckcCopy(9).Enabled = False
    '    ckcCopy(9).Value = vbUnchecked
    '    ckcCopy(10).Enabled = False
    '    ckcCopy(10).Value = vbUnchecked
    'Else
    '    ckcOverrideOptions(5).Enabled = True
    '    ckcCopy(5).Enabled = True
    '    ckcCopy(6).Enabled = True
    '    ckcCopy(7).Enabled = True
    '    ckcCopy(8).Enabled = True
    '    ckcCopy(9).Enabled = True
    '    ckcCopy(10).Enabled = True
    'End If
    mSetAudioTypes
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingLiveCopy_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingLiveLog_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingLiveLog.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "ULL-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingLiveLog.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingLiveLog.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingLiveLog_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingMatrix_Click(Index As Integer)
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingMatrix(Index).Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                If Index <> 2 Then
                    sgPasswordAddition = "UME-"
                Else
                    sgPasswordAddition = "TE-"
                End If
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingMatrix(Index).Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingMatrix(Index).Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingMatrix_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingMultiMedia_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingMultiMedia.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "UMM-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingMultiMedia.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingMultiMedia.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingMultiMedia_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingNTR_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingNTR.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "UN-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingNTR.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingNTR.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingNTR_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub


Private Sub ckcUsingRep_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingRep.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "UR-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingRep.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingRep.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingRep_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingRevenue_Click()
 Dim ilPasswordOk As Integer
    '9-28-05
    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingRevenue.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "URE-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingRevenue.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingRevenue.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingRevenue_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingSpecialResearch_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingSpecialResearch.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "USR-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingSpecialResearch.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingSpecialResearch.Value = vbUnchecked
        End If
    End If
    If ckcUsingSpecialResearch.Value = vbChecked Then
        edcComment(3).Visible = True
        lacComment(3).Caption = "Research Estimate Comment"
    ElseIf udcSiteTabs.Research(31) = vbChecked Then
        edcComment(3).Visible = True
        lacComment(3).Caption = "Research Override Comment"
    Else
        edcComment(3).Visible = False
        lacComment(3).Caption = ""
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingSpecialResearch_GotFocus()
    mCtrlGotFocusAndIgnoreChange cmcCommand(COMMAND_CANCEL)
End Sub

Private Sub ckcUsingSplitCopy_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingSplitCopy.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "USC-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingSplitCopy.Value = vbUnchecked
                Else
                    ckcRegionalCopy.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingSplitCopy.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingSplitCopy_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingSplitNetworks_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingSplitNetworks.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "USN-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingSplitNetworks.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingSplitNetworks.Value = vbUnchecked
        End If
        If ckcUsingSplitNetworks.Value = vbUnchecked Then
            ckcRegionMixLen.Value = vbUnchecked
            ckcRegionMixLen.Enabled = False
        Else
            ckcRegionMixLen.Enabled = True
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingSplitNetworks_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingSports_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingSports.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "US-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingSports.Value = vbUnchecked
                    'plcSports.Enabled = False
                Else
                    'plcSports.Enabled = True
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingSports.Value = vbUnchecked
            'plcSports.Enabled = False
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingSports_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcUsingTraffic_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcUsingTraffic.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "UT-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcUsingTraffic.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcUsingTraffic.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcUsingTraffic_GotFocus()
    mCtrlGotFocusAndIgnoreChange cmcCommand(COMMAND_CANCEL)
End Sub

Private Sub ckcVCreative_Click()
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcVCreative.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "VC-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcVCreative.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcVCreative.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub ckcVCreative_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

'8/17/21 - JW - TTP 10233 - Audacy: line summary export
Private Sub ckcCntrLineExport_Click()
    Dim ilPasswordOk As Integer
    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcCntrLineExport.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "UWE-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcCntrLineExport.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcCntrLineExport.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

'TTP 10205 - 6/21/21 - JW - WO Invoice Export Enable Checkbox
Private Sub ckcWOInvoiceExport_Click()
    Dim ilPasswordOk As Integer
    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcWOInvoiceExport.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "UWE-"
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcWOInvoiceExport.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcWOInvoiceExport.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
    
    'L.Bianchi 06/10/2021
    If ckcWOInvoiceExport.Value = vbUnchecked Then
        edcIE(0).Enabled = False
        edcIE(1).Enabled = False
        edcIE(2).Enabled = False
        'edcIE(0).Text = ""
        'edcIE(1).Text = ""
        'edcIE(2).Text = ""
    End If
    'L.Bianchi 06/10/2021
    If ckcWOInvoiceExport.Value = vbChecked Then
        edcIE(0).Enabled = True
        edcIE(1).Enabled = True
        edcIE(2).Enabled = True
        edcIE(0).Text = Trim$(tgSpfx.sInvExpProperty)
        edcIE(1).Text = Trim$(tgSpfx.sInvExpPrefix)
        edcIE(2).Text = Trim$(tgSpfx.sInvExpBillGroup)
    End If
   
End Sub

Private Sub ckcXDSBy_Click(Index As Integer)
    If Not bmIgnoreChange Then
        mChangeOccured
        If Index = BYBREAK Then
            If ckcXDSBy(Index).Value = vbChecked Then
                ckcXDSBy(2).Enabled = True
                ckcXDSBy(3).Enabled = True
                ckcXDSBy(ASTBREAK).Enabled = True
            Else
                ckcXDSBy(2).Enabled = False
                ckcXDSBy(3).Enabled = False
                ckcXDSBy(ASTBREAK).Enabled = False
                ckcXDSBy(2).Value = vbUnchecked
                ckcXDSBy(3).Value = vbUnchecked
                ckcXDSBy(ASTBREAK).Value = vbUnchecked
            End If
        '9114
        ElseIf Index = BYISCI Then
            If ckcXDSBy(Index).Value = vbChecked Then
                ckcXDSBy(ASTISCI).Enabled = True
            Else
                ckcXDSBy(ASTISCI).Enabled = False
                ckcXDSBy(ASTISCI).Value = vbUnchecked
            End If
        End If
        '100012 9664  astbreak = 4
'        If Index = 3 Or Index = ASTBREAK Then
'            Dim ilNotChosen As Integer
'            If Index = 3 Then
'                ilNotChosen = ASTBREAK
'            Else
'                ilNotChosen = 3
'            End If
'            ckcXDSBy(ilNotChosen).Value = vbUnchecked
'        End If
  End If
End Sub

Private Sub ckcXDSBy_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub cmcCommand_Click(Index As Integer)

    'TTP 10626 JJB 2023-01-10
    'Converted all the main form buttons to an array to increase the number of available controls to use (MAX = 256)
        
    Select Case Index
        Case COMMAND_DONE:
            If (Not imUpdateAllowed) Or (Not igPasswordOk) Then
                mTerminate
                Exit Sub
            End If
            Update_Data
            If imUpdateError Then
                Exit Sub
            End If
            mTerminate
        Case COMMAND_CANCEL:
            mTerminate
        Case COMMAND_SAVE:
            Update_Data
        Case COMMAND_UNDO:
            Dim ilRet As Integer
            sgSpfStamp = "~"    'Force read
            gSpfRead
            ilRet = mSafReadRec()
            ilRet = mSiteReadRec()
            ilRet = mNrfReadRec()
            mMoveRecToCtrl
        Case COMMAND_REPORT:
            Dim slStr As String
            igRptCallType = SITELIST
            
            If igTestSystem Then
                slStr = "SiteOpt^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
            Else
                slStr = "SiteOpt^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
            End If
           
            sgCommandStr = slStr
            RptList.Show vbModal
    End Select
    
End Sub

Private Sub cmcCommand_GotFocus(Index As Integer)
    'TTP 10626 JJB 2023-01-10
    'Converted all the main form buttons to an array to increase the number of available controls to use (MAX = 256)
    mCtrlGotFocusAndIgnoreChange cmcCommand(Index)
End Sub

Private Sub cmcLevelPrice_Click()
    Dim slLow As String
    Dim slHigh As String
    Dim slInc As String
    Dim ilLoop As Integer
    Dim llInc As Long

    slLow = Trim$(edcSchedule(0).Text)
    If slLow = "" Then
        edcSchedule(1).Text = ""
        For ilLoop = LBound(lmSSave) To UBound(lmSSave) Step 1
            lmSSave(ilLoop) = 0
        Next ilLoop
        pbcSchedule.Cls
        pbcSchedule_Paint
        Exit Sub
    End If
    If Val(slLow) <= 0 Then
        For ilLoop = LBound(lmSSave) To UBound(lmSSave) Step 1
            lmSSave(ilLoop) = 0
        Next ilLoop
        pbcSchedule.Cls
        pbcSchedule_Paint
        Exit Sub
    End If
    slHigh = Trim$(edcSchedule(1).Text)
    If slHigh = "" Then
        Beep
        edcSchedule(1).SetFocus
        Exit Sub
    End If
    If Val(slHigh) - Val(slLow) < 12 Then
        Beep
        edcSchedule(1).SetFocus
        Exit Sub
    End If
    slInc = gSubStr(slHigh, slLow)
    slInc = gDivStr(slInc, "12")
    llInc = Val(slInc)
    lmSSave(1) = Val(slLow)
    For ilLoop = 2 To 13 Step 1
        lmSSave(ilLoop) = lmSSave(ilLoop - 1) + llInc
    Next ilLoop
    lmSSave(13) = Val(slHigh)
    pbcSchedule.Cls
    pbcSchedule_Paint
End Sub

Private Sub cmcLevelPrice_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    imSRowNo = -1
End Sub

Private Sub cmcRCorpCal_Click()
    If (imUpdateAllowed) And (igPasswordOk) Then
        igUpdateOk = True
    Else
        igUpdateOk = False
    End If
    CorpCal.Show vbModal
End Sub

Private Sub Update_Data()

    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slStr As String
    Dim slMsg As String
    Dim ilRecLen As Integer     'SPF record length
    Dim ilFirst As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim hlSpf As Integer        'site Option file handle
    Dim hlSpfx As Integer       'TTP 10205 - 6/21/21 - JW - Extended Site Option file Handle
    Dim tlSaf As SAF
    Dim tlSite As SITE
    Dim tlNrf As NRF

    imUpdateError = False
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    imUpdateError = True
    'Dan M 4/28/09
    If imChangesOccured <= 0 Then
        imUpdateError = False
        Exit Sub
    ElseIf imChangesOccured > igChangesAllowed Or bmExceededChanges Then
        gMsgBox "You have made too many changes.  Current Changes will not be saved.", vbInformation, "Save Cancelled"
        bmExceededChanges = True
        mChangeLabel
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass  'Wait
    slStr = edcBBillDate(0).Text
    If slStr <> "" Then
        If Not gValidDate(slStr) Then
            Screen.MousePointer = vbDefault
            MsgBox "Last Standard Broadcast Billed Date not Valid"
            Exit Sub
        End If
        slDate = gObtainEndStd(slStr)
        If gDateValue(slStr) <> gDateValue(slDate) Then
            Screen.MousePointer = vbDefault
            MsgBox "Last Standard Broadcast Billed Date not End of Standard Broadcast Month"
            Exit Sub
        End If
    End If
    slStr = edcBBillDate(1).Text
    If slStr <> "" Then
        If Not gValidDate(slStr) Then
            Screen.MousePointer = vbDefault
            MsgBox "Last Calendar Billed Date not Valid"
            Exit Sub
        End If
        slDate = gObtainEndCal(slStr)
        If gDateValue(slStr) <> gDateValue(slDate) Then
            Screen.MousePointer = vbDefault
            MsgBox "Last Calendar Billed Date not End of Month"
            Exit Sub
        End If
    End If
    slStr = edcBBillDate(3).Text
    If slStr <> "" Then
        If Not gValidDate(slStr) Then
            Screen.MousePointer = vbDefault
            MsgBox "Last Week Billed Date not Valid"
            Exit Sub
        End If
        slDate = gObtainNextSunday(slStr)
        If gDateValue(slStr) <> gDateValue(slDate) Then
            Screen.MousePointer = vbDefault
            MsgBox "Last Week Billed Date not End of Week"
            Exit Sub
        End If
    End If
    If ckcUsingSpecialResearch.Value = vbChecked Then
        If Trim$(edcComment(3).Text) = "" Then
            Screen.MousePointer = vbDefault
            MsgBox "Research Estimate Comment must be defined before save allowed"
            Exit Sub
        End If
    End If
    If udcSiteTabs.Research(31) = vbChecked Then
        If Trim$(edcComment(3).Text) = "" Then
            Screen.MousePointer = vbDefault
            MsgBox "Research Override Comment must be defined before save allowed"
            Exit Sub
        End If
    End If
    If (ckcUsingSpecialResearch.Value = vbChecked) And (udcSiteTabs.Research(31) = vbChecked) Then
        Screen.MousePointer = vbDefault
        MsgBox "Research Override and Research Estimate features can not both be turned on at the same time, turn one Off"
        Exit Sub
    End If
    If (Val(edcGRetain(2).Text) <= 0) Or (Trim$(edcGRetain(2).Text) = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Number of months to retain Traffic spots must be defined"
        Exit Sub
    End If
    If ckcGUseAffSys(0).Value = vbChecked Then  'Using Affiliate System
        If (Val(edcGRetain(1).Text) <= 0) Or (Trim$(edcGRetain(1).Text) = "") Then
            Screen.MousePointer = vbDefault
            MsgBox "Number of months to retain Affiliate spots must be defined"
            Exit Sub
        End If
        If Val(edcGRetain(2).Text) < Val(edcGRetain(1).Text) Then
            Screen.MousePointer = vbDefault
            MsgBox "Number of months to retain Traffic spots must be equal or larger than the number of months to retain Affiliate spots"
            Exit Sub
        End If
    End If
    If ckcOptionFields(0).Value = vbChecked Then  'Option Fields (Right to Left):0=Projection;1=Bus Cat; 2=Share; 3=Rev Set; 4=Guar; 5=Billing Cycle; 6=Co-op; 7=Research
        If (Val(edcGRetain(3).Text) <= 0) Or (Trim$(edcGRetain(3).Text) = "") Then
            Screen.MousePointer = vbDefault
            MsgBox "Number of months to retain Projections must be defined"
            Exit Sub
        End If
    End If
    If ckcGUsePropSys.Value = vbChecked Then 'Using Proposal System
        If (Val(edcGRetain(5).Text) <= 0) Or (Trim$(edcGRetain(5).Text) = "") Then
            Screen.MousePointer = vbDefault
            MsgBox "Number of months to retain Proposals must be defined"
            Exit Sub
        End If
    End If
    If (Val(edcGRetain(7).Text) <= 0) Or (Trim$(edcGRetain(7).Text) = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Number of months to retain Expired Contracts must be defined"
        Exit Sub
    End If
    If Val(edcGRetain(7).Text) < Val(edcGRetain(2).Text) Then
        Screen.MousePointer = vbDefault
        MsgBox "Number of months to retain Expired Contracts must be equal or larger than the number of months to retain Traffic spots"
        Exit Sub
    End If
    If (Val(edcGRetain(8).Text) <= 0) Or (Trim$(edcGRetain(8).Text) = "") Then
        Screen.MousePointer = vbDefault
        MsgBox "Number of months to retain Payment and Revenue History must be defined"
        Exit Sub
    End If
    If Val(edcGRetain(8).Text) < Val(edcGRetain(7).Text) Then
        Screen.MousePointer = vbDefault
        MsgBox "Number of months to retain Payment and Revenue History must be equal or larger than the number of months to retain Expired Contracts"
        Exit Sub
    End If
    If (ckcUsingSpecialResearch.Value = vbUnchecked) And (udcSiteTabs.Research(31) = vbUnchecked) Then
        edcComment(3).Text = ""
    End If
    If (ckcOptionFields(17).Value = vbChecked) Then
        If (ckcOptionFields(19).Value = vbUnchecked) And (ckcOptionFields(20).Value = vbUnchecked) And (ckcOptionFields(21).Value = vbUnchecked) Then
            Screen.MousePointer = vbDefault
            MsgBox "Using Programmatic feature: CPP, CPM or Price must be checked"
            Exit Sub
        End If
    End If
    gGetSyncDateTime slSyncDate, slSyncTime
    ilFirst = True
    hlSpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ilRecLen = Len(tgSpf) 'btrRecordLength(hlSpf)  'Get and save record length
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, "cmcCommand(COMMAND_SAVE) (btrOpen):" & "Spf.Btr", SiteOpt
    On Error GoTo 0
    Do
        ilRet = btrGetFirst(hlSpf, tgSpf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_REC_LOCKED) Or (ilRet = BTRV_ERR_FILE_LOCKED) Then
            ilRet = btrClose(hlSpf)
            btrDestroy hlSpf
            Screen.MousePointer = vbDefault
            MsgBox "Unable to update Site Preference as Credit Alert being Set", vbCritical, "Site Preference"
            Exit Sub
        End If
        On Error GoTo cmcUpdateErr
        gBtrvErrorMsg ilRet, "cmcCommand(COMMAND_SAVE) (btrGetFirst):" & "Spf.Btr", SiteOpt
        On Error GoTo 0
        If (imUpdateAllowed) And (igPasswordOk) Then
            mMoveCtrlToRec
            tgSpf.lBCxfDisclaimer = lmBCxfDisclaimer     '2-20-03
            tgSpf.lCxfContrComment = lmCxfContrComment
            tgSpf.lCxfInsertComment = lmCxfInsertComment
            tgSpf.lCxfDemoEst = lmCxfEstComment
            If ilFirst Then
                slMsg = mUpdateComments(tgSpf.lBCxfDisclaimer, imInvCommentLen, smInvComment)
                lmBCxfDisclaimer = tgSpf.lBCxfDisclaimer
                'tmCxf.iStrLen = imInvCommentLen
                'tmCxf.sComment = Trim$(smInvComment)
                'imCxfRecLen = Len(tmCxf) - Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment)) ' + 2    '25 = fixed record length; 2=Length value which is part of the variable record
                'If tmCxf.lCode = 0 Then 'New
                '    If Len(Trim$(tmCxf.sComment)) > 2 Then  '-2 so control character at end not counted
                '        tmCxf.sComType = "D"
                '        tmCxf.sShProp = "N"
                '        tmCxf.sShSpot = "N"
                '        tmCxf.sShOrder = "N"
                '        tmCxf.sShInv = "Y"
                '        tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                '        tmCxf.lAutoCode = tmCxf.lCode
                '        ilRet = btrInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)
                '
                '        tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                '        tmCxf.lAutoCode = tmCxf.lCode
                '        tmCxf.iSourceID = tgUrf(0).iRemoteUserID
                '        gPackDate slSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
                '        gPackTime slSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
                '        imCxfRecLen = Len(tmCxf) - Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment))
                '        ilRet = btrUpdate(hmCxf, tmCxf, imCxfRecLen)
                '    Else
                '        tmCxf.lCode = 0
                '        ilRet = BTRV_ERR_NONE
                '    End If
                '    slMsg = "mSaveRec (btrInsert: Comment)"
                'Else 'Old record-Update
                '    If Len(Trim$(tmCxf.sComment)) > 2 Then  '-2 so the control character at end is not counted
                '        tmCxf.iSourceID = tgUrf(0).iRemoteUserID
                '        gPackDate slSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
                '        gPackTime slSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
                '        ilRet = btrUpdate(hmCxf, tmCxf, imCxfRecLen)
                '    Else
                '        ilRet = btrDelete(hmCxf)
                '        tmCxf.lCode = 0
                '        If tgSpf.sRemoteUsers = "Y" Then
                '            tmDsf.lCode = 0
                '            tmDsf.sFileName = "CXF"
                '            gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
                '            gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
                '            tmDsf.iRemoteID = tmCxf.iRemoteID
                '            tmDsf.lAutoCode = tmCxf.lAutoCode
                '            tmDsf.iSourceID = tgUrf(0).iRemoteUserID
                '            tmDsf.lCntrNo = 0
                '            ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
                '        End If
                '    End If
                '    slMsg = "mSaveRec (btrUpdate: Comment)"
                'End If
                'tgSpf.lBCxfDisclaimer = tmCxf.lCode

                '2-12-03  update contract comments
                slMsg = mUpdateComments(tgSpf.lCxfContrComment, imContrCommentLen, smContrComment)
                lmCxfContrComment = tgSpf.lCxfContrComment
                '2-12-03  update contract comments
                slMsg = mUpdateComments(tgSpf.lCxfInsertComment, imInsertCommentLen, smInsertComment)
                lmCxfInsertComment = tgSpf.lCxfInsertComment
                slMsg = mUpdateComments(tgSpf.lCxfDemoEst, imEstCommentLen, smEstComment)
                lmCxfEstComment = tgSpf.lCxfDemoEst
                '7942  XDS-Break
                If smHeadEndZoneChange <> UCase$(tmSaf.sXDSHeadEndZone) Then
                    mVatSetToGoToWeb 114
                    smHeadEndZoneChange = UCase$(tmSaf.sXDSHeadEndZone)
                End If
                On Error GoTo cmcUpdateErr
                gBtrvErrorMsg ilRet, slMsg, SiteOpt
                On Error GoTo 0
                ilFirst = False
            End If
        End If
        'tgSpf.lBCxfDisclaimer = tmCxf.lCode


        ''Locks
        'tgSpf.sLkCredit = edcGLock(0).Text
        'tgSpf.sLkLog = edcGLock(1).Text
        tgSpf.iUrfGCode = tgUrf(0).iCode
        ilRet = btrUpdate(hlSpf, tgSpf, ilRecLen)
        If tgSpf.iMnfInvTerms = 0 Then
            edcTerms.Text = sgDefaultTerms
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    slMsg = "cmcCommand(COMMAND_SAVE) (btrUpdate)"
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, slMsg, SiteOpt
    On Error GoTo 0
    ilRet = btrClose(hlSpf)
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, "cmcCommand(COMMAND_SAVE) (btrClose): " & "Spf.Btr", SiteOpt
    On Error GoTo 0
    btrDestroy hlSpf

    If ckcCSIBackup.Value = vbChecked Then
        ' Save the backup time back to the CSI_Server.ini file.
        If (Len(cbcBkupTime.Text) > 0) And (cbcBkupTime.Text <> "[None]") Then
            If Not gSaveINIValue(smCSIServerINIFile, "Backup", "StartTime", cbcBkupTime.Text) Then
                MsgBox "Unable to save the backup time."
            End If
        End If
        ' Save the day of weeks.
        smBUWeekDays = ""
        For ilLoop = 1 To 7
            If chkDOW(ilLoop - 1).Value = 1 Then
                smBUWeekDays = smBUWeekDays + "1"
            Else
                smBUWeekDays = smBUWeekDays + "0"
            End If
        Next
        If Not gSaveINIValue(smCSIServerINIFile, "Backup", "WeekDays", smBUWeekDays) Then
            MsgBox "Unable to save the backup day of week settings."
        End If
    End If
    
    'Schedule
    tmSaf.lStatementComment = lmCxfStatementComment
    slMsg = mUpdateComments(tmSaf.lStatementComment, imStatementCommentLen, smStatementComment)
    lmCxfStatementComment = tmSaf.lStatementComment
    tmSaf.lCitationComment = lmCxfCitationComment
    slMsg = mUpdateComments(tmSaf.lCitationComment, imCitationCommentLen, smCitationComment)
    lmCxfCitationComment = tmSaf.lCitationComment
    If tmSaf.iCode = 0 Then
        slMsg = "cmcCommand(COMMAND_SAVE) (btrInsert: SAF)"
        ilRet = btrInsert(hmSaf, tmSaf, imSafRecLen, INDEXKEY0)
    Else
        tmSafSrchKey1.iVefCode = 0
        ilRet = btrGetEqual(hmSaf, tlSaf, imSafRecLen, tmSafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        slMsg = "cmcCommand(COMMAND_SAVE) (btrUpdate: SAF)"
        ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
    End If

    bgHideHiddenLines = False
    If (Asc(tmSaf.sFeatures2) And HIDEHIDDENLINES) = HIDEHIDDENLINES Then
        bgHideHiddenLines = True
    End If

    'E-Mail
    If tgSite.lCode = 0 Then
        slMsg = "cmcCommand(COMMAND_SAVE) (btrInsert: Site.mkd)"
        ilRet = btrInsert(hmSite, tgSite, imSiteRecLen, INDEXKEY0)
    Else
        tmSiteSrchKey.lCode = 1
        ilRet = btrGetEqual(hmSite, tlSite, imSiteRecLen, tmSiteSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        slMsg = "cmcCommand(COMMAND_SAVE) (btrUpdate: Site.mkd)"
        ilRet = btrUpdate(hmSite, tgSite, imSiteRecLen)
    End If
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, slMsg, SiteOpt
    On Error GoTo 0

    If (ckcRN_Rep.Value = vbChecked) Or (ckcRN_Net.Value = vbChecked) Then
        If tgNrf.iCode = 0 Then
            slMsg = "cmcCommand(COMMAND_SAVE) (btrInsert: Nrf.Btr)"
            ilRet = btrInsert(hmNrf, tgNrf, imNrfRecLen, INDEXKEY0)
        Else
            tmNrfSrchKey.iCode = tgNrf.iCode
            ilRet = btrGetEqual(hmNrf, tlNrf, imNrfRecLen, tmNrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            slMsg = "cmcCommand(COMMAND_SAVE) (btrUpdate: Nrf.Btr)"
            ilRet = btrUpdate(hmNrf, tgNrf, imNrfRecLen)
        End If
        If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
            If tgNrf.sType = "R" Then
                sgRepDBID = tgNrf.sDBID
            End If
        ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
            If tgNrf.sType = "N" Then
                sgNetDBID = tgNrf.sDBID
            End If
        End If
    End If
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, slMsg, SiteOpt
    On Error GoTo 0

    ilRet = gObtainSAF()
    
    '07-13-15 Add or Update Cleint Name to EDS
    '08-25-15 404 {"Message":"No HTTP resource was found that matches the request URI 'https://www.counterpointcloud.com:44307/DataSync/ChangeNetworkName'."} {"oldName":"EntravisionXXX","newName":"Entravision"}
    If bgEDSIsActive Then
        If StrComp(smOrigClientName, Trim$(tgSpf.sGClient), vbTextCompare) <> 0 Then
            ilRet = gChangeNetworkName(Trim$(smOrigClientName), Trim$(tgSpf.sGClient))
            If Not ilRet Then
                'erro msg
            End If
        End If
    End If

    '----------------------------------
    'TTP 10205 - 6/21/21 - JW - Save Extended Site Options (SPFX)
    hlSpfx = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlSpfx, "", sgDBPath & "Spfx.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ilRecLen = Len(tgSpfx) 'btrRecordLength(hlSpfx)  'Get and save record length
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, "cmcCommand(COMMAND_SAVE) (btrOpen):" & "Spfx.Btr", SiteOpt
    On Error GoTo 0
    'Get 1st SPFX Record
    ilRet = btrGetFirst(hlSpfx, tgSpfx, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_REC_LOCKED) Or (ilRet = BTRV_ERR_FILE_LOCKED) Then
        ilRet = btrClose(hlSpfx)
        btrDestroy hlSpfx
        Screen.MousePointer = vbDefault
        MsgBox "Unable to update Extended Site Preference as Credit Alert being Set", vbCritical, "Site Preference"
        Exit Sub
    End If
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, "cmcCommand(COMMAND_SAVE) (btrGetFirst):" & "Spfx.Btr", SiteOpt
    'The above call wiptes out tgSpfx, read them in again
    mMoveCtrlToRec
    'Save SPFX Record
    ilRet = btrUpdate(hlSpfx, tgSpfx, ilRecLen)
    slMsg = "cmcCommand(COMMAND_SAVE) (btrUpdate)"
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, slMsg, SiteOpt
    On Error GoTo 0
    ilRet = btrClose(hlSpfx)
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, "cmcCommand(COMMAND_SAVE) (btrClose): " & "Spfx.Btr", SiteOpt
    On Error GoTo 0
    btrDestroy hlSpfx


    Screen.MousePointer = vbDefault
    'Traffic!lbcVehicle.Clear
    'Traffic!lbcVehicle.Tag = ""
    'sgVsfStamp = ""
    'sgUrfStamp = ""
    gUrfRead SiteOpt, sgUserName, True, tgUrf(), False  'Obtain user records
    igSiteSet = True
    imUpdateError = False
    Exit Sub
cmcUpdateErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    ilRet = btrClose(hlSpf)
    btrDestroy hlSpf
    imUpdateError = True
    Exit Sub
End Sub


Private Sub edcAPenny_GotFocus()
    'slStr = edcAPenny.Text
    'gUnformatStr slStr, UNFMTREMOVEDECZERO, slStr
    'edcAPenny.Text = slStr
    mCtrlGotFocusAndIgnoreChange edcAPenny
End Sub
Private Sub edcAPenny_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    ilPos = InStr(edcAPenny.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcAPenny.Text, ".")    'Disallow multi-decimal points
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
    slStr = edcAPenny.Text
    slStr = Left$(slStr, edcAPenny.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcAPenny.SelStart - edcAPenny.SelLength)
    If gCompNumberStr(slStr, "999.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcAPenny_LostFocus()
    'slStr = edcAPenny.Text
    'gFormatStr slStr, FMTDEFAULT, 2, slStr
    'edcAPenny.Text = slStr
End Sub
Private Sub edcExpt_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange edcExpt(Index)
End Sub
Private Sub edcExpt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slStr As String
    
    If Index = 0 Then
        ilKey = KeyAscii
        If Not gCheckKeyAscii(ilKey) Then
            KeyAscii = 0
            Exit Sub
        End If
    ElseIf Index = 1 Then
        slStr = UCase(Chr(KeyAscii))
        If (slStr <> "E") And (slStr <> "C") And (slStr <> "M") And (slStr <> "P") Then
            KeyAscii = 0
            Exit Sub
        End If
        KeyAscii = Asc(slStr)
        
    ElseIf Index = 2 Then
        'Filter characters (allow only BackSpace, numbers 0 thru 9
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        ilKey = KeyAscii
        If Not gCheckKeyAscii(ilKey) Then
            KeyAscii = 0
            Exit Sub
        End If
    ElseIf Index = 3 Then   'Invoice Export Delimiter characters
    End If
End Sub

Private Sub edcBarterLPD_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcBarterLPD_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcBBillDate_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub edcBBillDate_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    'ilKey = KeyAscii
    'If Not gCheckKeyAscii(ilKey) Then
    '    KeyAscii = 0
    '    Exit Sub
    'End If
End Sub

Private Sub edcBLogoSpaces_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange edcBLogoSpaces(Index)
End Sub

Private Sub edcBLogoSpaces_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY5)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcBNo_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange edcBNo(Index)
End Sub
Private Sub edcBNo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcBPayAddr_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange edcBPayAddr(Index)
End Sub
Private Sub edcBPayAddr_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcBPayName_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcBPayName
End Sub
Private Sub edcBPayName_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcBTTax_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

'Private Sub edcCNo_Change(Index As Integer)
'If Not bmIgnoreChange Then
'    mChangeOccuredTest "edcCNo_Change(" & Index & ")"
'End If
'End Sub

Private Sub edcCNo_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange edcCNo(Index)
End Sub
Private Sub edcCNo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub


Private Sub edcComment_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcDiscNo_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcDiscNo
End Sub
Private Sub edcDiscNo_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub


Private Sub edcEDI_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange edcEDI(Index)
End Sub

Private Sub edcEDI_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcGAddr_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange edcGAddr(Index)
End Sub
Private Sub edcGAddr_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer

    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcGAlertInterval_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcGAlertInterval
End Sub
Private Sub edcGAlertInterval_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcGClient_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcGClient
End Sub
Private Sub edcGClient_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcGClientAbbr_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcGClient
End Sub

Private Sub edcGClientAbbr_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcGRetain_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange edcGRetain(Index)
End Sub
Private Sub edcGRetain_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    If Index = 0 Then
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Else
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcGRetainPassword_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcGRetainPassword
End Sub
Private Sub edcGRetainPassword_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Sub edcInvExportId_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Sub edcInvExportId_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer

    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcLevelPrice_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcLevelPrice
End Sub


Private Sub edcLnOverride_GotFocus(Index As Integer)
    Dim ilPasswordOk As Integer

    If Index = 0 Then
        If imPWPrefSpotPct = 0 Then
            If (Trim$(tgUrf(0).sName) <> sgCPName) Then
                If igPasswordOk Then
                    imPWPrefSpotPct = 1
                    ilPasswordOk = igPasswordOk
                    sgPasswordAddition = "PPS-"
                    mProtectChangesAllowed Start
                    CSPWord.Show vbModal
                    mProtectChangesAllowed Done
                    If Not igPasswordOk Then
                        imPWPrefSpotPct = 0
                        cmcCommand(COMMAND_CANCEL).SetFocus
                    End If
                    igPasswordOk = ilPasswordOk
                    sgPasswordAddition = ""
                    imPWPrefSpotPct = 2
                    mCtrlGotFocusAndIgnoreChange edcLnOverride(Index)
                Else
                    cmcCommand(COMMAND_CANCEL).SetFocus
                End If
            End If
        End If
    Else
        If imWk1stSolo = 0 Then
            If (Trim$(tgUrf(0).sName) <> sgCPName) Then
                If igPasswordOk Then
                    imWk1stSolo = 1
                    ilPasswordOk = igPasswordOk
                    sgPasswordAddition = "W1S-"
                    mProtectChangesAllowed Start
                    CSPWord.Show vbModal
                    mProtectChangesAllowed Done
                    If Not igPasswordOk Then
                        imWk1stSolo = 0
                        cmcCommand(COMMAND_CANCEL).SetFocus
                    End If
                    igPasswordOk = ilPasswordOk
                    sgPasswordAddition = ""
                    imWk1stSolo = 2
                    If (edcLnOverride(Index).Text = "") Then
                        edcLnOverride(Index).Text = ".75"
                    End If
                    mCtrlGotFocusAndIgnoreChange edcLnOverride(Index)
                Else
                    cmcCommand(COMMAND_CANCEL).SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub edcLnOverride_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilPos As Integer    'Decimal point position
    Dim ilKey As Integer
    Dim slStr As String

    If Index = 0 Then
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        ilKey = KeyAscii
        If Not gCheckKeyAscii(ilKey) Then
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcLnOverride(Index).Text
        slStr = Left$(slStr, edcLnOverride(Index).SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcLnOverride(Index).SelStart - edcLnOverride(Index).SelLength)
        If gCompNumberStr(slStr, "100") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Else
        ilPos = InStr(edcRB.SelText, ".")
        If ilPos = 0 Then
            ilPos = InStr(edcRB.Text, ".")    'Disallow multi-decimal points
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
        ilKey = KeyAscii
        If Not gCheckKeyAscii(ilKey) Then
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcLnOverride(Index).Text
        slStr = Left$(slStr, edcLnOverride(Index).SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcLnOverride(Index).SelStart - edcLnOverride(Index).SelLength)
        If gCompNumberStr(slStr, ".99") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        If gCompNumberStr(slStr, ".1") < 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcLnOverride_LostFocus(Index As Integer)
    If Index = 0 Then
        If imPWPrefSpotPct <> 1 Then
            imPWPrefSpotPct = 0
        End If
    Else
        If imWk1stSolo <> 1 Then
            imWk1stSolo = 0
        End If
    End If
End Sub

Private Sub edcRB_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcRB
End Sub
Private Sub edcRB_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer    'Decimal point position
    Dim ilKey As Integer
    ilPos = InStr(edcRB.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcRB.Text, ".")    'Disallow multi-decimal points
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
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRCollectContact_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcRCollectContact
End Sub
Private Sub edcRCollectContact_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRCreditDate_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcRCreditDate
End Sub
Private Sub edcRCreditDate_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRCRP_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcRCRP
End Sub
Private Sub edcRCRP_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRNewCntr_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub edcRNewCntr_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRNRP_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcRNRP
End Sub
Private Sub edcRNRP_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRPctCredit_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcRPctCredit
End Sub
Private Sub edcRPctCredit_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer    'Decimal point position
    Dim ilKey As Integer
    ilPos = InStr(edcRPctCredit.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcRPctCredit.Text, ".")    'Disallow multi-decimal points
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
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRPRP_GotFocus()
    mCtrlGotFocusAndIgnoreChange edcRPRP
End Sub
Private Sub edcRPRP_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub


Private Sub edcSageIE_Change(Index As Integer)
    'TTP 10626 JJB 2023-01-10
    'Converted all the main form buttons to an array to increase the number of available controls to use (MAX = 256)
    If Not bmIgnoreChange Then
          mChangeOccured
          edcSageIE(Index).ToolTipText = edcSageIE(Index).Text
    End If
End Sub

Private Sub edcSageIE_GotFocus(Index As Integer)

    'TTP 10626 JJB 2023-01-10
    'Converted all the main form buttons to an array to increase the number of available controls to use (MAX = 256)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcSageIE_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'TTP 10626 JJB 2023-01-10
    'Converted all the main form buttons to an array to increase the number of available controls to use (MAX = 256)
    
    edcSageIE(Index).ToolTipText = edcSageIE(Index).Text
End Sub

Private Sub edcSales_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcSales_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim slStr As String
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    If Index = 0 Then
        slStr = edcSales(0).Text
        slStr = Left$(slStr, edcSales(0).SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSales(0).SelStart - edcSales(0).SelLength)
        If gCompNumberStr(slStr, "10000") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    ElseIf Index = 1 Then
        slStr = edcSales(1).Text
        slStr = Left$(slStr, edcSales(1).SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSales(1).SelStart - edcSales(1).SelLength)
        If gCompNumberStr(slStr, "40") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    ElseIf Index = 7 Then
        slStr = edcSales(7).Text
        slStr = Left$(slStr, edcSales(7).SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSales(7).SelStart - edcSales(7).SelLength)
        If gCompNumberStr(slStr, "13") > 0 Then     '9-16-11 chg max to 13
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    ElseIf Index = 8 Then
        slStr = edcSales(8).Text
        slStr = Left$(slStr, edcSales(8).SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSales(8).SelStart - edcSales(8).SelLength)
        If gCompNumberStr(slStr, "12") > 0 Then     '9-16-11 chg max to 12
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcSchedule_GotFocus(Index As Integer)
    mSSetShow imSBoxNo
    imSBoxNo = -1
    imSRowNo = -1
    mCtrlGotFocusAndIgnoreChange edcSchedule(Index)
End Sub

Private Sub edcSchedule_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    If Index = 2 Then
        'Filter characters (allow only BackSpace, numbers 0 thru 9
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcTerms_GotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
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
    'mInit sets Enabled for the objects
    If (igWinStatus(SITELIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        'plcGeneral.Enabled = False
        'plcSales.Enabled = False
        'plcAgyAdv.Enabled = False
        'plcInv.Enabled = False
        'plcCntr.Enabled = False
        'plcAccount.Enabled = False
        imUpdateAllowed = False
    Else
        'plcGeneral.Enabled = True
        'plcSales.Enabled = True
        'plcAgyAdv.Enabled = True
        'plcInv.Enabled = True
        'plcCntr.Enabled = True
        'plcAccount.Enabled = True
        'If igPassword = false, then only can update lock values
        imUpdateAllowed = True
    End If
    'If Not igPasswordOk Then
    '    plcGeneral.Enabled = False
    '    plcSales.Enabled = False
    '    plcAgyAdv.Enabled = False
    '    plcInv.Enabled = False
    '    plcCntr.Enabled = False
    '    plcAccount.Enabled = False
    '    'imUpdateAllowed = False    'allow only locks to be updated
    'Else
    '    plcGeneral.Enabled = True
    '    plcSales.Enabled = True
    '    plcAgyAdv.Enabled = True
    '    plcInv.Enabled = True
    '    plcCntr.Enabled = True
    '    plcAccount.Enabled = True
    'End If
    'Allow locks to be removed
    'If Not imUpdateAllowed Then
    '    cmcUpdate.Enabled = False
    'Else
    ' Dan M 4/21/09 would like to get rid of this.
      '  cmcUpdate.Enabled = True
    'End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    SiteOpt.Refresh
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
    mFixDisplay
' Dan M 4/21/09
    bmIgnoreChange = True
    mInit
    bmIgnoreChange = False
    If imTerminate Then
        Call cmcCommand_Click(COMMAND_CANCEL)
        Exit Sub
    End If
    If (igWinStatus(SITELIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        igPasswordOk = False
    'ElseIf (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
    'Remove Guide 2/14/03-Jim request
    ElseIf (Trim$(tgUrf(0).sName) = sgCPName) Then
        igPasswordOk = True
    Else
        CSPWord.Show vbModal
    End If

    If Not igPasswordOk Then
        plcGeneral.Enabled = False
        plcSales.Enabled = False
        plcComm.Enabled = False
        plcAgyAdv.Enabled = False
        plcInv.Enabled = False
        plcCntr.Enabled = False
        plcAccount.Enabled = False
        plcBackup.Enabled = False
        plcOptions.Enabled = False
        'plcSports.Enabled = False
        udcSiteTabs.Enabled = False
        plcComments.Enabled = False        '2-11-03
        'plcResearch.Enabled = False
        'imUpdateAllowed = False    'allow only locks to be updated
        ' Dan M  4/21/09
        cmcCommand(COMMAND_UNDO).Enabled = False
        cmcCommand(COMMAND_SAVE).Enabled = False
    Else
        ' Dan M 4/21/09 limit changes
        mTestUnlimitedChanges
        mChangeLabel True
        plcGeneral.Enabled = True
        plcSales.Enabled = True
        plcComm.Enabled = True
        plcAgyAdv.Enabled = True
        plcInv.Enabled = True
        plcCntr.Enabled = True
        plcAccount.Enabled = True
        plcBackup.Enabled = True
        plcOptions.Enabled = True
        udcSiteTabs.Enabled = True
        plcComments.Enabled = True         '2-11-03
        'plcResearch.Enabled = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    ilRet = btrClose(hmCxf)
    btrDestroy hmCxf
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmSaf)
    btrDestroy hmSaf
    ilRet = btrClose(hmSite)
    btrDestroy hmSite
    ilRet = btrClose(hmNrf)
    btrDestroy hmNrf
    'ilRet = btrClose(hmAst)
    'btrDestroy hmAst
    'ilRet = btrClose(hmAtt)
    'btrDestroy hmAtt
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
        
    Set SiteOpt = Nothing   'Remove data segment
    End
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDemoPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Demo list             *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mDemoPop()
    Dim ilRet As Integer
    ilRet = gPopMnfPlusFieldsBox(SiteOpt, cbcReallDemo, tgDemoCode(), sgDemoCodeTag, "D")
    'lbcAud.Height = gListBoxHeight(lbcAud.ListCount, 5)
    cbcReallDemo.AddItem "[Target]", 0
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
'*  ilAffOk                                                                               *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim llDate As Long
    Dim ilValue As Integer

    Screen.MousePointer = vbHourglass
    imLBSCtrls = 1
    imFirstActivate = True
    imTerminate = False
    imPWPrefSpotPct = 0
    imWk1stSolo = 0
    smStdTerms = "15 Days Upon Receipt"
    mParseCmmdLine
    'mInitParameter
    On Error GoTo 0
    hmCxf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCxf, "", sgDBPath & "CXF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CXF.Btr)", SiteOpt
    On Error GoTo 0
    hmMnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "MNF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: MNF.Btr)", SiteOpt
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)

    hmSaf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Saf.Btr)", SiteOpt
    On Error GoTo 0
    imSafRecLen = Len(tmSaf)

    hmSite = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSite, "", sgDBPath & "Site.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Site.Mkd)", SiteOpt
    On Error GoTo 0
    imSiteRecLen = Len(tgSite)

    hmNrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmNrf, "", sgDBPath & "Nrf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Nrf.Btr)", SiteOpt
    On Error GoTo 0
    imNrfRecLen = Len(tgNrf)

    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", SiteOpt
    On Error GoTo 0

    'ilAffOk = True
    'hmAtt = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    'ilRet = btrOpen(hmAtt, "", sgDBPath & "Att.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    ilAffOk = False
    'Else
   '
    '    hmAst = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    '    ilRet = btrOpen(hmAst, "", sgDBPath & "Ast.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    '    If ilRet <> BTRV_ERR_NONE Then
    '        ilAffOk = False
    '    End If
    'End If

'    hmDsf = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dsf.Btr)", SiteOpt
'    On Error GoTo 0
'    imDsfRecLen = Len(tmDsf) 'btrRecordLength(hmDsf)    'Get Cff size
    plcGeneral.BorderStyle = 0
    plcGeneral.Visible = True
    'plcGeneral.Move SiteOpt.Width / 2 - plcGeneral.Width / 2 - 60, 555
    plcGeneral.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcSales.BorderStyle = 0
    plcSales.Visible = False
    'plcSales.Move SiteOpt.Width / 2 - plcSales.Width / 2 - 60, 555
    plcSales.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcComm.BorderStyle = 0
    plcComm.Visible = False
    plcComm.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcSchedule.BorderStyle = 0
    plcSchedule.Visible = False
    plcSchedule.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcAgyAdv.BorderStyle = 0
    plcAgyAdv.Visible = False
    'plcAgyAdv.Move SiteOpt.Width / 2 - plcAgyAdv.Width / 2 - 60, 555
    plcAgyAdv.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcCntr.BorderStyle = 0
    plcCntr.Visible = False
    'plcCntr.Move SiteOpt.Width / 2 - plcCntr.Width / 2 - 60, 555
    plcCntr.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcCopy.BorderStyle = 0
    plcCopy.Visible = False
    'plcCntr.Move SiteOpt.Width / 2 - plcCntr.Width / 2 - 60, 555
    plcCopy.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcLog.BorderStyle = 0
    plcLog.Visible = False
    'plcCntr.Move SiteOpt.Width / 2 - plcCntr.Width / 2 - 60, 555
    plcLog.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcInv.BorderStyle = 0
    plcInv.Visible = False
    'plcInv.Move SiteOpt.Width / 2 - plcInv.Width / 2 - 60, 555
    plcInv.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcAccount.BorderStyle = 0
    plcAccount.Visible = False
    cmcRCorpCal.Visible = False
    'plcAccount.Move SiteOpt.Width / 2 - plcAccount.Width / 2 - 60, 555
    plcAccount.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    cmcRCorpCal.Move plcAccount.Left + 3105, plcAccount.Top + 60
    plcBackup.BorderStyle = 0
    plcBackup.Visible = False
    plcBackup.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcOptions.BorderStyle = 0
    plcOptions.Visible = False
    plcOptions.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    udcSiteTabs.Visible = False
    udcSiteTabs.Move tbcSelection.Left + 60, tbcSelection.Top + 355, tbcSelection.Width - 120, tbcSelection.Height - 510  'plcInv.Height
    plcComments.BorderStyle = 0   '2-11-03
    plcComments.Visible = False
    plcComments.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    'plcResearch.BorderStyle = 0   '2-11-03
    'plcResearch.Visible = False
    'plcResearch.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475

    mInitBox

    SiteOpt.Height = cmcCommand(COMMAND_REPORT).Top + 5 * cmcCommand(COMMAND_REPORT).Height / 3


    gCenterStdAlone SiteOpt
    'SiteOpt.Show

    'imcHelp.Picture = Traffic!imcHelp.Picture
    imFirstFocus = True
    imUpdateAllowed = True
    sgSpfStamp = "~"    'Force read
    gSpfRead
    mDemoPop
    ilRet = mPopPollUserListBox(SiteOpt, cbcUser(0), cbcUser(1))
    mGetLastBkup

    llDate = gGetEarliestTrafSpotDate(hmSdf, -1)
    If llDate <> -1 Then
        edcGRetain(4).Text = Format$(llDate, "m/d/yy")
    End If
    'If ilAffOk Then
    '    llDate = gGetEarliestAffSpotDate(hmAst, hmAtt, -1)
    '    If llDate <> -1 Then
    '        edcGRetain(6).Text = Format$(llDate, "m/d/yy")
    '    End If
    'End If
    ilRet = mSafReadRec()
    ilRet = mSiteReadRec()
    ilRet = mNrfReadRec()
    gUnpackDateLong tmSaf.iEarliestAffSpot(0), tmSaf.iEarliestAffSpot(1), llDate
    If llDate > 0 Then
        If llDate <> gDateValue("1/1/1990") Then
            edcGRetain(6).Text = Format$(llDate, "m/d/yy")
        End If
    End If
    mMoveRecToCtrl
    If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) And (igWinStatus(SITELIST) <> 2) Then
        plcGeneral.Enabled = False
        plcSales.Enabled = False
        plcComm.Enabled = False
        plcAgyAdv.Enabled = False
        plcCntr.Enabled = False
        plcInv.Enabled = False
        plcAccount.Enabled = False
        plcBackup.Enabled = False
        udcSiteTabs.Enabled = False
        plcComments.Enabled = False        '2-11-03
        'plcResearch.Enabled = False
    End If
    lacDBLocation = "[DB-> " & sgDBPath & "]"
    'rbcOption(0).Value = True   'General
    If tbcSelection.SelectedItem.Index <> 1 Then
        ''plcTabSelection.SelectedItem.Index = 1
        'SendKeys "%g", True
    Else
        tbcSelection_Click
    End If
    imPWStatus = 0

    ilValue = Asc(tgSpf.sUsingFeatures7)
    If (ilValue And CSIBACKUP) = CSIBACKUP Then
        ckcCSIBackup.Value = vbChecked
    Else
        ckcCSIBackup.Value = vbUnchecked
    End If
    Call ckcCSIBackup_Click ' Call this to enable or disable the controls on this screen.

    ilValue = Asc(tgSpf.sUsingFeatures2)
    If (((ilValue And SPLITCOPY) = SPLITCOPY) And ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
        lacCopy(3).Caption = "Split Network/Copy Station State Address by"
    ElseIf ((ilValue And SPLITCOPY) = SPLITCOPY) Then
        lacCopy(3).Caption = "Split Copy Station State Address by"
    Else
        lacCopy(3).Caption = "Split Network Station State Address by"
    End If
    '7948
    edcExpt(2).Visible = False
    lacAG(8).Visible = False
    'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
'    '10016
'    For ilRet = 0 To 2
'        rbcInvEmail(ilRet).Top = ckcInv(10).Top
'    Next ilRet
    mInvoiceEmailOptions
    
    'L.Bianchi 06/10/2021 'WO Inovice Export
    'TTP 10291: Traffic Site Options: WO Invoice Export Site Options settings become read-only once they're set
    'If (tgSpfx.iInvExpFeature And 1) > 0 Then
    If (tgSpfx.iInvExpFeature And INVEXP_AUDACYWO) > 0 Then
        edcIE(0).Enabled = True
        edcIE(1).Enabled = True
        edcIE(2).Enabled = True
    Else
        edcIE(0).Enabled = False
        edcIE(1).Enabled = False
        edcIE(2).Enabled = False
    End If
    
    ' JD 11/3/23  ' New option
    If tgSpfx.iLineCostType = 0 Then ' 1=Daily 0=Monthly
        rbcFlatRateAverageFormula(0).Value = False
        rbcFlatRateAverageFormula(1).Value = True
    Else
        rbcFlatRateAverageFormula(0).Value = True
        rbcFlatRateAverageFormula(1).Value = False
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
'*      Procedure Name:mInitParameter                  *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                      and set defaults               *
'*                                                     *
'*******************************************************
Private Sub mInitParameter()
End Sub
Private Sub mkcRCollectPhoneNo_GotFocus()
    mCtrlGotFocusAndIgnoreChange mkcRCollectPhoneNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                      and set defaults               *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
'
'   mMoveCtrlToRec
'   Where:
'
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilValue As Integer
    Dim ilValue2 As Integer
    Dim ilValue3 As Integer
    Dim ilValue10 As Integer
    Dim ilTerms As Integer  '0=Unchanged; 1=Add; 2=Update; 3=Delete
    Dim ilClient As Integer  '0=Unchanged; 1=Add; 2=Update; 3=Delete
    'General

    tgSpf.sGClient = Trim$(edcGClient.Text)
    For ilLoop = LBound(tgSpf.sGAddr) To UBound(tgSpf.sGAddr) Step 1
        tgSpf.sGAddr(ilLoop) = Trim$(edcGAddr(ilLoop))
    Next ilLoop

    ilClient = 0 'Unchanged
    slStr = Trim$(edcGClientAbbr.Text)
    If tgSpf.iMnfClientAbbr <> 0 Then
        tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If (Trim$(tmMnf.sType) = "J") And (Trim$(tmMnf.sUnitType) = "A") Then
                If (Trim$(slStr) <> "") Then
                    ilClient = 2
                Else
                    ilClient = 3
                End If
            Else
                If (slStr <> "") Then
                    ilClient = 1
                Else
                    tgSpf.iMnfClientAbbr = 0
                End If
            End If
        Else
            If (slStr <> "") Then
                ilClient = 1
            Else
                tgSpf.iMnfClientAbbr = 0
            End If
        End If
    Else
        If (slStr <> "") Then
            ilClient = 1 'Add
        End If
    End If
    gGetSyncDateTime smSyncDate, smSyncTime
    If ilClient = 1 Then
        tmMnf.iCode = 0
        tmMnf.sType = "J"
        tmMnf.sName = slStr
        tmMnf.sRPU = ""
        tmMnf.sUnitType = "A"
        tmMnf.iMerge = 0
        tmMnf.iGroupNo = 0
        tmMnf.sCodeStn = ""
        tmMnf.iRemoteID = 0
        tmMnf.iAutoCode = 0
        ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
        Do
            tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmMnf.iAutoCode = tmMnf.iCode
            gPackDate smSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
            gPackTime smSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
            ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        tgSpf.iMnfClientAbbr = tmMnf.iCode
    ElseIf ilClient = 2 Then
        tmMnf.sName = slStr
        ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
        tgSpf.iMnfClientAbbr = tmMnf.iCode
    ElseIf ilClient = 3 Then
        ilRet = btrDelete(hmMnf)
        tgSpf.iMnfClientAbbr = 0
    End If

    If rbcSystemType(1).Value Then
        tgSpf.sSystemType = "R"
    Else
        tgSpf.sSystemType = "N"
    End If
    'tgSpf.iGRetainCntr = Val(edcGRetain(0).Text) 'Lowest invoice #
    gPackDate edcGRetain(0).Text, tmSaf.iLastArchRunDate(0), tmSaf.iLastArchRunDate(1)
    tgSpf.iRetainAffSpot = Val(edcGRetain(1).Text) 'Lowest invoice #
    tgSpf.iRetainTrafSpot = Val(edcGRetain(2).Text) 'Lowest invoice #
    tmSaf.iRetainTrafProj = Val(edcGRetain(3).Text) 'Lowest invoice #
    tgSpf.iRetainTrafProp = Val(edcGRetain(5).Text) 'Lowest invoice #
    tmSaf.iRetainCntr = Val(edcGRetain(7).Text)
    tmSaf.iRetainPayRevHist = Val(edcGRetain(8).Text)
    tmSaf.iNoDaysRetainUAF = Val(edcGRetain(9).Text)
    
    'TTP 10626 JJB 2023-01-10
    tgSpfx.sSageTerm = Trim$(edcSageIE(SAGE_IE_TERMS).Text)
    tgSpfx.sSageAccount = Trim$(edcSageIE(SAGE_IE_ACCOUNT).Text)
    tgSpfx.sSageLocation = Trim$(edcSageIE(SAGE_IE_LOCATIONID).Text)
    tgSpfx.sSageDept = Trim$(edcSageIE(SAGE_IE_DEPTID).Text)
    ''''''''''''''''''''''''''
    
    'SOW Megaphone Phase 1 - New checkbox for daylight savings
    If ckcDaylightSavings.Value = vbChecked Then
        tgSpfx.iIntFeature = 1 'Honor DST
    Else
        tgSpfx.iIntFeature = 0 'Disabled
    End If
        
    ' JD 11/3/23 new option
    If rbcFlatRateAverageFormula(0).Value = True Then
        tgSpfx.iLineCostType = 1
    Else
        tgSpfx.iLineCostType = 0
    End If
    
    If tmSaf.iNoDaysRetainUAF = 0 Then
        tmSaf.iNoDaysRetainUAF = -1
    End If
    If ckcAllowFinalLogDisplay.Value = vbChecked Then
        tmSaf.sFinalLogDisplay = "Y"
    Else
        tmSaf.sFinalLogDisplay = "N"
    End If
    If ckcCntr(1).Value = vbChecked Then
        tmSaf.sProdProtMan = "Y"
    Else
        tmSaf.sProdProtMan = "N"
    End If
    If ckcSales(15).Value = vbChecked Then
        tmSaf.sAvailGreenBar = "Y"
    Else
        tmSaf.sAvailGreenBar = "N"
    End If
    If ckcSortSS.Value = vbChecked Then
        tmSaf.sInvoiceSort = "S"
    Else
        tmSaf.sInvoiceSort = "P"
    End If
    ilValue = 0
    If ckcUsingMatrix(1).Value = vbChecked Then
        ilValue = ilValue Or MATRIXCAL
    End If
    If ckcACodes(3).Value = vbChecked Then
        ilValue = ilValue Or ENGRHIDEMEDIACODE
    End If
    If ckcCntr(5).Value = vbChecked Then
        ilValue = ilValue Or SHOWAUDIOTYPEONBR
    End If
    If ckcSales(16).Value = vbChecked Then
        ilValue = ilValue Or SHOWPRICEONINSERTIONWITHACQUISTION
    End If
    If ckcSalesForce.Value = vbChecked Then
        ilValue = ilValue Or SALESFORCEEXPORT
    End If
    If ckcEfficio.Value = vbChecked Then
        ilValue = ilValue Or EFFICIOEXPORT
    End If
    If ckcJelli.Value = vbChecked Then
        ilValue = ilValue Or JELLIEXPORT
    End If
    If ckcCompensation.Value = vbChecked Then
        ilValue = ilValue Or COMPENSATION
    End If
    tmSaf.sFeatures1 = Chr$(ilValue)
    
    
    ilValue = 0
    If ckcEventRevenue.Value = vbChecked Then
        ilValue = ilValue Or EVENTREVENUE
    End If
    If ckcCntr(6).Value = vbChecked Then
        ilValue = ilValue Or HIDEHIDDENLINES
    End If
    If ckcCntr(8).Value = vbChecked Then
        ilValue = ilValue Or CANCELCLAUSEMANDATORY
    End If
    If ckcOptionFields(14).Value = vbChecked Then
        ilValue = ilValue Or EMAILDISTRIBUTION
    End If
    If ckcRUseTMP(4).Value = vbChecked Then
        ilValue = ilValue Or ACQUISITIONCOMMISSIONABLE
    End If
    If ckcRUseTMP(5).Value = vbChecked Then
        ilValue = ilValue Or PAYMENTONCOLLECTION
    End If
    If ckcUsingMatrix(2).Value = vbChecked Then
        ilValue = ilValue Or TABLEAUEXPORT
    End If
    If ckcUsingMatrix(3).Value = vbChecked Then
        ilValue = ilValue Or TABLEAUCAL
    End If
    tmSaf.sFeatures2 = Chr$(ilValue)
    
    ilValue = 0
    If ckcSales(17).Value = vbChecked Then
        ilValue = ilValue Or SUPPRESSNETCOMM
    End If
    If ckcInv(8).Value = vbChecked Then
        ilValue = ilValue Or REQSTATIONPOSTING
    End If
    If rbcSplitCopyState(1).Value = True Then
        ilValue = ilValue Or SPLITCOPYLICENSE
    End If
    If rbcSplitCopyState(2).Value = True Then
        ilValue = ilValue Or SPLITCOPYPHYSICAL
    End If
    If ckcCntr(7).Value = vbChecked Then
        ilValue = ilValue Or FREEZEDEFAULT
    End If
    If ckcInv(9).Value = vbChecked Then
        ilValue = ilValue Or INVEMAILINDEX
    End If
    If ckcInv(10).Value = vbChecked Then
        ilValue = ilValue Or INVSENDEMAILINDEX
    End If
    If ckcACodes(4).Value = vbChecked Then  'Suppress Zero Dollars from Invoice Export
        ilValue = ilValue Or SUPPRESSZERODOLLARINVEXPT
    End If
    tmSaf.sFeatures3 = Chr$(ilValue)
    
    ilValue = 0
    If ckcOptionFields(15).Value = vbChecked Then
        ilValue = ilValue Or FILEMAKERIMPORT
    End If
    If ckcOptionFields(16).Value = vbChecked Then
        ilValue = ilValue Or ACT1CODES
    End If
    If ckcCntr(9).Value = vbChecked Then
        ilValue = ilValue Or MKTNAMEONBR
    End If
    If ckcRUseTMP(6).Value = vbChecked Then
        ilValue = ilValue Or COMPRESSTRANSACTIONS
    End If
    If ckcCAvails(1).Value = vbChecked Then
        ilValue = ilValue Or AVAILINCLUDERESERVATION
    End If
    If ckcCAvails(2).Value = vbChecked Then
        ilValue = ilValue Or AVAILINCLUDEREMNANT
    End If
    If ckcCAvails(3).Value = vbChecked Then
        ilValue = ilValue Or AVAILINCLDEDIRECTRESPONSES
    End If
    If ckcCAvails(4).Value = vbChecked Then
        ilValue = ilValue Or AVAILINCLUDEPERINQUIRY
    End If
    tmSaf.sFeatures4 = Chr$(ilValue)
    
    ilValue = 0
    If ckcOptionFields(17).Value = vbChecked Then
        ilValue = ilValue Or PROGRAMMATICALLOWED
    End If
    If ckcOptionFields(18).Value = vbChecked Then
        ilValue = ilValue Or SHOWAVAILCOUNT
    End If
    If ckcOptionFields(19).Value = vbChecked Then
        ilValue = ilValue Or SHOWCPPTAB
    End If
    If ckcOptionFields(20).Value = vbChecked Then
        ilValue = ilValue Or SHOWCPMTAB
    End If
    If ckcOptionFields(21).Value = vbChecked Then
        ilValue = ilValue Or SHOWPRICETAB
    End If
    If ckcCntr(10).Value = vbChecked Then
        ilValue = ilValue Or PODCASTAUDPCT
    End If
    If ckcCntr(11).Value = vbChecked Then
        ilValue = ilValue Or SHOWDAYDROPDOWN
    End If
    If ckcOptionFields(25).Value = vbChecked Then
        ilValue = ilValue Or CSVAFFIDAVITIMPORT
    End If
    tmSaf.sFeatures5 = Chr$(ilValue)
    
    ilValue = 0
    If ckcSales(18).Value = vbChecked Then
        ilValue = ilValue Or BILLINGONINSERTIONS
    End If
    '9114
    If ckcXDSBy(ASTISCI).Value = vbChecked Then 'send hb/hbp with astcodes
        ilValue = ilValue Or UNITIDBYASTCODEFORISCI
    End If
    If ckcCntr(12).Value = vbChecked Then 'Signature on Proposals
        ilValue = ilValue Or SIGNATUREONPROPOSAL
    End If
    If ckcCntr(13).Value = vbChecked Then 'Calcultae research totals on line change
        ilValue = ilValue Or CALCULATESEARCHONLINECHG
    End If
    If ckcAEDI(2).Value = vbChecked Then 'Using EDI Client and Product codes
        ilValue = ilValue Or EDIAGYCODES
    End If
    'If ckcOptionFields(22).Value = vbChecked Then 'Adavance Avails (Avails/Protection/Research Tab)
    '    ilValue = ilValue Or ADVANCEAVAILS
    'End If
    If ckcOptionFields(23).Value = vbChecked Then 'RAB-Calendar
        ilValue = ilValue Or RABCALENDAR
    End If
    If ckcOptionFields(24).Value = vbChecked Then 'RAB-Calendar
        ilValue = ilValue Or OVERDUEEXPORT
    End If
    tmSaf.sFeatures6 = Chr$(ilValue)
    
    ilValue = 0
    If ckcOptionFields(26).Value = vbChecked Then 'RAB-Standard
        ilValue = ilValue Or RABSTD
    End If
    If ckcOptionFields(27).Value = vbChecked Then 'RAB-Calendar Spots
        ilValue = ilValue Or RABCALSPOTS
    End If
    If udcSiteTabs.Automation(29) = vbChecked Then   '11/5/20 - TTP # 10013 - iMediaTouch Replace COM with Media Code
        ilValue = ilValue Or IMEDIA_MEDIACODE
    End If
    '10016
    If rbcInvEmail(AIRONLY).Value = True Then
        ilValue = ilValue Or INVEMAILAIRONLY
    ElseIf rbcInvEmail(NTRONLY).Value = True Then
        ilValue = ilValue Or INVEMAILNTRONLY
    End If
    If ckcOptionFields(31) = vbChecked Then   'TTP # 9992 - Custom Rev Export
        ilValue = ilValue Or CUSTOMEXPORT
    End If
    If ckcInv(11) = vbChecked Then   'Bill Over-Delivered CPM Impressions
        ilValue = ilValue Or PODBILLOVERDELIVERED
    End If
    tmSaf.sFeatures7 = Chr$(ilValue)
    '10048
    ilValue = 0
    If ckcOptionFields(PODAIRTIMECKC).Value = vbChecked Then
        ilValue = ilValue Or PODAIRTIME
    End If
    If ckcOptionFields(PODSPOTSCKC).Value = vbChecked Then
        ilValue = ilValue Or PODSPOTS
    End If
    If ckcOptionFields(ADSERVERCKC).Value = vbChecked Then
        ilValue = ilValue Or PODADSERVER
    End If
    If ckcOptionFields(PODMIXCKC).Value = vbChecked Then
        ilValue = ilValue Or PODADSERVERVIEWONLY
    End If
    If ckcPodShowWk.Value = vbChecked Then
        ilValue = ilValue Or PODSHOWWKOF
    End If
    tmSaf.sFeatures8 = Chr$(ilValue)
    
    If ckcOptionFields(22).Value = vbChecked Then 'Adavance Avails (Avails/Protection/Research Tab)
        tmSaf.sAdvanceAvail = "Y"
    Else
        tmSaf.sAdvanceAvail = "N"
    End If
        
    tmSaf.sExcludeAudioTypeR = "N"
    tmSaf.sExcludeAudioTypeL = "N"
    tmSaf.sExcludeAudioTypeM = "N"
    tmSaf.sExcludeAudioTypeS = "N"
    tmSaf.sExcludeAudioTypeP = "N"
    tmSaf.sExcludeAudioTypeQ = "N"
    'If ckcCopy(5).Value = vbChecked Then
    '    tmSaf.sExcludeAudioTypeR = "Y"
    'End If
    If (ckcUsingLiveCopy.Value = vbChecked) Or (ckcCopy(1).Value = vbChecked) Then
        If ckcCopy(6).Value = vbChecked Then
            tmSaf.sExcludeAudioTypeL = "Y"
        End If
        If ckcCopy(7).Value = vbChecked Then
            tmSaf.sExcludeAudioTypeM = "Y"
        End If
        If ckcCopy(8).Value = vbChecked Then
            tmSaf.sExcludeAudioTypeS = "Y"
        End If
        If ckcCopy(9).Value = vbChecked Then
            tmSaf.sExcludeAudioTypeP = "Y"
        End If
        If ckcCopy(10).Value = vbChecked Then
            tmSaf.sExcludeAudioTypeQ = "Y"
        End If
    End If
    '6/2/14:
    tmSaf.sEventSubtotal1 = ""
    tmSaf.sEventSubtotal2 = ""
    tmSaf.sEventSubtotal1 = ""
    tmSaf.sEventSubtotal2 = ""
    
    If ckcInvoiceExport.Value = vbChecked Then
        tmSaf.sInvExpDelimiter = edcExpt(3).Text
    Else
        tmSaf.sInvExpDelimiter = ","
    End If
    
    If rbcGTBar(0).Value Then
        tgSpf.sGTBar = "C"
    Else
        tgSpf.sGTBar = "U"
    End If
    tgSpf.iGNoDaysPass = Val(edcGRetainPassword.Text)
    tgSpf.iGAlertInterval = Val(edcGAlertInterval.Text)
    If ckcSSelling.Value = vbChecked Then 'Selling vehicle
        tgSpf.sSSellNet = "Y"
    Else
        tgSpf.sSSellNet = "N"
    End If
    If ckcSDelivery.Value = vbChecked Then  'Delivery vehicle
        tgSpf.sSDelNet = "Y"
    Else
        tgSpf.sSDelNet = "N"
    End If
    'If ckcExport.Value = vbChecked Then  'Export Menu
    '    tgSpf.sExport = "Y"
    'Else
    '    tgSpf.sExport = "N"
    'End If
    'If ckcImport.Value = vbChecked Then  'Import Menu
    '    tgSpf.sImport = "Y"
    'Else
    '    tgSpf.sImport = "N"
    'End If
    ilValue = 0
    ilValue2 = 0
    ilValue3 = 0
    ilValue10 = 0
    '8-10-05 Types of automation equipment are stored in 2 1-byte character fields.
    'AutoType (Byte 1)Bits Right to Left:1=Dalet; 2=Prophet NexGen; 3=Scott; 4=Drake; 5=RCS; 6=Prophet Wizard, 7 = iMediaTouch
    'Autotype2 (Byte 2): bits from rt to lt: 1 = audio vault sat, 2 = audio vault air (unused for now), 3 = wireready, 4 = enco, 5 = R-N Rep, 6= R-N Net
    If udcSiteTabs.Automation(0) = vbChecked Then
        'tgSpf.sAutoType = "1"
        ilValue = DALET
    End If
    If udcSiteTabs.Automation(1) = vbChecked Then
        'tgSpf.sAutoType = "4"
        ilValue = ilValue Or DRAKE
    End If
    If udcSiteTabs.Automation(2) = vbChecked Then
        'tgSpf.sAutoType = "2"
        ilValue = ilValue Or PROPHETNEXGEN
    End If
    If udcSiteTabs.Automation(3) = vbChecked Then
        'tgSpf.sAutoType = "5"
        ilValue = ilValue Or RCS4DIGITCART
    End If
    If udcSiteTabs.Automation(4) = vbChecked Then
        'tgSpf.sAutoType = "3"
        ilValue = ilValue Or SCOTT
    End If
    If udcSiteTabs.Automation(5) = vbChecked Then
        'tgSpf.sAutoType = "3"
        ilValue = ilValue Or PROPHETWIZARD
    End If
    If udcSiteTabs.Automation(6) = vbChecked Then
        'tgSpf.sAutoType = "3"
        ilValue = ilValue Or PROPHETMEDIASTAR
    End If
    If udcSiteTabs.Automation(7) = vbChecked Then       '6-25-05 iMediaTouch
        'tgSpf.sAutoType = "3"
        ilValue = ilValue Or IMEDIATOUCH
    End If
    If udcSiteTabs.Automation(8) = vbChecked Then       '8-10-05 Audio Vault Sat
        'tgSpf.sAutoType = "3"
        ilValue2 = AUDIOVAULT
    End If
    If udcSiteTabs.Automation(9) = vbChecked Then       '6-25-05 iMediaTouch
        'tgSpf.sAutoType = "3"
        ilValue2 = ilValue2 Or WIREREADY
    End If
    If udcSiteTabs.Automation(10) = vbChecked Then       '9-12-06 Enco
        ilValue2 = ilValue2 Or ENCO
    End If
    If ckcRN_Rep.Value = vbChecked Then
        ilValue2 = ilValue2 Or RN_REP
    End If
    If ckcRN_Net.Value = vbChecked Then
        ilValue2 = ilValue2 Or RN_NET
    End If

    If udcSiteTabs.Automation(12) = vbChecked Then       '8-22-08 Simian
        ilValue2 = ilValue2 Or SIMIAN
    End If

    If udcSiteTabs.Automation(16) = vbChecked Then       '8-22-08 Simian
        ilValue2 = ilValue2 Or RCS5DIGITCART
    End If
    If udcSiteTabs.Automation(20) = vbChecked Then  '11-15-10 audio vault rps
        ilValue2 = ilValue2 Or AUDIOVAULTRPS
    End If

    If udcSiteTabs.Automation(21) = vbChecked Then  '11-15-10 audio vault rps
        ilValue3 = ilValue3 Or AUDIOVAULTAIR
    End If
    If udcSiteTabs.Automation(22) = vbChecked Then  '1/3/12: Wide Orbit
        ilValue3 = ilValue3 Or WIDEORBIT
    End If
    If udcSiteTabs.Automation(23) = vbChecked Then  '5/10/12: Jelli
        ilValue3 = ilValue3 Or JELLI
    End If
    If udcSiteTabs.Automation(24) = vbChecked Then  '5/10/12: Jelli
        ilValue3 = ilValue3 Or ENCOESPN
    End If
    If udcSiteTabs.Automation(25) = vbChecked Then  '8-16-13 Scott V5
        ilValue3 = ilValue3 Or SCOTT_V5
    End If
    
    If udcSiteTabs.Automation(26) = vbChecked Then  '1-5-16 Zetta
        ilValue3 = ilValue3 Or ZETTA
    End If

    If udcSiteTabs.Automation(27) = vbChecked Then  '5/10/18: Station Playlist
        ilValue3 = ilValue3 Or STATIONPLAYLIST
    End If
    If udcSiteTabs.Automation(28) = vbChecked Then  '1-5-16 Zetta
        ilValue3 = ilValue3 Or RADIOMAN
    End If


    tgSpf.sAutoType = Chr$(ilValue)
    tgSpf.sAutoType2 = Chr$(ilValue2)
    tgSpf.sAutoType3 = Chr$(ilValue3)
    'Sales
    'If rbcSMove(1).Value Then 'Using Proposal Contracts
    '    tgSpf.sSMove = "N"
    'Else
    '    tgSpf.sSMove = "Y"
    'End If
    If rbcSUseProd(1).Value Then 'Use Prod/Advt,Prod on spot screen
        tgSpf.sUseProdSptScr = "A"
    Else
        tgSpf.sUseProdSptScr = "P"
    End If
    If ckcSales(2).Value = vbChecked Then
        tgSpf.sHideGhostSptScr = "Y"
    Else
        tgSpf.sHideGhostSptScr = "N"
    End If
    tgSpf.iRptDollarMag = Val(edcSales(0).Text)
    If rbcUnitOr3060(0).Value Then 'For Billed & Booked, use E = Entered Date, A = AGed Date
        tgSpf.sUnitOr3060 = "U"
    Else
        tgSpf.sUnitOr3060 = "3"
    End If
    If rbcEqualize(0).Value Then 'Combo Avail Equalize by 30, 60 or None
        tgSpf.sAvailEqualize = "3"
    ElseIf rbcEqualize(1).Value Then
        tgSpf.sAvailEqualize = "6"
    Else
        tgSpf.sAvailEqualize = "N"
    End If
    If rbcSEnterAge(1).Value Then 'For Billed & Booked, use E = Entered Date, A = AGed Date
        tgSpf.sSEnterAgeDate = "A"
    Else
        tgSpf.sSEnterAgeDate = "E"
    End If
    If ckcSales(14).Value = vbChecked Then  'Remote Users
        tgSpf.sRemoteUsers = "Y"
    Else
        tgSpf.sRemoteUsers = "N"
    End If
    If udcSiteTabs.Research(2) Then 'Audience data Magnitude, in T = Thousands (000), H = Hundreds (00), N = Tens (0), U = Units()
        tgSpf.sSAudData = "H"
    ElseIf udcSiteTabs.Research(3) Then  'Audience data Magnitude, in T = Thousands (000), H = Hundreds (00), N = Tens (0), U = Units()
        tgSpf.sSAudData = "N"
    ElseIf udcSiteTabs.Research(4) Then  'Audience data Magnitude, in T = Thousands (000), H = Hundreds (00), N = Tens (0), U = Units()
        tgSpf.sSAudData = "U"
    Else
        tgSpf.sSAudData = "T"
    End If
    If udcSiteTabs.Research(11) Then  'GRP/CPP Calculation, by R = Rating, G = GRP 2 places, A = Audience
        tgSpf.sSGRPCPPCal = "R"
    ElseIf udcSiteTabs.Research(12) Then
        tgSpf.sSGRPCPPCal = "G"
    Else
        tgSpf.sSGRPCPPCal = "A"
    End If
    If ckcSales(0).Value = vbUnchecked Then 'Allow sub-company commissions
        tgSpf.sAllowMGs = "N"
    Else
        tgSpf.sAllowMGs = "Y"
    End If
    If ckcSMktBase.Value = vbChecked Then  'Market Base
        tgSpf.sMktBase = "Y"
    Else
        tgSpf.sMktBase = "N"
    End If
    If ckcSales(9).Value = vbChecked Then  'Allow Daily Buys
        tgSpf.sAllowDailyBuys = "Y"
    Else
        tgSpf.sAllowDailyBuys = "N"
    End If
    If rbcPLMove(0).Value Then 'Post Log Moves, by M = MG's, O = Outside, A = Ask
        tgSpf.sPLMove = "M"
    ElseIf rbcPLMove(1).Value Then
        tgSpf.sPLMove = "O"
    Else
        tgSpf.sPLMove = "A"
    End If
    If ckcSales(10).Value = vbChecked Then
        tgSpf.sBBsToAff = "Y"
    Else
        tgSpf.sBBsToAff = "N"
    End If
    tmSaf.iRptLenDefault(0) = Val(edcSales(2).Text)
    tmSaf.iRptLenDefault(1) = Val(edcSales(3).Text)
    tmSaf.iRptLenDefault(2) = Val(edcSales(4).Text)
    tmSaf.iRptLenDefault(3) = Val(edcSales(5).Text)
    tmSaf.iRptLenDefault(4) = Val(edcSales(6).Text)
    If ckcSales(11).Value = vbChecked Then  'Reschedule across calendar month
        tmSaf.sReSchdXCal = "N"
    Else
        tmSaf.sReSchdXCal = "Y"
    End If
    '2-7-03
    If rbcInsertAddr(1).Value Then  'get Insertion Address from P = payee (contract), S = Site (invoice tab), or Vehicle
        tgSpf.sInsertAddr = "S"
    ElseIf rbcInsertAddr(2).Value Then
        tgSpf.sInsertAddr = "V"
    Else
        tgSpf.sInsertAddr = "P"
    End If
    If ckcSales(1).Value = vbChecked Then  'Retain Spot Screen Date
        tgSpf.sSSRetainDate = "Y"
    Else
        tgSpf.sSSRetainDate = "N"
    End If
    If ckcSales(4).Value = vbChecked Then  'BR by Standard Quarter
        tgSpf.sSBrStdQt = "Y"
    Else
        tgSpf.sSBrStdQt = "N"
    End If
    'If ckcSSubCompany.Value = vbChecked Then  'Allow sub-company commissions
    If rbcCommBy(4).Value Then
        tgSpf.sSubCompany = "Y"
    Else
        tgSpf.sSubCompany = "N"
    End If
    'If ckcSCommByCntr.Value = vbChecked Then  'Salesperson commission by Contract
    If rbcCommBy(3).Value Then
        tgSpf.sCommByCntr = "Y"
    Else
        tgSpf.sCommByCntr = "N"
    End If
    If Val(edcSales(1).Text) > 40 Then
        tgSpf.iVehLen = 20
    Else
        tgSpf.iVehLen = Val(edcSales(1).Text)
    End If
    If rbcSInvCntr(0).Value Then        'Show Ordered, Update Ordered
        tgSpf.sInvAirOrder = "S"
    ElseIf rbcSInvCntr(1).Value Then    'Show Ordered, Update Aired
        tgSpf.sInvAirOrder = "O"
    ElseIf rbcSInvCntr(3).Value Then    'Show Aired minus Missed, Update Ordered
        tgSpf.sInvAirOrder = "2"
    Else
        tgSpf.sInvAirOrder = "A"        'As Aired
    End If
    If ckcSales(5).Value = vbChecked Then
        tgSpf.sCPkOrdered = "Y"
    Else
        tgSpf.sCPkOrdered = "N"
    End If
    'If (tgSpf.sInvAirOrder = "O") Or (tgSpf.sInvAirOrder = "S") Then
    '    tgSpf.sCPkAired = "N"
    'Else
        If ckcSales(6).Value = vbChecked Then
            tgSpf.sCPkAired = "Y"
        Else
            tgSpf.sCPkAired = "N"
        End If
    'End If
    If ckcSales(7).Value = vbChecked Then
        tgSpf.sCPkEqual = "Y"
    Else
        tgSpf.sCPkEqual = "N"
    End If
    'If (tgSpf.sCPkOrdered = "Y") Then
    '    If rbcBPkageGenMeth(0).Value Then
    '        tgSpf.iPkageGenMeth = 0
    '    Else
    '        tgSpf.iPkageGenMeth = 1
    '    End If
    'Else
        tgSpf.iPkageGenMeth = 0
    'End If
    'tgSpf.iSDay = Val(edcSHoldDays.Text)
    'tgSpf.sSchRNG = UCase$(edcSchRNG.Text)
    'tgSpf.sSchMal = UCase$(edcSchMal.Text)
    'tgSpf.sSchMdl = UCase$(edcSchMdl.Text)
    'tgSpf.sSchMil = UCase$(edcSchMil.Text)
    'tgSpf.sSchCycle = UCase$(edcSchCycle.Text)
    'tgSpf.sSchMove = UCase$(edcSchMove.Text)
    'tgSpf.sSchCompact = UCase$(edcSchCompact.Text)
    'tgSpf.sSchPreempt = UCase$(edcSchPreempt.Text)
    'tgSpf.sSchHour = UCase$(edcSchHour.Text)
    'tgSpf.sSchDay = UCase$(edcSchDay.Text)
    'tgSpf.sSchQH = UCase$(edcSchQH.Text)
    'tgSpf.sLkCredit = edcGLock(0).Text
    'tgSpf.sLkLog = edcGLock(1).Text
    'Agency/Advertiser
    'tgSpf.iAProd = Val(edcAProdNameSize.Text)
    If ckcOptionFields(9).Value = vbChecked Then  'Program exclusions
        tgSpf.sAExcl = "Y"
    Else
        tgSpf.sAExcl = "N"
    End If
    tgSpf.iATargets = Val(edcExpt(0).Text)
    If rbcAISCI(0).Value Then 'ISCI
        tgSpf.sAISCI = "A"
    ElseIf rbcAISCI(1).Value Then 'ISCI
        tgSpf.sAISCI = "X"
    ElseIf rbcAISCI(2).Value Then 'ISCI
        tgSpf.sAISCI = "Y"
    Else
        tgSpf.sAISCI = "N"
    End If
    If ckcAEDI(0).Value = vbChecked Then  'EDI for Contracts
        tgSpf.sAEDIC = "Y"
    Else
        tgSpf.sAEDIC = "N"
    End If
    If ckcAEDI(1).Value = vbChecked Then  'EDI for Invoices
        tgSpf.sAEDII = "Y"
    Else
        tgSpf.sAEDII = "N"
    End If
    If rbcAPrtStyle(0).Value Then 'Contract print style
        tgSpf.sAPrtStyle = "W"
    ElseIf rbcAPrtStyle(1).Value Then 'Contract print style
        tgSpf.sAPrtStyle = "N"
    Else
        tgSpf.sAPrtStyle = "A"
    End If
    If ckcACodes(0).Value = vbChecked Then  'Rep codes
        tgSpf.sARepCodes = "Y"
    Else
        tgSpf.sARepCodes = "N"
    End If
    If ckcACodes(1).Value = vbChecked Then  'Station codes
        tgSpf.sAStnCodes = "Y"
    Else
        tgSpf.sAStnCodes = "N"
    End If
    If ckcACodes(2).Value = vbChecked Then  'Agency codes
        tgSpf.sAAgyCodes = "Y"
    Else
        tgSpf.sAAgyCodes = "N"
    End If

    ilTerms = 0 'Unchanged
    slStr = Trim$(edcTerms.Text)
    If tgSpf.iMnfInvTerms <> 0 Then
        tmMnfSrchKey.iCode = tgSpf.iMnfInvTerms
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If (Trim$(tmMnf.sType) = "J") And (Trim$(tmMnf.sUnitType) = "D") Then
                If StrComp(sgDefaultTerms, slStr, vbTextCompare) <> 0 Then
                    If (Trim$(slStr) <> "") And (StrComp(slStr, smStdTerms, vbTextCompare) <> 0) Then
                        ilTerms = 2
                    Else
                        ilTerms = 3
                    End If
                End If
            Else
                If (slStr <> "") And (StrComp(sgDefaultTerms, slStr, vbTextCompare) <> 0) Then
                    ilTerms = 1
                Else
                    tgSpf.iMnfInvTerms = 0
                End If
            End If
        Else
            If (slStr <> "") And (StrComp(sgDefaultTerms, slStr, vbTextCompare) <> 0) Then
                ilTerms = 1
            Else
                tgSpf.iMnfInvTerms = 0
            End If
        End If
    Else
        If (slStr <> "") And (StrComp(sgDefaultTerms, slStr, vbTextCompare) <> 0) Then
            ilTerms = 1 'Add
        End If
    End If
    gGetSyncDateTime smSyncDate, smSyncTime
    If ilTerms = 1 Then
        tmMnf.iCode = 0
        tmMnf.sType = "J"
        tmMnf.sName = slStr
        tmMnf.sRPU = ""
        tmMnf.sUnitType = "D"
        tmMnf.iMerge = 0
        tmMnf.iGroupNo = 0
        tmMnf.sCodeStn = ""
        tmMnf.iRemoteID = 0
        tmMnf.iAutoCode = 0
        ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
        Do
            tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmMnf.iAutoCode = tmMnf.iCode
            gPackDate smSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
            gPackTime smSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
            ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        sgDefaultTerms = slStr
        tgSpf.iMnfInvTerms = tmMnf.iCode
    ElseIf ilTerms = 2 Then
        tmMnf.sName = slStr
        ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
        sgDefaultTerms = slStr
        tgSpf.iMnfInvTerms = tmMnf.iCode
    ElseIf ilTerms = 3 Then
        ilRet = btrDelete(hmMnf)
        sgDefaultTerms = smStdTerms
        tgSpf.iMnfInvTerms = 0
    End If
    'Penny
    slStr = edcAPenny.Text
    'gStrToPDN slStr, 2, 3, tgSpf.sRPenny
    tgSpf.lRPenny = gStrDecToLong(slStr, 2)
    'Contract/Copy
    tgSpf.lCLowestNo = Val(edcCNo(0).Text) 'Lowest contract #
    tgSpf.lCHighestNo = Val(edcCNo(1).Text)   'Highest contract #- Not using
    tgSpf.lCNextNo = Val(edcCNo(2).Text) 'Next number
    If ckcOptionFields(8).Value = vbChecked Then  'Using Estimate number
        tgSpf.sCEstNo = "Y"
    Else
        tgSpf.sCEstNo = "N"
    End If
    'If rbcCRot(0).Value Then 'Copy rotation
    '    tgSpf.sCRot = "I"
    'ElseIf rbcCRot(1).Value Then
    '    tgSpf.sCRot = "M"
    'Else
    '    tgSpf.sCRot = "V"
    'End If
    If ckcAllowPrelLog.Value = vbChecked Then  'Using Blackouts on Logs
        tgSpf.sAllowPrelLog = "Y"
    Else
        tgSpf.sAllowPrelLog = "N"
    End If
    If rbcVirtPkg(0).Value Then 'Default Log Copy (On=Y/Off=N)
        tgSpf.sVirtPkgCompute = "H"
    Else
        tgSpf.sVirtPkgCompute = "P"
    End If
    If rbcCSchdRemnant(1).Value Then 'Schedule Remnant Contracts
        tgSpf.sSchdRemnant = "Y"
    Else
        tgSpf.sSchdRemnant = "N"
    End If
    If rbcCSchdPromo(1).Value Then 'Schedule Remnant Contracts
        tgSpf.sSchdPromo = "Y"
    Else
        tgSpf.sSchdPromo = "N"
    End If
    If rbcCSchdPSA(1).Value Then 'Schedule Remnant Contracts
        tgSpf.sSchdPSA = "Y"
    Else
        tgSpf.sSchdPSA = "N"
    End If
'    If ckcCUseCartNo.Value = vbChecked Then  'Using Cart number
    If ckcCBump.Value = vbChecked Then  'Bump spots in past
        tgSpf.sCBumpPast = "Y"
    Else
        tgSpf.sCBumpPast = "N"
    End If
    'If rbcSRes(1).Value Then 'Using Reservation Orders
    '    tgSpf.sSUseResv = "N"
    'Else
    '    tgSpf.sSUseResv = "Y"
    'End If
    'If rbcSRem(1).Value Then 'Using Proposal Orders
    '    tgSpf.sSUseRem = "N"
    'Else
    '    tgSpf.sSUseRem = "Y"
    'End If
    'If rbcSDR(1).Value Then 'Using Direct Response Orders
    '    tgSpf.sSUseDR = "N"
    'Else
    '    tgSpf.sSUseDR = "Y"
    'End If
    'If rbcSPI(1).Value Then 'Using Per Inquiry Orders
    '    tgSpf.sSUsePI = "N"
    'Else
    '    tgSpf.sSUsePI = "Y"
    'End If
    'If rbcSPSA(1).Value Then 'Using PSA Orders
    '    tgSpf.sSUsePSA = "N"
    'Else
    '    tgSpf.sSUsePSA = "Y"
    'End If
    'If rbcSPromo(1).Value Then 'Using Promo Orders
    '    tgSpf.sSUsePromo = "N"
    'Else
    '    tgSpf.sSUsePromo = "Y"
    'End If
    If cbcReallDemo.ListIndex <= 0 Then
        tgSpf.iReallMnfDemo = 0
    Else
        slNameCode = tgDemoCode(cbcReallDemo.ListIndex - 1).sKey 'Traffic!lbcDemoCode.List(lbcDemo(ilLoop).ListIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        tgSpf.iReallMnfDemo = CInt(slCode)
    End If
    tgSpf.lDiscCurrCntrNo = Val(edcDiscNo.Text) 'Current Discrepancy Contract #
    If ckcCAvails(0).Value = vbChecked Then  'Include Missed in Demo Bar Avails
        tgSpf.sCIncludeMissDB = "Y"
    Else
        tgSpf.sCIncludeMissDB = "N"
    End If
    If ckcCWarnMsg.Value = vbChecked Then  'Contract Warning Messages
        tgSpf.sCWarnMsg = "Y"
    Else
        tgSpf.sCWarnMsg = "N"
    End If
    
    'TTP 11033 JJB 2024-05-11
    If ckcDisallowAuthorScheduling.Value = vbChecked Then  ''Schedule screen: disallow allow user who last edited a contract from being able to schedule it
        tgSpfx.iSchdFeature = SCHD_DISALLOWSAMEUSER
    Else
        tgSpfx.iSchdFeature = 0
    End If
    
    If ckcCAudPkg.Value = vbChecked Then  'Allow audience packages
        tgSpf.sCAudPkg = "Y"
    Else
        tgSpf.sCAudPkg = "N"
    End If
    '10842 no longer used
'    If ckcCLnStdQt.Value = vbChecked Then  'BR by Standard Quarter
'        tgSpf.sCLnStdQt = "Y"
'    Else
'        tgSpf.sCLnStdQt = "N"
'    End If
    'Options
    'If imPWStatus = 1 Then
        If ckcGUsePropSys.Value = vbChecked Then  'Using Proposal System
            tgSpf.sGUsePropSys = "Y"
        Else
            tgSpf.sGUsePropSys = "N"
        End If
        If ckcGUseAffSys(0).Value = vbChecked Then  'Using Affiliate System
            tgSpf.sGUseAffSys = "Y"
        Else
            tgSpf.sGUseAffSys = "N"
        End If
        If tgSpf.sSystemType <> "R" Then
            If ckcGUseAffFeed.Value = vbChecked Then  'Using Affiliate Feed (Regional Copy)
                tgSpf.sGUseAffFeed = "Y"
            Else
                tgSpf.sGUseAffFeed = "N"
            End If
        Else
            tgSpf.sGUseAffFeed = "N"
        End If
    'End If
    If ckcCUseSegments.Value = vbChecked Then  'Using Segments
        tgSpf.sCUseSegments = "Y"
    Else
        tgSpf.sCUseSegments = "N"
    End If
    If ckcUsingBBs.Value = vbChecked Then  'Using Billboards
        tgSpf.sUsingBBs = "Y"
    Else
        tgSpf.sUsingBBs = "N"
    End If
    ilValue = 0
    If ckcOptionFields(0).Value = vbChecked Then  'Option Fields (Right to Left):0=Projection;1=Bus Cat; 2=Share; 3=Rev Set; 4=Guar; 5=Billing Cycle; 6=Co-op; 7=Research
        ilValue = OFPROJECTION  '&H1
    End If
    If ckcOptionFields(1).Value = vbChecked Then
        ilValue = ilValue Or OFBUSCAT   '&H2
    End If
    If ckcOptionFields(2).Value = vbChecked Then
        ilValue = ilValue Or OFSHARE    '&H4
    End If
    If ckcOptionFields(3).Value = vbChecked Then
        ilValue = ilValue Or OFREVSET   '&H8
    End If
    If ckcOptionFields(4).Value = vbChecked Then
        ilValue = ilValue Or OFDELGUAR  '&H10
    End If
    If ckcOptionFields(5).Value = vbChecked Then
        ilValue = ilValue Or OFCALENDARBILL    '&H20
    End If
    '3/22/13: Feature has been hidden
    'If ckcOptionFields(6).Value = vbChecked Then
    '    ilValue = ilValue Or OFCOOPBILL '&H40
    'End If
    If ckcOptionFields(7).Value = vbChecked Then
        ilValue = ilValue Or OFRESEARCH '&H80
    End If

    tgSpf.sOptionFields = Chr$(ilValue)

    ilValue = 0
    If ckcOverrideOptions(0).Value = vbChecked Then  'Option Fields (Right to Left):0=Allocation %;1=Acquisition Cost; 2=1st Position; 3=Preferred Days/Times; 4=Solo Avails
        ilValue = SPALLOCATION
    End If
    '6/7/15: renamed acquisition to ntr acquisition and replace acquisition in site override with Barter in system options
    If ckcOverrideOptions(1).Value = vbChecked Then
        ilValue = ilValue Or SPNTRACQUISITION
    End If
    If ckcOverrideOptions(2).Value = vbChecked Then
        ilValue = ilValue Or SP1STPOSITION
    End If
    If ckcOverrideOptions(3).Value = vbChecked Then
        ilValue = ilValue Or SPPREFERREDDT
    End If
    If ckcOverrideOptions(4).Value = vbChecked Then
        ilValue = ilValue Or SPSOLOAVAIL
    End If
    If udcSiteTabs.Research(23) Then  '(0) and (1) in sUsingFeatures
        ilValue = ilValue Or BESTFITWEIGHTNONE
    End If
    If ckcInv(0).Value = vbChecked Then
        ilValue = ilValue Or BBSAMELINE
    End If
    tgSpf.sOverrideOptions = Chr$(ilValue)

    ilValue = 0
    '12/9/14: Removed Regional Copy as not used any longer
    'If ckcRegionalCopy.Value = vbChecked Then  'Using Split Networks
    '    ilValue = ilValue Or REGIONALCOPY
    'End If
    If ckcUsingSplitCopy.Value = vbChecked Then  'Using Split Networks
        ilValue = ilValue Or SPLITCOPY
    End If
    If ckcUsingSplitNetworks.Value = vbChecked Then  'Using Split Networks
        ilValue = ilValue Or SPLITNETWORKS
    End If
    If ckcUsingBarter.Value = vbChecked Then  'Using Split Networks
        ilValue = ilValue Or BARTER
    End If
    If ckcStrongPassword.Value = vbChecked Then  'Using Split Networks
        ilValue = ilValue Or STRONGPASSWORD
    End If
    If rbcRMerchPromo(1).Value Then
        ilValue = ilValue Or MERCHPROMOBYDOLLAR
    End If
    '5/31/19: Question has been hidden as Invoicing does not handle the mixture
    'If ckcSales(13).Value = vbChecked Then  'Mix Air Time and Rep on same contract
    '    ilValue = ilValue Or MIXAIRTIMEANDREP
    'End If
    If ckcGreatPlainGL.Value = vbChecked Then  'Mix Air Time and Rep on same contract
        ilValue = ilValue Or GREATPLAINS
    End If
    tgSpf.sUsingFeatures2 = Chr$(ilValue)

    ilValue = 0
    If ckcSales(8).Visible Then
        If ckcSales(8).Value = vbChecked Then  'Using Split Networks
            ilValue = ilValue Or USINGHUB
        End If
    End If
    If ckcTaxOn(0).Value = vbChecked Then  'Tax on Air time
        ilValue = ilValue Or TAXONAIRTIME
    End If
    If ckcTaxOn(1).Value = vbChecked Then  'Tax on Air Time
        ilValue = ilValue Or TAXONNTR
    End If
    If ckcCopy(1).Value = vbChecked Then  'Using Promo Copy
        ilValue = ilValue Or PROMOCOPY
    End If
    If ckcCopy(2).Value = vbChecked Then  'Media Copy by Vehicle
        ilValue = ilValue Or MEDIACODEBYVEH
    End If
    If udcSiteTabs.Automation(11) = vbChecked Then  'Include Media Definition
        ilValue = ilValue Or INCMEDIACODEAUDIOVAULT
    End If
    If rbcCSchdPromo(1).Value = True Then
        If ckcBookInto(0).Value = vbChecked Then
            ilValue = ilValue Or PROMOINTOCONTRACTAVAILS
        End If
    End If
    If rbcCSchdPSA(1).Value = True Then
        If ckcBookInto(1).Value = vbChecked Then
            ilValue = ilValue Or PSAINTOCONTRACTAVAILS
        End If
    End If
    tgSpf.sUsingFeatures3 = Chr$(ilValue)

    ilValue = 0
    If ckcInv(3).Value = vbChecked Then  'Lock Box by Vehicle
        ilValue = ilValue Or LOCKBOXBYVEHICLE
    End If
    If ckcSales(3).Value = vbChecked Then  'Lock Box by Vehicle
        ilValue = ilValue Or ALLOWMOVEONTODAY
    End If
    If ckcOptionFields(10).Value = vbChecked Then  'Lock Box by Vehicle
        ilValue = ilValue Or CHGBILLEDPRICE
    End If
    If rbcTaxOnAirTime(1).Value Then 'Tax on Air Time by USA
        ilValue = ilValue Or TAXBYUSA
    ElseIf rbcTaxOnAirTime(2).Value Then 'Tax on Air Time by Canadian
        ilValue = ilValue Or TAXBYCANADA
    End If
    If rbcInvSortBy(1).Value Then 'Sort by Vehicle
        ilValue = ilValue Or INVSORTBYVEHICLE
    End If
    tgSpf.sUsingFeatures4 = Chr$(ilValue)

    ilValue = 0
    If ckcRemoteExport.Value = vbChecked Then
        ilValue = ilValue Or REMOTEEXPORT
    End If
    If ckcRemoteImport.Value = vbChecked Then
        ilValue = ilValue Or REMOTEIMPORT
    End If
    If ckcInv(1).Value = vbChecked Then
        ilValue = ilValue Or COMBINEAIRNTR
    End If
    If rbcCSortBy(0).Value Then
        ilValue = ilValue Or CNTRINVSORTRC
    End If
    If rbcCSortBy(2).Value Then
        ilValue = ilValue Or CNTRINVSORTLN
    End If
    If ckcGUseAffSys(1).Value = vbChecked Then
        ilValue = ilValue Or STATIONINTERFACE
    End If
    If ckcGUseAffSys(2).Value = vbChecked Then
        ilValue = ilValue Or RADAR
    End If
    If ckcSuppressTimeForm1.Value = vbChecked Then
        ilValue = ilValue Or SUPPRESSTIMEFORM1
    End If
    tgSpf.sUsingFeatures5 = Chr$(ilValue)


    ilValue = 0
    If plcBB(0).Visible Then
        If rbcBBOnLine(1).Value = True Then
            ilValue = ilValue Or BBNOTSEPARATELINE
        End If
    End If
    If rbcBBType(1).Value Then
        ilValue = ilValue Or BBCLOSEST
    End If
    If ckcInstallment.Value = vbChecked Then
        ilValue = ilValue Or INSTALLMENT
        If rbcInstRev(0).Value Then
            ilValue = ilValue Or INSTALLMENTREVENUEEARNED
        End If
    End If
    If ckcGetPaidExport.Value = vbChecked Then
        ilValue = ilValue Or GETPAIDEXPORT
    End If
    If ckcDigital.Value = vbChecked Then
        ilValue = ilValue Or DIGITALCONTENT
    End If
    If ckcOptionFields(11).Value = vbChecked Then
        ilValue = ilValue Or GUARBYGRIMP
    End If
'    If ckcOptionFields(12).Value = vbChecked Then
'        ilValue = ilValue Or INVEXPORTPARAMETERS
'    End If

    If ckcInvoiceExport.Value = vbChecked Then
        ilValue = ilValue Or INVEXPORTPARAMETERS
    End If

    tgSpf.sUsingFeatures6 = Chr$(ilValue)

    ilValue = 0
    If ckcCSIBackup.Value = vbChecked Then
        ilValue = ilValue Or CSIBACKUP
    End If
    If rbcCommBy(0).Value Then
        ilValue = ilValue Or BONUSCOMM
    End If
    If rbcCommBy(2).Value Then
        ilValue = ilValue Or COMMFISCALYEAR
    End If
    If ckcRevenueExport.Value = vbChecked Then
        ilValue = ilValue Or EXPORTREVENUE
    End If
    'If udcSiteTabs.Automation(13) = vbChecked Then
    If ckcXDSBy(0).Value = vbChecked Then
        ilValue = ilValue Or XDIGITALISCIEXPORT
    End If
    If udcSiteTabs.Automation(14) = vbChecked Then
        ckcRegionMixLen.Value = vbUnchecked
        ilValue = ilValue Or WEGENEREXPORT
    End If
    If udcSiteTabs.Automation(15) = vbChecked Then
        ckcRegionMixLen.Value = vbUnchecked
        ilValue = ilValue Or OLAEXPORT
    End If
    If ckcRegionMixLen.Value = vbChecked Then
        ilValue = ilValue Or REGIONMIXLEN
    End If
    tgSpf.sUsingFeatures7 = Chr$(ilValue)

    ilValue = 0
    If ckcOverrideOptions(5).Value = vbChecked Then
        ilValue = ilValue Or LRMANDATORY
    End If
    If ckcCntr(0).Value = vbChecked Then
        ilValue = ilValue Or SHOWCMMTONDETAILPAGE
    End If
    If ckcMetroSplitCopy.Value = vbChecked Then
        ilValue = ilValue Or ALLOWMSASPLITCOPY
    End If
    If udcSiteTabs.Automation(17) = vbChecked Then
        ilValue = ilValue Or RIVENDELLEXPORT
    End If
    'If udcSiteTabs.Automation(18) = vbChecked Then
    If ckcXDSBy(1).Value = vbChecked Then
        ilValue = ilValue Or XDIGITALBREAKEXPORT
    End If
    If udcSiteTabs.Automation(19) = vbChecked Then
        ilValue = ilValue Or ISCIEXPORT
    End If
    If ckcPrefeed.Value = vbChecked Then
        ilValue = ilValue Or PREFEEDDEF
    End If
    If ckcInv(5).Value = vbChecked Then
        ilValue = ilValue Or REPBYDT
    End If
    tgSpf.sUsingFeatures8 = Chr$(ilValue)

    ilValue = 0
    If ckcAffiliateCRM.Value = vbChecked Then
        ilValue = ilValue Or AFFILIATECRM
    End If
    'If ckcPC1StPos.Value = vbChecked Then
    '    ilValue = ilValue Or PC1STPOS
    'End If
    If ckcProposalXML.Value = vbChecked Then
        ilValue = ilValue Or PROPOSALXML
    End If
    If ckcCopy(3).Value = vbChecked Then
        ilValue = ilValue Or LIMITISCI
    End If
'    If ckcCopy(4).Value = vbChecked Then
'        ilValue = ilValue Or IDCRESTRICTION
'    End If
    If ckcOptionFields(13).Value = vbChecked Then
        ilValue = ilValue Or WEEKLYBILL
    End If
    If ckcInv(6).Value = vbChecked Then
        ilValue = ilValue Or PRINTEDI
    End If
    If ckcSales(12).Value = vbChecked Then
        ilValue = ilValue Or WORDWRAPVEHICLE
    End If
    tgSpf.sUsingFeatures9 = Chr$(ilValue)

    'If ckcXDSBy(2).Value = vbChecked Then 'Add Advt/Prod to X-Digital ISCI Export
    '    tgSpf.sXSDAddAdvtToISCI = "Y"
    'Else
    '    tgSpf.sXSDAddAdvtToISCI = "N"
    'End If
    ilValue = 0
    If ckcXDSBy(2).Value = vbChecked Then 'Add Advt/Prod to X-Digital ISCI Export
        ilValue = ilValue Or ADDADVTTOISCI
    End If
    If ckcXDSBy(3).Value = vbChecked Then 'Add Advt/Prod to X-Digital ISCI Export
        ilValue = ilValue Or MIDNIGHTBASEDHOUR
    End If
    If ckcCntr(2).Value = vbChecked Then 'Add Advt/Prod to X-Digital ISCI Export
        ilValue = ilValue Or PKGLNRATEONBR
    End If
    If ckcCntr(3).Value = vbChecked Then 'Add Advt/Prod to X-Digital ISCI Export
        ilValue = ilValue Or REPLACEDELWKWITHFILLS
    End If
    If ckcVCreative.Value = vbChecked Then 'Add Advt/Prod to X-Digital ISCI Export
        ilValue = ilValue Or VCREATIVEEXPORT
    End If
    'If ckcCntr(4).Value = vbChecked Then 'Add Advt/Prod to X-Digital ISCI Export
    '    ilValue = ilValue Or CONTRACTVERIFY
    'End If
    If udcSiteTabs.Automation(13) = vbChecked Then       '8-22-08 Simian
        ilValue = ilValue Or WegenerIPump
    End If
    '9114
    If ckcXDSBy(ASTBREAK).Value = vbChecked Then 'send hb/hbp with astcodes
        ilValue = ilValue Or UNITIDBYASTCODEFORBREAK
    End If
    tgSpf.sUsingFeatures10 = Chr$(ilValue)

    '6/16/09:  Removed menu item Import Contracts because out of memory
    'If rbcImptCntr(0).Value Then 'Use Separation Rules
    '    tgSpf.sImptCntr = "Y"
    'ElseIf rbcImptCntr(1).Value Then 'Don't use separation rules
    '    tgSpf.sImptCntr = "N"
    'Else
        tgSpf.sImptCntr = "P"       'Prohibit import
    'End If
    If ckcUsingNTR.Value = vbChecked Then  'Using NTR
        tgSpf.sUsingNTR = "Y"
    Else
        tgSpf.sUsingNTR = "N"
    End If

    ilValue = 0
    If ckcUsingMatrix(0).Value = vbChecked Then  'Using Matrix Export
        ilValue = ilValue Or MATRIXEXPORT
    End If
    If ckcUsingRevenue.Value = vbChecked Then  'Using Matrix Export
        ilValue = ilValue Or REVENUEEXPORT
    End If
    If ckcUsingLiveCopy.Value = vbChecked Then  'Using Live Copy
        ilValue = ilValue Or LIVECOPY
    End If
    If ckcUsingMultiMedia.Value = vbChecked Then  'Using MultiMedia
        ilValue = ilValue Or MULTIMEDIA
    End If
    If ckcUsingLiveLog.Value = vbChecked Then  'Using Live Log
        ilValue = ilValue Or USINGLIVELOG
    End If
    'Part saved in sOverrideOptions
    If udcSiteTabs.Research(22) Then
        ilValue = ilValue Or BESTFITWEIGHT
    End If
    If udcSiteTabs.Research(31) = vbChecked Then
        ilValue = ilValue Or HIDDENOVERRIDE
    End If
    If ckcUsingRep.Value = vbChecked Then  'Using Live Log
        ilValue = ilValue Or USINGREP
    End If
    tgSpf.sUsingFeatures = Chr(ilValue)

    If ckcUsingSpecialResearch.Value = vbChecked Then
        tgSpf.sDemoEstAllowed = "Y"
    Else
        tgSpf.sDemoEstAllowed = "N"
    End If
    If ckcUsingTraffic.Value = vbChecked Then
        tgSpf.sUsingTraffic = "Y"
    Else
        tgSpf.sUsingTraffic = "N"
    End If
    'Invoicing
    tgSpf.sBPayName = Trim$(edcBPayName.Text)
    tgSpf.sBPayAddr(0) = Trim$(edcBPayAddr(0).Text)
    tgSpf.sBPayAddr(1) = Trim$(edcBPayAddr(1).Text)
    tgSpf.sBPayAddr(2) = Trim$(edcBPayAddr(2).Text)
    If rbcBLCycle(1).Value Then 'Local billing cycle
        tgSpf.sBLCycle = "C"
    ElseIf rbcBLCycle(2).Value Then 'Local billing cycle
        tgSpf.sBLCycle = "W"
    Else
        tgSpf.sBLCycle = "S"
    End If
    If rbcBRCycle(1).Value Then 'Regional billing cycle
        tgSpf.sBRCycle = "C"
    ElseIf rbcBRCycle(2).Value Then 'Regional billing cycle
        tgSpf.sBRCycle = "W"
    Else
        tgSpf.sBRCycle = "S"
    End If
    If rbcBNCycle(1).Value Then 'National billing cycle
        tgSpf.sBNCycle = "C"
    ElseIf rbcBNCycle(2).Value Then 'National billing cycle
        tgSpf.sBNCycle = "W"
    Else
        tgSpf.sBNCycle = "S"
    End If
    tgSpf.lBLowestNo = Val(edcBNo(0).Text) 'Lowest invoice #
    tgSpf.lBHighestNo = Val(edcBNo(1).Text)   'Highest invoice #
    tgSpf.lBNextNo = Val(edcBNo(2).Text) 'Next number
    slStr = edcBBillDate(0).Text
    If gValidDate(slStr) Then
        gPackDate slStr, tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1)
    Else
        slStr = ""
        gPackDate slStr, tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1)
    End If
    slStr = edcBBillDate(1).Text
    If gValidDate(slStr) Then
        gPackDate slStr, tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1)
    Else
        slStr = ""
        gPackDate slStr, tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1)
    End If
    slStr = edcBBillDate(2).Text
    If gValidDate(slStr) Then
        gPackDate slStr, tgSpf.iRepPrintDate(0), tgSpf.iRepPrintDate(1)
    Else
        slStr = ""
        gPackDate slStr, tgSpf.iRepPrintDate(0), tgSpf.iRepPrintDate(1)
    End If
    slStr = edcBBillDate(3).Text
    If gValidDate(slStr) Then
        gPackDate slStr, tmSaf.iBLastWeeklyDate(0), tmSaf.iBLastWeeklyDate(1)
    Else
        slStr = ""
        gPackDate slStr, tmSaf.iBLastWeeklyDate(0), tmSaf.iBLastWeeklyDate(1)
    End If
    slStr = edcExpt(2).Text
    If gValidDate(slStr) Then
        gPackDate slStr, tmSaf.iXDSLastImptDate(0), tmSaf.iXDSLastImptDate(1)
    Else
        slStr = "1/1/1970"
        gPackDate slStr, tmSaf.iXDSLastImptDate(0), tmSaf.iXDSLastImptDate(1)
    End If
    If rbcBCombine(1).Value Then 'Combine advertiser across vehicles on invoices
        tgSpf.sBCombine = "N"
    Else
        tgSpf.sBCombine = "Y"
    End If
    '8-23-01
    If ckcInv(2).Value = vbChecked Then  'Allow selective vehicles
        tgSpf.sInvVehSel = "Y"
    Else
        tgSpf.sInvVehSel = "N"
    End If
    tgSpf.sExport = edcBLogoSpaces(0).Text
    tgSpf.sImport = edcBLogoSpaces(1).Text

    tgSpf.sInvExportId = edcInvExportId.Text

    'If rbcBSepItem(1).Value Then 'Place item bill on separate invoice
    '    tgSpf.sBSepItem = "N"
    'Else
    '    tgSpf.sBSepItem = "Y"
    'End If
    If rbcBMissedDT(0).Value Then 'Place item bill on separate invoice
        tgSpf.sBMissedDT = "A"
    Else
        tgSpf.sBMissedDT = "R"
    End If
    If rbcPrintRepInv(1).Value Then 'Print Rep Invoice
        tgSpf.sRepRptForm = "V"
    Else
        tgSpf.sRepRptForm = "M"
    End If
    If rbcPostRepAffidavit(1).Value Then 'Post rep affidavit by Calendar
        tgSpf.sPostCalAff = "C"
    ElseIf rbcPostRepAffidavit(2).Value Then 'Post rep affidavit by Week
        tgSpf.sPostCalAff = "W"
    ElseIf rbcPostRepAffidavit(3).Value Then 'None. Post rep affidavit By Date/Time moved to Features8
        tgSpf.sPostCalAff = "N"
    Else
        tgSpf.sPostCalAff = "S" 'By standard broadcast month
    End If
    If ckcInv(4).Value = vbUnchecked Then 'Activate Day Complete testing
        tgSpf.sBActDayCompl = "N"
    Else
        tgSpf.sBActDayCompl = "Y"
    End If
    If rbcBLaserForm(1).Value Then 'Laser Form 1=Ordered, Aired & Recon; 2=Invoice & Affidavit; 3=Aired
        tgSpf.sBLaserForm = "2"
    ElseIf rbcBLaserForm(2).Value Then 'Laser Form 1=Ordered, Aired & Recon; 2=Invoice & Affidavit; 3=Aired
        tgSpf.sBLaserForm = "3"
    ElseIf rbcBLaserForm(3).Value Then      '3-8-12 3-column Aired
        tgSpf.sBLaserForm = "4"
    Else
        tgSpf.sBLaserForm = "1"
    End If
    tgSpf.sTax1Text = edcBTTax(0).Text
    tgSpf.sTax2Text = edcBTTax(1).Text
    If edcSales(7).Text <> "" Then
        tgSpf.iNoMnthNewBus = Val(edcSales(7).Text)
    Else
        tgSpf.iNoMnthNewBus = 0
    End If
    If edcSales(8).Text <> "" Then
        tgSpf.iNoMnthNewIsNew = Val(edcSales(8).Text)
    Else
        tgSpf.iNoMnthNewIsNew = 0
    End If
    If rbcNewBusYear(0).Value Then
        tgSpf.sNewBusYearType = "C"
    Else
        tgSpf.sNewBusYearType = "R"
        edcSales(7).Enabled = True
        edcSales(8).Enabled = True
    End If
    tgSpf.sEDICallLetter = edcEDI(0).Text
    tgSpf.sEDIMediaType = edcEDI(1).Text
    tgSpf.sEDIBand = edcEDI(2).Text
    
    If rbcDefFillInv(1).Value Then
        tgSpf.sDefFillInv = "N"
    Else
        tgSpf.sDefFillInv = "Y"
    End If
    If rbcBOrderDPShow(2).Value Then
        tgSpf.sBOrderDPShow = "B"
    ElseIf rbcBOrderDPShow(0).Value Then
        tgSpf.sBOrderDPShow = "N"
    Else
        tgSpf.sBOrderDPShow = "T"
    End If
    If rbcInvSpotTimeZone(0).Value Then
        tgSpf.sInvSpotTimeZone = "E"
    ElseIf rbcInvSpotTimeZone(1).Value Then
        tgSpf.sInvSpotTimeZone = "C"
    ElseIf rbcInvSpotTimeZone(2).Value Then
        tgSpf.sInvSpotTimeZone = "M"
    ElseIf rbcInvSpotTimeZone(3).Value Then
        tgSpf.sInvSpotTimeZone = "P"
    Else
        tgSpf.sInvSpotTimeZone = "N"
    End If
    
    'Accounting
    'tgSpf.iRCorp(0) = Val(edcRCorp1.Text) 'Corporate calendar # weeks (jan,..)
    'tgSpf.iRCorp(1) = Val(edcRCorp2.Text) 'Corporate calendar # weeks (feb,..)
    'tgSpf.iRCorp(2) = Val(edcRCorp3.Text) 'Corporate calendar # weeks (mar,..)
    'If rbcREnd(0).Value Then
    '    tgSpf.sRYEnd = "L"
    'Else
    '    tgSpf.sRYEnd = "D"
    'End If
    If rbcRCorpCal(1).Value Then
        tgSpf.sRUseCorpCal = "Y"
    Else
        tgSpf.sRUseCorpCal = "N"
    End If
    'lacRLastPay(1).Caption = "" 'Date last payment
    If rbcRCurrAmt(1).Value Then 'Include current amount in computing Credit limit
        tgSpf.sRCurrAmt = "N"
    Else
        tgSpf.sRCurrAmt = "Y"
    End If
    If rbcRUnbilled(1).Value Then 'Include unbilled amount in computing Credit limit
        tgSpf.sRUnbilled = "N"
    Else
        tgSpf.sRUnbilled = "Y"
    End If
    'If rbcRNewCntr(0).Value Then 'Include 1st week amount in computing Credit limit
    '    tgSpf.sRNewCntr = "W"
    'ElseIf rbcRNewCntr(1).Value Then    'Include 1st month
    '    tgSpf.sRNewCntr = "M"
    'ElseIf rbcRNewCntr(2).Value Then    'Include All
    '    tgSpf.sRNewCntr = "A"
    'Else
    '    tgSpf.sRNewCntr = "I"
    'End If
    tgSpf.iRNoWks = Val(edcRNewCntr.Text)
    slStr = edcRPctCredit.Text
    'gStrToPDN slStr, 0, 2, tgSpf.sRPctCredit
    tgSpf.iRPctCredit = gStrDecToInt(slStr, 0)
    If rbcRRP(1).Value Then
        tgSpf.sRRP = "C"
    ElseIf rbcRRP(2).Value Then
        tgSpf.sRRP = "F"
    Else
        tgSpf.sRRP = "S"
    End If
    slStr = edcRPRP.Text
    If gValidDate(slStr) Then
        gPackDate slStr, tgSpf.iRPRP(0), tgSpf.iRPRP(1)
    Else
        slStr = ""
        gPackDate slStr, tgSpf.iRPRP(0), tgSpf.iRPRP(1)
    End If
    slStr = edcRCRP.Text
    If gValidDate(slStr) Then
        gPackDate slStr, tgSpf.iRCRP(0), tgSpf.iRCRP(1)
    Else
        slStr = ""
        gPackDate slStr, tgSpf.iRCRP(0), tgSpf.iRCRP(1)
    End If
    slStr = edcRNRP.Text
    If gValidDate(slStr) Then
        gPackDate slStr, tgSpf.iRNRP(0), tgSpf.iRNRP(1)
    Else
        slStr = ""
        gPackDate slStr, tgSpf.iRNRP(0), tgSpf.iRNRP(1)
    End If
    slStr = edcRB.Text
    gStrToPDN slStr, 2, 6, tgSpf.sRB
    tgSpf.sRCollectContact = Trim$(edcRCollectContact.Text)
    gGetPhoneNo mkcRCollectPhoneNo, tgSpf.sRCollectPhoneNo
    slStr = edcRCreditDate.Text
    If gValidDate(slStr) Then
        gPackDate slStr, tgSpf.iRCreditDate(0), tgSpf.iRCreditDate(1)
    Else
        slStr = ""
        gPackDate slStr, tgSpf.iRCreditDate(0), tgSpf.iRCreditDate(1)
    End If
    'For ilLoop = 0 To UBound(smRGL) Step 1
    '    tgSpf.sRGLSuffix(ilLoop) = smRGL(ilLoop)
    'Next ilLoop
    'For ilLoop = 0 To UBound(smRName) Step 1
    '    tgSpf.sRName(ilLoop) = Trim$(smRName(ilLoop))
    '    tgSpf.sRTsfx(ilLoop) = Trim$(smRTsfx(ilLoop))
    '    tgSpf.sRAsfx(ilLoop) = Trim$(smRAsfx(ilLoop))
    'Next ilLoop
    If ckcRUseTMP(0).Value = vbChecked Then  'Using Trade Receivables
        tgSpf.sRUseTrade = "Y"
    Else
        tgSpf.sRUseTrade = "N"
    End If
    If ckcRUseTMP(1).Value = vbChecked Then  'Using Merchandising Receivables
        tgSpf.sRUseMerch = "Y"
    Else
        tgSpf.sRUseMerch = "N"
    End If
    If ckcRUseTMP(2).Value = vbChecked Then  'Using Promotion Receivables
        tgSpf.sRUsePromo = "Y"
    Else
        tgSpf.sRUsePromo = "N"
    End If
    slStr = edcBarterLPD.Text
    If gValidDate(slStr) Then
        gPackDate slStr, tgSpf.iBarterLPD(0), tgSpf.iBarterLPD(1)
    Else
        slStr = ""
        gPackDate slStr, tgSpf.iBarterLPD(0), tgSpf.iBarterLPD(1)
    End If
    If ckcRUseTMP(3).Value = vbChecked Then  'Cutoff
        tmSaf.sCreditLimitMsg = "C"
    Else
        tmSaf.sCreditLimitMsg = "W"
    End If

    '1-22-04 Default vehicle group for Reconciliation reports
    tgSpf.iReconcGroupNo = cbcReconGroup.ListIndex
    ''Locks
    'tgSpf.sLkCredit = edcGLock(0).Text
    'tgSpf.sLkLog = edcGLock(1).Text
    'tmCxf.iStrLen = Len(edcComment(2).Text)
    'tmCxf.sComment = Trim$(edcComment(2).Text) & Chr$(0) & Chr$(0) 'sgTB
    imInvCommentLen = Len(edcComment(2).Text)
    smInvComment = Trim$(edcComment(2).Text) '& Chr$(0) & Chr$(0) '2-12-03
    imContrCommentLen = Len(edcComment(0).Text)
    smContrComment = Trim$(edcComment(0).Text) '& Chr$(0) & Chr$(0)   '2-12-03
    imInsertCommentLen = Len(edcComment(1).Text)
    smInsertComment = Trim$(edcComment(1).Text) '& Chr$(0) & Chr$(0) '2-12-03
    imEstCommentLen = Len(edcComment(3).Text)
    smEstComment = Trim$(edcComment(3).Text) '& Chr$(0) & Chr$(0) '2-12-03
    imStatementCommentLen = Len(edcComment(4).Text)
    smStatementComment = Trim$(edcComment(4).Text) '& Chr$(0) & Chr$(0) '2-12-03
    imCitationCommentLen = Len(edcComment(5).Text)
    smCitationComment = Trim$(edcComment(5).Text) '& Chr$(0) & Chr$(0) '2-12-03

    'Automation
    If tgSpf.sSystemType = "R" Then
        If cbcUser(0).ListIndex > 0 Then
            tgSpf.iPriUrfCode = cbcUser(0).ItemData(cbcUser(0).ListIndex)
        Else
            tgSpf.iPriUrfCode = 0
        End If
        If cbcUser(1).ListIndex > 0 Then
            tgSpf.iSecUrfCode = cbcUser(1).ItemData(cbcUser(1).ListIndex)
        Else
            tgSpf.iSecUrfCode = 0
        End If
    Else
        tgSpf.iPriUrfCode = 0
        tgSpf.iSecUrfCode = 0
    End If
    'Using Sports
    ilValue = 0
    If ckcUsingSports.Value = vbChecked Then  'Automation Equipment
        ilValue = USINGSPORTS
        If udcSiteTabs.Sports(0) = vbChecked Then
            ilValue = ilValue Or PREEMPTREGPROG
        End If
        If udcSiteTabs.Sports(1) = vbChecked Then
            ilValue = ilValue Or USINGFEED
        End If
        If udcSiteTabs.Sports(2) = vbChecked Then
            ilValue = ilValue Or USINGLANG
        End If
    End If
    tgSpf.sSportInfo = Chr$(ilValue)
    
    If ckcUsingSports.Value = vbChecked Then  'Automation Equipment
        tmSaf.sEventTitle1 = udcSiteTabs.EventTitle(1)
        tmSaf.sEventTitle2 = udcSiteTabs.EventTitle(2)
        tmSaf.sEventSubtotal1 = udcSiteTabs.EventSubtotalTitle(1)
        tmSaf.sEventSubtotal2 = udcSiteTabs.EventSubtotalTitle(2)
    Else
        tmSaf.sEventTitle1 = ""
        tmSaf.sEventTitle2 = ""
        tmSaf.sEventSubtotal1 = ""
        tmSaf.sEventSubtotal2 = ""
    End If


    'Copy
    If rbcCUseCartNo(0).Value Then
        tgSpf.sUseCartNo = "Y"
    ElseIf rbcCUseCartNo(2).Value Then
        tgSpf.sUseCartNo = "B"
    Else
        tgSpf.sUseCartNo = "N"
    End If
    If rbcTapeShowForm(1).Value Then
        tgSpf.sTapeShowForm = "C"
    Else
        tgSpf.sTapeShowForm = "A"
    End If
    If ckcCopy(0).Value = vbChecked Then  'Using Blackouts on Logs
        tgSpf.sCBlackoutLog = "Y"
    Else
        tgSpf.sCBlackoutLog = "N"
    End If
    If rbcDefLogCopy(0).Value Then 'Default Log Copy (On=Y/Off=N)
        tgSpf.sCDefLogCopy = "Y"
    Else
        tgSpf.sCDefLogCopy = "N"
    End If
    ilValue = 0
    If rbcMGCopyAssign(0).Value Then 'MG Assign by
        ilValue = MGORIGVEHONLY
    ElseIf rbcMGCopyAssign(1).Value Then
        ilValue = MGSCHVEHONLY
    Else
        ilValue = MGEITHERVEH
    End If
    If rbcFillCopyAssign(0).Value Then 'MG Assign by
        ilValue = ilValue Or FILLORIGVEHONLY
    ElseIf rbcFillCopyAssign(1).Value Then
        ilValue = ilValue Or FILLSCHVEHONLY
    Else
        ilValue = ilValue Or FILLEITHERVEH
    End If
    If rbcMGRules(1).Value Then
        ilValue = ilValue Or MGRULESINCOPY
    End If
    If udcSiteTabs.Research(41) = vbChecked Then
        ilValue = ilValue Or RSCHCUSTDEMO
    End If
    tgSpf.sMOFCopyAssign = Chr$(ilValue)


    'Schedule
    If ckcCmmlSchStatus.Value = vbChecked Then  '
        tgSpf.sCmmlSchStatus = "A"
    Else
        tgSpf.sCmmlSchStatus = "R"
    End If
    If (Trim$(edcSchedule(0).Text) = "") Or (Trim$(edcSchedule(1).Text) = "") Or (lmSSave(1) = 0) Then
        tmSaf.lLowPrice = 0
        tmSaf.lLevelToPrice(0) = 0
        tmSaf.lLevelToPrice(1) = 0
        tmSaf.lLevelToPrice(2) = 0
        tmSaf.lLevelToPrice(3) = 0
        tmSaf.lLevelToPrice(4) = 0
        tmSaf.lLevelToPrice(5) = 0
        tmSaf.lLevelToPrice(6) = 0
        tmSaf.lLevelToPrice(7) = 0
        tmSaf.lLevelToPrice(8) = 0
        tmSaf.lLevelToPrice(9) = 0
        tmSaf.lLevelToPrice(10) = 0
        'tmSaf.lLevelToPrice(11) = 0
        tmSaf.lHighPrice = 0
    Else
        tmSaf.lLowPrice = lmSSave(1)
        tmSaf.lLevelToPrice(0) = lmSSave(2)
        tmSaf.lLevelToPrice(1) = lmSSave(3)
        tmSaf.lLevelToPrice(2) = lmSSave(4)
        tmSaf.lLevelToPrice(3) = lmSSave(5)
        tmSaf.lLevelToPrice(4) = lmSSave(6)
        tmSaf.lLevelToPrice(5) = lmSSave(7)
        tmSaf.lLevelToPrice(6) = lmSSave(8)
        tmSaf.lLevelToPrice(7) = lmSSave(9)
        tmSaf.lLevelToPrice(8) = lmSSave(10)
        tmSaf.lLevelToPrice(9) = lmSSave(11)
        tmSaf.lLevelToPrice(10) = lmSSave(12)
        tmSaf.lHighPrice = lmSSave(13)
    End If
    If ckcOverrideOptions(3).Value = vbChecked Then  'Option Fields (Right to Left):0=Allocation %;1=Acquisition Cost; 2=1st Position; 3=Preferred Days/Times; 4=Solo Avails
        tmSaf.iPreferredPct = Trim$(Str$(edcLnOverride(0).Text))
    Else
        'Retain percent even if feature turned off
        If (tmSaf.iPreferredPct < 0) Or (tmSaf.iPreferredPct > 100) Then
            tmSaf.iPreferredPct = 0
        End If
    End If
    If Trim$(edcLnOverride(1).Text) <> "" Then
        tmSaf.iWk1stSoloIndex = gStrDecToInt(edcLnOverride(1).Text, 2)
    Else
        tmSaf.iWk1stSoloIndex = 0
    End If
    If rbcInvISCIForm(1).Value Then
        tmSaf.sInvISCIForm = "L"
    ElseIf rbcInvISCIForm(2).Value Then
        tmSaf.sInvISCIForm = "W"
    Else
        tmSaf.sInvISCIForm = "R"
    End If

    'Great Plain
    tmSaf.lGPBatchNo = Val(edcGP(0).Text)
    tmSaf.sGPPrefixChar = Trim$(edcGP(1).Text)
    tmSaf.sGPCustomerNo = Trim$(edcGP(2).Text)
    'dan 6/28/11 no longer using from Address and from name is tls
    'E-Mail
    tgSite.sEmailHost = udcSiteTabs.Email(1)
    tgSite.sEmailAcctName = udcSiteTabs.Email(2)
    tgSite.sEmailPassword = udcSiteTabs.Email(3)
    tgSite.iEmailPort = Val(udcSiteTabs.Email(4))
    tgSite.sEmailFromName = udcSiteTabs.Email(5)
'    tgSite.sEmailFromAddress = udcSiteTabs.Email(6)

    If (ckcRN_Net.Value = vbChecked) Or (ckcRN_Rep.Value = vbChecked) Then
        tgNrf.sDBID = udcSiteTabs.RepNet(1)
        tgNrf.sFTPUserID = udcSiteTabs.RepNet(2)
        tgNrf.sFTPUserPW = udcSiteTabs.RepNet(3)
        tgNrf.iFTPPort = Val(udcSiteTabs.RepNet(4))
        tgNrf.sFTPAddress = udcSiteTabs.RepNet(5)
        tgNrf.sFTPImportDir = udcSiteTabs.RepNet(6)
        tgNrf.sFTPExportDir = udcSiteTabs.RepNet(7)
        tgNrf.sIISRootURL = udcSiteTabs.RepNet(8)
        tgNrf.sIISRegSection = udcSiteTabs.RepNet(9)
        If (ckcRN_Net.Value = vbChecked) Then
            tgNrf.sType = "N"
        Else
            tgNrf.sType = "R"
        End If
    End If

    tgSpf.sWegenerGroupChar = udcSiteTabs.Wegener()
    tmSaf.sIPumpZone = udcSiteTabs.WegenerIPump()
    slStr = edcSchedule(2).Text
    If gValidDate(slStr) Then
        gPackDate slStr, tmSaf.iVCreativeDate(0), tmSaf.iVCreativeDate(1)
    Else
        slStr = "1/1/1990"
        gPackDate slStr, tmSaf.iVCreativeDate(0), tmSaf.iVCreativeDate(1)
    End If
    If udcSiteTabs.Automation(18) = vbChecked Then
        tmSaf.sGenAutoFileWOSpt = "Y"
    Else
        tmSaf.sGenAutoFileWOSpt = "N"
    End If

    If ckcInv(7).Value = vbChecked Then
        tmSaf.sXMidSpotsBill = "A"
    Else
        tmSaf.sXMidSpotsBill = "O"
    End If
    If udcSiteTabs.Research(42) = vbChecked Then
        tmSaf.sHideDemoOnBR = "Y"
    Else
        tmSaf.sHideDemoOnBR = "N"
    End If
    If udcSiteTabs.Research(43) = vbChecked Then
        tmSaf.sAudByPackage = "Y"
    Else
        tmSaf.sAudByPackage = "N"
    End If
    tmSaf.sXDSHeadEndZone = "E"
    slStr = UCase$(edcExpt(1).Text)
    If (slStr = "C") Or (slStr = "M") Or (slStr = "P") Then
        tmSaf.sXDSHeadEndZone = slStr
    End If
    tmSaf.sSyncCopyInRot = "N"
    If ckcCopy(4).Value = vbChecked Then
        tmSaf.sSyncCopyInRot = "Y"
    End If
    
    'TTP 10205 - 6/21/21 - JW - Set SPFX Extended Site Options
    ilValue = 0
    tgSpfx.sInvExpProperty = ""
    tgSpfx.sInvExpPrefix = ""
    tgSpfx.sInvExpBillGroup = ""
    If ckcWOInvoiceExport.Value = vbChecked Then
       'tgSpfx.iInvExpFeature = 1
       ilValue = ilValue Or INVEXP_AUDACYWO
       'L.Bianchi 06/10/2021
       tgSpfx.sInvExpProperty = edcIE(0).Text
       tgSpfx.sInvExpPrefix = edcIE(1).Text
       tgSpfx.sInvExpBillGroup = edcIE(2).Text
    Else
        tgSpfx.iInvExpFeature = 0
    End If
    '8/17/21 - JW - TTP 10233 - Audacy: line summary export
    If ckcCntrLineExport.Value = vbChecked Then
        'tgSpfx.iInvExpFeature = tgSpfx.iInvExpFeature + 2
        ilValue = ilValue Or INVEXP_AUDACYLINE
    End If
    'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
    If ckcInv(12).Value = 1 Then
        ilValue = ilValue Or INVEXP_SELECTIVEEMAIL
    End If
    tgSpfx.iInvExpFeature = ilValue
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilValue As Integer
    Dim ilValue2 As Integer
    Dim ilValue3 As Integer
    Dim ilLevel As Integer

    imIgnoreClickEvent = True
    udcSiteTabs.Action 7, 0

    'General
    edcGClient.Text = Trim$(tgSpf.sGClient)
    smOrigClientName = Trim$(tgSpf.sGClient)
    For ilLoop = LBound(tgSpf.sGAddr) To UBound(tgSpf.sGAddr) Step 1
        edcGAddr(ilLoop) = Trim$(tgSpf.sGAddr(ilLoop))
    Next ilLoop

    If tgSpf.iMnfClientAbbr <> 0 Then
        tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If Trim$(tmMnf.sUnitType) = "A" Then
                edcGClientAbbr.Text = Trim$(tmMnf.sName)
            Else
                edcGClientAbbr.Text = ""
            End If
        Else
            edcGClientAbbr.Text = ""
        End If
    Else
        edcGClientAbbr.Text = ""
    End If

    If tgSpf.sSystemType = "R" Then
        rbcSystemType(1).Value = True
    Else
        rbcSystemType(0).Value = True
    End If

    If tmSaf.sFinalLogDisplay = "Y" Then
        ckcAllowFinalLogDisplay.Value = vbChecked
    Else
        ckcAllowFinalLogDisplay.Value = vbUnchecked
    End If
    If tmSaf.sProdProtMan = "Y" Then
        ckcCntr(1).Value = vbChecked
    Else
        ckcCntr(1).Value = vbUnchecked
    End If
    If tmSaf.sAvailGreenBar = "Y" Then
        ckcSales(15).Value = vbChecked
    Else
        ckcSales(15).Value = vbUnchecked
    End If
    If tmSaf.sInvoiceSort = "S" Then
        ckcSortSS.Value = vbChecked
    Else
        ckcSortSS.Value = vbUnchecked
    End If
    ilValue = Asc(tmSaf.sFeatures1)
    If (ilValue And MATRIXCAL) = MATRIXCAL Then 'Matrix-Cal
        ckcUsingMatrix(1).Value = vbChecked
    Else
        ckcUsingMatrix(1).Value = vbUnchecked
    End If
    If (ilValue And ENGRHIDEMEDIACODE) = ENGRHIDEMEDIACODE Then 'Engineering Export: Hide Media Code
        ckcACodes(3).Value = vbChecked
    Else
        ckcACodes(3).Value = vbUnchecked
    End If
    If (ilValue And SHOWAUDIOTYPEONBR) = SHOWAUDIOTYPEONBR Then 'Show audio type on proposal/order
        ckcCntr(5).Value = vbChecked
    Else
        ckcCntr(5).Value = vbUnchecked
    End If
    If (ilValue And SHOWPRICEONINSERTIONWITHACQUISTION) = SHOWPRICEONINSERTIONWITHACQUISTION Then 'Show Spot Prices on Insertion Orders if Acquistion Exist
        ckcSales(16).Value = vbChecked
    Else
        ckcSales(16).Value = vbUnchecked
    End If
    If (ilValue And SALESFORCEEXPORT) = SALESFORCEEXPORT Then 'Sales Force
        ckcSalesForce.Value = vbChecked
    Else
        ckcSalesForce.Value = vbUnchecked
    End If
    If (ilValue And EFFICIOEXPORT) = EFFICIOEXPORT Then 'Efficio export
        ckcEfficio.Value = vbChecked
    Else
        ckcEfficio.Value = vbUnchecked
    End If
    If (ilValue And JELLIEXPORT) = JELLIEXPORT Then 'Jelli export
        ckcJelli.Value = vbChecked
    Else
        ckcJelli.Value = vbUnchecked
    End If
    If (ilValue And COMPENSATION) = COMPENSATION Then 'COMPENSATION
        ckcCompensation.Value = vbChecked
    Else
        ckcCompensation.Value = vbUnchecked
    End If
    
    ilValue = Asc(tmSaf.sFeatures2)
    If (ilValue And EVENTREVENUE) = EVENTREVENUE Then 'Event Revenue
        ckcEventRevenue.Value = vbChecked
    Else
        ckcEventRevenue.Value = vbUnchecked
    End If
    If (ilValue And HIDEHIDDENLINES) = HIDEHIDDENLINES Then 'Event Revenue
        ckcCntr(6).Value = vbChecked
    Else
        ckcCntr(6).Value = vbUnchecked
    End If
    If (ilValue And CANCELCLAUSEMANDATORY) = CANCELCLAUSEMANDATORY Then 'Event Revenue
        ckcCntr(8).Value = vbChecked
    Else
        ckcCntr(8).Value = vbUnchecked
    End If
    If (ilValue And EMAILDISTRIBUTION) = EMAILDISTRIBUTION Then 'E-Mail distribution system
        ckcOptionFields(14).Value = vbChecked
    Else
        ckcOptionFields(14).Value = vbUnchecked
    End If
    If (ilValue And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then 'Acquisition Commissionable
        ckcRUseTMP(4).Value = vbChecked
    Else
        ckcRUseTMP(4).Value = vbUnchecked
    End If
    If (ilValue And PAYMENTONCOLLECTION) = PAYMENTONCOLLECTION Then 'Payment on Collection
        ckcRUseTMP(5).Value = vbChecked
    Else
        ckcRUseTMP(5).Value = vbUnchecked
    End If
    If (ilValue And TABLEAUEXPORT) = TABLEAUEXPORT Then 'Tableau
        ckcUsingMatrix(2).Value = vbChecked
    Else
        ckcUsingMatrix(2).Value = vbUnchecked
    End If
    If (ilValue And TABLEAUCAL) = TABLEAUCAL Then 'Tableau
        ckcUsingMatrix(3).Value = vbChecked
    Else
        ckcUsingMatrix(3).Value = vbUnchecked
    End If

    ilValue = Asc(tmSaf.sFeatures3)
    If (ilValue And SUPPRESSNETCOMM) = SUPPRESSNETCOMM Then 'Suppress Net and Commision on insertion orders
        ckcSales(17).Value = vbChecked
    Else
        ckcSales(17).Value = vbUnchecked
    End If
    
    If (ilValue And REQSTATIONPOSTING) = REQSTATIONPOSTING Then 'Require Station Posting Prior to Invoicing
        ckcInv(8).Value = vbChecked
    Else
        ckcInv(8).Value = vbUnchecked
    End If

    If (ilValue And SPLITCOPYLICENSE) = SPLITCOPYLICENSE Then 'Require Station Posting Prior to Invoicing
        rbcSplitCopyState(1).Value = True
    ElseIf (ilValue And SPLITCOPYPHYSICAL) = SPLITCOPYPHYSICAL Then
        rbcSplitCopyState(2).Value = True
    Else
        rbcSplitCopyState(0).Value = True
    End If

    If (ilValue And FREEZEDEFAULT) = FREEZEDEFAULT Then 'Freeze calculation default
        ckcCntr(7).Value = vbChecked
    Else
        ckcCntr(7).Value = vbUnchecked
    End If

    If (ilValue And INVEMAILINDEX) = INVEMAILINDEX Then 'Invoice E-Mail: Activate
        ckcInv(9).Value = vbChecked
    Else
        ckcInv(9).Value = vbUnchecked
    End If
    If (ilValue And INVSENDEMAILINDEX) = INVSENDEMAILINDEX Then 'Invoice E-Mail: Automatic
        ckcInv(10).Value = vbChecked
    Else
        ckcInv(10).Value = vbUnchecked
    End If
    If ckcInv(9).Value = vbUnchecked Then
        ckcInv(10).Value = vbUnchecked
        ckcInv(10).Enabled = False
    End If
    
    If (ilValue And SUPPRESSZERODOLLARINVEXPT) = SUPPRESSZERODOLLARINVEXPT Then 'Suppress Zero Dollars from Invoice Export
        ckcACodes(4).Value = vbChecked
    Else
        ckcACodes(4).Value = vbUnchecked
    End If
    
    ilValue = Asc(tmSaf.sFeatures4)
    If (ilValue And FILEMAKERIMPORT) = FILEMAKERIMPORT Then 'Filemaker Contract Import
        ckcOptionFields(15).Value = vbChecked
    Else
        ckcOptionFields(15).Value = vbUnchecked
    End If
    If (ilValue And ACT1CODES) = ACT1CODES Then 'ACT1 Codes
        ckcOptionFields(16).Value = vbChecked
    Else
        ckcOptionFields(16).Value = vbUnchecked
    End If
    If (ilValue And MKTNAMEONBR) = MKTNAMEONBR Then 'Market name on BR
        ckcCntr(9).Value = vbChecked
    Else
        ckcCntr(9).Value = vbUnchecked
    End If
    If (ilValue And COMPRESSTRANSACTIONS) = COMPRESSTRANSACTIONS Then 'Compress Transactions
        ckcRUseTMP(6).Value = vbChecked
    Else
        ckcRUseTMP(6).Value = vbUnchecked
    End If
    If (ilValue And AVAILINCLUDERESERVATION) = AVAILINCLUDERESERVATION Then 'Compress Transactions
        ckcCAvails(1).Value = vbChecked
    Else
        ckcCAvails(1).Value = vbUnchecked
    End If
    If (ilValue And AVAILINCLUDEREMNANT) = AVAILINCLUDEREMNANT Then 'Compress Transactions
        ckcCAvails(2).Value = vbChecked
    Else
        ckcCAvails(2).Value = vbUnchecked
    End If
    If (ilValue And AVAILINCLDEDIRECTRESPONSES) = AVAILINCLDEDIRECTRESPONSES Then 'Compress Transactions
        ckcCAvails(3).Value = vbChecked
    Else
        ckcCAvails(3).Value = vbUnchecked
    End If
    If (ilValue And AVAILINCLUDEPERINQUIRY) = AVAILINCLUDEPERINQUIRY Then 'Compress Transactions
        ckcCAvails(4).Value = vbChecked
    Else
        ckcCAvails(4).Value = vbUnchecked
    End If
    
    ilValue = Asc(tmSaf.sFeatures5)
    If (ilValue And PROGRAMMATICALLOWED) = PROGRAMMATICALLOWED Then 'Programmatic Allowed
        ckcOptionFields(17).Value = vbChecked
        If (ilValue And SHOWAVAILCOUNT) = SHOWAVAILCOUNT Then 'Show Avail Count
            ckcOptionFields(18).Value = vbChecked
        Else
            ckcOptionFields(18).Value = vbUnchecked
        End If
        If (ilValue And SHOWCPPTAB) = SHOWCPPTAB Then 'Show CPP Tab
            ckcOptionFields(19).Value = vbChecked
        Else
            ckcOptionFields(19).Value = vbUnchecked
        End If
        If (ilValue And SHOWCPMTAB) = SHOWCPMTAB Then 'Show CPM Tab
            ckcOptionFields(20).Value = vbChecked
        Else
            ckcOptionFields(20).Value = vbUnchecked
        End If
        If (ilValue And SHOWPRICETAB) = SHOWPRICETAB Then 'Show Price Tab
            ckcOptionFields(21).Value = vbChecked
        Else
            ckcOptionFields(21).Value = vbUnchecked
        End If
    Else
        ckcOptionFields(17).Value = vbUnchecked
        ckcOptionFields(18).Value = vbUnchecked
        ckcOptionFields(19).Value = vbUnchecked
        ckcOptionFields(20).Value = vbUnchecked
        ckcOptionFields(21).Value = vbUnchecked
        ckcOptionFields(18).Enabled = False
        ckcOptionFields(19).Enabled = False
        ckcOptionFields(20).Enabled = False
        ckcOptionFields(21).Enabled = False
    End If
    If (ilValue And SHOWDAYDROPDOWN) = SHOWDAYDROPDOWN Then 'Show Day Dropdown if Flight Button allowed
        ckcCntr(11).Value = vbChecked
    Else
        ckcCntr(11).Value = vbUnchecked
    End If
    If (ilValue And CSVAFFIDAVITIMPORT) = CSVAFFIDAVITIMPORT Then 'CSV Affidavit Export
        ckcOptionFields(25).Value = vbChecked
    Else
        ckcOptionFields(25).Value = vbUnchecked
    End If
    
    ilValue = Asc(tmSaf.sFeatures6)
    If (ilValue And BILLINGONINSERTIONS) = BILLINGONINSERTIONS Then 'Insertion Order include Monthly Billed Summary
        ckcSales(18).Value = vbChecked
    Else
        ckcSales(18).Value = vbUnchecked
    End If
    '9114
    If (ilValue And UNITIDBYASTCODEFORISCI) = UNITIDBYASTCODEFORISCI Then 'Insertion Order include Monthly Billed Summary
        ckcXDSBy(ASTISCI).Value = vbChecked
    Else
        ckcXDSBy(ASTISCI).Value = vbUnchecked
    End If
    If (ilValue And SIGNATUREONPROPOSAL) = SIGNATUREONPROPOSAL Then 'Print signature line on proposals
        ckcCntr(12).Value = vbChecked
    Else
        ckcCntr(12).Value = vbUnchecked
    End If
    If (ilValue And CALCULATESEARCHONLINECHG) = CALCULATESEARCHONLINECHG Then 'Calculate Research Totals on Line Change
        ckcCntr(13).Value = vbChecked
    Else
        ckcCntr(13).Value = vbUnchecked
    End If
    If (ilValue And EDIAGYCODES) = EDIAGYCODES Then 'Using EDI Client and Product codes
        ckcAEDI(2).Value = vbChecked
    Else
        ckcAEDI(2).Value = vbUnchecked
    End If
    'If (ilValue And ADVANCEAVAILS) = ADVANCEAVAILS Then 'Advance Avails (Avail/Protection/Research tab)
    '    ckcOptionFields(22).Value = vbChecked
    'Else
    '    ckcOptionFields(22).Value = vbUnchecked
    'End If
    If (ilValue And RABCALENDAR) = RABCALENDAR Then 'RAB-Calendar
        ckcOptionFields(23).Value = vbChecked
    Else
        ckcOptionFields(23).Value = vbUnchecked
    End If
    If (ilValue And OVERDUEEXPORT) = OVERDUEEXPORT Then 'Affiliate Overdue export
        ckcOptionFields(24).Value = vbChecked
    Else
        ckcOptionFields(24).Value = vbUnchecked
    End If
   
    ilValue = Asc(tmSaf.sFeatures7)
    If (ilValue And RABSTD) = RABSTD Then 'RAB-Standard
        ckcOptionFields(26).Value = vbChecked
    Else
        ckcOptionFields(26).Value = vbUnchecked
    End If
    If (ilValue And RABCALSPOTS) = RABCALSPOTS Then 'RAB-Calendar Spots
        ckcOptionFields(27).Value = vbChecked
    Else
        ckcOptionFields(27).Value = vbUnchecked
    End If
    If (ilValue And IMEDIA_MEDIACODE) = IMEDIA_MEDIACODE Then '11/5/20 - TTP # 10013 - iMediaTouch Replace COM with Media Code
        udcSiteTabs.Automation(29) = vbChecked
    Else
        udcSiteTabs.Automation(29) = vbUnchecked
    End If
    If (ilValue And CUSTOMEXPORT) = CUSTOMEXPORT Then  'TTP # 9992 - Custom Rev Export
        ckcOptionFields(31) = vbChecked
    Else
        ckcOptionFields(31) = vbUnchecked
    End If
    If (ilValue And PODBILLOVERDELIVERED) = PODBILLOVERDELIVERED Then  'Bill Over-Delivered CPM Impressions
        ckcInv(11) = vbChecked
    Else
        ckcInv(11) = vbUnchecked
    End If
    '10016
    If (ilValue And INVEMAILAIRONLY) = INVEMAILAIRONLY Then
        rbcInvEmail(AIRONLY).Value = True
    ElseIf (ilValue And INVEMAILNTRONLY) = INVEMAILNTRONLY Then
        rbcInvEmail(NTRONLY).Value = True
    Else
        rbcInvEmail(AIRANDNTR).Value = True
    End If
    '10048
    ilValue = Asc(tmSaf.sFeatures8)
    If (ilValue And PODAIRTIME) = PODAIRTIME Then
        ckcOptionFields(PODAIRTIMECKC).Value = vbChecked
    Else
        ckcOptionFields(PODAIRTIMECKC).Value = vbUnchecked
    End If
    If (ilValue And PODSPOTS) = PODSPOTS Then
        ckcOptionFields(PODSPOTSCKC).Value = vbChecked
    Else
        ckcOptionFields(PODSPOTSCKC).Value = vbUnchecked
    End If
    If (ilValue And PODADSERVER) = PODADSERVER Then
        ckcOptionFields(ADSERVERCKC).Value = vbChecked
    Else
        ckcOptionFields(ADSERVERCKC).Value = vbUnchecked
    End If
    If (ilValue And PODADSERVERVIEWONLY) = PODADSERVERVIEWONLY Then
        ckcOptionFields(PODMIXCKC).Value = vbChecked
    Else
        ckcOptionFields(PODMIXCKC).Value = vbUnchecked
    End If
    If (ilValue And PODSHOWWKOF) = PODSHOWWKOF Then
        ckcPodShowWk.Value = vbChecked
    Else
        ckcPodShowWk.Value = vbUnchecked
    End If
    mPodcastOptions -10
    If tmSaf.sAdvanceAvail = "Y" Then  'Advance Avails (Avail/Protection/Research tab)
        ckcOptionFields(22).Value = vbChecked
    Else
        ckcOptionFields(22).Value = vbUnchecked
    End If

    If (ilValue And PODCASTAUDPCT) = PODCASTAUDPCT Then 'Include Audience % on Podcast
        ckcCntr(10).Value = vbChecked
    Else
        ckcCntr(10).Value = vbUnchecked
    End If
    'edcGRetain(0).Text = Trim$(Str$(tgSpf.iGRetainCntr)) 'Retain Contract
    gUnpackDate tmSaf.iLastArchRunDate(0), tmSaf.iLastArchRunDate(1), slStr
    edcGRetain(0).Text = slStr
    If tgSpf.iRetainAffSpot = 0 Then
        tgSpf.iRetainAffSpot = 24
    End If
    If tgSpf.iRetainTrafSpot = 0 Then
        tgSpf.iRetainTrafSpot = 24
    End If
    If tmSaf.iRetainTrafProj = 0 Then
        tmSaf.iRetainTrafProj = 24
    End If
    If tgSpf.iRetainTrafProp = 0 Then
        tgSpf.iRetainTrafProp = 12
    End If
    If tmSaf.iRetainCntr = 0 Then
        tmSaf.iRetainCntr = 60
    End If
    If tmSaf.iRetainPayRevHist = 0 Then
        tmSaf.iRetainPayRevHist = 60
    End If
    
    edcGRetain(1).Text = Trim$(Str$(tgSpf.iRetainAffSpot)) 'Retain Rotation
    edcGRetain(2).Text = Trim$(Str$(tgSpf.iRetainTrafSpot)) 'Retain Contract
    edcGRetain(3).Text = Trim$(Str$(tmSaf.iRetainTrafProj)) 'Retain Contract
    edcGRetain(5).Text = Trim$(Str$(tgSpf.iRetainTrafProp)) 'Retain Dead
    edcGRetain(7).Text = Trim$(Str$(tmSaf.iRetainCntr))
    edcGRetain(8).Text = Trim$(Str$(tmSaf.iRetainPayRevHist))
    If tmSaf.iNoDaysRetainUAF > 0 Then
        edcGRetain(9).Text = Trim$(Str$(tmSaf.iNoDaysRetainUAF))
    ElseIf tmSaf.iNoDaysRetainUAF = 0 Then
        edcGRetain(9).Text = "5"
    Else
        edcGRetain(9).Text = "0"
    End If
    If tgSpf.sGTBar = "C" Then
        rbcGTBar(0).Value = True
    Else
        rbcGTBar(1).Value = True
    End If
    edcGRetainPassword.Text = Trim$(Str$(tgSpf.iGNoDaysPass))
    edcGAlertInterval.Text = Trim$(Str$(tgSpf.iGAlertInterval))
    If tgSpf.sSSellNet = "N" Then 'Selling/Airing networt
        ckcSSelling.Value = vbUnchecked
    Else
        ckcSSelling.Value = vbChecked
    End If
    If tgSpf.sSDelNet = "N" Then 'Selling/Airing networt
        ckcSDelivery.Value = vbUnchecked
    Else
        ckcSDelivery.Value = vbChecked
    End If
    'If tgSpf.sExport = "N" Then 'Export Menu
    '    ckcExport.Value = vbUnchecked
    'Else
    '    ckcExport.Value = vbChecked
    'End If
    'If tgSpf.sImport = "N" Then 'Export Menu
    '    ckcImport.Value = vbUnchecked
    'Else
    '    ckcImport.Value = vbChecked
    'End If
    'edcGLock(1).Text = tgSpf.sLkLog
    ilValue = Asc(tgSpf.sAutoType)  'Automation Equipment
    ilValue2 = Asc(tgSpf.sAutoType2)    'continuation of automation equipment types
    ilValue3 = Asc(tgSpf.sAutoType3)    'continuation of automation equipment types
    If (ilValue And DALET) = DALET Then 'Dalet)
        udcSiteTabs.Automation(0) = vbChecked
    End If
    If (ilValue And PROPHETNEXGEN) = PROPHETNEXGEN Then 'Prophet
        udcSiteTabs.Automation(2) = vbChecked
    End If
    If (ilValue And SCOTT) = SCOTT Then 'Scott
        udcSiteTabs.Automation(4) = vbChecked
    End If
    If (ilValue And DRAKE) = DRAKE Then 'Drake
        udcSiteTabs.Automation(1) = vbChecked
    End If
    If (ilValue And RCS4DIGITCART) = RCS4DIGITCART Then   'RCS
        udcSiteTabs.Automation(3) = vbChecked
    End If
    If (ilValue And PROPHETWIZARD) = PROPHETWIZARD Then   'Prophet Wizard
        udcSiteTabs.Automation(5) = vbChecked
    End If
    If (ilValue And PROPHETMEDIASTAR) = PROPHETMEDIASTAR Then   'Prophet MediaStar
        udcSiteTabs.Automation(6) = vbChecked
    End If
    If (ilValue And IMEDIATOUCH) = IMEDIATOUCH Then   'iMediaTouch
        udcSiteTabs.Automation(7) = vbChecked
    End If
    If (ilValue2 And AUDIOVAULT) = AUDIOVAULT Then   '8-10-05 Audio Vault Sat
        udcSiteTabs.Automation(8) = vbChecked
    End If
    If (ilValue2 And WIREREADY) = WIREREADY Then   '6/6/06 Wire Ready
        udcSiteTabs.Automation(9) = vbChecked
    End If
    If (ilValue2 And ENCO) = ENCO Then   '9-12-06
        udcSiteTabs.Automation(10) = vbChecked
    End If
    If (ilValue2 And RN_REP) = RN_REP Then
        ckcRN_Rep.Value = vbChecked
    End If
    If (ilValue2 And RN_NET) = RN_NET Then
        ckcRN_Net.Value = vbChecked
    End If

    If (ilValue2 And SIMIAN) = SIMIAN Then         '8-22-08
        udcSiteTabs.Automation(12) = vbChecked
    End If
    If (ilValue2 And RCS5DIGITCART) = RCS5DIGITCART Then   'RCS
        udcSiteTabs.Automation(16) = vbChecked
    End If
    If (ilValue2 And AUDIOVAULTRPS) = AUDIOVAULTRPS Then   '11-15-10 audio vault rps
        udcSiteTabs.Automation(20) = vbChecked
    End If

    If (ilValue3 And AUDIOVAULTAIR) = AUDIOVAULTAIR Then   '2/17/11 audio vault aired
        udcSiteTabs.Automation(21) = vbChecked
    End If
    If (ilValue3 And WIDEORBIT) = WIDEORBIT Then   '1/3/12: Wide Orbit
        udcSiteTabs.Automation(22) = vbChecked
    End If
    If (ilValue3 And JELLI) = JELLI Then   '5/10/12: Jelli
        udcSiteTabs.Automation(23) = vbChecked
    End If
    If (ilValue3 And ENCOESPN) = ENCOESPN Then   '5/10/12: Jelli
        udcSiteTabs.Automation(24) = vbChecked
    End If
    If (ilValue3 And SCOTT_V5) = SCOTT_V5 Then   '8-16-13 Scott V5
        udcSiteTabs.Automation(25) = vbChecked
    End If
    
    If (ilValue3 And ZETTA) = ZETTA Then   '1-5-16
        udcSiteTabs.Automation(26) = vbChecked
    End If

    If (ilValue3 And STATIONPLAYLIST) = STATIONPLAYLIST Then   '1-5-16
        udcSiteTabs.Automation(27) = vbChecked
    End If

    If (ilValue3 And RADIOMAN) = RADIOMAN Then   '5/1/20
        udcSiteTabs.Automation(28) = vbChecked
    End If

    'Sales
    'If tgSpf.sSMove = "Y" Then 'Using Proposal Contracts
    '    rbcSMove(0).Value = True
    'Else
    '    rbcSMove(1).Value = True
    'End If
    If tgSpf.sUseProdSptScr = "A" Then 'Use product or advertiser, product on spot screen
        rbcSUseProd(1).Value = True
    Else
        rbcSUseProd(0).Value = True
    End If
    If tgSpf.sHideGhostSptScr = "Y" Then
        ckcSales(2).Value = vbChecked
    Else
        ckcSales(2).Value = vbUnchecked
    End If
    edcSales(0).Text = Trim$(Str$(tgSpf.iRptDollarMag))
    If tgSpf.sUnitOr3060 = "U" Then 'Use ageing date (vs entered date)
        rbcUnitOr3060(0).Value = True
    Else
        rbcUnitOr3060(1).Value = True        'use entered date (tran date)
    End If
    If tgSpf.sAvailEqualize = "3" Then 'Combo Avail Equalize by 30, 60 or None
        rbcEqualize(0).Value = True
    ElseIf tgSpf.sAvailEqualize = "6" Then
        rbcEqualize(1).Value = True
    Else
        rbcEqualize(2).Value = True
    End If
    If tgSpf.sSEnterAgeDate = "A" Then 'Use ageing date (vs entered date)
        rbcSEnterAge(1).Value = True
    Else
        rbcSEnterAge(0).Value = True        'use entered date (tran date)
    End If
    If tgSpf.sRemoteUsers = "Y" Then 'Remote Users
        ckcSales(14).Value = vbChecked
    Else
        ckcSales(14).Value = vbUnchecked
    End If
    If tgSpf.sSAudData = "H" Then 'Audience Data: T=Thousand; H=Hundred; N = Tens; U = Units
        udcSiteTabs.Research(2) = True
    ElseIf tgSpf.sSAudData = "N" Then 'Audience Data: T=Thousand; H=Hundred; N = Tens; U = Units
        udcSiteTabs.Research(3) = True
    ElseIf tgSpf.sSAudData = "U" Then 'Audience Data: T=Thousand; H=Hundred; N = Tens; U = Units
        udcSiteTabs.Research(4) = True
    Else
        udcSiteTabs.Research(1) = True
    End If
    If tgSpf.sSGRPCPPCal = "R" Then 'GRP/CPP Calculation: R=Rating; G=GRP 2 places; A=Audience
        udcSiteTabs.Research(11) = True
    ElseIf tgSpf.sSGRPCPPCal = "G" Then
        udcSiteTabs.Research(12) = True
    Else
        udcSiteTabs.Research(13) = True
    End If
    If tgSpf.sAllowMGs = "N" Then 'Allow MG's
        ckcSales(0).Value = vbUnchecked
    Else
        ckcSales(0).Value = vbChecked
    End If
    If tgSpf.sMktBase = "Y" Then 'Market Base
        ckcSMktBase.Value = vbChecked
    Else
        ckcSMktBase.Value = vbUnchecked
    End If
    If tgSpf.sAllowDailyBuys = "Y" Then 'Allow daily buys
        ckcSales(9).Value = vbChecked
    Else
        ckcSales(9).Value = vbUnchecked
    End If
    If tgSpf.sPLMove = "M" Then 'PostLog Moves: M=MG's; O=Outsides; A=Ask
        rbcPLMove(0).Value = True
    ElseIf tgSpf.sPLMove = "O" Then
        rbcPLMove(1).Value = True
    Else
        rbcPLMove(2).Value = True
    End If
    If tgSpf.sBBsToAff = "Y" Then
        ckcSales(10).Value = vbChecked
    Else
        ckcSales(10).Value = vbUnchecked
    End If
    If tmSaf.iRptLenDefault(0) > 0 Then
        edcSales(2).Text = Trim$(Str$(tmSaf.iRptLenDefault(0)))
    Else
        edcSales(2).Text = ""
    End If
    If tmSaf.iRptLenDefault(1) > 0 Then
        edcSales(3).Text = Trim$(Str$(tmSaf.iRptLenDefault(1)))
    Else
        edcSales(3).Text = ""
    End If
    If tmSaf.iRptLenDefault(2) > 0 Then
        edcSales(4).Text = Trim$(Str$(tmSaf.iRptLenDefault(2)))
    Else
        edcSales(4).Text = ""
    End If
    If tmSaf.iRptLenDefault(3) > 0 Then
        edcSales(5).Text = Trim$(Str$(tmSaf.iRptLenDefault(3)))
    Else
        edcSales(5).Text = ""
    End If
    If tmSaf.iRptLenDefault(4) > 0 Then
        edcSales(6).Text = Trim$(Str$(tmSaf.iRptLenDefault(4)))
    Else
        edcSales(6).Text = ""
    End If
    If tmSaf.sReSchdXCal = "N" Then  'Reschedule across calendar month
        ckcSales(11).Value = vbChecked
    Else
        ckcSales(11).Value = vbUnchecked
    End If

     '2-7-03
    If tgSpf.sInsertAddr = "S" Then 'get Insertion Address from P = payee (contract), S = Site (invoice tab), or Vehicle
        rbcInsertAddr(1).Value = True
    ElseIf tgSpf.sInsertAddr = "V" Then
        rbcInsertAddr(2).Value = True
    Else
        rbcInsertAddr(0).Value = True       'default to payee
    End If
    If tgSpf.sSSRetainDate = "Y" Then 'Retain Date
        ckcSales(1).Value = vbChecked
    Else
        ckcSales(1).Value = vbUnchecked
    End If
    If tgSpf.sSBrStdQt = "Y" Then 'BR by standard Quarter
        ckcSales(4).Value = vbChecked
    Else
        ckcSales(4).Value = vbUnchecked
    End If
    rbcCommBy(5).Value = True
    If tgSpf.sSubCompany = "Y" Then 'BR by standard Quarter
        'ckcSSubCompany.Value = vbChecked
        rbcCommBy(4).Value = True
    Else
        'ckcSSubCompany.Value = vbUnchecked
    End If
    If tgSpf.sCommByCntr = "Y" Then 'BR by standard Quarter
        'ckcSCommByCntr.Value = vbChecked
        rbcCommBy(3).Value = True
    Else
        'ckcSCommByCntr.Value = vbUnchecked
    End If
    If (tgSpf.iVehLen > 40) Or (tgSpf.iVehLen < 0) Then
        tgSpf.iVehLen = 20
    End If
    edcSales(1).Text = Trim$(Str$(tgSpf.iVehLen))

    If tgSpf.sInvAirOrder = "S" Then 'Show Ordered, Update Ordered
        rbcSInvCntr(0).Value = True
    ElseIf tgSpf.sInvAirOrder = "O" Then 'Show Ordered, Update Aired
        rbcSInvCntr(1).Value = True
    ElseIf tgSpf.sInvAirOrder = "2" Then 'Show Aired minus Missed, Update Order
        rbcSInvCntr(3).Value = True
    Else
        rbcSInvCntr(2).Value = True     'As Aired
    End If
    If tgSpf.sCPkOrdered = "Y" Then
        ckcSales(5).Value = vbChecked
    Else
        ckcSales(5).Value = vbUnchecked
    End If
    If tgSpf.sCPkAired = "Y" Then
        ckcSales(6).Value = vbChecked
    Else
        ckcSales(6).Value = vbUnchecked
    End If
    If tgSpf.sCPkEqual = "Y" Then
        ckcSales(7).Value = vbChecked
    Else
        ckcSales(7).Value = vbUnchecked
    End If
    'If tgSpf.iPkageGenMeth = 0 Then
        rbcBPkageGenMeth(0).Value = True
    'Else
    '    rbcBPkageGenMeth(1).Value = True
    'End If
    'edcSHoldDays.Text = Trim$(Str$(tgSpf.iSDay))
    'edcSchRNG.Text = tgSpf.sSchRNG
    'edcSchMal.Text = tgSpf.sSchMal
    'edcSchMdl.Text = tgSpf.sSchMdl
    'edcSchMil.Text = tgSpf.sSchMil
    'edcSchCycle.Text = tgSpf.sSchCycle
    'edcSchMove.Text = tgSpf.sSchMove
    'edcSchCompact.Text = tgSpf.sSchCompact
    'edcSchPreempt.Text = tgSpf.sSchPreempt
    'edcSchHour.Text = tgSpf.sSchHour
    'edcSchDay.Text = tgSpf.sSchDay
    'edcSchQH.Text = tgSpf.sSchQH
    'edcGLock(0).Text = tgSpf.sLkCredit
    'Agency/Advertiser
    'edcAProdNameSize.Text = Trim$(Str$(tgSpf.iAProd))
    If tgSpf.sAExcl = "N" Then 'Program exclusions
        ckcOptionFields(9).Value = vbUnchecked
    Else
        ckcOptionFields(9).Value = vbChecked
    End If
    edcExpt(0).Text = Trim$(Str$(tgSpf.iATargets))
    If tgSpf.sAISCI = "A" Then 'ISCI on invoices
        rbcAISCI(0).Value = True
    ElseIf tgSpf.sAISCI = "X" Then 'ISCI on invoices
        rbcAISCI(1).Value = True
    ElseIf tgSpf.sAISCI = "Y" Then 'ISCI on invoices
        rbcAISCI(2).Value = True
    Else
        rbcAISCI(3).Value = True
    End If
    If tgSpf.sAEDIC = "N" Then 'EDI for contract
        ckcAEDI(0).Value = vbUnchecked
    Else
        ckcAEDI(0).Value = vbChecked
    End If
    If tgSpf.sAEDII = "N" Then 'EDI for Invoices
        ckcAEDI(1).Value = vbUnchecked
        ckcAEDI(2).Value = vbUnchecked
        ckcAEDI(2).Enabled = False
    Else
        ckcAEDI(1).Value = vbChecked
        ckcAEDI(2).Enabled = True
    End If
    If tgSpf.sAPrtStyle = "W" Then 'Contract print style
        rbcAPrtStyle(0).Value = True
    ElseIf tgSpf.sAPrtStyle = "N" Then
        rbcAPrtStyle(1).Value = True
    Else
        rbcAPrtStyle(2).Value = True
    End If
    If tgSpf.sARepCodes = "N" Then 'Rep codes
        ckcACodes(0).Value = vbUnchecked
    Else
        ckcACodes(0).Value = vbChecked
    End If
    If tgSpf.sAStnCodes = "N" Then 'Station codes
        ckcACodes(1).Value = vbUnchecked
    Else
        ckcACodes(1).Value = vbChecked
    End If
    If tgSpf.sAAgyCodes = "N" Then 'Agency codes
        ckcACodes(2).Value = vbUnchecked
    Else
        ckcACodes(2).Value = vbChecked
    End If
    If tgSpf.iMnfInvTerms <> 0 Then
        tmMnfSrchKey.iCode = tgSpf.iMnfInvTerms
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If Trim$(tmMnf.sUnitType) = "D" Then
                edcTerms.Text = Trim$(tmMnf.sName)
            Else
                edcTerms.Text = sgDefaultTerms
            End If
        Else
            edcTerms.Text = sgDefaultTerms
        End If
    Else
        edcTerms.Text = sgDefaultTerms
    End If
    'Penny variance
    'gPDNToStr tgSpf.sRPenny, 2, slStr
    slStr = gLongToStrDec(tgSpf.lRPenny, 2)
    edcAPenny.Text = slStr
    'Contract/Copy
    edcCNo(0).Text = Trim$(Str$(tgSpf.lCLowestNo)) 'Lowest contract #
    edcCNo(1).Text = Trim$(Str$(tgSpf.lCHighestNo))   'Highest contract #- Not using
    edcCNo(2).Text = Trim$(Str$(tgSpf.lCNextNo)) 'Next number
    If tgSpf.sCEstNo = "N" Then 'Using estimate number
        ckcOptionFields(8).Value = vbUnchecked
    Else
        ckcOptionFields(8).Value = vbChecked
    End If
    'If tgSpf.sCRot = "I" Then 'Copy rotation
    '    rbcCRot(0).Value = True
    'ElseIf tgSpf.sCRot = "M" Then
    '    rbcCRot(1).Value = True
    'Else
    '    rbcCRot(2).Value = True
    'End If
    If tgSpf.sAllowPrelLog = "Y" Then 'allow preliminary logs
        ckcAllowPrelLog.Value = vbChecked
    Else
        ckcAllowPrelLog.Value = vbUnchecked
    End If
    If tgSpf.sVirtPkgCompute = "H" Then 'Default Log Copy (Y=On; N=Off)
        rbcVirtPkg(0).Value = True
    Else
        rbcVirtPkg(1).Value = True
    End If
    If tgSpf.sSchdRemnant = "Y" Then 'Schedule Remnant Contracts like Standard Contracts
        rbcCSchdRemnant(1).Value = True
    Else
        rbcCSchdRemnant(0).Value = True
    End If
    If tgSpf.sSchdPromo = "Y" Then 'Schedule Remnant Contracts like Standard Contracts
        rbcCSchdPromo(1).Value = True
    Else
        rbcCSchdPromo(0).Value = True
    End If
    If tgSpf.sSchdPSA = "Y" Then 'Schedule Remnant Contracts like Standard Contracts
        rbcCSchdPSA(1).Value = True
    Else
        rbcCSchdPSA(0).Value = True
    End If
    If tgSpf.sCBumpPast = "N" Then 'Using estimate number
        ckcCBump.Value = vbUnchecked
    Else
        ckcCBump.Value = vbChecked
    End If
    '10842
'    If tgSpf.sCLnStdQt = "Y" Then 'BR by standard Quarter
'        ckcCLnStdQt.Value = vbChecked
'    Else
'        ckcCLnStdQt.Value = vbUnchecked
'    End If
    edcBPayName.Text = Trim$(tgSpf.sBPayName)
    edcBPayAddr(0).Text = Trim$(tgSpf.sBPayAddr(0))
    edcBPayAddr(1).Text = Trim$(tgSpf.sBPayAddr(1))
    edcBPayAddr(2).Text = Trim$(tgSpf.sBPayAddr(2))
    'If tgSpf.sSUseResv = "Y" Then 'Using Reservation Contracts
    '    rbcSRes(0).Value = True
    'Else
    '    rbcSRes(1).Value = True
    'End If
    'If tgSpf.sSUseRem = "Y" Then 'Using Proposal Contracts
    '    rbcSRem(0).Value = True
    'Else
    '    rbcSRem(1).Value = True
    'End If
    'If tgSpf.sSUseDR = "Y" Then 'Using Direct Response Contracts
    '    rbcSDR(0).Value = True
    'Else
    '    rbcSDR(1).Value = True
    'End If
    'If tgSpf.sSUsePI = "Y" Then 'Using per inquiry Contracts
    '    rbcSPI(0).Value = True
    'Else
    '    rbcSPI(1).Value = True
    'End If
    'If tgSpf.sSUsePSA = "Y" Then  'Using PSA Contracts
    '    rbcSPSA(0).Value = True
    'Else
    '    rbcSPSA(1).Value = True
    'End If
    'If tgSpf.sSUsePromo = "Y" Then 'Using Promo
    '    rbcSPromo(0).Value = True
    'Else
    '    rbcSPromo(1).Value = True
    'End If
    If tgSpf.iReallMnfDemo > 0 Then
        For ilLoop = LBound(tgDemoCode) To UBound(tgDemoCode) - 1 Step 1
            slNameCode = tgDemoCode(ilLoop).sKey  'Traffic!lbcDemoCode.List(lbcDemo(ilLoop).ListIndex - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tgSpf.iReallMnfDemo = CInt(slCode) Then
                cbcReallDemo.ListIndex = ilLoop + 1
                Exit For
            End If
        Next ilLoop
    Else
        cbcReallDemo.ListIndex = 0
    End If
    gUnpackDate tgSpf.iReallDate(0), tgSpf.iReallDate(1), slStr
    smReallDateCaption = slStr
    plcReallDate_Paint
    If tgSpf.lDiscCurrCntrNo > 0 Then
        edcDiscNo.Text = Trim$(Str$(tgSpf.lDiscCurrCntrNo)) 'Current Discrepancy contract #
    Else
        edcDiscNo.Text = ""
    End If
    gUnpackDate tgSpf.iDiscDateRun(0), tgSpf.iDiscDateRun(1), slStr
    If gValidDate(slStr) Then
        smDiscDateCaption = slStr
    Else
        smDiscDateCaption = ""
    End If
    plcDiscDate_Paint
    If tgSpf.sCIncludeMissDB = "Y" Then 'Include Missed in Demo Bar Avails
        ckcCAvails(0).Value = vbChecked
    Else
        ckcCAvails(0).Value = vbUnchecked
    End If
    
    If tgSpf.sCWarnMsg = "Y" Then 'Contract Warning Message
        ckcCWarnMsg.Value = vbChecked
    Else
        ckcCWarnMsg.Value = vbUnchecked
    End If
     
    'TTP 11033 JJB 2024-05-11
    If (tgSpfx.iSchdFeature And SCHD_DISALLOWSAMEUSER) = SCHD_DISALLOWSAMEUSER Then  'Schedule screen: disallow allow user who last edited a contract from being able to schedule it
        ckcDisallowAuthorScheduling.Value = vbChecked
    Else
        ckcDisallowAuthorScheduling.Value = vbUnchecked
    End If
    
    If tgSpf.sCAudPkg = "Y" Then 'Contract Warning Message
        ckcCAudPkg.Value = vbChecked
    Else
        ckcCAudPkg.Value = vbUnchecked
    End If
    'Options
    If tgSpf.sGUsePropSys = "Y" Then 'Using Proposal System
        ckcGUsePropSys.Value = vbChecked
    Else
        ckcGUsePropSys.Value = vbUnchecked
    End If
    If tgSpf.sGUseAffSys = "Y" Then 'Using Affiliate System
        ckcGUseAffSys(0).Value = vbChecked
    Else
        ckcGUseAffSys(0).Value = vbUnchecked
    End If
    If tgSpf.sSystemType <> "R" Then
        If tgSpf.sGUseAffFeed = "Y" Then 'Using Affiliate Feed (Regional Copy)
            ckcGUseAffFeed.Value = vbChecked
        Else
            ckcGUseAffFeed.Value = vbUnchecked
        End If
        frcPollUsers.Visible = False
    Else
        ckcGUseAffFeed.Visible = False
        frcPollUsers.Visible = True
    End If
    If tgSpf.sCUseSegments = "Y" Then 'Using Segments
        ckcCUseSegments.Value = vbChecked
    Else
        ckcCUseSegments.Value = vbUnchecked
    End If
    If tgSpf.sUsingBBs = "Y" Then 'Using Segments
        ckcUsingBBs.Value = vbChecked
    Else
        ckcUsingBBs.Value = vbUnchecked
    End If
    '6/16/09:  Removed menu item Import Contracts because out of memory
    'If tgSpf.sImptCntr = "Y" Then 'Use Separation rules with Moves
    '    rbcImptCntr(0).Value = True
    'ElseIf tgSpf.sImptCntr = "N" Then 'Don't Use Separation rules with Moves
    '    rbcImptCntr(1).Value = True
    'Else
    '    rbcImptCntr(2).Value = True
    'End If
    If tgSpf.sUsingNTR = "Y" Then 'Using NTR
        ckcUsingNTR.Value = vbChecked
    Else
        ckcUsingNTR.Value = vbUnchecked
    End If
    'If tgSpf.sUsingMatrix = "Y" Then 'Using Matrix Export
    '    ckcUsingMatrix.Value = vbChecked
    'Else
    '    ckcUsingMatrix.Value = vbUnchecked
    'End If
    ilValue = Asc(tgSpf.sUsingFeatures)  'Option Fields in Orders/Proposals
    If (ilValue And MATRIXEXPORT) = MATRIXEXPORT Then
        ckcUsingMatrix(0).Value = vbChecked
    Else
        ckcUsingMatrix(0).Value = vbUnchecked
    End If
    If (ilValue And REVENUEEXPORT) = REVENUEEXPORT Then
        ckcUsingRevenue.Value = vbChecked
    Else
        ckcUsingRevenue.Value = vbUnchecked
    End If
    If (ilValue And LIVECOPY) = LIVECOPY Then
        ckcUsingLiveCopy.Value = vbChecked
    Else
        ckcUsingLiveCopy.Value = vbUnchecked
    End If
    If (ilValue And MULTIMEDIA) = MULTIMEDIA Then
        ckcUsingMultiMedia.Value = vbChecked
    Else
        ckcUsingMultiMedia.Value = vbUnchecked
    End If
    If (ilValue And USINGLIVELOG) = USINGLIVELOG Then
        ckcUsingLiveLog.Value = vbChecked
    Else
        ckcUsingLiveLog.Value = vbUnchecked
    End If
    'Part in sOverrideOptions
    If (ilValue And BESTFITWEIGHT) = BESTFITWEIGHT Then
        udcSiteTabs.Research(22) = True
    ElseIf (Asc(tgSpf.sOverrideOptions) And BESTFITWEIGHTNONE) <> BESTFITWEIGHTNONE Then
        udcSiteTabs.Research(21) = True
    End If
    If (ilValue And HIDDENOVERRIDE) = HIDDENOVERRIDE Then
        udcSiteTabs.Research(31) = vbChecked
    Else
        udcSiteTabs.Research(31) = vbUnchecked
    End If
    If (ilValue And USINGREP) = USINGREP Then
        ckcUsingRep.Value = vbChecked
    Else
        ckcUsingRep.Value = vbUnchecked
    End If

    If tgSpf.sDemoEstAllowed = "Y" Then
        ckcUsingSpecialResearch.Value = vbChecked
    Else
        ckcUsingSpecialResearch.Value = vbUnchecked
    End If
    If tgSpf.sUsingTraffic <> "N" Then
        ckcUsingTraffic.Value = vbChecked
    Else
        ckcUsingTraffic.Value = vbUnchecked
    End If
    ilValue = Asc(tgSpf.sOptionFields)  'Option Fields in Orders/Proposals
    'If (ilValue And &H1) = &H1 Then 'Projections
    If (ilValue And OFPROJECTION) = OFPROJECTION Then 'Projections
        ckcOptionFields(0).Value = vbChecked
    Else
        ckcOptionFields(0).Value = vbUnchecked
    End If
    'If (ilValue And &H2) = &H2 Then 'Business Category
    If (ilValue And OFBUSCAT) = OFBUSCAT Then 'Business Category
        ckcOptionFields(1).Value = vbChecked
    Else
        ckcOptionFields(1).Value = vbUnchecked
    End If
    'If (ilValue And &H4) = &H4 Then 'Share
    If (ilValue And OFSHARE) = OFSHARE Then 'Share
        ckcOptionFields(2).Value = vbChecked
    Else
        ckcOptionFields(2).Value = vbUnchecked
    End If
    'If (ilValue And &H8) = &H8 Then 'Revenue Set
    If (ilValue And OFREVSET) = OFREVSET Then 'Revenue Set
        ckcOptionFields(3).Value = vbChecked
    Else
        ckcOptionFields(3).Value = vbUnchecked
    End If
    'If (ilValue And &H10) = &H10 Then   'Delivery Guarantee %
    If (ilValue And OFDELGUAR) = OFDELGUAR Then   'Delivery Guarantee %
        ckcOptionFields(4).Value = vbChecked
    Else
        ckcOptionFields(4).Value = vbUnchecked
    End If
    'If (ilValue And &H20) = &H20 Then   'Billing Cycle
    If (ilValue And OFCALENDARBILL) = OFCALENDARBILL Then   'Billing Cycle
        ckcOptionFields(5).Value = vbChecked
    Else
        ckcOptionFields(5).Value = vbUnchecked
    End If
    'If (ilValue And &H40) = &H40 Then   'Co-op Billing
    If (ilValue And OFCOOPBILL) = OFCOOPBILL Then   'Co-op Billing
        ckcOptionFields(6).Value = vbChecked
    Else
        ckcOptionFields(6).Value = vbUnchecked
    End If
    'If (ilValue And &H80) = &H80 Then   'Research
    If (ilValue And OFRESEARCH) = OFRESEARCH Then   'Research
        ckcOptionFields(7).Value = vbChecked
    Else
        ckcOptionFields(7).Value = vbUnchecked
    End If

    ilValue = Asc(tgSpf.sOverrideOptions)  'Option Fields in Orders/Proposals
    If (ilValue And SPALLOCATION) = SPALLOCATION Then 'Allocation %
        ckcOverrideOptions(0).Value = vbChecked
    Else
        ckcOverrideOptions(0).Value = vbUnchecked
    End If
    '6/7/15: renamed acquisition to ntr acquisition and replaced acquisition in site override with Barter in system options
    If (ilValue And SPNTRACQUISITION) = SPNTRACQUISITION Then 'Acquisition Cose
        ckcOverrideOptions(1).Value = vbChecked
    Else
        ckcOverrideOptions(1).Value = vbUnchecked
    End If
    If (ilValue And SP1STPOSITION) = SP1STPOSITION Then '1st Position
        ckcOverrideOptions(2).Value = vbChecked
    Else
        ckcOverrideOptions(2).Value = vbUnchecked
    End If
    If (ilValue And SPPREFERREDDT) = SPPREFERREDDT Then 'Preferred Days/Times
        ckcOverrideOptions(3).Value = vbChecked
    Else
        ckcOverrideOptions(3).Value = vbUnchecked
    End If
    If (ilValue And SPSOLOAVAIL) = SPSOLOAVAIL Then   'Solo Avails
        ckcOverrideOptions(4).Value = vbChecked
    Else
        ckcOverrideOptions(4).Value = vbUnchecked
    End If
    'Also part in sUsingFeatures
    If (ilValue And BESTFITWEIGHTNONE) = BESTFITWEIGHTNONE Then
        udcSiteTabs.Research(23) = True
    End If
    If (ilValue And BBSAMELINE) = BBSAMELINE Then   'BBs on Same Line
        ckcInv(0).Value = vbChecked
    Else
        ckcInv(0).Value = vbUnchecked
    End If




    ilValue = Asc(tgSpf.sUsingFeatures2)  'Option Fields in Orders/Proposals
    '12/9/14: Removed Regional Copy as not used any longer
    'If (ilValue And REGIONALCOPY) = REGIONALCOPY Then
    '    ckcRegionalCopy.Value = vbChecked
    'Else
        ckcRegionalCopy.Value = vbUnchecked
    'End If
    If (ilValue And SPLITCOPY) = SPLITCOPY Then
        ckcUsingSplitCopy.Value = vbChecked
    Else
        ckcUsingSplitCopy.Value = vbUnchecked
    End If
    If (ilValue And SPLITNETWORKS) = SPLITNETWORKS Then
        ckcUsingSplitNetworks.Value = vbChecked
    Else
        ckcUsingSplitNetworks.Value = vbUnchecked
    End If
    If (ilValue And BARTER) = BARTER Then
        ckcUsingBarter.Value = vbChecked
    Else
        ckcUsingBarter.Value = vbUnchecked
    End If
    If (ilValue And STRONGPASSWORD) = STRONGPASSWORD Then
        ckcStrongPassword.Value = vbChecked
    Else
        ckcStrongPassword.Value = vbUnchecked
    End If
    If (ilValue And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
        rbcRMerchPromo(1).Value = True
    Else
        rbcRMerchPromo(0).Value = True
    End If
    '5/31/19: Turning this feature off as invoicing does not handle it (Either rep or air time billed, the other is not)
    ''5/6/15: Re-activate this feature
    'If (ilValue And MIXAIRTIMEANDREP) = MIXAIRTIMEANDREP Then
    '    ckcSales(13).Value = vbChecked
    'Else
        ckcSales(13).Value = vbUnchecked
    'End If
    If (ilValue And GREATPLAINS) = GREATPLAINS Then
        ckcGreatPlainGL.Value = vbChecked
        frcGP.Enabled = True
    Else
        ckcGreatPlainGL.Value = vbUnchecked
        frcGP.Enabled = False
    End If

    ilValue = Asc(tgSpf.sUsingFeatures3)  'Option Fields in Orders/Proposals
    If (ilValue And USINGHUB) = USINGHUB Then
        ckcSales(8).Value = vbChecked
    Else
        ckcSales(8).Value = vbUnchecked
    End If
    If (ilValue And TAXONAIRTIME) = TAXONAIRTIME Then
        ckcTaxOn(0).Value = vbChecked
    Else
        ckcTaxOn(0).Value = vbUnchecked
    End If
    If (ilValue And TAXONNTR) = TAXONNTR Then
        ckcTaxOn(1).Value = vbChecked
    Else
        ckcTaxOn(1).Value = vbUnchecked
    End If
    If (ilValue And PROMOCOPY) = PROMOCOPY Then 'Using Promo copy
        ckcCopy(1).Value = vbChecked
    Else
        ckcCopy(1).Value = vbUnchecked
    End If
    If (ilValue And MEDIACODEBYVEH) = MEDIACODEBYVEH Then   'Media copy by Vehicle
        ckcCopy(2).Value = vbChecked
    Else
        ckcCopy(2).Value = vbUnchecked
    End If
    If (ilValue And INCMEDIACODEAUDIOVAULT) = INCMEDIACODEAUDIOVAULT Then
        udcSiteTabs.Automation(11) = vbChecked
    Else
        udcSiteTabs.Automation(11) = vbUnchecked
    End If
    If rbcCSchdPSA(0).Value = True Then
        ckcBookInto(1).Value = vbUnchecked
        ckcBookInto(1).Enabled = False
    Else
        ckcBookInto(1).Enabled = True
        If (ilValue And PSAINTOCONTRACTAVAILS) = PSAINTOCONTRACTAVAILS Then
            ckcBookInto(1).Value = vbChecked
        Else
            ckcBookInto(1).Value = vbUnchecked
        End If
    End If
    If rbcCSchdPromo(0).Value = True Then
        ckcBookInto(0).Value = vbUnchecked
        ckcBookInto(0).Enabled = False
    Else
        ckcBookInto(0).Enabled = True
        If (ilValue And PROMOINTOCONTRACTAVAILS) = PROMOINTOCONTRACTAVAILS Then
            ckcBookInto(0).Value = vbChecked
        Else
            ckcBookInto(0).Value = vbUnchecked
        End If
    End If

    ilValue = Asc(tgSpf.sUsingFeatures4)  'Option Fields in Orders/Proposals
    If (ilValue And LOCKBOXBYVEHICLE) = LOCKBOXBYVEHICLE Then
        ckcInv(3).Value = vbChecked
    Else
        ckcInv(3).Value = vbUnchecked
    End If
    If (ilValue And ALLOWMOVEONTODAY) = ALLOWMOVEONTODAY Then
        ckcSales(3).Value = vbChecked
    Else
        ckcSales(3) = vbUnchecked
    End If
    If (ilValue And CHGBILLEDPRICE) = CHGBILLEDPRICE Then
        ckcOptionFields(10).Value = vbChecked
    Else
        ckcOptionFields(10) = vbUnchecked
    End If
    If (ilValue And TAXBYUSA) = TAXBYUSA Then
        rbcTaxOnAirTime(1).Value = True
    ElseIf (ilValue And TAXBYCANADA) = TAXBYCANADA Then
        rbcTaxOnAirTime(2).Value = True
    Else
        rbcTaxOnAirTime(0).Value = True
    End If
    If (ilValue And INVSORTBYVEHICLE) = INVSORTBYVEHICLE Then
        rbcInvSortBy(1).Value = True
    Else
        rbcInvSortBy(0).Value = True
    End If

    ilValue = Asc(tgSpf.sUsingFeatures5)  'Option Fields in Orders/Proposals
    If (ilValue And REMOTEEXPORT) = REMOTEEXPORT Then
        ckcRemoteExport.Value = vbChecked
    Else
        ckcRemoteExport.Value = vbUnchecked
    End If
    If (ilValue And REMOTEIMPORT) = REMOTEIMPORT Then
        ckcRemoteImport.Value = vbChecked
    Else
        ckcRemoteImport = vbUnchecked
    End If
    If (ilValue And COMBINEAIRNTR) = COMBINEAIRNTR Then
        ckcInv(1).Value = vbChecked
    Else
        ckcInv(1) = vbUnchecked
    End If
    If (ilValue And CNTRINVSORTRC) = CNTRINVSORTRC Then
        rbcCSortBy(0).Value = True
    ElseIf (ilValue And CNTRINVSORTLN) = CNTRINVSORTLN Then
        rbcCSortBy(2).Value = True
    Else
        rbcCSortBy(1).Value = True
    End If
    If (ilValue And STATIONINTERFACE) = STATIONINTERFACE Then
        ckcGUseAffSys(1).Value = vbChecked
    Else
        ckcGUseAffSys(1) = vbUnchecked
    End If
    If (ilValue And RADAR) = RADAR Then
        ckcGUseAffSys(2).Value = vbChecked
    Else
        ckcGUseAffSys(2) = vbUnchecked
    End If
    If (ilValue And SUPPRESSTIMEFORM1) = SUPPRESSTIMEFORM1 Then
        ckcSuppressTimeForm1.Value = vbChecked
    Else
        ckcSuppressTimeForm1.Value = vbUnchecked
    End If

    ilValue = Asc(tgSpf.sUsingFeatures6)  'Option Fields in Orders/Proposals
    If plcBB(0).Visible Then
        If (ilValue And BBNOTSEPARATELINE) = BBNOTSEPARATELINE Then
            rbcBBOnLine(1).Value = True
        Else
            rbcBBOnLine(0).Value = True
        End If
    Else
        rbcBBOnLine(0).Value = True
    End If
    If (ilValue And BBCLOSEST) = BBCLOSEST Then
        rbcBBType(1).Value = True
    Else
        rbcBBType(0).Value = True
    End If
    If (ilValue And INSTALLMENT) = INSTALLMENT Then
        ckcInstallment.Value = vbChecked
        If (ilValue And INSTALLMENTREVENUEEARNED) = INSTALLMENTREVENUEEARNED Then
            rbcInstRev(0).Value = True
        Else
            rbcInstRev(1).Value = True
        End If
    Else
        ckcInstallment = vbUnchecked
        rbcInstRev(1).Value = True
        rbcInstRev(0).Enabled = False
    End If
    If (ilValue And GETPAIDEXPORT) = GETPAIDEXPORT Then
        ckcGetPaidExport.Value = vbChecked
    Else
        ckcGetPaidExport.Value = vbUnchecked
    End If
    If (ilValue And DIGITALCONTENT) = DIGITALCONTENT Then
        ckcDigital.Value = vbChecked
    Else
        ckcDigital.Value = vbUnchecked
    End If
    If (ilValue And GUARBYGRIMP) = GUARBYGRIMP Then
        ckcOptionFields(11).Value = vbChecked
    Else
        ckcOptionFields(11).Value = vbUnchecked
    End If
'    If (ilValue And INVEXPORTPARAMETERS) = INVEXPORTPARAMETERS Then
'        ckcOptionFields(12).Value = vbChecked
'    Else
'        ckcOptionFields(12).Value = vbUnchecked
'    End If

    If (ilValue And INVEXPORTPARAMETERS) = INVEXPORTPARAMETERS Then
        ckcInvoiceExport.Value = vbChecked
    Else
        ckcInvoiceExport.Value = vbUnchecked
    End If
    If ckcInvoiceExport.Value = vbUnchecked Then
        ckcACodes(4).Enabled = False
        ckcACodes(4).Value = vbUnchecked
    Else
        ckcACodes(4).Enabled = True
    End If
    
    ilValue = Asc(tgSpf.sUsingFeatures7)
    If (ilValue And CSIBACKUP) = CSIBACKUP Then
        ckcCSIBackup.Value = vbChecked
    Else
        ckcCSIBackup.Value = vbUnchecked
    End If
    rbcCommBy(1).Value = True
    If (ilValue And BONUSCOMM) = BONUSCOMM Then
        rbcCommBy(0).Value = True
    End If
    rbcCommBy(6).Value = True
    If (ilValue And COMMFISCALYEAR) = COMMFISCALYEAR Then
        rbcCommBy(2).Value = True
    End If
    If (ilValue And EXPORTREVENUE) = EXPORTREVENUE Then
        ckcRevenueExport.Value = vbChecked
    Else
        ckcRevenueExport.Value = vbUnchecked
    End If
    If (ilValue And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT Then
        'udcSiteTabs.Automation(13) = vbChecked
        ckcXDSBy(0).Value = vbChecked
    Else
        'udcSiteTabs.Automation(13) = vbUnchecked
        ckcXDSBy(0).Value = vbUnchecked
    End If
    If (ilValue And WEGENEREXPORT) = WEGENEREXPORT Then
        ckcRegionMixLen.Enabled = False
        udcSiteTabs.Automation(14) = vbChecked
    End If
    If (ilValue And OLAEXPORT) = OLAEXPORT Then
        ckcRegionMixLen.Enabled = False
        udcSiteTabs.Automation(15) = vbChecked
    End If
    If ckcRegionMixLen.Enabled Then
        If ckcUsingSplitNetworks.Value = vbUnchecked Then
            ckcRegionMixLen.Value = vbUnchecked
            ckcRegionMixLen.Enabled = False
        Else
            If (ilValue And REGIONMIXLEN) = REGIONMIXLEN Then
                ckcRegionMixLen.Value = vbChecked
            Else
                ckcRegionMixLen.Value = vbUnchecked
            End If
        End If
    Else
        ckcRegionMixLen.Value = vbUnchecked
    End If

    ilValue = Asc(tgSpf.sUsingFeatures8)
    If (ilValue And LRMANDATORY) = LRMANDATORY Then
        ckcOverrideOptions(5).Value = vbChecked
    Else
        ckcOverrideOptions(5).Value = vbUnchecked
    End If
    If (ilValue And SHOWCMMTONDETAILPAGE) = SHOWCMMTONDETAILPAGE Then
        ckcCntr(0).Value = vbChecked
    Else
        ckcCntr(0).Value = vbUnchecked
    End If
    If (ilValue And ALLOWMSASPLITCOPY) = ALLOWMSASPLITCOPY Then
        ckcMetroSplitCopy.Value = vbChecked
    Else
        ckcMetroSplitCopy.Value = vbUnchecked
    End If
    If (ilValue And RIVENDELLEXPORT) = RIVENDELLEXPORT Then
        udcSiteTabs.Automation(17) = vbChecked
    Else
        udcSiteTabs.Automation(17) = vbUnchecked
    End If
    If (ilValue And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT Then
        'udcSiteTabs.Automation(18) = vbChecked
        ckcXDSBy(1).Value = vbChecked
    Else
        'udcSiteTabs.Automation(18) = vbUnchecked
        ckcXDSBy(1).Value = vbUnchecked
    End If
    If (ilValue And ISCIEXPORT) = ISCIEXPORT Then
        udcSiteTabs.Automation(19) = vbChecked
    Else
        udcSiteTabs.Automation(19) = vbUnchecked
    End If
    If (ilValue And PREFEEDDEF) = PREFEEDDEF Then
        ckcPrefeed.Value = vbChecked
    Else
        ckcPrefeed.Value = vbUnchecked
    End If
    If (ilValue And REPBYDT) = REPBYDT Then
        ckcInv(5).Value = vbChecked
    Else
        ckcInv(5).Value = vbUnchecked
    End If
    ilValue = Asc(tgSpf.sUsingFeatures9)
    If (ilValue And AFFILIATECRM) = AFFILIATECRM Then
        ckcAffiliateCRM.Value = vbChecked
    Else
        ckcAffiliateCRM.Value = vbUnchecked
    End If
    'If (ilValue And PC1STPOS) = PC1STPOS Then
    '    ckcPC1StPos.Value = vbChecked
    'Else
    '    ckcPC1StPos.Value = vbUnchecked
    'End If
    If (ilValue And PROPOSALXML) = PROPOSALXML Then
        ckcProposalXML.Value = vbChecked
    Else
        ckcProposalXML.Value = vbUnchecked
    End If
    If (ilValue And LIMITISCI) = LIMITISCI Then
        ckcCopy(3).Value = vbChecked
    Else
        ckcCopy(3).Value = vbUnchecked
    End If
'    If (ilValue And IDCRESTRICTION) = IDCRESTRICTION Then
'        ckcCopy(4).Value = vbChecked
'    Else
'        ckcCopy(4).Value = vbUnchecked
'    End If
    If (ilValue And WEEKLYBILL) = WEEKLYBILL Then
        ckcOptionFields(13).Value = vbChecked
    Else
        ckcOptionFields(13).Value = vbUnchecked
    End If
    If (ilValue And PRINTEDI) = PRINTEDI Then
        ckcInv(6).Value = vbChecked
    Else
        ckcInv(6).Value = vbUnchecked
    End If
    If (ilValue And WORDWRAPVEHICLE) = WORDWRAPVEHICLE Then
        ckcSales(12).Value = vbChecked
    Else
        ckcSales(12).Value = vbUnchecked
    End If

    'If tgSpf.sXSDAddAdvtToISCI = "Y" Then 'Add Advt/Prod to X-Digital ISCI
    '    ckcXDSBy(2).Value = vbChecked
    'Else
    '    ckcXDSBy(2).Value = vbUnchecked
    'End If
    ilValue = Asc(tgSpf.sUsingFeatures10)
    If (ilValue And ADDADVTTOISCI) = ADDADVTTOISCI Then
        ckcXDSBy(2).Value = vbChecked
    Else
        ckcXDSBy(2).Value = vbUnchecked
    End If
    If (ilValue And MIDNIGHTBASEDHOUR) = MIDNIGHTBASEDHOUR Then
        ckcXDSBy(3).Value = vbChecked
    Else
        ckcXDSBy(3).Value = vbUnchecked
    End If
    If (ilValue And PKGLNRATEONBR) = PKGLNRATEONBR Then
        ckcCntr(2).Value = vbChecked
    Else
        ckcCntr(2).Value = vbUnchecked
    End If
    If (ilValue And REPLACEDELWKWITHFILLS) = REPLACEDELWKWITHFILLS Then
        ckcCntr(3).Value = vbChecked
    Else
        ckcCntr(3).Value = vbUnchecked
    End If
    If (ilValue And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
        ckcVCreative.Value = vbChecked
    Else
        ckcVCreative.Value = vbUnchecked
    End If
    'If (ilValue And CONTRACTVERIFY) = CONTRACTVERIFY Then
    '    ckcCntr(4).Value = vbChecked
    'Else
    '    ckcCntr(4).Value = vbUnchecked
    'End If
    If (ilValue And WegenerIPump) = WegenerIPump Then 'Dalet)
        udcSiteTabs.Automation(13) = vbChecked
    Else
        udcSiteTabs.Automation(13) = vbUnchecked
    End If
    '9114
    If (ilValue And UNITIDBYASTCODEFORBREAK) = UNITIDBYASTCODEFORBREAK Then
        ckcXDSBy(ASTBREAK).Value = vbChecked
    Else
        ckcXDSBy(ASTBREAK).Value = vbUnchecked
    End If


    'Invoicing
    If tgSpf.sBLCycle = "C" Then 'Local billing cycle
        rbcBLCycle(1).Value = True
    ElseIf tgSpf.sBLCycle = "W" Then 'Local billing cycle
        rbcBLCycle(2).Value = True
    Else
        rbcBLCycle(0).Value = True
    End If
    If tgSpf.sBRCycle = "C" Then 'Regional billing cycle
        rbcBRCycle(1).Value = True
    ElseIf tgSpf.sBRCycle = "W" Then 'Regional billing cycle
        rbcBRCycle(2).Value = True
    Else
        rbcBRCycle(0).Value = True
    End If
    If tgSpf.sBNCycle = "C" Then 'National billing cycle
        rbcBNCycle(1).Value = True
    ElseIf tgSpf.sBNCycle = "W" Then 'National billing cycle
        rbcBNCycle(2).Value = True
    Else
        rbcBNCycle(0).Value = True
    End If
    edcBNo(0).Text = Trim$(Str$(tgSpf.lBLowestNo)) 'Lowest invoice #
    edcBNo(1).Text = Trim$(Str$(tgSpf.lBHighestNo))   'Highest invoice #
    edcBNo(2).Text = Trim$(Str$(tgSpf.lBNextNo)) 'Next number
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr
    If gValidDate(slStr) Then
        edcBBillDate(0).Text = slStr
    Else
        edcBBillDate(0).Text = ""
    End If
    gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), slStr
    If gValidDate(slStr) Then
        edcBBillDate(1).Text = slStr
    Else
        edcBBillDate(1).Text = ""
    End If
    gUnpackDate tgSpf.iRepPrintDate(0), tgSpf.iRepPrintDate(1), slStr
    If gValidDate(slStr) Then
        edcBBillDate(2).Text = slStr
    Else
        edcBBillDate(2).Text = ""
    End If
    gUnpackDate tmSaf.iBLastWeeklyDate(0), tmSaf.iBLastWeeklyDate(1), slStr
    If gValidDate(slStr) Then
        If gDateValue(slStr) <> gDateValue("1/1/1990") Then
            edcBBillDate(3).Text = slStr
        Else
            edcBBillDate(3).Text = ""
        End If
    Else
        edcBBillDate(3).Text = ""
    End If
    gUnpackDate tmSaf.iXDSLastImptDate(0), tmSaf.iXDSLastImptDate(1), slStr
    If gValidDate(slStr) Then
        If gDateValue(slStr) <> gDateValue("1/1/1970") Then
            edcExpt(2).Text = slStr
        Else
            edcExpt(2).Text = ""
        End If
    Else
        edcExpt(2).Text = ""
    End If
    If tgSpf.sBCombine = "N" Then 'Combine advertiser across vehicles on invoices
        rbcBCombine(1).Value = True
    Else
        rbcBCombine(0).Value = True
    End If
    '8-23-01
    If tgSpf.sInvVehSel = "Y" Then 'Allow selective vehicle invoicing
        ckcInv(2).Value = vbChecked
    Else
        ckcInv(2).Value = vbUnchecked
    End If
    edcBLogoSpaces(0).Text = Trim$(Str$(Val(tgSpf.sExport)))
    edcBLogoSpaces(1).Text = Trim$(Str$(Val(tgSpf.sImport)))

    edcInvExportId.Text = Trim$(tgSpf.sInvExportId)

    'If tgSpf.sBSepItem = "N" Then 'Place item bill on separate invoice
    '    rbcBSepItem(1).Value = True
    'Else
    '    rbcBSepItem(0).Value = True
    'End If
    If (tgSpf.sInvAirOrder = "S") Or (tgSpf.sInvAirOrder = "O") Then 'Show Ordered, Update Ordered
        If tgSpf.sBMissedDT = "A" Then 'Place item bill on separate invoice
            rbcBMissedDT(0).Value = True
        Else
            rbcBMissedDT(1).Value = True
        End If
        rbcBMissedDT(0).Enabled = True
        rbcBMissedDT(1).Enabled = True
    Else
        rbcBMissedDT(1).Value = True
        rbcBMissedDT(0).Enabled = False
        rbcBMissedDT(1).Enabled = False
    End If
    If tgSpf.sRepRptForm = "V" Then 'Print Rep Invoice
        rbcPrintRepInv(1).Value = True
    Else
        rbcPrintRepInv(0).Value = True
    End If
    If tgSpf.sPostCalAff = "C" Then 'Print Rep Invoice by calendar
        rbcPostRepAffidavit(1).Value = True
    ElseIf tgSpf.sPostCalAff = "W" Then 'Print Rep Invoice by week
        rbcPostRepAffidavit(2).Value = True
    ElseIf tgSpf.sPostCalAff = "N" Then 'None.  Print Rep Invoice by date/time moved to Features8
        rbcPostRepAffidavit(3).Value = True
    Else
        rbcPostRepAffidavit(0).Value = True 'by standard broadcast month
    End If
    If tgSpf.sBActDayCompl = "N" Then 'Using Remnant Contracts
        ckcInv(4).Value = vbUnchecked
    Else
        ckcInv(4).Value = vbChecked
    End If
    If tgSpf.sBLaserForm = "2" Then 'Laser Form: 1=Ordered, Aired & Recon; 2=Invoice & Affidavit; 3=Aired
        rbcBLaserForm(1).Value = True
    ElseIf tgSpf.sBLaserForm = "3" Then 'Laser Form: 1=Ordered, Aired & Recon; 2=Invoice & Affidavit; 3=Aired
        rbcBLaserForm(2).Value = True
    ElseIf tgSpf.sBLaserForm = "4" Then     '3-8-12 3-column aired
        rbcBLaserForm(3).Value = True
    Else
        rbcBLaserForm(0).Value = True
    End If
    'gPDNToStr tgSpf.sBTax(0), 2, slStr
    edcBTTax(0).Text = Trim$(tgSpf.sTax1Text)
    'gPDNToStr tgSpf.sBTax(1), 2, slStr
    edcBTTax(1).Text = Trim$(tgSpf.sTax2Text)
    edcSales(7).Text = Trim$(Str$(tgSpf.iNoMnthNewBus))
    edcSales(8).Text = Trim$(Str$(tgSpf.iNoMnthNewIsNew))
    If tgSpf.sNewBusYearType = "C" Then
        rbcNewBusYear(0).Value = True
        edcSales(7).Enabled = False
        edcSales(8).Enabled = False
    Else
        rbcNewBusYear(1).Value = True
        edcSales(7).Enabled = True
        edcSales(8).Enabled = True
    End If
    edcEDI(0).Text = Trim$(tgSpf.sEDICallLetter)
    edcEDI(1).Text = Trim$(tgSpf.sEDIMediaType)
    edcEDI(2).Text = Trim$(tgSpf.sEDIBand)
    If tgSpf.sDefFillInv = "N" Then
        rbcDefFillInv(1).Value = True
    Else
        rbcDefFillInv(0).Value = True
    End If
    If tgSpf.sBOrderDPShow = "B" Then
        rbcBOrderDPShow(2).Value = True
    ElseIf tgSpf.sBOrderDPShow = "N" Then
        rbcBOrderDPShow(0).Value = True
    Else
        rbcBOrderDPShow(1).Value = True
    End If
    If tgSpf.sInvSpotTimeZone = "E" Then
        rbcInvSpotTimeZone(0).Value = True
    ElseIf tgSpf.sInvSpotTimeZone = "C" Then
        rbcInvSpotTimeZone(1).Value = True
    ElseIf tgSpf.sInvSpotTimeZone = "M" Then
        rbcInvSpotTimeZone(2).Value = True
    ElseIf tgSpf.sInvSpotTimeZone = "P" Then
        rbcInvSpotTimeZone(3).Value = True
    Else
        rbcInvSpotTimeZone(4).Value = True
    End If
        
    'Accounting
    'edcRCorp1.Text = Trim$(Str$(tgSpf.iRCorp(0))) 'Corporation calendar number of week (Jan,..)
    'edcRCorp2.Text = Trim$(Str$(tgSpf.iRCorp(1))) 'Corporation calendar number of week (Feb,..)
    'edcRCorp3.Text = Trim$(Str$(tgSpf.iRCorp(2))) 'Corporation calendar number of week (Mar,..)
    'If tgSpf.sRYEnd = "L" Then
    '    rbcREnd(0).Value = True
    'Else
    '    rbcREnd(1).Value = True
    'End If
    If tgSpf.sRUseCorpCal = "Y" Then
        rbcRCorpCal(1).Value = True
    Else
        rbcRCorpCal(0).Value = True
    End If
    gUnpackDate tgSpf.iRLastPurgedDate(0), tgSpf.iRLastPurgedDate(1), slStr
    lacLastPurgedDate(1).Caption = " " & slStr
    gUnpackDate tgSpf.iRLastPay(0), tgSpf.iRLastPay(1), slStr
    smRLastPayCaption = " " & slStr 'Date last payment
    'plcRLastPay_Paint
    lacAG(19).Caption = slStr
    If tgSpf.sRCurrAmt = "N" Then 'Include current amount in computing Credit limit
        rbcRCurrAmt(1).Value = True
    Else
        rbcRCurrAmt(0).Value = True
    End If
    If tgSpf.sRUnbilled = "N" Then 'Include unbilled amount in computing Credit limit
        rbcRUnbilled(1).Value = True
    Else
        rbcRUnbilled(0).Value = True
    End If
    'If tgSpf.sRNewCntr = "W" Then 'Include 1st week amount in computing Credit limit
    '    rbcRNewCntr(0).Value = True
    'ElseIf tgSpf.sRNewCntr = "M" Then    'Include 1st month
    '    rbcRNewCntr(1).Value = True
    'ElseIf tgSpf.sRNewCntr = "A" Then    'Include All
    '    rbcRNewCntr(2).Value = True
    'Else
    '    rbcRNewCntr(3).Value = True
    'End If
    If tgSpf.iRNoWks > 0 Then
        edcRNewCntr.Text = Trim$(Str$(tgSpf.iRNoWks))
    Else
        edcRNewCntr.Text = ""
    End If
    'gPDNToStr tgSpf.sRPctCredit, 0, slStr
    slStr = gIntToStrDec(tgSpf.iRPctCredit, 0)
    If tgSpf.sRRP = "C" Then
        rbcRRP(1).Value = True
    ElseIf tgSpf.sRRP = "F" Then
        rbcRRP(2).Value = True
    Else
        rbcRRP(0).Value = True
    End If
    edcRPctCredit.Text = slStr
    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slStr
    If gValidDate(slStr) Then
        edcRPRP.Text = slStr
    Else
        edcRPRP.Text = ""
    End If
    gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slStr
    If gValidDate(slStr) Then
        edcRCRP.Text = slStr
    Else
        edcRCRP.Text = ""
    End If
    gUnpackDate tgSpf.iRNRP(0), tgSpf.iRNRP(1), slStr
    If gValidDate(slStr) Then
        edcRNRP.Text = slStr
    Else
        edcRNRP.Text = ""
    End If
    gPDNToStr tgSpf.sRB, 2, slStr
    If edcRPRP.Text <> "" Then
        edcRB.Text = slStr
    Else
        edcRB.Text = ""
    End If
    edcRCollectContact.Text = Trim$(tgSpf.sRCollectContact)
    gSetPhoneNo tgSpf.sRCollectPhoneNo, mkcRCollectPhoneNo
    gUnpackDate tgSpf.iRCreditDate(0), tgSpf.iRCreditDate(1), slStr
    If gValidDate(slStr) Then
        edcRCreditDate.Text = slStr
    Else
        edcRCreditDate.Text = ""
    End If
    'For ilLoop = 0 To UBound(smRGL) Step 1
    '    smRGL(ilLoop) = Trim$(tgSpf.sRGLSuffix(ilLoop))
    'Next ilLoop
    'For ilLoop = 0 To UBound(smRName) Step 1
    '    smRName(ilLoop) = Trim$(tgSpf.sRName(ilLoop))
    '    smRTsfx(ilLoop) = Trim$(tgSpf.sRTsfx(ilLoop))
    '    smRAsfx(ilLoop) = Trim$(tgSpf.sRAsfx(ilLoop))
    'Next ilLoop

    '1-22-04 Vehicle Group default for Reconciliation reports
    If tgSpf.iReconcGroupNo > 6 Then
        tgSpf.iReconcGroupNo = 0
    End If
    cbcReconGroup.ListIndex = tgSpf.iReconcGroupNo

    If tgSpf.sRUseTrade = "Y" Then 'Using Trade
        ckcRUseTMP(0).Value = vbChecked
    Else
        ckcRUseTMP(0).Value = vbUnchecked
    End If
    If tgSpf.sRUseMerch = "Y" Then 'Using Merchandising
        ckcRUseTMP(1).Value = vbChecked
    Else
        ckcRUseTMP(1).Value = vbUnchecked
    End If
    If tgSpf.sRUsePromo = "Y" Then 'Using Promotion
        ckcRUseTMP(2).Value = vbChecked
    Else
        ckcRUseTMP(2).Value = vbUnchecked
    End If
    gUnpackDate tgSpf.iBarterLPD(0), tgSpf.iBarterLPD(1), slStr
    If gValidDate(slStr) Then
        edcBarterLPD.Text = slStr
    Else
        edcBarterLPD.Text = ""
    End If
    If tmSaf.sCreditLimitMsg = "C" Then  'Cutoff
        ckcRUseTMP(3).Value = vbChecked
    Else
        ckcRUseTMP(3).Value = vbUnchecked
    End If

    tmCxfSrchKey.lCode = tgSpf.lBCxfDisclaimer
    lmBCxfDisclaimer = tgSpf.lBCxfDisclaimer    '2-20-03
    If tgSpf.lBCxfDisclaimer <> 0 Then
        tmCxf.sComment = ""
        imCxfRecLen = Len(tmCxf)
        ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            tmCxf.lCode = 0
            'tmCxf.iStrLen = 0
            tmCxf.sComment = ""
            lmBCxfDisclaimer = 0            '2-20-03
        End If
    Else
        tmCxf.lCode = 0
        'tmCxf.iStrLen = 0
        tmCxf.sComment = ""
    End If
    'If tmCxf.iStrLen > 0 Then
    '    edcComment(2).Text = Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
    'Else
    '    edcComment(2).Text = ""
    'End If
    edcComment(2).Text = gStripChr0(tmCxf.sComment)

    tmCxfSrchKey.lCode = tgSpf.lCxfContrComment          '2-12-03 Contract comments
    lmCxfContrComment = tgSpf.lCxfContrComment      '2-20-03
    If tgSpf.lCxfContrComment <> 0 Then
        tmCxf.sComment = ""
        imCxfRecLen = Len(tmCxf)
        ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            tmCxf.lCode = 0
            'tmCxf.iStrLen = 0
            tmCxf.sComment = ""
            lmCxfContrComment = 0           '2-20-03
        End If
    Else
        tmCxf.lCode = 0
        'tmCxf.iStrLen = 0
        tmCxf.sComment = ""
    End If
    'If tmCxf.iStrLen > 0 Then
    '    edcComment(0).Text = Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
    'Else
    '    edcComment(0).Text = ""
    'End If
    edcComment(0).Text = gStripChr0(tmCxf.sComment)

    tmCxfSrchKey.lCode = tgSpf.lCxfInsertComment          '2-12-03 Insertion comments
    lmCxfInsertComment = tgSpf.lCxfInsertComment        '2-20-03
    If tgSpf.lCxfInsertComment <> 0 Then
        tmCxf.sComment = ""
        imCxfRecLen = Len(tmCxf)
        ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            tmCxf.lCode = 0
            'tmCxf.iStrLen = 0
            tmCxf.sComment = ""
            lmCxfInsertComment = 0              '2-20-03
        End If
    Else
        tmCxf.lCode = 0
        'tmCxf.iStrLen = 0
        tmCxf.sComment = ""
    End If
    'If tmCxf.iStrLen > 0 Then
    '    edcComment(1).Text = Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
    'Else
    '    edcComment(1).Text = ""
    'End If
    edcComment(1).Text = gStripChr0(tmCxf.sComment)

    tmCxfSrchKey.lCode = tgSpf.lCxfDemoEst          '2-12-03 Contract comments
    lmCxfEstComment = tgSpf.lCxfDemoEst     '2-20-03
    If tgSpf.lCxfDemoEst <> 0 Then
        tmCxf.sComment = ""
        imCxfRecLen = Len(tmCxf)
        ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            tmCxf.lCode = 0
            'tmCxf.iStrLen = 0
            tmCxf.sComment = ""
            lmCxfEstComment = 0           '2-20-03
        End If
    Else
        tmCxf.lCode = 0
        'tmCxf.iStrLen = 0
        tmCxf.sComment = ""
    End If
    'If tmCxf.iStrLen > 0 Then
    '    edcComment(3).Text = Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
    'Else
    '    edcComment(3).Text = ""
    'End If
    edcComment(3).Text = gStripChr0(tmCxf.sComment)

    tmCxfSrchKey.lCode = tmSaf.lStatementComment
    lmCxfStatementComment = tmSaf.lStatementComment
    If tmSaf.lStatementComment <> 0 Then
        tmCxf.sComment = ""
        imCxfRecLen = Len(tmCxf)
        ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            tmCxf.lCode = 0
            tmCxf.sComment = ""
            lmCxfStatementComment = 0           '2-20-03
        End If
    Else
        tmCxf.lCode = 0
        tmCxf.sComment = ""
    End If
    edcComment(4).Text = gStripChr0(tmCxf.sComment)

    tmCxfSrchKey.lCode = tmSaf.lCitationComment
    lmCxfCitationComment = tmSaf.lCitationComment
    If tmSaf.lCitationComment <> 0 Then
        tmCxf.sComment = ""
        imCxfRecLen = Len(tmCxf)
        ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            tmCxf.lCode = 0
            tmCxf.sComment = ""
            lmCxfCitationComment = 0           '2-20-03
        End If
    Else
        tmCxf.lCode = 0
        tmCxf.sComment = ""
    End If
    edcComment(5).Text = gStripChr0(tmCxf.sComment)

    'Automation
    cbcUser(0).ListIndex = -1
    cbcUser(1).ListIndex = -1
    If tgSpf.sSystemType = "R" Then
        cbcUser(0).ListIndex = 0
        For ilLoop = 0 To cbcUser(0).ListCount - 1
            ilValue = cbcUser(0).ItemData(ilLoop)
            If ilValue = tgSpf.iPriUrfCode Then
                cbcUser(0).ListIndex = ilLoop
                cbcUser(0).Text = cbcUser(0).List(ilLoop)
                Exit For
            End If
        Next ilLoop
        For ilLoop = 0 To cbcUser(1).ListCount - 1
            cbcUser(1).ListIndex = 0
            ilValue = cbcUser(1).ItemData(ilLoop)
            If ilValue = tgSpf.iSecUrfCode Then
                cbcUser(1).ListIndex = ilLoop
                cbcUser(1).Text = cbcUser(1).List(ilLoop)
                Exit For
            End If
        Next ilLoop
    End If

    If tgSpf.sCmmlSchStatus = "A" Then 'Using Trade
        ckcCmmlSchStatus.Value = vbChecked
    Else
        ckcCmmlSchStatus.Value = vbUnchecked
    End If
    'Price Levels
    If (tmSaf.iCode = 0) Or (tmSaf.lLowPrice = 0) Or (tmSaf.lHighPrice = 0) Then
        edcSchedule(0).Text = ""
        edcSchedule(1).Text = ""
        For ilLoop = LBound(lmSSave) To UBound(lmSSave) Step 1
            lmSSave(ilLoop) = 0
        Next ilLoop
    Else
        edcSchedule(0).Text = Trim$(Str$(tmSaf.lLowPrice))
        edcSchedule(1).Text = Trim$(Str$(tmSaf.lHighPrice))
        'lmSSave(LBound(lmSSave)) = tmSaf.lLowPrice
        lmSSave(LBONE) = tmSaf.lLowPrice
        lmSSave(UBound(lmSSave) - 1) = tmSaf.lHighPrice
        'For ilLoop = LBound(lmSSave) + 1 To UBound(lmSSave) - 2 Step 1
        ilLevel = LBound(tmSaf.lLevelToPrice)
        For ilLoop = LBONE + 1 To UBound(lmSSave) - 2 Step 1
            lmSSave(ilLoop) = tmSaf.lLevelToPrice(ilLevel)
            ilLevel = ilLevel + 1
        Next ilLoop
    End If
    If ckcOverrideOptions(3).Value = vbChecked Then  'Option Fields (Right to Left):0=Allocation %;1=Acquisition Cost; 2=1st Position; 3=Preferred Days/Times; 4=Solo Avails
        edcLnOverride(0).Text = Trim$(Str$(tmSaf.iPreferredPct))
        edcLnOverride(0).Enabled = True
    Else
        If (tmSaf.iPreferredPct >= 0) And (tmSaf.iPreferredPct <= 100) Then
            edcLnOverride(0).Text = Trim$(Str$(tmSaf.iPreferredPct))
        Else
            edcLnOverride(0).Text = ""
        End If
        edcLnOverride(0).Enabled = False
    End If
    If (tmSaf.iWk1stSoloIndex > 0) And (tmSaf.iWk1stSoloIndex < 100) Then
        edcLnOverride(1).Text = gIntToStrDec(tmSaf.iWk1stSoloIndex, 2)
    Else
        edcLnOverride(1).Text = ""
    End If
    If tmSaf.sInvISCIForm = "L" Then
        rbcInvISCIForm(1).Value = True
    ElseIf tmSaf.sInvISCIForm = "W" Then
        rbcInvISCIForm(2).Value = True
    Else
        rbcInvISCIForm(0).Value = True
    End If

    'Great Plain
    edcGP(0).Text = Trim$(Str$(tmSaf.lGPBatchNo))
    If tmSaf.lGPBatchNo <= 0 Then
        edcGP(0).BackColor = &HFFFF00
        edcGP(0).Enabled = True
    Else
        edcGP(0).BackColor = LIGHTYELLOW
        edcGP(0).Enabled = False
    End If
    edcGP(1).Text = Trim$(tmSaf.sGPPrefixChar)
    edcGP(2).Text = Trim$(tmSaf.sGPCustomerNo)
    'Copy
    If tgSpf.sUseCartNo = "N" Then 'Using cart number
'        ckcCUseCartNo.Value = vbUnchecked
        rbcCUseCartNo(1).Value = True
    ElseIf tgSpf.sUseCartNo = "B" Then
        rbcCUseCartNo(2).Value = True
    Else
'        ckcCUseCartNo.Value = vbChecked
        rbcCUseCartNo(0).Value = True
    End If
    If tgSpf.sTapeShowForm = "C" Then 'Tape Show Form
        rbcTapeShowForm(1).Value = True
    Else
        rbcTapeShowForm(0).Value = True
    End If
    If tgSpf.sCBlackoutLog = "Y" Then 'Using Blackouts on Logs
        ckcCopy(0).Value = vbChecked
    Else
        ckcCopy(0).Value = vbUnchecked
    End If
    If tgSpf.sCDefLogCopy = "Y" Then 'Default Log Copy (Y=On; N=Off)
        rbcDefLogCopy(0).Value = True
    Else
        rbcDefLogCopy(1).Value = True
    End If
    ilValue = Asc(tgSpf.sMOFCopyAssign)  'Mg Copy Assignment
    If (ilValue And MGORIGVEHONLY) = MGORIGVEHONLY Then 'Allocation %
        rbcMGCopyAssign(0).Value = True
    ElseIf (ilValue And MGSCHVEHONLY) = MGSCHVEHONLY Then 'Allocation %
        rbcMGCopyAssign(1).Value = True
    Else
        rbcMGCopyAssign(2).Value = True
    End If
    If (ilValue And FILLORIGVEHONLY) = FILLORIGVEHONLY Then 'Allocation %
        rbcFillCopyAssign(0).Value = True
    ElseIf (ilValue And FILLSCHVEHONLY) = FILLSCHVEHONLY Then 'Allocation %
        rbcFillCopyAssign(1).Value = True
    Else
        rbcFillCopyAssign(2).Value = True
    End If
    If (ilValue And MGRULESINCOPY) = MGRULESINCOPY Then
        rbcMGRules(1).Value = True
    Else
        rbcMGRules(0).Value = True
    End If
    If (ilValue And RSCHCUSTDEMO) = RSCHCUSTDEMO Then
        udcSiteTabs.Research(41) = vbChecked
    Else
        udcSiteTabs.Research(41) = vbUnchecked
    End If

    'Sports
    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    If (ilValue And USINGSPORTS) = USINGSPORTS Then 'Using Sports
        ckcUsingSports.Value = vbChecked
        If (ilValue And PREEMPTREGPROG) = PREEMPTREGPROG Then
            udcSiteTabs.Sports(0) = vbChecked
        Else
            udcSiteTabs.Sports(0) = vbUnchecked
        End If
        If (ilValue And USINGFEED) = USINGFEED Then
            udcSiteTabs.Sports(1) = vbChecked
        Else
            udcSiteTabs.Sports(1) = vbUnchecked
        End If
        If (ilValue And USINGLANG) = USINGLANG Then
            udcSiteTabs.Sports(2) = vbChecked
        Else
            udcSiteTabs.Sports(2) = vbUnchecked
        End If
        udcSiteTabs.EventTitle(1) = Trim(tmSaf.sEventTitle1)
        udcSiteTabs.EventTitle(2) = Trim$(tmSaf.sEventTitle2)
        udcSiteTabs.EventSubtotalTitle(1) = Trim(tmSaf.sEventSubtotal1)
        udcSiteTabs.EventSubtotalTitle(2) = Trim$(tmSaf.sEventSubtotal2)
    Else
        ckcUsingSports.Value = vbUnchecked
        udcSiteTabs.Sports(0) = vbUnchecked
        udcSiteTabs.Sports(1) = vbUnchecked
        udcSiteTabs.Sports(2) = vbUnchecked
        udcSiteTabs.EventTitle(1) = ""
        udcSiteTabs.EventTitle(2) = ""
        udcSiteTabs.EventSubtotalTitle(1) = ""
        udcSiteTabs.EventSubtotalTitle(2) = ""
    End If
    If ckcUsingSpecialResearch.Value = vbChecked Then
        edcComment(3).Visible = True
        lacComment(3).Caption = "Research Estimate Comment"
    ElseIf udcSiteTabs.Research(31) = vbChecked Then
        edcComment(3).Visible = True
        lacComment(3).Caption = "Research Override Comment"
    Else
        edcComment(3).Visible = False
        lacComment(3).Caption = ""
    End If

    'Set Tax fields
    If (rbcTaxOnAirTime(0).Value = True) Then
        edcBTTax(0).Text = ""
        edcBTTax(1).Text = ""
        edcBTTax(0).Enabled = False
        edcBTTax(1).Enabled = False
        ckcTaxOn(0).Enabled = False
        ckcTaxOn(1).Enabled = False
        ckcTaxOn(0).Value = vbUnchecked
        ckcTaxOn(1).Value = vbUnchecked
    Else
        edcBTTax(0).Enabled = True
        edcBTTax(1).Enabled = True
        ckcTaxOn(0).Enabled = True
        ckcTaxOn(1).Enabled = True
    End If

    If ckcInstallment.Value = vbUnchecked Then
        rbcInstRev(1).Value = True
        rbcInstRev(0).Enabled = False
        ckcInv(1).Enabled = True
    Else
        rbcInstRev(0).Enabled = True
        ckcInv(1).Value = vbChecked
        ckcInv(1).Enabled = False
    End If

    'E-Mail
    '6/28/11 Dan lost from address
    udcSiteTabs.Email(1) = Trim$(tgSite.sEmailHost)
    udcSiteTabs.Email(2) = Trim$(tgSite.sEmailAcctName)
    udcSiteTabs.Email(3) = Trim$(tgSite.sEmailPassword)
    udcSiteTabs.Email(4) = Trim$(Str$(tgSite.iEmailPort))
    If Trim$(tgSite.sEmailFromName) = "1" Or Trim$(tgSite.sEmailFromName) = "0" Then
        udcSiteTabs.Email(5) = Trim$(tgSite.sEmailFromName)
    End If
    'udcSiteTabs.Email(5) = Trim$(tgSite.sEmailFromName)
   ' udcSiteTabs.Email(6) = Trim$(tgSite.sEmailFromAddress)

    If (ckcRN_Net.Value = vbUnchecked) And (ckcRN_Rep.Value = vbUnchecked) Then
        udcSiteTabs.Action 6, 0
    Else
        udcSiteTabs.RepNet(1) = Trim$(tgNrf.sDBID)
        udcSiteTabs.RepNet(2) = Trim$(tgNrf.sFTPUserID)
        udcSiteTabs.RepNet(3) = Trim$(tgNrf.sFTPUserPW)
        udcSiteTabs.RepNet(4) = Trim$(Str$(tgNrf.iFTPPort))
        udcSiteTabs.RepNet(5) = Trim$(tgNrf.sFTPAddress)
        udcSiteTabs.RepNet(6) = Trim$(tgNrf.sFTPImportDir)
        udcSiteTabs.RepNet(7) = Trim$(tgNrf.sFTPExportDir)
        udcSiteTabs.RepNet(8) = Trim$(tgNrf.sIISRootURL)
        udcSiteTabs.RepNet(9) = Trim$(tgNrf.sIISRegSection)
    End If
    
    udcSiteTabs.Wegener = Trim$(tgSpf.sWegenerGroupChar)
    udcSiteTabs.WegenerIPump = Trim$(tmSaf.sIPumpZone)
    gUnpackDate tmSaf.iVCreativeDate(0), tmSaf.iVCreativeDate(1), slStr
    If gValidDate(slStr) Then
        If gDateValue(slStr) <> gDateValue("1/1/1990") Then
            edcSchedule(2).Text = slStr
        Else
            edcSchedule(2).Text = ""
        End If
    Else
        edcSchedule(2).Text = ""
    End If
    
    If tmSaf.sGenAutoFileWOSpt = "Y" Then
        udcSiteTabs.Automation(18) = vbChecked
    Else
        udcSiteTabs.Automation(18) = vbUnchecked
    End If
    
    If tmSaf.sXMidSpotsBill = "A" Then
        ckcInv(7).Value = vbChecked
    Else
        ckcInv(7).Value = vbUnchecked
    End If
    If tmSaf.sHideDemoOnBR = "Y" Then
        udcSiteTabs.Research(42) = vbChecked
    Else
        udcSiteTabs.Research(42) = vbUnchecked
    End If
    If tmSaf.sAudByPackage = "Y" Then
        udcSiteTabs.Research(43) = vbChecked
    Else
        udcSiteTabs.Research(43) = vbUnchecked
    End If
    edcExpt(1).Text = UCase$(tmSaf.sXDSHeadEndZone)
    '7942
    smHeadEndZoneChange = edcExpt(1).Text
    ckcCopy(4).Value = vbUnchecked
    If tmSaf.sSyncCopyInRot = "Y" Then
        ckcCopy(4).Value = vbChecked
    End If
    
    '4/15/19: this code must be after the controls are set
    If (ckcUsingLiveCopy.Value = vbUnchecked) And (ckcCopy(1).Value = vbUnchecked) Then
        'ckcCopy(5).Enabled = True
        'If tmSaf.sExcludeAudioTypeR = "Y" Then
        '    ckcCopy(5).Value = vbChecked
        'Else
            ckcCopy(5).Value = vbUnchecked
        'End If
        ckcCopy(6).Enabled = False
        ckcCopy(6).Value = vbUnchecked
        ckcCopy(7).Enabled = False
        ckcCopy(7).Value = vbUnchecked
        ckcCopy(8).Enabled = False
        ckcCopy(8).Value = vbUnchecked
        ckcCopy(9).Enabled = False
        ckcCopy(9).Value = vbUnchecked
        ckcCopy(10).Enabled = False
        ckcCopy(10).Value = vbUnchecked
    ElseIf (ckcUsingLiveCopy.Value = vbUnchecked) Then
        'ckcCopy(5).Enabled = True
        'If tmSaf.sExcludeAudioTypeR = "Y" Then
        '    ckcCopy(5).Value = vbChecked
        'Else
            ckcCopy(5).Value = vbUnchecked
        'End If
        ckcCopy(6).Enabled = False
        ckcCopy(6).Value = vbChecked
        ckcCopy(7).Enabled = False
        ckcCopy(7).Value = vbChecked
        ckcCopy(8).Enabled = True
        If tmSaf.sExcludeAudioTypeS = "Y" Then
            ckcCopy(8).Value = vbChecked
        Else
            ckcCopy(8).Value = vbUnchecked
        End If
        ckcCopy(9).Enabled = False
        ckcCopy(9).Value = vbChecked
        ckcCopy(10).Enabled = False
        ckcCopy(10).Value = vbChecked
    ElseIf (ckcCopy(1).Value = vbUnchecked) Then
        'ckcCopy(5).Enabled = True
        'If tmSaf.sExcludeAudioTypeR = "Y" Then
        '    ckcCopy(5).Value = vbChecked
        'Else
            ckcCopy(5).Value = vbUnchecked
        'End If
        ckcCopy(6).Enabled = True
        If tmSaf.sExcludeAudioTypeL = "Y" Then
            ckcCopy(6).Value = vbChecked
        Else
            ckcCopy(6).Value = vbUnchecked
        End If
        ckcCopy(7).Enabled = False
        ckcCopy(7).Value = vbChecked
        ckcCopy(8).Enabled = False
        ckcCopy(8).Value = vbChecked
        ckcCopy(9).Enabled = True
        If tmSaf.sExcludeAudioTypeP = "Y" Then
            ckcCopy(9).Value = vbChecked
        Else
            ckcCopy(9).Value = vbUnchecked
        End If
        ckcCopy(10).Enabled = False
        ckcCopy(10).Value = vbChecked
    Else
        'ckcCopy(5).Enabled = True
        'If tmSaf.sExcludeAudioTypeR = "Y" Then
        '    ckcCopy(5).Value = vbChecked
        'Else
            ckcCopy(5).Value = vbUnchecked
        'End If
        ckcCopy(6).Enabled = True
        If tmSaf.sExcludeAudioTypeL = "Y" Then
            ckcCopy(6).Value = vbChecked
        Else
            ckcCopy(6).Value = vbUnchecked
        End If
        ckcCopy(7).Enabled = True
        If tmSaf.sExcludeAudioTypeM = "Y" Then
            ckcCopy(7).Value = vbChecked
        Else
            ckcCopy(7).Value = vbUnchecked
        End If
        ckcCopy(8).Enabled = True
        If tmSaf.sExcludeAudioTypeS = "Y" Then
            ckcCopy(8).Value = vbChecked
        Else
            ckcCopy(8).Value = vbUnchecked
        End If
        ckcCopy(9).Enabled = True
        If tmSaf.sExcludeAudioTypeP = "Y" Then
            ckcCopy(9).Value = vbChecked
        Else
            ckcCopy(9).Value = vbUnchecked
        End If
        ckcCopy(10).Enabled = True
        If tmSaf.sExcludeAudioTypeQ = "Y" Then
            ckcCopy(10).Value = vbChecked
        Else
            ckcCopy(10).Value = vbUnchecked
        End If
    End If
    If ckcInvoiceExport.Value = vbChecked Then
        edcExpt(3).Text = tmSaf.sInvExpDelimiter
    Else
        edcExpt(3).Text = ","
    End If
    
    'TTP 10205 - 6/21/21 - JW - Get SPFX Extended Site Features : spfxInvExpFeature - Audacy WO Invoice Export
    edcIE(0).Text = Trim$(tgSpfx.sInvExpProperty)
    edcIE(1).Text = Trim$(tgSpfx.sInvExpPrefix)
    edcIE(2).Text = Trim$(tgSpfx.sInvExpBillGroup)
    
    'TTP 10626 JJB 2023-01-10
    edcSageIE(SAGE_IE_TERMS).Text = Trim$(tgSpfx.sSageTerm)
    edcSageIE(SAGE_IE_ACCOUNT).Text = Trim$(tgSpfx.sSageAccount)
    edcSageIE(SAGE_IE_LOCATIONID).Text = Trim$(tgSpfx.sSageLocation)
    edcSageIE(SAGE_IE_DEPTID).Text = Trim$(tgSpfx.sSageDept)
    ''''''''''''''''''''''''
    
   'SOW Megaphone Phase 1
   ckcDaylightSavings.Value = IIF(tgSpfx.iIntFeature = 0, vbUnchecked, vbChecked)
    
    '8/17/21 - JW - TTP 10233 - Audacy: line summary export
    'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
'    Select Case tgSpfx.iInvExpFeature
'        Case 0 'All Disabled
'            ckcWOInvoiceExport.Value = vbUnchecked
'            ckcCntrLineExport.Value = vbUnchecked
'        Case 1 'WO Invoice Export' Enabled
'            ckcWOInvoiceExport.Value = vbChecked
'        Case 2 'WO InvLine Export' Enabled
'            ckcCntrLineExport.Value = vbChecked
'        Case 3 'both 'WO Invoice Export' AND 'WO InvLine Export' Enabled
'            ckcWOInvoiceExport.Value = vbChecked
'            ckcCntrLineExport.Value = vbChecked
'    End Select
    ilValue = tgSpfx.iInvExpFeature
    If (ilValue And INVEXP_AUDACYWO) = INVEXP_AUDACYWO Then
        ckcWOInvoiceExport.Value = vbChecked
    Else
        ckcWOInvoiceExport.Value = vbUnchecked
    End If
    If (ilValue And INVEXP_AUDACYLINE) = INVEXP_AUDACYLINE Then
        ckcCntrLineExport.Value = vbChecked
    Else
        ckcCntrLineExport.Value = vbUnchecked
    End If
    If (ilValue And INVEXP_SELECTIVEEMAIL) = INVEXP_SELECTIVEEMAIL Then
        ckcInv(12).Value = vbChecked
    Else
        ckcInv(12).Value = vbUnchecked
    End If
    
    
    
    imIgnoreClickEvent = False
    udcSiteTabs.Action 7, 1
    imChangesOccured = 0
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

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload SiteOpt
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcLevelSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcLevelSTab.hWnd Then
        Exit Sub
    End If
    mSSetShow imSBoxNo
    ilBox = imSBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imSRowNo = 2
                ilBox = LEVEL2INDEX
                imSBoxNo = ilBox
                mSEnableBox ilBox
                Exit Sub
            Case LEVEL2INDEX 'Time (first control within header)
                imSBoxNo = -1
                imSRowNo = -1
                cmcCommand(0).SetFocus
            Case Else
                ilBox = imSBoxNo - 1
                ilFound = True
        End Select
    Loop While Not ilFound
    imSBoxNo = ilBox
    mSEnableBox ilBox
End Sub

Private Sub pbcLevelTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcLevelTab.hWnd Then
        Exit Sub
    End If
    mSSetShow imSBoxNo
    ilBox = imSBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imSRowNo = 2
                ilBox = LEVEL14INDEX
                imSBoxNo = ilBox
                mSEnableBox ilBox
                Exit Sub
            Case LEVEL14INDEX 'Time (first control within header)
                imSBoxNo = -1
                imSRowNo = -1
                cmcCommand(0).SetFocus
                Exit Sub
            Case Else
                ilBox = imSBoxNo + 1
                ilFound = True
        End Select
    Loop While Not ilFound
    imSBoxNo = ilBox
    mSEnableBox ilBox
End Sub

Private Sub pbcSchedule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim ilBox As Integer

    If (Trim$(edcSchedule(0).Text) = "") Or (Trim$(edcSchedule(1).Text) = "") Then
        pbcSchedule.Cls
        Exit Sub
    End If
    For ilRow = 2 To 2 Step 1
        For ilBox = imLBSCtrls To UBound(tmSCtrls) Step 1
            If (X >= tmSCtrls(ilBox).fBoxX) And (X <= (tmSCtrls(ilBox).fBoxX + tmSCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmSCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmSCtrls(ilBox).fBoxY + tmSCtrls(ilBox).fBoxH)) Then
                    If ilBox = LEVEL15INDEX Then
                        Beep
                        Exit Sub
                    End If
                    ilRowNo = ilRow
                    mSSetShow imSBoxNo
                    imSBoxNo = ilBox
                    imSRowNo = 2
                    mSEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mSSetFocus imSBoxNo
End Sub

Private Sub pbcSchedule_Paint()
    Dim ilRow As Integer
    Dim ilBox As Integer
    Dim slStr As String
    Dim slLow As String
    Dim slHigh As String

    slLow = Trim$(edcSchedule(0).Text)
    If slLow = "" Then
        pbcSchedule.Cls
        Exit Sub
    End If
    If Val(slLow) <= 0 Then
        pbcSchedule.Cls
        Exit Sub
    End If
    slHigh = Trim$(edcSchedule(1).Text)
    If slHigh = "" Then
        pbcSchedule.Cls
        Exit Sub
    End If
    If Val(slHigh) - Val(slLow) < 12 Then
        pbcSchedule.Cls
        Exit Sub
    End If
    For ilRow = 1 To 2 Step 1
        'For ilBox = LBound(lmSSave) To UBound(lmSSave) Step 1
        For ilBox = LBONE To UBound(lmSSave) Step 1
            If ilRow = 1 Then
                'If ilBox = LBound(lmSSave) Then
                If ilBox = LBONE Then
                    slStr = ".01"
                Else
                    slStr = Trim$(Str$(lmSSave(ilBox - 1) + 1))
                End If
            Else
                'If ilBox = LBound(lmSSave) Then
                If ilBox = LBONE Then
                    slStr = edcSchedule(0).Text
                ElseIf ilBox = UBound(lmSSave) Then
                    slStr = "Above"
                Else
                    slStr = Trim$(Str$(lmSSave(ilBox)))
                End If
            End If
            gSetShow pbcSchedule, slStr, tmSCtrls(ilBox)
            pbcSchedule.CurrentX = tmSCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcSchedule.CurrentY = tmSCtrls(ilBox).fBoxY + (ilRow - 1) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            pbcSchedule.Print tmSCtrls(ilBox).sShow
        Next ilBox
    Next ilRow
End Sub


Private Sub plcAccount_Click()
    pbcClickFocus.SetFocus
End Sub



Private Sub plcBackup_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcBackup_Paint()
    plcBackup.CurrentX = 0
    plcBackup.CurrentY = 0
    'plcBackup.Print "Backup Settings"

End Sub



Private Sub plcBB_Paint(Index As Integer)
    plcBB(Index).CurrentX = 0
    plcBB(Index).CurrentY = -30
    If Index = 0 Then
        plcBB(0).Print "Show BB's on Separate Lines"
    Else
        plcBB(1).Print "BB Buy Type"
    End If
End Sub

Private Sub plcCntr_Click()
    pbcClickFocus.SetFocus
End Sub



Private Sub plcDiscDate_Paint()
    plcDiscDate.CurrentX = 0
    plcDiscDate.CurrentY = 0
    plcDiscDate.Print smDiscDateCaption
End Sub



Private Sub plcGeneral_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcInstRev_Paint()
    plcInstRev.CurrentX = 0
    plcInstRev.CurrentY = -30
    plcInstRev.Print "Installment Method as"
End Sub

Private Sub plcInv_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcComments_Click()
    pbcClickFocus.SetFocus
End Sub



Private Sub plcInvISCIForm_Paint()
    plcInvISCIForm.CurrentX = 0
    plcInvISCIForm.CurrentY = -30
    plcInvISCIForm.Print "For ISCI"
End Sub

Private Sub plcInvSortBy_Paint()
    plcInvSortBy.CurrentX = 0
    plcInvSortBy.CurrentY = -30
    plcInvSortBy.Print "Sort by"
End Sub

Private Sub plcInvSpotTimeZone_Paint()
    plcInvSpotTimeZone.CurrentX = 0
    plcInvSpotTimeZone.CurrentY = -30
    plcInvSpotTimeZone.Print "Change Invoice Spot Times to"
End Sub

Private Sub plcMerchPromo_Paint()
    plcMerchPromo.CurrentX = 0
    plcMerchPromo.CurrentY = -30
    plcMerchPromo.Print "Merchandising and Promotional Defined by"
End Sub

Private Sub plcPostRepAffidavit_Paint()
    plcPostRepAffidavit.CurrentX = 0
    plcPostRepAffidavit.CurrentY = -30
    plcPostRepAffidavit.Print "Post Rep by                              and/or Counts by "
End Sub

Private Sub plcPrintRepInv_Paint()
    plcPrintRepInv.CurrentX = 0
    plcPrintRepInv.CurrentY = -30
    plcPrintRepInv.Print "Print Rep Invoices by"
End Sub

Private Sub plcReallDate_Paint()
    plcReallDate.Cls
    plcReallDate.CurrentX = 0
    plcReallDate.CurrentY = -30
    plcReallDate.Print smReallDateCaption
End Sub


Private Sub plcSales_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcSSale_Paint(Index As Integer)
    plcSSale(Index).CurrentX = 0
    plcSSale(Index).CurrentY = -30
    Select Case Index
        Case 0
            plcSSale(Index).Print "On Spot Screen Use "
        Case 1
            plcSSale(Index).Print "For Billed and Booked - Display Adjustments for"
        Case 2
            plcSSale(Index).Print "Packages Allowed"
        Case 3
            plcSSale(Index).Print "Set Post Log Moved Spots as"
        Case 4
            plcSSale(Index).CurrentY = 0
            plcSSale(Index).Print "Address from"
        Case 5
            plcSSale(Index).Print "Default Avails Report to"
        Case 6
            plcSSale(Index).Print "Combo Avail Report Equalize by"
    End Select
End Sub

Private Sub plcTaxOnAirTime_Paint()
    plcTaxOnAirTime.CurrentX = 0
    plcTaxOnAirTime.CurrentY = -30
    plcTaxOnAirTime.Print "Tax Region"
End Sub

Private Sub rbcAISCI_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub rbcAPrtStyle_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub



Private Sub rbcBCombine_Click(Index As Integer)
    If rbcBCombine(Index).Value Then
        If Index = 0 Then
            ckcInv(1).Enabled = True
        Else
            ckcInv(1).Enabled = False
            ckcInv(1).Value = vbUnchecked
            ckcInstallment.Value = vbUnchecked
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub rbcBCombine_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange rbcBCombine(Index)
End Sub
Private Sub rbcBLCycle_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub rbcBNCycle_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub rbcBOrderDPShow_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub rbcBPkageGenMeth_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub rbcBRCycle_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub



Private Sub rbcCSchdPromo_Click(Index As Integer)
    If rbcCSchdPromo(Index).Value Then
        If Index = 0 Then
            ckcBookInto(0).Value = vbUnchecked
            ckcBookInto(0).Enabled = False
        ElseIf Index = 1 Then
            ckcBookInto(0).Enabled = True
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub rbcCSchdPromo_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcCSchdPSA_Click(Index As Integer)
    If rbcCSchdPSA(Index).Value Then
        If Index = 0 Then
            ckcBookInto(1).Value = vbUnchecked
            ckcBookInto(1).Enabled = False
        ElseIf Index = 1 Then
            ckcBookInto(1).Enabled = True
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub rbcCSchdPSA_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcCSchdRemnant_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcCSortBy_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcCUseCartNo_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcDefFillInv_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcEqualize_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcEqualize_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcFlatRateAverageFormula_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcGTBar_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub


Private Sub rbcImptCntr_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub



Private Sub rbcInvEmail_Click(Index As Integer)
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub rbcInvISCIForm_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcInvSortBy_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange rbcInvSortBy(Index)
End Sub


Private Sub rbcInvSpotTimeZone_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcInvSpotTimeZone_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange rbcInvSpotTimeZone(Index)
End Sub

Private Sub rbcNewBusYear_Click(Index As Integer)
    'dan M 10/24/11 added bmIgnoreChange
    If Index = 0 Then
        bmIgnoreChange = True
        edcSales(7).Enabled = False
        edcSales(8).Enabled = False
        edcSales(7).Text = "12"
        edcSales(8).Text = "12"
        bmIgnoreChange = False
    Else
        edcSales(7).Enabled = True
        edcSales(8).Enabled = True
        bmIgnoreChange = False
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub rbcPLMove_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcPostRepAffidavit_Click(Index As Integer)
    If Index = 3 Then
        If (rbcPostRepAffidavit(Index).Value) Then
            '1/8/10:  Allow Time posting if Using Rep or using RN_Rep
            'If (ckcUsingRep.Value = vbUnchecked) Or (ckcRN_Rep.Value = vbUnchecked) Then
            If (ckcUsingRep.Value = vbUnchecked) And (ckcRN_Rep.Value = vbUnchecked) Then
                rbcPostRepAffidavit(0).Value = True
            End If
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub rbcPrintRepInv_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange rbcPrintRepInv(Index)
End Sub

Private Sub rbcRCorpCal_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcRCorpCal(Index).Value
    'End of coded added
    If Value Then
        If Index = 0 Then
            cmcRCorpCal.Enabled = False
        Else
            cmcRCorpCal.Enabled = True
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub
Private Sub rbcRCorpCal_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
Private Sub rbcRCurrAmt_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange rbcRCurrAmt(Index)
End Sub
Private Sub rbcRRP_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange rbcRRP(Index)
End Sub
Private Sub rbcRUnbilled_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange rbcRUnbilled(Index)
End Sub
Private Sub rbcSInvCntr_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSInvCntr(Index).Value
    'End of coded added
    If Value Then
        'Mary request 12/13/00
        'If Index = 0 Then
        '    ckcSales(6).Value = False
        '    ckcSales(6).Enabled = False
        'Else
        '    ckcSales(6).Enabled = True
        'End If
        If Index <= 1 Then
            rbcBMissedDT(0).Enabled = True
            rbcBMissedDT(1).Enabled = True
        Else
            rbcBMissedDT(1).Value = False
            rbcBMissedDT(0).Enabled = False
            rbcBMissedDT(1).Enabled = False
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub
Private Sub rbcSInvCntr_GotFocus(Index As Integer)
    Dim ilPasswordOk As Integer
    mCtrlGotFocusAndIgnoreChange ActiveControl
    If imPWStatus = 0 Then
        imPWStatus = -2
        If (Trim$(tgUrf(0).sName) <> sgCPName) Then
            ilPasswordOk = igPasswordOk
            mProtectChangesAllowed Start
            CSPWord.Show vbModal
            mProtectChangesAllowed Done
            DoEvents
            If Not igPasswordOk Then
                imPWStatus = -1
                igPasswordOk = ilPasswordOk
                If tgSpf.sInvAirOrder = "S" Then 'Show Ordered, Update Ordered
                    rbcSInvCntr(0).Value = True
                ElseIf tgSpf.sInvAirOrder = "O" Then 'Show Ordered, Update Aired
                    rbcSInvCntr(1).Value = True
                ElseIf tgSpf.sInvAirOrder = "2" Then 'Show Aired minus Missed, Update Order
                    rbcSInvCntr(3).Value = True
                Else
                    rbcSInvCntr(2).Value = True     'As Aired
                End If
                cmcCommand(COMMAND_CANCEL).SetFocus
                Exit Sub
            Else
                imPWStatus = 1
                igPasswordOk = ilPasswordOk
            End If
        Else
            imPWStatus = 1
        End If
    Else
        If imPWStatus = -1 Then
            If tgSpf.sInvAirOrder = "S" Then 'Show Ordered, Update Ordered
                rbcSInvCntr(0).Value = True
            ElseIf tgSpf.sInvAirOrder = "O" Then 'Show Ordered, Update Aired
                rbcSInvCntr(1).Value = True
            ElseIf tgSpf.sInvAirOrder = "2" Then 'Show Aired minus Missed, Update Order
                rbcSInvCntr(3).Value = True
            Else
                rbcSInvCntr(2).Value = True     'As Aired
            End If
            cmcCommand(COMMAND_CANCEL).SetFocus
        End If
    End If
End Sub

Private Sub rbcSplitCopyState_Click(Index As Integer)
    If Not bmIgnoreChange Then
        mChangeOccured
    End If
End Sub

Private Sub rbcSplitCopyState_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange rbcSplitCopyState(Index)
End Sub

Private Sub rbcSUseProd_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange rbcSUseProd(Index)
End Sub
Private Sub plcRUnbilled_Paint()
    plcRUnbilled.CurrentX = 0
    plcRUnbilled.CurrentY = -30
    plcRUnbilled.Print "Unbilled $'s"
End Sub
Private Sub plcRCurrAR_Paint()
    plcRCurrAR.CurrentX = 0
    plcRCurrAR.CurrentY = -30
    plcRCurrAR.Print "Current A/R"
End Sub
Private Sub plcRRP_Paint()
    plcRRP.CurrentX = 0
    plcRRP.CurrentY = -30
    plcRRP.Print "Reconciling Period"
End Sub
Private Sub plcRCorpCal_Paint()
    plcRCorpCal.CurrentX = 0
    plcRCorpCal.CurrentY = -30
    plcRCorpCal.Print "Corporate Calendar"
End Sub
Private Sub plcAPrtStyle_Paint()
    plcAPrtStyle.CurrentX = 0
    plcAPrtStyle.CurrentY = -30
    plcAPrtStyle.Print "Contract Form"
End Sub
Private Sub plcAISCI_Paint()
    plcAISCI.CurrentX = 0
    plcAISCI.CurrentY = -30
    plcAISCI.Print "Using ISCI Codes on Invoices"
End Sub
Private Sub plcVirtPkg_Paint()
    plcVirtPkg.CurrentX = 0
    plcVirtPkg.CurrentY = -30
    plcVirtPkg.Print "For Package Line"
End Sub
Private Sub plcCSchd_Paint(Index As Integer)
    plcCSchd(Index).CurrentX = 0
    plcCSchd(Index).CurrentY = -30
    If Index = 0 Then
        plcCSchd(Index).Print "Remnant Contracts Scheduled"
    ElseIf Index = 1 Then
        plcCSchd(Index).Print "Promo Contracts Scheduled"
    ElseIf Index = 2 Then
        plcCSchd(Index).Print "PSA Contracts Scheduled"
    End If
End Sub
Private Sub plcDefLogCopy_Paint()
    plcDefLogCopy.CurrentX = 0
    plcDefLogCopy.CurrentY = -30
    plcDefLogCopy.Print "default Log 'Assign Copy'"
End Sub
Private Sub plcGTBar_Paint()
    plcGTBar.CurrentX = 0
    plcGTBar.CurrentY = -30
    plcGTBar.Print "In Screen Title Bar Show"
End Sub
Private Sub plcBMissedDT_Paint()
    plcBMissedDT.CurrentX = 0
    plcBMissedDT.CurrentY = -30
    plcBMissedDT.Print "'as Ordered' Missed, Show"
End Sub
Private Sub plcBLaserForm_Paint()
    plcBLaserForm.CurrentX = 0
    plcBLaserForm.CurrentY = -30
    plcBLaserForm.Print "Laser Print Form"
End Sub
Private Sub plcBOrderDPShow_Paint()
    plcBOrderDPShow.CurrentX = 0
    plcBOrderDPShow.CurrentY = -30
    plcBOrderDPShow.Print "Ordered Daypart Show"
End Sub
Private Sub plcBPkageGenMeth_Paint()
    plcBPkageGenMeth.CurrentX = 0
    plcBPkageGenMeth.CurrentY = -30
    plcBPkageGenMeth.Print "Calculate Virtual Package $ by"
End Sub
Private Sub plcSInvCntr_Paint()
    '10048
    plcSInvCntr.CurrentX = 120
    plcSInvCntr.CurrentY = -30
    plcSInvCntr.Print "Invoice All Contracts"
End Sub
Private Sub plcDefFillInv_Paint()
    plcDefFillInv.CurrentX = 0
    plcDefFillInv.CurrentY = -30
    plcDefFillInv.Print "default 'Bonus on Invoices' as"
End Sub
Private Sub plcBCombine_Paint()
    plcBCombine.CurrentX = 0
    plcBCombine.CurrentY = -30
    plcBCombine.Print "All Vehicles on Same Invoice"
End Sub
Private Sub plcBNCycle_Paint()
    plcBNCycle.CurrentX = 0
    plcBNCycle.CurrentY = -30
    plcBNCycle.Print "National"
End Sub
Private Sub plcBRCycle_Paint()
    plcBRCycle.CurrentX = 0
    plcBRCycle.CurrentY = -30
    plcBRCycle.Print "Regional"
End Sub
Private Sub plcBLCycle_Paint()
    '10048
    plcBLCycle.CurrentX = -15
    plcBLCycle.CurrentY = -30
    plcBLCycle.Print "default Billing Cycle:  Local"
End Sub

Private Sub rbcSystemType_Click(Index As Integer)
    If rbcSystemType(Index).Value Then
        If Index = 0 Then
            ckcGUseAffFeed.Visible = True
            frcPollUsers.Visible = False
        Else
            ckcGUseAffFeed.Visible = False
            frcPollUsers.Visible = True
        End If
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub rbcSystemType_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcTapeShowForm_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcTaxOnAirTime_Click(Index As Integer)
    If (rbcTaxOnAirTime(0).Value = True) Then
        edcBTTax(0).Text = ""
        edcBTTax(1).Text = ""
        edcBTTax(0).Enabled = False
        edcBTTax(1).Enabled = False
        ckcTaxOn(0).Enabled = False
        ckcTaxOn(1).Enabled = False
        ckcTaxOn(0).Value = vbUnchecked
        ckcTaxOn(1).Value = vbUnchecked
    Else
        edcBTTax(0).Enabled = True
        edcBTTax(1).Enabled = True
        ckcTaxOn(0).Enabled = True
        ckcTaxOn(1).Enabled = True
    End If
    If Not bmIgnoreChange Then
      mChangeOccured
    End If
End Sub

Private Sub rbcTaxOnAirTime_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange rbcTaxOnAirTime(Index)
End Sub

Private Sub tbcSelection_Click()
    Dim ilIndex As Integer

    ilIndex = tbcSelection.SelectedItem.Index - 1
    plcGeneral.Visible = False
    plcSales.Visible = False
    plcComm.Visible = False
    plcSchedule.Visible = False
    plcAgyAdv.Visible = False
    plcCntr.Visible = False
    plcCopy.Visible = False
    plcLog.Visible = False
    plcInv.Visible = False
    plcAccount.Visible = False
    cmcRCorpCal.Visible = False
    plcBackup.Visible = False
    plcOptions.Visible = False
    If udcSiteTabs.Visible Then
        udcSiteTabs.Action 5, True
        udcSiteTabs.Visible = False
    End If
    plcComments.Visible = False     '2-11-03
    'plcResearch.Visible = False
    Select Case ilIndex
        Case 0
            plcGeneral.Visible = True
        Case 1
            plcSales.Visible = True
        Case 2
            plcComm.Visible = True
        Case 3
            'plcResearch.Visible = True
            igUpdateAllowed = imUpdateAllowed
            udcSiteTabs.Action 2, True
            udcSiteTabs.Visible = True
            Screen.MousePointer = vbDefault
        Case 4
            plcSchedule.Visible = True
        Case 5 'Export
            plcAgyAdv.Visible = True
        Case 6
            plcCntr.Visible = True
        Case 7
            plcCopy.Visible = True
        Case 8
            plcLog.Visible = True
        Case 9
            plcInv.Visible = True
        Case 10
            plcAccount.Visible = True
            cmcRCorpCal.Visible = True
        Case 11
            plcBackup.Visible = True
        Case 12
            plcOptions.Visible = True
        Case 13
            If ckcUsingSports.Value = vbChecked Then
                igUpdateAllowed = imUpdateAllowed
                udcSiteTabs.Action 2, True
                udcSiteTabs.Visible = True
                Screen.MousePointer = vbDefault
            Else
                igUpdateAllowed = imUpdateAllowed
                udcSiteTabs.Action 2, False
                udcSiteTabs.Visible = True
                Screen.MousePointer = vbDefault
            End If
        Case 14
            'plcAutomation.Visible = True
            igUpdateAllowed = imUpdateAllowed
            udcSiteTabs.Action 2, True
            udcSiteTabs.Visible = True
            Screen.MousePointer = vbDefault
        Case 15
            plcComments.Visible = True '2-11-03
        ' Dan M 11/04/09 email site tab eliminated..6/23/11 reinstalled
        Case 16
            igUpdateAllowed = imUpdateAllowed
            udcSiteTabs.Action 2, True
            udcSiteTabs.Visible = True
            Screen.MousePointer = vbDefault
        Case 17
            igUpdateAllowed = imUpdateAllowed
            udcSiteTabs.Action 2, True
            udcSiteTabs.Visible = True
            Screen.MousePointer = vbDefault
    End Select

End Sub

Private Sub mGetLastBkup()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slLastDate                                                                            *
'******************************************************************************************

    Dim slBasePath As String
    Dim slFileName As String
    Dim llBkupSize As Long
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slBackupHour As String
    On Error GoTo mGenErr:

    'Build the base path to the Save directory
    ilPos = InStrRev(sgExePath, "\", Len(sgExePath) - 1)
    slBasePath = Left$(sgExePath, ilPos) & "SaveData\"

    smCSIServerINIFile = sgExePath & "\CSI_Server.ini"
    If Not gLoadINIValue(smCSIServerINIFile, "MainSettings", "LastBackupFileName", slFileName) Then
        edcLastBkupName.Text = "     *** No Backup Exists ***"
        edcLastBkupSize.Text = "0KB"
        Exit Sub
    End If

    If gLoadINIValue(smCSIServerINIFile, "Backup", "StartTime", slBackupHour) Then
        cbcBkupTime.Text = slBackupHour
    End If
    If gLoadINIValue(smCSIServerINIFile, "Backup", "WeekDays", smBUWeekDays) Then
        If Len(smBUWeekDays) = 7 Then
            For ilLoop = 1 To 7
                If Mid(smBUWeekDays, ilLoop, 1) = 1 Then
                    chkDOW(ilLoop - 1).Value = 1
                End If
            Next
        End If
    End If

    'On Error GoTo mIgnoreErr
    ilRet = 0
    'edcLastBkupDtTime.Text = FileDateTime(slBasePath & "\" & slFileName)
    ilRet = gFileExist(slBasePath & "\" & slFileName)
    'If ilRet = -1 Then
    If ilRet <> 0 Then
        edcLastBkupName.Text = "     *** No Backup Exists ***"
        edcLastBkupSize.Text = "0KB"
        Exit Sub
    End If
    edcLastBkupDtTime.Text = gFileDateTime(slBasePath & "\" & slFileName)
    llBkupSize = Round(FileLen(slBasePath & "\" & slFileName) / 1024)
    edcLastBkupSize.Text = CStr(llBkupSize) & " KB"
    edcLastBkupName.Text = slFileName
    edcLastBkupLoc.Text = Left$(sgExePath, ilPos) & "SaveData\"
    Exit Sub

mIgnoreErr:
    ilRet = -1
    Resume Next

mGenErr:
    MsgBox "An error has occured in mGetLastBkup."
End Sub


Public Function mUpdateComments(llSpfComment As Long, ilCommentLen As Integer, slComment As String) As String
Dim ilRet As Integer
Dim slMsg As String
Dim slSyncDate As String
Dim slSyncTime As String

    slMsg = ""
    tmCxf.lCode = llSpfComment
    'tmCxf.iStrLen = ilCommentLen
    tmCxf.sComment = Trim$(slComment) & Chr(0)
    imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment)) ' + 2    '25 = fixed record length; 2=Length value which is part of the variable record
    If tmCxf.lCode = 0 Then 'New
        'If Len(Trim$(tmCxf.sComment)) > 2 Then  '-2 so control character at end not counted
        If Trim$(slComment) <> "" Then
            tmCxf.sComType = "D"
            tmCxf.sShProp = "N"
            tmCxf.sShSpot = "N"
            tmCxf.sShOrder = "N"
            tmCxf.sShInv = "Y"
            tmCxf.sShInsertion = "N"
            tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
            tmCxf.lAutoCode = tmCxf.lCode
            ilRet = btrInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)

            tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
            tmCxf.lAutoCode = tmCxf.lCode
            tmCxf.iSourceID = tgUrf(0).iRemoteUserID
            gPackDate slSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
            gPackTime slSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
            imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment))
            ilRet = btrUpdate(hmCxf, tmCxf, imCxfRecLen)
        Else
            tmCxf.lCode = 0
            ilRet = BTRV_ERR_NONE
        End If
        slMsg = "mSaveRec (btrInsert: Comment)"
    Else 'Old record-Update
        '2-12-03 reread orig record
        tmCxfSrchKey.lCode = llSpfComment
        tmCxf.sComment = ""
        imCxfRecLen = Len(tmCxf)
        ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            tmCxf.lCode = 0
            'tmCxf.iStrLen = 0
            tmCxf.sComment = ""
            slComment = ""
        Else
            'tmCxf.iStrLen = ilCommentLen
            tmCxf.sComment = Trim$(slComment) & Chr$(0)
        End If

        'If Len(Trim$(tmCxf.sComment)) > 2 Then  '-2 so the control character at end is not counted
        If Trim$(slComment) <> "" Then  '-2 so the control character at end is not counted

            imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment)) ' + 2    '25 = fixed record length; 2=Length value which is part of the variable record

            tmCxf.iSourceID = tgUrf(0).iRemoteUserID
            gPackDate slSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
            gPackTime slSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
            ilRet = btrUpdate(hmCxf, tmCxf, imCxfRecLen)
        Else
            ilRet = btrDelete(hmCxf)
            tmCxf.lCode = 0
'            If tgSpf.sRemoteUsers = "Y" Then
'                tmDsf.lCode = 0
'                tmDsf.sFileName = "CXF"
'                gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'                gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'                tmDsf.iRemoteID = tmCxf.iRemoteID
'                tmDsf.lAutoCode = tmCxf.lAutoCode
'                tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'                tmDsf.lCntrNo = 0
'                ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'            End If
        End If
        slMsg = "mSaveRec (btrUpdate: Comment)"
    End If
    llSpfComment = tmCxf.lCode
    mUpdateComments = slMsg
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSEnableBox                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSEnableBox(ilBoxNo As Integer)
    If (ilBoxNo < imLBSCtrls) Or (ilBoxNo > UBound(tmSCtrls)) Then
        Exit Sub
    End If
    If imSRowNo <> 2 Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case LEVEL2INDEX To LEVEL14INDEX
            edcLevelPrice.Width = tmSCtrls(ilBoxNo).fBoxW
            edcLevelPrice.MaxLength = 0
            gMoveTableCtrl pbcSchedule, edcLevelPrice, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - 1) * (fgBoxGridH + 15)
            edcLevelPrice.Text = Trim$(Str$(lmSSave(ilBoxNo - LEVEL2INDEX + 1)))
            edcLevelPrice.Enabled = True
            edcLevelPrice.Visible = True  'Set visibility
            edcLevelPrice.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSSetShow                       *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSSetShow(ilBoxNo As Integer)
    Dim slStr As String

    If (ilBoxNo < imLBSCtrls) Or (ilBoxNo > UBound(tmSCtrls)) Then
        Exit Sub
    End If
    If (imSRowNo <> 2) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case LEVEL2INDEX To LEVEL14INDEX
            edcLevelPrice.Visible = False
            slStr = edcLevelPrice.Text
            If lmSSave(ilBoxNo - LEVEL2INDEX + 1) <> Val(slStr) Then
                lmSSave(ilBoxNo - LEVEL2INDEX + 1) = Val(slStr)
                If ilBoxNo = LEVEL2INDEX Then
                    edcSchedule(0).Text = slStr
                End If
                If ilBoxNo = LEVEL14INDEX Then
                    edcSchedule(1).Text = slStr
                End If
                pbcSchedule.Cls
                pbcSchedule_Paint
            End If
    End Select

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSSetFocus                      *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSSetFocus(ilBoxNo As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************


    If (ilBoxNo < imLBSCtrls) Or (ilBoxNo > UBound(tmSCtrls)) Then
        Exit Sub
    End If
    If (imSRowNo <> 2) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case LEVEL2INDEX To LEVEL14INDEX
            edcLevelPrice.Visible = True
            edcLevelPrice.SetFocus
    End Select

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
'*  flTextHeight                                                                          *
'******************************************************************************************


    gSetCtrl tmSCtrls(LEVEL2INDEX), 525, 225, 555, fgBoxGridH
    tmSCtrls(LEVEL2INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL3INDEX), 1095, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL3INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL4INDEX), 1665, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL4INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL5INDEX), 2235, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL5INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL6INDEX), 2805, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL6INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL7INDEX), 3375, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL7INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL8INDEX), 3945, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL8INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL9INDEX), 4515, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL9INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL10INDEX), 5085, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL10INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL11INDEX), 5655, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL11INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL12INDEX), 6225, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL12INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL13INDEX), 6795, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL13INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL14INDEX), 7365, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL14INDEX).iReq = False
    gSetCtrl tmSCtrls(LEVEL15INDEX), 7935, tmSCtrls(LEVEL2INDEX).fBoxY, 555, fgBoxGridH
    tmSCtrls(LEVEL15INDEX).iReq = False

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSafReadRec                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mSafReadRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mSafReadRecErr                                                                        *
'******************************************************************************************

'
'   iRet = mSafReadRec()
'   Where:
'       ilVefCode(I) - Vehicle Code
'       slType (I) - L=Log; C=CP and O= Other
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    tmSafSrchKey1.iVefCode = 0
    ilRet = btrGetEqual(hmSaf, tmSaf, imSafRecLen, tmSafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        'Add Record
        tmSaf.sSchRNG = "T"
        tmSaf.sSchMaL = "H"
        tmSaf.sSchMdL = "D"
        tmSaf.sSchMiL = "Q"
        tmSaf.sSchCycle = "1"
        tmSaf.sSchMove = "N"
        tmSaf.sSchCompact = "N"
        tmSaf.sSchPreempt = "E"
        tmSaf.sSchHour = "B"
        tmSaf.sSchDay = "B"
        tmSaf.sSchQH = "B"
        tmSaf.sName = ""
        tmSaf.iVefCode = 0
        tmSaf.lLowPrice = 0
        tmSaf.lLevelToPrice(0) = 0
        tmSaf.lLevelToPrice(1) = 0
        tmSaf.lLevelToPrice(2) = 0
        tmSaf.lLevelToPrice(3) = 0
        tmSaf.lLevelToPrice(4) = 0
        tmSaf.lLevelToPrice(5) = 0
        tmSaf.lLevelToPrice(6) = 0
        tmSaf.lLevelToPrice(7) = 0
        tmSaf.lLevelToPrice(8) = 0
        tmSaf.lLevelToPrice(9) = 0
        tmSaf.lLevelToPrice(10) = 0
        'tmSaf.lLevelToPrice(11) = 0
        tmSaf.lHighPrice = 0
        tmSaf.iPreferredPct = 0
        tmSaf.iWk1stSoloIndex = 0
        tmSaf.sInvISCIForm = "R"
        tmSaf.lGPBatchNo = 0
        tmSaf.sGPPrefixChar = ""
        tmSaf.sGPCustomerNo = ""
        tmSaf.iRptLenDefault(0) = 0
        tmSaf.iRptLenDefault(1) = 0
        tmSaf.iRptLenDefault(2) = 0
        tmSaf.iRptLenDefault(3) = 0
        tmSaf.iRptLenDefault(4) = 0
        tmSaf.iNoDaysRetainUAF = 5
        tmSaf.sFinalLogDisplay = "N"
        gPackDate "1/1/1990", tmSaf.iLastArchRunDate(0), tmSaf.iLastArchRunDate(1)
        gPackDate "1/1/1990", tmSaf.iEarliestAffSpot(0), tmSaf.iEarliestAffSpot(1)
        gPackDate "1/1/1990", tmSaf.iBLastWeeklyDate(0), tmSaf.iBLastWeeklyDate(1)
        gPackDate "1/1/1970", tmSaf.iXDSLastImptDate(0), tmSaf.iXDSLastImptDate(1)
        tmSaf.sFinalLogDisplay = "N"
        tmSaf.sProdProtMan = "N"
        tmSaf.sAvailGreenBar = "N"
        tmSaf.sIPumpZone = ""
        gPackDate "1/1/1990", tmSaf.iVCreativeDate(0), tmSaf.iVCreativeDate(1)
        tmSaf.sInvoiceSort = "P"
        tmSaf.sGenAutoFileWOSpt = "N"
        tmSaf.sFeatures1 = Chr(0)
        tmSaf.sFeatures2 = Chr(0)
        tmSaf.sEventTitle1 = ""
        tmSaf.sEventTitle2 = ""
        tmSaf.sEventSubtotal1 = ""
        tmSaf.sEventSubtotal2 = ""
        tmSaf.sCreditLimitMsg = "W"
        tmSaf.sXMidSpotsBill = "O"
        tmSaf.sHideDemoOnBR = "N"
        tmSaf.sClientSentToEDS = "N"
        tmSaf.sXDSHeadEndZone = "E"
        tmSaf.sSyncCopyInRot = "N"
        tmSaf.sFeatures3 = Chr(0)
        tmSaf.lStatementComment = 0
        tmSaf.sAudByPackage = "N"
        tmSaf.sFeatures4 = Chr(0)
        tmSaf.sFeatures5 = Chr(0)
        tmSaf.sFeatures6 = Chr(0)
        tmSaf.lCitationComment = 0
        tmSaf.sExcludeAudioTypeR = "N"       ' Exclude Autio Type Recorded
                                                 ' Commercial (RC: R). (Y/N, Test
                                                 ' for Y. Treat Blank same as N)
        tmSaf.sExcludeAudioTypeL = "N"         ' Exclude Autio Type Live
                                                 ' Commercial (LC: L). (Y/N, Test
                                                 ' for Y. Treat Blank same as N)
        tmSaf.sExcludeAudioTypeM = "N"         ' Exclude Autio Type Live Promo
                                                 ' (LP: M). (Y/N, Test for Y. Treat
                                                 ' Blank same as N)
        tmSaf.sExcludeAudioTypeS = "N"         ' Exclude Autio Type Recorded Promo
                                                 ' (RP: S). (Y/N, Test for Y. Treat
                                                 ' Blank same as N)
        tmSaf.sExcludeAudioTypeP = "N"         ' Exclude Autio Type Pre-Recorded
                                                 ' Live Commercial (PC: P). (Y/N,
                                                 ' Test for Y. Treat Blank same as
                                                 ' N)
        tmSaf.sExcludeAudioTypeQ = "N"        ' Exclude Autio Type Pre-Recorded
        tmSaf.sInvExpDelimiter = ","            'Invoice export delimiter
        tmSaf.sFeatures7 = Chr(0)
        tmSaf.sUnused = ""
        tmSaf.iCode = 0  'Autoincrement
    End If
    mSafReadRec = True
    Exit Function
mSafReadRecErr: 'VBC NR
    On Error GoTo 0
    mSafReadRec = False
    Exit Function
End Function


Private Function mSiteReadRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mSiteReadRecErr                                                                       *
'******************************************************************************************

    Dim ilRet As Integer    'Return status

    tmSiteSrchKey.lCode = 1
    ilRet = btrGetEqual(hmSite, tgSite, imSiteRecLen, tmSiteSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        'Add Record
        tgSite.lCode = 0
        tgSite.iMarketron = 0
        tgSite.iWeb = 0
        tgSite.iNoDaysDelq = 0
        'tgSite.iDaysResendISCI = 0
        tgSite.iUnused3 = 0
        tgSite.iCMCmtCode = 0
        tgSite.iWMCmtCode = 0
        tgSite.iPMCmtCode = 0
        tgSite.iOMCmtCode = 0
        tgSite.iOMNoWeeks = 0
        tgSite.iDaysRetainSpots = 0
        tgSite.lAdminArttCode = 0
        tgSite.sVehicleStn = "N"
        'tgSite.sISCIExport = "B"
        ''tgSite.iNoMnthRetain = 999
        ''gPackDate "1/1/1970", tgSite.iLastDateArch(0), tgSite.iLastDateArch(1)
        tgSite.sUsingServAgree = "N"
        tgSite.sSyncMulticast = "N"
        tgSite.iBandCmtCode = 0
        tgSite.sUnused2 = ""
        tgSite.sShowVehType = "N"
        tgSite.sExportCart4 = ""    ' Use 4 character Cart Numbers
        tgSite.sExportCart5 = ""    ' Use 4 character Cart Numbers
        tgSite.sCCEMail = ""
        tgSite.iDayRetainPost = 0       ' Number of days to retain web
                                                 ' posted spots
        tgSite.sChngPswd = ""    ' Flag to allow web users to change
                                                 ' passwords or not - Y/N
        gPackDate "1/1/1970", tgSite.iOMMinDate(0), tgSite.iOMMinDate(1)
        tgSite.sEmailHost = ""   ' Host name/SMTP server address or
                                                 ' URL
        tgSite.iEmailPort = 0       ' Port Number for SMTP server
        tgSite.sEmailAcctName = ""  ' Account name for SMTP
                                                 ' credentials.
        tgSite.sEmailPassword = ""   ' Password for the SMTP credentials
        tgSite.sEmailFromAddress = ""    ' The email address of the sender.
        tgSite.sEmailFromName = ""  ' Used as from email name.  Like
                                                 ' "Dial Global Netowrk" or "ABC
                                                 ' Clearance Dept."
        tgSite.iNCRWks = 60
        tgSite.sWebSuppressLog = "N"
        tgSite.sMultiVehWebPost = "N"
        tgSite.iWebNoDyKeepMiss = 60
        tgSite.sAllowBonusSpots = "N"
        tgSite.sShowContrDate = "N"
        tgSite.sDDF092710 = "Y"
        tgSite.sUsingViero = "N"
        tgSite.sStationToXDS = "N"
        tgSite.sAgreementToXDS = "N"
        tgSite.sMissedDateTime = "N"
        tgSite.sCompliantBy = "P"
        tgSite.iWebNoDyViewPost = 7
        gPackDate "1/1/1970", tgSite.iRqtDate(0), tgSite.iRqtDate(1)
        gPackTime "12AM", tgSite.iRqtTime(0), tgSite.iRqtTime(1)
        tgSite.sGenTransparent = "N"
        tgSite.sProgToXDS = "N"
        gPackDate "1/1/1970", tgSite.iSSBDate(0), tgSite.iSSBDate(1)
        gPackTime "12AM", tgSite.iSSBTime(0), tgSite.iSSBTime(1)
        tgSite.sSupportXDSDelay = "N"
        tgSite.sAllowAutoPost = "N"
        tgSite.sWithinMissMonth = "N"       ' Within missed standard month (Y=Yes; N or Blank=No). Test for Y
        tgSite.sLastWk1stWk = "N"           ' Miised in last week of the standard month, MG allowed in 1st
                                            ' week of next month (Y=Yes; N or Blank = No). Test for Y
        tgSite.sSkipHiatusWk = "N"          ' Disallow MG in skipped Weeks (Hiatus). (Y=Yes; N or Blank=No). Test for Y
        tgSite.sValidDaysOnly = "N"         ' Book only into valid flight days. (Y=Yes; N or Blank=No). Test for Y
        tgSite.sTimeRange = "N"             ' MG with time range (O=Flight Order Times; P=Pledge Times; N or
                                            ' Blank=Any Time).Test for O or P
        tgSite.sISCIPolicy = "O"            ' ISCI Policy: O or Blank=Order; A=Advertiser after Order failed. Test for A
        tgSite.iMGDays = -1
        tgSite.iCompetSepTime = 0
        tgSite.sAllowMGSpots = "Y"
        tgSite.sAllowReplSpots = "Y"
        tgSite.sNoMissedReason = "N"
        tgSite.sDefaultEstDay = "N"
        tgSite.sWebPostInFuture = "Y"
        tgSite.sMissedMGBypass = "N"
        'tgSite.sUnused = ""
        tgSite.sRADARMultiAir = "S"
        tgSite.lAstMaxLastValue = 0
        gPackDate "1/1/1970", tgSite.iAstMaxDate(0), tgSite.iAstMaxDate(1)
        '9926
        tgSite.iUMCmtCode = 0

    End If
    mSiteReadRec = True
    Exit Function
mSiteReadRecErr: 'VBC NR
    On Error GoTo 0
    mSiteReadRec = False
    Exit Function
End Function

Private Function mNrfReadRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mNtrReadRecErr                                                                        *
'******************************************************************************************

    Dim ilRet As Integer    'Return status

    If (((Asc(tgSpf.sAutoType2)) And RN_REP) <> RN_REP) And (((Asc(tgSpf.sAutoType2)) And RN_NET) <> RN_NET) Then
        mNrfReadRec = True
        Exit Function
    End If
    ilRet = btrGetFirst(hmNrf, tgNrf, imNrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If (((Asc(tgSpf.sAutoType2)) And RN_REP) = RN_REP) Then
            If tgNrf.sType = "R" Then
                mNrfReadRec = True
                Exit Function
            End If
        ElseIf (((Asc(tgSpf.sAutoType2)) And RN_NET) = RN_NET) Then
            If tgNrf.sType = "N" Then
                mNrfReadRec = True
                Exit Function
            End If
        End If
        ilRet = btrGetNext(hmNrf, tgNrf, imNrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    tgNrf.iCode = 0
    tgNrf.sType = ""
    tgNrf.sName = ""
    tgNrf.sDBID = ""
    tgNrf.sContactName = ""
    tgNrf.sPhoneNo = ""
    tgNrf.sEMail = ""
    tgNrf.sSlspFirstName = ""
    tgNrf.sSlspLastName = ""
    tgNrf.sFTPIsOn = "Y"
    tgNrf.sFTPUserID = ""
    tgNrf.sFTPUserPW = ""
    tgNrf.sFTPAddress = ""
    tgNrf.sFTPImportDir = ""
    tgNrf.sFTPExportDir = ""
    tgNrf.sInsertionType = ""
    tgNrf.sIISRootURL = ""
    tgNrf.sIISRegSection = ""
    tgNrf.iFTPPort = 0
    tgNrf.sUnused = ""
    mNrfReadRec = True
    Exit Function
mNtrReadRecErr: 'VBC NR
    On Error GoTo 0
    mNrfReadRec = False
    Exit Function
End Function

Private Function mPopPollUserListBox(frm As Form, lbcPoll1 As Control, lbcPoll2 As Control) As Integer
'
'   ilRet = mPopBkupUserListBox (MainForm, cbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       cbcCtrl (I)- List box control that will be populated with names
'       ilRet (O)- True=list was either populated or repopulated
'                  False=List was OK- it didn't require populating
'

    Dim ilRecLen As Integer     'URF record length
    Dim hlUrf As Integer        'User Option file handle
    Dim tlUrf As URF
    Dim ilRet As Integer

    mPopPollUserListBox = True
    hlUrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mPopPollUserListBoxErr
    gBtrvErrorMsg ilRet, "mPopPollUserListBox (btrOpen):" & "Urf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlUrf)  'Get and save record length
    lbcPoll1.Clear
    lbcPoll2.Clear

    lbcPoll1.AddItem " [None]"
    lbcPoll1.ItemData(lbcPoll1.NewIndex) = 0
    lbcPoll2.AddItem " [None]"
    lbcPoll2.ItemData(lbcPoll2.NewIndex) = 0

    ilRet = btrGetFirst(hlUrf, tlUrf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        gUrfDecrypt tlUrf
        If (tlUrf.sDelete <> "Y") Then
            gFindMatch tlUrf.sRept, 0, lbcPoll1
            If gLastFound(lbcPoll1) < 0 Then
                If tlUrf.iCode > 2 Then
                    If Trim$(tlUrf.sRept) <> "" Then
                        lbcPoll1.AddItem " " & Trim$(tlUrf.sRept)
                        lbcPoll2.AddItem " " & Trim$(tlUrf.sRept)
                    Else
                        lbcPoll1.AddItem " " & Trim$(tlUrf.sRept)
                        lbcPoll2.AddItem " " & Trim$(tlUrf.sRept)
                    End If
                    lbcPoll1.ItemData(lbcPoll1.NewIndex) = tlUrf.iCode
                    lbcPoll2.ItemData(lbcPoll2.NewIndex) = tlUrf.iCode
                End If
            End If
        End If
        ilRet = btrGetNext(hlUrf, tlUrf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        On Error GoTo mPopPollUserListBoxErr
        gBtrvErrorMsg ilRet, "mPopPollUserListBox (btrGetFirst):" & "Urf.Btr", frm
        On Error GoTo 0
    End If
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    Exit Function
mPopPollUserListBoxErr:
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    gDbg_HandleError "SiteOpt: mPopPollUserListBox"
    mPopPollUserListBox = False
    Exit Function
End Function
Public Sub mChangeOccured()
    imChangesOccured = imChangesOccured + 1
    mChangeLabel
    bmIgnoreChange = True
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
Private Sub mCtrlGotFocusAndIgnoreChange(Ctrl As Control)
'   Dan M 4/22/09 hijacked gCtrlGotFocus and added bmIgnoreChange to help control textboxes and counting changes
    bmIgnoreChange = False
    gCtrlGotFocus Ctrl
End Sub

Public Sub mProtectChangesAllowed(blStart As ProtectChanges)
Static isSaveChangesAllowed As Integer
If blStart Then
    isSaveChangesAllowed = igChangesAllowed
Else
    igChangesAllowed = isSaveChangesAllowed
    If Not igPasswordOk Then   'And imChangesOccured > 1
        imChangesOccured = imChangesOccured - 1
    End If
End If
End Sub
Private Sub mFixDisplay()
    Dim c As Integer
    For c = 0 To 2
        frcOption(c).Left = 100
    Next c
    frcOption(4).Left = 100
    ckcOptionFields(17).Left = ckcOptionFields(17).Left + 60
    frcOption(6).Left = 100 'Export/Import Options
End Sub
'automated events Dan M 4/22/09

Private Sub cbcUser_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcAPenny_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcExpt_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcBarterLPD_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcBBillDate_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcBLogoSpaces_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcBNo_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcBPayAddr_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcBPayName_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcCNo_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcComment_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcDiscNo_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcEDI_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcGAddr_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcGAlertInterval_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcGClient_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcGClientAbbr_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcGRetain_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcGRetainPassword_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Sub edcInvExportId_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcLevelPrice_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcLnOverride_Change(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcRB_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcRCollectContact_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcRCreditDate_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcRCRP_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcRNewCntr_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcRNRP_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcRPctCredit_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcRPRP_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcSales_Change(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub edcSchedule_Change(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub edcTerms_Change()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcBCombine_Change(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
'begin rbc and ckc


Private Sub ckcACodes_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcDaylightSavings_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcAEDI_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
    If Index = 1 Then
        If ckcAEDI(Index).Value = vbChecked Then
            ckcAEDI(2).Enabled = True
        Else
            ckcAEDI(2).Enabled = False
            ckcAEDI(2).Value = vbUnchecked
        End If
    End If
End Sub

Private Sub ckcAllowPrelLog_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
Private Sub ckcBookInto_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcCAudPkg_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub


Private Sub ckcCBump_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
'10843 removed ckc
'Private Sub ckcCLnStdQt_Click()
'    If Not bmIgnoreChange Then
'          mChangeOccured
'    End If
'End Sub

Private Sub ckcCAvails_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcCmmlSchStatus_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
Private Sub ckcCUseSegments_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcCWarnMsg_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcDisallowAuthorScheduling_Click()

    Dim ilPasswordOk As Integer
    '10048 podcast
    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If ckcDisallowAuthorScheduling.Value = vbChecked Then
                ilPasswordOk = igPasswordOk
                mProtectChangesAllowed Start
                CSPWord.Show vbModal
                mProtectChangesAllowed Done
                If Not igPasswordOk Then
                    ckcDisallowAuthorScheduling.Value = vbUnchecked
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            ckcDisallowAuthorScheduling.Value = vbUnchecked
        End If
    End If
   
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcRUseTMP_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcSDelivery_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcSMktBase_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcSSelling_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
Private Sub ckcSuppressTimeForm1_Click()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub ckcTaxOn_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
Private Sub rbcAISCI_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcAPrtStyle_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub


Private Sub rbcBLCycle_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcBNCycle_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcBOrderDPShow_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcBPkageGenMeth_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcBRCycle_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub



Private Sub rbcCSchdRemnant_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcCSortBy_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcCUseCartNo_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcDefFillInv_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcGTBar_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcImptCntr_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcInvISCIForm_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcInvSortBy_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcPLMove_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcPrintRepInv_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub


Private Sub rbcRCurrAmt_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcRRP_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcRUnbilled_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub


Private Sub rbcSUseProd_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub


Private Sub rbcTapeShowForm_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

'controls that weren't using gotfocus:

Private Sub rbcCommBy_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcCSIBackup_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcLastBkupLoc_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcLastBkupName_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcLastBkupSize_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcLastBkupDtTime_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

'Private Sub cbcBkupTime_gotFocus()
'    mCtrlGotFocusAndIgnoreChange ActiveControl
'End Sub

Private Sub rbcInstRev_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcRMerchPromo_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

'Private Sub cbcReconGroup_gotFocus()
'    mCtrlGotFocusAndIgnoreChange ActiveControl
'End Sub

Private Sub edcGP_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcVirtPkg_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcUnitOr3060_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcInsertAddr_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcSEnterAge_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcPostRepAffidavit_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub edcBTTax_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcBMissedDT_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcBLaserForm_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub ckcRegionMixLen_gotFocus()
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcBBType_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcBBOnLine_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcFillCopyAssign_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcMGCopyAssign_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcMGRules_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub

Private Sub rbcDefLogCopy_gotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
End Sub
'above controls _click event
Private Sub rbcCommBy_Click(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub


Private Sub rbcInstRev_Click(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub rbcRMerchPromo_Click(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub rbcVirtPkg_Click(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub rbcUnitOr3060_Click(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub rbcInsertAddr_Click(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub rbcSEnterAge_Click(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub rbcBMissedDT_Click(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub rbcBLaserForm_Click(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub ckcRegionMixLen_Click()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub rbcBBType_Click(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub rbcBBOnLine_Click(Index As Integer)
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub rbcFillCopyAssign_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcMGCopyAssign_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcMGRules_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub rbcDefLogCopy_Click(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
Private Sub edcLastBkupLoc_Change()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub edcLastBkupName_Change()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub edcLastBkupSize_Change()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub edcLastBkupDtTime_Change()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub


Private Sub edcBTTax_Change(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
Private Sub mVatSetToGoToWeb(ilVendor As Integer)
    '7942
    Dim rst_pw As ADODB.Recordset

    On Error GoTo ErrHand
        
    SQLQuery = "UPDATE VAT_Vendor_Agreement SET vatSentToWeb = '' WHERE vatWvtVendorId = " & ilVendor
    'Set rst_pw = cnn.Execute(SQLQuery)
    Set rst_pw = gSQLSelectCall(SQLQuery)
    Exit Sub
    
ErrHand:
End Sub

Private Sub mSetAudioTypes()
    ckcCopy(6).Enabled = True
    ckcCopy(7).Enabled = True
    ckcCopy(8).Enabled = True
    ckcCopy(9).Enabled = True
    ckcCopy(10).Enabled = True
    If (ckcUsingLiveCopy.Value = vbUnchecked) And (ckcCopy(1).Value = vbUnchecked) Then
        ckcOverrideOptions(5).Value = vbUnchecked
        ckcOverrideOptions(5).Enabled = False
        'ckcCopy(5).Enabled = True
        ckcCopy(6).Enabled = False
        ckcCopy(6).Value = vbChecked
        ckcCopy(7).Enabled = False
        ckcCopy(7).Value = vbChecked
        ckcCopy(8).Enabled = False
        ckcCopy(8).Value = vbChecked
        ckcCopy(9).Enabled = False
        ckcCopy(9).Value = vbChecked
        ckcCopy(10).Enabled = False
        ckcCopy(10).Value = vbChecked
    ElseIf (ckcUsingLiveCopy.Value = vbUnchecked) Then
        ckcCopy(6).Enabled = False
        ckcCopy(6).Value = vbChecked
        ckcCopy(7).Enabled = False
        ckcCopy(7).Value = vbChecked
        ckcCopy(9).Enabled = False
        ckcCopy(9).Value = vbChecked
        ckcCopy(10).Enabled = False
        ckcCopy(10).Value = vbChecked
    ElseIf (ckcCopy(1).Value = vbUnchecked) Then
        ckcCopy(7).Enabled = False
        ckcCopy(7).Value = vbChecked
        ckcCopy(8).Enabled = False
        ckcCopy(8).Value = vbChecked
        ckcCopy(10).Enabled = False
        ckcCopy(10).Value = vbChecked
    Else
        ckcOverrideOptions(5).Enabled = True
        'ckcCopy(5).Enabled = True
        ckcCopy(6).Enabled = True
        ckcCopy(7).Enabled = True
        ckcCopy(8).Enabled = True
        ckcCopy(9).Enabled = True
        ckcCopy(10).Enabled = True
    End If

End Sub
Private Sub mInvoiceEmailOptions()
'10016
    Dim blCombine As Boolean
    Dim blUnCombined As Boolean
    
    blCombine = False
    blUnCombined = False
    If ckcInv(INVAUTO).Value = vbChecked Then
        'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
        ckcInv(INVSELECTIVE).Value = vbUnchecked
        ckcInv(INVSELECTIVE).Enabled = False
        blCombine = True
        If ckcInv(INVCOMBINE).Value <> vbChecked Then
            blUnCombined = True
        End If
    Else
        'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
        ckcInv(INVSELECTIVE).Enabled = True
    End If
    'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
    If ckcInv(INVSELECTIVE).Value = vbChecked Then
        ckcInv(INVAUTO).Value = vbUnchecked
        ckcInv(INVAUTO).Enabled = False
    Else
        ckcInv(INVAUTO).Enabled = True
    End If
    
    rbcInvEmail(AIRANDNTR).Enabled = blCombine
    rbcInvEmail(AIRONLY).Enabled = blUnCombined
    rbcInvEmail(NTRONLY).Enabled = blUnCombined
    
    If ckcInv(INVAUTO).Value <> vbChecked Then
        rbcInvEmail(AIRANDNTR).Value = True
    End If
    If ckcInv(INVCOMBINE).Value = vbChecked Then
        rbcInvEmail(AIRANDNTR).Value = True
    End If
    If rbcInvEmail(AIRONLY).Value = False And rbcInvEmail(NTRONLY).Value = False Then
        rbcInvEmail(AIRANDNTR).Value = True
    End If
End Sub
Private Sub mPodcastOptions(ilIndex As Integer)
    If ilIndex = PODAIRTIMECKC Or ilIndex = PODSPOTSCKC Or ilIndex = ADSERVERCKC Or ilIndex = PODMIXCKC Or ilIndex < 0 Then
        ckcOptionFields(PODMIXCKC).Enabled = False
        ckcPodShowWk.Enabled = False
        If ckcOptionFields(PODAIRTIMECKC).Value = vbChecked Then
            If ckcOptionFields(PODSPOTSCKC).Value = vbChecked Or ckcOptionFields(ADSERVERCKC).Value = vbChecked Then
                ckcOptionFields(PODMIXCKC).Enabled = True
                If ckcOptionFields(PODSPOTSCKC).Value = vbChecked And ckcOptionFields(PODMIXCKC).Value = vbChecked Then
                    'for mixed packages. only enabled if 3 boxes checked!
                    ckcPodShowWk.Enabled = True
                End If
            End If
        End If
        If ckcOptionFields(PODMIXCKC).Enabled = False Then
            ckcOptionFields(PODMIXCKC).Value = vbUnchecked
        End If
        If ckcPodShowWk.Enabled = False Then
            ckcPodShowWk.Value = vbUnchecked
        End If
    End If
End Sub
'L.Bianchi 06/10/2021 WO Inovice Export
Private Sub edcIE_Change(Index As Integer)
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub
'L.Bianchi 06/10/2021 WO Inovice Export
Private Sub edcIE_GotFocus(Index As Integer)
    mCtrlGotFocusAndIgnoreChange ActiveControl
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
    'If igWinStatus(INVOICESJOB) = 0 Then
    '    imTerminate = True
    'End If
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

