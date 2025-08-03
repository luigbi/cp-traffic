VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form VehOpt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7230
   ClientLeft      =   495
   ClientTop       =   1830
   ClientWidth     =   11910
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
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7230
   ScaleWidth      =   11910
   Begin VB.PictureBox plcAccounting 
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
      Left            =   120
      ScaleHeight     =   5835
      ScaleWidth      =   10875
      TabIndex        =   329
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   10935
      Begin VB.CheckBox ckcXDSave 
         Alignment       =   1  'Right Justify
         Caption         =   "Cue code by zone"
         Height          =   210
         Index           =   7
         Left            =   8280
         TabIndex        =   410
         Top             =   5550
         Width           =   1815
      End
      Begin VB.CheckBox ckcXDSave 
         Caption         =   "NAS"
         Height          =   225
         Index           =   6
         Left            =   9480
         TabIndex        =   482
         Top             =   5205
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox ckcXDSave 
         Caption         =   "HDD"
         Height          =   225
         Index           =   5
         Left            =   8820
         TabIndex        =   481
         Top             =   5205
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox ckcXDSave 
         Caption         =   "CF"
         Height          =   225
         Index           =   4
         Left            =   8280
         TabIndex        =   480
         Top             =   5205
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox edcXDISCIPrefix 
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
         Left            =   6105
         MaxLength       =   6
         TabIndex        =   479
         Top             =   5145
         Width           =   780
      End
      Begin VB.TextBox edcInterfaceID 
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
         Left            =   4185
         MaxLength       =   8
         TabIndex        =   406
         Top             =   5130
         Width           =   780
      End
      Begin VB.CheckBox ckcXDSave 
         Alignment       =   1  'Right Justify
         Caption         =   "Honor Program ID ""Merge"""
         Height          =   225
         Index           =   3
         Left            =   825
         TabIndex        =   404
         Top             =   4950
         Width           =   2565
      End
      Begin VB.TextBox edcExport 
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
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   408
         Top             =   5505
         Width           =   1245
      End
      Begin VB.CheckBox ckcAffExport 
         Alignment       =   1  'Right Justify
         Caption         =   "Export ISCI by Pledge (Unchecked means by Feed)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3405
         TabIndex        =   409
         Top             =   5550
         Width           =   4530
      End
      Begin VB.CheckBox ckcXDSave 
         Caption         =   "NAS"
         Height          =   225
         Index           =   2
         Left            =   9480
         TabIndex        =   403
         Top             =   4695
         Width           =   645
      End
      Begin VB.CheckBox ckcXDSave 
         Caption         =   "HDD"
         Height          =   225
         Index           =   1
         Left            =   8820
         TabIndex        =   402
         Top             =   4695
         Width           =   645
      End
      Begin VB.CheckBox ckcXDSave 
         Caption         =   "CF"
         Height          =   225
         Index           =   0
         Left            =   8280
         TabIndex        =   401
         Top             =   4695
         Width           =   510
      End
      Begin VB.PictureBox pbcXDXMLForm 
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
         Left            =   1845
         ScaleHeight     =   210
         ScaleWidth      =   1185
         TabIndex        =   398
         Top             =   4695
         Width           =   1185
      End
      Begin VB.TextBox edcXDISCIPrefix 
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
         Left            =   6105
         MaxLength       =   6
         TabIndex        =   400
         Top             =   4650
         Width           =   780
      End
      Begin VB.CheckBox ckcAffExport 
         Alignment       =   1  'Right Justify
         Caption         =   "OLA Export"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   9705
         TabIndex        =   411
         Top             =   5475
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.TextBox edcInterfaceID 
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
         Left            =   4185
         MaxLength       =   8
         TabIndex        =   399
         Top             =   4650
         Width           =   780
      End
      Begin VB.TextBox edcRadarCode 
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
         Left            =   7830
         MaxLength       =   5
         TabIndex        =   388
         Top             =   3345
         Width           =   780
      End
      Begin VB.TextBox edcARBCode 
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
         Left            =   1140
         MaxLength       =   6
         TabIndex        =   386
         Top             =   3345
         Width           =   780
      End
      Begin VB.CheckBox ckcExportSQL 
         Alignment       =   1  'Right Justify
         Caption         =   "Vehicle Exported to SQL"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5370
         TabIndex        =   396
         Top             =   4335
         Width           =   2520
      End
      Begin VB.CheckBox ckcKCGenRot 
         Alignment       =   1  'Right Justify
         Caption         =   "KenCast:  Include Rotation Reference in Envelope"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   395
         Top             =   4335
         Width           =   4680
      End
      Begin VB.TextBox edcEDASWindow 
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
         Left            =   4260
         MaxLength       =   6
         TabIndex        =   394
         Top             =   3990
         Width           =   1005
      End
      Begin VB.CheckBox ckcStnFdInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Include in Cross Reference"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   5070
         TabIndex        =   392
         Top             =   3735
         Width           =   2640
      End
      Begin VB.CheckBox ckcStnFdInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Cart #'s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3330
         TabIndex        =   391
         Top             =   3735
         Width           =   1515
      End
      Begin VB.TextBox edcStnFdCode 
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
         MaxLength       =   2
         TabIndex        =   390
         Top             =   3690
         Width           =   390
      End
      Begin VB.TextBox edcExpCntrVehNo 
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
         Left            =   5595
         MaxLength       =   2
         TabIndex        =   387
         Top             =   3345
         Width           =   390
      End
      Begin VB.PictureBox plcExpBkCpyCart 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   7590
         TabIndex        =   371
         TabStop         =   0   'False
         Top             =   2040
         Width           =   7590
         Begin VB.OptionButton rbcExpBkCpyCart 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3600
            TabIndex        =   373
            Top             =   0
            Width           =   555
         End
         Begin VB.OptionButton rbcExpBkCpyCart 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2895
            TabIndex        =   372
            Top             =   0
            Width           =   645
         End
      End
      Begin VB.TextBox edcIFGroupNo 
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
         Left            =   1755
         MaxLength       =   1
         TabIndex        =   384
         Top             =   3015
         Width           =   390
      End
      Begin VB.PictureBox plcIFTime 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   8370
         TabIndex        =   380
         TabStop         =   0   'False
         Top             =   2760
         Width           =   8370
         Begin VB.OptionButton rbcIFTime 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   6990
            TabIndex        =   382
            Top             =   -15
            Width           =   555
         End
         Begin VB.OptionButton rbcIFTime 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   6330
            TabIndex        =   381
            Top             =   0
            Width           =   645
         End
      End
      Begin VB.PictureBox plcIFSelling 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   7590
         TabIndex        =   377
         TabStop         =   0   'False
         Top             =   2520
         Width           =   7590
         Begin VB.OptionButton rbcIFSelling 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   5970
            TabIndex        =   379
            Top             =   0
            Width           =   555
         End
         Begin VB.OptionButton rbcIFSelling 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   5280
            TabIndex        =   378
            Top             =   0
            Width           =   645
         End
      End
      Begin VB.PictureBox plcIFBulk 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   7590
         TabIndex        =   374
         TabStop         =   0   'False
         Top             =   2280
         Width           =   7590
         Begin VB.OptionButton rbcIFBulk 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   4920
            TabIndex        =   375
            Top             =   0
            Width           =   645
         End
         Begin VB.OptionButton rbcIFBulk 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   5595
            TabIndex        =   376
            Top             =   0
            Width           =   555
         End
      End
      Begin VB.TextBox edcIFDPNo 
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
         Left            =   7935
         MaxLength       =   1
         TabIndex        =   370
         Top             =   1575
         Width           =   390
      End
      Begin VB.TextBox edcIFProgCode 
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
         Left            =   6540
         MaxLength       =   5
         TabIndex        =   369
         Top             =   1575
         Width           =   915
      End
      Begin VB.TextBox edcIFZone 
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
         Left            =   5670
         MaxLength       =   3
         TabIndex        =   368
         Top             =   1575
         Width           =   615
      End
      Begin VB.TextBox edcIFDPNo 
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
         Left            =   7935
         MaxLength       =   1
         TabIndex        =   367
         Top             =   1245
         Width           =   390
      End
      Begin VB.TextBox edcIFProgCode 
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
         Left            =   6540
         MaxLength       =   5
         TabIndex        =   366
         Top             =   1245
         Width           =   915
      End
      Begin VB.TextBox edcIFZone 
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
         Left            =   5670
         MaxLength       =   3
         TabIndex        =   365
         Top             =   1245
         Width           =   615
      End
      Begin VB.TextBox edcIFDPNo 
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
         Left            =   7935
         MaxLength       =   1
         TabIndex        =   364
         Top             =   915
         Width           =   390
      End
      Begin VB.TextBox edcIFProgCode 
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
         Left            =   6540
         MaxLength       =   5
         TabIndex        =   363
         Top             =   915
         Width           =   915
      End
      Begin VB.TextBox edcIFZone 
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
         Left            =   5670
         MaxLength       =   3
         TabIndex        =   362
         Top             =   915
         Width           =   615
      End
      Begin VB.TextBox edcIFDPNo 
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
         Left            =   7935
         MaxLength       =   1
         TabIndex        =   361
         Top             =   585
         Width           =   390
      End
      Begin VB.TextBox edcIFProgCode 
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
         Left            =   6540
         MaxLength       =   5
         TabIndex        =   360
         Top             =   585
         Width           =   915
      End
      Begin VB.TextBox edcIFZone 
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
         Left            =   5670
         MaxLength       =   3
         TabIndex        =   359
         Top             =   585
         Width           =   615
      End
      Begin VB.TextBox edcIFPST 
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
         Left            =   4470
         MaxLength       =   7
         TabIndex        =   356
         Top             =   1575
         Width           =   750
      End
      Begin VB.TextBox edcIFPST 
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
         Left            =   3675
         MaxLength       =   7
         TabIndex        =   355
         Top             =   1575
         Width           =   750
      End
      Begin VB.TextBox edcIFPST 
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
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   354
         Top             =   1575
         Width           =   750
      End
      Begin VB.TextBox edcIFPST 
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
         Left            =   2085
         MaxLength       =   7
         TabIndex        =   353
         Top             =   1575
         Width           =   750
      End
      Begin VB.TextBox edcIFPST 
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
         Left            =   1290
         MaxLength       =   7
         TabIndex        =   352
         Top             =   1575
         Width           =   750
      End
      Begin VB.TextBox edcIFMST 
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
         Left            =   4470
         MaxLength       =   7
         TabIndex        =   350
         Top             =   1245
         Width           =   750
      End
      Begin VB.TextBox edcIFMST 
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
         Left            =   3675
         MaxLength       =   7
         TabIndex        =   349
         Top             =   1245
         Width           =   750
      End
      Begin VB.TextBox edcIFMST 
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
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   348
         Top             =   1245
         Width           =   750
      End
      Begin VB.TextBox edcIFMST 
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
         Left            =   2085
         MaxLength       =   7
         TabIndex        =   347
         Top             =   1245
         Width           =   750
      End
      Begin VB.TextBox edcIFMST 
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
         Left            =   1290
         MaxLength       =   7
         TabIndex        =   346
         Top             =   1245
         Width           =   750
      End
      Begin VB.TextBox edcIFCST 
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
         Left            =   4470
         MaxLength       =   7
         TabIndex        =   344
         Top             =   915
         Width           =   750
      End
      Begin VB.TextBox edcIFCST 
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
         Left            =   3675
         MaxLength       =   7
         TabIndex        =   343
         Top             =   915
         Width           =   750
      End
      Begin VB.TextBox edcIFCST 
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
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   342
         Top             =   915
         Width           =   750
      End
      Begin VB.TextBox edcIFCST 
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
         Left            =   2085
         MaxLength       =   7
         TabIndex        =   341
         Top             =   915
         Width           =   750
      End
      Begin VB.TextBox edcIFCST 
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
         Left            =   1290
         MaxLength       =   7
         TabIndex        =   340
         Top             =   915
         Width           =   750
      End
      Begin VB.TextBox edcIFEST 
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
         Left            =   4470
         MaxLength       =   7
         TabIndex        =   338
         Top             =   585
         Width           =   750
      End
      Begin VB.TextBox edcIFEST 
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
         Left            =   3675
         MaxLength       =   7
         TabIndex        =   337
         Top             =   585
         Width           =   750
      End
      Begin VB.TextBox edcIFEST 
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
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   336
         Top             =   585
         Width           =   750
      End
      Begin VB.TextBox edcIFEST 
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
         Left            =   2085
         MaxLength       =   7
         TabIndex        =   335
         Top             =   585
         Width           =   750
      End
      Begin VB.TextBox edcIFEST 
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
         Left            =   1290
         MaxLength       =   7
         TabIndex        =   334
         Top             =   585
         Width           =   750
      End
      Begin VB.Label lacCode 
         Appearance      =   0  'Flat
         Caption         =   "X-Digital: National ISCI Model                   Vehicle ID                        ISCI Prefix"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   405
         Top             =   5205
         Width           =   8340
      End
      Begin VB.Label lacCode 
         Appearance      =   0  'Flat
         Caption         =   $"Vehopt.frx":0000
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   407
         Top             =   5550
         Width           =   1725
      End
      Begin VB.Label lacCode 
         Appearance      =   0  'Flat
         Caption         =   $"Vehopt.frx":0089
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   397
         Top             =   4695
         Width           =   8340
      End
      Begin VB.Label lacCode 
         Appearance      =   0  'Flat
         Caption         =   "ARB Code                                    Contract Export Vehicle Extension #                          RADAR Code"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   385
         Top             =   3375
         Width           =   7605
      End
      Begin VB.Label lacEDASWindow 
         Appearance      =   0  'Flat
         Caption         =   "EDAS Time Window Duration                          seconds"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1845
         TabIndex        =   393
         Top             =   4020
         Width           =   4275
      End
      Begin VB.Label lacStnFdCode 
         Appearance      =   0  'Flat
         Caption         =   "Station Feed Export:  Vehicle ID "
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   389
         Top             =   3720
         Width           =   2625
      End
      Begin VB.Label lacIFGroupNo 
         Appearance      =   0  'Flat
         Caption         =   "Bulk Feed Group #             (Use Negative ""-"" to indicate exclude from Bulk Feed)"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   383
         Top             =   3045
         Width           =   7920
      End
      Begin VB.Label lacIFDP 
         Appearance      =   0  'Flat
         Caption         =   "Zone     Program Code    Daypart #"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   5775
         TabIndex        =   358
         Top             =   375
         Width           =   2940
      End
      Begin VB.Label lacIFDP 
         Appearance      =   0  'Flat
         Caption         =   "Program Code Daypart Mapping"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   5880
         TabIndex        =   357
         Top             =   180
         Width           =   2685
      End
      Begin VB.Label lacIFDP 
         Appearance      =   0  'Flat
         Caption         =   "PST Zone"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   405
         TabIndex        =   351
         Top             =   1605
         Width           =   870
      End
      Begin VB.Label lacIFDP 
         Appearance      =   0  'Flat
         Caption         =   "MST Zone"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   405
         TabIndex        =   345
         Top             =   1275
         Width           =   870
      End
      Begin VB.Label lacIFDP 
         Appearance      =   0  'Flat
         Caption         =   "CST Zone"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   405
         TabIndex        =   339
         Top             =   945
         Width           =   870
      End
      Begin VB.Label lacIFDP 
         Appearance      =   0  'Flat
         Caption         =   "1st            2nd           3rd            4th            5th"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1545
         TabIndex        =   332
         Top             =   375
         Width           =   3690
      End
      Begin VB.Label lacIFDP 
         Appearance      =   0  'Flat
         Caption         =   "Daypart End Times"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   2520
         TabIndex        =   331
         Top             =   180
         Width           =   1560
      End
      Begin VB.Label lacIFComm 
         Appearance      =   0  'Flat
         Caption         =   "Commercial Summary Report Information"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   330
         Top             =   0
         Width           =   4665
      End
      Begin VB.Label lacIFDP 
         Appearance      =   0  'Flat
         Caption         =   "EST Zone"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   405
         TabIndex        =   333
         Top             =   615
         Width           =   870
      End
   End
   Begin VB.PictureBox plcParticipant 
      Height          =   5160
      Left            =   11520
      ScaleHeight     =   5100
      ScaleWidth      =   8745
      TabIndex        =   171
      Top             =   1170
      Visible         =   0   'False
      Width           =   8805
      Begin V81Vehicle.CSI_Calendar csiParticipantDate 
         Height          =   315
         Left            =   4185
         TabIndex        =   174
         Top             =   675
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         Text            =   "03/15/2024"
         BorderStyle     =   1
         CSI_ShowDropDownOnFocus=   -1  'True
         CSI_InputBoxBoxAlignment=   0
         CSI_CalBackColor=   16777130
         CSI_CloseCalAfterSelection=   0   'False
         CSI_CalDateFormat=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CSI_CurDayBackColor=   16777215
         CSI_CurDayForeColor=   0
         CSI_ForceMondaySelectionOnly=   0   'False
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   -1  'True
         CSI_DefaultDateType=   4
      End
      Begin VB.CommandButton cmcClear 
         Appearance      =   0  'Flat
         Caption         =   "Clear"
         Height          =   240
         Left            =   7560
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   795
         Width           =   1050
      End
      Begin VB.PictureBox pbcPartSetFocus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   8505
         ScaleHeight     =   45
         ScaleWidth      =   45
         TabIndex        =   434
         TabStop         =   0   'False
         Top             =   540
         Width           =   45
      End
      Begin VB.PictureBox pbcArrow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         Picture         =   "Vehopt.frx":0116
         ScaleHeight     =   180
         ScaleWidth      =   105
         TabIndex        =   176
         TabStop         =   0   'False
         Top             =   1140
         Visible         =   0   'False
         Width           =   105
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
         Left            =   3150
         Picture         =   "Vehopt.frx":0420
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   2595
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox edcProdPct 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         HelpContextID   =   8
         Left            =   5580
         MaxLength       =   5
         TabIndex        =   182
         Top             =   2850
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox edcVehGpDropDown 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   420
         MaxLength       =   20
         TabIndex        =   179
         TabStop         =   0   'False
         Top             =   2595
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.TextBox edcSSDropDown 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   450
         MaxLength       =   20
         TabIndex        =   181
         TabStop         =   0   'False
         Top             =   3555
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.PictureBox pbcIntUpdateRvf 
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
         Left            =   5625
         ScaleHeight     =   210
         ScaleWidth      =   1260
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   2130
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.PictureBox pbcExtUpdateRvf 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   6315
         ScaleHeight     =   210
         ScaleWidth      =   1260
         TabIndex        =   186
         TabStop         =   0   'False
         Top             =   2445
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.ListBox lbcVehGp 
         Appearance      =   0  'Flat
         Height          =   240
         ItemData        =   "Vehopt.frx":051A
         Left            =   2385
         List            =   "Vehopt.frx":051C
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   2235
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.ListBox lbcSSource 
         Appearance      =   0  'Flat
         Height          =   240
         ItemData        =   "Vehopt.frx":051E
         Left            =   1650
         List            =   "Vehopt.frx":0520
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   1740
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.ComboBox cbcParticipant 
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
         Left            =   5310
         TabIndex        =   172
         Top             =   60
         Width           =   3315
      End
      Begin VB.PictureBox pbcParticipantSTab 
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
         Left            =   45
         ScaleHeight     =   60
         ScaleWidth      =   120
         TabIndex        =   177
         Top             =   225
         Width           =   120
      End
      Begin VB.PictureBox pbcParticipantTab 
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
         ScaleWidth      =   120
         TabIndex        =   187
         Top             =   4485
         Width           =   120
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdParticipant 
         Height          =   3795
         Left            =   135
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   1110
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   6694
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   16777215
         GridColor       =   -2147483635
         GridColorFixed  =   -2147483635
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.Label lacParticipant 
         Appearance      =   0  'Flat
         Caption         =   "Select either [New] or Date"
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
         Left            =   5325
         TabIndex        =   501
         Top             =   375
         Width           =   1950
      End
      Begin VB.Label lacParticipant 
         Appearance      =   0  'Flat
         Caption         =   "First Participant listed must be the Owner"
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
         Index           =   1
         Left            =   150
         TabIndex        =   435
         Top             =   4920
         Width           =   5790
      End
      Begin VB.Label lacParticipant 
         Appearance      =   0  'Flat
         Caption         =   "Effective Standard Broadcast Month Start Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   173
         Top             =   750
         Width           =   4080
      End
   End
   Begin VB.PictureBox plcSchedule 
      Height          =   1425
      Index           =   1
      Left            =   11160
      ScaleHeight     =   1365
      ScaleWidth      =   8235
      TabIndex        =   497
      Top             =   1680
      Visible         =   0   'False
      Width           =   8295
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
         Index           =   5
         Left            =   4215
         MaxLength       =   5
         TabIndex        =   499
         Top             =   345
         Width           =   1080
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "Break 1 @ 6:11; Break 2 @ 6:13: Distance of 120, then breaks are treated as one. Distance of 60, then breaks are separate"
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
         Height          =   225
         Index           =   10
         Left            =   75
         TabIndex        =   500
         Top             =   690
         Width           =   7935
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "+/- Max Distance to Adjacent Break (in seconds)"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   105
         TabIndex        =   498
         Top             =   360
         Width           =   3915
      End
   End
   Begin V81Vehicle.VehOptTabs udcVehOptTabs 
      Height          =   1110
      Left            =   435
      TabIndex        =   483
      Top             =   6825
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1958
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
      Height          =   5715
      Left            =   240
      ScaleHeight     =   5655
      ScaleWidth      =   10965
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   11025
      Begin V81Vehicle.CSI_ComboBoxList cbcCsiGeneric 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   41
         Top             =   2880
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   503
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin VB.TextBox edcGen 
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
         Left            =   6240
         TabIndex        =   43
         Top             =   2880
         Width           =   2625
      End
      Begin VB.TextBox edcGen 
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
         Left            =   9015
         MaxLength       =   4
         TabIndex        =   64
         Top             =   4290
         Width           =   675
      End
      Begin VB.TextBox edcGen 
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
         Left            =   10425
         MaxLength       =   2
         TabIndex        =   65
         Top             =   4290
         Width           =   345
      End
      Begin VB.Frame frcExport 
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   3
         Left            =   135
         TabIndex        =   18
         Top             =   720
         Width           =   9195
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "Unowned-Station"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   12
            Left            =   6735
            TabIndex        =   23
            Top             =   0
            Width           =   1905
         End
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "Unowned-Network"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   11
            Left            =   4680
            TabIndex        =   22
            Top             =   0
            Width           =   1890
         End
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "Owned-Station"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   10
            Left            =   2970
            TabIndex        =   21
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "Owned-Network"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   9
            Left            =   1095
            TabIndex        =   20
            Top             =   -15
            Width           =   1680
         End
         Begin VB.Label lacTitle 
            Caption         =   "Ownership"
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   19
            Top             =   -15
            Width           =   1305
         End
      End
      Begin VB.PictureBox plcShowAirTime 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   120
         ScaleHeight     =   270
         ScaleWidth      =   9120
         TabIndex        =   55
         Top             =   4020
         Width           =   9120
         Begin VB.OptionButton rbcShowAirDate 
            Caption         =   "Exact Date"
            Height          =   210
            Index           =   0
            Left            =   2550
            TabIndex        =   56
            Top             =   0
            Width           =   1410
         End
         Begin VB.OptionButton rbcShowAirDate 
            Caption         =   "Week of Date"
            Height          =   210
            Index           =   1
            Left            =   3975
            TabIndex        =   57
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.Frame frcExport 
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   2
         Left            =   3045
         TabIndex        =   6
         Top             =   135
         Width           =   5655
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "Embedded"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   7
            Left            =   1095
            TabIndex        =   8
            Top             =   0
            Width           =   1350
         End
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "ROS"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   8
            Left            =   2595
            TabIndex        =   9
            Top             =   0
            Width           =   960
         End
         Begin VB.Label lacTitle 
            Caption         =   "Delivery"
            Height          =   225
            Index           =   1
            Left            =   180
            TabIndex        =   7
            Top             =   0
            Width           =   1305
         End
      End
      Begin VB.TextBox edcGen 
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
         Left            =   6090
         MaxLength       =   1
         TabIndex        =   62
         Top             =   4290
         Width           =   345
      End
      Begin VB.Frame frcRemoteImport 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5355
         TabIndex        =   29
         Top             =   960
         Width           =   5490
         Begin VB.OptionButton rbcRemoteImport 
            Caption         =   "Affiliate Spots"
            Height          =   210
            Index           =   1
            Left            =   3000
            TabIndex        =   32
            Top             =   0
            Width           =   1470
         End
         Begin VB.OptionButton rbcRemoteImport 
            Caption         =   "Insertion Order"
            Height          =   210
            Index           =   0
            Left            =   1335
            TabIndex        =   31
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton rbcRemoteImport 
            Caption         =   "None"
            Height          =   210
            Index           =   2
            Left            =   4560
            TabIndex        =   33
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lacRemoteImport 
            Caption         =   "Remote Import"
            Height          =   240
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   1275
         End
      End
      Begin VB.PictureBox plcShowRateOnInsertion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4200
         ScaleHeight     =   240
         ScaleWidth      =   4425
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3195
         Width           =   4425
         Begin VB.OptionButton rbcShowRateOnInsertion 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3780
            TabIndex        =   49
            Top             =   0
            Width           =   570
         End
         Begin VB.OptionButton rbcShowRateOnInsertion 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   3120
            TabIndex        =   48
            Top             =   0
            Width           =   630
         End
      End
      Begin VB.PictureBox plcSplitCopy 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   3645
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3195
         Width           =   3645
         Begin VB.OptionButton rbcSplitCopy 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2325
            TabIndex        =   45
            Top             =   0
            Width           =   630
         End
         Begin VB.OptionButton rbcSplitCopy 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   2970
            TabIndex        =   46
            Top             =   0
            Width           =   570
         End
      End
      Begin VB.TextBox edcEMail 
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
         Left            =   780
         TabIndex        =   73
         Top             =   5340
         Width           =   7635
      End
      Begin VB.Frame frcRemoteExport 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   4815
         Begin VB.OptionButton rbcRemoteExport 
            Caption         =   "None"
            Height          =   210
            Index           =   2
            Left            =   3720
            TabIndex        =   28
            Top             =   0
            Width           =   720
         End
         Begin VB.OptionButton rbcRemoteExport 
            Caption         =   "Insertion Order"
            Height          =   210
            Index           =   0
            Left            =   1335
            TabIndex        =   26
            Top             =   0
            Width           =   1590
         End
         Begin VB.OptionButton rbcRemoteExport 
            Caption         =   "Log"
            Height          =   210
            Index           =   1
            Left            =   3000
            TabIndex        =   27
            Top             =   0
            Width           =   630
         End
         Begin VB.Label lacRemoteExport 
            Caption         =   "Remote Export"
            Height          =   240
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   1290
         End
      End
      Begin VB.TextBox edcGen 
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
         Index           =   10
         Left            =   2025
         MaxLength       =   3
         TabIndex        =   59
         Top             =   4305
         Width           =   675
      End
      Begin VB.ListBox lbcFeed 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   8325
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3510
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox edcGen 
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
         Left            =   6090
         MaxLength       =   20
         TabIndex        =   71
         Top             =   4995
         Width           =   2580
      End
      Begin VB.TextBox edcGen 
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
         Left            =   2205
         MaxLength       =   20
         TabIndex        =   70
         Top             =   4995
         Width           =   2580
      End
      Begin VB.TextBox edcGen 
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
         Left            =   6090
         MaxLength       =   20
         TabIndex        =   68
         Top             =   4650
         Width           =   2580
      End
      Begin VB.TextBox edcGen 
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
         Left            =   2205
         MaxLength       =   20
         TabIndex        =   67
         Top             =   4650
         Width           =   2580
      End
      Begin VB.TextBox edcGen 
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
         Left            =   4755
         MaxLength       =   4
         TabIndex        =   61
         Top             =   4290
         Width           =   675
      End
      Begin VB.PictureBox plcShowAirTime 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   450
         Index           =   0
         Left            =   120
         ScaleHeight     =   450
         ScaleWidth      =   8685
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3510
         Width           =   8685
         Begin VB.OptionButton rbcShowAirTime 
            Caption         =   "AvailTimes"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   4080
            TabIndex        =   54
            Top             =   225
            Width           =   1500
         End
         Begin VB.OptionButton rbcShowAirTime 
            Caption         =   "Daypart Times"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   2550
            TabIndex        =   53
            Top             =   225
            Width           =   1500
         End
         Begin VB.OptionButton rbcShowAirTime 
            Caption         =   "Hour-Separated Spot Times"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   4080
            TabIndex        =   52
            Top             =   -15
            Width           =   2595
         End
         Begin VB.OptionButton rbcShowAirTime 
            Caption         =   "Spot Times"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   2550
            TabIndex        =   51
            Top             =   -15
            Width           =   1245
         End
      End
      Begin VB.PictureBox plcGSAGroupNo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   105
         ScaleHeight     =   225
         ScaleWidth      =   1920
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   4290
         Width           =   1920
      End
      Begin VB.PictureBox pbcGTZTab 
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
         Left            =   75
         ScaleHeight     =   60
         ScaleWidth      =   90
         TabIndex        =   40
         Top             =   3195
         Width           =   90
      End
      Begin VB.PictureBox pbcGTZSTab 
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
         Left            =   105
         ScaleHeight     =   60
         ScaleWidth      =   15
         TabIndex        =   34
         Top             =   2010
         Width           =   15
      End
      Begin VB.PictureBox plcGMedium 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         ScaleHeight     =   210
         ScaleWidth      =   10725
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   10725
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "Podcast/Ad Server"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   6
            Left            =   7920
            TabIndex        =   17
            Top             =   0
            Width           =   1920
         End
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "Radio ROS Net"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   3090
            TabIndex        =   13
            Top             =   -15
            Width           =   1440
         End
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "Radio"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   810
            TabIndex        =   11
            Top             =   0
            Width           =   810
         End
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "TV"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   4770
            TabIndex        =   14
            Top             =   -15
            Width           =   690
         End
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "Cable"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   6960
            TabIndex        =   16
            Top             =   0
            Width           =   765
         End
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "Radio Net"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   1785
            TabIndex        =   12
            Top             =   -15
            Width           =   1065
         End
         Begin VB.OptionButton rbcGMedium 
            Caption         =   "TV Network"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   5475
            TabIndex        =   15
            Top             =   -15
            Width           =   1275
         End
      End
      Begin VB.PictureBox pbcGTZ 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1365
         Left            =   780
         Picture         =   "Vehopt.frx":0522
         ScaleHeight     =   1365
         ScaleWidth      =   8460
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1230
         Width           =   8460
         Begin VB.PictureBox pbcGTZToggle 
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
            Left            =   4065
            ScaleHeight     =   210
            ScaleWidth      =   870
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.CommandButton cmcGTZDropDown 
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
            Left            =   5490
            Picture         =   "Vehopt.frx":25ED8
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox edcGTZDropDown 
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
            Left            =   1590
            MaxLength       =   3
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   675
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.TextBox edcGen 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   5
         Top             =   120
         Width           =   1290
      End
      Begin VB.Label lacGEDI 
         Appearance      =   0  'Flat
         Caption         =   "Ad Server Vendor                                                                   Vendor Vehicle ID"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   42
         Top             =   2880
         Width           =   7770
      End
      Begin VB.Label lacGEDI 
         Appearance      =   0  'Flat
         Caption         =   "vCreative: Call Letters                    Band"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   7125
         TabIndex        =   63
         Top             =   4320
         Width           =   3240
      End
      Begin VB.Label lacEMail 
         Appearance      =   0  'Flat
         Caption         =   "E-Mail"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   72
         Top             =   5385
         Width           =   600
      End
      Begin VB.Label lacTitle 
         Appearance      =   0  'Flat
         Caption         =   "Time Zones"
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   0
         Left            =   195
         TabIndex        =   432
         Top             =   1635
         Width           =   630
      End
      Begin VB.Label lacFed 
         Appearance      =   0  'Flat
         Caption         =   "Fed: Feed zone indicate with Asterisk; Non-Feed zone indicate with First Letter of Feed Zone Letter of zone primary zone"
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
         Left            =   2940
         TabIndex        =   315
         Top             =   2625
         Width           =   5790
      End
      Begin VB.Label lacGBGL 
         Appearance      =   0  'Flat
         Caption         =   "Billed:       Revenue G/L #                                                                 Trade G/L #"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   69
         Top             =   5025
         Width           =   5910
      End
      Begin VB.Label lacGAGL 
         Appearance      =   0  'Flat
         Caption         =   "Accrued:  Revenue G/L #                                                                 Trade G/L #"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   66
         Top             =   4680
         Width           =   5910
      End
      Begin VB.Label lacGEDI 
         Appearance      =   0  'Flat
         Caption         =   "EDI: Call Letters                    Band"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   3390
         TabIndex        =   60
         Top             =   4320
         Width           =   2730
      End
      Begin VB.Label lacGSignOn 
         Appearance      =   0  'Flat
         Caption         =   "Sign-on Time"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   135
         Width           =   1365
      End
   End
   Begin VB.PictureBox plcSports 
      Height          =   5835
      Left            =   9495
      ScaleHeight     =   5775
      ScaleWidth      =   11175
      TabIndex        =   466
      Top             =   6045
      Visible         =   0   'False
      Width           =   11235
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
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   485
         Top             =   1695
         Width           =   2520
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
         Index           =   3
         Left            =   8640
         MaxLength       =   15
         TabIndex        =   484
         Top             =   1695
         Width           =   2520
      End
      Begin VB.ListBox lbcSeason 
         Height          =   1320
         ItemData        =   "Vehopt.frx":25FD2
         Left            =   6675
         List            =   "Vehopt.frx":25FD4
         TabIndex        =   478
         Top             =   330
         Width           =   2310
      End
      Begin V81Vehicle.CSI_RTFEdit edcTextFt 
         Height          =   1770
         Left            =   1200
         TabIndex        =   476
         Top             =   3945
         Width           =   9600
         _ExtentX        =   13732
         _ExtentY        =   3334
         Text            =   $"Vehopt.frx":25FD6
         FontName        =   ""
         FontSize        =   0
      End
      Begin V81Vehicle.CSI_RTFEdit edcTextHd 
         Height          =   1770
         Left            =   1200
         TabIndex        =   474
         Top             =   2070
         Width           =   9600
         _ExtentX        =   13732
         _ExtentY        =   3334
         Text            =   $"Vehopt.frx":26058
         FontName        =   ""
         FontSize        =   0
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
         Index           =   4
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   472
         Top             =   1335
         Width           =   1080
      End
      Begin VB.CheckBox ckcSch 
         Caption         =   "Agreement Pledge by Events"
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   470
         Top             =   1035
         Width           =   2850
      End
      Begin VB.CheckBox ckcSch 
         Caption         =   "Allow Sport Spots to be Moved/Fill to Non-Sports Vehicles"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   467
         Top             =   135
         Width           =   6015
      End
      Begin VB.CheckBox ckcSch 
         Caption         =   "Allow Sport Spots to be Moved/Fill to other Sports Vehicles"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   468
         Top             =   435
         Width           =   6315
      End
      Begin VB.CheckBox ckcSch 
         Caption         =   "Allow Spots from Non-Sports Vehicles to be Moved/Fill to Sport Vehicle"
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   469
         Top             =   735
         Width           =   6270
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "Event Title 1 (Ex: Visiting Team)"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   487
         Top             =   1740
         Width           =   2730
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "@ Event Title 2 (Ex: Home Team)"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   5895
         TabIndex        =   486
         Top             =   1740
         Width           =   2730
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "Default Season"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   6675
         TabIndex        =   477
         Top             =   90
         Width           =   2250
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "Pledge Clearance Required Count"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   471
         Top             =   1380
         Width           =   2955
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "Pledge Declaration Footer"
         ForeColor       =   &H80000008&
         Height          =   690
         Index           =   4
         Left            =   120
         TabIndex        =   475
         Top             =   3900
         Width           =   990
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "Pledge Declaration Header"
         ForeColor       =   &H80000008&
         Height          =   900
         Index           =   3
         Left            =   120
         TabIndex        =   473
         Top             =   2055
         Width           =   1095
      End
   End
   Begin VB.PictureBox plcExport 
      Height          =   5880
      Left            =   4515
      ScaleHeight     =   5820
      ScaleWidth      =   10890
      TabIndex        =   438
      Top             =   7275
      Width           =   10950
      Begin VB.Frame frcExport 
         Caption         =   "Affiliate"
         Height          =   5685
         Index           =   1
         Left            =   5280
         TabIndex        =   457
         Top             =   90
         Width           =   5580
         Begin VB.CheckBox ckcAffExport 
            Caption         =   "Station Compensation"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   458
            Top             =   375
            Width           =   2565
         End
         Begin VB.TextBox edcExport 
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
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   463
            Top             =   1370
            Width           =   345
         End
         Begin VB.CheckBox ckcAffExport 
            Caption         =   "Wegener-iPump"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   461
            Top             =   1410
            Width           =   1815
         End
         Begin VB.CheckBox ckcAffExport 
            Caption         =   "Wegener-Compel"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   460
            Top             =   1065
            Width           =   1935
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Pledge Vs Air (CSV)"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   459
            Top             =   720
            Width           =   2160
         End
         Begin VB.Label lacCode 
            Appearance      =   0  'Flat
            Caption         =   "Media Code Event Type Override"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   2070
            TabIndex        =   462
            Top             =   1410
            Width           =   2805
         End
      End
      Begin VB.Frame frcExport 
         Caption         =   "Traffic"
         Height          =   5685
         Index           =   0
         Left            =   30
         TabIndex        =   437
         Top             =   90
         Width           =   5235
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Custom Revenue Export"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   20
            Left            =   120
            TabIndex        =   153
            Top             =   3135
            Width           =   2535
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "RAB"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   19
            Left            =   2760
            TabIndex        =   453
            Top             =   1410
            Width           =   930
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Tableau"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   18
            Left            =   2760
            TabIndex        =   454
            Top             =   720
            Width           =   1470
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Jelli"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   17
            Left            =   120
            TabIndex        =   451
            Top             =   5205
            Width           =   1470
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Matrix"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   16
            Left            =   2760
            TabIndex        =   452
            Top             =   375
            Width           =   1470
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "CnC Network Inventory"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   15
            Left            =   120
            TabIndex        =   442
            Top             =   1755
            Width           =   2265
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Linkup"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   14
            Left            =   120
            TabIndex        =   448
            Top             =   4170
            Width           =   1470
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Enigneering Feed - ESPN"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   12
            Left            =   120
            TabIndex        =   450
            Top             =   4860
            Width           =   2505
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Enco"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   11
            Left            =   120
            TabIndex        =   447
            Top             =   3825
            Width           =   1080
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "CnC Spots"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   443
            Top             =   2100
            Width           =   1725
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Audio MP2"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   9
            Left            =   120
            TabIndex        =   440
            Top             =   1065
            Width           =   1725
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Audio ISCI Titles"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   8
            Left            =   120
            TabIndex        =   439
            Top             =   720
            Width           =   1725
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "AirWave"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   7
            Left            =   120
            TabIndex        =   169
            Top             =   375
            Width           =   1725
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Clearance Spots"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   441
            Top             =   1410
            Width           =   1725
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Dallas Feed"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   446
            Top             =   3480
            Width           =   1335
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Phoenix Log"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   2760
            TabIndex        =   455
            Top             =   1065
            Width           =   1470
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "New York EAS Feed"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   2700
            TabIndex        =   456
            Top             =   5205
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Commercial Change"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   120
            TabIndex        =   444
            Top             =   2445
            Width           =   2055
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Corporate"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   445
            Top             =   2790
            Width           =   1230
         End
         Begin VB.CheckBox ckcIFExport 
            Caption         =   "Engineering Feed - ASP"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   6
            Left            =   120
            TabIndex        =   449
            Top             =   4515
            Width           =   2505
         End
      End
   End
   Begin VB.PictureBox plcGreatPlains 
      Height          =   3720
      Left            =   9870
      ScaleHeight     =   3660
      ScaleWidth      =   8745
      TabIndex        =   151
      Top             =   5640
      Visible         =   0   'False
      Width           =   8805
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
         Index           =   8
         Left            =   1230
         MaxLength       =   15
         TabIndex        =   168
         Top             =   3285
         Width           =   1725
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
         Index           =   7
         Left            =   3045
         MaxLength       =   10
         TabIndex        =   167
         Top             =   2850
         Width           =   1500
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
         Index           =   6
         Left            =   3045
         MaxLength       =   10
         TabIndex        =   165
         Top             =   2460
         Width           =   1500
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
         Index           =   5
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   170
         Top             =   2070
         Width           =   1500
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
         Index           =   4
         Left            =   3045
         MaxLength       =   10
         TabIndex        =   163
         Top             =   1680
         Width           =   1500
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
         Index           =   3
         Left            =   4305
         MaxLength       =   10
         TabIndex        =   161
         Top             =   1290
         Width           =   1470
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
         Index           =   2
         Left            =   3045
         MaxLength       =   10
         TabIndex        =   159
         Top             =   900
         Width           =   1500
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
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   157
         Top             =   510
         Width           =   1500
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
         Left            =   1485
         MaxLength       =   10
         TabIndex        =   155
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label lacGP 
         Appearance      =   0  'Flat
         Caption         =   "Vendor ID"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   464
         Top             =   3285
         Width           =   990
      End
      Begin VB.Label lacGP 
         Appearance      =   0  'Flat
         Caption         =   "Primary Code Receivables - Trade"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   465
         Top             =   2895
         Width           =   2925
      End
      Begin VB.Label lacGP 
         Appearance      =   0  'Flat
         Caption         =   "Primary Code Gross Sales - Trade"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   166
         Top             =   2505
         Width           =   2835
      End
      Begin VB.Label lacGP 
         Appearance      =   0  'Flat
         Caption         =   "Primary Code Receivables - Cash"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   162
         Top             =   1740
         Width           =   2835
      End
      Begin VB.Label lacGP 
         Appearance      =   0  'Flat
         Caption         =   "Primary Code Agency Commission - Cash/Trade"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   160
         Top             =   1335
         Width           =   4155
      End
      Begin VB.Label lacGP 
         Appearance      =   0  'Flat
         Caption         =   "Primary Code Gross Sales - Cash"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   158
         Top             =   945
         Width           =   2835
      End
      Begin VB.Label lacGP 
         Appearance      =   0  'Flat
         Caption         =   "Branch Code - Trade"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   164
         Top             =   2115
         Width           =   1845
      End
      Begin VB.Label lacGP 
         Appearance      =   0  'Flat
         Caption         =   "Branch Code - Cash"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   156
         Top             =   555
         Width           =   1755
      End
      Begin VB.Label lacGP 
         Appearance      =   0  'Flat
         Caption         =   "Division Code"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   154
         Top             =   165
         Width           =   1215
      End
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6645
      TabIndex        =   1
      Top             =   0
      Width           =   3930
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   6
      Left            =   7560
      Top             =   5955
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
      Left            =   5340
      TabIndex        =   314
      TabStop         =   0   'False
      Top             =   -120
      Visible         =   0   'False
      Width           =   435
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
      TabIndex        =   429
      TabStop         =   0   'False
      Top             =   5025
      Width           =   75
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Height          =   240
      Left            =   7050
      TabIndex        =   428
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
      Height          =   240
      Left            =   5925
      TabIndex        =   427
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   240
      Left            =   4815
      TabIndex        =   426
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   240
      Left            =   3690
      TabIndex        =   425
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   240
      Left            =   2565
      TabIndex        =   424
      Top             =   6930
      Width           =   1050
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10590
      Top             =   390
   End
   Begin VB.PictureBox plcSales 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   10095
      ScaleHeight     =   5760
      ScaleWidth      =   10740
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   5430
      Visible         =   0   'False
      Width           =   10800
      Begin VB.PictureBox plcAudioType 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   5460
         ScaleHeight     =   435
         ScaleWidth      =   5355
         TabIndex        =   126
         Top             =   3735
         Width           =   5355
         Begin VB.OptionButton rbcAudioType 
            Caption         =   "Pre-Rec Promo"
            Height          =   210
            Index           =   5
            Left            =   3540
            TabIndex        =   133
            Top             =   195
            Width           =   1590
         End
         Begin VB.OptionButton rbcAudioType 
            Caption         =   "Pre-Rec Cmml"
            Height          =   210
            Index           =   4
            Left            =   2010
            TabIndex        =   132
            Top             =   195
            Width           =   1485
         End
         Begin VB.OptionButton rbcAudioType 
            Caption         =   "Rec Promo"
            Height          =   210
            Index           =   3
            Left            =   720
            TabIndex        =   131
            Top             =   195
            Width           =   1215
         End
         Begin VB.OptionButton rbcAudioType 
            Caption         =   "Live Promo"
            Height          =   210
            Index           =   2
            Left            =   3540
            TabIndex        =   130
            Top             =   0
            Width           =   1290
         End
         Begin VB.OptionButton rbcAudioType 
            Caption         =   "Rec Cmml"
            Height          =   210
            Index           =   0
            Left            =   720
            TabIndex        =   128
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton rbcAudioType 
            Caption         =   "Live Cmml"
            Height          =   210
            Index           =   1
            Left            =   2010
            TabIndex        =   129
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lacSales 
            Appearance      =   0  'Flat
            Caption         =   "Audio Default"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   0
            TabIndex        =   127
            Top             =   -15
            Width           =   690
         End
      End
      Begin VB.TextBox edcBB 
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
         Left            =   990
         TabIndex        =   152
         Top             =   5385
         Width           =   7425
      End
      Begin VB.TextBox edcBB 
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
         Left            =   990
         TabIndex        =   150
         Top             =   5010
         Width           =   7425
      End
      Begin VB.ComboBox cbcMedia 
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
         Left            =   4200
         TabIndex        =   125
         Top             =   3795
         Width           =   960
      End
      Begin VB.PictureBox plcLGrid 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   6885
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   4725
         Width           =   6885
         Begin VB.OptionButton rbcLGrid 
            Caption         =   "Half"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   5055
            TabIndex        =   147
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton rbcLGrid 
            Caption         =   "Full"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   4380
            TabIndex        =   146
            Top             =   0
            Width           =   585
         End
         Begin VB.OptionButton rbcLGrid 
            Caption         =   "Quarter"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   5730
            TabIndex        =   148
            Top             =   0
            Width           =   930
         End
      End
      Begin VB.PictureBox pbcSalesSTab 
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
         Left            =   -45
         ScaleHeight     =   60
         ScaleWidth      =   120
         TabIndex        =   112
         Top             =   5535
         Width           =   120
      End
      Begin VB.TextBox edcSec 
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
         Left            =   6720
         MaxLength       =   4
         TabIndex        =   143
         Top             =   4335
         Width           =   480
      End
      Begin VB.TextBox edcYear 
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   136
         Top             =   4335
         Width           =   840
      End
      Begin VB.PictureBox plcInvBy 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2040
         ScaleHeight     =   255
         ScaleWidth      =   2055
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   4380
         Width           =   2055
         Begin VB.OptionButton rbcInvBy 
            Caption         =   "Year"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   139
            Top             =   0
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton rbcInvBy 
            Caption         =   "Week"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   138
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.CheckBox ckcRollover 
         Caption         =   "Rollover"
         Height          =   240
         Left            =   7560
         TabIndex        =   144
         Top             =   4380
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox edcInventory 
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
         Left            =   5520
         MaxLength       =   4
         TabIndex        =   141
         Top             =   4335
         Width           =   720
      End
      Begin VB.PictureBox plcBillSA 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   6255
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   3555
         Width           =   6255
         Begin VB.OptionButton rbcBillSA 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1980
            TabIndex        =   122
            Top             =   0
            Width           =   630
         End
         Begin VB.OptionButton rbcBillSA 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   123
            Top             =   0
            Width           =   570
         End
      End
      Begin VB.PictureBox plcSAdvtSep 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5145
         ScaleHeight     =   240
         ScaleWidth      =   3750
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   75
         Width           =   3750
         Begin VB.OptionButton rbcSAdvtSep 
            Caption         =   "Time"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2100
            TabIndex        =   79
            Top             =   0
            Width           =   720
         End
         Begin VB.OptionButton rbcSAdvtSep 
            Caption         =   "Break"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2835
            TabIndex        =   80
            Top             =   0
            Width           =   780
         End
      End
      Begin VB.PictureBox plcSMoveLLD 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   6255
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   3300
         Width           =   6255
         Begin VB.OptionButton rbcSMoveLLD 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   5610
            TabIndex        =   120
            Top             =   0
            Width           =   540
         End
         Begin VB.OptionButton rbcSMoveLLD 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   4890
            TabIndex        =   119
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.TextBox edcSLen 
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
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   117
         Top             =   2820
         Width           =   630
      End
      Begin VB.PictureBox plcSSalesperson 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   6960
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   2475
         Width           =   6960
         Begin VB.OptionButton rbcSSalesperson 
            Caption         =   "on Collections"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   5115
            TabIndex        =   109
            Top             =   0
            Width           =   1530
         End
         Begin VB.OptionButton rbcSSalesperson 
            Caption         =   "on Billing"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   4050
            TabIndex        =   108
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.PictureBox plcSMove 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   8025
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   1350
         Width           =   8025
         Begin VB.OptionButton rbcSMove 
            Caption         =   "MG's"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   4050
            TabIndex        =   96
            Top             =   0
            Width           =   720
         End
         Begin VB.OptionButton rbcSMove 
            Caption         =   "MG's or Outsides"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   4845
            TabIndex        =   97
            Top             =   0
            Width           =   1800
         End
      End
      Begin VB.PictureBox plcSOverbook 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   4035
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1110
         Width           =   4035
         Begin VB.OptionButton rbcSOverbook 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3060
            TabIndex        =   94
            Top             =   0
            Width           =   540
         End
         Begin VB.OptionButton rbcSOverbook 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2415
            TabIndex        =   93
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox plcSSellout 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   8550
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   870
         Width           =   8550
         Begin VB.OptionButton rbcSSellout 
            Caption         =   "Matching Seconds"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   4140
            TabIndex        =   91
            Top             =   0
            Width           =   1860
         End
         Begin VB.OptionButton rbcSSellout 
            Caption         =   "Units"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1545
            TabIndex        =   88
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton rbcSSellout 
            Caption         =   "Units and Seconds"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2280
            TabIndex        =   89
            Top             =   0
            Width           =   1830
         End
         Begin VB.OptionButton rbcSSellout 
            Caption         =   "30 Second Units"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   7020
            TabIndex        =   90
            Top             =   -15
            Visible         =   0   'False
            Width           =   1605
         End
      End
      Begin VB.PictureBox plcSCompetitive 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   8250
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   315
         Width           =   8250
         Begin VB.OptionButton rbcSCompetitive 
            Caption         =   "not Back to Back"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   5355
            TabIndex        =   84
            Top             =   0
            Width           =   1650
         End
         Begin VB.OptionButton rbcSCompetitive 
            Caption         =   "Break"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   4530
            TabIndex        =   83
            Top             =   0
            Width           =   780
         End
         Begin VB.OptionButton rbcSCompetitive 
            Caption         =   "Separation Length"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2640
            TabIndex        =   82
            Top             =   0
            Width           =   1830
         End
      End
      Begin VB.PictureBox plcSCommission 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   3825
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   45
         Width           =   3825
         Begin VB.OptionButton rbcSCommission 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2520
            TabIndex        =   76
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton rbcSCommission 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3225
            TabIndex        =   77
            Top             =   0
            Width           =   540
         End
      End
      Begin VB.PictureBox pbcSBreak 
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
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   8715
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   1605
         Width           =   8715
         Begin VB.OptionButton rbcSBreak 
            Caption         =   "N/A"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   6
            Left            =   2145
            TabIndex        =   106
            Top             =   630
            Width           =   645
         End
         Begin VB.OptionButton rbcSBreak 
            Caption         =   "Announcers with Separate Groups"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   5
            Left            =   5430
            TabIndex        =   105
            Top             =   420
            Width           =   3165
         End
         Begin VB.OptionButton rbcSBreak 
            Caption         =   "Announcers with Clustered Groups"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   2145
            TabIndex        =   104
            Top             =   420
            Width           =   3255
         End
         Begin VB.OptionButton rbcSBreak 
            Caption         =   "Media Code with Separate Groups"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   5430
            TabIndex        =   101
            Top             =   0
            Width           =   3135
         End
         Begin VB.OptionButton rbcSBreak 
            Caption         =   "Media Code with Clustered Groups"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2145
            TabIndex        =   100
            Top             =   0
            Width           =   3210
         End
         Begin VB.OptionButton rbcSBreak 
            Caption         =   "Lengths with Clustered Groups"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   2145
            TabIndex        =   102
            Top             =   210
            Width           =   3165
         End
         Begin VB.OptionButton rbcSBreak 
            Caption         =   "Lengths with Separate Groups"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   5430
            TabIndex        =   103
            Top             =   210
            Width           =   3030
         End
         Begin VB.Label lacSales 
            Appearance      =   0  'Flat
            Caption         =   "In Breaks, order Spots by"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   99
            Top             =   0
            Width           =   2130
         End
      End
      Begin VB.PictureBox pbcSalesTab 
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
         Left            =   45
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   115
         Top             =   4410
         Width           =   105
      End
      Begin VB.PictureBox pbcSSpotLen 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   1425
         Picture         =   "Vehopt.frx":260DA
         ScaleHeight     =   450
         ScaleWidth      =   5610
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   2760
         Width           =   5610
         Begin VB.TextBox edcSSpotLG 
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
            Left            =   600
            MaxLength       =   4
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   210
            Visible         =   0   'False
            Width           =   540
         End
      End
      Begin VB.TextBox edcSCompSepLen 
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
         Left            =   3420
         MaxLength       =   10
         TabIndex        =   86
         Top             =   555
         Width           =   1080
      End
      Begin VB.Label lacBB 
         Appearance      =   0  'Flat
         Caption         =   "Close BB"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   502
         Top             =   5430
         Width           =   780
      End
      Begin VB.Label lacBB 
         Appearance      =   0  'Flat
         Caption         =   "Open BB"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   149
         Top             =   5055
         Width           =   780
      End
      Begin VB.Label lacSales 
         Caption         =   "Restrict Copy Assignment to Media Definition"
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   124
         Top             =   3825
         Width           =   3960
      End
      Begin VB.Label lacSec 
         Caption         =   "Sec"
         Height          =   240
         Left            =   6360
         TabIndex        =   142
         Top             =   4380
         Width           =   495
      End
      Begin VB.Label lacSales 
         Caption         =   "Network Inventory-"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   134
         Top             =   4095
         Width           =   1695
      End
      Begin VB.Label lacInventory 
         Caption         =   "Inventory- Min"
         Height          =   240
         Left            =   4200
         TabIndex        =   140
         Top             =   4380
         Width           =   1215
      End
      Begin VB.Label lacYear 
         Caption         =   "Year"
         Height          =   240
         Left            =   480
         TabIndex        =   135
         Top             =   4380
         Width           =   495
      End
      Begin VB.Label lacSales 
         Appearance      =   0  'Flat
         Caption         =   "Default Length"
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   3
         Left            =   7215
         TabIndex        =   116
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label lacSLength 
         Appearance      =   0  'Flat
         Caption         =   "Group #s"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   465
         TabIndex        =   111
         Top             =   2985
         Width           =   870
      End
      Begin VB.Label lacSales 
         Appearance      =   0  'Flat
         Caption         =   "Separation Length, Product Protection"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   85
         Top             =   570
         Width           =   3285
      End
      Begin VB.Label lacSLength 
         Appearance      =   0  'Flat
         Caption         =   "Spot Lengths"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   110
         Top             =   2775
         Width           =   1275
      End
   End
   Begin VB.PictureBox plcLog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5745
      Left            =   10200
      ScaleHeight     =   5685
      ScaleWidth      =   10905
      TabIndex        =   251
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   10965
      Begin VB.TextBox edcLog 
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
         Left            =   2595
         MaxLength       =   250
         TabIndex        =   313
         Top             =   5310
         Width           =   6000
      End
      Begin VB.TextBox edcLNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   9390
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   293
         Top             =   2355
         Visible         =   0   'False
         Width           =   7800
      End
      Begin VB.CheckBox ckcLog 
         Caption         =   "For Airing Vehicles, Honor Zero Units"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   6810
         TabIndex        =   296
         Top             =   3330
         Width           =   3435
      End
      Begin VB.TextBox edcLog 
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
         Left            =   5940
         MaxLength       =   40
         TabIndex        =   307
         Top             =   4005
         Width           =   3885
      End
      Begin VB.PictureBox pbcMerge 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   120
         ScaleHeight     =   270
         ScaleWidth      =   3480
         TabIndex        =   303
         Top             =   4020
         Width           =   3480
         Begin VB.OptionButton rbcMerge 
            Caption         =   "Merge"
            Height          =   210
            Index           =   4
            Left            =   1110
            TabIndex        =   304
            Top             =   0
            Width           =   1065
         End
         Begin VB.OptionButton rbcMerge 
            Caption         =   "Separate"
            Height          =   210
            Index           =   5
            Left            =   2175
            TabIndex        =   305
            Top             =   -15
            Width           =   1065
         End
      End
      Begin VB.PictureBox pbcMerge 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   5115
         ScaleHeight     =   270
         ScaleWidth      =   3120
         TabIndex        =   300
         Top             =   3735
         Width           =   3120
         Begin VB.OptionButton rbcMerge 
            Caption         =   "Merge"
            Height          =   210
            Index           =   2
            Left            =   780
            TabIndex        =   301
            Top             =   0
            Width           =   1065
         End
         Begin VB.OptionButton rbcMerge 
            Caption         =   "Separate"
            Height          =   210
            Index           =   3
            Left            =   1830
            TabIndex        =   302
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.PictureBox pbcMerge 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   120
         ScaleHeight     =   270
         ScaleWidth      =   3570
         TabIndex        =   297
         Top             =   3735
         Width           =   3570
         Begin VB.OptionButton rbcMerge 
            Caption         =   "Separate"
            Height          =   210
            Index           =   1
            Left            =   2175
            TabIndex        =   299
            Top             =   0
            Width           =   1065
         End
         Begin VB.OptionButton rbcMerge 
            Caption         =   "Merge"
            Height          =   210
            Index           =   0
            Left            =   1110
            TabIndex        =   298
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.PictureBox plcGenLog 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   10125
         TabIndex        =   252
         TabStop         =   0   'False
         Top             =   75
         Width           =   10125
         Begin VB.OptionButton rbcGenLog 
            Caption         =   "Live Log or Merge"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   8310
            TabIndex        =   436
            Top             =   0
            Width           =   1815
         End
         Begin VB.OptionButton rbcGenLog 
            Caption         =   "Merge into Pre-empt Vehicle"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   5580
            TabIndex        =   256
            Top             =   0
            Width           =   2685
         End
         Begin VB.OptionButton rbcGenLog 
            Caption         =   "Live Log && Print"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   3930
            TabIndex        =   255
            Top             =   0
            Width           =   1650
         End
         Begin VB.OptionButton rbcGenLog 
            Caption         =   "None"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3150
            TabIndex        =   254
            Top             =   0
            Width           =   720
         End
         Begin VB.OptionButton rbcGenLog 
            Caption         =   "Print"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   253
            Top             =   0
            Width           =   750
         End
      End
      Begin VB.PictureBox pbcLYN 
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
         Left            =   1980
         ScaleHeight     =   210
         ScaleWidth      =   540
         TabIndex        =   292
         TabStop         =   0   'False
         Top             =   2745
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox edcLDropDown 
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
         Left            =   690
         MaxLength       =   3
         TabIndex        =   291
         TabStop         =   0   'False
         Top             =   2670
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.PictureBox pbcLSTab 
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
         Left            =   -15
         ScaleHeight     =   135
         ScaleWidth      =   60
         TabIndex        =   289
         Top             =   2640
         Width           =   60
      End
      Begin VB.PictureBox pbcLTab 
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
         Left            =   45
         ScaleHeight     =   120
         ScaleWidth      =   15
         TabIndex        =   294
         Top             =   3450
         Width           =   15
      End
      Begin VB.PictureBox pbcLogForm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   120
         Picture         =   "Vehopt.frx":276A8
         ScaleHeight     =   960
         ScaleWidth      =   8310
         TabIndex        =   290
         TabStop         =   0   'False
         Top             =   2280
         Width           =   8310
      End
      Begin VB.TextBox edcLAffDate 
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
         Left            =   2190
         MaxLength       =   10
         TabIndex        =   287
         Top             =   1815
         Width           =   1155
      End
      Begin VB.TextBox edcLAffDate 
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
         Left            =   5460
         MaxLength       =   10
         TabIndex        =   288
         Top             =   1815
         Width           =   1155
      End
      Begin VB.PictureBox plcLCopyOnAir 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6450
         ScaleHeight     =   255
         ScaleWidth      =   3900
         TabIndex        =   283
         TabStop         =   0   'False
         Top             =   1530
         Visible         =   0   'False
         Width           =   3900
         Begin VB.OptionButton rbcLCopyOnAir 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   3060
            TabIndex        =   285
            Top             =   0
            Width           =   555
         End
         Begin VB.OptionButton rbcLCopyOnAir 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   2430
            TabIndex        =   284
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.TextBox edcLLDCpyAsgn 
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
         Left            =   2235
         MaxLength       =   10
         TabIndex        =   258
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox edcLDate 
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
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   262
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox edcLDate 
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
         Left            =   5235
         MaxLength       =   10
         TabIndex        =   260
         Top             =   360
         Width           =   1155
      End
      Begin VB.PictureBox plcLAffTimes 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         ScaleHeight     =   300
         ScaleWidth      =   6360
         TabIndex        =   278
         TabStop         =   0   'False
         Top             =   1155
         Width           =   6360
         Begin VB.ComboBox cbcSvLog 
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
            Index           =   2
            Left            =   5055
            TabIndex        =   281
            Top             =   0
            Width           =   960
         End
         Begin VB.ComboBox cbcSvLog 
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
            Index           =   1
            Left            =   3510
            TabIndex        =   280
            Top             =   0
            Width           =   960
         End
         Begin VB.ComboBox cbcSvLog 
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
            Index           =   0
            Left            =   2040
            TabIndex        =   279
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox plcLAffCPs 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         ScaleHeight     =   300
         ScaleWidth      =   6105
         TabIndex        =   274
         TabStop         =   0   'False
         Top             =   795
         Width           =   6105
         Begin VB.ComboBox cbcLog 
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
            Index           =   2
            Left            =   5055
            TabIndex        =   277
            Top             =   0
            Width           =   960
         End
         Begin VB.ComboBox cbcLog 
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
            Index           =   1
            Left            =   3510
            TabIndex        =   276
            Top             =   0
            Width           =   960
         End
         Begin VB.ComboBox cbcLog 
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
            Index           =   0
            Left            =   2040
            TabIndex        =   275
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox plcLTiming 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6225
         ScaleHeight     =   255
         ScaleWidth      =   3060
         TabIndex        =   271
         TabStop         =   0   'False
         Top             =   375
         Visible         =   0   'False
         Width           =   3060
         Begin VB.OptionButton rbcLTiming 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2235
            TabIndex        =   273
            Top             =   0
            Width           =   540
         End
         Begin VB.OptionButton rbcLTiming 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1575
            TabIndex        =   272
            Top             =   0
            Width           =   570
         End
      End
      Begin VB.PictureBox plcLDaylight 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6195
         ScaleHeight     =   255
         ScaleWidth      =   3060
         TabIndex        =   268
         TabStop         =   0   'False
         Top             =   195
         Visible         =   0   'False
         Width           =   3060
         Begin VB.OptionButton rbcLDaylight 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   2385
            TabIndex        =   270
            Top             =   0
            Width           =   540
         End
         Begin VB.OptionButton rbcLDaylight 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   1680
            TabIndex        =   269
            Top             =   0
            Width           =   570
         End
      End
      Begin VB.PictureBox plcLZone 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6240
         ScaleHeight     =   255
         ScaleWidth      =   4860
         TabIndex        =   263
         TabStop         =   0   'False
         Top             =   15
         Visible         =   0   'False
         Width           =   4860
         Begin VB.OptionButton rbcLZone 
            Caption         =   "Eastern"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   990
            TabIndex        =   264
            Top             =   0
            Width           =   930
         End
         Begin VB.OptionButton rbcLZone 
            Caption         =   "Central"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   1935
            TabIndex        =   265
            Top             =   0
            Width           =   870
         End
         Begin VB.OptionButton rbcLZone 
            Caption         =   "Mountain"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   2865
            TabIndex        =   266
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton rbcLZone 
            Caption         =   "Pacific"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   3900
            TabIndex        =   267
            Top             =   0
            Width           =   930
         End
      End
      Begin VB.CheckBox ckcUnsoldBlank 
         Caption         =   "Retain Unsold Time if No Replacement exist for Suppressed Spot"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   295
         Top             =   3330
         Width           =   6045
      End
      Begin VB.TextBox edcLog 
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
         Left            =   2610
         MaxLength       =   250
         TabIndex        =   309
         Top             =   4440
         Width           =   6000
      End
      Begin VB.TextBox edcLog 
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
         Left            =   2610
         MaxLength       =   250
         TabIndex        =   311
         Top             =   4875
         Width           =   6000
      End
      Begin VB.CheckBox ckcLog 
         Caption         =   "Suppress Rotation Comments on Logs"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   282
         Top             =   1530
         Width           =   4500
      End
      Begin VB.Label lacLog 
         Appearance      =   0  'Flat
         Caption         =   "Log Export Drive/Path"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   312
         Top             =   5355
         Width           =   2490
      End
      Begin VB.Label lacLog 
         Appearance      =   0  'Flat
         Caption         =   "Override Affidavit Name"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   3690
         TabIndex        =   306
         Top             =   4020
         Width           =   2490
      End
      Begin VB.Label lacLog 
         Appearance      =   0  'Flat
         Caption         =   "Automation Import Drive/Path"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   310
         Top             =   4920
         Width           =   2490
      End
      Begin VB.Label lacLog 
         Appearance      =   0  'Flat
         Caption         =   "Automation Export Drive/Path"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   308
         Top             =   4485
         Width           =   2490
      End
      Begin VB.Label lacLAffLogDate 
         Appearance      =   0  'Flat
         Caption         =   "Affiliate: Last Log Date                                                Last C.P. Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   286
         Top             =   1845
         Width           =   5190
      End
      Begin VB.Label lacLLDCpyAsgn 
         Appearance      =   0  'Flat
         Caption         =   "Last Date Copy Assigned"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   257
         Top             =   405
         Width           =   2100
      End
      Begin VB.Label lacLLogDate 
         Appearance      =   0  'Flat
         Caption         =   "Preliminary"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   6540
         TabIndex        =   261
         Top             =   390
         Width           =   1020
      End
      Begin VB.Label lacLLogDate 
         Appearance      =   0  'Flat
         Caption         =   "Last Log Date: Final Log Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   3600
         TabIndex        =   259
         Top             =   390
         Width           =   1635
      End
   End
   Begin VB.PictureBox plcPSAPromo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4980
      Left            =   10410
      ScaleHeight     =   4920
      ScaleWidth      =   8640
      TabIndex        =   240
      TabStop         =   0   'False
      Top             =   4950
      Visible         =   0   'False
      Width           =   8700
      Begin VB.PictureBox pbcMPromoTab 
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
         Height          =   30
         Left            =   5190
         ScaleHeight     =   30
         ScaleWidth      =   15
         TabIndex        =   250
         Top             =   4905
         Width           =   15
      End
      Begin VB.PictureBox pbcMPromoSTab 
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
         Left            =   5100
         ScaleHeight     =   60
         ScaleWidth      =   45
         TabIndex        =   247
         Top             =   255
         Width           =   45
      End
      Begin VB.PictureBox pbcMPSATab 
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
         Height          =   45
         Left            =   945
         ScaleHeight     =   45
         ScaleWidth      =   30
         TabIndex        =   245
         Top             =   4875
         Width           =   30
      End
      Begin VB.PictureBox pbcMPSASTab 
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
         Left            =   930
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   242
         Top             =   300
         Width           =   75
      End
      Begin VB.PictureBox pbcPromo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4950
         Left            =   5325
         Picture         =   "Vehopt.frx":4AA96
         ScaleHeight     =   4920
         ScaleWidth      =   2970
         TabIndex        =   248
         Top             =   0
         Width           =   3000
         Begin VB.TextBox edcMPromo 
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
            Left            =   525
            MaxLength       =   2
            TabIndex        =   249
            TabStop         =   0   'False
            Top             =   210
            Visible         =   0   'False
            Width           =   795
         End
      End
      Begin VB.PictureBox pbcPSA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4950
         Left            =   1155
         Picture         =   "Vehopt.frx":52B38
         ScaleHeight     =   4920
         ScaleWidth      =   2970
         TabIndex        =   243
         TabStop         =   0   'False
         Top             =   0
         Width           =   3000
         Begin VB.TextBox edcMPSA 
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
            Left            =   525
            MaxLength       =   2
            TabIndex        =   244
            TabStop         =   0   'False
            Top             =   600
            Visible         =   0   'False
            Width           =   795
         End
      End
      Begin VB.Label lacPromo 
         Appearance      =   0  'Flat
         Caption         =   "Promo"
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
         Left            =   4560
         TabIndex        =   246
         Top             =   2535
         Width           =   660
      End
      Begin VB.Label lacPSA 
         Appearance      =   0  'Flat
         Caption         =   "PSA"
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
         Left            =   645
         TabIndex        =   241
         Top             =   2505
         Width           =   420
      End
   End
   Begin VB.PictureBox plcVirtual 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   10545
      ScaleHeight     =   4200
      ScaleWidth      =   8910
      TabIndex        =   413
      TabStop         =   0   'False
      Top             =   4650
      Visible         =   0   'False
      Width           =   8970
      Begin VB.PictureBox pbcVirtTab 
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
         Left            =   4080
         ScaleHeight     =   60
         ScaleWidth      =   90
         TabIndex        =   422
         Top             =   3735
         Width           =   90
      End
      Begin VB.PictureBox pbcVirtSTab 
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
         Left            =   4140
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   418
         Top             =   555
         Width           =   75
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
         HelpContextID   =   8
         Left            =   5730
         MaxLength       =   8
         TabIndex        =   421
         Top             =   1380
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.PictureBox pbcVehicle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   4335
         Picture         =   "Vehopt.frx":5ABDA
         ScaleHeight     =   3165
         ScaleWidth      =   4170
         TabIndex        =   420
         TabStop         =   0   'False
         Top             =   585
         Visible         =   0   'False
         Width           =   4170
         Begin VB.Label lacVehFrame 
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
            Left            =   15
            TabIndex        =   412
            Top             =   390
            Visible         =   0   'False
            Width           =   4155
         End
      End
      Begin VB.PictureBox plcVehNames 
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
         Height          =   3300
         Index           =   0
         Left            =   105
         ScaleHeight     =   3240
         ScaleWidth      =   2835
         TabIndex        =   414
         TabStop         =   0   'False
         Top             =   510
         Width           =   2895
         Begin VB.ListBox lbcVehNames 
            Appearance      =   0  'Flat
            Height          =   3180
            Left            =   30
            TabIndex        =   415
            Top             =   30
            Width           =   2775
         End
      End
      Begin VB.PictureBox pbcArrowMove 
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
         Index           =   0
         Left            =   3855
         Picture         =   "Vehopt.frx":69334
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   431
         TabStop         =   0   'False
         Top             =   1335
         Width           =   180
      End
      Begin VB.PictureBox pbcArrowMove 
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
         Index           =   1
         Left            =   3270
         Picture         =   "Vehopt.frx":6940E
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   430
         TabStop         =   0   'False
         Top             =   2385
         Width           =   180
      End
      Begin VB.CommandButton cmcMoveToVehicle 
         Appearance      =   0  'Flat
         Caption         =   "M&ove   "
         Height          =   300
         Left            =   3165
         TabIndex        =   416
         Top             =   1275
         Width           =   945
      End
      Begin VB.CommandButton cmcMoveToVehName 
         Appearance      =   0  'Flat
         Caption         =   "    Mo&ve"
         Height          =   300
         Left            =   3165
         TabIndex        =   417
         Top             =   2325
         Width           =   945
      End
      Begin VB.PictureBox plcVehNames 
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
         Height          =   3300
         Index           =   1
         Left            =   4290
         ScaleHeight     =   3240
         ScaleWidth      =   4530
         TabIndex        =   419
         TabStop         =   0   'False
         Top             =   510
         Width           =   4590
         Begin VB.VScrollBar vbcVehicle 
            Height          =   3180
            LargeChange     =   14
            Left            =   4245
            Max             =   0
            TabIndex        =   423
            Top             =   75
            Width           =   270
         End
      End
      Begin VB.Label lacVehMsg 
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
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   75
         TabIndex        =   316
         Top             =   3930
         Visible         =   0   'False
         Width           =   8880
      End
   End
   Begin VB.PictureBox plcProducer 
      Height          =   4515
      Left            =   10755
      ScaleHeight     =   4455
      ScaleWidth      =   8670
      TabIndex        =   318
      TabStop         =   0   'False
      Top             =   4470
      Width           =   8730
      Begin VB.ListBox lbcExpCommAudio 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   6300
         TabIndex        =   327
         TabStop         =   0   'False
         Top             =   1830
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ListBox lbcExpProgAudio 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   5190
         TabIndex        =   326
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ListBox lbcContentProvider 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   5835
         TabIndex        =   325
         TabStop         =   0   'False
         Top             =   1665
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ListBox lbcProducer 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   3180
         TabIndex        =   322
         TabStop         =   0   'False
         Top             =   1620
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.PictureBox pbcPTab 
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
         Left            =   30
         ScaleHeight     =   60
         ScaleWidth      =   0
         TabIndex        =   328
         Top             =   1440
         Width           =   0
      End
      Begin VB.PictureBox pbcProducer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   135
         Picture         =   "Vehopt.frx":694E8
         ScaleHeight     =   1065
         ScaleWidth      =   4620
         TabIndex        =   320
         TabStop         =   0   'False
         Top             =   120
         Width           =   4620
         Begin VB.PictureBox pbcCommEmbedded 
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
            Left            =   3390
            ScaleHeight     =   210
            ScaleWidth      =   1185
            TabIndex        =   324
            TabStop         =   0   'False
            Top             =   1170
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox edcPDropdown 
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
            Left            =   2985
            MaxLength       =   3
            TabIndex        =   321
            TabStop         =   0   'False
            Top             =   420
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton cmcPDropdown 
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
            Left            =   3570
            Picture         =   "Vehopt.frx":7EC0E
            TabIndex        =   323
            TabStop         =   0   'False
            Top             =   435
            Visible         =   0   'False
            Width           =   195
         End
      End
      Begin VB.PictureBox pbcPSTab 
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
         Left            =   0
         ScaleHeight     =   60
         ScaleWidth      =   15
         TabIndex        =   319
         Top             =   975
         Width           =   15
      End
   End
   Begin VB.PictureBox plcSchedule 
      Height          =   4515
      Index           =   0
      Left            =   10800
      ScaleHeight     =   4455
      ScaleWidth      =   8745
      TabIndex        =   229
      Top             =   4320
      Visible         =   0   'False
      Width           =   8805
      Begin VB.CheckBox ckcSch 
         Caption         =   "First In, Stays In when Scheduling Orders (Suppress Pre-empting)"
         Height          =   270
         Index           =   4
         Left            =   105
         TabIndex        =   230
         Top             =   105
         Width           =   5865
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
         TabIndex        =   238
         TabStop         =   0   'False
         Top             =   1470
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmcLevelPrice 
         Appearance      =   0  'Flat
         Caption         =   "Generate Price Levels"
         Height          =   255
         Left            =   5385
         TabIndex        =   235
         Top             =   735
         Width           =   2175
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
         TabIndex        =   239
         Top             =   2265
         Width           =   120
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
         Left            =   45
         ScaleHeight     =   60
         ScaleWidth      =   120
         TabIndex        =   236
         Top             =   1125
         Width           =   120
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
         TabIndex        =   234
         Top             =   915
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
         Index           =   0
         Left            =   3900
         MaxLength       =   5
         TabIndex        =   232
         Top             =   510
         Width           =   1080
      End
      Begin VB.PictureBox pbcSchedule 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   90
         Picture         =   "Vehopt.frx":7ED08
         ScaleHeight     =   630
         ScaleWidth      =   8520
         TabIndex        =   237
         Top             =   1425
         Width           =   8520
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
         TabIndex        =   433
         Top             =   2100
         Width           =   5790
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "Highest Spot Price to Generate Price Levels"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   105
         TabIndex        =   233
         Top             =   930
         Width           =   3765
      End
      Begin VB.Label lacSchedule 
         Appearance      =   0  'Flat
         Caption         =   "Lowest Spot Price to Generate Price Levels"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   231
         Top             =   525
         Width           =   3705
      End
   End
   Begin VB.PictureBox plcBarter 
      Height          =   5895
      Left            =   11100
      ScaleHeight     =   5835
      ScaleWidth      =   11130
      TabIndex        =   188
      Top             =   3945
      Visible         =   0   'False
      Width           =   11190
      Begin VB.Frame frcBarter 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   3420
         TabIndex        =   194
         Top             =   270
         Width           =   7455
         Begin VB.TextBox edcBarter 
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
            Left            =   5745
            MaxLength       =   10
            TabIndex        =   201
            Top             =   0
            Width           =   1635
         End
         Begin VB.OptionButton rbcBarterMethod 
            Caption         =   "Radio Station Invoice"
            Height          =   210
            Index           =   6
            Left            =   2970
            TabIndex        =   197
            Top             =   15
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.OptionButton rbcBarterMethod 
            Caption         =   "None"
            Height          =   210
            Index           =   5
            Left            =   2070
            TabIndex        =   196
            Top             =   15
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lacBarter 
            Appearance      =   0  'Flat
            Caption         =   "P/W"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   8
            Left            =   5295
            TabIndex        =   198
            Top             =   15
            Width           =   450
         End
         Begin VB.Label lacBarter 
            Appearance      =   0  'Flat
            Caption         =   "Post Log Import Source"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   0
            TabIndex        =   195
            Top             =   15
            Visible         =   0   'False
            Width           =   2160
         End
      End
      Begin VB.Frame frcBarter 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   189
         Top             =   30
         Width           =   8220
         Begin VB.OptionButton rbcBarterMethod 
            Caption         =   "Marketron"
            Height          =   210
            Index           =   7
            Left            =   5040
            TabIndex        =   496
            Top             =   30
            Width           =   1320
         End
         Begin VB.OptionButton rbcBarterMethod 
            Caption         =   "None"
            Height          =   210
            Index           =   8
            Left            =   2640
            TabIndex        =   192
            Top             =   30
            Width           =   840
         End
         Begin VB.OptionButton rbcBarterMethod 
            Caption         =   "Wide Orbit"
            Height          =   210
            Index           =   9
            Left            =   3600
            TabIndex        =   191
            Top             =   30
            Width           =   1350
         End
         Begin VB.Label lacBarter 
            Appearance      =   0  'Flat
            Caption         =   "Send XML Insertion Order to"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   7
            Left            =   0
            TabIndex        =   190
            Top             =   15
            Width           =   2460
         End
      End
      Begin VB.Frame frcBarterEnable 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   2
         Left            =   360
         TabIndex        =   493
         Top             =   645
         Visible         =   0   'False
         Width           =   7635
         Begin VB.PictureBox pbcAcqSTab 
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
            Index           =   1
            Left            =   7425
            ScaleHeight     =   90
            ScaleWidth      =   105
            TabIndex        =   202
            Top             =   135
            Width           =   105
         End
         Begin VB.PictureBox pbcAcqTab 
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
            Index           =   1
            Left            =   7245
            ScaleHeight     =   90
            ScaleWidth      =   105
            TabIndex        =   205
            Top             =   555
            Width           =   105
         End
         Begin VB.PictureBox pbcAcq 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   1
            Left            =   1320
            Picture         =   "Vehopt.frx":904DA
            ScaleHeight     =   450
            ScaleWidth      =   5610
            TabIndex        =   203
            TabStop         =   0   'False
            Top             =   120
            Width           =   5610
            Begin VB.TextBox edcBarter 
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
               Index           =   4
               Left            =   30
               MaxLength       =   4
               TabIndex        =   204
               TabStop         =   0   'False
               Top             =   210
               Visible         =   0   'False
               Width           =   540
            End
         End
         Begin VB.Label lacBarter 
            Appearance      =   0  'Flat
            Caption         =   "Spot Lengths"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   0
            TabIndex        =   199
            Top             =   120
            Width           =   1275
         End
         Begin VB.Label lacBarter 
            Appearance      =   0  'Flat
            Caption         =   "Index"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   6
            Left            =   0
            TabIndex        =   200
            Top             =   345
            Width           =   870
         End
      End
      Begin VB.CheckBox ckcBarter 
         Caption         =   "Include on Insertion Orders"
         Height          =   300
         Index           =   0
         Left            =   135
         TabIndex        =   193
         Top             =   255
         Width           =   3060
      End
      Begin VB.Frame frcBarterEnable 
         BorderStyle     =   0  'None
         Height          =   1035
         Index           =   1
         Left            =   3615
         TabIndex        =   488
         Top             =   1695
         Visible         =   0   'False
         Width           =   7410
         Begin VB.TextBox edcBarter 
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
            Left            =   3045
            TabIndex        =   492
            Top             =   570
            Width           =   1080
         End
         Begin VB.TextBox edcBarter 
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
            Left            =   2160
            TabIndex        =   490
            Top             =   30
            Width           =   1080
         End
         Begin VB.Label lacBarter 
            Appearance      =   0  'Flat
            Caption         =   "Station Acquisition Commission"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   0
            TabIndex        =   491
            Top             =   615
            Width           =   2880
         End
         Begin VB.Label lacBarter 
            Appearance      =   0  'Flat
            Caption         =   "Effective Start Date"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   0
            TabIndex        =   489
            Top             =   75
            Width           =   1845
         End
      End
      Begin VB.ComboBox cbcBarter 
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
         Left            =   5670
         TabIndex        =   206
         Top             =   720
         Width           =   2760
      End
      Begin VB.Frame frcBarterEnable 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4425
         Index           =   0
         Left            =   825
         TabIndex        =   207
         Top             =   1440
         Visible         =   0   'False
         Width           =   8490
         Begin VB.PictureBox pbcAcq 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   630
            Index           =   0
            Left            =   0
            Picture         =   "Vehopt.frx":91AA8
            ScaleHeight     =   630
            ScaleWidth      =   8340
            TabIndex        =   211
            TabStop         =   0   'False
            Top             =   420
            Width           =   8340
            Begin VB.TextBox edcBarter 
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
               Index           =   3
               Left            =   1590
               TabIndex        =   212
               TabStop         =   0   'False
               Top             =   210
               Visible         =   0   'False
               Width           =   660
            End
         End
         Begin VB.PictureBox pbcAcqTab 
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
            Index           =   0
            Left            =   0
            ScaleHeight     =   90
            ScaleWidth      =   105
            TabIndex        =   213
            Top             =   1050
            Width           =   105
         End
         Begin VB.PictureBox pbcAcqSTab 
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
            Index           =   0
            Left            =   15
            ScaleHeight     =   60
            ScaleWidth      =   120
            TabIndex        =   210
            Top             =   315
            Width           =   120
         End
         Begin VB.Frame frcBarter 
            Caption         =   "Barter Method"
            Height          =   2415
            Index           =   0
            Left            =   0
            TabIndex        =   214
            Top             =   1095
            Width           =   8355
            Begin VB.TextBox edcBarterMethod 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Height          =   300
               Index           =   5
               Left            =   6480
               MaxLength       =   10
               TabIndex        =   495
               Top             =   1020
               Width           =   1080
            End
            Begin VB.TextBox edcBarterMethod 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
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
               Height          =   300
               Index           =   4
               Left            =   6480
               MaxLength       =   10
               TabIndex        =   494
               Top             =   615
               Width           =   1080
            End
            Begin VB.ComboBox cbcPerPeriod 
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
               Index           =   2
               ItemData        =   "Vehopt.frx":A2D3A
               Left            =   3495
               List            =   "Vehopt.frx":A2D47
               TabIndex        =   224
               Top             =   1455
               Width           =   1140
            End
            Begin VB.ComboBox cbcPerPeriod 
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
               Index           =   1
               ItemData        =   "Vehopt.frx":A2D5E
               Left            =   2130
               List            =   "Vehopt.frx":A2D6B
               TabIndex        =   220
               Top             =   1005
               Width           =   1140
            End
            Begin VB.ComboBox cbcPerPeriod 
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
               Index           =   0
               ItemData        =   "Vehopt.frx":A2D82
               Left            =   2385
               List            =   "Vehopt.frx":A2D8F
               TabIndex        =   217
               Top             =   600
               Width           =   1140
            End
            Begin VB.OptionButton rbcBarterMethod 
               Caption         =   "All Cash"
               Height          =   210
               Index           =   0
               Left            =   180
               TabIndex        =   215
               Top             =   255
               Width           =   1095
            End
            Begin VB.TextBox edcBarterMethod 
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
               Left            =   4245
               MaxLength       =   5
               TabIndex        =   218
               Top             =   615
               Width           =   1080
            End
            Begin VB.TextBox edcBarterMethod 
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
               Left            =   4245
               MaxLength       =   5
               TabIndex        =   221
               Top             =   1020
               Width           =   1080
            End
            Begin VB.TextBox edcBarterMethod 
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
               Left            =   5340
               MaxLength       =   5
               TabIndex        =   225
               Top             =   1455
               Width           =   720
            End
            Begin VB.TextBox edcBarterMethod 
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
               Left            =   1500
               MaxLength       =   5
               TabIndex        =   223
               Top             =   1455
               Width           =   690
            End
            Begin VB.OptionButton rbcBarterMethod 
               Caption         =   "None"
               Height          =   210
               Index           =   4
               Left            =   180
               TabIndex        =   226
               Top             =   1875
               Width           =   840
            End
            Begin VB.OptionButton rbcBarterMethod 
               Caption         =   "After Every                    Paid Spots per  Week               , Allow                     Free Spots per Week "
               Height          =   210
               Index           =   3
               Left            =   180
               TabIndex        =   222
               Top             =   1470
               Width           =   8085
            End
            Begin VB.OptionButton rbcBarterMethod 
               Caption         =   "Pay When Minutes per  Week                Exceed                                Balance"
               Height          =   225
               Index           =   1
               Left            =   180
               TabIndex        =   216
               Top             =   660
               Width           =   6240
            End
            Begin VB.OptionButton rbcBarterMethod 
               Caption         =   "Pay When Units per  Week                  Exceed                                   Balance"
               Height          =   210
               Index           =   2
               Left            =   180
               TabIndex        =   219
               Top             =   1065
               Width           =   6240
            End
         End
         Begin VB.TextBox edcInsertComment 
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
            Height          =   525
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   228
            Top             =   3810
            Width           =   8310
         End
         Begin VB.TextBox edcBarter 
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
            TabIndex        =   209
            Top             =   30
            Width           =   1080
         End
         Begin VB.Label lacBarter 
            Appearance      =   0  'Flat
            Caption         =   "Insertion Comment"
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
            Left            =   0
            TabIndex        =   227
            Top             =   3570
            Width           =   2115
         End
         Begin VB.Label lacBarter 
            Appearance      =   0  'Flat
            Caption         =   "Effective Start Date"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   208
            Top             =   75
            Width           =   1845
         End
      End
      Begin VB.Line lncBarter 
         Index           =   1
         X1              =   300
         X2              =   10275
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line lncBarter 
         Index           =   0
         X1              =   285
         X2              =   10260
         Y1              =   600
         Y2              =   600
      End
   End
   Begin ComctlLib.TabStrip tbcSelection 
      Height          =   6360
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11218
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   13
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Genl"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Sales"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "S&chd"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sp&orts"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&PSA/Promo"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Log"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "P&roducer"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Export"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Interface"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Barter"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "P&articipant"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Grea&t Plains G/L"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "A&ffiliate Log"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label plcScreen 
      Caption         =   "Vehicle Options"
      Height          =   195
      Left            =   15
      TabIndex        =   0
      Top             =   -45
      Width           =   1815
   End
   Begin VB.Label lacType 
      Alignment       =   1  'Right Justify
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
      Height          =   195
      Left            =   2115
      TabIndex        =   317
      Top             =   0
      Width           =   3075
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   5670
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "VehOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Vehopt.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmVafSrchKey                  tmSafSrchKey                  tmPifSrchKey1             *
'*  tmPifDates                                                                            *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: VehOpt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Vehicle Option input screen code
Option Explicit
Option Compare Text
Dim imVpfChanged As Integer     'Force read at end if changed
Dim tmVehNamesCode() As SORTCODE
Dim smVehNamesCodeTag As String
Dim tmVehicle() As SORTCODE
Dim smVehicleTag As String
Dim tmFeedCode() As SORTCODE
Dim smFeedCodeTag As String
Dim tmProducerCode() As SORTCODE
Dim smProducerCodeTag As String
Dim tmContentProviderCode() As SORTCODE
Dim smContentProviderCodeTag As String
Dim tmProducerOrProviderCode() As SORTCODE

Dim imSSChgMode As Integer
Dim tmSSourceCode() As SORTCODE
Dim smSSourceCodeTag As String
Dim smSUpdateRvf() As String * 1
Dim imVehGpChgMode As Integer
Dim tmVehGpCode() As SORTCODE
Dim smVehGpCodeTag As String
Dim smSMnfStamp As String
Dim tmSMnf() As MNF
Dim tmMediaCode() As SORTCODE
Dim smMediaCodeTag As String
Dim lmPartEnableRow As Integer
Dim lmPartEnableCol As Integer
Dim imCtrlVisible As Integer
Dim lmPartTopRow As Long
Dim imIgnoreScroll As Integer
Dim imFromArrow As Integer
Dim imInitNoRows As Integer
Dim imCbcParticipantListIndex As Integer
Dim imIgnorePartChg As Integer
Dim lmLastBilledDate As Long
Dim bmInDateChg As Boolean

Dim imVpfIndex As Integer
Dim imSelectedIndex As Integer
Dim imPartMissAndReq As Integer 'Is participant missing and is required
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer 'True = Processing arrow key- retain current list box visibly
                                'False= Make list box invisible
Dim imUpdateAllowed As Integer    'User can update records
Dim imIgnoreClickEvent As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imIgnoreChg As Integer  'Ignore changes to fields during Move Record to Controls
Dim imAltered As Integer    'Indicates if any field has been altered
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim tmVpf As VPF            'Vehicle preference record image
Dim tmSrchKey As VPFKEY0    'Vpf key record image
Dim hmVef As Integer
Dim tmVef As VEF            'Vehicle preference record image
Dim tmVefSrchKey As INTKEY0    'Vef key record image
Dim imVefRecLen As Integer
Dim hmVof As Integer
Dim tmVof As VOF            'Vehicle preference record image
Dim tmVofSrchKey As VOFKEY0    'Vef key record image
Dim imVofRecLen As Integer
Dim tmLVof As VOF            'Vehicle preference record image
Dim tmCVof As VOF            'Vehicle preference record image
Dim tmOVof As VOF            'Vehicle preference record image
Dim hmVbf As Integer
Dim tmVbfSrchKey As LONGKEY0    'Vef key record image
Dim tmVbfSrchKey1 As VBFKEY1    'Vef key record image
Dim imVBfRecLen As Integer
Dim tmVbf() As VBF
Dim tmVbfIndex As VBF
Dim smVbfComment() As String
Dim imVBFIndex As Integer
Dim lmVbfCode As Long
Dim imVbfChg As Integer
Dim hmVaf As Integer
Dim tmVafSrchKey1 As VBFKEY1    'Vef key record image
Dim imVafRecLen As Integer
Dim tmVaf As VAF
Dim hmSaf As Integer
Dim tmSaf As SAF            'Schedule Attributes record image
Dim tmSafSrchKey1 As SAFKEY1    'Vef key record image
Dim imSafRecLen As Integer
Dim hmVff As Integer            'Multiname file handle
Dim imVffRecLen As Integer      'MNF record length
Dim tmVffSrchKey1 As INTKEY0
Dim tmVff As VFF
Dim smXDXMLForm As String       'S=H#B# (Hour and Break); A=H#B#P# (Hour; Break and Position); P= ISCI
Dim imVffChg As Integer
Dim hmVtf As Integer            'Multiname file handle
Dim imVtfRecLen As Integer      'MNF record length
Dim tmVtfSrchKey As LONGKEY0    'Vsf key record image
Dim tmVtfSrchKey1 As VTFKEY1
Dim tmVtf As VTF
Dim lmT1VtfCode As Long 'Event Title 1
Dim lmT2VtfCode As Long 'Event Title 2
Dim hmVsf As Integer 'Name and address file handle
Dim tmVsf As VSF       'Vsf record image
Dim tmVsfSrchKey As LONGKEY0    'Vsf key record image
Dim imVsfRecLen As Integer        'Vsf record length
Dim hmCef As Integer    'Comment file handle
Dim tmCef As CEF        'CEF record image
Dim tmCefSrchKey As LONGKEY0    'CEF key record image
Dim imCefRecLen As Integer        'CEF record length
Dim tmRnf As RNF                'RNF record image
Dim imRnfRecLen As Integer      'RnF record length
Dim hmRnf As Integer            'Report Name file handle
Dim tmArf As ARF                'ARF record image
Dim tmArfSrchKey As INTKEY0     'ARF key 0 image
Dim imArfRecLen As Integer      'ARF record length
Dim hmArf As Integer            'ARF file handle

Dim tmGhf As GHF                'GHF record image
Dim tmGhfSrchKey1 As GHFKEY1    'GHF key 0 image
Dim imGhfRecLen As Integer      'GHF record length
Dim hmGhf As Integer            'GHF file handle
Dim tmSeasonInfo() As SEASONINFO

Dim hmPif As Integer
Dim imPifRecLen As Integer
Dim tmPifSrchKey As LONGKEY0
Dim tmPifRec() As PIFREC
Dim imPifChg As Integer
Dim smOrigPartStartDate As String

Dim imVefChg As Integer

'NIF network inventory added 7-21-05
Dim tmNif As NIF                'NIF record image
Dim tmNifSrchKey1 As NIFKEY1     'NIF key 1 image
Dim imNifRecLen As Integer      'NIF record length
Dim hmNif As Integer            'NIF file handle
Dim lmInventory As Long       'amount of inventory defined
Dim imYear As Integer
Dim imInventoryAltered As Integer
Dim smRollover As String * 1
Dim smByWeekOrYear As String * 1        'is the input by week or

Dim imFirstTime As Integer
Dim imGTZBoxNo As Integer
Dim imSSpotLenBoxNo As Integer
Dim imLevelPriceBoxNo As Integer
Dim imAcqCostBoxNo As Integer
Dim imAcqIndexBoxNo As Integer
Dim imMPSABoxNo As Integer  '1 thru 72 (-1=not assigned)
Dim imMPromoBoxNo As Integer '1 thru 72 (-1=not assigned)
Dim imTZMaxCtrls As Integer
Dim tmTZCtrls(0 To 13)  As FIELDAREA    'Time zone fields
Dim imLBTZCtrls As Integer
Dim imTZX As Integer
Dim imTZY As Integer
Dim imTZW As Integer
Dim imTZH As Integer
Dim imSpotLGX As Integer
Dim imSpotLGY As Integer
Dim imSpotLGW As Integer
Dim imSpotLGH As Integer
Dim imAcqCostX As Integer
Dim imAcqCostY As Integer
Dim imAcqCostW As Integer
Dim imAcqCostH As Integer
Dim imAcqIndexX As Integer
Dim imAcqIndexY As Integer
Dim imAcqIndexW As Integer
Dim imAcqIndexH As Integer
Dim imPsaX As Integer
Dim imPsaY As Integer
Dim imPsaW As Integer
Dim imPsaH As Integer
Dim imPromoX As Integer
Dim imPromoY As Integer
Dim imPromoW As Integer
Dim imPromoH As Integer
Dim fmTextHeight As Single  'Standard text height
'Producer
'Dim tmPCtrls(1 To 5)  As FIELDAREA    'Producer Fields
Dim tmPCtrls(0 To 4)  As FIELDAREA    'Producer Fields
Dim imLBPCtrls As Integer
Dim imPBoxNo As Integer
'Log Form
Dim tmLCtrls(0 To 14) As FIELDAREA
Dim imLBLCtrls As Integer
Dim smLSave(0 To 3, 0 To 3) As String
Dim imLSave(0 To 11, 0 To 3) As Integer
Dim smLShow(0 To 14, 0 To 3) As String
Dim imLRowNo As Integer
Dim imLBoxNo As Integer
Dim smFTP As String
Dim smLiveWindow As String
'2/28/19: Add Cart on Web
Dim smCartOnWeb As String
Dim smEMail As String
'Moved to vehopt as out of room for static variables
Dim smAutoExpt As String
Dim smAutoImpt As String
Dim smLogExpt As String
Dim smOpenBB As String
Dim smCloseBB As String

'Schedule Price
Dim tmSCtrls(0 To 14) As FIELDAREA
Dim imLBSCtrls As Integer
Dim lmSSave(0 To 14) As Long
Dim imSBoxNo As Integer
Dim imSRowNo As Integer

'Virtual Vehicle
'Combo Box Field Areas
Dim tmVirtCtrls(0 To 3) As FIELDAREA
Dim imLBVirtCtrls As Integer
Dim smVirtSave() As String          'Index 1:Vehicle Name; 2:# spots; 3:% of $'s
Dim imVirtSave() As Integer         'Index 1: Vehicle Code
Dim smVirtShow() As String
Dim imVirtBoxNo As Integer
Dim imVirtRowNo As Integer
Dim imVirtChgVeh As Integer   'True=Values changed; False=Nothing changed
Dim imOrigNoVehicles As Integer
Dim imVirtDoubleClick As Integer
Dim imVirtSettingValue As Integer
Dim imVirtError As Integer
Dim imPartError As Integer
Dim imBarterError As Integer
Dim imNameError As Integer
'Dim smInitLgVehNm As String
'Dim smInitLgHd1 As String
'Dim smInitLgFt1 As String
'Dim smInitLgFt2 As String
Dim imLevelAltered As Integer
Dim imLogAltered As Integer
Dim imProducerAltered As Integer
Dim imGreatPlainsAltered As Integer
Dim imOrigSAGroupNo As Integer
Dim imAskedSAGroupNo As Integer
Dim smISCIAvailForm As String
Dim smHBHBPAvailForm As String

Dim smAllowMGSpots As String
Dim smAllowReplSpots As String
Dim smNoMissedReason As String

Dim imInInit As Integer 'In Init routine
'Combo Box Field Areas
Const VEHINDEX = 1
Const NOSPOTSINDEX = 2
Const PERCENTINDEX = 3  'Percent control/field
'10933
Private Const CUEZONE = 7
'Drag
Dim imDragIndexSrce As Integer  '
Dim imDragIndexDest As Integer  '
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragButton As Integer 'Value 1= Left button; 2=Right button; 4=Middle button
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer  'Shift state when mouse down event occurrs
Dim imDragSrce As Integer 'Values defined below
Dim imDragDest As Integer 'Values defined below
'9114
Dim smOriginalXdsProgramCode As String
Dim bmOriginalXdsHonorMerge As Boolean
'10050 use for podcast stuff too
Dim cmt_rst As ADODB.Recordset
Dim bmCitationDefined As Boolean
'10050
'Dim smAdVendorVehNameOriginal As String
'10981
Dim smVendorExternalIDOriginal As String
Dim tmVendorInfo() As VendorInfo
Dim imCurrentVendorInfoIndex As Integer
'10894 removed
'Dim bmPodCastAllowedInSite As Boolean
'Dim bmAirTimeAllowedInSite As Boolean
'10463
Dim bmAirPodDoubleClick As Boolean

Const LBONE = 1
Const ADJBD = 1


Const DRAGVEHNAME = 1
Const DRAGVEHICLE = 2
Const GNAMEINDEX = 1
Const GFEDZONEINDEX = 2
Const GLOCALADJINDEX = 3
Const GFEEDADJINDEX = 4
Const GVERDISPLINDEX = 5    '5, 6 and 7
Const GCMMLSCHINDEX = 9
Const GFEDDELIVERYINDEX = 10
Const GFEEDINDEX = 11 'was 8
Const GBUSINDEX = 12 'was 9
Const GSCHDINDEX = 13 'was 10
Const LNODAYSINDEX = 1
Const LSKIPINDEX = 2
Const LLENINDEX = 3
Const LPRODINDEX = 4
Const LTITLEINDEX = 5
Const LISCIINDEX = 6
Const LDAYPARTINDEX = 7
Const LTIMEINDEX = 8
Const LLINEINDEX = 9
Const LHOURINDEX = 10
Const LLOADINDEX = 11
Const LHEADERINDEX = 12
Const LFOOT1INDEX = 13
Const LFOOT2INDEX = 14

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

Const PRODUCERINDEX = 1
'Const CONTENTPROVIDERINDEX = 2
Const EXPPROGAUDIOINDEX = 2 '3
Const EXPCOMMAUDIOINDEX = 3 '4
Const COMMEMBEDDEDINDEX = 4 '5

Const SSOURCEINDEX = 0
Const PARTINDEX = 1
Const INTUPDATEINDEX = 2
Const EXTUPDATEINDEX = 3
Const PRODPCTINDEX = 4
Const PIFCODEINDEX = 5
'10050
Const ADVENDOR = 0 ' 3
Const EDICALLLETTERS = 4
Const EDIBAND = 5
Const VEDICALLLETTERS = 6
Const VEDIBAND = 7
'10980 replace
'Const ADVENDORVEHICLENAME = 8   '4
Const ADVENDOREXTERNALIDINDEX = 8

Const Signon = 9
Const SAGROUPNO = 10
Const ADVENDORLABEL = 2
Const PODCASTRBC = 6


'*******************************************************
'*                                                     *
'*      Procedure Name:mContentProviderBranch          *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Lock   *
'*                      Box and process                *
'*                      communication back from Lock   *
'*                      Box                            *
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
Private Function mContentProviderBranch() As Integer
'
'   ilRet = mContentProviderBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
'    ilRet = gOptionalLookAhead(edcPDropdown, lbcContentProvider, imBSMode, slStr)
    If imPBoxNo = EXPPROGAUDIOINDEX Then
        ilRet = gOptionalLookAhead(edcPDropdown, lbcExpProgAudio, imBSMode, slStr)
    Else
        ilRet = gOptionalLookAhead(edcPDropdown, lbcExpCommAudio, imBSMode, slStr)
    End If
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcPDropdown.Text = "[None]") Then
        imDoubleClickName = False
        mContentProviderBranch = False
        Exit Function
    End If
    If igWinStatus(VEHICLESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mContentProviderBranch = True
        'lbcContentProvider.SetFocus
        If imPBoxNo = EXPPROGAUDIOINDEX Then
            lbcExpProgAudio.SetFocus
        Else
            lbcExpCommAudio.SetFocus
        End If
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(LOCKBOXESLIST)) Then
    '    imDoubleClickName = False
    '    mContentProviderBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass
    sgArfCallType = "C"
    igArfCallSource = CALLSOURCEVEHOPT
    If edcPDropdown.Text = "[New]" Then
        sgArfName = ""
    Else
        sgArfName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Agency^Test\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
        Else
            slStr = "Agency^Prod\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Agency^Test^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    Else
    '        slStr = "Agency^Prod^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "NmAddr.Exe " & slStr, 1)
    'Agency.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    NmAddr.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgArfName)
    igArfCallSource = Val(sgArfName)
    ilParse = gParseItem(slStr, 2, "\", sgArfName)
    'Agency.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mContentProviderBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igArfCallSource = CALLDONE Then  'Done
        igArfCallSource = CALLNONE
'        gSetMenuState True
        If imPBoxNo = EXPPROGAUDIOINDEX Then
            lbcExpProgAudio.Clear
        Else
            lbcExpCommAudio.Clear
        End If
        smContentProviderCodeTag = ""
        mContentProviderPop
        If imTerminate Then
            mContentProviderBranch = False
            Exit Function
        End If
'        mProducerOrProviderPop
        If imPBoxNo = EXPPROGAUDIOINDEX Then
            gFindMatch sgArfName, 1, lbcExpProgAudio
            sgArfName = ""
            If gLastFound(lbcExpProgAudio) > 0 Then
                imChgMode = True
                lbcExpProgAudio.ListIndex = gLastFound(lbcExpProgAudio)
                edcPDropdown.Text = lbcExpProgAudio.List(lbcExpProgAudio.ListIndex)
                imChgMode = False
                mContentProviderBranch = False
                mSetCommands
            Else
                imChgMode = True
                lbcExpProgAudio.Height = gListBoxHeight(lbcExpProgAudio.ListCount, 6)
                lbcExpProgAudio.ListIndex = 1
                edcPDropdown.Text = lbcExpProgAudio.List(lbcExpProgAudio.ListIndex)
                imChgMode = False
                mSetCommands
                edcPDropdown.SetFocus
                Exit Function
            End If
        Else
            gFindMatch sgArfName, 1, lbcExpCommAudio
            sgArfName = ""
            If gLastFound(lbcExpCommAudio) > 0 Then
                imChgMode = True
                lbcExpCommAudio.ListIndex = gLastFound(lbcExpCommAudio)
                edcPDropdown.Text = lbcExpCommAudio.List(lbcExpCommAudio.ListIndex)
                imChgMode = False
                mContentProviderBranch = False
                mSetCommands
            Else
                imChgMode = True
                lbcExpCommAudio.Height = gListBoxHeight(lbcExpCommAudio.ListCount, 6)
                lbcExpCommAudio.ListIndex = 1
                edcPDropdown.Text = lbcExpCommAudio.List(lbcExpCommAudio.ListIndex)
                imChgMode = False
                mSetCommands
                edcPDropdown.SetFocus
                Exit Function
            End If
        End If
    End If
    If igArfCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mPEnableBox imPBoxNo
        Exit Function
    End If
    If igArfCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mPEnableBox imPBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function




'*******************************************************
'*                                                     *
'*      Procedure Name:mProducerBranch                 *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Lock   *
'*                      Box and process                *
'*                      communication back from Lock   *
'*                      Box                            *
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
Private Function mProducerBranch() As Integer
'
'   ilRet = mProducerBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcPDropdown, lbcProducer, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcPDropdown.Text = "[None]") Then
        imDoubleClickName = False
        mProducerBranch = False
        Exit Function
    End If
    If igWinStatus(VEHICLESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mProducerBranch = True
        lbcProducer.SetFocus
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(LOCKBOXESLIST)) Then
    '    imDoubleClickName = False
    '    mProducerBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass
    sgArfCallType = "P"
    igArfCallSource = CALLSOURCEVEHOPT
    If edcPDropdown.Text = "[New]" Then
        sgArfName = ""
    Else
        sgArfName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Agency^Test\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
        Else
            slStr = "Agency^Prod\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Agency^Test^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    Else
    '        slStr = "Agency^Prod^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "NmAddr.Exe " & slStr, 1)
    'Agency.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    NmAddr.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgArfName)
    igArfCallSource = Val(sgArfName)
    ilParse = gParseItem(slStr, 2, "\", sgArfName)
    'Agency.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mProducerBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igArfCallSource = CALLDONE Then  'Done
        igArfCallSource = CALLNONE
'        gSetMenuState True
        lbcProducer.Clear
        smProducerCodeTag = ""
        mProducerPop
        If imTerminate Then
            mProducerBranch = False
            Exit Function
        End If
'        mProducerOrProviderPop
        gFindMatch sgArfName, 1, lbcProducer
        sgArfName = ""
        If gLastFound(lbcProducer) > 0 Then
            imChgMode = True
            lbcProducer.ListIndex = gLastFound(lbcProducer)
            edcPDropdown.Text = lbcProducer.List(lbcProducer.ListIndex)
            imChgMode = False
            mProducerBranch = False
            mSetCommands
        Else
            imChgMode = True
            lbcProducer.Height = gListBoxHeight(lbcProducer.ListCount, 6)
            lbcProducer.ListIndex = 1
            edcPDropdown.Text = lbcProducer.List(lbcProducer.ListIndex)
            imChgMode = False
            mSetCommands
            edcPDropdown.SetFocus
            Exit Function
        End If
    End If
    If igArfCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mPEnableBox imPBoxNo
        Exit Function
    End If
    If igArfCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mPEnableBox imPBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mContentProviderPop             *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Lock Box list box     *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mContentProviderPop()
'
'   mContentProvPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffset(0) As Integer
    Dim ilRet As Integer
    Dim slProgName As String
    Dim ilProgIndex As Integer
    Dim slCommName As String
    Dim ilCommIndex As Integer
    Dim ilLoop As Integer

    ilProgIndex = lbcExpProgAudio.ListIndex
    If ilProgIndex >= 1 Then
        slProgName = lbcExpProgAudio.List(ilProgIndex)
    End If
    ilCommIndex = lbcExpCommAudio.ListIndex
    If ilCommIndex >= 1 Then
        slCommName = lbcExpCommAudio.List(ilCommIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "C"
    ilOffset(0) = gFieldOffset("Arf", "ArfType") '2
    'ilRet = gIMoveListBox(Agency, lbcContentProvider, lbcContentProviderCode, "Arf.Btr", gFieldOffset("Arf", "ArfID"), 10, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(VehOpt, lbcExpProgAudio, tmContentProviderCode(), smContentProviderCodeTag, "Arf.Btr", gFieldOffset("Arf", "ArfName"), 40, ilfilter(), slFilter(), ilOffset())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mContentProviderPopErr
        gCPErrorMsg ilRet, "mContentProviderPop (gIMoveListBox)", VehOpt
        On Error GoTo 0
        lbcExpCommAudio.Clear
        For ilLoop = 0 To lbcExpProgAudio.ListCount - 1 Step 1
            lbcExpCommAudio.AddItem lbcExpProgAudio.List(ilLoop)
        Next ilLoop
        lbcExpProgAudio.AddItem "[None]", 0  'Force as first item on list
        lbcExpProgAudio.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilProgIndex > 1 Then
            gFindMatch slProgName, 2, lbcExpProgAudio
            If gLastFound(lbcExpProgAudio) > 1 Then
                lbcExpProgAudio.ListIndex = gLastFound(lbcExpProgAudio)
            Else
                lbcExpProgAudio.ListIndex = -1
            End If
        Else
            lbcExpProgAudio.ListIndex = ilProgIndex
        End If
        imChgMode = False
        lbcExpCommAudio.AddItem "[None]", 0  'Force as first item on list
        lbcExpCommAudio.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilCommIndex > 1 Then
            gFindMatch slCommName, 2, lbcExpCommAudio
            If gLastFound(lbcExpCommAudio) > 1 Then
                lbcExpCommAudio.ListIndex = gLastFound(lbcExpCommAudio)
            Else
                lbcExpCommAudio.ListIndex = -1
            End If
        Else
            lbcExpCommAudio.ListIndex = ilCommIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mContentProviderPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mProducerPop                    *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Lock Box list box     *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mProducerPop()
'
'   mLkBoxPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffset(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcProducer.ListIndex
    If ilIndex > 1 Then
        slName = lbcProducer.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "P"
    ilOffset(0) = gFieldOffset("Arf", "ArfType") '2
    ilRet = gIMoveListBox(VehOpt, lbcProducer, tmProducerCode(), smProducerCodeTag, "Arf.Btr", gFieldOffset("Arf", "ArfName"), 40, ilfilter(), slFilter(), ilOffset())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mProducerPopErr
        gCPErrorMsg ilRet, "mProducerPop (gIMoveListBox)", VehOpt
        On Error GoTo 0
        lbcProducer.AddItem "[None]", 0  'Force as first item on list
        lbcProducer.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcProducer
            If gLastFound(lbcProducer) > 1 Then
                lbcProducer.ListIndex = gLastFound(lbcProducer)
            Else
                lbcProducer.ListIndex = -1
            End If
        Else
            lbcProducer.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mProducerPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub


Private Sub cbcBarter_Change()
    Dim ilLoop As Integer
    Dim ilSvVbfChg As Integer
    
    ilSvVbfChg = imVbfChg
    mMoveCtrlToVbf
    lmVbfCode = cbcBarter.ItemData(cbcBarter.ListIndex)
    mClearBarterCtrl False
    If lmVbfCode >= 0 Then
        If tmVef.sType = "R" Then
            frcBarterEnable(0).Enabled = True
            For ilLoop = 0 To UBound(tmVbf) - 1 Step 1
                If lmVbfCode = tmVbf(ilLoop).lCode Then
                    imVBFIndex = ilLoop
                    mMoveVbfToCtrl
                    Exit For
                End If
            Next ilLoop
        ElseIf (tmVef.sType = "C") Or (tmVef.sType = "S") Then
            frcBarterEnable(1).Enabled = True
            frcBarterEnable(2).Enabled = True
            For ilLoop = 0 To UBound(tmVbf) - 1 Step 1
                If lmVbfCode = tmVbf(ilLoop).lCode Then
                    imVBFIndex = ilLoop
                    mMoveVbfToCtrl
                    Exit For
                End If
            Next ilLoop
        End If
    End If
    imVbfChg = ilSvVbfChg
    mSetCommands
End Sub

Private Sub cbcBarter_Click()
    cbcBarter_Change
End Sub

Private Sub cbcCsiGeneric_DblClick(Index As Integer)
    '10050
    If Index = ADVENDOR And imChgMode = False Then
        'don't go if choosing 'none'
        mGoToAdVendorForm
    End If
    mSetCommands
End Sub

Private Sub cbcCsiGeneric_LostFocus(Index As Integer)
If Index = ADVENDOR And imChgMode = False Then
     If cbcCsiGeneric(ADVENDOR).ListIndex = 0 Then
        mGoToAdVendorForm
        mSetCommands
     End If
End If
End Sub

Private Sub cbcCsiGeneric_OnChange(Index As Integer)
    '10981
    If Index = ADVENDOR And imChgMode = False And cbcCsiGeneric(ADVENDOR).ListIndex > -1 Then
        mVendorsLoadAndSelect False, cbcCsiGeneric(ADVENDOR).GetItemData(cbcCsiGeneric(ADVENDOR).ListIndex)
        mVendorSetExtID
        mVendorEnableOptions
    End If
    'set imLogAltered?
    imLogAltered = True
    mSetCommands
End Sub

Private Sub cbcLog_Change(Index As Integer)
    imLogAltered = True
    mSetCommands
End Sub
Private Sub cbcLog_Click(Index As Integer)
    cbcLog_Change Index
End Sub


Private Sub cbcLog_GotFocus(Index As Integer)
    mLSetShow imLBoxNo
    imLBoxNo = -1
    imLRowNo = -1
End Sub
Private Sub cbcMedia_Change()
    imVffChg = True
    mSetCommands
End Sub

Private Sub cbcMedia_Click()
    cbcMedia_Change
End Sub

Private Sub cbcMedia_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcParticipant_Change()
    Dim ilSvPifChg As Integer
    Dim slStartDate As String

    If imIgnorePartChg Then
        Exit Sub
    End If
    ilSvPifChg = imPifChg
    If Not mGridFieldsOk() Then
        imIgnorePartChg = True
        cbcParticipant.ListIndex = imCbcParticipantListIndex
        imIgnorePartChg = False
        Exit Sub
    End If
    mMoveCtrlToPif
    'slStartDate = edcParticipantDate.Text
    slStartDate = csiParticipantDate.Text
    If gValidDate(slStartDate) And (imCbcParticipantListIndex >= 0) Then
        If (imCbcParticipantListIndex = 0) Then
            cbcParticipant.AddItem slStartDate & "-" & "New"
        End If
        If (imCbcParticipantListIndex > 0) Then
            If InStr(1, cbcParticipant.List(imCbcParticipantListIndex), "-New", vbTextCompare) > 0 Then
                cbcParticipant.List(imCbcParticipantListIndex) = slStartDate & "-" & "New"
            End If
        End If
    End If
    mClearParticipantCtrl True
    imCbcParticipantListIndex = cbcParticipant.ListIndex
    mMovePifToCtrl
    '5/18/18: removing  this feature as it is causing an endless loop
    'Deselect new so that it can be selected a second time if need to add more then one
    'If imCbcParticipantListIndex = 0 Then
    '    imIgnorePartChg = True
    '    cbcParticipant.ListIndex = -1
    '    imIgnorePartChg = False
    'End If
    imPifChg = ilSvPifChg
    mSetCommands
End Sub

Private Sub cbcParticipant_Click()
    cbcParticipant_Change
End Sub

Private Sub cbcParticipant_GotFocus()
    mPartSetShow
End Sub



Private Sub cbcPerPeriod_Click(Index As Integer)
Dim ilPos As Integer
Dim slStr As String

    If Index = 2 Then
        'rbcBarterMethod(3)
        slStr = RTrim(rbcBarterMethod(3).Caption)
        ilPos = InStrRev(slStr, " ")
        slStr = Mid(slStr, 1, ilPos)
        If cbcPerPeriod(Index).ListIndex = 0 Then
            slStr = slStr & "Week"
        ElseIf cbcPerPeriod(Index).ListIndex = 1 Then
            slStr = slStr & "Month"
        Else
            slStr = slStr & "Year"
        End If
        rbcBarterMethod(3).Caption = Trim$(slStr)
    End If
End Sub

Private Sub cbcSelect_Change()
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slCode As String
    Dim ilCode As Integer
    Dim slStr As String
    Dim ilValue As Integer

    If imChgMode = False Then
        imChgMode = True
        imIgnoreChg = True
        Screen.MousePointer = vbHourglass  'Wait
        lacType.Caption = ""
        If cbcSelect.Text <> "" Then
            gManLookAhead cbcSelect, imBSMode, imSelectedIndex
        End If
        imSelectedIndex = cbcSelect.ListIndex
        If imSelectedIndex >= 0 Then
            slNameCode = tmVehicle(imSelectedIndex).sKey   'Traffic!lbcVehicle.List(imSelectedIndex)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo cbcSelectErr
            gCPErrorMsg ilRet, "cbcSelect_Change (gParseItem field 2)", VehOpt
            On Error GoTo 0
            ilCode = CInt(slCode)
            tmVefSrchKey.iCode = ilCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo cbcSelectErr
            gBtrvErrorMsg ilRet, "cbcSelect (btrGetEqual):" & "Vef.btr", VehOpt
            On Error GoTo cbcSelectErr
            imVpfIndex = gVpfFind(VehOpt, ilCode)
            'Move here from below as used by mVbfReadRec
            mMoveRec tgVpf(imVpfIndex), tmVpf
            On Error GoTo 0
            imGTZBoxNo = -1
            imSSpotLenBoxNo = -1
            imAcqCostBoxNo = -1
            imAcqIndexBoxNo = -1
            imLevelPriceBoxNo = -1
            imMPSABoxNo = -1
            imMPromoBoxNo = -1
            imLBoxNo = -1
            imLRowNo = -1
            If tmVef.sType = "V" Then
                lbcVehNames.Clear   'Clear since not all vehicle names might be shown
                'lbcVehNamesCode.Clear
                ReDim tmVehNamesCode(0 To 0) As SORTCODE
                'lbcVehNamesCode.Tag = ""
                smVehNamesCodeTag = ""
                mVirtVehPop
                ilRet = mVsfReadRec(tmVef.lVsfCode, SETFORREADONLY)
                mVirtMoveRecToCtrl
                pbcVehicle.Cls
                pbcVehicle_Paint
                'rbcOption(5).Visible = True
            Else
                'rbcOption(5).Visible = False
            End If
            If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "R") Or (tmVef.sType = "G") Then
                rbcRemoteExport(0).Enabled = True
                rbcRemoteExport(1).Enabled = True
                rbcRemoteExport(2).Enabled = True
            Else
                rbcRemoteExport(0).Enabled = False
                rbcRemoteExport(1).Enabled = False
                rbcRemoteExport(2).Enabled = True
            End If
            If (tmVef.sType <> "P") And (tmVef.sType <> "N") And (tmVef.sType <> "L") Then
                frcExport(3).Visible = True
            Else
                frcExport(3).Visible = False
            End If
            If (tmVef.sType = "C") Or (tmVef.sType = "S") Then
                rbcRemoteImport(0).Enabled = True
                rbcRemoteImport(2).Enabled = True
            Else
                rbcRemoteImport(0).Enabled = False
                rbcRemoteImport(2).Enabled = True
            End If
            If ((tmVef.sType = "C") Or (tmVef.sType = "G")) And (tgSpf.sGUseAffSys = "Y") Then
                rbcRemoteImport(1).Enabled = True
            Else
                rbcRemoteImport(1).Enabled = False
            End If
            ilRet = mSafReadRec(tmVef.iCode)
            mClearBarterCtrl True
            ilRet = mVbfReadRec(tmVef.iCode)
            ilRet = mVffReadRec(tmVef.iCode)
            If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "N") Or (tmVef.sType = "R") Or (tmVef.sType = "G")) And ((Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS) Then
                ilRet = mVafReadRec(tmVef.iCode)
            End If
            If ((tmVef.sType = "C") And (tmVef.iVefCode <= 0)) Or (tmVef.sType = "A") Or (tmVef.sType = "L") Or (tmVef.sType = "G") Then
                ckcExportSQL.Enabled = True
            Else
                ckcExportSQL.Enabled = False
                ckcExportSQL.Value = vbUnchecked
            End If
            If tmVef.sType = "G" Then
                ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
                If (ilValue And USINGSPORTS) = USINGSPORTS Then 'Using Sports
                    If (ilValue And PREEMPTREGPROG) = PREEMPTREGPROG Then
                        rbcGenLog(3).Enabled = True
                    Else
                        rbcGenLog(3).Enabled = False
                    End If
                Else
                    rbcGenLog(3).Enabled = False
                End If
                
                mVtfReadRec
                
            Else
                rbcGenLog(3).Enabled = False
            End If
            If (Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) <> USINGLIVELOG Then 'Using Live Log
                rbcGenLog(4).Enabled = False
            Else
                rbcGenLog(4).Enabled = rbcGenLog(3).Enabled
            End If
            'Show conventional even if reference Log vehicle as Log Dates are set
            'If ((tmVef.sType = "C") And (tmVef.iVefCode = 0)) Or (tmVef.sType = "L") Or (tmVef.sType = "A") Then
            If ((tmVef.sType = "C") Or (tmVef.sType = "L") Or (tmVef.sType = "A") Or (tmVef.sType = "S") Or (tmVef.sType = "G")) And (igVpfType <> 1) Then
                'Moved to tab control
                'rbcOption(3).Enabled = True
                'If (tgSpf.sGUseAffSys = "Y") And (igVpfType <> 1) Then
                If (igVpfType <> 1) Then
                    ilRet = mVofReadRec(tmVef.iCode, "L")
                    ilRet = mVofReadRec(tmVef.iCode, "C")
                    ilRet = mVofReadRec(tmVef.iCode, "O")
                    ckcUnsoldBlank.Enabled = True
                End If
            Else
                'If rbcOption(3).Value Then
                '    rbcOption(1).Value = True
                'End If
                'rbcOption(3).Enabled = False
                ckcUnsoldBlank.Enabled = False
'                If tbcSelection.SelectedItem.Index = 4 Then
'                    SendKeys "%g", True
'                End If
            End If
            'If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Or (tmVef.sType = "T") Or (tmVef.sType = "R") Or (tmVef.sType = "N")) And (igVpfType <> 1) Then
            imPartMissAndReq = False
            If (igVpfType <> 1) And (tmVef.sType <> "L") Then
                mPopParticipantDates
                If UBound(tmPifRec) <= LBound(tmPifRec) Then
                    imPartMissAndReq = True
                End If
            Else
                ReDim tmPifRec(0 To 0) As PIFREC
            End If
            If Not imInInit Then
                igVehNewToVehOpt = False
            End If
            'If (tbcSelection.SelectedItem.Index = 3) Or (tbcSelection.SelectedItem.Index = 4) Or (tbcSelection.SelectedItem.Index = 5) Or (tbcSelection.SelectedItem.Index = 6) Or (tbcSelection.SelectedItem.Index = 8) Or (tbcSelection.SelectedItem.Index = 9) Then
                tbcSelection_Click
            'End If
'            If tbcSelection.SelectedItem.Index = 4 Then
'                If ((tmVef.sType = "C") Or (tmVef.sType = "L") Or (tmVef.sType = "A") Or (tmVef.sType = "S")) And (igVpfType <> 1) Then
'                    plcLog.Visible = True
'                Else
'                    plcLog.Visible = False
'                End If
'            End If
            imAskedSAGroupNo = False
            'mMoveRec tgVpf(imVpfIndex), tmVpf
            imAskedSAGroupNo = False
            mMoveRecToCtrl tmVpf, tmVff
            mSeasonPop
            imLogAltered = False
            imLevelAltered = False
            imAskedSAGroupNo = False
            imProducerAltered = False
            imInventoryAltered = False  '7-21-05
            imGreatPlainsAltered = False
            If (tmVef.sType = "S") Or (tmVef.sType = "A") Then
                edcGen(SAGROUPNO).Enabled = True
                If imOrigSAGroupNo <> 0 Then
                    imAskedSAGroupNo = True
                End If
            Else
                edcGen(SAGROUPNO).Enabled = False
            End If
            If tmVef.sState = "D" Then
                slStr = ", Dormant"
                lacType.ForeColor = Red
            Else
                slStr = ", Active"
                lacType.ForeColor = DARKGREEN
            End If
            Select Case tmVef.sType
                Case "C"
                    lacType.Caption = "Conventional" & slStr
                Case "S"
                    lacType.Caption = "Selling" & slStr
                Case "A"
                    lacType.Caption = "Airing" & slStr
                Case "L"
                    lacType.Caption = "Log" & slStr
                Case "V"
                    lacType.Caption = "Virtual" & slStr
                Case "T"
                    lacType.Caption = "Simulcast" & slStr
                Case "P"
                    If tmVef.lPvfCode <= 0 Then
                        lacType.Caption = "Package-Dynamic" & slStr
                    Else
                        lacType.Caption = "Package-Standard" & slStr
                    End If
                Case "R"
                    lacType.Caption = "Rep" & slStr
                Case "G"
                    lacType.Caption = "Sport" & slStr
                Case "I"
                    lacType.Caption = "Import" & slStr
                Case "N"
                    lacType.Caption = "NTR" & slStr
            End Select
            'Test if virtual vehicle-
        End If
        Screen.MousePointer = vbDefault
        If imPartMissAndReq Then
            imIgnoreChg = False
        End If
        mSetCommands
        imIgnoreChg = False
        imChgMode = False
    End If
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imIgnoreChg = True
    imTerminate = True
    Exit Sub
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcSelect_DropDown()
    'mPopulate
    If imTerminate Then
        Exit Sub
    End If
End Sub
Private Sub cbcSelect_GotFocus()
    'mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If imFirstTime Then
        'Not necessary as same as vehicle.frm (cbcSelect)
'        If imUpdateAllowed = False Then
'            mSendHelpMsg "BF"
'        Else
'            mSendHelpMsg "BT"
'        End If
        gShowBranner imUpdateAllowed
        If igVehOptCallSource = CALLNONE Then
            If cbcSelect.ListCount > 0 Then
                cbcSelect.ListIndex = 0
            End If
        Else
            gFindMatch sgVehNameToVehOpt, 0, cbcSelect
            If gLastFound(cbcSelect) >= 0 Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                imIgnoreChg = True
                If imPartMissAndReq Then
                    imIgnoreChg = False
                End If
                mSetCommands
                imIgnoreChg = False
                'rbcOption(0).SetFocus
                If tbcSelection.SelectedItem.Index <> 1 Then
                    'SendKeys "%g", True
                Else
                    tbcSelection_Click
                End If
                imFirstTime = False
                Exit Sub
            Else
                cbcSelect.ListIndex = 0
            End If
        End If
        imFirstTime = False
    Else
        mRemoveFocus
    End If
    gCtrlGotFocus cbcSelect
    imIgnoreChg = True
    mSetCommands
    imIgnoreChg = False
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
Private Sub cbcSvLog_Change(Index As Integer)
    imLogAltered = True
    mSetCommands
End Sub
Private Sub cbcSvLog_Click(Index As Integer)
    cbcSvLog_Change Index
End Sub
Private Sub cbcSvLog_GotFocus(Index As Integer)
    mLSetShow imLBoxNo
    imLBoxNo = -1
    imLRowNo = -1
End Sub

Private Sub ckcBarter_Click(Index As Integer)
    imVffChg = True
    mSetCommands
End Sub

Private Sub ckcExportSQL_Click()
    mSetCommands
End Sub

Private Sub ckcExportSQL_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcLog_Click(Index As Integer)
    imVffChg = True
    mSetCommands
End Sub

Private Sub ckcLog_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcIFExport_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcIFExport(Index).Value = vbChecked Then
        Value = True
    End If
    If (Index = 7) Or (Index = 8) Or (Index = 9) Or (Index = 10) Or (Index = 11) Or (Index = 12) Or (Index = 13) Or (Index = 14) Or (Index = 15) Or (Index = 16) Or (Index = 17) Or (Index = 18) Or (Index = 20) Then        'Air wave or Export NY ESPN or Pledge vs Air
        imVffChg = True
    End If
    If Index = 19 Then
        imVefChg = True
    End If
    'End of coded added
    mSetCommands
End Sub

Private Sub ckcKCGenRot_Click()
    mSetCommands
End Sub

Private Sub ckcKCGenRot_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcSch_Click(Index As Integer)
    imVffChg = True
    '9213 temp remove when no longer needed
    If ckcSch(3).Value = vbChecked Then
        udcVehOptTabs.AffLog(0) = vbUnchecked
    End If
    mSetCommands
End Sub

Private Sub ckcSch_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcStnFdInfo_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcStnFdInfo(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    mSetCommands
End Sub
Private Sub ckcStnFdInfo_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcUnsoldBlank_Click()
    mSetCommands
End Sub

Private Sub ckcUnsoldBlank_GotFocus()
    mLSetShow imLBoxNo
    imLBoxNo = -1
    imLRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcAffExport_Click(Index As Integer)
    If (Index = 3) Or (Index = 4) Then
        imVffChg = True
    End If
    mSetCommands
End Sub

Private Sub ckcAffExport_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub


Private Sub ckcXDSave_Click(Index As Integer)
    imVffChg = True
    mSetCommands
End Sub

Private Sub ckcXDSave_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcCancel_Click()
    If (igVpfType <> 1) And (tmVef.sType <> "L") Then
        If imPartMissAndReq Then
            Exit Sub
        End If
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mRemoveFocus
    gCtrlGotFocus cmcCancel
End Sub

Private Sub cmcClear_Click()
    mClearParticipantCtrl False
End Sub

Private Sub cmcClear_GotFocus()
    mRemoveFocus
End Sub

Private Sub cmcDone_Click()
    Dim slMess As String
    Dim ilRes As Integer
    Dim ilRet As Integer
    
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    mSetChg tmVpf, imVpfIndex, imIgnoreChg, imAltered

    If imAltered Then
        slMess = "Save Changes to " & cbcSelect.Text
        ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
        If ilRes = vbCancel Then
            'rbcOption(0).Value = True
            If tbcSelection.SelectedItem.Index <> 1 Then
                ''plcTabSelection.SelectedItem.Index = 1
                'SendKeys "%g", True
            Else
                tbcSelection_Click
            End If
            'edcGSignOn.SetFocus
            Exit Sub
        End If
        If ilRes = vbYes Then
            cmcUpdate_Click
            If imTerminate Then
                Exit Sub
            End If
            If imVirtError Then
                Exit Sub
            End If
            If imPartError Then
                Exit Sub
            End If
            If imBarterError Then
                Exit Sub
            End If
            If imNameError Then
                Exit Sub
            End If
        Else
            If imPartMissAndReq Then
                Exit Sub
            End If
        End If
    Else
        If imPartMissAndReq Then
            Exit Sub
        End If
    End If
    sgVpfStamp = ""
    ilRet = gVpfRead()
    sgVffStamp = ""
    ilRet = gVffRead()
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mRemoveFocus
    gCtrlGotFocus cmcDone
End Sub

Private Sub cmcDropDown_Click()
    Select Case lmPartEnableCol
        Case SSOURCEINDEX
            lbcSSource.Visible = Not lbcSSource.Visible
            edcSSDropDown.SelStart = 0
            edcSSDropDown.SelLength = Len(edcSSDropDown.Text)
            edcSSDropDown.SetFocus
        Case PARTINDEX
            lbcVehGp.Visible = Not lbcVehGp.Visible
            edcVehGpDropDown.SelStart = 0
            edcVehGpDropDown.SelLength = Len(edcVehGpDropDown.Text)
            edcVehGpDropDown.SetFocus
    End Select
End Sub

Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcGTZDropDown_Click()
    If (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDINDEX Then  'Feed
        lbcFeed.Visible = Not lbcFeed.Visible
    End If
End Sub
Private Sub cmcGTZDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcLevelPrice_Click()
    Dim slLow As String
    Dim slHigh As String
    Dim slInc As String
    Dim ilLoop As Integer
    Dim llInc As Long

    slLow = Trim$(edcSchedule(0).Text)
    If slLow = "" Then
        If lmSSave(1) <> 0 Then
            imLevelAltered = True
        End If
        edcSchedule(1).Text = ""
        For ilLoop = LBONE To UBound(lmSSave) Step 1
            lmSSave(ilLoop) = 0
        Next ilLoop
        pbcSchedule.Cls
        pbcSchedule_Paint
        Exit Sub
    End If
    If Val(slLow) <= 0 Then
        For ilLoop = LBONE To UBound(lmSSave) Step 1
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
    If lmSSave(1) <> Val(slLow) Then
        imLevelAltered = True
    End If
    lmSSave(1) = Val(slLow)
    For ilLoop = 2 To 13 Step 1
        If lmSSave(ilLoop) <> lmSSave(ilLoop - 1) + llInc Then
            imLevelAltered = True
        End If
        lmSSave(ilLoop) = lmSSave(ilLoop - 1) + llInc
    Next ilLoop
    If lmSSave(13) <> Val(slHigh) Then
        imLevelAltered = True
    End If
    lmSSave(13) = Val(slHigh)
    pbcSchedule.Cls
    pbcSchedule_Paint
    mSetCommands
End Sub

Private Sub cmcLevelPrice_GotFocus()
    mSSetShow imSBoxNo
    imSBoxNo = -1
    imSRowNo = -1
End Sub

Private Sub cmcMoveToVehicle_Click()
    Dim ilUpper As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilIndex As Integer
    If imOrigNoVehicles > 0 Then
        Exit Sub
    End If
    If lbcVehNames.ListIndex >= 0 Then
        ilUpper = UBound(smVirtSave, 2)
        If ilUpper >= UBound(tmVsf.iFSCode) + 1 Then    'Leave one spot unused for schd and reports
            MsgBox "Only" & Str$(UBound(tmVsf.iFSCode)) & "selections allowed", vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        slNameCode = tmVehNamesCode(lbcVehNames.ListIndex).sKey    'lbcVehNamesCode.List(lbcVehNames.ListIndex)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        smVirtSave(1, ilUpper) = slName
        imVirtSave(1, ilUpper) = Val(slCode)
        gSetShow pbcVehicle, slName, tmVirtCtrls(VEHINDEX)
        smVirtShow(1, ilUpper) = tmVirtCtrls(VEHINDEX).sShow
        smVirtSave(2, ilUpper) = "1"
        gSetShow pbcVehicle, smVirtSave(2, ilUpper), tmVirtCtrls(NOSPOTSINDEX)
        smVirtShow(2, ilUpper) = tmVirtCtrls(NOSPOTSINDEX).sShow
        smVirtSave(3, ilUpper) = ""
        slStr = ""
        gSetShow pbcVehicle, slStr, tmVirtCtrls(PERCENTINDEX)
        smVirtShow(3, ilUpper) = tmVirtCtrls(PERCENTINDEX).sShow
        'lbcVehNamesCode.RemoveItem lbcVehNames.ListIndex
        ilIndex = lbcVehNames.ListIndex
        gRemoveItemFromSortCode ilIndex, tmVehNamesCode()
        lbcVehNames.RemoveItem lbcVehNames.ListIndex
        ilUpper = ilUpper + 1
        ReDim Preserve smVirtSave(0 To 3, 0 To ilUpper) As String
        ReDim Preserve imVirtSave(0 To 1, 0 To ilUpper) As Integer
        ReDim Preserve smVirtShow(0 To 3, 0 To ilUpper) As String
        imVirtSettingValue = True
        'vbcVehicle.Min = LBound(smVirtShow, 2)
        If UBound(smVirtShow, 2) - 1 <= vbcVehicle.LargeChange Then
            vbcVehicle.Max = LBONE  'LBound(smVirtShow, 2)
        Else
            vbcVehicle.Max = UBound(smVirtShow, 2) - vbcVehicle.LargeChange '- 1
        End If
        'vbcVehicle.Value = vbcVehicle.Min
         If imVirtSettingValue Then
            pbcVehicle.Cls
            pbcVehicle_Paint
            imVirtSettingValue = False
         End If
         imVirtChgVeh = True
         mSetCommands
    End If
End Sub
Private Sub cmcMoveToVehicle_GotFocus()
    mVirtSetShow imVirtBoxNo
    imVirtBoxNo = -1
    imVirtRowNo = -1
End Sub
Private Sub cmcMoveToVehName_Click()
    Dim slName As String * 20
    Dim slNameCode As String
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim slGpSort As String
    Dim slVehSort As String
    If imOrigNoVehicles > 0 Then
        Exit Sub
    End If
    If imVirtRowNo >= 1 Then
        'Remove row
        slName = smVirtSave(1, imVirtRowNo)
        slGpSort = "000"
        slVehSort = "000"
        'slNameCode = slName & "\" & Trim$(Str$(imVirtSave(1, imVirtRowNo)))
        slNameCode = slGpSort & "|" & slVehSort & "|" & slName & "\" & Trim$(Str$(imVirtSave(1, imVirtRowNo)))
        'lbcVehNamesCode.AddItem slNameCode
        gAddItemToSortCode slNameCode, tmVehNamesCode(), True
        'For ilLoop = 0 To lbcVehNamesCode.ListCount - 1 Step 1
        '    If slNameCode = lbcVehNamesCode.List(ilLoop) Then
        '        lbcVehNames.AddItem slName, ilLoop
        '        Exit For
        '    End If
        'Next ilLoop
        For ilLoop = 0 To UBound(tmVehNamesCode) - 1 Step 1
            If Trim$(slNameCode) = Trim$(tmVehNamesCode(ilLoop).sKey) Then
                lbcVehNames.AddItem slName, ilLoop
                Exit For
            End If
        Next ilLoop
        For ilLoop = imVirtRowNo To UBound(smVirtSave, 2) - 2 Step 1
            smVirtSave(1, ilLoop) = smVirtSave(1, ilLoop + 1)
            smVirtSave(2, ilLoop) = smVirtSave(2, ilLoop + 1)
            smVirtSave(3, ilLoop) = smVirtSave(3, ilLoop + 1)
            imVirtSave(1, ilLoop) = imVirtSave(1, ilLoop + 1)
            smVirtShow(1, ilLoop) = smVirtShow(1, ilLoop + 1)
            smVirtShow(2, ilLoop) = smVirtShow(2, ilLoop + 1)
            smVirtShow(3, ilLoop) = smVirtShow(3, ilLoop + 1)
        Next ilLoop
        ilUpper = UBound(smVirtSave, 2) - 1
        ReDim Preserve smVirtSave(0 To 3, 0 To ilUpper) As String
        ReDim Preserve imVirtSave(0 To 1, 0 To ilUpper) As Integer
        ReDim Preserve smVirtShow(0 To 3, 0 To ilUpper) As String
        smVirtSave(1, ilUpper) = ""
        smVirtSave(2, ilUpper) = ""
        smVirtSave(3, ilUpper) = ""
        imVirtSave(1, ilUpper) = 0
        smVirtShow(1, ilUpper) = ""
        smVirtShow(2, ilUpper) = ""
        smVirtShow(3, ilUpper) = ""
        imVirtSettingValue = True
        'vbcVehicle.Min = LBound(smVirtShow, 2)
        If UBound(smVirtShow, 2) - 1 <= vbcVehicle.LargeChange Then
            vbcVehicle.Max = LBONE  'LBound(smVirtShow, 2)
        Else
            vbcVehicle.Max = UBound(smVirtShow, 2) - vbcVehicle.LargeChange '- 1
        End If
        'vbcVehicle.Value = vbcVehicle.Min
         If imVirtSettingValue Then
            pbcVehicle.Cls
            pbcVehicle_Paint
            imVirtSettingValue = False
         End If
         imVirtChgVeh = True
         imVirtRowNo = -1
         mSetCommands
    End If
End Sub
Private Sub cmcMoveToVehName_GotFocus()
    mVirtSetShow imVirtBoxNo
    imVirtBoxNo = -1
    'imVirtRowNo = -1
End Sub

Private Sub cmcPDropdown_Click()
    Select Case imPBoxNo
        Case PRODUCERINDEX
            lbcProducer.Visible = Not lbcProducer.Visible
'        Case CONTENTPROVIDERINDEX
'            lbcContentProvider.Visible = Not lbcContentProvider.Visible
        Case EXPPROGAUDIOINDEX
            lbcExpProgAudio.Visible = Not lbcExpProgAudio.Visible
        Case EXPCOMMAUDIOINDEX
            lbcExpCommAudio.Visible = Not lbcExpCommAudio.Visible
    End Select
End Sub

Private Sub cmcPDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'10071 removed
'Private Sub cmcReport_Click()
'    Dim slStr As String
'    'Unload IconTraf
'    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
'    '    Exit Sub
'    'End If
'    igRptCallType = VEHICLESLIST
'    igRptType = 1
'    ''MousePointer = vbHourGlass  'Wait
'    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'    'If (Not igStdAloneMode) And (igShowHelpMsg) Then
'        If igTestSystem Then
'            slStr = "Vehicle^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
'        Else
'            slStr = "Vehicle^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
'        End If
'    'Else
'    '    If igTestSystem Then
'    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
'    '    Else
'    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
'    '    End If
'    'End If
'    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
'    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
'    'VehOpt.Enabled = False
'    'Do While Not igChildDone
'     '   DoEvents
'    'Loop
'    'slStr = sgDoneMsg
'    'VehOpt.Enabled = True
'    'Vehicle!edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
'    'For ilLoop = 0 To 10
'    '    DoEvents
'    'Next ilLoop
'    ''MousePointer = vbDefault    'Default
'    sgCommandStr = slStr
'    RptList.Show vbModal
'End Sub
Private Sub cmcReport_GotFocus()
    mRemoveFocus
    gCtrlGotFocus cmcReport
End Sub
Private Sub cmcUndo_Click()
    Dim ilRet As Integer

    ilRet = mSafReadRec(tmVef.iCode)
    ilRet = mVbfReadRec(tmVef.iCode)
    ilRet = mVffReadRec(tmVef.iCode)
    If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "N") Or (tmVef.sType = "R") Or (tmVef.sType = "G")) And ((Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS) Then
        ilRet = mVafReadRec(tmVef.iCode)
    End If
    If ((tmVef.sType = "C") Or (tmVef.sType = "L") Or (tmVef.sType = "A") Or (tmVef.sType = "S") Or (tmVef.sType = "G")) And (igVpfType <> 1) Then
        If (igVpfType <> 1) Then
            ilRet = mVofReadRec(tmVef.iCode, "L")
            ilRet = mVofReadRec(tmVef.iCode, "C")
            ilRet = mVofReadRec(tmVef.iCode, "O")
        End If
    End If
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Or (tmVef.sType = "T") Or (tmVef.sType = "R") Or (tmVef.sType = "N")) And (igVpfType <> 1) Then
    If (igVpfType <> 1) And (tmVef.sType <> "L") Then
        mPopParticipantDates
    Else
        ReDim tmPifRec(0 To 0) As PIFREC
    End If
    mMoveRec tgVpf(imVpfIndex), tmVpf
    mMoveRecToCtrl tmVpf, tmVff
    imLogAltered = False
    imLevelAltered = False
    imProducerAltered = False
    imInventoryAltered = False          '7-21-05
    imGreatPlainsAltered = False
    If tmVef.sType = "V" Then
        lbcVehNames.Clear
        'lbcVehNamesCode.Clear
        'lbcVehNamesCode.Tag = ""
        ReDim tmVehNamesCode(0 To 0) As SORTCODE
        smVehNamesCodeTag = ""
        mVirtVehPop
        ilRet = mVsfReadRec(tmVef.lVsfCode, SETFORREADONLY)
        mVirtMoveRecToCtrl
        pbcVehicle.Cls
        pbcVehicle_Paint
        imVirtBoxNo = -1
        imVirtRowNo = -1
        imVirtChgVeh = False
    End If
    'If rbcOption(0).Value Then
    '    rbcOption_Click 0
    'Else
    '    rbcOption(0).Value = True
    'End If
    If tbcSelection.SelectedItem.Index <> 1 Then
        'SendKeys "%g", True
    Else
        tbcSelection_Click
    End If
    '12/24/15: Clear enables/disables
    imIgnoreChg = True
    mSetCommands
    imIgnoreChg = False
    DoEvents
    'edcGSignOn.SetFocus
End Sub
Private Sub cmcUndo_GotFocus()
    mRemoveFocus
    gCtrlGotFocus cmcUndo
End Sub
Private Sub cmcUpdate_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slSyncDate                    slSyncTime                    ilWeeks                   *
'*                                                                                        *
'******************************************************************************************

    Dim ilRet As Integer
    Dim slMsg As String
    Dim ilRecLen As Integer     'SPF record length
    Dim hlVpf As Integer        'site Option file handle
    Dim ilRow As Integer
    Dim slType As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilMatchCef As Integer
    Dim tlVof As VOF
    Dim tlSaf As SAF
    Dim tlVaf As VAF
    Dim tlVff As VFF
    Dim slTDate As String
    Dim ilRetFlag As Integer
    Dim ilValue As Integer
    Dim ilVbf As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim tlVbf As VBF
    Dim ilOwnerMnfCode As Integer
    Dim ilSSMnfCode As Integer
    Dim llOwnerDate As Long
    Dim llDate As Long
    Dim ilVef As Integer
    Dim ilPif As Integer
    Dim tlPif As PIF
    Dim ilAtLeast1Asterisk As Integer   'at least 1 time zone has an *
    Dim ilATLeast1NonAsterisk As Integer    'at least 1 time zone has a non *
    Dim ilUsingTZ As Integer                'using times for vehicle
    Dim ilFindAsterisk As Integer
    Dim ilFoundAsterisk As Integer
    Dim ilVff As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilInterfaceID As Integer
    Dim llRet As Long
    
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    imVirtError = False
    imPartError = False
    imBarterError = False
    imNameError = False
    '10894 ok to change from podcast?
    If rbcGMedium(PODCASTRBC).Value = False And tmVpf.sGMedium = "P" Then
        'block saving if podcast vehicle used in a contract.
        If mIsUsedInDigitalLine(tmVef.iCode) Then
            ilRet = MsgBox("Medium was previously used as Podcast.  Cannot change Medium", vbOKOnly + vbExclamation, "Invalid Medium")
            rbcGMedium(PODCASTRBC).Value = True
            '10981
            'mGetPodcastInfo
            mVendorSetAndEnableInfoOriginal
            Exit Sub
        End If
    End If
    If tmVef.sType = "V" Then
        mVirtSetShow imVirtBoxNo
        If mVirtTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
            'rbcOption(5).Value = True
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mVirtEnableBox imVirtBoxNo
            imVirtError = True
            Exit Sub
        End If
    ElseIf (tmVef.sType = "C") Or (tmVef.sType = "A") Then
        slStr = Trim$(edcStnFdCode.Text)
        If (slStr <> "") Then
            If (Asc(slStr) <> 0) Then
                For ilLoop = LBound(tgVpf) To UBound(tgVpf) Step 1
                    If (slStr = Trim$(tgVpf(ilLoop).sStnFdCode)) And (tgVpf(ilLoop).iVefKCode <> tmVef.iCode) Then
                        ilRet = MsgBox("'Station Feed Export: Vehicle ID' is Not Unique", vbOKOnly + vbExclamation, "Station Code")
                        imVirtError = True
                        Exit Sub
                    End If
                Next ilLoop
            End If
        End If
    End If
   
    'If smXDXMLForm = "ISCI" Then
    If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) Then
        If Val(edcInterfaceID(1).Text) > 32000 Then
            ilRet = MsgBox("X-Digital Vehicle ID can not be greater the 32000", vbOKOnly + vbExclamation, "Vehicle ID")
            imVirtError = True
            Exit Sub
        End If
    End If
    If Not mOKName() Then
        imNameError = True
        Exit Sub
    End If
    mMoveCtrlToVbf
    mAdjVbfDates
    ilValue = Asc(tgSpf.sUsingFeatures2)
    ''If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G")) And (ilValue And BARTER) = BARTER Then
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S")) And (ilValue And BARTER) = BARTER Then
    If (tmVef.sType = "R") And ((ilValue And BARTER) = BARTER) Then
        For ilVbf = 0 To UBound(tmVbf) - 1 Step 1
            If tmVbf(ilVbf).lCode = 0 Then
                gUnpackDateLong tmVbf(ilVbf).iStartDate(0), tmVbf(ilVbf).iStartDate(1), llStartDate
                If llStartDate > 0 Then
                    If gWeekDayLong(llStartDate) <> 0 Then
                        ilRet = MsgBox("Start Date of [New] Barter must be a Monday", vbOKOnly + vbExclamation, "Station Code")
                        imBarterError = True
                        Exit Sub
                    End If
                End If
                Exit For
            End If
        Next ilVbf
    ElseIf ((tmVef.sType = "C") Or (tmVef.sType = "S")) And ((ilValue And BARTER) = BARTER) Then
        For ilVbf = 0 To UBound(tmVbf) - 1 Step 1
            If tmVbf(ilVbf).lCode = 0 Then
                gUnpackDateLong tmVbf(ilVbf).iStartDate(0), tmVbf(ilVbf).iStartDate(1), llStartDate
                If llStartDate > 0 Then
                    If gWeekDayLong(llStartDate) <> 0 Then
                        ilRet = MsgBox("Start Date of [New] Barter must be the Start Date of the Broadcast Month", vbOKOnly + vbExclamation, "Station Code")
                        imBarterError = True
                        Exit Sub
                    Else
                        If llStartDate <> gDateValue(gObtainStartStd(Format(llStartDate, "m/d/yy"))) Then
                            ilRet = MsgBox("Start Date of [New] Barter must be the Start Date of the Broadcast Month", vbOKOnly + vbExclamation, "Station Code")
                            imBarterError = True
                            Exit Sub
                        End If
                    End If
                End If
                Exit For
            End If
        Next ilVbf
    End If
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Or (tmVef.sType = "T") Or (tmVef.sType = "R") Or (tmVef.sType = "N")) And (igVpfType <> 1) Then
    If (igVpfType <> 1) And (tmVef.sType <> "L") Then
        If Not mGridFieldsOk() Then
            ilRet = MsgBox("Participant not Complete", vbOKOnly + vbExclamation, "Error")
            imPartError = True
            Exit Sub
        End If
        mMoveCtrlToPif
        ilRet = mCheckPartStartDate()
        If ilRet = 1 Then
            ilRet = MsgBox("New Participant Start date can't be prior to last Billed date", vbOKOnly + vbExclamation, "Error")
            imPartError = True
            Exit Sub
        ElseIf ilRet = 2 Then
            ilRet = MsgBox("New Participant Start date must be on the start of a Standard Broadcast month", vbOKOnly + vbExclamation, "Error")
            imPartError = True
            Exit Sub
        End If
        If UBound(tmPifRec) <= LBound(tmPifRec) Then
            ilRet = MsgBox("Participant(s) must be defined", vbOKOnly + vbExclamation, "Error")
            imPartError = True
            Exit Sub
        End If
        mAdjPifDates
    End If
    Screen.MousePointer = vbHourglass  'Wait
    hlVpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ilRecLen = Len(tgVpf(imVpfIndex)) 'btrRecordLength(hlVpf)  'Get and save record length
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, "cmcUpdate (btrOpen):" & "Vpf.Btr", VehOpt
    On Error GoTo 0
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    tmSrchKey.iVefKCode = tgVpf(imVpfIndex).iVefKCode
    ilRet = btrGetEqual(hlVpf, tgVpf(imVpfIndex), ilRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record so it can be updated
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, "cmcUpdate (btrGetEqual):" & "Vpf.Btr", VehOpt
    On Error GoTo 0
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    ilInterfaceID = tgVpf(imVpfIndex).iInterfaceID
    mMoveCtrlToRec tmVpf
    
     '10-4-11 test for correct input of time zones and feed flags
    'If (tgSpf.sGUseAffSys = "Y") Then       'only test the validity of time zone feed flag if using affiliate.  The fed flag is dual purpose field
        ilAtLeast1Asterisk = False
        ilATLeast1NonAsterisk = False
        ilUsingTZ = False
        'For ilLoop = 1 To 5
        For ilLoop = 0 To 4
            If Trim(tmVpf.sGZone(ilLoop)) <> "" Then
                ilUsingTZ = True
                If tmVpf.sGFed(ilLoop) = "*" Then
                    ilAtLeast1Asterisk = True
                End If
                If Trim(tmVpf.sGFed(ilLoop)) <> "*" Then
                    ilATLeast1NonAsterisk = True
                End If
            End If
        Next ilLoop
        'if at least 1 asterisk exists, any zone not having an asterisk must reference
        'an asterisk zone
        If ilUsingTZ And ilAtLeast1Asterisk And ilATLeast1NonAsterisk Then       'using time zone, continue for validity check if theres at least 1 time zone line with an
            'For ilLoop = 1 To 5
            For ilLoop = 0 To 4
                ilFoundAsterisk = False
                                                                                    'asterisk, and one timezone line without an asterisk
                If Trim$(tmVpf.sGZone(ilLoop)) <> "" And Trim$(tmVpf.sGFed(ilLoop)) <> "*" Then        'time zone description exists and doesnt have an *, needs to reference one of the * time zones
                    'this zone has to reference a zone with an *, look for one
                    'For ilFindAsterisk = 1 To 5
                    For ilFindAsterisk = 0 To 4
                        If tmVpf.sGFed(ilFindAsterisk) = "*" And (tmVpf.sGFed(ilLoop)) = Mid(tmVpf.sGZone(ilFindAsterisk), 1, 1) Then   'this zone
                            ilFoundAsterisk = True
                            Exit For
                        End If
                    Next ilFindAsterisk
                    
                    If Not ilFoundAsterisk Then
                        'error message
                        ilFoundAsterisk = ilFoundAsterisk
                        gMsgBox Trim$(tmVpf.sGZone(ilLoop)) & " Zone Fed column incorrectly defined.  Must reference zone with an *", vbOKOnly + vbExclamation, "Error"
                        Screen.MousePointer = vbDefault
                        mSetCommands
                        Exit Sub
                    End If
                End If
            Next ilLoop
        End If
    'End If
    '7-21-05 update Network Inventory if necessary
    If imYear > 0 Then
        'date last entered or changed
        slTDate = Format$(gNow(), "m/d/yy")


        'see if anything exists for the year
        tmNifSrchKey1.iVefCode = tmVef.iCode
        tmNifSrchKey1.iYear = imYear
        ilRet = btrGetEqual(hmNif, tmNif, imNifRecLen, tmNifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            gPackDate slTDate, tmNif.iEnterDate(0), tmNif.iEnterDate(1)
            tmNif.iUrfCode = tgUrf(0).iCode     'user added/changed
            tmNif.sAllowRollover = smRollover
            tmNif.sInvWkYear = smByWeekOrYear
            mUpdateNifCounts
            ilRet = btrUpdate(hmNif, tmNif, imNifRecLen)
            On Error GoTo cmcUpdateErr
            gBtrvErrorMsg ilRet, "cmcUpdate (btrGetEqual):" & "Nif.Btr", VehOpt
            On Error GoTo 0

        Else
            gPackDate slTDate, tmNif.iEnterDate(0), tmNif.iEnterDate(1)
            tmNif.iUrfCode = tgUrf(0).iCode     'user added/changed
            tmNif.sAllowRollover = smRollover
            tmNif.sInvWkYear = smByWeekOrYear
            tmNif.iYear = imYear
            tmNif.iVefCode = tmVef.iCode
            tmNif.lCode = 0
            mUpdateNifCounts
            ilRet = btrInsert(hmNif, tmNif, imNifRecLen, INDEXKEY0)
            On Error GoTo cmcUpdateErr
            gBtrvErrorMsg ilRet, "cmcUpInsert (btrGetEqual):" & "Nif.Btr", VehOpt
            On Error GoTo 0
        End If
    End If


    tmArfSrchKey.iCode = tmVpf.iFTPArfCode
    If tmArfSrchKey.iCode <> 0 Then
        ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            If Trim$(smFTP) <> "" Then
                tmArf.sFTP = smFTP
                ilRet = btrUpdate(hmArf, tmArf, imArfRecLen)
            Else
                ilRet = btrDelete(hmArf)
                tmVpf.iFTPArfCode = 0
            End If
        Else
            If Trim$(smFTP) <> "" Then
                mInitArf
                tmArf.sFTP = smFTP
                ilRet = btrInsert(hmArf, tmArf, imArfRecLen, INDEXKEY0)
                tmVpf.iFTPArfCode = tmArf.iCode
            Else
                tmVpf.iFTPArfCode = 0
            End If
        End If
    Else
        If Trim$(smFTP) <> "" Then
            mInitArf
            tmArf.sFTP = smFTP
            ilRet = btrInsert(hmArf, tmArf, imArfRecLen, INDEXKEY0)
            tmVpf.iFTPArfCode = tmArf.iCode
        Else
            tmVpf.iFTPArfCode = 0
        End If
    End If

    'tmArfSrchKey.iCode = tmVpf.iAutoExptArfCode
    'If tmArfSrchKey.iCode <> 0 Then
    '    ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    '    If ilRet = BTRV_ERR_NONE Then
    '        If Trim$(smAutoExpt) <> "" Then
    '            tmArf.sFTP = smAutoExpt
    '            ilRet = btrUpdate(hmArf, tmArf, imArfRecLen)
    '        Else
    '            ilRet = btrDelete(hmArf)
    '            tmVpf.iAutoExptArfCode = 0
    '        End If
    '    Else
    '        If Trim$(smAutoExpt) <> "" Then
    '            mInitArf
    '            tmArf.sFTP = smAutoExpt
    '            ilRet = btrInsert(hmArf, tmArf, imArfRecLen, INDEXKEY0)
    '            tmVpf.iAutoExptArfCode = tmArf.iCode
    '        Else
    '            tmVpf.iAutoExptArfCode = 0
    '        End If
    '    End If
    'Else
    '    If Trim$(smAutoExpt) <> "" Then
    '        mInitArf
    '        tmArf.sFTP = smAutoExpt
    '        ilRet = btrInsert(hmArf, tmArf, imArfRecLen, INDEXKEY0)
    '        tmVpf.iAutoExptArfCode = tmArf.iCode
    '    Else
    '        tmVpf.iAutoExptArfCode = 0
    '    End If
    'End If
    tmVpf.iAutoExptArfCode = mUpdateArf(tmVpf.iAutoExptArfCode, smAutoExpt)

    'tmArfSrchKey.iCode = tmVpf.iAutoImptArfCode
    'If tmArfSrchKey.iCode <> 0 Then
    '    ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    '    If ilRet = BTRV_ERR_NONE Then
    '        If Trim$(smAutoImpt) <> "" Then
    '            tmArf.sFTP = smAutoImpt
    '            ilRet = btrUpdate(hmArf, tmArf, imArfRecLen)
    '        Else
    '            ilRet = btrDelete(hmArf)
    '            tmVpf.iAutoImptArfCode = 0
    '        End If
    '    Else
    '        If Trim$(smAutoImpt) <> "" Then
    '            mInitArf
    '            tmArf.sFTP = smAutoImpt
    '            ilRet = btrInsert(hmArf, tmArf, imArfRecLen, INDEXKEY0)
    '            tmVpf.iAutoImptArfCode = tmArf.iCode
    '        Else
    '            tmVpf.iAutoImptArfCode = 0
    '        End If
    '    End If
    'Else
    '    If Trim$(smAutoImpt) <> "" Then
    '        mInitArf
    '        tmArf.sFTP = smAutoImpt
    '        ilRet = btrInsert(hmArf, tmArf, imArfRecLen, INDEXKEY0)
    '        tmVpf.iAutoImptArfCode = tmArf.iCode
    '    Else
    '        tmVpf.iAutoImptArfCode = 0
    '    End If
    'End If
    tmVpf.iAutoImptArfCode = mUpdateArf(tmVpf.iAutoImptArfCode, smAutoImpt)

    'Log Form
    If ((tmVef.sType = "C") And (tmVef.iVefCode = 0)) Or (tmVef.sType = "L") Or (tmVef.sType = "A") Or (tmVef.sType = "G") Then
        'If (tgSpf.sGUseAffSys = "Y") And (igVpfType <> 1) Then
        If (igVpfType <> 1) Then
            For ilRow = 1 To 3 Step 1
                If ilRow = 1 Then
                    tmVof = tmLVof
                    slType = "L"
                ElseIf ilRow = 2 Then
                    tmVof = tmCVof
                    slType = "C"
                Else
                    tmVof = tmOVof
                    slType = "O"
                End If
                ilMatchCef = False
                tmCefSrchKey.lCode = tmVof.lHd1CefCode
                If tmCefSrchKey.lCode <> 0 Then
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        tmCef.lCode = 0
                    Else
                        slStr = gStripChr0(tmCef.sComment)
                        'If StrComp(Trim$(Left$(tmCef.sComment, tmCef.iStrLen)), Trim$(smLSave(1, ilRow)), vbTextCompare) = 0 Then
                        If StrComp(slStr, Trim$(smLSave(1, ilRow)), vbTextCompare) = 0 Then
                            ilMatchCef = True
                        End If
                    End If
                Else
                    tmCef.lCode = 0
                End If
                If Not ilMatchCef Then
                    'tmCef.iStrLen = Len(Trim$(smLSave(1, ilRow)))
                    tmCef.sComment = Trim$(smLSave(1, ilRow)) & Chr$(0) '& Chr$(0) 'sgTB
                    imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
                    'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
                    If Trim$(smLSave(1, ilRow)) <> "" Then
                        'DL-9/26/03
                        'Force new comment because of the double comments instead of testing to see if comment in more then one vof
                        'Double caused by modeling
                        'If tmCef.lCode = 0 Then
                            tmCef.lCode = 0 'Autoincrement
                            ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                        'Else
                        '    ilRet = btrUpdate(hmCef, tmCef, imCefRecLen)
                        'End If
                        If ilRet = BTRV_ERR_NONE Then
                            tmVof.lHd1CefCode = tmCef.lCode
                        Else
                            tmVof.lHd1CefCode = 0
                        End If
                    Else
                        If tmCef.lCode <> 0 Then
                            ilRet = btrDelete(hmCef)
                        End If
                        tmVof.lHd1CefCode = 0
                    End If
                End If
                ilMatchCef = False
                tmCefSrchKey.lCode = tmVof.lFt1CefCode
                If tmCefSrchKey.lCode <> 0 Then
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        tmCef.lCode = 0
                    Else
                        slStr = gStripChr0(tmCef.sComment)
                        'If StrComp(Trim$(Left$(tmCef.sComment, tmCef.iStrLen)), Trim$(smLSave(2, ilRow)), vbTextCompare) = 0 Then
                        If StrComp(slStr, Trim$(smLSave(2, ilRow)), vbTextCompare) = 0 Then
                            ilMatchCef = True
                        End If
                    End If
                Else
                    tmCef.lCode = 0
                End If
                If Not ilMatchCef Then
                    'tmCef.iStrLen = Len(Trim$(smLSave(2, ilRow)))
                    tmCef.sComment = Trim$(smLSave(2, ilRow)) & Chr$(0) '& Chr$(0) 'sgTB
                    imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
                    'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
                    If Trim$(smLSave(2, ilRow)) <> "" Then
                        'DL-9/26/03
                        'Force new comment because of the double comments instead of testing to see if comment in more then one vof
                        'Double caused by modeling
                        'If tmCef.lCode = 0 Then
                            tmCef.lCode = 0 'Autoincrement
                            ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                        'Else
                        '    ilRet = btrUpdate(hmCef, tmCef, imCefRecLen)
                        'End If
                        If ilRet = BTRV_ERR_NONE Then
                            tmVof.lFt1CefCode = tmCef.lCode
                        Else
                            tmVof.lFt1CefCode = 0
                        End If
                    Else
                        If tmCef.lCode <> 0 Then
                            ilRet = btrDelete(hmCef)
                        End If
                        tmVof.lFt1CefCode = 0
                    End If
                End If
                ilMatchCef = False
                tmCefSrchKey.lCode = tmVof.lFt2CefCode
                If tmCefSrchKey.lCode <> 0 Then
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        tmCef.lCode = 0
                    Else
                        slStr = gStripChr0(tmCef.sComment)
                        If StrComp(slStr, Trim$(smLSave(3, ilRow)), vbTextCompare) = 0 Then
                            ilMatchCef = True
                        End If
                    End If
                Else
                    tmCef.lCode = 0
                End If
                If Not ilMatchCef Then
                    'tmCef.iStrLen = Len(Trim$(smLSave(3, ilRow)))
                    tmCef.sComment = Trim$(smLSave(3, ilRow)) & Chr$(0) '& Chr$(0) 'sgTB
                    imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
                    'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
                    If Trim$(smLSave(3, ilRow)) <> "" Then
                        'DL-9/26/03
                        'Force new comment because of the double comments instead of testing to see if comment in more then one vof
                        'Double caused by modeling
                        'If tmCef.lCode = 0 Then
                            tmCef.lCode = 0 'Autoincrement
                            ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                        'Else
                        '    ilRet = btrUpdate(hmCef, tmCef, imCefRecLen)
                        'End If
                        If ilRet = BTRV_ERR_NONE Then
                            tmVof.lFt2CefCode = tmCef.lCode
                        Else
                            tmVof.lFt2CefCode = 0
                        End If
                    Else
                        If tmCef.lCode <> 0 Then
                            ilRet = btrDelete(hmCef)
                        End If
                        tmVof.lFt2CefCode = 0
                    End If
                End If
                                
                Do
                    tmVofSrchKey.iVefCode = tmVef.iCode
                    tmVofSrchKey.sType = slType
                    ilRet = btrGetEqual(hmVof, tlVof, imVofRecLen, tmVofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    slMsg = "cmcUpdate (btrGetEqual: VOF)"
                    On Error GoTo cmcUpdateErr
                    gBtrvErrorMsg ilRet, slMsg, VehOpt
                    On Error GoTo 0
                    ilRet = btrUpdate(hmVof, tmVof, imVofRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                slMsg = "cmcUpdate (btrUpdate: VOF)"
                On Error GoTo cmcUpdateErr
                gBtrvErrorMsg ilRet, slMsg, VehOpt
                On Error GoTo 0
                If ilRow = 1 Then
                    tmLVof = tmVof
                ElseIf ilRow = 2 Then
                    tmCVof = tmVof
                Else
                    tmOVof = tmVof
                End If
            Next ilRow
        End If
    End If
    'E-Mail
    If (igVpfType <> 1) Then
        ilMatchCef = False
        tmCefSrchKey.lCode = tmVpf.lEMailCefCode
        If tmCefSrchKey.lCode <> 0 Then
            tmCef.sComment = ""
            imCefRecLen = Len(tmCef)    '1009
            ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                tmCef.lCode = 0
            Else
                slStr = gStripChr0(tmCef.sComment)
                'If StrComp(Trim$(Left$(tmCef.sComment, tmCef.iStrLen)), Trim$(smEMail), vbTextCompare) = 0 Then
                If StrComp(slStr, Trim$(smEMail), vbTextCompare) = 0 Then
                    ilMatchCef = True
                End If
            End If
        Else
            tmCef.lCode = 0
        End If
        If Not ilMatchCef Then
            'tmCef.iStrLen = Len(Trim$(smEMail))
            tmCef.sComment = Trim$(smEMail) & Chr$(0) '& Chr$(0) 'sgTB
            imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
            'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
            If Trim$(smEMail) <> "" Then
                'DL-9/26/03
                'Force new comment because of the double comments instead of testing to see if comment in more then one vof
                'Double caused by modeling
                'If tmCef.lCode = 0 Then
                    tmCef.lCode = 0 'Autoincrement
                    ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                'Else
                '    ilRet = btrUpdate(hmCef, tmCef, imCefRecLen)
                'End If
                If ilRet = BTRV_ERR_NONE Then
                    tmVpf.lEMailCefCode = tmCef.lCode
                Else
                    tmVpf.lEMailCefCode = 0
                End If
            Else
                If tmCef.lCode <> 0 Then
                    ilRet = btrDelete(hmCef)
                End If
                tmVpf.lEMailCefCode = 0
            End If
        End If
    Else
        tmVpf.lEMailCefCode = 0
    End If

    'Schedule
    If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Then
        '2/6/20: Retain saf because using features bit values.  In the scheduling of spots, if LowPrice and HighPrice <= 0, site is used
        If tmSaf.iCode = 0 Then
            'If tmSaf.lLowPrice > 0 Then
                slMsg = "cmcUpdate (btrInsert: SAF)"
                ilRet = btrInsert(hmSaf, tmSaf, imSafRecLen, INDEXKEY0)
            'Else
            '    ilRet = BTRV_ERR_NONE
            'End If
        Else
            tmSafSrchKey1.iVefCode = tmVef.iCode
            ilRet = btrGetEqual(hmSaf, tlSaf, imSafRecLen, tmSafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                'If tmSaf.lLowPrice <= 0 Then
                '    slMsg = "cmcUpdate (btrDelete: SAF)"
                '    ilRet = btrDelete(hmSaf)
                '    ilRetFlag = mSafReadRec(tmVef.iCode)
                'Else
                    slMsg = "cmcUpdate (btrUpdate: SAF)"
                    ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
                'End If
            Else
                'If tmSaf.lLowPrice > 0 Then
                    slMsg = "cmcUpdate (btrInsert: SAF)"
                    ilRet = btrInsert(hmSaf, tmSaf, imSafRecLen, INDEXKEY0)
               ' Else
               '     ilRet = BTRV_ERR_NONE
               ' End If
            End If
        End If
        On Error GoTo cmcUpdateErr
        gBtrvErrorMsg ilRet, slMsg, VehOpt
        On Error GoTo 0
        ilRet = gObtainSAF()
    End If
    If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "N") Or (tmVef.sType = "R") Or (tmVef.sType = "G")) And ((Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS) Then
        If tmVaf.lCode = 0 Then
            slMsg = "cmcUpdate (btrInsert: VAF)"
            ilRet = btrInsert(hmVaf, tmVaf, imVafRecLen, INDEXKEY0)
        Else
            tmVafSrchKey1.iVefCode = tmVef.iCode
            ilRet = btrGetEqual(hmVaf, tlVaf, imVafRecLen, tmVafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                slMsg = "cmcUpdate (btrUpdate: VAF)"
                ilRet = btrUpdate(hmVaf, tmVaf, imVafRecLen)
            Else
                slMsg = "cmcUpdate (btrInsert: VAF)"
                ilRet = btrInsert(hmVaf, tmVaf, imVafRecLen, INDEXKEY0)
            End If
        End If
        On Error GoTo cmcUpdateErr
        gBtrvErrorMsg ilRet, slMsg, VehOpt
        On Error GoTo 0
    End If

    'Export X-Digital
    tlVff = tmVff
    ilRet = mVffReadRec(tmVef.iCode)
    If ilRet Then
        If tmVef.sType = "G" Then
            tmVff.lPledgeHdVtfCode = mSaveVtf("H")
            tmVff.lPledgeFtVtfCode = mSaveVtf("F")
            llRet = mSaveVtf("1")
            llRet = mSaveVtf("2")
        Else
            tmVff.lPledgeHdVtfCode = 0
            tmVff.lPledgeFtVtfCode = 0
        End If
        tmVff.sXDISCIPrefix = Trim$(edcXDISCIPrefix(0).Text)
        tmVff.sXDSISCIPrefix = Trim$(edcXDISCIPrefix(1).Text)
        If (ilInterfaceID <> Val(Trim$(edcInterfaceID(1).Text))) Then
            If (Val(Trim$(edcInterfaceID(1).Text)) > 0) Then
                If tmVff.sSentToXDSStatus <> "N" Then
                    tmVff.sSentToXDSStatus = "M"
                End If
            Else
                tmVff.sSentToXDSStatus = "Y"
            End If
        ElseIf (Val(Trim$(edcInterfaceID(1).Text)) = 0) Then
            tmVff.sSentToXDSStatus = "Y"
        End If
        'If smXDXMLForm = "ISCI" Then
        '    tmVff.sXDXMLForm = "P"
        '    tmVff.sXDProgCodeID = ""
        'ElseIf smXDXMLForm = "H#B#P#" Then
        If smXDXMLForm = "H#B#P#" Then
            tmVff.sXDXMLForm = "A"
            tmVff.sXDProgCodeID = Trim$(edcInterfaceID(0).Text)
        ElseIf smXDXMLForm = "H#B#" Then
            tmVff.sXDXMLForm = "S"
            tmVff.sXDProgCodeID = Trim$(edcInterfaceID(0).Text)
        Else
            tmVff.sXDXMLForm = ""
            tmVff.sXDProgCodeID = ""
        End If
        If ckcXDSave(0).Value = vbUnchecked Then
            tmVff.sXDSaveCF = "N"
        Else
            tmVff.sXDSaveCF = "Y"
        End If
        If ckcXDSave(1).Value = vbChecked Then
            tmVff.sXDSaveHDD = "Y"
        Else
            tmVff.sXDSaveHDD = "N"
        End If
        If ckcXDSave(2).Value = vbChecked Then
            tmVff.sXDSaveNAS = "Y"
        Else
            tmVff.sXDSaveNAS = "N"
        End If
        If ckcXDSave(4).Value = vbUnchecked Then
            tmVff.sXDSSaveCF = "N"
        Else
            tmVff.sXDSSaveCF = "Y"
        End If
        If ckcXDSave(5).Value = vbChecked Then
            tmVff.sXDSSaveHDD = "Y"
        Else
            tmVff.sXDSSaveHDD = "N"
        End If
        If ckcXDSave(6).Value = vbChecked Then
            tmVff.sXDSSaveNAS = "Y"
        Else
            tmVff.sXDSSaveNAS = "N"
        End If
        If udcVehOptTabs.AffLog(0) = vbChecked Then
            tmVff.sMGsOnWeb = "Y"
        Else
            tmVff.sMGsOnWeb = "N"
        End If
        If udcVehOptTabs.AffLog(1) = vbChecked Then
            tmVff.sReplacementOnWeb = "Y"
        Else
            tmVff.sReplacementOnWeb = "N"
        End If
        If udcVehOptTabs.AffLog(2) = vbChecked Then
            tmVff.sHideCommOnWeb = "Y"
        Else
            tmVff.sHideCommOnWeb = "N"
        End If
        tmVff.sAirWavePrgID = Trim$(edcExport(0).Text)
        tmVff.sIPumpEventTypeOV = Trim$(edcExport(1).Text)
        If ckcIFExport(7).Value = vbChecked Then
            tmVff.sExportAirWave = "Y"
        Else
            tmVff.sExportAirWave = "N"
        End If
        If ckcIFExport(8).Value = vbChecked Then
            tmVff.sExportAudio = "Y"
        Else
            tmVff.sExportAudio = "N"
        End If
        If ckcIFExport(9).Value = vbChecked Then
            tmVff.sExportMP2 = "Y"
        Else
            tmVff.sExportMP2 = "N"
        End If
        If ckcIFExport(10).Value = vbChecked Then
            tmVff.sExportCnCSpot = "Y"
        Else
            tmVff.sExportCnCSpot = "N"
        End If
        If ckcIFExport(11).Value = vbChecked Then
            tmVff.sExportEnco = "Y"
        Else
            tmVff.sExportEnco = "N"
        End If
        If ckcIFExport(12).Value = vbChecked Then
            tmVff.sExportNYESPN = "Y"
        Else
            tmVff.sExportNYESPN = "N"
        End If
        If ckcIFExport(13).Value = vbChecked Then
            tmVff.sPledgeVsAir = "Y"
        Else
            tmVff.sPledgeVsAir = "N"
        End If
        If ckcIFExport(14).Value = vbChecked Then
            tmVff.sExportEncoESPN = "Y"
        Else
            tmVff.sExportEncoESPN = "N"
        End If
        If ckcIFExport(15).Value = vbChecked Then
            tmVff.sExportCnCNetInv = "Y"
        Else
            tmVff.sExportCnCNetInv = "N"
        End If
        If ckcIFExport(16).Value = vbChecked Then
            tmVff.sExportMatrix = "Y"
        Else
            tmVff.sExportMatrix = "N"
        End If
        If ckcIFExport(18).Value = vbChecked Then
            tmVff.sExportTableau = "Y"
        Else
            tmVff.sExportTableau = "N"
        End If
        If ckcIFExport(17).Value = vbChecked Then
            tmVff.sExportJelli = "Y"
        Else
            tmVff.sExportJelli = "N"
        End If
        If ckcAffExport(3).Value = vbChecked Then
            tmVff.sExportIPump = "Y"
        Else
            tmVff.sExportIPump = "N"
        End If
        If ckcAffExport(4).Value = vbChecked Then
            tmVff.sStationComp = "Y"
        Else
            tmVff.sStationComp = "N"
        End If
        For ilVff = LBound(tlVff.sFedDelivery) To UBound(tlVff.sFedDelivery) Step 1
            tmVff.sFedDelivery(ilVff) = tlVff.sFedDelivery(ilVff)
        Next ilVff
        
        tmVff.sMoveSportToNon = "N"
        tmVff.sMoveSportToSport = "N"
        tmVff.sMoveNonToSport = "N"
        tmVff.sPledgeByEvent = "N"
        If tmVef.sType = "G" Then
            If ckcSch(0).Value = vbChecked Then
                tmVff.sMoveSportToNon = "Y"
            End If
            If ckcSch(1).Value = vbChecked Then
                tmVff.sMoveSportToSport = "Y"
            End If
            If ckcSch(2).Value = vbChecked Then
                tmVff.sMoveNonToSport = "Y"
            End If
            If ckcSch(3).Value = vbChecked Then
                tmVff.sPledgeByEvent = "Y"
            End If
            If edcSchedule(4).Text <> "" Then
                tmVff.iPledgeClearance = edcSchedule(4).Text
            Else
                tmVff.iPledgeClearance = 0
            End If
        End If
        tmVff.sMergeTraffic = ""
        tmVff.sMergeAffiliate = ""
        tmVff.sMergeWeb = ""
        If tmVef.iVefCode > 0 Then
            If rbcMerge(1).Value Then
                tmVff.sMergeTraffic = "S"
            Else
                tmVff.sMergeTraffic = "M"
            End If
            If rbcMerge(3).Value Then
                tmVff.sMergeAffiliate = "S"
            Else
                tmVff.sMergeAffiliate = "M"
            End If
            If rbcMerge(5).Value Then
                tmVff.sMergeWeb = "S"
            Else
                tmVff.sMergeWeb = "M"
            End If
        End If
        tmVff.lSeasonGhfCode = 0
        If lbcSeason.ListIndex >= 0 Then
            tmVff.lSeasonGhfCode = lbcSeason.ItemData(lbcSeason.ListIndex)
        End If
        tmVff.sWebName = edcLog(3).Text
        tmVff.iMcfCode = 0
        If cbcMedia.ListIndex > 0 Then
            slNameCode = tmMediaCode(cbcMedia.ListIndex - 1).sKey  'lbcMediaCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If ilRet = CP_MSG_NONE Then
                tmVff.iMcfCode = Val(slCode)
            End If
        End If

        ilMatchCef = False
        tmCefSrchKey.lCode = tmVff.lBBOpenCefCode
        If tmCefSrchKey.lCode <> 0 Then
            tmCef.sComment = ""
            imCefRecLen = Len(tmCef)    '1009
            ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                tmCef.lCode = 0
            Else
                slStr = gStripChr0(tmCef.sComment)
                If StrComp(slStr, Trim$(edcBB(0).Text), vbTextCompare) = 0 Then
                    ilMatchCef = True
                End If
            End If
        Else
            tmCef.lCode = 0
        End If
        If Not ilMatchCef Then
            tmCef.sComment = Trim$(edcBB(0)) & Chr$(0) '& Chr$(0) 'sgTB
            imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
            If Trim$(edcBB(0)) <> "" Then
                'DL-9/26/03
                'Force new comment because of the double comments instead of testing to see if comment in more then one vof
                'Double caused by modeling
                tmCef.lCode = 0 'Autoincrement
                ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                If ilRet = BTRV_ERR_NONE Then
                    tmVff.lBBOpenCefCode = tmCef.lCode
                Else
                    tmVff.lBBOpenCefCode = 0
                End If
            Else
                If tmCef.lCode <> 0 Then
                    ilRet = btrDelete(hmCef)
                End If
                tmVff.lBBOpenCefCode = 0
            End If
        End If

        ilMatchCef = False
        tmCefSrchKey.lCode = tmVff.lBBCloseCefCode
        If tmCefSrchKey.lCode <> 0 Then
            tmCef.sComment = ""
            imCefRecLen = Len(tmCef)    '1009
            ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                tmCef.lCode = 0
            Else
                slStr = gStripChr0(tmCef.sComment)
                If StrComp(slStr, Trim$(edcBB(1).Text), vbTextCompare) = 0 Then
                    ilMatchCef = True
                End If
            End If
        Else
            tmCef.lCode = 0
        End If
        If Not ilMatchCef Then
            tmCef.sComment = Trim$(edcBB(1)) & Chr$(0) '& Chr$(0) 'sgTB
            imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
            If Trim$(edcBB(1)) <> "" Then
                'DL-9/26/03
                'Force new comment because of the double comments instead of testing to see if comment in more then one vof
                'Double caused by modeling
                tmCef.lCode = 0 'Autoincrement
                ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                If ilRet = BTRV_ERR_NONE Then
                    tmVff.lBBCloseCefCode = tmCef.lCode
                Else
                    tmVff.lBBCloseCefCode = 0
                End If
            Else
                If tmCef.lCode <> 0 Then
                    ilRet = btrDelete(hmCef)
                End If
                tmVff.lBBCloseCefCode = 0
            End If
        End If
    
        tmVff.iLiveCompliantAdj = Val(smLiveWindow)
        '8032
'        If (rbcBarterMethod(7).Value) And (tmVef.sType = "R") Then
'            tmVff.sOnXMLInsertion = "W"
'        Else
'            tmVff.sOnXMLInsertion = "N"
'        End If
        tmVff.sOnXMLInsertion = "N"
        '8132
        If tmVef.sType = "R" Or tmVef.sType = "C" Then
            If (rbcBarterMethod(STATIONXMLWIDEORBIT).Value) Then
                tmVff.sOnXMLInsertion = "W"
            ElseIf rbcBarterMethod(STATIONXMLMARKETRON).Value Then
                tmVff.sOnXMLInsertion = "M"
            End If
        End If
                
        If (tmVef.sType = "R") Or (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "N") Then
            If ckcBarter(0).Value = vbChecked Then
                tmVff.sOnInsertions = "Y"
            Else
                tmVff.sOnInsertions = "N"
            End If
        Else
            tmVff.sOnInsertions = "N"
        End If
        tmVff.sStationPassword = ""
        If (tmVef.sType = "C") Or (tmVef.sType = "S") Then
            If rbcBarterMethod(6).Value Then
                tmVff.sPostLogSource = "S"
                tmVff.sStationPassword = edcBarter(5).Text
            Else
                tmVff.sPostLogSource = "N"
            End If
        Else
            tmVff.sPostLogSource = "N"
        End If
        tmVff.iConflictWinLen = 0
        If (tmVef.sType = "A") Then
            If ckcLog(0).Value = vbChecked Then
                tmVff.sHonorZeroUnits = "Y"
            Else
                tmVff.sHonorZeroUnits = "N"
            End If
            If edcSchedule(5).Text <> "" Then
                tmVff.iConflictWinLen = Val(edcSchedule(5).Text)
            End If
        Else
            tmVff.sHonorZeroUnits = "N"
        End If
        If (tmVef.sType <> "S") Then
            If ckcLog(1).Value = vbChecked Then
                tmVff.sHideCommOnLog = "Y"
            Else
                tmVff.sHideCommOnLog = "N"
            End If
        Else
            tmVff.sHideCommOnLog = "N"
        End If
        '2/28/19: Add Cart on Web
        tmVff.sCartOnWeb = smCartOnWeb
        If rbcAudioType(1).Value = True Then
            tmVff.sDefaultAudioType = "L"
        ElseIf rbcAudioType(2).Value = True Then
            tmVff.sDefaultAudioType = "M"
        ElseIf rbcAudioType(3).Value = True Then
            tmVff.sDefaultAudioType = "S"
        ElseIf rbcAudioType(4).Value = True Then
            tmVff.sDefaultAudioType = "P"
        ElseIf rbcAudioType(5).Value = True Then
            tmVff.sDefaultAudioType = "Q"
        Else
            tmVff.sDefaultAudioType = "R"
        End If
        
'        'TTP 9992
        If ckcIFExport(20).Value = vbChecked Then
            tmVff.sExportCustom = "Y"
        Else
            tmVff.sExportCustom = "N"
        End If
        
        smLogExpt = Trim$(edcLog(0).Text)
        tmVff.iLogExptArfCode = mUpdateArf(tmVff.iLogExptArfCode, smLogExpt)
        tmVff.sASICallLetters = Trim$(edcGen(VEDICALLLETTERS).Text)
        tmVff.sASIBand = Trim$(edcGen(VEDIBAND).Text)
        '10981
        tmVff.lAdVehNameCefCode = 0
        '10050 podcast  do we have a previously defined advendorName? If so, may have to delete.
        If rbcGMedium(PODCASTRBC).Value Then
            'excludes 'none' and 'new'
            If cbcCsiGeneric(ADVENDOR).ListIndex > 1 Then
                tmVff.iAvfCode = cbcCsiGeneric(ADVENDOR).GetItemData(cbcCsiGeneric(ADVENDOR).ListIndex)
                '10981
                If Not mWriteToVVC() Then
                    'issue
                End If
            End If
            '10981
           ' tmVff.lAdVehNameCefCode = mAdVendorNameToCef()
        '10894
        'was podcast
        ElseIf tmVpf.sGMedium = "P" Then
            tmVff.iAvfCode = 0
        '10981 noticed double test of 'sGMedium'
'            'was podcast
'            If tmVpf.sGMedium = "P" Then
'                'test if used in pcf and block saving if so.
'                If tmVpf.sGMedium = "P" Then
'                    tmVff.iAvfCode = 0
'                'delete cef code as needed
'                ElseIf tmVff.lAdVehNameCefCode > 0 Then
'                    tmVff.lAdVehNameCefCode = mAdVendorNameToCef()
'                End If
'            End If
        End If
        '10933
        If ckcXDSave(CUEZONE).Value = vbChecked Then
            tmVff.sXDEventZone = "Y"
        Else
            tmVff.sXDEventZone = "N"
        End If
        slMsg = "cmcUpdate (btrUpdate: VFF)"
        ilRet = btrUpdate(hmVff, tmVff, imVffRecLen)
        On Error GoTo cmcUpdateErr
        gBtrvErrorMsg ilRet, slMsg, VehOpt
        On Error GoTo 0
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If tmVff.iCode = tgVff(ilVff).iCode Then
                tgVff(ilVff) = tmVff
                Exit For
            End If
        Next ilVff
        '9114
        mXdsVendorChange
    End If

    ilValue = Asc(tgSpf.sUsingFeatures2)
    ''If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G")) And (ilValue And BARTER) = BARTER Then
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S")) And (ilValue And BARTER) = BARTER Then
    If (tmVef.sType = "R") And ((ilValue And BARTER) = BARTER) Then
        For ilVbf = 0 To UBound(tmVbf) - 1 Step 1
            gUnpackDateLong tmVbf(ilVbf).iStartDate(0), tmVbf(ilVbf).iStartDate(1), llStartDate
            If llStartDate > 0 Then
                ilMatchCef = False
                tmCefSrchKey.lCode = tmVbf(ilVbf).lInsertionCefCode
                If tmCefSrchKey.lCode <> 0 Then
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        tmCef.lCode = 0
                    Else
                        slStr = gStripChr0(tmCef.sComment)
                        'If StrComp(Trim$(Left$(tmCef.sComment, tmCef.iStrLen)), Trim$(smVbfComment(ilVbf)), vbTextCompare) = 0 Then
                        If StrComp(slStr, Trim$(smVbfComment(ilVbf)), vbTextCompare) = 0 Then
                            ilMatchCef = True
                        End If
                    End If
                Else
                    tmCef.lCode = 0
                End If
                If Not ilMatchCef Then
                    'tmCef.iStrLen = Len(Trim$(smVbfComment(ilVbf)))
                    tmCef.sComment = Trim$(smVbfComment(ilVbf)) & Chr$(0) '& Chr$(0) 'sgTB
                    imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
                    'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
                    If Trim$(smVbfComment(ilVbf)) <> "" Then
                        'DL-9/26/03
                        'Force new comment because of the double comments instead of testing to see if comment in more then one vof
                        'Double caused by modeling
                        'If tmCef.lCode = 0 Then
                            tmCef.lCode = 0 'Autoincrement
                            ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                        'Else
                        '    ilRet = btrUpdate(hmCef, tmCef, imCefRecLen)
                        'End If
                        If ilRet = BTRV_ERR_NONE Then
                            tmVbf(ilVbf).lInsertionCefCode = tmCef.lCode
                        Else
                            tmVbf(ilVbf).lInsertionCefCode = 0
                        End If
                    Else
                        If tmCef.lCode <> 0 Then
                            ilRet = btrDelete(hmCef)
                        End If
                        tmVbf(ilVbf).lInsertionCefCode = 0
                    End If
                End If
                If tmVbf(ilVbf).lCode > 0 Then
                    slMsg = "cmcUpdate (btrGetEqual: VBF)"
                    tmVbfSrchKey.lCode = tmVbf(ilVbf).lCode
                    ilRet = btrGetEqual(hmVbf, tlVbf, imVBfRecLen, tmVbfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        gUnpackDateLong tmVbf(ilVbf).iEndDate(0), tmVbf(ilVbf).iEndDate(1), llEndDate

                        If (llEndDate < llStartDate) And (llEndDate <> 0) Then
                            slMsg = "cmcUpdate (btrDelete: VBF)"
                            ilRet = btrDelete(hmVbf)
                        Else
                            slMsg = "cmcUpdate (btrUpdate: VBF)"
                            ilRet = btrUpdate(hmVbf, tmVbf(ilVbf), imVBfRecLen)
                        End If
                    End If
                Else
                    slMsg = "cmcUpdate (btrInsert: VBF)"
                    ilRet = btrInsert(hmVbf, tmVbf(ilVbf), imVBfRecLen, INDEXKEY0)
                End If
                On Error GoTo cmcUpdateErr
                gBtrvErrorMsg ilRet, slMsg, VehOpt
                On Error GoTo 0
            End If
        Next ilVbf
    ElseIf ((tmVef.sType = "C") Or (tmVef.sType = "S")) And ((ilValue And BARTER) = BARTER) Then
        '5/4/20: The handling of the Acq adjustment by spot length is not consistent
        '        It is added only with vehicle type C or S but in the contract when clicking in the acq override field, it is testing for Vehicle type R
        '        see mSetDefAcqCost and mAcqCostPop
        '        this record type is method I, whereas the commission record method is N
        If tmVbfIndex.lCode > 0 Then
            slMsg = "cmcUpdate (btrGetEqual: VBF)"
            tmVbfSrchKey.lCode = tmVbfIndex.lCode
            ilRet = btrGetEqual(hmVbf, tlVbf, imVBfRecLen, tmVbfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                slMsg = "cmcUpdate (btrUpdate: VBF)"
                ilRet = btrUpdate(hmVbf, tmVbfIndex, imVBfRecLen)
            End If
        Else
            slMsg = "cmcUpdate (btrInsert: VBF)"
            ilRet = btrInsert(hmVbf, tmVbfIndex, imVBfRecLen, INDEXKEY0)
        End If
        On Error GoTo cmcUpdateErr
        gBtrvErrorMsg ilRet, slMsg, VehOpt
        On Error GoTo 0
        For ilVbf = 0 To UBound(tmVbf) - 1 Step 1
            gUnpackDateLong tmVbf(ilVbf).iStartDate(0), tmVbf(ilVbf).iStartDate(1), llStartDate
            If llStartDate > 0 Then
                If tmVbf(ilVbf).lCode > 0 Then
                    slMsg = "cmcUpdate (btrGetEqual: VBF)"
                    tmVbfSrchKey.lCode = tmVbf(ilVbf).lCode
                    ilRet = btrGetEqual(hmVbf, tlVbf, imVBfRecLen, tmVbfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        gUnpackDateLong tmVbf(ilVbf).iEndDate(0), tmVbf(ilVbf).iEndDate(1), llEndDate
        
                        If (llEndDate < llStartDate) And (llEndDate <> 0) Then
                            slMsg = "cmcUpdate (btrDelete: VBF)"
                            ilRet = btrDelete(hmVbf)
                        Else
                            slMsg = "cmcUpdate (btrUpdate: VBF)"
                            ilRet = btrUpdate(hmVbf, tmVbf(ilVbf), imVBfRecLen)
                        End If
                    End If
                Else
                    slMsg = "cmcUpdate (btrInsert: VBF)"
                    ilRet = btrInsert(hmVbf, tmVbf(ilVbf), imVBfRecLen, INDEXKEY0)
                End If
                On Error GoTo cmcUpdateErr
                gBtrvErrorMsg ilRet, slMsg, VehOpt
                On Error GoTo 0
            End If
        Next ilVbf
    End If
    'Update Participant
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Or (tmVef.sType = "T") Or (tmVef.sType = "R") Or (tmVef.sType = "N")) And (igVpfType <> 1) Then
    If (igVpfType <> 1) And (tmVef.sType <> "L") Then
        llOwnerDate = 0
        ilOwnerMnfCode = 0
        ilSSMnfCode = 0
        For ilPif = 0 To UBound(tmPifRec) - 1 Step 1
           If tmPifRec(ilPif).tPif.lCode >= 0 Then
                gUnpackDateLong tmPifRec(ilPif).tPif.iStartDate(0), tmPifRec(ilPif).tPif.iStartDate(1), llDate
                If (llDate > llOwnerDate) And (tmPifRec(ilPif).tPif.iSeqNo = 1) Then
                    ilOwnerMnfCode = tmPifRec(ilPif).tPif.iMnfGroup
                    llOwnerDate = llDate
                    ilSSMnfCode = tmPifRec(ilPif).tPif.iMnfSSCode
                End If
            End If
            If tmPifRec(ilPif).tPif.lCode < 0 Then
                'Delete record
                Do
                    tmPifSrchKey.lCode = -tmPifRec(ilPif).tPif.lCode
                    ilRet = btrGetEqual(hmPif, tlPif, imPifRecLen, tmPifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    slMsg = "cmcUpdate (btrGetEqual: PIF)"
                    On Error GoTo cmcUpdateErr
                    gBtrvErrorMsg ilRet, slMsg, VehOpt
                    On Error GoTo 0
                    ilRet = btrDelete(hmPif)
                Loop While ilRet = BTRV_ERR_CONFLICT
                slMsg = "cmcUpdate (btrDelete: PIF)"
            ElseIf tmPifRec(ilPif).tPif.lCode = 0 Then
                'Insert record
                ilRet = btrInsert(hmPif, tmPifRec(ilPif).tPif, imPifRecLen, INDEXKEY0)
                slMsg = "cmcUpdate (btrInsert: PIF)"
            ElseIf tmPifRec(ilPif).tPif.lCode > 0 Then
                'Update record
                Do
                    tmPifSrchKey.lCode = tmPifRec(ilPif).tPif.lCode
                    ilRet = btrGetEqual(hmPif, tlPif, imPifRecLen, tmPifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    slMsg = "cmcUpdate (btrGetEqual: PIF)"
                    On Error GoTo cmcUpdateErr
                    gBtrvErrorMsg ilRet, slMsg, VehOpt
                    On Error GoTo 0
                    ilRet = btrUpdate(hmPif, tmPifRec(ilPif).tPif, imPifRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                slMsg = "cmcUpdate (btrUpdate: PIF)"
            End If
            On Error GoTo cmcUpdateErr
            gBtrvErrorMsg ilRet, slMsg, VehOpt
            On Error GoTo 0
        Next ilPif
        Do
            tmVefSrchKey.iCode = tmVef.iCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tmVef.iOwnerMnfCode = ilOwnerMnfCode
                tmVef.iSSMnfCode = ilSSMnfCode
                If ckcIFExport(19).Value = vbChecked Then
                    tmVef.sExportRAB = "Y"
                Else
                    tmVef.sExportRAB = "N"
                End If
                ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If tmVef.iCode = tgMVef(ilVef).iCode Then
                tgMVef(ilVef).iOwnerMnfCode = ilOwnerMnfCode
                tgMVef(ilVef).iSSMnfCode = ilSSMnfCode
                If ckcIFExport(19).Value = vbChecked Then
                    tgMVef(ilVef).sExportRAB = "Y"
                Else
                    tgMVef(ilVef).sExportRAB = "N"
                End If
                Exit For
            End If
        Next ilVef
        
        '11/26/17
        gFileChgdUpdate "vef.btr", False

    End If
    imVpfChanged = True
    mMoveRec tmVpf, tgVpf(imVpfIndex)
    'gGetSyncDateTime slSyncDate, slSyncTime
    'gPackDate slSyncDate, tgVpf(imVpfIndex).iSyncDate(0), tgVpf(imVpfIndex).iSyncDate(1)
    'gPackTime slSyncTime, tgVpf(imVpfIndex).iSyncTime(0), tgVpf(imVpfIndex).iSyncTime(1)
    ilRet = btrUpdate(hlVpf, tgVpf(imVpfIndex), ilRecLen)
    slMsg = "cmcUpdate (btrUpdate)"
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, slMsg, VehOpt
    On Error GoTo 0
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    ilRet = btrClose(hlVpf)
    On Error GoTo cmcUpdateErr
    gBtrvErrorMsg ilRet, "cmcUpdate (btrClose): " & "Vpf.Btr", VehOpt
    On Error GoTo 0
    btrDestroy hlVpf
    
    '11/26/17
    gFileChgdUpdate "vpf.btr", True
    
    If tmVef.sType = "V" Then
        ilRet = mVirtSaveRec()
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
    End If
    
    '10/16/15: Moved to vehicles mSaveRec
    ''07-13-15
    ''Test for vehicles being stations and if on insertions enabled update network link else remove network link
    ''verified 8/25/15
    'If bgEDSIsActive Then
    '    ilRet = gUpdateVehicle(tmVef.iCode)
    'End If
    
    imPartMissAndReq = False
    If (igVpfType <> 1) And (tmVef.sType <> "L") Then
        mPopParticipantDates
        If UBound(tmPifRec) <= LBound(tmPifRec) Then
            imPartMissAndReq = True
        End If
    Else
        ReDim tmPifRec(0 To 0) As PIFREC
    End If
    igVehNewToVehOpt = False
    'smInitLgVehNm = edcLgVehNm.Text
    'smInitLgHd1 = edcLgHd1.Text
    'smInitLgFt1 = edcLgFt1.Text
    'smInitLgFt2 = edcLgFt2.Text
    mClearBarterCtrl False
    ilRet = mVbfReadRec(tmVef.iCode)
    If ((tmVef.sType = "C") Or (tmVef.sType = "S")) And ((ilValue And BARTER) = BARTER) Then
        mAcqIndexPaint
    End If
    imVffChg = False
    imVbfChg = False
    imPifChg = False
    imVefChg = False
    imLogAltered = False
    imLevelAltered = False
    imProducerAltered = False
    imInventoryAltered = False          '7-21-05
    imGreatPlainsAltered = False
    imVirtBoxNo = -1
    imVirtRowNo = -1
    imVirtChgVeh = False
    imLBoxNo = -1
    imLRowNo = -1
    imIgnoreChg = True
    mSetCommands
    imIgnoreChg = False
    Screen.MousePointer = vbDefault
    cbcSelect.Enabled = True
    cbcSelect.SetFocus
    Exit Sub
cmcUpdateErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    ilRet = btrClose(hlVpf)
    btrDestroy hlVpf
    Resume Next
End Sub
Private Sub cmcUpdate_GotFocus()
    mRemoveFocus
    gCtrlGotFocus cmcUpdate
End Sub

Private Sub csiParticipantDate_GotFocus()
    csiParticipantDate.ZOrder vbBringToFront
    If imCbcParticipantListIndex = -1 Then
        DoEvents
        cbcParticipant.SetFocus
        DoEvents
        Exit Sub
    End If
    smOrigPartStartDate = csiParticipantDate.Text
    If csiParticipantDate.Text = "" Then
        csiParticipantDate.Text = gIncOneDay(gObtainEndStd(gNow()))
    End If
    mPartSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub csiParticipantDate_Validate(Cancel As Boolean)
    Dim slDate As String
    If bmInDateChg Then
        Exit Sub
    End If
    If csiParticipantDate.Text <> "" Then
        bmInDateChg = True
        slDate = csiParticipantDate.Text
        csiParticipantDate.Text = gObtainStartStd(slDate)
        bmInDateChg = False
    End If
End Sub

Private Sub edcARBCode_Change()
    mSetCommands
End Sub

Private Sub edcARBCode_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub edcBarter_Change(Index As Integer)
    If Index <= 2 Then
        imVbfChg = True
    ElseIf Index = 3 Then
        If imAcqCostBoxNo >= 0 Then
            If imAcqCostBoxNo <= (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD) - 1 Then
                If tmVbf(imVBFIndex).lDefAcqCost(imAcqCostBoxNo - ADJBD + 1) <> gStrDecToLong(edcBarter(3).Text, 2) Then
                    imVbfChg = True
                End If
            Else
                If tmVbf(imVBFIndex).lActAcqCost(imAcqCostBoxNo - (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD) + 1 - ADJBD) <> gStrDecToLong(edcBarter(3).Text, 2) Then
                    imVbfChg = True
                End If
            End If
        End If
    ElseIf Index = 4 Then
        If imAcqIndexBoxNo >= 0 Then
            If tmVbfIndex.lDefAcqCost(imAcqIndexBoxNo + 1 - ADJBD) <> gStrDecToLong(edcBarter(4).Text, 2) Then
                imVbfChg = True
            End If
        End If
    ElseIf Index = 5 Then
        imVbfChg = True
    End If
    mSetCommands
End Sub

Private Sub edcBarter_GotFocus(Index As Integer)
    If Index <= 2 Then
        imAcqCostBoxNo = -1
        imAcqIndexBoxNo = -1
    ElseIf Index = 3 Then
        imAcqIndexBoxNo = -1
    ElseIf Index = 4 Then
        imAcqCostBoxNo = -1
    ElseIf Index = 5 Then
        If Not bmCitationDefined Then
            MsgBox "The Station Posting Citation in Site---->Comment tab must be defined prior to entering the User Password", vbOKOnly + vbExclamation, "Vehicle Option"
            pbcClickFocus.SetFocus
        End If
        imAcqCostBoxNo = -1
    End If
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcBarter_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilPos As Integer
    
    If Index <= 1 Then
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    ElseIf Index = 2 Then
        ilPos = InStr(edcBarter(2).SelText, ".")
        If ilPos = 0 Then
            ilPos = InStr(edcBarter(2).Text, ".")    'Disallow multi-decimal points
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
    ElseIf Index = 3 Then
        If imAcqCostBoxNo >= 0 Then
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
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
    ElseIf Index = 4 Then
        If imAcqIndexBoxNo >= 0 Then
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
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
    End If
End Sub

Private Sub edcBarter_LostFocus(Index As Integer)
    If Index = 3 Then
        If imAcqCostBoxNo >= 0 Then
            If imAcqCostBoxNo <= (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD) - 1 Then
                tmVbf(imVBFIndex).lDefAcqCost(imAcqCostBoxNo + 1 - ADJBD) = gStrDecToLong(edcBarter(3).Text, 2)
            Else
                tmVbf(imVBFIndex).lActAcqCost(imAcqCostBoxNo - (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD) + 1 - ADJBD) = gStrDecToLong(edcBarter(3).Text, 2)
            End If
        End If
    ElseIf Index = 4 Then
        If imAcqIndexBoxNo >= 0 Then
            tmVbfIndex.lDefAcqCost(imAcqIndexBoxNo + 1 - ADJBD) = gStrDecToLong(edcBarter(4).Text, 2)
        End If
    End If
End Sub

Private Sub edcBarterMethod_Change(Index As Integer)
    imVbfChg = True
    mSetCommands
End Sub

Private Sub edcBarterMethod_GotFocus(Index As Integer)
    imAcqCostBoxNo = -1
    imAcqIndexBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcBarterMethod_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcBB_Change(Index As Integer)
    imVffChg = True
    mSetCommands
End Sub

Private Sub edcBB_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_Change()
    If Not imIgnoreChg Then
        imVirtChgVeh = True
    End If
    mSetCommands
End Sub
Private Sub edcDropDown_GotFocus()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slTotalPct As String
    If imVirtBoxNo = PERCENTINDEX Then
        slTotalPct = "0"
        If edcDropDown.Text = "" Then
            For ilLoop = LBONE To UBound(smVirtSave, 2) - 1 Step 1
                slStr = smVirtSave(3, ilLoop)
                If slStr <> "" Then
                    slTotalPct = gAddStr(slStr, slTotalPct)
                End If
            Next ilLoop
            If Val(slTotalPct) <= 100 Then
                edcDropDown.Text = gSubStr("100.0000", slTotalPct)
            Else
                edcDropDown.Text = "0"
            End If
            'imVirtChgVeh = True
            mSetCommands
        End If
    End If
    gCtrlGotFocus edcDropDown
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    If imVirtBoxNo = NOSPOTSINDEX Then
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    ElseIf imVirtBoxNo = PERCENTINDEX Then
        ilPos = InStr(edcDropDown.SelText, ".")
        If ilPos = 0 Then
            ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
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
    End If
End Sub

Private Sub edcEDASWindow_Change()
    mSetCommands
End Sub

Private Sub edcEDASWindow_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcEDASWindow_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcEMail_Change()
    mSetCommands
End Sub

Private Sub edcEMail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcExpCntrVehNo_Change()
    mSetCommands
End Sub
Private Sub edcExpCntrVehNo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcExpCntrVehNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub edcExport_Change(Index As Integer)
    imVffChg = True
    mSetCommands
End Sub

Private Sub edcExport_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub


Private Sub edcGEDI_Change(Index As Integer)
    'If Index >= 2 Then
    '    imVffChg = True
    'End If
    'mSetCommands
End Sub
Private Sub edcGEDI_GotFocus(Index As Integer)
    ''10050
    'If Index <> ADVENDORVEHICLENAME Then
    '    edcGTZDropDown_LostFocus
    '    pbcGTZToggle.Visible = False
    '    lbcFeed.Visible = False
    '    cmcGTZDropDown.Visible = False
    '    edcGTZDropDown.Visible = False
    '    imGTZBoxNo = -1
    '    gCtrlGotFocus ActiveControl
    'End If
End Sub

Private Sub edcGEDI_KeyPress(Index As Integer, KeyAscii As Integer)
    ''10050
    'If Index <> ADVENDORVEHICLENAME Then
    '    'Convert to upper case
    '    If KeyAscii >= 97 And KeyAscii <= 122 Then
    '        KeyAscii = KeyAscii - 32
    '    End If
    'End If
End Sub

Private Sub edcGen_Change(Index As Integer)
    Dim ilRes As Integer

    If (Index >= VEDICALLLETTERS And Index <= VEDIBAND) Or Index = ADVENDOREXTERNALIDINDEX Then
        'turns out this isn't tied to vff in any way, but mSetCommands know a change happened.
        imVffChg = True
    End If
    If Index = SAGROUPNO Then
        If imAskedSAGroupNo Then
            If (tmVef.sType = "S") Or (tmVef.sType = "A") Then
                If imOrigSAGroupNo <> Val(edcGen(SAGROUPNO).Text) Then
                    imAskedSAGroupNo = False
                    ilRes = MsgBox("Before changing the Group #, the links should be terminated.  Have you terminated the Links?", vbYesNo + vbQuestion, "Update")
                    If ilRes = vbNo Then
                        edcGen(SAGROUPNO).Text = Trim$(Str$(imOrigSAGroupNo))
                    End If
                End If
            End If
        End If
    End If
    mSetCommands
End Sub

Private Sub edcGen_GotFocus(Index As Integer)
    Dim slStr As String
    '10050
    If Index <> ADVENDOREXTERNALIDINDEX Then
        edcGTZDropDown_LostFocus
        pbcGTZToggle.Visible = False
        lbcFeed.Visible = False
        cmcGTZDropDown.Visible = False
        edcGTZDropDown.Visible = False
        imGTZBoxNo = -1
        If Index = Signon Then
            slStr = edcGen(Signon).Text
            slStr = gUnformatTime(slStr)
            edcGen(Signon).Text = slStr
        End If
        gCtrlGotFocus ActiveControl
    End If
End Sub

Private Sub edcGen_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    '10050
    If (Index >= EDICALLLETTERS) And (Index < ADVENDOREXTERNALIDINDEX) Then
        'Convert to upper case
        If KeyAscii >= 97 And KeyAscii <= 122 Then
            KeyAscii = KeyAscii - 32
        End If
    End If
    If Index = Signon Then
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

Private Sub edcGen_LostFocus(Index As Integer)
    Dim slStr As String
    If Index = Signon Then
        slStr = edcGen(Index).Text
        If Not gValidTime(slStr) Then
            Beep
            'edcGSignOn.SetFocus
            edcGen(Signon).SetFocus
        End If
        slStr = gFormatTime(slStr, "A", "1")
        edcGen(Index).Text = slStr
    End If
End Sub

Private Sub edcGP_Change(Index As Integer)
    If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "N") Or (tmVef.sType = "R") Or (tmVef.sType = "G")) And ((Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS) Then
        imGreatPlainsAltered = False
        If StrComp(Trim$(edcGP(0).Text), Trim$(tmVaf.sDivisionCode), vbTextCompare) <> 0 Then
            imGreatPlainsAltered = True
        End If
        If StrComp(Trim$(edcGP(1).Text), Trim$(tmVaf.sBranchCodeCash), vbTextCompare) <> 0 Then
            imGreatPlainsAltered = True
        End If
        If StrComp(Trim$(edcGP(2).Text), Trim$(tmVaf.sPCGrossSalesCash), vbTextCompare) <> 0 Then
            imGreatPlainsAltered = True
        End If
        If StrComp(Trim$(edcGP(3).Text), Trim$(tmVaf.sPCAgyCommCash), vbTextCompare) <> 0 Then
            imGreatPlainsAltered = True
        End If
        If StrComp(Trim$(edcGP(4).Text), Trim$(tmVaf.sPCRecvCash), vbTextCompare) <> 0 Then
            imGreatPlainsAltered = True
        End If
        If StrComp(Trim$(edcGP(5).Text), Trim$(tmVaf.sBranchCodeTrade), vbTextCompare) <> 0 Then
            imGreatPlainsAltered = True
        End If
        If StrComp(Trim$(edcGP(6).Text), Trim$(tmVaf.sPCGrossSalesTrade), vbTextCompare) <> 0 Then
            imGreatPlainsAltered = True
        End If
        If StrComp(Trim$(edcGP(7).Text), Trim$(tmVaf.sPCRecvTrade), vbTextCompare) <> 0 Then
            imGreatPlainsAltered = True
        End If
        '8-25-10 added
        If StrComp(Trim$(edcGP(8).Text), Trim$(tmVaf.sVendorID), vbTextCompare) <> 0 Then
            imGreatPlainsAltered = True
        End If
    End If
    mSetCommands
End Sub

Private Sub edcGP_GotFocus(Index As Integer)
    gCtrlGotFocus edcGP(Index)
End Sub

Private Sub edcGSignOn_Change()
    'mSetCommands
End Sub
Private Sub edcGSignOn_GotFocus()
    'Dim slStr As String
    'edcGTZDropDown_LostFocus
    'pbcGTZToggle.Visible = False
    'lbcFeed.Visible = False
    'cmcGTZDropDown.Visible = False
    'edcGTZDropDown.Visible = False
    'imGTZBoxNo = -1
    'slStr = edcGSignOn.Text
    'slStr = gUnformatTime(slStr)
    'edcGSignOn.Text = slStr
    'gCtrlGotFocus edcGSignOn
End Sub
Private Sub edcGSignOn_KeyPress(KeyAscii As Integer)
    'Dim ilFound As Integer
    'Dim ilLoop As Integer
    'If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
    '    ilFound = False
    '    For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
    '        If KeyAscii = igLegalTime(ilLoop) Then
    '            ilFound = True
    '            Exit For
    '        End If
    '    Next ilLoop
    '    If Not ilFound Then
    '        Beep
    '        KeyAscii = 0
    '        Exit Sub
    '    End If
    'End If
End Sub
Private Sub edcGSignOn_LostFocus()
    'Dim slStr As String
    'slStr = edcGSignOn.Text
    'If Not gValidTime(slStr) Then
    '    Beep
    '    edcGSignOn.SetFocus
    'End If
    'slStr = gFormatTime(slStr, "A", "1")
    'edcGSignOn.Text = slStr
End Sub
Private Sub edcGTZDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer

    If (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDINDEX Then  'Feed
        imLbcArrowSetting = True
        ilRet = gOptionalLookAhead(edcGTZDropDown, lbcFeed, imBSMode, slStr)
        If ilRet = 1 Then
            lbcFeed.ListIndex = 1
        End If
        imLbcArrowSetting = False
    End If
    mSetCommands
End Sub
Private Sub edcGTZDropDown_DblClick()
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub edcGTZDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcGTZDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcGTZDropDown_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    If (imGTZBoxNo Mod imTZMaxCtrls) = GNAMEINDEX Then  'Zone Name
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GLOCALADJINDEX Then  'Local Adj Time
        If (KeyAscii = KEYNEG) And (edcGTZDropDown.SelStart <> 0) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        If (KeyAscii <> KEYBACKSPACE) And (KeyAscii <> KEYNEG) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcGTZDropDown.Text
        slStr = Left$(slStr, edcGTZDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcGTZDropDown.SelStart - edcGTZDropDown.SelLength)
        If gCompNumberStr(slStr, "12") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDADJINDEX Then  'Feed Adj Time
        If (KeyAscii = KEYPOS) And (edcGTZDropDown.SelStart <> 0) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        If (KeyAscii = KEYNEG) And (edcGTZDropDown.SelStart <> 0) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        If (KeyAscii <> KEYBACKSPACE) And (KeyAscii <> KEYPOS) And (KeyAscii <> KEYNEG) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcGTZDropDown.Text
        slStr = Left$(slStr, edcGTZDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcGTZDropDown.SelStart - edcGTZDropDown.SelLength)
        If gCompNumberStr(slStr, "9") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    ElseIf ((imGTZBoxNo Mod imTZMaxCtrls) >= GVERDISPLINDEX) And ((imGTZBoxNo Mod imTZMaxCtrls) <= GVERDISPLINDEX + 3) Then  'Versions displacement
        If (KeyAscii <> KEYBACKSPACE) And (KeyAscii <> KEYPOS) And (KeyAscii <> KEYNEG) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcGTZDropDown.Text
        slStr = Left$(slStr, edcGTZDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcGTZDropDown.SelStart - edcGTZDropDown.SelLength)
        If gCompNumberStr(slStr, "99") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GCMMLSCHINDEX Then  'Version for cmml schd
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEDZONEINDEX Then  'Fed (events transmitted)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDINDEX Then  'Feed
        If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
            If edcGTZDropDown.SelLength <> 0 Then    'avoid deleting two characters
                imBSMode = True 'Force deletion of character prior to selected text
            End If
        End If
    Else    'Bus and schedule
    End If
End Sub
Private Sub edcGTZDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        If (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDINDEX Then  'Feed
            gProcessArrowKey Shift, KeyCode, lbcFeed, imLbcArrowSetting
            edcGTZDropDown.SelStart = 0
            edcGTZDropDown.SelLength = Len(edcGTZDropDown.Text)
        End If
    End If
End Sub
Private Sub edcGTZDropDown_LostFocus()
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    If imGTZBoxNo <= 0 Then
        Exit Sub
    End If
    ilIndex = ((imGTZBoxNo - 1) \ imTZMaxCtrls) + 1
    If (imGTZBoxNo Mod imTZMaxCtrls) = GNAMEINDEX Then
        tmVpf.sGZone(ilIndex - 1) = edcGTZDropDown.Text
        If Trim$(tmVpf.sGZone(ilIndex - 1)) = "" Then
            tmVpf.iGLocalAdj(ilIndex - 1) = 0
            tmVpf.iGFeedAdj(ilIndex - 1) = 0
            tmVpf.iGV1Z(ilIndex - 1) = 0
            tmVpf.iGV2Z(ilIndex - 1) = 0
            tmVpf.iGV3Z(ilIndex - 1) = 0
            tmVpf.iGV4Z(ilIndex - 1) = 0
            tmVpf.sGCSVer(ilIndex - 1) = ""
            tmVpf.sGFed(ilIndex - 1) = ""
            tmVpf.iGMnfNCode(ilIndex - 1) = 0 'Val(edcGTZDropDown.Text)
            tmVpf.sGBus(ilIndex - 1) = ""
            tmVpf.sGSked(ilIndex - 1) = ""
            tmVff.sFedDelivery(ilIndex - 1) = ""
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GLOCALADJINDEX Then  'Local time adj
        If tmVpf.sGZone(ilIndex - 1) <> "" Then
            tmVpf.iGLocalAdj(ilIndex - 1) = Val(edcGTZDropDown.Text)
        Else
            tmVpf.iGLocalAdj(ilIndex - 1) = 0
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDADJINDEX Then  'Feed time adj
        If (tmVpf.sGZone(ilIndex - 1) <> "") And (edcGTZDropDown.Text <> "") Then
            tmVpf.iGFeedAdj(ilIndex - 1) = Val(edcGTZDropDown.Text)
        Else
            tmVpf.iGFeedAdj(ilIndex - 1) = 0
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GVERDISPLINDEX Then  'Versions displacement
        If (tmVpf.sGZone(ilIndex - 1) <> "") And (edcGTZDropDown.Text <> "") Then
            tmVpf.iGV1Z(ilIndex - 1) = Val(edcGTZDropDown.Text)
        Else
            tmVpf.iGV1Z(ilIndex - 1) = 0
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GVERDISPLINDEX + 1 Then 'Versions displacement
        If (tmVpf.sGZone(ilIndex - 1) <> "") And (edcGTZDropDown.Text <> "") Then
            tmVpf.iGV2Z(ilIndex - 1) = Val(edcGTZDropDown.Text)
        Else
            tmVpf.iGV2Z(ilIndex - 1) = 0
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GVERDISPLINDEX + 2 Then 'Versions displacement
        If (tmVpf.sGZone(ilIndex - 1) <> "") And (edcGTZDropDown.Text <> "") Then
            tmVpf.iGV3Z(ilIndex - 1) = Val(edcGTZDropDown.Text)
        Else
            tmVpf.iGV3Z(ilIndex - 1) = 0
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GVERDISPLINDEX + 3 Then 'Versions displacement
        If (tmVpf.sGZone(ilIndex - 1) <> "") And (edcGTZDropDown.Text <> "") Then
            tmVpf.iGV4Z(ilIndex - 1) = Val(edcGTZDropDown.Text)
        Else
            tmVpf.iGV4Z(ilIndex - 1) = 0
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GCMMLSCHINDEX Then  'Version for Cmml Schd
'        If (tmVpf.sGZone(ilIndex) <> "") And (edcGTZDropDown.Text <> "") Then
'            tmVpf.sGCSVer(ilIndex) = Left$(edcGTZDropDown.Text, 1)
'        Else
'            tmVpf.sGCSVer(ilIndex) = ""
'        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEDZONEINDEX Then  'Fed over Dish (Yes or No)
        tmVpf.sGFed(ilIndex - 1) = edcGTZDropDown.Text
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDINDEX Then  'Feed
        If (tmVpf.sGZone(ilIndex - 1) <> "") And (edcGTZDropDown.Text <> "") Then
            ilFound = False
            If lbcFeed.ListIndex > 1 Then
                slNameCode = tmFeedCode(lbcFeed.ListIndex - 2).sKey    'lbcFeedCode.List(lbcFeed.ListIndex - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet = CP_MSG_NONE Then
                    ilFound = True
                Else
                    ilFound = False
                End If
            End If
            If ilFound Then
                tmVpf.iGMnfNCode(ilIndex - 1) = Val(slCode)
            Else
                tmVpf.iGMnfNCode(ilIndex - 1) = 0 'Val(edcGTZDropDown.Text)
            End If
        Else
            tmVpf.iGMnfNCode(ilIndex - 1) = 0
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GBUSINDEX Then  'Bus
        If (tmVpf.sGZone(ilIndex - 1) <> "") And (edcGTZDropDown.Text <> "") Then
            tmVpf.sGBus(ilIndex - 1) = edcGTZDropDown.Text
        Else
            tmVpf.sGBus(ilIndex - 1) = ""
        End If
    Else    'Schedule
        If (tmVpf.sGZone(ilIndex - 1) <> "") And (edcGTZDropDown.Text <> "") Then
            tmVpf.sGSked(ilIndex - 1) = edcGTZDropDown.Text
        Else
            tmVpf.sGSked(ilIndex - 1) = ""
        End If
    End If
    mSetCommands
End Sub
Private Sub edcGTZDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        If (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDINDEX Then  'Feed
                If imTabDirection = -1 Then  'Right To Left
                    pbcGTZSTab.SetFocus
                Else
                    pbcGTZTab.SetFocus
                End If
                Exit Sub
        End If
        imDoubleClickName = False
    End If
End Sub
Private Sub edcIFCST_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcIFCST_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcIFCST_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
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
End Sub
Private Sub edcIFDPNo_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcIFDPNo_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcIFDPNo_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < Asc("1")) Or (KeyAscii > Asc("5"))) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcIFEST_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcIFEST_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcIFEST_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
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
End Sub
Private Sub edcIFGroupNo_Change()
    mSetCommands
End Sub
Private Sub edcIFGroupNo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcIFGroupNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub edcIFMST_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcIFMST_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcIFMST_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
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
End Sub
Private Sub edcIFProgCode_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcIFProgCode_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcIFPST_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcIFPST_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcIFPST_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
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
End Sub
Private Sub edcIFZone_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcIFZone_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcIFZone_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub edcInsertComment_Change()
    imVbfChg = True
    mSetCommands
End Sub

Private Sub edcInsertComment_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcInterfaceID_Change(Index As Integer)
    If (smXDXMLForm = "H#B#P#" Or smXDXMLForm = "H#B#") And (Index = 0) Then
        imVffChg = True
    End If
    mSetCommands
End Sub

Private Sub edcInterfaceID_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcInterfaceID_KeyPress(Index As Integer, KeyAscii As Integer)
    'If smXDXMLForm = "ISCI" Then
    If Index = 1 Then
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcInventory_Change()
    imInventoryAltered = True
    mSetCommands
End Sub

Private Sub edcInventory_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcLAffDate_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcLAffDate_GotFocus(Index As Integer)
    mLSetShow imLBoxNo
    imLBoxNo = -1
    imLRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcLAffDate_KeyPress(Index As Integer, KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLDate_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcLDate_GotFocus(Index As Integer)
    mLSetShow imLBoxNo
    imLBoxNo = -1
    imLRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcLDate_KeyPress(Index As Integer, KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLDropDown_Change()
    If imIgnoreChg Then
        Exit Sub
    End If
    imLogAltered = True
    mSetCommands
End Sub
Private Sub edcLDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcLDropDown_KeyPress(KeyAscii As Integer)
    If imLBoxNo = LNODAYSINDEX Then
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcLevelPrice_Change()
    mSetCommands
End Sub

Private Sub edcLevelPrice_GotFocus()
    gCtrlGotFocus edcLevelPrice
End Sub

Private Sub edcLevelPrice_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcLLDCpyAsgn_Change()
    mSetCommands
End Sub
Private Sub edcLLDCpyAsgn_GotFocus()
    mLSetShow imLBoxNo
    imLBoxNo = -1
    imLRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcLLDCpyAsgn_KeyPress(KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLNote_Change()
    If imIgnoreChg Then
        Exit Sub
    End If
    imLogAltered = True
    mSetCommands
End Sub
Private Sub edcLNote_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcLog_Change(Index As Integer)
    If Index = 3 Then
        imVffChg = True
    End If
    mSetCommands
End Sub

Private Sub edcLog_GotFocus(Index As Integer)
    mLSetShow imLBoxNo
    imLBoxNo = -1
    imLRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcMPromo_Change()
    mSetCommands
End Sub
Private Sub edcMPromo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcMPromo_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcMPromo.Text
    slStr = Left$(slStr, edcMPromo.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcMPromo.SelStart - edcMPromo.SelLength)
    If gCompNumberStr(slStr, "99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcMPromo_LostFocus()
    If (imMPromoBoxNo >= 1) And (imMPromoBoxNo <= 24) Then
        tmVpf.iMMFPromo(imMPromoBoxNo - LBONE) = Val(edcMPromo.Text)
    ElseIf (imMPromoBoxNo >= 25) And (imMPromoBoxNo <= 48) Then
        tmVpf.iMSAPromo(imMPromoBoxNo - LBONE - 24) = Val(edcMPromo.Text)
    ElseIf (imMPromoBoxNo >= 49) And (imMPromoBoxNo - LBONE <= 72) Then
        tmVpf.iMSUPromo(imMPromoBoxNo - LBONE - 48) = Val(edcMPromo.Text)
    End If
    edcMPromo.Visible = False
End Sub
Private Sub edcMPSA_Change()
    mSetCommands
End Sub
Private Sub edcMPSA_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcMPSA_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcMPSA.Text
    slStr = Left$(slStr, edcMPSA.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcMPSA.SelStart - edcMPSA.SelLength)
    If gCompNumberStr(slStr, "99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcMPSA_LostFocus()
    If (imMPSABoxNo >= 1) And (imMPSABoxNo <= 24) Then
        tmVpf.iMMFPSA(imMPSABoxNo - LBONE) = Val(edcMPSA.Text)
    ElseIf (imMPSABoxNo >= 25) And (imMPSABoxNo <= 48) Then
        tmVpf.iMSAPSA(imMPSABoxNo - LBONE - 24) = Val(edcMPSA.Text)
    ElseIf (imMPSABoxNo >= 49) And (imMPSABoxNo <= 72) Then
        tmVpf.iMSUPSA(imMPSABoxNo - LBONE - 48) = Val(edcMPSA.Text)
    End If
    edcMPSA.Visible = False
End Sub


Private Sub edcPDropdown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    If imIgnoreChg Then
        Exit Sub
    End If
    imProducerAltered = True
    Select Case imPBoxNo
        Case PRODUCERINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcPDropdown, lbcProducer, imBSMode, slStr)
            If ilRet = 1 Then
                lbcProducer.ListIndex = -1
            End If
'        Case CONTENTPROVIDERINDEX
'            imLbcArrowSetting = True
'            ilRet = gOptionalLookAhead(edcPDropdown, lbcContentProvider, imBSMode, slStr)
'            If ilRet = 1 Then
'                lbcContentProvider.ListIndex = -1
'            End If
        Case EXPPROGAUDIOINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcPDropdown, lbcExpProgAudio, imBSMode, slStr)
            If ilRet = 1 Then
                lbcExpProgAudio.ListIndex = -1
            End If
        Case EXPCOMMAUDIOINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcPDropdown, lbcExpCommAudio, imBSMode, slStr)
            If ilRet = 1 Then
                lbcExpCommAudio.ListIndex = -1
            End If
    End Select
    mSetCommands
End Sub

Private Sub edcPDropdown_DblClick()
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub

Private Sub edcPDropdown_GotFocus()
    Select Case imPBoxNo
        Case PRODUCERINDEX
'        Case CONTENTPROVIDERINDEX
        Case EXPPROGAUDIOINDEX
        Case EXPCOMMAUDIOINDEX
    End Select
End Sub

Private Sub edcPDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcPDropdown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcPDropdown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcPDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imPBoxNo
            Case PRODUCERINDEX
                gProcessArrowKey Shift, KeyCode, lbcProducer, imLbcArrowSetting
'            Case CONTENTPROVIDERINDEX
'                gProcessArrowKey Shift, KeyCode, lbcContentProvider, imLbcArrowSetting
            Case EXPPROGAUDIOINDEX
                gProcessArrowKey Shift, KeyCode, lbcExpProgAudio, imLbcArrowSetting
            Case EXPCOMMAUDIOINDEX
                gProcessArrowKey Shift, KeyCode, lbcExpCommAudio, imLbcArrowSetting
        End Select
        edcPDropdown.SelStart = 0
        edcPDropdown.SelLength = Len(edcPDropdown.Text)
    End If
End Sub

Private Sub edcPDropdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imPBoxNo
            'Case PRODUCERINDEX, CONTENTPROVIDERINDEX, EXPPROGAUDIOINDEX, EXPCOMMAUDIOINDEX
            Case PRODUCERINDEX, EXPPROGAUDIOINDEX, EXPCOMMAUDIOINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcPSTab.SetFocus
                Else
                    pbcPTab.SetFocus
                End If
                Exit Sub
        End Select
        imDoubleClickName = False
    End If
End Sub

Private Sub edcProdPct_Change()
    imPifChg = True
End Sub

Private Sub edcProdPct_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcProdPct_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcProdPct.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcProdPct.Text, ".")    'Disallow multi-decimal points
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
    slStr = edcProdPct.Text
    slStr = Left$(slStr, edcProdPct.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcProdPct.SelStart - edcProdPct.SelLength)
    If gCompNumberStr(slStr, "100.00") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcRadarCode_Change()
    mSetCommands
End Sub

Private Sub edcRadarCode_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSAGroupNo_Change()
    'Dim ilRes As Integer

    'If imAskedSAGroupNo Then
    '    If (tmVef.sType = "S") Or (tmVef.sType = "A") Then
    '        If imOrigSAGroupNo <> Val(edcSAGroupNo.Text) Then
    '            imAskedSAGroupNo = False
    '            ilRes = MsgBox("Before changing the Group #, the links should be terminated.  Have you terminated the Links?", vbYesNo + vbQuestion, "Update")
    '            If ilRes = vbNo Then
    '                edcSAGroupNo.Text = Trim$(str$(imOrigSAGroupNo))
    '            End If
    '        End If
    '    End If
    'End If
    'mSetCommands
End Sub
Private Sub edcSAGroupNo_GotFocus()
    'edcGTZDropDown_LostFocus
    'pbcGTZToggle.Visible = False
    'lbcFeed.Visible = False
    'cmcGTZDropDown.Visible = False
    'edcGTZDropDown.Visible = False
    'imGTZBoxNo = -1
    'gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSchedule_Change(Index As Integer)
    If (Index = 4) Or (Index = 5) Then
        imVffChg = True
    End If
    mSetCommands
End Sub

Private Sub edcSchedule_GotFocus(Index As Integer)
    If Index <> 5 Then
        mSSetShow imSBoxNo
        imSBoxNo = -1
        imSRowNo = -1
    End If
    gCtrlGotFocus edcSchedule(Index)
End Sub

Private Sub edcSchedule_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Index = 2) Or (Index = 3) Then
        Exit Sub
    End If
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcSCompSepLen_Change()
    mSetCommands
End Sub
Private Sub edcSCompSepLen_GotFocus()
    gCtrlGotFocus edcSCompSepLen
End Sub
Private Sub edcSCompSepLen_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
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
End Sub
Private Sub edcSCompSepLen_LostFocus()
    Dim slStr As String
    slStr = edcSCompSepLen.Text
    If Not gValidLength(slStr) Then
        Beep
        edcSCompSepLen.SetFocus
    End If
End Sub



Private Sub edcSec_Change()
    imInventoryAltered = True
    mSetCommands
End Sub

Private Sub edcSec_GotFocus()
   gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSLen_Change()
    mSetCommands
End Sub
Private Sub edcSLen_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcSLen_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcSSDropDown_Change()
    Dim ilRet As Integer
    Dim slStr As String
    imLbcArrowSetting = True
    ilRet = gOptionalLookAhead(edcSSDropDown, lbcSSource, imBSMode, slStr)
    If ilRet = 1 Then
        lbcSSource.ListIndex = 0
    End If
    imLbcArrowSetting = False
    imPifChg = True
End Sub

Private Sub edcSSDropDown_DblClick()
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub

Private Sub edcSSDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcSSDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcSSDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcSSDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcSSource, imLbcArrowSetting
        edcSSDropDown.SelStart = 0
        edcSSDropDown.SelLength = Len(edcSSDropDown.Text)
    End If
End Sub

Private Sub edcSSDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        If imTabDirection = -1 Then  'Right To Left
            pbcParticipantSTab.SetFocus
        Else
            pbcParticipantTab.SetFocus
        End If
        Exit Sub
        imDoubleClickName = False
    End If
End Sub

Private Sub edcSSpotLG_Change()
    mSetCommands
End Sub
Private Sub edcSSpotLG_GotFocus()
    gCtrlGotFocus edcSSpotLG
End Sub
Private Sub edcSSpotLG_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcSSpotLG.Text
    slStr = Left$(slStr, edcSSpotLG.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSSpotLG.SelStart - edcSSpotLG.SelLength)
    If gCompNumberStr(slStr, "4000") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcSSpotLG_LostFocus()
    If imSSpotLenBoxNo >= 0 Then
        If imSSpotLenBoxNo <= UBound(tmVpf.iSLen) Then
            tmVpf.iSLen(imSSpotLenBoxNo) = Val(edcSSpotLG.Text)
            mSetAcqLen
            mSetAcqIndexLen
        Else
            tmVpf.iSLenGroup(imSSpotLenBoxNo - UBound(tmVpf.iSLen) - 1) = Val(edcSSpotLG.Text)
        End If
    End If
    edcSSpotLG.Visible = False
End Sub
Private Sub edcStnFdCode_Change()
    mSetCommands
End Sub
Private Sub edcStnFdCode_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcStnFdCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub edcTextFt_Change()
    imVffChg = True
    mSetCommands
End Sub

Private Sub edcTextHd_Change()
    imVffChg = True
    mSetCommands
End Sub

Private Sub edcVehGpDropDown_Change()
    Dim ilRet As Integer
    Dim slStr As String
    imLbcArrowSetting = True
    ilRet = gOptionalLookAhead(edcVehGpDropDown, lbcVehGp, imBSMode, slStr)
    If ilRet = 1 Then
        lbcVehGp.ListIndex = 0
    End If
    imLbcArrowSetting = False
    imPifChg = True
End Sub

Private Sub edcVehGpDropDown_DblClick()
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub

Private Sub edcVehGpDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcVehGpDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcVehGpDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcVehGpDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcVehGp, imLbcArrowSetting
        edcVehGpDropDown.SelStart = 0
        edcVehGpDropDown.SelLength = Len(edcVehGpDropDown.Text)
    End If
End Sub

Private Sub edcVehGpDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        If imTabDirection = -1 Then  'Right To Left
            pbcParticipantSTab.SetFocus
        Else
            pbcParticipantTab.SetFocus
        End If
        Exit Sub
        imDoubleClickName = False
    End If
End Sub

Private Sub edcXDISCIPrefix_Change(Index As Integer)
    imVffChg = True
    mSetCommands
End Sub

Private Sub edcXDISCIPrefix_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcYear_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilLoop                        slMinutes                 *
'*  slSeconds                                                                             *
'******************************************************************************************

Dim ilLen As Integer

    ilLen = Len(edcYear)        '7-21-05 determine if year entered for network inventory
    If ilLen = 4 Then           'has a 4-digit year been entered
        If imYear <> Val(edcYear.Text) Then
            imYear = Val(edcYear.Text)
            imInventoryAltered = True
            mObtainNIFYear
        End If
    End If
End Sub

Private Sub edcYear_GotFocus()
    gCtrlGotFocus ActiveControl
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
    If (igWinStatus(VEHICLESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        plcAccounting.Enabled = False
        plcGeneral.Enabled = False
        plcLog.Enabled = False
        plcPSAPromo.Enabled = False
        plcSales.Enabled = False
        plcVirtual.Enabled = False
        udcVehOptTabs.Enabled = False
        imUpdateAllowed = False
    Else
        plcAccounting.Enabled = True
        plcGeneral.Enabled = True
        plcLog.Enabled = True
        plcPSAPromo.Enabled = True
        plcSales.Enabled = True
        plcVirtual.Enabled = True
        udcVehOptTabs.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    VehOpt.Refresh
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
        If (cbcSelect.Enabled) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If ilReSet Then
            cbcSelect.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    If (igWinStatus(VEHICLESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        igPasswordOk = False
    Else
        igPasswordOk = True
    End If
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    Erase tmVehNamesCode
    Erase tmFeedCode
    Erase tmProducerCode
    Erase tmContentProviderCode
    Erase tmProducerOrProviderCode
    Erase tmVehicle
    Erase tmVbf
    Erase tmPifRec
    Erase smVbfComment
    Erase tmSSourceCode
    Erase smSUpdateRvf
    Erase tmVehGpCode
    Erase tmSMnf
    Erase tmSeasonInfo
    Erase tmMediaCode

    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    btrExtClear hmSaf   'Clear any previous extend operation
    ilRet = btrClose(hmSaf)
    btrDestroy hmSaf
    btrExtClear hmVaf   'Clear any previous extend operation
    ilRet = btrClose(hmVaf)
    btrDestroy hmVaf
    btrExtClear hmVbf   'Clear any previous extend operation
    ilRet = btrClose(hmVbf)
    btrDestroy hmVbf
    btrExtClear hmVof   'Clear any previous extend operation
    ilRet = btrClose(hmVof)
    btrDestroy hmVof
    btrExtClear hmVff   'Clear any previous extend operation
    ilRet = btrClose(hmVff)
    btrDestroy hmVff
    btrExtClear hmVtf   'Clear any previous extend operation
    ilRet = btrClose(hmVtf)
    btrDestroy hmVtf
    btrExtClear hmVsf   'Clear any previous extend operation
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    btrExtClear hmCef   'Clear any previous extend operation
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    btrExtClear hmRnf   'Clear any previous extend operation
    ilRet = btrClose(hmRnf)
    btrDestroy hmRnf
    btrExtClear hmArf   'Clear any previous extend operation
    ilRet = btrClose(hmArf)
    btrDestroy hmArf
    btrExtClear hmGhf   'Clear any previous extend operation
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    btrExtClear hmPif   'Clear any previous extend operation
    ilRet = btrClose(hmPif)
    btrDestroy hmPif
    ilRet = btrClose(hmNif)      '7-21-05
    btrDestroy hmNif

    Set VehOpt = Nothing   'Remove data segment

End Sub

Private Sub grdParticipant_EnterCell()
    mPartSetShow
End Sub

Private Sub grdParticipant_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmPartTopRow = grdParticipant.TopRow
    grdParticipant.Redraw = False
End Sub

Private Sub grdParticipant_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

    Dim ilCol As Integer
    Dim ilRow As Integer

    imIgnoreScroll = False
    If Y < grdParticipant.RowHeight(0) Then
        grdParticipant.Redraw = True
        Exit Sub
    End If
    If imCbcParticipantListIndex = -1 Then
        cbcParticipant.SetFocus
        Exit Sub
    End If
    'If Trim$(edcParticipantDate.Text) = "" Then
    '    If edcParticipantDate.Enabled Then
    '        edcParticipantDate.SetFocus
    '    End If
    '    Exit Sub
    'End If
    If Trim$(csiParticipantDate.Text) = "" Then
        If csiParticipantDate.GetEnabled Then
            csiParticipantDate.SetFocus
        End If
        Exit Sub
    End If
    pbcArrow.Visible = False
    ilCol = grdParticipant.MouseCol
    ilRow = grdParticipant.MouseRow
    If ilCol < grdParticipant.FixedCols Then
        grdParticipant.Redraw = True
        Exit Sub
    End If
    If ilRow < grdParticipant.FixedRows Then
        grdParticipant.Redraw = True
        Exit Sub
    End If
    If Not mPartColOk(ilRow, ilCol) Then
        grdParticipant.Redraw = True
        Exit Sub
    End If
    If grdParticipant.TextMatrix(ilRow, SSOURCEINDEX) = "" Then
        grdParticipant.Redraw = False
        Do
            ilRow = ilRow - 1
        Loop While grdParticipant.TextMatrix(ilRow, SSOURCEINDEX) = ""
        grdParticipant.Row = ilRow + 1
        grdParticipant.Col = SSOURCEINDEX
        grdParticipant.Redraw = True
    Else
        grdParticipant.Row = ilRow
        grdParticipant.Col = ilCol
    End If
    grdParticipant.Redraw = True
    lmPartTopRow = grdParticipant.TopRow
    'If Not mPartColOk() Then
    '    pbcArrow.Move grdParticipant.Left - pbcArrow.Width - 30, grdParticipant.Top + grdParticipant.RowPos(grdParticipant.Row) + (grdParticipant.RowHeight(grdParticipant.Row) - pbcArrow.Height) / 2
    '    pbcArrow.Visible = True
    '    Exit Sub
    'End If

    mPartEnableBox
End Sub

Private Sub grdParticipant_Scroll()
    If imIgnoreScroll Then  'Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdParticipant.Redraw = False Then
        grdParticipant.Redraw = True
        If lmPartTopRow < grdParticipant.FixedRows Then
            grdParticipant.TopRow = grdParticipant.FixedRows
        Else
            grdParticipant.TopRow = lmPartTopRow
        End If
        grdParticipant.Refresh
        grdParticipant.Redraw = False
    End If
    If (imCtrlVisible) And (grdParticipant.Row >= grdParticipant.FixedRows) And (grdParticipant.Col >= 0) And (grdParticipant.Col < grdParticipant.Cols - 1) Then
        If grdParticipant.RowIsVisible(grdParticipant.Row) Then
            pbcArrow.Move grdParticipant.Left - pbcArrow.Width - 30, grdParticipant.Top + grdParticipant.RowPos(grdParticipant.Row) + (grdParticipant.RowHeight(grdParticipant.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            mPartSetFocus
        Else
            pbcPartSetFocus.SetFocus
            edcSSDropDown.Visible = False  'Set visibility
            lbcSSource.Visible = False
            edcVehGpDropDown.Visible = False  'Set visibility
            lbcVehGp.Visible = False
            pbcIntUpdateRvf.Visible = False  'Set visibility
            pbcExtUpdateRvf.Visible = False  'Set visibility
            edcProdPct.Visible = False  'Set visibility
            cmcDropDown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcContentProvider_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcContentProvider, edcPDropdown, imChgMode, imLbcArrowSetting
    End If

End Sub

Private Sub lbcContentProvider_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcContentProvider_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcContentProvider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcContentProvider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcContentProvider, edcPDropdown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcPSTab.SetFocus
        Else
            pbcPTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcExpCommAudio_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcExpCommAudio, edcPDropdown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcExpCommAudio_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcExpCommAudio_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcExpCommAudio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcExpCommAudio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcExpCommAudio, edcPDropdown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcPSTab.SetFocus
        Else
            pbcPTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcExpProgAudio_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcExpProgAudio, edcPDropdown, imChgMode, imLbcArrowSetting
    End If

End Sub

Private Sub lbcExpProgAudio_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcExpProgAudio_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcExpProgAudio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcExpProgAudio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcExpProgAudio, edcPDropdown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcPSTab.SetFocus
        Else
            pbcPTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcFeed_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcFeed, edcGTZDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcFeed_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcFeed_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcFeed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcFeed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcFeed, edcGTZDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcGTZSTab.SetFocus
        Else
            pbcGTZTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcProducer_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcProducer, edcPDropdown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcProducer_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcProducer_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcProducer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcProducer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcProducer, edcPDropdown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcPSTab.SetFocus
        Else
            pbcPTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcSeason_Click()
    imVffChg = True
    mSetCommands
End Sub

Private Sub lbcSSource_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcSSource, edcSSDropDown, imSSChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcSSource_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcSSource_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcSSource_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcSSource_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcSSource, edcSSDropDown, imSSChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcParticipantSTab.SetFocus
        Else
            pbcParticipantTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcVehGp_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcVehGp, edcVehGpDropDown, imVehGpChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcVehGp_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcVehGp_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcVehGp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcVehGp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcVehGp, edcVehGpDropDown, imVehGpChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcParticipantSTab.SetFocus
        Else
            pbcParticipantTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcVehNames_DblClick()
    If imOrigNoVehicles > 0 Then
        Exit Sub
    End If
    If imUpdateAllowed Then
        imVirtDoubleClick = True    'Double click event is followed by a mouse up event
                            'Process the double click event in the mouse up event
                            'to avoid the mouse up event being in next form
    End If
End Sub
Private Sub lbcVehNames_DragDrop(Source As Control, X As Single, Y As Single)
    If imDragDest = -1 Then
        mVirtClearDrag
        Exit Sub
    End If
    Select Case imDragSrce
        Case DRAGVEHNAME
        Case DRAGVEHICLE
            cmcMoveToVehName_Click
    End Select
    mVirtClearDrag
End Sub
Private Sub lbcVehNames_GotFocus()
    mVirtSetShow imVirtBoxNo
    imVirtBoxNo = -1
    imVirtRowNo = -1
End Sub
Private Sub lbcVehNames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then  'Right Mouse
        Exit Sub
    End If
    If imOrigNoVehicles > 0 Then
        Exit Sub
    End If
    If imUpdateAllowed Then
        fmDragX = X
        fmDragY = Y
        imDragButton = Button
        imDragType = 0
        imDragShift = Shift
        imDragSrce = DRAGVEHNAME
        imDragIndexDest = -1
        tmcDrag.Enabled = True  'Start timer to see if drag or click
    End If
End Sub
Private Sub lbcVehNames_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Exit Sub
    End If
    If imOrigNoVehicles > 0 Then
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If imVirtDoubleClick Then
        imVirtDoubleClick = False
        cmcMoveToVehicle_Click
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFeedBranch                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to invoice*
'*                      sorting and process            *
'*                      communication back from        *
'*                      invoice sorting                *
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
Private Function mFeedBranch() As Integer
'
'   ilRet = mFeedBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcGTZDropDown, lbcFeed, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcGTZDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mFeedBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(FEEDTYPESLIST)) Then
    '    imDoubleClickName = False
    '    mFeedBranch = True
    '    mSetGTZ
    '    Exit Function
    'End If
    'MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "N"
    igMNmCallSource = CALLSOURCEVEHOPT
    If edcGTZDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'Vehicle!edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (igShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Vehicle^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Vehicle^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'VehOpt.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'VehOpt.Enabled = True
    'Vehicle!edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mFeedBranch = True
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
        lbcFeed.Clear
        mFeedPop
        If imTerminate Then
            mFeedBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcFeed
        sgMNmName = ""
        If gLastFound(lbcFeed) > 0 Then
            imChgMode = True
            lbcFeed.ListIndex = gLastFound(lbcFeed)
            edcGTZDropDown.Text = lbcFeed.List(lbcFeed.ListIndex)
            edcGTZDropDown_LostFocus    'Value set into record (no mSetShow)
            imChgMode = False
            mFeedBranch = False
        Else
            imChgMode = True
            lbcFeed.ListIndex = 1
            edcGTZDropDown.Text = lbcFeed.List(1)
            imChgMode = False
            edcGTZDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mSetGTZ
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mSetGTZ
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFeedPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Feed list             *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mFeedPop()
'
'   mFeedPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffset(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcFeed.ListIndex
    If ilIndex > 1 Then
        slName = lbcFeed.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "N"
    ilOffset(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(VehOpt, lbcFeed, lbcFeedCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(VehOpt, lbcFeed, tmFeedCode(), smFeedCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffset())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mFeedPopErr
        gCPErrorMsg ilRet, "mFeedPop (gIMoveListBox)", VehOpt
        On Error GoTo 0
        lbcFeed.AddItem "[None]", 0  'Force as first item on list
        lbcFeed.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcFeed
            If gLastFound(lbcFeed) > 1 Then
                lbcFeed.ListIndex = gLastFound(lbcFeed)
            Else
                lbcFeed.ListIndex = -1
            End If
        Else
            lbcFeed.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mFeedPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGTZPaint                       *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint system area              *
'*                                                     *
'*******************************************************
Private Sub mGTZPaint(tlVpf As VPF, tlVff As VFF)
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilMaxCol As Integer
    flY = imTZY
    '3/1/13: Allow table table to be fill in for selling vehicles
    'If (tmVef.sType = "S") Then
    '    ilMaxCol = 1
    'Else
        ilMaxCol = imTZMaxCtrls
    'End If
    For ilCol = 1 To ilMaxCol Step 1
        flX = tmTZCtrls(ilCol).fBoxX + fgBoxInsetX
        For ilRow = 1 To 5 Step 1
            gPaintArea pbcGTZ, tmTZCtrls(ilCol).fBoxX, flY + 15, tmTZCtrls(ilCol).fBoxW - fgBoxInsetX - 15, imTZH - 45, WHITE
            pbcGTZ.CurrentX = tmTZCtrls(ilCol).fBoxX + fgBoxInsetX
            pbcGTZ.CurrentY = flY - 15
            If Trim$(tlVpf.sGZone(ilRow - 1)) <> "" Then
                If ilCol = GNAMEINDEX Then
                    pbcGTZ.Print Trim$(tlVpf.sGZone(ilRow - 1))
                ElseIf ilCol = GFEDZONEINDEX Then   'Version for commercial schedule
                    pbcGTZ.Print tlVpf.sGFed(ilRow - 1)
                ElseIf ilCol = GLOCALADJINDEX Then   'Local time adj
                    pbcGTZ.Print Trim$(Str$(tlVpf.iGLocalAdj(ilRow - 1)))
                ElseIf ilCol = GFEEDADJINDEX Then   'Feed Adj
                    pbcGTZ.Print Trim$(Str$(tlVpf.iGFeedAdj(ilRow - 1)))
                ElseIf ilCol = GVERDISPLINDEX Then   'Versions displacement
                    pbcGTZ.Print Trim$(Str$(tlVpf.iGV1Z(ilRow - 1)))
                ElseIf ilCol = GVERDISPLINDEX + 1 Then 'Versions displacement
                    pbcGTZ.Print Trim$(Str$(tlVpf.iGV2Z(ilRow - 1)))
                ElseIf ilCol = GVERDISPLINDEX + 2 Then 'Versions displacement
                    pbcGTZ.Print Trim$(Str$(tlVpf.iGV3Z(ilRow - 1)))
                ElseIf ilCol = GVERDISPLINDEX + 3 Then 'Versions displacement
                    pbcGTZ.Print Trim$(Str$(tlVpf.iGV4Z(ilRow - 1)))
                ElseIf ilCol = GCMMLSCHINDEX Then   'Version for commercial schedule
                    If tlVpf.sGCSVer(ilRow - 1) = "A" Then
                        pbcGTZ.Print "All"
                    ElseIf tlVpf.sGCSVer(ilRow - 1) = "O" Then
                        pbcGTZ.Print "Original"
                   End If
                ElseIf ilCol = GFEDDELIVERYINDEX Then   'Version for commercial schedule
                    If tlVff.sFedDelivery(ilRow - 1) = "Y" Then
                        pbcGTZ.Print "Yes"
                    ElseIf tlVff.sFedDelivery(ilRow - 1) = "N" Then
                        pbcGTZ.Print "No"
                    End If
                ElseIf ilCol = GFEEDINDEX Then   'Feed
                    ilFound = False
                    For ilLoop = 0 To UBound(tmFeedCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
                        slNameCode = tmFeedCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If ilRet = CP_MSG_NONE Then
                            If Val(slCode) = tmVpf.iGMnfNCode(ilRow - 1) Then
                                ilFound = True
                                Exit For
                            End If
                        Else
                            ilFound = False
                        End If
                    Next ilLoop
                    If ilFound Then
                        slStr = lbcFeed.List(ilLoop + 2)
                        gSetShow pbcGTZ, slStr, tmTZCtrls(ilCol)
                        pbcGTZ.Print tmTZCtrls(ilCol).sShow
                    End If
                ElseIf ilCol = GBUSINDEX Then   'Version for commercial schedule
                    pbcGTZ.Print Trim$(tlVpf.sGBus(ilRow - 1))
                Else
                    pbcGTZ.Print Trim$(tlVpf.sGSked(ilRow - 1))
                End If
            End If
            flY = flY + imTZH - 15
        Next ilRow
        flY = imTZY
    Next ilCol
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
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
'*  llAdj                                                                                 *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilValue As Integer
    Dim llStdDate As Long
    Dim llCalDate As Long
    Dim slStr As String

    Screen.MousePointer = vbHourglass
    imLBPCtrls = 1
    imLBLCtrls = 1
    imLBSCtrls = 1
    imLBVirtCtrls = 1
    imLBTZCtrls = 1
    imVpfChanged = False
    bmInDateChg = False
    imTerminate = False
    imFirstActivate = True
    imCtrlVisible = False
    tbcSelection.Move 75, cbcSelect.Top + cbcSelect.Height + 30
    plcGeneral.BorderStyle = 0
    plcGeneral.Visible = True
    'plcGeneral.Move VehOpt.Width / 2 - plcGeneral.Width / 2 - 60, 660   '555
    plcGeneral.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcSales.BorderStyle = 0
    plcSales.Visible = False
    'plcSales.Move VehOpt.Width / 2 - plcSales.Width / 2 - 60, 660   '555
    plcSales.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcSchedule(0).BorderStyle = 0
    plcSchedule(0).Visible = False
    plcSchedule(0).Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcSchedule(1).BorderStyle = 0
    plcSchedule(1).Visible = False
    plcSchedule(1).Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcSports.BorderStyle = 0
    plcSports.Visible = False
    plcSports.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '460
    plcPSAPromo.BorderStyle = 0
    plcPSAPromo.Visible = False
    'plcPSAPromo.Move VehOpt.Width / 2 - plcPSAPromo.Width / 2 - 60, 660 '555
    plcPSAPromo.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '460
    plcLog.BorderStyle = 0
    plcLog.Visible = False
    'plcLog.Move VehOpt.Width / 2 - plcLog.Width / 2 - 60, 660   '555
    plcLog.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcProducer.BorderStyle = 0
    plcProducer.Visible = False
    'plcAccounting.Move VehOpt.Width / 2 - plcAccounting.Width / 2 - 60, 660 '555
    plcProducer.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    '1/17/10:  Move embedded to Affiliate Export ISCI screen
    If tgSpf.sGUseAffSys = "Y" Then
        pbcProducer.Height = 1065   '1425
    Else
        pbcProducer.Height = 375
    End If
    plcAccounting.BorderStyle = 0
    plcAccounting.Visible = False
    'plcAccounting.Move VehOpt.Width / 2 - plcAccounting.Width / 2 - 60, 660 '555
    plcAccounting.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    'No tab for virtual as this feature has been removed
    plcVirtual.BorderStyle = 0
    plcVirtual.Visible = False
    'plcVirtual.Move VehOpt.Width / 2 - plcVirtual.Width / 2 - 60, 660 '555
    plcVirtual.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    plcGreatPlains.BorderStyle = 0
    plcGreatPlains.Visible = False
    plcGreatPlains.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    If ((Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) <> GREATPLAINS) Then
        plcGreatPlains.Enabled = False
    End If
    
    udcVehOptTabs.Visible = False
    udcVehOptTabs.Move tbcSelection.Left + 60, tbcSelection.Top + 355, tbcSelection.Width - 120, tbcSelection.Height - 510  'plcInv.Height
    
    frcBarterEnable(1).Left = frcBarterEnable(0).Left
    frcBarterEnable(1).Top = frcBarterEnable(0).Top
    plcBarter.BorderStyle = 0
    plcBarter.Visible = False
    'plcLog.Move VehOpt.Width / 2 - plcLog.Width / 2 - 60, 660   '555
    plcBarter.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
   
    plcExport.BorderStyle = 0
    plcExport.Visible = False
    plcExport.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    
    plcParticipant.BorderStyle = 0
    plcParticipant.Visible = False
    plcParticipant.Move tbcSelection.Left + 60, tbcSelection.Top + 355  '475
    imPartMissAndReq = False
    If tgSpf.sSystemType = "R" Then
        'frcRemoteExport.Visible = False
        pbcGTZ.Enabled = False
        pbcGTZ.Visible = False
        pbcGTZSTab.Enabled = False
        pbcGTZSTab.Visible = False
        pbcGTZTab.Enabled = False
        pbcGTZTab.Visible = False
        lacFed.Visible = False
        lacTitle(0).Visible = False
        lacTitle(1).Visible = False
        'llAdj = plcGSAGroupNo.Top - frcRemoteExport.Top - 60
        'plcGSAGroupNo.Top = plcGSAGroupNo.Top - llAdj
        'edcSAGroupNo.Top = edcSAGroupNo.Top - llAdj
        'plcShowAirTime.Top = plcShowAirTime.Top - llAdj
        'lacGEDICallLetter.Top = lacGEDICallLetter.Top - llAdj
        'edcGEDICallLetter.Top = edcGEDICallLetter.Top - llAdj
        'lacGAGL.Top = lacGAGL.Top - llAdj
        'edcGen(0).Top = edcGen(0).Top - llAdj
        'edcGen(1).Top = edcGen(1).Top - llAdj
        'lacGBGL.Top = lacGBGL.Top - llAdj
        'edcGen(2).Top = edcGen(2).Top - llAdj
        'edcGen(3).Top = edcGen(3).Top - llAdj
    End If
    ilValue = Asc(tgSpf.sUsingFeatures)  'Option Fields in Orders/Proposals
    If (ilValue And USINGLIVELOG) <> USINGLIVELOG Then 'Using Live Log
        rbcGenLog(2).Enabled = False
    Else
        rbcGenLog(2).Enabled = True
    End If
    If (Asc(tgSpf.sUsingFeatures5) And REMOTEEXPORT) <> REMOTEEXPORT Then
        frcRemoteExport.Visible = False
        rbcRemoteExport(2).Value = True
    Else
        frcRemoteExport.Visible = True
    End If
    If (Asc(tgSpf.sUsingFeatures5) And REMOTEIMPORT) <> REMOTEIMPORT Then
        frcRemoteImport.Visible = False
        rbcRemoteImport(2).Value = True
    Else
        frcRemoteImport.Visible = True
    End If
    If tgSpf.sGUseAffSys <> "Y" Then
        rbcRemoteImport(1).Enabled = False
    End If
    pbcXDXMLForm.Enabled = False
    edcInterfaceID(0).Enabled = False
    edcInterfaceID(1).Enabled = False
    edcXDISCIPrefix(0).Enabled = False
    edcXDISCIPrefix(1).Enabled = False
    ckcXDSave(0).Enabled = False
    ckcXDSave(1).Enabled = False
    ckcXDSave(2).Enabled = False
    ckcXDSave(3).Enabled = False
    ckcXDSave(4).Visible = False
    ckcXDSave(5).Visible = False
    ckcXDSave(6).Visible = False
    '10933
    ckcXDSave(CUEZONE).Enabled = False
    If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) Or ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
        If ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
            pbcXDXMLForm.Enabled = True
            edcInterfaceID(0).Enabled = True
            edcXDISCIPrefix(0).Enabled = True
            ckcXDSave(0).Enabled = True
            ckcXDSave(1).Enabled = True
            ckcXDSave(2).Enabled = True
            ckcXDSave(3).Enabled = True
            '10933
            ckcXDSave(CUEZONE).Enabled = True
            smHBHBPAvailForm = "X-Digital: Avail Form                                Program ID                        ISCI Prefix                        Delivery Save-"
            lacCode(3).Caption = smHBHBPAvailForm
        End If
        If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) Then
            edcInterfaceID(1).Enabled = True
            edcXDISCIPrefix(1).Enabled = True
            ckcXDSave(4).Visible = True
            ckcXDSave(5).Visible = True
            ckcXDSave(6).Visible = True
            smISCIAvailForm = "X-Digital: National ISCI Model                   Vehicle ID                        ISCI Prefix                        Delivery Save-"
            lacCode(4).Caption = smISCIAvailForm
        End If
    Else
        If ((Asc(tgSpf.sUsingFeatures8) And ISCIEXPORT) = ISCIEXPORT) Then
            pbcXDXMLForm.Visible = False
            lacCode(3).Caption = "ISCI Export: Avail Form                              Program ID"
            edcInterfaceID(0).Enabled = True
            edcXDISCIPrefix(0).Visible = False
            ckcXDSave(0).Visible = False
            ckcXDSave(1).Visible = False
            ckcXDSave(2).Visible = False
            lacCode(4).Enabled = False
            edcInterfaceID(1).Enabled = False
            edcXDISCIPrefix(1).Enabled = False
            ckcXDSave(3).Enabled = False
            ckcXDSave(4).Enabled = False
            ckcXDSave(5).Enabled = False
            ckcXDSave(6).Enabled = False
        End If
    End If
    mInitParameter
    'smISCIAvailForm = "X-Digital: National ISCI Model                   Vehicle ID                        ISCI Prefix"
    'lacCode(4).Caption = smISCIAvailForm
    'smHBHBPAvailForm = "X-Digital: Avail Form                                Program ID                        ISCI Prefix                         Delivery Save-"
    'lacCode(3).Caption = smHBHBPAvailForm
    imTabDirection = 0  'Left to right movement
    imVBFIndex = -1
    imVbfChg = False
    imPifChg = False
    imVffChg = False
    imVefChg = False
    imFirstTime = True
    imGTZBoxNo = -1
    imSSpotLenBoxNo = -1
    imLevelPriceBoxNo = -1
    imAcqCostBoxNo = -1
    imAcqIndexBoxNo = -1
    imMPSABoxNo = -1
    imMPromoBoxNo = -1
    imVirtBoxNo = -1
    imVirtRowNo = -1
    imVirtSettingValue = False
    imVirtChgVeh = False
    imDragIndexSrce = -1
    imDragSrce = -1
    lmPartTopRow = -1
    imIgnoreScroll = False
    imCbcParticipantListIndex = -1
    lmPartEnableRow = -1
    lmPartEnableCol = -1
    imIgnorePartChg = False
    imVirtDoubleClick = False
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", VehOpt
    On Error GoTo 0
    imVefRecLen = Len(tmVef)

    hmVof = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVof, "", sgDBPath & "Vof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vof.Btr)", VehOpt
    On Error GoTo 0
    imVofRecLen = Len(tmVof)


    hmVbf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVbf, "", sgDBPath & "Vbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vbf.Btr)", VehOpt
    On Error GoTo 0
    ReDim tmVbf(0 To 0) As VBF
    imVBfRecLen = Len(tmVbf(0))


    hmSaf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Saf.Btr)", VehOpt
    On Error GoTo 0
    imSafRecLen = Len(tmSaf)

    hmVaf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVaf, "", sgDBPath & "Vaf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vaf.Btr)", VehOpt
    On Error GoTo 0
    imVafRecLen = Len(tmVaf)

    hmVff = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVff, "", sgDBPath & "Vff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vff.Btr)", VehOpt
    On Error GoTo 0
    imVffRecLen = Len(tmVff)

    hmVtf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVtf, "", sgDBPath & "Vtf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vtf.Btr)", VehOpt
    On Error GoTo 0
    imVtfRecLen = Len(tmVtf)

    hmVsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", VehOpt
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmCef = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cef.Btr)", VehOpt
    On Error GoTo 0
    imCefRecLen = Len(tmCef)
    hmRnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRnf, "", sgDBPath & "Rnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rnf.Btr)", VehOpt
    On Error GoTo 0
    imRnfRecLen = Len(tmRnf)
    hmArf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmArf, "", sgDBPath & "Arf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Arf.Btr)", VehOpt
    On Error GoTo 0
    imArfRecLen = Len(tmArf)

    hmGhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ghf.Btr)", VehOpt
    On Error GoTo 0
    imGhfRecLen = Len(tmGhf)


    hmPif = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPif, "", sgDBPath & "Pif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pif.Btr)", VehOpt
    On Error GoTo 0
    ReDim tmPifRec(0 To 0) As PIFREC
    imPifRecLen = Len(tmPifRec(0).tPif)

    'Open NIF Network Inventory 7-21-05
    hmNif = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmNif, "", sgDBPath & "Nif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Nif.Btr)", VehOpt
    On Error GoTo 0
    imNifRecLen = Len(tmNif)

    mGetAffiliateSite
    
    mInitBox
    cbcSelect.Clear  'Force list to be populated
    mVirtVehPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ilRet = gVffRead()
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    smSMnfStamp = ""
    ilRet = gObtainMnfForType("S", smSMnfStamp, tmSMnf())
    If Not imTerminate Then
        lbcSSource.Clear
        mSSourcePop
    End If
    If Not imTerminate Then
        lbcVehGp.Clear
        mVehGpPop
    End If
    VehOpt.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    'gCenterModalForm VehOpt
    gCenterStdAlone VehOpt
    imIgnoreChg = True
    mSetCommands
    imIgnoreChg = False
    If tgSpf.sSystemType <> "R" Then
        'If tgSpf.sGUseAffSys <> "Y" Then
        '    lacFed.Visible = False
        'Else
            lacFed.Visible = True
        'End If
    Else
        lacFed.Visible = False
    End If
    If ((Asc(tgSpf.sUsingFeatures7) And WEGENEREXPORT) <> WEGENEREXPORT) Then
        ckcAffExport(0).Enabled = False
    End If
     If ((Asc(tgSpf.sAutoType3) And ENCOESPN) <> ENCOESPN) Then
        ckcIFExport(14).Enabled = False
    End If
    If ((Asc(tgSpf.sUsingFeatures) And MATRIXEXPORT) <> MATRIXEXPORT) And ((Asc(tgSaf(0).sFeatures1) And MATRIXCAL) <> MATRIXCAL) Then
        ckcIFExport(16).Enabled = False
    End If
    If ((Asc(tgSaf(0).sFeatures1) And JELLIEXPORT) <> JELLIEXPORT) Then      'yes, show audio type on lines
        ckcIFExport(17).Enabled = False
    End If
    If ((Asc(tgSaf(0).sFeatures2) And TABLEAUEXPORT) <> TABLEAUEXPORT) And ((Asc(tgSaf(0).sFeatures2) And TABLEAUCAL) <> TABLEAUCAL) Then      'yes, show audio type on lines
        ckcIFExport(18).Enabled = False
    End If
    If (Asc(tgSpf.sUsingFeatures10) And WegenerIPump) <> WegenerIPump Then
        ckcAffExport(3).Enabled = False
        edcExport(1).Enabled = False
        lacCode(2).Enabled = False
    End If
    If ((Asc(tgSaf(0).sFeatures6) And RABCALENDAR) <> RABCALENDAR) And ((Asc(tgSaf(0).sFeatures7) And RABSTD) <> RABSTD) And ((Asc(tgSaf(0).sFeatures7) And RABCALSPOTS) <> RABCALSPOTS) Then
        ckcIFExport(19).Enabled = False
    End If
    'TTP 9992
    If ((Asc(tgSaf(0).sFeatures7) And CUSTOMEXPORT) <> CUSTOMEXPORT) Then
        ckcIFExport(20).Enabled = False
    End If

    
    '4/14/19
    If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then  'Using Live Copy
        rbcAudioType(0).Enabled = True
        rbcAudioType(1).Enabled = True
        rbcAudioType(2).Enabled = True
        rbcAudioType(3).Enabled = True
        rbcAudioType(4).Enabled = True
        rbcAudioType(5).Enabled = True
        
        If tgSaf(0).sExcludeAudioTypeL = "Y" Then
            rbcAudioType(1).Enabled = False
        End If
        If tgSaf(0).sExcludeAudioTypeM = "Y" Then
            rbcAudioType(2).Enabled = False
        End If
        If tgSaf(0).sExcludeAudioTypeS = "Y" Then
            rbcAudioType(3).Enabled = False
        End If
        If tgSaf(0).sExcludeAudioTypeP = "Y" Then
            rbcAudioType(4).Enabled = False
        End If
        If tgSaf(0).sExcludeAudioTypeQ = "Y" Then
            rbcAudioType(5).Enabled = False
        End If
        If tgSaf(0).sExcludeAudioTypeR = "Y" Then
            rbcAudioType(0).Enabled = False
        Else
            rbcAudioType(0).Enabled = True
        End If
    Else
        rbcAudioType(0).Enabled = True
        rbcAudioType(1).Enabled = False
        rbcAudioType(2).Enabled = False
        rbcAudioType(3).Enabled = False
        rbcAudioType(4).Enabled = False
        rbcAudioType(5).Enabled = False
    End If
   'If ((Asc(tgSpf.sUsingFeatures7) And OLAEXPORT) <> OLAEXPORT) Then
    '    ckcAffExport(1).Enabled = False
    'End If
    lbcFeed.Clear
    mFeedPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lbcProducer.Clear
    mProducerPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    mMediaPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'lbcContentProvider.Clear
    lbcExpProgAudio.Clear
    lbcExpCommAudio.Clear
    mContentProviderPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
'    mProducerOrProviderPop
'    If imTerminate Then
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    mRnfPop
    If tgSpf.sGUseAffSys <> "Y" Then
        lacLAffLogDate(0).Visible = False
        edcLAffDate(0).Visible = False
        edcLAffDate(1).Visible = False
        
        udcVehOptTabs.Enabled = False
        
        'pbcLogForm.Visible = False
        'pbcLSTab.Visible = False
        'pbcLTab.Visible = False
    End If
    If igVpfType = 1 Then
        'Moved to tab control
        'rbcOption(2).Enabled = False
        'rbcOption(3).Enabled = False
        'rbcOption(4).Enabled = False

        pbcLogForm.Visible = False
        pbcLSTab.Visible = False
        pbcLTab.Visible = False
    End If
    gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llStdDate
    gUnpackDateLong tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), llCalDate
    If llStdDate < llCalDate Then
        lmLastBilledDate = llCalDate
    Else
        lmLastBilledDate = llStdDate
    End If
    
    bmCitationDefined = False
    SQLQuery = "Select * From cxf_Hdr_Ln_Comments Where cxfCode  = " & tgSaf(0).lCitationComment
    'Set cmt_rst = cnn.Execute(SQLQuery)
    Set cmt_rst = gSQLSelectCall(SQLQuery)
    If Not cmt_rst.EOF Then
        slStr = gStripChr0(cmt_rst!cxfComment)
        If slStr <> "" Then
            bmCitationDefined = True
        End If
    End If
    '10050 podcast
    cbcCsiGeneric(ADVENDOR).BackColor = &HFFFF00
    '10981
'    smAdVendorVehNameOriginal = ""
'    imCurrentVendorInfoIndex = -1
    edcGen(ADVENDOREXTERNALIDINDEX).MaxLength = 70
    mVendorsLoadAndSelect True, 0
    '10894 remove
   ' mInitGeneralMedium
    imInInit = False
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    '8032 turn on/off here
    rbcBarterMethod(STATIONXMLMARKETRON).Visible = False
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
    Dim flTextHeight As Single  'Standard text height
    Dim ilRow As Integer

    'Time zone info
    gSetCtrl tmTZCtrls(GNAMEINDEX), 465, 360, 600, fgBoxGridH
    tmTZCtrls(GNAMEINDEX).iReq = False
    gSetCtrl tmTZCtrls(GFEDZONEINDEX), 1080, tmTZCtrls(GNAMEINDEX).fBoxY, 375, fgBoxGridH
    tmTZCtrls(GFEDZONEINDEX).iReq = False
    gSetCtrl tmTZCtrls(GLOCALADJINDEX), 1470, tmTZCtrls(GNAMEINDEX).fBoxY, 795, fgBoxGridH
    tmTZCtrls(GLOCALADJINDEX).iReq = False
    gSetCtrl tmTZCtrls(GFEEDADJINDEX), 2280, tmTZCtrls(GNAMEINDEX).fBoxY, 795, fgBoxGridH
    tmTZCtrls(GFEEDADJINDEX).iReq = False
    gSetCtrl tmTZCtrls(GVERDISPLINDEX), 3090, tmTZCtrls(GNAMEINDEX).fBoxY, 405, fgBoxGridH
    tmTZCtrls(GVERDISPLINDEX).iReq = False
    gSetCtrl tmTZCtrls(GVERDISPLINDEX + 1), 3510, tmTZCtrls(GNAMEINDEX).fBoxY, 405, fgBoxGridH
    tmTZCtrls(GVERDISPLINDEX + 1).iReq = False
    gSetCtrl tmTZCtrls(GVERDISPLINDEX + 2), 3930, tmTZCtrls(GNAMEINDEX).fBoxY, 405, fgBoxGridH
    tmTZCtrls(GVERDISPLINDEX + 2).iReq = False
    gSetCtrl tmTZCtrls(GVERDISPLINDEX + 3), 4350, tmTZCtrls(GNAMEINDEX).fBoxY, 405, fgBoxGridH
    tmTZCtrls(GVERDISPLINDEX + 3).iReq = False
    gSetCtrl tmTZCtrls(GCMMLSCHINDEX), 4770, tmTZCtrls(GNAMEINDEX).fBoxY, 870, fgBoxGridH
    tmTZCtrls(GCMMLSCHINDEX).iReq = False
    gSetCtrl tmTZCtrls(GFEDDELIVERYINDEX), 5655, tmTZCtrls(GNAMEINDEX).fBoxY, 495, fgBoxGridH
    tmTZCtrls(GFEDDELIVERYINDEX).iReq = False
    gSetCtrl tmTZCtrls(GFEEDINDEX), 6165, tmTZCtrls(GNAMEINDEX).fBoxY, 1485, fgBoxGridH
    tmTZCtrls(GFEEDINDEX).iReq = False
    gSetCtrl tmTZCtrls(GBUSINDEX), 7665, tmTZCtrls(GNAMEINDEX).fBoxY, 375, fgBoxGridH
    tmTZCtrls(GBUSINDEX).iReq = False
    gSetCtrl tmTZCtrls(GSCHDINDEX), 8055, tmTZCtrls(GNAMEINDEX).fBoxY, 375, fgBoxGridH
    tmTZCtrls(GSCHDINDEX).iReq = False
    'Log Form
    'No Days
    edcLNote.Move pbcLogForm.Left, pbcLogForm.Top - edcLNote.Height - 90
    gSetCtrl tmLCtrls(LNODAYSINDEX), 465, 375, 690, fgBoxGridH
    tmLCtrls(LNODAYSINDEX).iReq = False
    'Skip Page
    gSetCtrl tmLCtrls(LSKIPINDEX), 1170, tmLCtrls(LNODAYSINDEX).fBoxY, 690, fgBoxGridH
    tmLCtrls(LSKIPINDEX).iReq = False
    'Length
    gSetCtrl tmLCtrls(LLENINDEX), 1875, tmLCtrls(LNODAYSINDEX).fBoxY, 465, fgBoxGridH
    tmLCtrls(LLENINDEX).iReq = False
    'Product
    gSetCtrl tmLCtrls(LPRODINDEX), 2355, tmLCtrls(LNODAYSINDEX).fBoxY, 555, fgBoxGridH
    tmLCtrls(LPRODINDEX).iReq = False
    'Creative Title
    gSetCtrl tmLCtrls(LTITLEINDEX), 2925, tmLCtrls(LNODAYSINDEX).fBoxY, 420, fgBoxGridH
    tmLCtrls(LTITLEINDEX).iReq = False
    'ISCI
    gSetCtrl tmLCtrls(LISCIINDEX), 3360, tmLCtrls(LNODAYSINDEX).fBoxY, 420, fgBoxGridH
    tmLCtrls(LISCIINDEX).iReq = False
    'Daypart
    gSetCtrl tmLCtrls(LDAYPARTINDEX), 3795, tmLCtrls(LNODAYSINDEX).fBoxY, 555, fgBoxGridH
    tmLCtrls(LDAYPARTINDEX).iReq = False
    'Time
    gSetCtrl tmLCtrls(LTIMEINDEX), 4365, tmLCtrls(LNODAYSINDEX).fBoxY, 690, fgBoxGridH
    tmLCtrls(LTIMEINDEX).iReq = False
    'Time Line
    gSetCtrl tmLCtrls(LLINEINDEX), 5070, tmLCtrls(LNODAYSINDEX).fBoxY, 690, fgBoxGridH
    tmLCtrls(LLINEINDEX).iReq = False
    'Hour
    gSetCtrl tmLCtrls(LHOURINDEX), 5775, tmLCtrls(LNODAYSINDEX).fBoxY, 420, fgBoxGridH
    tmLCtrls(LHOURINDEX).iReq = False
    'Load
    gSetCtrl tmLCtrls(LLOADINDEX), 6210, tmLCtrls(LNODAYSINDEX).fBoxY, 510, fgBoxGridH
    tmLCtrls(LLOADINDEX).iReq = False
    'Header
    gSetCtrl tmLCtrls(LHEADERINDEX), 6735, tmLCtrls(LNODAYSINDEX).fBoxY, 510, fgBoxGridH
    tmLCtrls(LHEADERINDEX).iReq = False
    'Note 1
    gSetCtrl tmLCtrls(LFOOT1INDEX), 7260, tmLCtrls(LNODAYSINDEX).fBoxY, 510, fgBoxGridH
    tmLCtrls(LFOOT1INDEX).iReq = False
    'Note 2
    gSetCtrl tmLCtrls(LFOOT2INDEX), 7785, tmLCtrls(LNODAYSINDEX).fBoxY, 510, fgBoxGridH
    tmLCtrls(LFOOT2INDEX).iReq = False

    'Producer
    gSetCtrl tmPCtrls(PRODUCERINDEX), 30, 30, 4560, fgBoxStH
    tmPCtrls(PRODUCERINDEX).iReq = False
'    gSetCtrl tmPCtrls(CONTENTPROVIDERINDEX), 30, 375, 4560, fgBoxStH
'    tmPCtrls(CONTENTPROVIDERINDEX).iReq = False
'    gSetCtrl tmPCtrls(EXPPROGAUDIOINDEX), 30, 720, 4560, fgBoxStH
'    tmPCtrls(EXPPROGAUDIOINDEX).iReq = False
'    gSetCtrl tmPCtrls(EXPCOMMAUDIOINDEX), 30, 1065, 3345, fgBoxStH
'    tmPCtrls(EXPCOMMAUDIOINDEX).iReq = False
'    gSetCtrl tmPCtrls(COMMEMBEDDEDINDEX), 3390, 1065, 1200, fgBoxStH
'    tmPCtrls(COMMEMBEDDEDINDEX).iReq = False
    gSetCtrl tmPCtrls(EXPPROGAUDIOINDEX), 30, 375, 4560, fgBoxStH
    tmPCtrls(EXPPROGAUDIOINDEX).iReq = False
    gSetCtrl tmPCtrls(EXPCOMMAUDIOINDEX), 30, 720, 4560, fgBoxStH
    tmPCtrls(EXPCOMMAUDIOINDEX).iReq = False
    gSetCtrl tmPCtrls(COMMEMBEDDEDINDEX), 30, 1065, 4560, fgBoxStH
    tmPCtrls(COMMEMBEDDEDINDEX).iReq = False

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

    mGridParticipantLayout
    mGridParticipantColumnWidths
    mGridParticipantColumns
    ilRow = grdParticipant.FixedRows
    Do
        If ilRow + 1 > grdParticipant.Rows Then
            grdParticipant.AddItem ""
        End If
        grdParticipant.RowHeight(ilRow) = fgBoxGridH + 15
        ilRow = ilRow + 1
    Loop While grdParticipant.RowIsVisible(ilRow - 1)
    imInitNoRows = grdParticipant.Rows - grdParticipant.FixedRows
    mGridParticipantLayout
    gGrid_IntegralHeight grdParticipant, fgBoxGridH + 30 'fgBoxGridH + 0     '15
    grdParticipant.Height = grdParticipant.Height - 30

    'Virtual vehicle
    flTextHeight = pbcVehicle.TextHeight("1")
    'Position panel and picture areas with panel
    plcVehNames(0).Move 105, 510, lbcVehNames.Width + fgPanelAdj, lbcVehNames.Height + fgPanelAdj
    lbcVehNames.Move 60, 60
    plcVehNames(1).Move 4290, 510, pbcVehicle.Width + vbcVehicle.Width + fgPanelAdj, pbcVehicle.Height + fgPanelAdj
    pbcVehicle.Move plcVehNames(0).Left + fgBevelX, plcVehNames(0).Top + fgBevelY
    'Set either Lock box or Agency DP service picture visible
    pbcVehicle.Visible = True
    'Vehicle
    gSetCtrl tmVirtCtrls(VEHINDEX), 30, 225, 2445, fgBoxGridH
    tmVirtCtrls(VEHINDEX).iReq = False
    '# of spots
    gSetCtrl tmVirtCtrls(NOSPOTSINDEX), 2490, tmVirtCtrls(VEHINDEX).fBoxY, 825, fgBoxGridH
    '% of $'s
    gSetCtrl tmVirtCtrls(PERCENTINDEX), 3330, tmVirtCtrls(VEHINDEX).fBoxY, 825, fgBoxGridH
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
    fmTextHeight = pbcPSA.TextHeight("1")
    imTZX = 465
    imTZY = 360
    imTZH = edcGTZDropDown.Height
    imTZW = edcGTZDropDown.Width
    imPsaX = 525
    imPsaY = 210
    imPsaH = edcMPSA.Height
    imPsaW = edcMPSA.Width
    imPromoX = 525
    imPromoY = 210
    imPromoH = edcMPromo.Height
    imPromoW = edcMPromo.Width
    imSpotLGX = 30
    imSpotLGY = 15
    imSpotLGH = edcSSpotLG.Height
    imSpotLGW = edcSSpotLG.Width
    imAcqCostX = 1590
    imAcqCostY = 210
    imAcqCostH = edcBarter(3).Height
    imAcqCostW = edcBarter(3).Width
    imAcqIndexX = 30
    imAcqIndexY = 210
    imAcqIndexH = edcBarter(4).Height
    imAcqIndexW = edcBarter(4).Width
    imTZMaxCtrls = UBound(tmTZCtrls)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLEnableBox                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mLEnableBox(ilBoxNo As Integer)
'
'   mVirtEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBLCtrls) Or (ilBoxNo > UBound(tmLCtrls)) Then
        Exit Sub
    End If
    If (imLRowNo < 1) Or (imLRowNo >= 4) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case LNODAYSINDEX '# of Days
            imIgnoreChg = True
            edcLDropDown.Width = tmLCtrls(ilBoxNo).fBoxW
            edcLDropDown.MaxLength = 2
            gMoveTableCtrl pbcLogForm, edcLDropDown, tmLCtrls(ilBoxNo).fBoxX, tmLCtrls(ilBoxNo).fBoxY + (imLRowNo - 1) * (fgBoxGridH + 15)
            edcLDropDown.Text = Trim$(Str$(imLSave(ilBoxNo, imLRowNo)))
            edcLDropDown.Enabled = True
            edcLDropDown.Visible = True  'Set visibility
            edcLDropDown.SetFocus
            imIgnoreChg = False
        Case LSKIPINDEX, LLENINDEX, LPRODINDEX, LTITLEINDEX, LISCIINDEX, LDAYPARTINDEX, LTIMEINDEX, LLINEINDEX, LHOURINDEX
            pbcLYN.Width = tmLCtrls(ilBoxNo).fBoxW
            gMoveTableCtrl pbcLogForm, pbcLYN, tmLCtrls(ilBoxNo).fBoxX, tmLCtrls(ilBoxNo).fBoxY + (imLRowNo - 1) * (fgBoxGridH + 15)
            pbcLYN.Visible = True  'Set visibility
            pbcLYN.SetFocus
        Case LLOADINDEX 'Load
            imIgnoreChg = True
            edcLDropDown.Width = tmLCtrls(ilBoxNo).fBoxW
            edcLDropDown.MaxLength = 2
            gMoveTableCtrl pbcLogForm, edcLDropDown, tmLCtrls(ilBoxNo).fBoxX, tmLCtrls(ilBoxNo).fBoxY + (imLRowNo - 1) * (fgBoxGridH + 15)
            edcLDropDown.Text = Trim$(Str$(imLSave(ilBoxNo, imLRowNo)))
            edcLDropDown.Enabled = True
            edcLDropDown.Visible = True  'Set visibility
            edcLDropDown.SetFocus
            imIgnoreChg = False
        Case LHEADERINDEX, LFOOT1INDEX, LFOOT2INDEX
            imIgnoreChg = True
            edcLDropDown.Width = tmLCtrls(ilBoxNo).fBoxW
            edcLDropDown.MaxLength = 0
            gMoveTableCtrl pbcLogForm, edcLDropDown, tmLCtrls(ilBoxNo).fBoxX, tmLCtrls(ilBoxNo).fBoxY + (imLRowNo - 1) * (fgBoxGridH + 15)
            edcLDropDown.Text = Left$(smLSave(ilBoxNo - LHEADERINDEX + 1, imLRowNo), 10)
            edcLDropDown.Visible = True  'Set visibility
            edcLDropDown.Enabled = False
            edcLNote.Text = smLSave(ilBoxNo - LHEADERINDEX + 1, imLRowNo)
            edcLNote.Visible = True  'Set visibility
            edcLNote.SetFocus
            imIgnoreChg = False
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLSetFocus                      *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mLSetFocus(ilBoxNo As Integer)
'
'   mLSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBLCtrls) Or (ilBoxNo > UBound(tmLCtrls)) Then
        Exit Sub
    End If
    If (imLRowNo < 1) Or (imLRowNo >= 4) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case LNODAYSINDEX '# of Days
            edcLDropDown.Visible = True  'Set visibility
            edcLDropDown.SetFocus
        Case LSKIPINDEX, LLENINDEX, LPRODINDEX, LTITLEINDEX, LISCIINDEX, LDAYPARTINDEX, LTIMEINDEX, LLINEINDEX, LHOURINDEX
            pbcLYN.Visible = True  'Set visibility
            pbcLYN.SetFocus
        Case LLOADINDEX 'Load Factor
            edcLDropDown.Visible = True  'Set visibility
            edcLDropDown.SetFocus
        Case LHEADERINDEX, LFOOT1INDEX, LFOOT2INDEX
            edcLDropDown.Visible = True  'Set visibility
            edcLNote.Visible = True  'Set visibility
            edcLNote.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLSetShow                       *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mLSetShow(ilBoxNo As Integer)
'
'   mLSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    Dim slComment As String
    Dim slChar As String

    If (ilBoxNo < imLBLCtrls) Or (ilBoxNo > UBound(tmLCtrls)) Then
        Exit Sub
    End If
    If (imLRowNo < 1) Or (imLRowNo >= 4) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case LNODAYSINDEX
            edcLDropDown.Visible = False  'Set visibility
            slStr = edcLDropDown.Text
            imLSave(ilBoxNo, imLRowNo) = Val(slStr)
            gSetShow pbcLogForm, slStr, tmLCtrls(ilBoxNo)
            smLShow(ilBoxNo, imLRowNo) = tmLCtrls(ilBoxNo).sShow
        Case LSKIPINDEX, LLENINDEX, LPRODINDEX, LTITLEINDEX, LISCIINDEX, LDAYPARTINDEX, LTIMEINDEX, LLINEINDEX, LHOURINDEX
            pbcLYN.Visible = False
            If imLSave(ilBoxNo, imLRowNo) = 1 Then
                slStr = "Yes"
            ElseIf imLSave(ilBoxNo, imLRowNo) = 0 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcLogForm, slStr, tmLCtrls(ilBoxNo)
            smLShow(ilBoxNo, imLRowNo) = tmLCtrls(ilBoxNo).sShow
        Case LLOADINDEX
            edcLDropDown.Visible = False  'Set visibility
            slStr = edcLDropDown.Text
            imLSave(ilBoxNo, imLRowNo) = Val(slStr)
            gSetShow pbcLogForm, slStr, tmLCtrls(ilBoxNo)
            smLShow(ilBoxNo, imLRowNo) = tmLCtrls(ilBoxNo).sShow
        Case LHEADERINDEX, LFOOT1INDEX, LFOOT2INDEX
            edcLDropDown.Visible = False  'Set visibility
            edcLNote.Visible = False  'Set visibility
            slStr = edcLNote.Text
            smLSave(ilBoxNo - LHEADERINDEX + 1, imLRowNo) = slStr
            slComment = ""
            For ilLoop = 1 To Len(slStr) Step 1
                slChar = Mid$(slStr, ilLoop, 1)
                If Asc(slChar) < 32 Then
                    Exit For
                End If
                slComment = slComment & slChar
            Next ilLoop
            gSetShow pbcLogForm, slComment, tmLCtrls(ilBoxNo)
            smLShow(ilBoxNo, imLRowNo) = tmLCtrls(ilBoxNo).sShow
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                      and set defaults               *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(tlVpf As VPF)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilWeeks                       ilInvPerWeek                                            *
'******************************************************************************************

'
'   mMoveCtrlToRec
'   Where:
'
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slTime As String
    Dim ilRnf As Integer
    Dim slName As String
    Dim ilRow As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilValue As Integer

    'General
    slStr = edcGen(Signon).Text     'edcGSignOn.Text
    gPackTime slStr, tlVpf.iGTime(0), tlVpf.iGTime(1)
    'edcGSignOn.Text = slStr
    edcGen(Signon).Text = slStr
    If rbcGMedium(0).Value Then
        tlVpf.sGMedium = "R"
    ElseIf rbcGMedium(1).Value Then
        tlVpf.sGMedium = "T"
    ElseIf rbcGMedium(2).Value Then
        tlVpf.sGMedium = "N"
    ElseIf rbcGMedium(3).Value Then
        tlVpf.sGMedium = "V"
    ElseIf rbcGMedium(4).Value Then
        tlVpf.sGMedium = "C"
    ElseIf rbcGMedium(5).Value Then
        tlVpf.sGMedium = "S"
    '10050
    ElseIf rbcGMedium(PODCASTRBC).Value Then     ''1-16-14 podcast
        tlVpf.sGMedium = "P"
    End If
    If rbcGMedium(7).Value Then
        tlVpf.sEmbeddedOrROS = "E"
    ElseIf rbcGMedium(8).Value Then
        tlVpf.sEmbeddedOrROS = "R"
    End If
    ilValue = 0
    If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Then
        If (Asc(tgSpf.sUsingFeatures5) And REMOTEEXPORT) = REMOTEEXPORT Then
            If rbcRemoteExport(0).Value Then
                ilValue = ilValue Or EXPORTINSERTION
            ElseIf rbcRemoteExport(1).Value Then
                ilValue = ilValue Or EXPORTLOG
            End If
        End If
        If (Asc(tgSpf.sUsingFeatures5) And REMOTEIMPORT) = REMOTEIMPORT Then
            If rbcRemoteImport(0).Value Then
                ilValue = ilValue Or IMPORTINSERTION
            ElseIf rbcRemoteImport(1).Value Then
                ilValue = ilValue Or IMPORTAFFILIATESPOTS
            End If
        End If
    End If
    If ckcAffExport(2).Value = vbChecked Then
        ilValue = ilValue Or EXPORTISCIBYPLEDGE
    End If
    If udcVehOptTabs.WebInfo(3) = vbChecked Then
        ilValue = ilValue Or SUPPRESSWEBLOG
    End If
    If udcVehOptTabs.WebInfo(4) = vbChecked Then
        ilValue = ilValue Or EXPORTPOSTEDTIMES
    End If
    tlVpf.sUsingFeatures1 = Chr(ilValue)
    
    ilValue = 0
    If ckcXDSave(3).Value = vbChecked Then
        ilValue = ilValue Or XDSAPPLYMERGE
    End If
    If rbcShowAirDate(1).Value = True Then
        ilValue = ilValue Or INVOICEAIRDATEWO
    End If
    tlVpf.sUsingFeatures2 = Chr(ilValue)

    If rbcSMoveLLD(0).Value Then
        tlVpf.sMoveLLD = "Y"
    Else
        tlVpf.sMoveLLD = "N"
    End If
    If rbcBillSA(0).Value Then
        tlVpf.sBillSA = "Y"
    Else
        tlVpf.sBillSA = "N"
    End If

    ' tlVpf.iUrfGCode = tlFromRec.iUrfGCode     'Counterpoint
    'tlVpf.iGMoFull = Val(edcGPast.Text)      'month retain unpacked
    'tlVpf.iGMoPack = Val(edcGHistory.Text)   'months retain packed
    'If rbcSPrice(0).Value Then
    '    tlVpf.sGPriceStat = "Y"
    'Else
    '    tlVpf.sGPriceStat = "N"
    'End If
    If frcExport(3).Visible Then
        If rbcGMedium(10).Value Then
            tlVpf.sOwnership = "B"
        ElseIf rbcGMedium(11).Value Then
            tlVpf.sOwnership = "C"
        ElseIf rbcGMedium(12).Value Then
            tlVpf.sOwnership = "D"
        Else
            tlVpf.sOwnership = "A"
        End If
    Else
        tlVpf.sOwnership = "A"
    End If
    If rbcGenLog(1).Value Then
        tlVpf.sGenLog = "N"
    ElseIf rbcGenLog(2).Value Then
        tlVpf.sGenLog = "L"
    ElseIf rbcGenLog(3).Value Then
        tlVpf.sGenLog = "M"
    ElseIf rbcGenLog(4).Value Then
        tlVpf.sGenLog = "A"
    Else
        tlVpf.sGenLog = "Y"
    End If
    If rbcLGrid(0).Value Then
        tlVpf.sGGridRes = "F"
    ElseIf rbcLGrid(1).Value Then
        tlVpf.sGGridRes = "H"
    Else
        tlVpf.sGGridRes = "Q"
    End If
    'If rbcLScripts(0).Value Then
    '    tlVpf.sGScript = "Y"
    'Else
    '    tlVpf.sGScript = "N"
    'End If
    If ckcIFExport(5).Value = vbUnchecked Then
        tlVpf.sExpHiCorp = "N"
    Else
        tlVpf.sExpHiCorp = "Y"
    End If
    If rbcShowAirTime(1).Value Then
        tlVpf.sShowTime = "H"
    ElseIf rbcShowAirTime(2).Value Then
        tlVpf.sShowTime = "D"
    ElseIf rbcShowAirTime(3).Value Then
        tlVpf.sShowTime = "A"
    Else
        tlVpf.sShowTime = "S"
    End If
    If (tmVef.sType = "S") Or (tmVef.sType = "A") Then
        tlVpf.iSAGroupNo = Val(edcGen(SAGROUPNO).Text)
    Else
        tlVpf.iSAGroupNo = 0
    End If
    tlVpf.sEDICallLetters = Trim$(edcGen(EDICALLLETTERS).Text)  'Trim$(edcGEDI(0).Text)
    tlVpf.sEDIBand = Trim$(edcGen(EDIBAND).Text) 'Trim$(edcGEDI(1).Text)
    tlVpf.sAccruedRevenue = Trim$(edcGen(0).Text)
    tlVpf.sAccruedTrade = Trim$(edcGen(1).Text)
    tlVpf.sBilledRevenue = Trim$(edcGen(2).Text)
    tlVpf.sBilledTrade = Trim$(edcGen(3).Text)
    'Sales
    If rbcSCommission(0).Value Then
        tlVpf.sSVarComm = "Y"
    Else
        tlVpf.sSVarComm = "N"
    End If
    If rbcSAdvtSep(0).Value Then
        tlVpf.sAdvtSep = "T"
    ElseIf rbcSAdvtSep(1).Value Then
        tlVpf.sAdvtSep = "B"
    End If
    If rbcSCompetitive(0).Value Then
        tlVpf.sSCompType = "T"
    ElseIf rbcSCompetitive(1).Value Then
        tlVpf.sSCompType = "B"
    ElseIf rbcSCompetitive(2).Value Then
        tlVpf.sSCompType = "N"
    End If
    If tlVpf.sSCompType = "T" Then
        slStr = edcSCompSepLen.Text
        gPackLength slStr, tlVpf.iSCompLen(0), tlVpf.iSCompLen(1)
    Else
        tlVpf.iSCompLen(0) = 0
        tlVpf.iSCompLen(1) = 0
    End If
    If rbcSSellout(0).Value Then
        tlVpf.sSSellOut = "U"
    ElseIf rbcSSellout(1).Value Then
        tlVpf.sSSellOut = "B"
    ElseIf rbcSSellout(2).Value Then
        tlVpf.sSSellOut = "T"
    Else
        tlVpf.sSSellOut = "M"
    End If
    If rbcSOverbook(0).Value Then
        tlVpf.sSOverBook = "Y"
    Else
        tlVpf.sSOverBook = "N"
    End If
    If rbcSMove(0).Value Then
        tlVpf.sSForceMG = "W"   'Always
    Else
        tlVpf.sSForceMG = "A"   'Ask
    End If

    '7-21-05 update network inventory count
    If edcYear.Text <> "" Then           'was anything entered to create or update?
        'tmNIf.iYear = Val(edcYear.Text)
        If ckcRollover.Value = vbChecked Then       'ok to use unused inventory in the past
            'tmNIf.sAllowRollover = "Y"
            smRollover = "Y"
        Else
            'tmNIf.sAllowRollover = "N"
            smRollover = "N"
        End If
        If rbcInvBy(0).Value = True Then            'do inventory by week
            'tmNIf.sInvWkYear = "W"
            smByWeekOrYear = "W"
        Else
            'tmNIf.sInvWkYear = "Y"
            smByWeekOrYear = "Y"
        End If

    Else                                'no record to update
        tmNif.lCode = 0
        tmNif.iVefCode = 0
        tmNif.iYear = 0
    End If



    'If rbcSIntoRC(0).Value Then
    '    tlVpf.sSIntoRC = "Y"
    'Else
    '    tlVpf.sSIntoRC = "N"
    'End If
    'If rbcSPTA(0).Value Then
    '    tlVpf.sSPTA = "P"
    'ElseIf rbcSPTA(1).Value Then
    '    tlVpf.sSPTA = "T"
    'Else
    '    tlVpf.sSPTA = "A"
    'End If
    'If rbcSMixed(0).Value Then
    '    tlVpf.sSPlaceNet = "F"
    'ElseIf rbcSMixed(1).Value Then
    '    tlVpf.sSPlaceNet = "L"
    'Else
    '    tlVpf.sSPlaceNet = "N"
    'End If
    
    If rbcSBreak(0).Value Then
        tlVpf.sSAvailOrder = "1"
    ElseIf rbcSBreak(1).Value Then
        tlVpf.sSAvailOrder = "2"
    ElseIf rbcSBreak(2).Value Then
        tlVpf.sSAvailOrder = "3"
    ElseIf rbcSBreak(3).Value Then
        tlVpf.sSAvailOrder = "4"
    ElseIf rbcSBreak(4).Value Then
        tlVpf.sSAvailOrder = "5"
    ElseIf rbcSBreak(5).Value Then
        tlVpf.sSAvailOrder = "6"
    Else
        tlVpf.sSAvailOrder = "7"
    End If
'    pbcSSpotLen.Cls
'    pbcSSpotLen_Paint   'Spot length and group number
    If rbcSSalesperson(0).Value Then
        tlVpf.sSCommCalc = "B"
    Else
        tlVpf.sSCommCalc = "C"
    End If
    'PSA/Promo
    'Log
    slStr = edcLLDCpyAsgn.Text
    gPackDate slStr, tlVpf.iLLastDateCpyAsgn(0), tlVpf.iLLastDateCpyAsgn(1)    'Last log date
    slStr = edcLDate(0).Text
    gPackDate slStr, tlVpf.iLLD(0), tlVpf.iLLD(1)    'Last log date
    slStr = edcLDate(1).Text
    gPackDate slStr, tlVpf.iLPD(0), tlVpf.iLPD(1)    'Last preliminary date
    slStr = edcLAffDate(0).Text
    gPackDate slStr, tlVpf.iLastLog(0), tlVpf.iLastLog(1)    'Last log date
    slStr = edcLAffDate(1).Text
    gPackDate slStr, tlVpf.iLastCP(0), tlVpf.iLastCP(1)    'Last preliminary date
    If rbcLZone(0).Value Then
        tlVpf.slTimeZone = "E"
    ElseIf rbcLZone(1).Value Then
        tlVpf.slTimeZone = "C"
    ElseIf rbcLZone(2).Value Then
        tlVpf.slTimeZone = "M"
    Else
        tlVpf.slTimeZone = "P"
    End If
    If rbcLDaylight(0).Value Then
        tlVpf.sLDayLight = "Y"
    Else
        tlVpf.sLDayLight = "N"
    End If
    If rbcLTiming(0).Value Then
        tlVpf.sLTiming = "Y"
    Else
        tlVpf.sLTiming = "N"
    End If
    'If rbcLLen(0).Value Then
    '    tlVpf.sLAvailLen = "Y"
    'Else
    '    tlVpf.sLAvailLen = "N"
    'End If
    tlVpf.iSDLen = Val(edcSLen.Text)
    'If rbcLAffCPs(0).Value Then
    '    tlVpf.sAffCPs = "Y"
    'Else
    '    tlVpf.sAffCPs = "N"
    'End If
    'If rbcLAffTimes(0).Value Then
    '    tlVpf.sAffTimes = "Y"
    'Else
    '    tlVpf.sAffTimes = "N"
    'End If
    'If rbcLCut(0).Value Then
    '    tlVpf.sLShowCut = "C"
    'ElseIf rbcLCut(1).Value Then
    '    tlVpf.sLShowCut = "I"
    'ElseIf rbcLCut(2).Value Then
    '    tlVpf.sLShowCut = "B"
    'Else
    '    tlVpf.sLShowCut = "N"
    'End If
    'If rbcLTime(0).Value Then
    '    tlVpf.slTimeFormat = "A"
    'Else
    '    tlVpf.slTimeFormat = "M"
    'End If
    If tmVef.sType = "A" Then
        If rbcLCopyOnAir(0).Value Then
            tlVpf.sCopyOnAir = "Y"
        Else
            tlVpf.sCopyOnAir = "N"
        End If
    Else
        tlVpf.sCopyOnAir = "N"
    End If
    tlVpf.iRnfLogCode = 0
    If cbcLog(0).ListIndex > 0 Then
        slName = Trim$(cbcLog(0).List(cbcLog(0).ListIndex))
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If StrComp(Trim$(tgRnfList(ilRnf).tRnf.sName), slName, 1) = 0 Then
                tlVpf.iRnfLogCode = tgRnfList(ilRnf).tRnf.iCode
                Exit For
            End If
        Next ilRnf
    End If
    tlVpf.iRnfCertCode = 0
    If cbcLog(1).ListIndex > 0 Then
        slName = Trim$(cbcLog(1).List(cbcLog(1).ListIndex))
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If StrComp(Trim$(tgRnfList(ilRnf).tRnf.sName), slName, 1) = 0 Then
                tlVpf.iRnfCertCode = tgRnfList(ilRnf).tRnf.iCode
                Exit For
            End If
        Next ilRnf
    End If
    tlVpf.iRnfPlayCode = 0
    If cbcLog(2).ListIndex > 0 Then
        slName = Trim$(cbcLog(2).List(cbcLog(2).ListIndex))
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If StrComp(Trim$(tgRnfList(ilRnf).tRnf.sName), slName, 1) = 0 Then
                tlVpf.iRnfPlayCode = tgRnfList(ilRnf).tRnf.iCode
                Exit For
            End If
        Next ilRnf
    End If
    tlVpf.iRnfSvLogCode = 0
    If cbcSvLog(0).ListIndex > 0 Then
        slName = Trim$(cbcSvLog(0).List(cbcSvLog(0).ListIndex))
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If StrComp(Trim$(tgRnfList(ilRnf).tRnf.sName), slName, 1) = 0 Then
                tlVpf.iRnfSvLogCode = tgRnfList(ilRnf).tRnf.iCode
                Exit For
            End If
        Next ilRnf
    End If
    tlVpf.iRnfSvCertCode = 0
    If cbcSvLog(1).ListIndex > 0 Then
        slName = Trim$(cbcSvLog(1).List(cbcSvLog(1).ListIndex))
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If StrComp(Trim$(tgRnfList(ilRnf).tRnf.sName), slName, 1) = 0 Then
                tlVpf.iRnfSvCertCode = tgRnfList(ilRnf).tRnf.iCode
                Exit For
            End If
        Next ilRnf
    End If
    tlVpf.iRnfSvPlayCode = 0
    If cbcSvLog(2).ListIndex > 0 Then
        slName = Trim$(cbcSvLog(2).List(cbcSvLog(2).ListIndex))
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If StrComp(Trim$(tgRnfList(ilRnf).tRnf.sName), slName, 1) = 0 Then
                tlVpf.iRnfSvPlayCode = tgRnfList(ilRnf).tRnf.iCode
                Exit For
            End If
        Next ilRnf
    End If
    If ckcUnsoldBlank.Value = vbUnchecked Then
        tlVpf.sUnsoldBlank = "N"
    Else
        tlVpf.sUnsoldBlank = "Y"
    End If
    If udcVehOptTabs.WebInfo(0) = vbChecked Then
        tlVpf.sAvailNameOnWeb = "Y"
    Else
        tlVpf.sAvailNameOnWeb = "N"
    End If
    If udcVehOptTabs.WebInfo(1) = vbChecked Then
        tlVpf.sWebLogFeedTime = "Y"
    Else
        tlVpf.sWebLogFeedTime = "N"
    End If
    If udcVehOptTabs.WebInfo(2) = vbChecked Then
        tlVpf.sWebLogSummary = "Y"
    Else
        tlVpf.sWebLogSummary = "N"
    End If
    'FTP saved in cmcUpdate
    'smFTP = Trim$(edcLog(0).Text)
    smFTP = Trim$(udcVehOptTabs.FTPInfo())
    smLiveWindow = Trim$(udcVehOptTabs.LiveWindow())
    '2/28/19: Add Cart on Web
    If udcVehOptTabs.WebInfo(5) = vbChecked Then
        smCartOnWeb = "Y"
    Else
        smCartOnWeb = "N"
    End If
    
    smAutoExpt = Trim$(edcLog(1).Text)
    smAutoImpt = Trim$(edcLog(2).Text)
    'Comments set in cmcUpdate
    'Log Form
    If ((tmVef.sType = "C") And (tmVef.iVefCode = 0)) Or (tmVef.sType = "L") Or (tmVef.sType = "A") Or (tmVef.sType = "G") Then
        'If (tgSpf.sGUseAffSys = "Y") And (igVpfType <> 1) Then
        If (igVpfType <> 1) Then
            For ilRow = 1 To 3 Step 1
                If ilRow = 1 Then
                    tmVof = tmLVof
                ElseIf ilRow = 2 Then
                    tmVof = tmCVof
                Else
                    tmVof = tmOVof
                End If
                tmVof.iNoDaysCP = imLSave(LNODAYSINDEX, ilRow)
                If imLSave(LSKIPINDEX, ilRow) = 1 Then
                    tmVof.sSkipPage = "Y"
                Else
                    tmVof.sSkipPage = "N"
                End If
                If imLSave(LLENINDEX, ilRow) = 1 Then
                    tmVof.sShowLen = "Y"
                Else
                    tmVof.sShowLen = "N"
                End If
                If imLSave(LPRODINDEX, ilRow) = 1 Then
                    tmVof.sShowProduct = "Y"
                Else
                    tmVof.sShowProduct = "N"
                End If
                If imLSave(LTITLEINDEX, ilRow) = 1 Then
                    tmVof.sShowCreative = "Y"
                Else
                    tmVof.sShowCreative = "N"
                End If
                If imLSave(LISCIINDEX, ilRow) = 1 Then
                    tmVof.sShowISCI = "Y"
                Else
                    tmVof.sShowISCI = "N"
                End If
                mLSetShow LISCIINDEX
                If imLSave(LDAYPARTINDEX, ilRow) = 1 Then
                    tmVof.sShowDP = "Y"
                Else
                    tmVof.sShowDP = "N"
                End If
                If imLSave(LTIMEINDEX, ilRow) = 1 Then
                    tmVof.sShowAirTime = "Y"
                Else
                    tmVof.sShowAirTime = "N"
                End If
                If imLSave(LLINEINDEX, ilRow) = 1 Then
                    tmVof.sShowAirLine = "Y"
                Else
                    tmVof.sShowAirLine = "N"
                End If
                If imLSave(LHOURINDEX, ilRow) = 1 Then
                    tmVof.sShowHour = "Y"
                Else
                    tmVof.sShowHour = "N"
                End If
                tmVof.iLoadFactor = imLSave(LLOADINDEX, ilRow)
                If ilRow = 1 Then
                    tmLVof = tmVof
                ElseIf ilRow = 2 Then
                    tmCVof = tmVof
                Else
                    tmOVof = tmVof
                End If
            Next ilRow
        End If
    End If
    'Schedule
    If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Then
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
            tmSaf.iPreferredPct = 0
            tmSaf.iWk1stSoloIndex = 0
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
    End If
    If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Then
        ilValue = 0
        If ckcSch(4).Value = vbChecked Then
            ilValue = ilValue Or SUPPRESSPREEMPTION
        End If
        tmSaf.sFeatures1 = Chr(ilValue)
    Else
        tmSaf.sFeatures1 = Chr(0)
    End If
    'Great Plains
    If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "N") Or (tmVef.sType = "R") Or (tmVef.sType = "G")) And ((Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS) Then
        tmVaf.sDivisionCode = edcGP(0).Text
        tmVaf.sBranchCodeCash = edcGP(1).Text
        tmVaf.sPCGrossSalesCash = edcGP(2).Text
        tmVaf.sPCAgyCommCash = edcGP(3).Text
        tmVaf.sPCRecvCash = edcGP(4).Text
        tmVaf.sBranchCodeTrade = edcGP(5).Text
        tmVaf.sPCGrossSalesTrade = edcGP(6).Text
        tmVaf.sPCRecvTrade = edcGP(7).Text
        tmVaf.sVendorID = edcGP(8).Text         '8-25-10
    End If
    
    'Producer
    ilFound = False
    If lbcProducer.ListIndex > 1 Then
        slNameCode = tmProducerCode(lbcProducer.ListIndex - 2).sKey    'lbcFeedCode.List(lbcFeed.ListIndex - 2)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If ilRet = CP_MSG_NONE Then
            ilFound = True
        Else
            ilFound = False
        End If
    End If
    If ilFound Then
        tlVpf.iProducerArfCode = Val(slCode)
    Else
        tlVpf.iProducerArfCode = 0
    End If
    If tgSpf.sGUseAffSys = "Y" Then
'        ilFound = False
'        If lbcContentProvider.ListIndex > 1 Then
'            slNameCode = tmContentProviderCode(lbcContentProvider.ListIndex - 2).sKey    'lbcFeedCode.List(lbcFeed.ListIndex - 2)
'            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'            If ilRet = CP_MSG_NONE Then
'                ilFound = True
'            Else
'                ilFound = False
'            End If
'        End If
'        If ilFound Then
'            tlVpf.iProviderArfCode = Val(slCode)
'        Else
'            tlVpf.iProviderArfCode = 0
'        End If
        ilFound = False
        If lbcExpProgAudio.ListIndex > 1 Then
            slNameCode = tmContentProviderCode(lbcExpProgAudio.ListIndex - 2).sKey    'lbcFeedCode.List(lbcFeed.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If ilRet = CP_MSG_NONE Then
                ilFound = True
            Else
                ilFound = False
            End If
        End If
        If ilFound Then
            tlVpf.iProgProvArfCode = Val(slCode)
        Else
            tlVpf.iProgProvArfCode = 0
        End If
        ilFound = False
        If lbcExpCommAudio.ListIndex > 1 Then
            slNameCode = tmContentProviderCode(lbcExpCommAudio.ListIndex - 2).sKey    'lbcFeedCode.List(lbcFeed.ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If ilRet = CP_MSG_NONE Then
                ilFound = True
            Else
                ilFound = False
            End If
        End If
        If ilFound Then
            tlVpf.iCommProvArfCode = Val(slCode)
        Else
            tlVpf.iCommProvArfCode = 0
        End If
    Else
        tlVpf.iProgProvArfCode = 0
        tlVpf.iCommProvArfCode = 0
    End If

    'Interface
    For ilLoop = 1 To 5 Step 1
        slTime = Trim$(edcIFEST(ilLoop - 1).Text)
        If slTime <> "" Then
            slTime = gConvertTime(slTime)
            If slTime <> "12:00AM" Then
                tlVpf.iESTEndTime(ilLoop - 1) = Minute(slTime) + 60 * Hour(slTime)
            Else
                tlVpf.iESTEndTime(ilLoop - 1) = 24 * 60
            End If
        Else
            tlVpf.iESTEndTime(ilLoop - 1) = 0
        End If
        slTime = Trim$(edcIFCST(ilLoop - 1).Text)
        If slTime <> "" Then
            slTime = gConvertTime(slTime)
            If slTime <> "12:00AM" Then
                tlVpf.iCSTEndTime(ilLoop - 1) = Minute(slTime) + 60 * Hour(slTime)
            Else
                tlVpf.iCSTEndTime(ilLoop - 1) = 24 * 60
            End If
        Else
            tlVpf.iCSTEndTime(ilLoop - 1) = 0
        End If
        slTime = Trim$(edcIFMST(ilLoop - 1).Text)
        If slTime <> "" Then
            slTime = gConvertTime(slTime)
            If slTime <> "12:00AM" Then
                tlVpf.iMSTEndTime(ilLoop - 1) = Minute(slTime) + 60 * Hour(slTime)
            Else
                tlVpf.iMSTEndTime(ilLoop - 1) = 24 * 60
            End If
        Else
            tlVpf.iMSTEndTime(ilLoop - 1) = 0
        End If
        slTime = Trim$(edcIFPST(ilLoop - 1).Text)
        If slTime <> "" Then
            slTime = gConvertTime(slTime)
            If slTime <> "12:00AM" Then
                tlVpf.iPSTEndTime(ilLoop - 1) = Minute(slTime) + 60 * Hour(slTime)
            Else
                tlVpf.iPSTEndTime(ilLoop - 1) = 24 * 60
            End If
        Else
            tlVpf.iPSTEndTime(ilLoop - 1) = 0
        End If
    Next ilLoop
    For ilLoop = 1 To 4 Step 1
        tlVpf.sMapZone(ilLoop - 1) = Trim$(edcIFZone(ilLoop - 1).Text)
        tlVpf.sMapProgCode(ilLoop - 1) = Trim$(edcIFProgCode(ilLoop - 1).Text)
        tlVpf.iMapDPNo(ilLoop - 1) = Val(edcIFDPNo(ilLoop - 1).Text)
    Next ilLoop
    If ckcIFExport(0).Value = vbChecked Then
        tlVpf.sExpHiClear = "Y"
    Else
        tlVpf.sExpHiClear = "N"
    End If
    If ckcIFExport(1).Value = vbChecked Then
        tlVpf.sExpHiDallas = "Y"
    Else
        tlVpf.sExpHiDallas = "N"
    End If
    If ckcIFExport(2).Value = vbChecked Then
        tlVpf.sExpHiPhoenix = "Y"
    Else
        tlVpf.sExpHiPhoenix = "N"
    End If
    If ckcIFExport(3).Value = vbChecked Then
        tlVpf.sExpHiNY = "Y"
    Else
        tlVpf.sExpHiNY = "N"
    End If
    If ckcIFExport(6).Value = vbChecked Then
        tlVpf.sExpHiNYISCI = "Y"
    Else
        tlVpf.sExpHiNYISCI = "N"
    End If
    If ckcIFExport(4).Value = vbChecked Then
        tlVpf.sExpHiCmmlChg = "Y"
    Else
        tlVpf.sExpHiCmmlChg = "N"
    End If
    If rbcExpBkCpyCart(0).Value Then
        tlVpf.sExpBkCpyCart = "Y"
    Else
        tlVpf.sExpBkCpyCart = "N"
    End If
    If rbcIFBulk(0).Value Then
        tlVpf.sBulkXFer = "Y"
    Else
        tlVpf.sBulkXFer = "N"
    End If
    If rbcIFSelling(0).Value Then
        tlVpf.sClearAsSell = "Y"
    Else
        tlVpf.sClearAsSell = "N"
    End If
    If rbcIFTime(0).Value Then
        tlVpf.sClearChgTime = "Y"
    Else
        tlVpf.sClearChgTime = "N"
    End If
'    gUnpackLength tlVpf.iBCal(0), tlVpf.iBCal(1), slStr    'Last log date
'    lacBInvDate(0).Text = slStr
'    gUnpackLength tlVpf.iBStd(0), tlVpf.iBStd(1), slStr    'Last preliminary date
'    lacBInvDate(1).Text = slStr
    tlVpf.sGGroupNo = edcIFGroupNo.Text   'Bulk Feed Group Number
    tlVpf.sExpVehNo = edcExpCntrVehNo.Text   'Contract Export Vehicle Number
    tlVpf.sARBCode = edcARBCode.Text   'Arbitron Code
    tlVpf.sRadarCode = edcRadarCode.Text   'Radar Code
    tlVpf.sStnFdCode = edcStnFdCode.Text   'Station feed code
    If ckcStnFdInfo(0).Value = vbChecked Then
        tlVpf.sStnFdCart = "Y"
    Else
        tlVpf.sStnFdCart = "N"
    End If
    If ckcStnFdInfo(1).Value = vbChecked Then
        tlVpf.sStnFdXRef = "Y"
    Else
        tlVpf.sStnFdXRef = "N"
    End If
    tlVpf.lEDASWindow = Val(edcEDASWindow.Text)
    If ckcKCGenRot.Value = vbChecked Then
        tlVpf.sKCGenRot = "Y"
    Else
        tlVpf.sKCGenRot = "N"
    End If
    If ckcExportSQL.Value = vbChecked Then
        tlVpf.sExportSQL = "Y"
    Else
        tlVpf.sExportSQL = "N"
    End If
    ilValue = Asc(tgSpf.sUsingFeatures2)
    ''If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G")) And (ilValue And BARTER) = BARTER Then
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S")) And (ilValue And BARTER) = BARTER Then
    If (tmVef.sType = "R") And ((ilValue And BARTER) = BARTER) Then
'        tlVpf.lDefAcqCost = Val(edcBarter(0).Text)
'        tlVpf.lActAcqCost = Val(edcBarter(1).Text)
'        If rbcBarterMethod(0).Value = True Then
'            tlVpf.sBarterMethod = "A"
'        ElseIf rbcBarterMethod(1).Value = True Then
'            tlVpf.sBarterMethod = "M"
'            tlVpf.iBarterThreshold = Val(edcBarterMethod(0).Text)
'        ElseIf rbcBarterMethod(2).Value = True Then
'            tlVpf.sBarterMethod = "U"
'            tlVpf.iBarterThreshold = Val(edcBarterMethod(1).Text)
'        ElseIf rbcBarterMethod(3).Value = True Then
'            tlVpf.sBarterMethod = "X"
'            tlVpf.iBarterXFree = Val(edcBarterMethod(2).Text)
'            tlVpf.iBarterYSold = Val(edcBarterMethod(3).Text)
'        ElseIf rbcBarterMethod(4).Value = True Then
'            tlVpf.sBarterMethod = "N"
'            tlVpf.iBarterThreshold = 0
'            tlVpf.iBarterXFree = 0
'            tlVpf.iBarterYSold = 0
'        End If
    End If
    ilValue = Asc(tgSpf.sUsingFeatures2)
    tlVpf.sAllowSplitCopy = "N"
    '6/19/07:  Jim:  Allow defineition for Conventional, Airing and Game.  Always set on fro Packages and Selling
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or ((tmVef.sType = "A") And (rbcLCopyOnAir(0).Value)) Or (tmVef.sType = "G") Or (tmVef.sType = "P")) And ((ilValue And SPLITCOPY) = SPLITCOPY) Then
    '5/11/11: Allow Selling vehicles to be set as No
    'If ((tmVef.sType = "C") Or (tmVef.sType = "A") Or (tmVef.sType = "G")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
    If ((tmVef.sType = "C") Or (tmVef.sType = "A") Or (tmVef.sType = "G") Or (tmVef.sType = "S")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
        If rbcSplitCopy(0).Value Then
            tlVpf.sAllowSplitCopy = "Y"
        Else
            tlVpf.sAllowSplitCopy = "N"
        End If
    'ElseIf ((tmVef.sType = "S") Or (tmVef.sType = "P")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
    ElseIf ((tmVef.sType = "P")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
        tlVpf.sAllowSplitCopy = "Y"
    End If
    tlVpf.sShowRateOnInsert = "N"
    'If ((tmVef.sType = "C") And (tmVef.iVefCode = 0)) Or (tmVef.sType = "L") Or (tmVef.sType = "A") Or (tmVef.sType = "G") Then
        If rbcShowRateOnInsertion(0).Value Then
            tlVpf.sShowRateOnInsert = "Y"
        Else
            tlVpf.sShowRateOnInsert = "N"
        End If
    'End If
    tlVpf.sWegenerExport = "N"
    If ((Asc(tgSpf.sUsingFeatures7) And WEGENEREXPORT) = WEGENEREXPORT) Then
        If ckcAffExport(0).Value = vbChecked Then
            tlVpf.sWegenerExport = "Y"
        End If
    End If
    tlVpf.sOLAExport = "N"
    'If ((Asc(tgSpf.sUsingFeatures7) And OLAEXPORT) = OLAEXPORT) Then
    '    If ckcAffExport(1).Value = vbChecked Then
    '        tlVpf.sOLAExport = "Y"
    '    End If
    'End If
    smEMail = Trim$(edcEMail.Text)
    'If smXDXMLForm = "ISCI" Then
    If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) Then
        tlVpf.iInterfaceID = Val(edcInterfaceID(1).Text)
    Else
        tlVpf.iInterfaceID = 0
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRec                        *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record to record          *
'*                                                     *
'*******************************************************
Private Sub mMoveRec(tlFromRec As VPF, tlToRec As VPF)
    tlToRec = tlFromRec
'    tlToRec.iVEFKCode = tlFromRec.iVEFKCode  'Input Vehicle
'    tlToRec.iGTime(0) = tlFromRec.iGTime(0)  'Default to 12M
'    tlToRec.iGTime(1) = tlFromRec.iGTime(1)
'    tlToRec.sGMedium = tlFromRec.sGMedium   'Radio
'    tlToRec.iUrfGCode = tlFromRec.iUrfGCode     'Counterpoint
'    tlToRec.iGMoFull = tlFromRec.iGMoFull      'month retain unpacked
'    tlToRec.iGMoPack = tlFromRec.iGMoPack     'months retain packed
'    tlToRec.sSVarComm = tlFromRec.sSVarComm    'Agency commission fixed
'    tlToRec.sSCompType = tlFromRec.sSCompType  'Separate competitive by break
'    tlToRec.iSCompLen(0) = tlFromRec.iSCompLen(0)  '0 for competitive time separation
'    tlToRec.iSCompLen(1) = tlFromRec.iSCompLen(1)
'    tlToRec.sSSellout = tlFromRec.sSSellout   'Sellout by Units
'    tlToRec.sSOverbook = tlFromRec.sSOverbook  'Disallow overbooking
'    tlToRec.sSForceMG = tlFromRec.sSForceMG   'Don't set spots moved outside limit as MG's
'    tlToRec.sSPlaceNet = tlFromRec.sSPlaceNet  'Place network spots first
'    tlToRec.sSAvailOrder = tlFromRec.sSAvailOrder    'Avail order= N/A
'    For ilLoop = 0 To 9     'Length and group #
'        tlToRec.isLen(ilLoop) = tlFromRec.isLen(ilLoop)
'        tlToRec.isLenGroup(ilLoop) = tlFromRec.isLenGroup(ilLoop)
'    Next ilLoop
'    tlToRec.sSCommCalc = tlFromRec.sSCommCalc  'Calculate salesperson commission on collected amount
'    tlToRec.sCPrtStyle = tlFromRec.sCPrtStyle  'Contract print style= Narrow
'    tlToRec.sCTargets = tlFromRec.sCTargets   'Don't include demo targets
'    tlToRec.sCPriceStat = tlFromRec.sCPriceStat 'Don't compute spot price statisitics
'    tlToRec.sPGridRes = tlFromRec.sPGridRes   'default grid resolution is Full hour
'    tlToRec.sYScript = tlFromRec.sYScript    'Using scripts
'    tlToRec.iLLD(0) = tlFromRec.iLLD(0)       'No date- set to zero
'    tlToRec.iLLD(1) = tlFromRec.iLLD(1)
'    tlToRec.iLPD(0) = tlFromRec.iLPD(0)       'No date- set to zero
'    tlToRec.iLPD(1) = tlFromRec.iLPD(1)
'    tlToRec.sLTimeZone = tlFromRec.sLTimeZone  'Time zone = pacific
'    tlToRec.sLDayLight = tlFromRec.sLDayLight  'Daylight saving time
'    tlToRec.sLTiming = tlFromRec.sLTiming    'Using log timing
'    tlToRec.sLAvailLen = tlFromRec.sLAvailLen  'Show length of unsold avails
'    tlToRec.sLBBs = tlFromRec.sLBBs       'Show bbs on schedule line as BBs
'    tlToRec.sLShowCut = tlFromRec.sLShowCut    'Show cut/instruction # or neither
'    tlToRec.sLTimeFormat = tlFromRec.sLTimeFormat    'Time format AM/PM
'    tlToRec.sBDefCycle = tlFromRec.sBDefCycle  'default billing cycle (standard)
'    tlToRec.iBCal(0) = tlFromRec.iBCal(0)      'No date for last calendar date
'    tlToRec.iBCal(1) = tlFromRec.iBCal(1)
'    tlToRec.iBStd(0) = tlFromRec.iBStd(0)     'No date for last standard date
'    tlToRec.iBStd(1) = tlFromRec.iBStd(1)
'    tlToRec.iBWk(0) = tlFromRec.iBWk(0)      'No date for last weekly date
'    tlToRec.iBWk(1) = tlFromRec.iBWk(1)
'    tlToRec.sBLockBox = tlFromRec.sBLockBox   'Don't use lock boxes
'    tlToRec.sBAgyDPServ = tlFromRec.sBAgyDPServ 'Don't use Agency DP service
'    tlToRec.sRPenny = tlFromRec.sRPenny
'    'Assume all G/L are undefined
'    For ilLoop = 0 To 33 Step 1
'        tlToRec.sRGL(ilLoop) = tlFromRec.sRGL(ilLoop)
'    Next ilLoop
'    'Assume no user defined transactions
'    For ilLoop = 0 To 9 Step 1
'        tlToRec.sRPrefix(ilLoop) = tlFromRec.sRPrefix(ilLoop)
'        tlToRec.sRSuffix(ilLoop) = tlFromRec.sRSuffix(ilLoop)
'        tlToRec.sRSign(ilLoop) = tlFromRec.sRSign(ilLoop)
'        tlToRec.sRUGL(ilLoop) = tlFromRec.sRUGL(ilLoop)
'    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl(tlVpf As VPF, tlVff As VFF)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilYear                                                                                *
'******************************************************************************************

'
'   mMoveRecToCtrl
'   Where:
'
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilRnf As Integer
    Dim ilList As Integer
    Dim ilRnfCode As Integer
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim ilFound As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slTDate As String
    Dim ilMonth As Integer
    Dim ilValue As Integer
    Dim ilLevel As Integer

    imIgnoreClickEvent = True
    
    udcVehOptTabs.Action 7, 0
    
    ' tlVpf.iVEFKCode  'Input Vehicle
    'General
    If tgSpf.sSystemType = "R" Then
        For ilLoop = LBound(tlVpf.sGZone) To UBound(tlVpf.sGZone) Step 1
            tlVpf.sGZone(ilLoop) = ""   'Zone name
            tlVpf.sGFed(ilLoop) = ""
            tlVpf.iGLocalAdj(ilLoop) = 0    'Local adjustment
            tlVpf.iGFeedAdj(ilLoop) = 0
            tlVpf.iGV1Z(ilLoop) = 0
            tlVpf.iGV2Z(ilLoop) = 0
            tlVpf.iGV3Z(ilLoop) = 0
            tlVpf.iGV4Z(ilLoop) = 0
            tlVpf.sGCSVer(ilLoop) = ""   'Don't Create log for time zone
            tlVpf.iGMnfNCode(ilLoop) = 0
            tlVpf.sGBus(ilLoop) = ""   'Don't Create log for time zone
            tlVpf.sGSked(ilLoop) = ""   'Don't Create log for time zone
            tlVff.sFedDelivery(ilLoop) = ""
        Next ilLoop
    End If
    gUnpackTime tlVpf.iGTime(0), tlVpf.iGTime(1), "A", "1", slStr
    'edcGSignOn.Text = slStr
    edcGen(Signon).Text = slStr
    Select Case tlVpf.sGMedium
        Case "R"
            rbcGMedium(0).Value = True
        Case "T"
            rbcGMedium(1).Value = True
        Case "N"
            rbcGMedium(2).Value = True
        Case "V"
            rbcGMedium(3).Value = True
        Case "C"
            rbcGMedium(4).Value = True
        Case "S"
            rbcGMedium(5).Value = True
        '10050
        Case "P"                                '1-16-14 podcast
            rbcGMedium(PODCASTRBC).Value = True
    End Select
    Select Case tlVpf.sEmbeddedOrROS
        Case "E"
            rbcGMedium(7).Value = True
        Case "R"
            rbcGMedium(8).Value = True
        Case Else
            rbcGMedium(8).Value = True
    End Select
    rbcRemoteExport(2).Value = True
    rbcRemoteImport(2).Value = True
    If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Then
        'Select Case tlVpf.sFeedLogOrder
        '    Case "I"
        '        rbcRemoteExport(0).Value = True
        '    Case "A"
        '        rbcRemoteExport(1).Value = True
        '    Case Else
        '        rbcRemoteExport(2).Value = True
        'End Select
        If (Asc(tgSpf.sUsingFeatures5) And REMOTEEXPORT) = REMOTEEXPORT Then
            If (Asc(tlVpf.sUsingFeatures1) And EXPORTINSERTION) = EXPORTINSERTION Then
                rbcRemoteExport(0).Value = True
            ElseIf (Asc(tlVpf.sUsingFeatures1) And EXPORTLOG) = EXPORTLOG Then
                rbcRemoteExport(1).Value = True
            End If
        End If
        If (Asc(tgSpf.sUsingFeatures5) And REMOTEIMPORT) = REMOTEIMPORT Then
            If (Asc(tlVpf.sUsingFeatures1) And IMPORTINSERTION) = IMPORTINSERTION Then
                rbcRemoteImport(0).Value = True
            ElseIf (Asc(tlVpf.sUsingFeatures1) And IMPORTAFFILIATESPOTS) = IMPORTAFFILIATESPOTS Then
                rbcRemoteImport(1).Value = True
            End If
        End If
    End If
    ckcAffExport(2).Value = vbUnchecked
    If (Asc(tlVpf.sUsingFeatures1) And EXPORTISCIBYPLEDGE) = EXPORTISCIBYPLEDGE Then
        ckcAffExport(2).Value = vbChecked
    End If
    udcVehOptTabs.WebInfo(3) = vbUnchecked
    If (Asc(tlVpf.sUsingFeatures1) And SUPPRESSWEBLOG) = SUPPRESSWEBLOG Then
        udcVehOptTabs.WebInfo(3) = vbChecked
    End If
    udcVehOptTabs.WebInfo(4) = vbUnchecked
    If (Asc(tlVpf.sUsingFeatures1) And EXPORTPOSTEDTIMES) = EXPORTPOSTEDTIMES Then
        udcVehOptTabs.WebInfo(4) = vbChecked
    End If
    
    ckcXDSave(3).Value = vbUnchecked
    If (Asc(tlVpf.sUsingFeatures2) And XDSAPPLYMERGE) = XDSAPPLYMERGE Then
        ckcXDSave(3).Value = vbChecked
    End If
    If (Asc(tlVpf.sUsingFeatures2) And INVOICEAIRDATEWO) = INVOICEAIRDATEWO Then
        rbcShowAirDate(1).Value = True
    Else
        rbcShowAirDate(0).Value = True
    End If
    '10933 10975
    If tlVff.sXDEventZone = "Y" Then
        ckcXDSave(CUEZONE).Value = vbChecked
    Else
         ckcXDSave(CUEZONE).Value = vbUnchecked
    End If
    Select Case tlVpf.sMoveLLD  'Move Today to LLD
        Case "Y"
            rbcSMoveLLD(0).Value = True
        Case Else
            rbcSMoveLLD(1).Value = True
    End Select
    If tmVef.sType = "S" Then
        plcBillSA.Enabled = True
        Select Case tlVpf.sBillSA  'Bill Selling or Airing
            Case "Y"
                rbcBillSA(0).Value = True
            Case Else
                rbcBillSA(1).Value = True
        End Select
    Else
        plcBillSA.Enabled = False
        rbcBillSA(1).Value = True
    End If
    ' tlVpf.iUrfGCode = tlFromRec.iUrfGCode     'Counterpoint
    'edcGPast.Text = Trim$(Str$(tlVpf.iGMoFull))      'month retain unpacked
    'edcGHistory.Text = Trim$(Str$(tlVpf.iGMoPack))   'months retain packed
    'Select Case tlVpf.sGPriceStat  'Spot Price statistics
    '    Case "Y"
    '        rbcSPrice(0).Value = True
    '    Case Else
    '        rbcSPrice(1).Value = True
    'End Select
    If frcExport(3).Visible Then
        Select Case tlVpf.sOwnership
            Case "B"
                rbcGMedium(10).Value = True
            Case "C"
                rbcGMedium(11).Value = True
            Case "D"
                rbcGMedium(12).Value = True
            Case Else
                rbcGMedium(9).Value = True
        End Select
    Else
        rbcGMedium(9).Value = True
    End If
    
    Select Case tlVpf.sGenLog  'Generate Log
        Case "N"
            rbcGenLog(1).Value = True
        Case "L"
            rbcGenLog(2).Value = True
        Case "M"
            rbcGenLog(3).Value = True
        Case "A"
            rbcGenLog(4).Value = True
        Case Else
            rbcGenLog(0).Value = True
    End Select
    Select Case tlVpf.sGGridRes  'Grid resolution
        Case "F"
            rbcLGrid(0).Value = True
        Case "H"
            rbcLGrid(1).Value = True
        Case Else
            rbcLGrid(2).Value = True
    End Select
    'Select Case tlVpf.sGScript  'Scripts
    '    Case "Y"
    '        rbcLScripts(0).Value = True
    '    Case Else
    '        rbcLScripts(1).Value = True
    'End Select
    Select Case tlVpf.sExpHiCorp
        Case "N"
            ckcIFExport(5).Value = vbUnchecked
        Case Else
            ckcIFExport(5).Value = vbChecked
    End Select
    pbcGTZ.Cls
    mGTZPaint tlVpf, tlVff     'Log zone generation
    Select Case tlVpf.sShowTime  'Scripts
        Case "H"
            rbcShowAirTime(1).Value = True
        Case "D"
            rbcShowAirTime(2).Value = True
        Case "A"
            rbcShowAirTime(3).Value = True
        Case "S"
            rbcShowAirTime(0).Value = True
        Case Else
            rbcShowAirTime(0).Value = False
            rbcShowAirTime(1).Value = False
            rbcShowAirTime(2).Value = False
            rbcShowAirTime(3).Value = False
    End Select
    If (tmVef.sType = "S") Or (tmVef.sType = "A") Then
        edcGen(SAGROUPNO).Text = Trim$(Str$(tlVpf.iSAGroupNo))
        imOrigSAGroupNo = tlVpf.iSAGroupNo
    Else
        edcGen(SAGROUPNO).Text = ""
        imOrigSAGroupNo = 0
    End If
    'edcGEDI(0).Text = Trim$(tlVpf.sEDICallLetters)
    'edcGEDI(1).Text = Trim$(tlVpf.sEDIBand)
    edcGen(EDICALLLETTERS).Text = Trim$(tlVpf.sEDICallLetters)
    edcGen(EDIBAND).Text = Trim$(tlVpf.sEDIBand)
    edcGen(0).Text = Trim$(tlVpf.sAccruedRevenue)
    edcGen(1).Text = Trim$(tlVpf.sAccruedTrade)
    edcGen(2).Text = Trim$(tlVpf.sBilledRevenue)
    edcGen(3).Text = Trim$(tlVpf.sBilledTrade)
    'Sales
    Select Case tlVpf.sSVarComm     'Agency commission fixed
        Case "Y"
            rbcSCommission(0).Value = True
        Case "N"
            rbcSCommission(1).Value = True
    End Select
    Select Case tlVpf.sAdvtSep  'Separate advertiser by time or break
        Case "T"
            rbcSAdvtSep(0).Value = True
        Case "B"
            rbcSAdvtSep(1).Value = True
        Case Else
            rbcSAdvtSep(0).Value = False
            rbcSAdvtSep(1).Value = False
    End Select
    Select Case tlVpf.sSCompType  'Separate competitive by break
        Case "T"
            rbcSCompetitive(0).Value = True
        Case "B"
            rbcSCompetitive(1).Value = True
        Case Else
            rbcSCompetitive(2).Value = True
    End Select
    If tlVpf.sSCompType = "T" Then
        gUnpackLength tlVpf.iSCompLen(0), tlVpf.iSCompLen(1), "2", True, slStr
    Else
        slStr = ""
    End If
    edcSCompSepLen.Text = slStr
    Select Case tlVpf.sSSellOut  'Sellout
        Case "U"
            rbcSSellout(0).Value = True
        Case "B"
            rbcSSellout(1).Value = True
        Case "T"
            rbcSSellout(2).Value = True
        Case Else
            rbcSSellout(3).Value = True
    End Select
    Select Case tlVpf.sSOverBook  'Overbook
        Case "Y"
            rbcSOverbook(0).Value = True
        Case Else
            rbcSOverbook(1).Value = True
    End Select
    Select Case tlVpf.sSForceMG  'Overbook
        Case "W"
            rbcSMove(0).Value = True
        Case Else
            rbcSMove(1).Value = True
    End Select
    'Select Case tlVpf.sSIntoRC  'Contract spots into Network avails
    '    Case "Y"
    '        rbcSIntoRC(0).Value = True
    '    Case Else
    '        rbcSIntoRC(1).Value = True
    'End Select
    'Select Case tlVpf.sSPTA  'Keep spots with program, time or ask
    '    Case "P"
    '        rbcSPTA(0).Value = True
    '    Case "T"
    '        rbcSPTA(1).Value = True
    '    Case Else
    '        rbcSPTA(2).Value = True
    'End Select
    'Select Case tlVpf.sSPlaceNet  'Network spot placement
    '    Case "F"
    '        rbcSMixed(0).Value = True
    '    Case "L"
    '        rbcSMixed(1).Value = True
    '    Case Else
    '        rbcSMixed(2).Value = True
    'End Select
    Select Case tlVpf.sSAvailOrder  'spot placement
        Case "1"
            rbcSBreak(0).Value = True
        Case "2"
            rbcSBreak(1).Value = True
        Case "3"
            rbcSBreak(2).Value = True
        Case "4"
            rbcSBreak(3).Value = True
        Case "5"
            rbcSBreak(4).Value = True
        Case "6"
            rbcSBreak(5).Value = True
        Case Else
            rbcSBreak(6).Value = True
    End Select

    slTDate = Format$(gNow, "m/d/yy")
    gObtainMonthYear 0, slTDate, ilMonth, imYear
    mObtainNIFYear

    pbcSSpotLen.Cls
    mSSpotLenPaint tlVpf   'Spot length and group number
    Select Case tlVpf.sSCommCalc  'Salesperson commission
        Case "B"
            rbcSSalesperson(0).Value = True
        Case "C"
            rbcSSalesperson(1).Value = True
    End Select
    'PSA/Promo
    pbcPSA.Cls
    mMPsaPaint tlVpf     'PSA
    pbcPromo.Cls
    mMPromoPaint tlVpf   'Promo
    'Log
    gUnpackDate tlVpf.iLLastDateCpyAsgn(0), tlVpf.iLLastDateCpyAsgn(1), slStr   'Last log date
    edcLLDCpyAsgn.Text = slStr
    gUnpackDate tlVpf.iLLD(0), tlVpf.iLLD(1), slStr   'Last log date
    edcLDate(0).Text = slStr
    gUnpackDate tlVpf.iLPD(0), tlVpf.iLPD(1), slStr   'Last preliminary date
    edcLDate(1).Text = slStr
    gUnpackDate tlVpf.iLastLog(0), tlVpf.iLastLog(1), slStr   'Last log date
    edcLAffDate(0).Text = slStr
    gUnpackDate tlVpf.iLastCP(0), tlVpf.iLastCP(1), slStr   'Last preliminary date
    edcLAffDate(1).Text = slStr
    Select Case tlVpf.slTimeZone  'Grid resolution
        Case "E"
            rbcLZone(0).Value = True
        Case "C"
            rbcLZone(1).Value = True
        Case "M"
            rbcLZone(2).Value = True
        Case Else
            rbcLZone(3).Value = True
    End Select
    Select Case tlVpf.sLDayLight  'Daylight saving
        Case "Y"
            rbcLDaylight(0).Value = True
        Case Else
            rbcLDaylight(1).Value = True
    End Select
    Select Case tlVpf.sLTiming  'Log Timing
        Case "Y"
            rbcLTiming(0).Value = True
        Case Else
            rbcLTiming(1).Value = True
    End Select
    'Select Case tlVpf.sLAvailLen  'Show lengths on unsold avails
    '    Case "Y"
    '        rbcLLen(0).Value = True
    '    Case Else
    '        rbcLLen(1).Value = True
    'End Select
    edcSLen.Text = Trim$(Str$(tlVpf.iSDLen))      'default Length
    'Select Case tlVpf.sAffCPs  'Create Affiliate Custom CPs
    '    Case "Y"
    '        rbcLAffCPs(0).Value = True
    '    Case Else
    '        rbcLAffCPs(1).Value = True
    'End Select
    'Select Case tlVpf.sAffTimes  'Post Affiliate Exact Times
    '    Case "Y"
    '        rbcLAffTimes(0).Value = True
    '    Case Else
    '        rbcLAffTimes(1).Value = True
    'End Select
    'Select Case tlVpf.sLShowCut  'Show cut/instruction #
    '    Case "C"
    '        rbcLCut(0).Value = True
    '    Case "I"
    '        rbcLCut(1).Value = True
    '    Case "B"
    '        rbcLCut(2).Value = True
    '    Case Else
    '        rbcLCut(3).Value = True
    'End Select
    'Select Case tlVpf.slTimeFormat  'Time format
    '    Case "A"
    '        rbcLTime(0).Value = True
    '    Case Else
    '        rbcLTime(1).Value = True
    'End Select
    If tmVef.sType = "A" Then
        Select Case tlVpf.sCopyOnAir  'Allow Copy on Airing Vehicles
            Case "Y"
                rbcLCopyOnAir(0).Value = True
            Case Else
                rbcLCopyOnAir(1).Value = True
        End Select
    Else
        rbcLCopyOnAir(1).Value = True
    End If
    For ilLoop = 0 To 5 Step 1
        Select Case ilLoop
            Case 0
                ilRnfCode = tlVpf.iRnfLogCode
            Case 1
                ilRnfCode = tlVpf.iRnfCertCode
            Case 2
                ilRnfCode = tlVpf.iRnfPlayCode
            Case 3
                ilRnfCode = tlVpf.iRnfSvLogCode
            Case 4
                ilRnfCode = tlVpf.iRnfSvCertCode
            Case 5
                ilRnfCode = tlVpf.iRnfSvPlayCode
        End Select
        If ilLoop <= 2 Then
            cbcLog(ilLoop).ListIndex = 0
            For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
                If tgRnfList(ilRnf).tRnf.iCode = ilRnfCode Then
                    For ilList = 0 To cbcLog(ilLoop).ListCount - 1 Step 1
                        If cbcLog(ilLoop).List(ilList) = UCase$(Trim$(tgRnfList(ilRnf).tRnf.sName)) Then
                            cbcLog(ilLoop).ListIndex = ilList
                            Exit For
                        End If
                    Next ilList
                    Exit For
                End If
            Next ilRnf
        Else
            cbcSvLog(ilLoop - 3).ListIndex = 0
            For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
                If tgRnfList(ilRnf).tRnf.iCode = ilRnfCode Then
                    For ilList = 0 To cbcSvLog(ilLoop - 3).ListCount - 1 Step 1
                        If cbcSvLog(ilLoop - 3).List(ilList) = UCase$(Trim$(tgRnfList(ilRnf).tRnf.sName)) Then
                            cbcSvLog(ilLoop - 3).ListIndex = ilList
                            Exit For
                        End If
                    Next ilList
                    Exit For
                End If
            Next ilRnf
        End If
    Next ilLoop
    If tlVpf.sUnsoldBlank = "N" Then
        ckcUnsoldBlank.Value = vbUnchecked
    Else
        ckcUnsoldBlank.Value = vbChecked
    End If
    If tlVpf.sAvailNameOnWeb = "Y" Then
        udcVehOptTabs.WebInfo(0) = vbChecked
    Else
        udcVehOptTabs.WebInfo(0) = vbUnchecked
    End If
    If tlVpf.sWebLogFeedTime = "Y" Then
        udcVehOptTabs.WebInfo(1) = vbChecked
    Else
       udcVehOptTabs.WebInfo(1) = vbUnchecked
    End If
    If tlVpf.sWebLogSummary = "Y" Then
        udcVehOptTabs.WebInfo(2) = vbChecked
    Else
        udcVehOptTabs.WebInfo(2) = vbUnchecked
    End If
    tmArfSrchKey.iCode = tlVpf.iFTPArfCode
    If tmArfSrchKey.iCode <> 0 Then
        ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smFTP = Trim$(tmArf.sFTP)
        Else
            smFTP = ""
        End If
    Else
        smFTP = ""
    End If
    'edcLog(0).Text = smFTP
    udcVehOptTabs.FTPInfo() = smFTP
    smLiveWindow = tmVff.iLiveCompliantAdj
    If Val(smLiveWindow) <= 0 Then
        udcVehOptTabs.LiveWindow() = "5"
    Else
        udcVehOptTabs.LiveWindow() = smLiveWindow
    End If
    '2/28/19: Add Cart on Web
    smCartOnWeb = tmVff.sCartOnWeb
    If smCartOnWeb = "Y" Then
        udcVehOptTabs.WebInfo(5) = vbChecked
    Else
        udcVehOptTabs.WebInfo(5) = vbUnchecked
    End If
    
    If tmVff.sDefaultAudioType = "L" Then
        rbcAudioType(1).Value = True
    ElseIf tmVff.sDefaultAudioType = "M" Then
        rbcAudioType(2).Value = True
    ElseIf tmVff.sDefaultAudioType = "S" Then
        rbcAudioType(3).Value = True
    ElseIf tmVff.sDefaultAudioType = "P" Then
        rbcAudioType(4).Value = True
    ElseIf tmVff.sDefaultAudioType = "Q" Then
        rbcAudioType(5).Value = True
    Else
        rbcAudioType(0).Value = True
    End If
    
    tmArfSrchKey.iCode = tmVff.iLogExptArfCode
    If tmArfSrchKey.iCode <> 0 Then
        ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smLogExpt = Trim$(tmArf.sFTP)
        Else
            smLogExpt = ""
        End If
    Else
        smLogExpt = ""
    End If
    edcLog(0).Text = smLogExpt
            
    'edcGEDI(2).Text = Trim$(tmVff.sASICallLetters)
    'edcGEDI(3).Text = Trim$(tmVff.sASIBand)
    edcGen(VEDICALLLETTERS).Text = Trim$(tmVff.sASICallLetters)
    edcGen(VEDIBAND).Text = Trim$(tmVff.sASIBand)
            
            
    tmArfSrchKey.iCode = tlVpf.iAutoExptArfCode
    If tmArfSrchKey.iCode <> 0 Then
        ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smAutoExpt = Trim$(tmArf.sFTP)
        Else
            smAutoExpt = ""
        End If
    Else
        smAutoExpt = ""
    End If
    edcLog(1).Text = smAutoExpt

    tmArfSrchKey.iCode = tlVpf.iAutoImptArfCode
    If tmArfSrchKey.iCode <> 0 Then
        ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smAutoImpt = Trim$(tmArf.sFTP)
        Else
            smAutoImpt = ""
        End If
    Else
        smAutoImpt = ""
    End If
    edcLog(2).Text = smAutoImpt

    'Log Form
    pbcLogForm.Cls
    If ((tmVef.sType = "C") And (tmVef.iVefCode = 0)) Or (tmVef.sType = "L") Or (tmVef.sType = "A") Or (tmVef.sType = "G") Then
        'If (tgSpf.sGUseAffSys = "Y") And (igVpfType <> 1) Then
        If (igVpfType <> 1) Then
            For ilRow = 1 To 3 Step 1
                If ilRow = 1 Then
                    tmVof = tmLVof
                ElseIf ilRow = 2 Then
                    tmVof = tmCVof
                Else
                    tmVof = tmOVof
                End If
                imLSave(LNODAYSINDEX, ilRow) = tmVof.iNoDaysCP
                slStr = Trim$(Str$(imLSave(LNODAYSINDEX, ilRow)))
                gSetShow pbcLogForm, slStr, tmLCtrls(LNODAYSINDEX)
                smLShow(LNODAYSINDEX, ilRow) = tmLCtrls(LNODAYSINDEX).sShow
                imLRowNo = ilRow
                If tmVof.sSkipPage = "Y" Then
                    imLSave(LSKIPINDEX, ilRow) = 1
                Else
                    imLSave(LSKIPINDEX, ilRow) = 0
                End If
                mLSetShow LSKIPINDEX
                If tmVof.sShowLen = "Y" Then
                    imLSave(LLENINDEX, ilRow) = 1
                Else
                    imLSave(LLENINDEX, ilRow) = 0
                End If
                mLSetShow LLENINDEX
                If tmVof.sShowProduct = "Y" Then
                    imLSave(LPRODINDEX, ilRow) = 1
                Else
                    imLSave(LPRODINDEX, ilRow) = 0
                End If
                mLSetShow LPRODINDEX
                If tmVof.sShowCreative = "Y" Then
                    imLSave(LTITLEINDEX, ilRow) = 1
                Else
                    imLSave(LTITLEINDEX, ilRow) = 0
                End If
                mLSetShow LTITLEINDEX
                If tmVof.sShowISCI = "Y" Then
                    imLSave(LISCIINDEX, ilRow) = 1
                Else
                    imLSave(LISCIINDEX, ilRow) = 0
                End If
                mLSetShow LISCIINDEX
                If tmVof.sShowDP = "Y" Then
                    imLSave(LDAYPARTINDEX, ilRow) = 1
                Else
                    imLSave(LDAYPARTINDEX, ilRow) = 0
                End If
                mLSetShow LDAYPARTINDEX
                If tmVof.sShowAirTime = "Y" Then
                    imLSave(LTIMEINDEX, ilRow) = 1
                Else
                    imLSave(LTIMEINDEX, ilRow) = 0
                End If
                mLSetShow LTIMEINDEX
                If tmVof.sShowAirLine = "Y" Then
                    imLSave(LLINEINDEX, ilRow) = 1
                Else
                    imLSave(LLINEINDEX, ilRow) = 0
                End If
                mLSetShow LLINEINDEX
                If tmVof.sShowHour = "Y" Then
                    imLSave(LHOURINDEX, ilRow) = 1
                Else
                    imLSave(LHOURINDEX, ilRow) = 0
                End If
                mLSetShow LHOURINDEX
                imLSave(LLOADINDEX, ilRow) = tmVof.iLoadFactor
                slStr = Trim$(Str$(imLSave(LLOADINDEX, ilRow)))
                gSetShow pbcLogForm, slStr, tmLCtrls(LLOADINDEX)
                smLShow(LLOADINDEX, ilRow) = tmLCtrls(LLOADINDEX).sShow
                tmCefSrchKey.lCode = tmVof.lHd1CefCode
                If tmCefSrchKey.lCode <> 0 Then
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        'If tmCef.iStrLen > 0 Then
                        '    smLSave(1, ilRow) = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                        'Else
                        '    smLSave(1, ilRow) = ""
                        'End If
                        smLSave(1, ilRow) = gStripChr0(tmCef.sComment)
                    Else
                        smLSave(1, ilRow) = ""
                    End If
                Else
                    smLSave(1, ilRow) = ""
                End If
                slStr = smLSave(1, ilRow)
                gSetShow pbcLogForm, slStr, tmLCtrls(LHEADERINDEX)
                smLShow(LHEADERINDEX, ilRow) = tmLCtrls(LHEADERINDEX).sShow
                'smInitLgHd1 = edcLgHd1.Text

                tmCefSrchKey.lCode = tmVof.lFt1CefCode
                If tmCefSrchKey.lCode <> 0 Then
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        'If tmCef.iStrLen > 0 Then
                        '    smLSave(2, ilRow) = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                        'Else
                        '    smLSave(2, ilRow) = ""
                        'End If
                        smLSave(2, ilRow) = gStripChr0(tmCef.sComment)
                    Else
                        smLSave(2, ilRow) = ""
                    End If
                Else
                    smLSave(2, ilRow) = ""
                End If
                slStr = smLSave(2, ilRow)
                gSetShow pbcLogForm, slStr, tmLCtrls(LFOOT1INDEX)
                smLShow(LFOOT1INDEX, ilRow) = tmLCtrls(LFOOT1INDEX).sShow
                'smInitLgFt1 = edcLgFt1.Text

                tmCefSrchKey.lCode = tmVof.lFt2CefCode
                If tmCefSrchKey.lCode <> 0 Then
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        'If tmCef.iStrLen > 0 Then
                        '    smLSave(3, ilRow) = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                        'Else
                        '    smLSave(3, ilRow) = ""
                        'End If
                        smLSave(3, ilRow) = gStripChr0(tmCef.sComment)
                    Else
                        smLSave(3, ilRow) = ""
                    End If
                Else
                    smLSave(3, ilRow) = ""
                End If
                slStr = smLSave(3, ilRow)
                gSetShow pbcLogForm, slStr, tmLCtrls(LFOOT2INDEX)
                smLShow(LFOOT2INDEX, ilRow) = tmLCtrls(LFOOT2INDEX).sShow
                'smInitLgFt2 = edcLgFt2.Text
            Next ilRow
            

        End If
    End If
    'E-Mail
    If (igVpfType <> 1) Then
        tmCefSrchKey.lCode = tlVpf.lEMailCefCode
        If tmCefSrchKey.lCode <> 0 Then
            tmCef.sComment = ""
            imCefRecLen = Len(tmCef)    '1009
            ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                'If tmCef.iStrLen > 0 Then
                '    smEMail = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                'Else
                '    smEMail = ""
                'End If
                smEMail = gStripChr0(tmCef.sComment)
            Else
                smEMail = ""
            End If
        Else
            smEMail = ""
        End If
        edcEMail.Text = smEMail
    Else
        smEMail = ""
        edcEMail.Text = ""
    End If
    pbcLogForm_Paint
    'Price Levels
    If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Then
        If (tmSaf.iCode = 0) Or (tmSaf.lLowPrice = 0) Or (tmSaf.lHighPrice = 0) Then
            edcSchedule(0).Text = ""
            edcSchedule(1).Text = ""
            For ilLoop = LBONE To UBound(lmSSave) Step 1
                lmSSave(ilLoop) = 0
            Next ilLoop
        Else
            edcSchedule(0).Text = Trim$(Str$(tmSaf.lLowPrice))
            edcSchedule(1).Text = Trim$(Str$(tmSaf.lHighPrice))
            lmSSave(LBONE) = tmSaf.lLowPrice
            lmSSave(UBound(lmSSave) - 1) = tmSaf.lHighPrice
            ilLevel = LBound(tmSaf.lLevelToPrice)
            For ilLoop = LBONE + 1 To UBound(lmSSave) - 2 Step 1
                lmSSave(ilLoop) = tmSaf.lLevelToPrice(ilLevel)  '(ilLoop - LBONE)
                ilLevel = ilLevel + 1
            Next ilLoop
        End If
        
    End If
    If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Then
        If (Asc(tmSaf.sFeatures1) And SUPPRESSPREEMPTION) = SUPPRESSPREEMPTION Then
            ckcSch(4).Value = vbChecked
        Else
            ckcSch(4).Value = vbUnchecked
        End If
    Else
        ckcSch(4).Value = vbUnchecked
    End If
    
    If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "N") Or (tmVef.sType = "R") Or (tmVef.sType = "G")) And ((Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS) Then
        edcGP(0).Text = Trim$(tmVaf.sDivisionCode)
        edcGP(1).Text = Trim$(tmVaf.sBranchCodeCash)
        edcGP(2).Text = Trim$(tmVaf.sPCGrossSalesCash)
        edcGP(3).Text = Trim$(tmVaf.sPCAgyCommCash)
        edcGP(4).Text = Trim$(tmVaf.sPCRecvCash)
        edcGP(5).Text = Trim$(tmVaf.sBranchCodeTrade)
        edcGP(6).Text = Trim$(tmVaf.sPCGrossSalesTrade)
        edcGP(7).Text = Trim$(tmVaf.sPCRecvTrade)
        edcGP(8).Text = Trim$(tmVaf.sVendorID)              '8-25-10
    End If
    imGreatPlainsAltered = False
    ''Export X-Digital
    'edcXDISCIPrefix(0).Text = Trim$(tmVff.sXDISCIPrefix)
    'If tmVff.sXDXMLForm = "P" Then
        'smXDXMLForm = "ISCI"
    If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) Then
        If (tlVpf.iInterfaceID > 0) Then
            edcInterfaceID(1).Text = Trim$(Str$(tlVpf.iInterfaceID))
            edcXDISCIPrefix(1).Text = Trim$(tmVff.sXDSISCIPrefix)
        Else
            edcInterfaceID(1).Text = ""
            edcXDISCIPrefix(1).Text = ""
        End If
    Else
        edcInterfaceID(1).Text = ""
        edcXDISCIPrefix(1).Text = ""
    End If
    'ElseIf tmVff.sXDXMLForm = "A" Then
    If ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
        If tmVff.sXDXMLForm = "A" Then
            smXDXMLForm = "H#B#P#"
            edcInterfaceID(0).Text = tmVff.sXDProgCodeID
            edcXDISCIPrefix(0).Text = Trim$(tmVff.sXDISCIPrefix)
        ElseIf tmVff.sXDXMLForm = "S" Then
            smXDXMLForm = "H#B#"
            edcInterfaceID(0).Text = tmVff.sXDProgCodeID
            edcXDISCIPrefix(0).Text = Trim$(tmVff.sXDISCIPrefix)
        Else
            smXDXMLForm = ""
            edcInterfaceID(0).Text = ""
            edcXDISCIPrefix(0).Text = ""
        End If
    Else
        smXDXMLForm = ""
        edcInterfaceID(0).Text = ""
        edcXDISCIPrefix(0).Text = ""
    End If
    pbcXDXMLForm_Paint
    '9114 note that ckcXDSave(3) must be set before this call
    smOriginalXdsProgramCode = edcInterfaceID(0).Text
    If ckcXDSave(3).Value = vbChecked Then
        bmOriginalXdsHonorMerge = True
    Else
        bmOriginalXdsHonorMerge = False
    End If
    Select Case tmVff.sXDSaveCF  'Show carts on Station feed
        Case "N"
            ckcXDSave(0).Value = vbUnchecked
        Case Else
            ckcXDSave(0).Value = vbChecked
    End Select
    Select Case tmVff.sXDSaveHDD  'Show carts on Station feed
        Case "Y"
            ckcXDSave(1).Value = vbChecked
        Case Else
            ckcXDSave(1).Value = vbUnchecked
    End Select
    Select Case tmVff.sXDSaveNAS  'Show carts on Station feed
        Case "Y"
            ckcXDSave(2).Value = vbChecked
        Case Else
            ckcXDSave(2).Value = vbUnchecked
    End Select
    Select Case tmVff.sXDSSaveCF  'Show carts on Station feed
        Case "N"
            ckcXDSave(4).Value = vbUnchecked
        Case Else
            ckcXDSave(4).Value = vbChecked
    End Select
    Select Case tmVff.sXDSSaveHDD  'Show carts on Station feed
        Case "Y"
            ckcXDSave(5).Value = vbChecked
        Case Else
            ckcXDSave(5).Value = vbUnchecked
    End Select
    Select Case tmVff.sXDSSaveNAS  'Show carts on Station feed
        Case "Y"
            ckcXDSave(6).Value = vbChecked
        Case Else
            ckcXDSave(6).Value = vbUnchecked
    End Select
    
    If (smNoMissedReason = "Y") Or (smAllowMGSpots = "N") Then
        udcVehOptTabs.AffLog(0) = vbUnchecked
        udcVehOptTabs.AllowMGEnabled = False
    Else
        Select Case tmVff.sMGsOnWeb
            Case "Y"
                udcVehOptTabs.AffLog(0) = vbChecked
            Case Else
                udcVehOptTabs.AffLog(0) = vbUnchecked
        End Select
    End If
    If smAllowReplSpots = "N" Then
        udcVehOptTabs.AffLog(1) = vbUnchecked
        udcVehOptTabs.AllowReplEnabled = False
    Else
        Select Case tmVff.sReplacementOnWeb
            Case "Y"
                udcVehOptTabs.AffLog(1) = vbChecked
            Case Else
                udcVehOptTabs.AffLog(1) = vbUnchecked
        End Select
    End If
    
    Select Case tmVff.sHideCommOnWeb
        Case "Y"
            udcVehOptTabs.AffLog(2) = vbChecked
        Case Else
            udcVehOptTabs.AffLog(2) = vbUnchecked
    End Select
    
    edcExport(0).Text = Trim$(tmVff.sAirWavePrgID)
    edcExport(1).Text = Trim$(tmVff.sIPumpEventTypeOV)
    Select Case tmVff.sExportAirWave
        Case "Y"
            ckcIFExport(7).Value = vbChecked
        Case Else
            ckcIFExport(7).Value = vbUnchecked
    End Select
    Select Case tmVff.sExportAudio
        Case "Y"
            ckcIFExport(8).Value = vbChecked
        Case Else
            ckcIFExport(8).Value = vbUnchecked
    End Select
    Select Case tmVff.sExportMP2
        Case "Y"
            ckcIFExport(9).Value = vbChecked
        Case Else
            ckcIFExport(9).Value = vbUnchecked
    End Select
    Select Case tmVff.sExportCnCSpot
        Case "Y"
            ckcIFExport(10).Value = vbChecked
        Case Else
            ckcIFExport(10).Value = vbUnchecked
    End Select
    Select Case tmVff.sExportEnco
        Case "Y"
            ckcIFExport(11).Value = vbChecked
        Case Else
            ckcIFExport(11).Value = vbUnchecked
    End Select
    Select Case tmVff.sExportNYESPN
        Case "Y"
            ckcIFExport(12).Value = vbChecked
        Case Else
            ckcIFExport(12).Value = vbUnchecked
    End Select
    Select Case tmVff.sPledgeVsAir
        Case "Y"
            ckcIFExport(13).Value = vbChecked
        Case Else
            ckcIFExport(13).Value = vbUnchecked
    End Select
    Select Case tmVff.sExportEncoESPN
        Case "Y"
            ckcIFExport(14).Value = vbChecked
        Case Else
            ckcIFExport(14).Value = vbUnchecked
    End Select
    Select Case tmVff.sExportCnCNetInv
        Case "Y"
            ckcIFExport(15).Value = vbChecked
        Case Else
            ckcIFExport(15).Value = vbUnchecked
    End Select
    Select Case tmVff.sExportMatrix
        Case "Y"
            ckcIFExport(16).Value = vbChecked
        Case Else
            ckcIFExport(16).Value = vbUnchecked
    End Select
    Select Case tmVff.sExportTableau
        Case "Y"
            ckcIFExport(18).Value = vbChecked
        Case Else
            ckcIFExport(18).Value = vbUnchecked
    End Select
    Select Case tmVff.sExportJelli
        Case "Y"
            ckcIFExport(17).Value = vbChecked
        Case Else
            ckcIFExport(17).Value = vbUnchecked
    End Select
    Select Case tmVff.sExportIPump
        Case "Y"
            ckcAffExport(3).Value = vbChecked
        Case Else
            ckcAffExport(3).Value = vbUnchecked
    End Select
    Select Case tmVff.sStationComp
        Case "Y"
            ckcAffExport(4).Value = vbChecked
        Case Else
            ckcAffExport(4).Value = vbUnchecked
    End Select
    If (tmVef.sType = "R") Or (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "N") Then
        If tmVff.sOnInsertions = "N" Then
            ckcBarter(0).Value = vbUnchecked
        Else
            ckcBarter(0).Value = vbChecked
        End If
    Else
        ckcBarter(0).Value = vbUnchecked
    End If
    'TTP 9992 - Custom Rev Export
    Select Case tmVff.sExportCustom
        Case "Y"
            ckcIFExport(20).Value = vbChecked
        Case Else
            ckcIFExport(20).Value = vbUnchecked
    End Select

    edcSchedule(5).Text = ""
    If (tmVef.sType = "A") Then
        ckcLog(0).Visible = True
        If tmVff.sHonorZeroUnits = "Y" Then
            ckcLog(0).Value = vbChecked
        Else
            ckcLog(0).Value = vbUnchecked
        End If
        If tmVff.iConflictWinLen > 0 Then
            edcSchedule(5).Text = tmVff.iConflictWinLen
        End If
    Else
        ckcLog(0).Visible = False
        ckcLog(0).Value = vbUnchecked
    End If

    If (tmVef.sType <> "S") Then
        ckcLog(1).Enabled = True
        If tmVff.sHideCommOnLog = "Y" Then
            ckcLog(1).Value = vbChecked
        Else
            ckcLog(1).Value = vbUnchecked
        End If
    Else
        ckcLog(1).Enabled = False
        ckcLog(1).Value = vbUnchecked
    End If

    edcBarter(5).Text = ""
    lacBarter(8).Enabled = False
    edcBarter(5).Enabled = False
    If (tmVef.sType = "C") Or (tmVef.sType = "S") Then
        If tmVff.sPostLogSource = "S" Then
            rbcBarterMethod(6).Value = True
            edcBarter(5).Text = Trim$(tmVff.sStationPassword)
            lacBarter(8).Enabled = True
            edcBarter(5).Enabled = True
        Else
            rbcBarterMethod(5).Value = True
        End If
    Else
        rbcBarterMethod(5).Value = True
    End If


    
    tmCefSrchKey.lCode = tmVff.lBBOpenCefCode
    If tmCefSrchKey.lCode <> 0 Then
        tmCef.sComment = ""
        imCefRecLen = Len(tmCef)    '1009
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smOpenBB = gStripChr0(tmCef.sComment)
        Else
            smOpenBB = ""
        End If
    Else
        smOpenBB = ""
    End If
    edcBB(0).Text = smOpenBB
    tmCefSrchKey.lCode = tmVff.lBBCloseCefCode
    If tmCefSrchKey.lCode <> 0 Then
        tmCef.sComment = ""
        imCefRecLen = Len(tmCef)    '1009
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smCloseBB = gStripChr0(tmCef.sComment)
        Else
            smCloseBB = ""
        End If
    Else
        smCloseBB = ""
    End If
    edcBB(1).Text = smCloseBB
    
    imVffChg = False
    'Producer
    pbcProducer.Cls
    ilFound = False
    For ilLoop = 0 To UBound(tmProducerCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
        slNameCode = tmProducerCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If ilRet = CP_MSG_NONE Then
            If Val(slCode) = tlVpf.iProducerArfCode Then
                ilFound = True
                Exit For
            End If
        Else
            ilFound = False
        End If
    Next ilLoop
    If ilFound Then
        lbcProducer.ListIndex = ilLoop + 2
    Else
        lbcProducer.ListIndex = 1   '[None]
    End If
    If lbcProducer.ListIndex <= 0 Then
        slStr = ""
    Else
        slStr = lbcProducer.List(lbcProducer.ListIndex)
    End If
    gSetShow pbcProducer, slStr, tmPCtrls(PRODUCERINDEX)
    If tgSpf.sGUseAffSys = "Y" Then
'        ilFound = False
'        For ilLoop = 0 To UBound(tmContentProviderCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
'            slNameCode = tmContentProviderCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
'            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'            If ilRet = CP_MSG_NONE Then
'                If Val(slCode) = tlVpf.iProviderArfCode Then
'                    ilFound = True
'                    Exit For
'                End If
'            Else
'                ilFound = False
'            End If
'        Next ilLoop
'        If ilFound Then
'            lbcContentProvider.ListIndex = ilLoop + 2
'        Else
'            lbcContentProvider.ListIndex = 1   '[None]
'        End If
'        If lbcContentProvider.ListIndex <= 0 Then
'            slStr = ""
'        Else
'            slStr = lbcContentProvider.List(lbcContentProvider.ListIndex)
'        End If
'        gSetShow pbcProducer, slStr, tmPCtrls(CONTENTPROVIDERINDEX)
        ilFound = False
'        For ilLoop = 0 To UBound(tmProducerOrProviderCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
'            slNameCode = tmProducerOrProviderCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
        For ilLoop = 0 To UBound(tmContentProviderCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
            slNameCode = tmContentProviderCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If ilRet = CP_MSG_NONE Then
                If Val(slCode) = tlVpf.iProgProvArfCode Then
                    ilFound = True
                    Exit For
                End If
            Else
                ilFound = False
            End If
        Next ilLoop
        If ilFound Then
            lbcExpProgAudio.ListIndex = ilLoop + 2
        Else
            lbcExpProgAudio.ListIndex = 1  '[None]
        End If
        If lbcExpProgAudio.ListIndex < 0 Then
            slStr = ""
        Else
            slStr = lbcExpProgAudio.List(lbcExpProgAudio.ListIndex)
        End If
        gSetShow pbcProducer, slStr, tmPCtrls(EXPPROGAUDIOINDEX)
        ilFound = False
'        For ilLoop = 0 To UBound(tmProducerOrProviderCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
'            slNameCode = tmProducerOrProviderCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
        For ilLoop = 0 To UBound(tmContentProviderCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
            slNameCode = tmContentProviderCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If ilRet = CP_MSG_NONE Then
                If Val(slCode) = tlVpf.iCommProvArfCode Then
                    ilFound = True
                    Exit For
                End If
            Else
                ilFound = False
            End If
        Next ilLoop
        If ilFound Then
            lbcExpCommAudio.ListIndex = ilLoop + 2
        Else
            lbcExpCommAudio.ListIndex = 1   '[None]
        End If
        If lbcExpCommAudio.ListIndex < 0 Then
            slStr = ""
        Else
            slStr = lbcExpCommAudio.List(lbcExpCommAudio.ListIndex)
        End If
        gSetShow pbcProducer, slStr, tmPCtrls(EXPCOMMAUDIOINDEX)
    Else
        'lbcContentProvider.ListIndex = 1
        lbcExpProgAudio.ListIndex = 0
        lbcExpCommAudio.ListIndex = 0
    End If
    pbcProducer_Paint
    'Interface
    For ilLoop = 1 To 5 Step 1
        If tlVpf.iESTEndTime(ilLoop - 1) > 0 Then
            edcIFEST(ilLoop - 1).Text = gFormatTimeLong(60 * CLng(tlVpf.iESTEndTime(ilLoop - 1)), "A", "2")
        Else
            edcIFEST(ilLoop - 1).Text = ""
        End If
        If tlVpf.iCSTEndTime(ilLoop - 1) > 0 Then
            edcIFCST(ilLoop - 1).Text = gFormatTimeLong(60 * CLng(tlVpf.iCSTEndTime(ilLoop - 1)), "A", "2")
        Else
            edcIFCST(ilLoop - 1).Text = ""
        End If
        If tlVpf.iMSTEndTime(ilLoop - 1) > 0 Then
            edcIFMST(ilLoop - 1).Text = gFormatTimeLong(60 * CLng(tlVpf.iMSTEndTime(ilLoop - 1)), "A", "2")
        Else
            edcIFMST(ilLoop - 1).Text = ""
        End If
        If tlVpf.iPSTEndTime(ilLoop - 1) > 0 Then
            edcIFPST(ilLoop - 1).Text = gFormatTimeLong(60 * CLng(tlVpf.iPSTEndTime(ilLoop - 1)), "A", "2")
        Else
            edcIFPST(ilLoop - 1).Text = ""
        End If
    Next ilLoop
    For ilLoop = 1 To 4 Step 1
        edcIFZone(ilLoop - 1).Text = Trim$(tlVpf.sMapZone(ilLoop - 1))
        edcIFProgCode(ilLoop - 1).Text = Trim$(tlVpf.sMapProgCode(ilLoop - 1))
        If tlVpf.iMapDPNo(ilLoop - 1) > 0 Then
            edcIFDPNo(ilLoop - 1).Text = Trim$(Str$(tlVpf.iMapDPNo(ilLoop - 1)))
        Else
            edcIFDPNo(ilLoop - 1).Text = ""
        End If
    Next ilLoop
    Select Case tlVpf.sExpHiClear  'Clearance Spots
        Case "Y"
            ckcIFExport(0).Value = vbChecked
        Case Else
            ckcIFExport(0).Value = vbUnchecked
    End Select
    Select Case tlVpf.sExpHiDallas  'Dallas Feed
        Case "Y"
            ckcIFExport(1).Value = vbChecked
        Case Else
            ckcIFExport(1).Value = vbUnchecked
    End Select
    Select Case tlVpf.sExpHiPhoenix  'Phoenix Feed
        Case "Y"
            ckcIFExport(2).Value = vbChecked
        Case Else
            ckcIFExport(2).Value = vbUnchecked
    End Select
    Select Case tlVpf.sExpHiNY  'New York EAS Feed
        Case "Y"
            ckcIFExport(3).Value = vbChecked
        Case Else
            ckcIFExport(3).Value = vbUnchecked
    End Select
    Select Case tlVpf.sExpHiNYISCI  'New York ASP Feed
        Case "Y"
            ckcIFExport(6).Value = vbChecked
        Case Else
            ckcIFExport(6).Value = vbUnchecked
    End Select
    Select Case tlVpf.sExpHiCmmlChg  'Commercial Change
        Case "Y"
            ckcIFExport(4).Value = vbChecked
        Case Else
            ckcIFExport(4).Value = vbUnchecked
    End Select
    Select Case tlVpf.sExpBkCpyCart  'Show Cart on Bulk Feed
        Case "Y"
            rbcExpBkCpyCart(0).Value = True
        Case Else
            rbcExpBkCpyCart(1).Value = True
    End Select
    Select Case tlVpf.sBulkXFer  'Bulk Feed Cross Reference
        Case "Y"
            rbcIFBulk(0).Value = True
        Case Else
            rbcIFBulk(1).Value = True
    End Select
    Select Case tlVpf.sClearAsSell  'In Clearance treat as selling
        Case "Y"
            rbcIFSelling(0).Value = True
        Case Else
            rbcIFSelling(1).Value = True
    End Select
    Select Case tlVpf.sClearChgTime  'In Clearance change times
        Case "Y"
            rbcIFTime(0).Value = True
        Case Else
            rbcIFTime(1).Value = True
    End Select
    'gUnpackDate tlVpf.iACal(0), tlVpf.iACal(1), slStr    'Last log date
    'lacAInvDate(0).Caption = slStr
    'gUnpackDate tlVpf.iAStd(0), tlVpf.iAStd(1), slStr    'Last preliminary date
    'lacAInvDate(1).Caption = slStr
    edcIFGroupNo.Text = Trim$(tlVpf.sGGroupNo)  'Bulk Feed Group Number
    edcExpCntrVehNo.Text = Trim$(tlVpf.sExpVehNo)  'Contract Export Vehicle Number
    edcARBCode.Text = Trim$(tlVpf.sARBCode) 'Arbitron Code
    edcRadarCode.Text = Trim$(tlVpf.sRadarCode) 'Radar Code
    edcStnFdCode.Text = Trim$(tlVpf.sStnFdCode)  'Station Feed code
    Select Case tlVpf.sStnFdCart  'Show carts on Station feed
        Case "Y"
            ckcStnFdInfo(0).Value = vbChecked
        Case Else
            ckcStnFdInfo(0).Value = vbUnchecked
    End Select
    Select Case tlVpf.sStnFdXRef  'Show carts in cross reference from Station feed
        Case "N"
            ckcStnFdInfo(1).Value = vbUnchecked
        Case Else
            ckcStnFdInfo(1).Value = vbChecked
    End Select
    edcEDASWindow.Text = Trim$(Str$(tlVpf.lEDASWindow))
    Select Case tlVpf.sKCGenRot
        Case "N"
            ckcKCGenRot.Value = vbUnchecked
        Case Else
            ckcKCGenRot.Value = vbChecked
    End Select
    Select Case tlVpf.sExportSQL
        Case "Y"
            ckcExportSQL.Value = vbChecked
        Case Else
            ckcExportSQL.Value = vbUnchecked
    End Select
    ilValue = Asc(tgSpf.sUsingFeatures2)
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G")) And ((ilValue And BARTER) = BARTER) Then
    If (tmVef.sType = "R") And ((ilValue And BARTER) = BARTER) Then
'        edcBarter(0).Text = Trim$(Str$(tlVpf.lDefAcqCost))
'        edcBarter(1).Text = Trim$(Str$(tlVpf.lActAcqCost))
        edcBarterMethod(0).Text = ""
        edcBarterMethod(1).Text = ""
        edcBarterMethod(2).Text = ""
        edcBarterMethod(3).Text = ""
        edcBarterMethod(0).Enabled = False
        edcBarterMethod(1).Enabled = False
        edcBarterMethod(2).Enabled = False
        edcBarterMethod(3).Enabled = False
        rbcBarterMethod(0).Value = False
        rbcBarterMethod(1).Value = False
        rbcBarterMethod(2).Value = False
        rbcBarterMethod(3).Value = False
        rbcBarterMethod(4).Value = False
        cbcPerPeriod(0).ListIndex = -1
        cbcPerPeriod(1).ListIndex = -1
        cbcPerPeriod(2).ListIndex = -1
'        Select Case Trim$(tlVpf.sBarterMethod)
'            Case "A"
'                rbcBarterMethod(0).Value = True
'            Case "M"
'                rbcBarterMethod(1).Value = True
'                edcBarterMethod(0).Text = Trim$(Str$(tlVpf.iBarterThreshold))
'                edcBarterMethod(0).Enabled = True
'            Case "U"
'                rbcBarterMethod(2).Value = True
'                edcBarterMethod(1).Text = Trim$(Str$(tlVpf.iBarterThreshold))
'                edcBarterMethod(1).Enabled = True
'            Case "X"
'                rbcBarterMethod(3).Value = True
'                edcBarterMethod(2).Text = Trim$(Str$(tlVpf.iBarterXFree))
'                edcBarterMethod(3).Text = Trim$(Str$(tlVpf.iBarterYSold))
'                edcBarterMethod(2).Enabled = True
'                edcBarterMethod(3).Enabled = True
'            Case Else
'                rbcBarterMethod(4).Value = True
'        End Select
    ElseIf ((tmVef.sType = "C") Or (tmVef.sType = "S")) And ((ilValue And BARTER) = BARTER) Then
    End If
    '8132
    If (tmVef.sType = "R") Or tmVef.sType = "C" Then
        If tmVff.sOnXMLInsertion = "W" Then
            rbcBarterMethod(STATIONXMLWIDEORBIT).Value = True
        '8032
        ElseIf tmVff.sOnXMLInsertion = "M" Then
            rbcBarterMethod(STATIONXMLMARKETRON).Value = True
        Else
            rbcBarterMethod(STATIONXMLNONE).Value = True
        End If
    End If
    ilValue = Asc(tgSpf.sUsingFeatures2)
    rbcSplitCopy(0).Value = False
    rbcSplitCopy(1).Value = False
    '6/19/07:  Jim:  Allow defineition for Conventional, Airing and Game.  Always set on fro Packages and Selling
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or ((tmVef.sType = "A") And (rbcLCopyOnAir(0).Value)) Or (tmVef.sType = "G") Or (tmVef.sType = "P")) And ((ilValue And SPLITCOPY) = SPLITCOPY) Then
    '5/11/11: Allow selling to be set to No
    'If ((tmVef.sType = "C") Or (tmVef.sType = "A") Or (tmVef.sType = "G")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
    If ((tmVef.sType = "C") Or (tmVef.sType = "A") Or (tmVef.sType = "G") Or (tmVef.sType = "S")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
        If tlVpf.sAllowSplitCopy = "Y" Then
            rbcSplitCopy(0).Value = True
        Else
            rbcSplitCopy(1).Value = True
        End If
    'ElseIf ((tmVef.sType = "S") Or (tmVef.sType = "P")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
    ElseIf ((tmVef.sType = "P")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
        rbcSplitCopy(0).Value = True
    Else
        rbcSplitCopy(1).Value = True
    End If
    'If ((tmVef.sType = "C") And (tmVef.iVefCode = 0)) Or (tmVef.sType = "L") Or (tmVef.sType = "A") Or (tmVef.sType = "G") Then
        If tlVpf.sShowRateOnInsert = "Y" Then
            rbcShowRateOnInsertion(0).Value = True
        Else
            rbcShowRateOnInsertion(1).Value = True
        End If
    'End If
    ckcAffExport(0).Value = vbUnchecked
    If ((Asc(tgSpf.sUsingFeatures7) And WEGENEREXPORT) = WEGENEREXPORT) Then
        If tlVpf.sWegenerExport = "Y" Then
            ckcAffExport(0).Value = vbChecked
        End If
    End If
    If tmVef.sType = "G" Then
        If tlVff.sMoveSportToNon = "Y" Then
            ckcSch(0).Value = vbChecked
        Else
            ckcSch(0).Value = vbUnchecked
        End If
        If tlVff.sMoveSportToSport = "Y" Then
            ckcSch(1).Value = vbChecked
        Else
            ckcSch(1).Value = vbUnchecked
        End If
        If tlVff.sMoveNonToSport = "Y" Then
            ckcSch(2).Value = vbChecked
        Else
            ckcSch(2).Value = vbUnchecked
        End If
        If tlVff.sPledgeByEvent = "Y" Then
            ckcSch(3).Value = vbChecked
            '9213 temp block choosing both MGs and 'Pledge by event'.  Pledge takes precedence
            udcVehOptTabs.AffLog(0) = vbUnchecked
'    End If
        Else
            ckcSch(3).Value = vbUnchecked
        End If
        If tlVff.iPledgeClearance > 0 Then
            edcSchedule(4).Text = tlVff.iPledgeClearance
        Else
            edcSchedule(4).Text = ""
        End If
        ckcSch(0).Enabled = True
        ckcSch(1).Enabled = True
        ckcSch(2).Enabled = True
        ckcSch(3).Enabled = True
        edcSchedule(4).Enabled = True
    Else
        ckcSch(0).Value = vbUnchecked
        ckcSch(1).Value = vbUnchecked
        ckcSch(2).Value = vbUnchecked
        ckcSch(3).Value = vbUnchecked
        ckcSch(0).Enabled = False
        ckcSch(1).Enabled = False
        ckcSch(2).Enabled = False
        ckcSch(3).Enabled = False
        edcSchedule(4).Enabled = False
        edcSchedule(4).Text = ""
    End If
    If tmVef.iVefCode > 0 Then
        For ilLoop = 0 To 5 Step 1
            rbcMerge(ilLoop).Value = False
            rbcMerge(ilLoop).Enabled = True
        Next ilLoop
        If tlVff.sMergeTraffic = "S" Then
            rbcMerge(1).Value = True
        Else
            rbcMerge(0).Value = True
        End If
        If tlVff.sMergeAffiliate = "S" Then
            rbcMerge(3).Value = True
        Else
            rbcMerge(2).Value = True
        End If
        If tlVff.sMergeWeb = "S" Then
            rbcMerge(5).Value = True
        Else
            rbcMerge(4).Value = True
        End If
    Else
        For ilLoop = 0 To 5 Step 1
            rbcMerge(ilLoop).Value = False
            rbcMerge(ilLoop).Enabled = False
        Next ilLoop
    End If
    edcLog(3).Text = Trim$(tlVff.sWebName)
    'ckcAffExport(1).Value = vbUnchecked
    'If ((Asc(tgSpf.sUsingFeatures7) And OLAEXPORT) = OLAEXPORT) Then
    '    If tlVpf.sOLAExport = "Y" Then
    '        ckcAffExport(1).Value = vbChecked
    '    End If
    'End If
    cbcMedia.ListIndex = 0
    If tlVff.iMcfCode > 0 Then
        For ilLoop = 0 To UBound(tmMediaCode) - 1 Step 1  'lbcMediaCode.ListCount - 1 Step 1
            slNameCode = tmMediaCode(ilLoop).sKey    'lbcMediaCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If ilRet = CP_MSG_NONE Then
                If tlVff.iMcfCode = Val(slCode) Then
                    cbcMedia.ListIndex = ilLoop + 1
                    Exit For
                End If
            End If
        Next ilLoop
    End If
    If tmVef.sExportRAB = "Y" Then
        ckcIFExport(19).Value = vbChecked
    Else
        ckcIFExport(19).Value = vbUnchecked
    End If
    '10894 removed 10050 podcast
    'mEnableGeneralMedium tlVpf.sGMedium, tmVef.iCode
    If rbcGMedium(PODCASTRBC).Value Then
        '10981
        'mGetPodcastInfo
        mVendorSetAndEnableInfoOriginal
        'after SetAndEnable because I want imCurrentVendorInfoIndex already set
        mVendorLegacyAdjustBoostr

    Else
        mVendorEnableOptions
    End If
    imIgnoreClickEvent = False
    udcVehOptTabs.Action 7, 1
    imVefChg = False
    imVffChg = False
    
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRUsrAreaPaint                  *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint user collection area     *
'*                                                     *
'*******************************************************
Private Sub mMPromoPaint(tlVpf As VPF)
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim flX As Single
    Dim flY As Single
    flX = imPromoX + fgBoxInsetX
    flY = imPromoY
    For ilCol = 0 To 2 Step 1
        For ilRow = 1 To 24 Step 1
            gPaintArea pbcPromo, flX, flY + 15, imPromoW - fgBoxInsetX - 15, imPromoH - 45, WHITE
            pbcPromo.CurrentX = flX
            pbcPromo.CurrentY = flY - 15
            If ilCol = 0 Then
                pbcPromo.Print Trim$(Str$(tlVpf.iMMFPromo(ilRow - 1)))
            ElseIf ilCol = 1 Then
                pbcPromo.Print Trim$(Str$(tlVpf.iMSAPromo(ilRow - 1)))
            Else
                pbcPromo.Print Trim$(Str$(tlVpf.iMSUPromo(ilRow - 1)))
            End If
            flY = flY + imPromoH - 15
        Next ilRow
        flX = flX + imPromoW + 15
        flY = imPromoY
    Next ilCol
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMPsaPaint                      *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint system area              *
'*                                                     *
'*******************************************************
Private Sub mMPsaPaint(tlVpf As VPF)
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim flX As Single
    Dim flY As Single
    flX = imPsaX + fgBoxInsetX
    flY = imPsaY
    For ilCol = 0 To 2 Step 1
        For ilRow = 1 To 24 Step 1
            gPaintArea pbcPSA, flX, flY + 15, imPsaW - fgBoxInsetX - 15, imPsaH - 45, WHITE
            pbcPSA.CurrentX = flX
            pbcPSA.CurrentY = flY - 15
            If ilCol = 0 Then
                pbcPSA.Print Trim$(Str$(tlVpf.iMMFPSA(ilRow - 1)))
            ElseIf ilCol = 1 Then
                pbcPSA.Print Trim$(Str$(tlVpf.iMSAPSA(ilRow - 1)))
            Else
                pbcPSA.Print Trim$(Str$(tlVpf.iMSUPSA(ilRow - 1)))
            End If
            flY = flY + imPsaH - 15
        Next ilRow
        flX = flX + imPsaW + 15
        flY = imPsaY
    Next ilCol
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim llFilter As Long

    'ilFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHLOG + VEHVIRTUAL + VEHLOGVEHICLE + ACTIVEVEH + DORMANTVEH
    If igVpfType <> 1 Then
        'llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHLOG + VEHVIRTUAL + VEHLOGVEHICLE + VEHSIMUL + VEHREP_WO_CLUSTER + ACTIVEVEH + DORMANTVEH
        'Add rep vehicles so that spot lengths can be defined
        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHLOG + VEHVIRTUAL + VEHLOGVEHICLE + VEHSIMUL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHSPORT + VEHIMPORTAFFILIATESPOTS + VEHNTR + ACTIVEVEH + DORMANTVEH
    Else
        llFilter = VEHPACKAGE + ACTIVEVEH + DORMANTVEH
    End If
    'ilRet = gPopUserVehicleBox(VehOpt, ilFilter, cbcSelect, Traffic!lbcVehicle)
    ilRet = gPopUserVehicleBox(VehOpt, llFilter, cbcSelect, tmVehicle(), smVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", VehOpt
        On Error GoTo 0
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub mRemoveFocus()
    If plcGeneral.Visible Then
        edcGTZDropDown_LostFocus
        pbcGTZToggle.Visible = False
        lbcFeed.Visible = False
        cmcGTZDropDown.Visible = False
        edcGTZDropDown.Visible = False
        imGTZBoxNo = -1
    End If
    If plcSales.Visible Then
        edcSSpotLG.Visible = False
        imSSpotLenBoxNo = -1
    End If
    If plcSchedule(0).Visible Then
        edcLevelPrice.Visible = False
        imLevelPriceBoxNo = -1
    End If
    If plcPSAPromo.Visible Then
        edcMPSA.Visible = False
        imMPSABoxNo = -1
        edcMPromo.Visible = False
        imMPromoBoxNo = -1
    End If
    If plcVirtual.Visible Then
        mVirtSetShow imVirtBoxNo
        imVirtBoxNo = -1
        imVirtRowNo = -1
    End If
    If plcSchedule(0).Visible Then
        mSSetShow imSBoxNo
        imSBoxNo = -1
        imSRowNo = -1
    End If
    If plcLog.Visible Then
        mLSetShow imLBoxNo
        imLBoxNo = -1
        imLRowNo = -1
    End If
    If plcProducer.Visible Then
        mPSetShow imPBoxNo
        imPBoxNo = -1
    End If
    If plcBarter.Visible Then
        edcBarter(3).Visible = False
        imAcqCostBoxNo = -1
        edcBarter(4).Visible = False
        imAcqIndexBoxNo = -1
    End If
    If plcParticipant.Visible Then
        mPartSetShow
        lmPartEnableRow = -1
        lmPartEnableCol = -1
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLogPop                         *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Log, CP and Other     *
'*                      list boxes                     *
'*                                                     *
'*******************************************************
Private Sub mRnfPop()
    Dim ilLoop As Integer
    Dim ilLen As Integer
    Dim slChar As String
    Dim ilOk As Integer
    Dim ilValue As Integer
    Dim slName As String
    gObtainRNF hmRnf
    cbcLog(0).Clear
    cbcLog(1).Clear
    cbcLog(2).Clear
    cbcSvLog(0).Clear
    cbcSvLog(1).Clear
    cbcSvLog(2).Clear
    For ilLoop = 0 To UBound(tgRnfList) - 1 Step 1
        If tgRnfList(ilLoop).tRnf.sType = "R" Then
            ilLen = Len(Trim$(tgRnfList(ilLoop).tRnf.sName))
            slChar = UCase$(Left$(tgRnfList(ilLoop).tRnf.sName, 1))
            If (ilLen > 1) And (ilLen < 4) Then
                ilOk = False
                If slChar = "L" Then
                    ilOk = True
                ElseIf slChar = "C" Then
                    ilOk = True
                'ElseIf slChar = "O" Then
                '    ilOk = True
                End If
                If ilOk Then
                    ilValue = Asc(Mid$(tgRnfList(ilLoop).tRnf.sName, 2, 1))
                    If (ilValue < Asc("0")) Or (ilValue > Asc("9")) Then
                        ilOk = False
                    End If
                End If
                If ilOk Then
                    slName = UCase$(Trim$(tgRnfList(ilLoop).tRnf.sName))
                    If slChar = "L" Then
                        cbcLog(0).AddItem slName
                        cbcLog(2).AddItem slName
                        cbcSvLog(0).AddItem slName
                        cbcSvLog(2).AddItem slName
                    ElseIf slChar = "C" Then
                        cbcLog(1).AddItem slName
                        cbcSvLog(1).AddItem slName
                        cbcLog(2).AddItem slName        '2-16-01 Allow "C" in addition to "L" types
                        cbcSvLog(2).AddItem slName
                    'ElseIf slChar = "O" Then
                    '    lbcOther.AddItem slName
                    End If
                End If
            End If
        End If
    Next ilLoop
    cbcLog(0).AddItem "[None]", 0
    cbcLog(1).AddItem "[None]", 0
    cbcLog(2).AddItem "[None]", 0
    cbcSvLog(0).AddItem "[None]", 0
    cbcSvLog(1).AddItem "[None]", 0
    cbcSvLog(2).AddItem "[None]", 0
    On Error GoTo 0
    Exit Sub

    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set clear and set altered flag *
'*                                                     *
'*******************************************************
Private Sub mSetChg(tlVpf As VPF, ilVpfIndex As Integer, ilIgnoreChg As Integer, ilAltered As Integer)
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slTime As String
    Dim slDate As String
    Dim ilValue As Integer

    ilAltered = False
    If ilIgnoreChg Then    'Bypass changing during move record to controls
        Exit Sub
    End If
    If imVffChg Then
        ilAltered = True
        Exit Sub
    End If
    If imVbfChg Then
        ilAltered = True
        Exit Sub
    End If
    If imVefChg Then
        ilAltered = True
        Exit Sub
    End If
    If imPifChg Then
        ilAltered = True
        Exit Sub
    End If
    If imPartMissAndReq Then
        ilAltered = True
        Exit Sub
    End If
    If imLogAltered Or imInventoryAltered Then
        ilAltered = True
        Exit Sub
    End If
    If imLevelAltered Then
        ilAltered = True
        Exit Sub
    End If
    If imProducerAltered Then
        ilAltered = True
        Exit Sub
    End If
    If imGreatPlainsAltered Then
        ilAltered = True
        Exit Sub
    End If
    ' tlToRec.iVEFKCode = tlFromRec.iVEFKCode  'Input Vehicle
    gUnpackTime tlVpf.iGTime(0), tlVpf.iGTime(1), "A", "1", slStr
    slStr = gConvertTime(slStr)
    slTime = edcGen(Signon).Text 'edcGSignOn.Text
    slTime = gConvertTime(slTime)
    If (Hour(slStr) <> Hour(slTime)) Or (Minute(slStr) <> Minute(slTime)) Or (Second(slStr) <> Second(slTime)) Then
        ilAltered = True
        Exit Sub
    End If
    Select Case tlVpf.sGMedium
        Case "R"
            If rbcGMedium(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "T"
            If rbcGMedium(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "N"
            If rbcGMedium(2).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "V"
            If rbcGMedium(3).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "C"
            If rbcGMedium(4).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "S"
            If rbcGMedium(5).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "P"            '1-16-14 podcast
            If rbcGMedium(PODCASTRBC).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    Select Case tlVpf.sEmbeddedOrROS
        Case "E"
            If rbcGMedium(7).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "R"
            If rbcGMedium(8).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Then
        'Select Case tlVpf.sFeedLogOrder
        '    Case "I"
        '        If rbcRemoteExport(0).Value <> True Then
        '            ilAltered = True
        '            Exit Sub
        '        End If
        '    Case "A"
        '        If rbcRemoteExport(1).Value <> True Then
        '            ilAltered = True
        '            Exit Sub
        '        End If
        '    Case Else
        '        If rbcRemoteExport(2).Value <> True Then
        '            ilAltered = True
        '            Exit Sub
        '        End If
        'End Select
        If (Asc(tgSpf.sUsingFeatures5) And REMOTEEXPORT) = REMOTEEXPORT Then
            If (Asc(tlVpf.sUsingFeatures1) And EXPORTINSERTION) = EXPORTINSERTION Then
                If rbcRemoteExport(0).Value <> True Then
                    ilAltered = True
                End If
            ElseIf (Asc(tlVpf.sUsingFeatures1) And EXPORTLOG) = EXPORTLOG Then
                If rbcRemoteExport(1).Value <> True Then
                    ilAltered = True
                End If
            Else
                If ((Asc(tlVpf.sUsingFeatures1) And EXPORTINSERTION) <> EXPORTINSERTION) And ((Asc(tlVpf.sUsingFeatures1) And EXPORTLOG) <> EXPORTLOG) Then
                    If rbcRemoteExport(2).Value <> True Then
                        ilAltered = True
                    End If
                End If
            End If
        End If
        If (Asc(tgSpf.sUsingFeatures5) And REMOTEIMPORT) = REMOTEIMPORT Then
            If (Asc(tlVpf.sUsingFeatures1) And IMPORTINSERTION) = IMPORTINSERTION Then
                If rbcRemoteImport(0).Value <> True Then
                    ilAltered = True
                End If
            ElseIf (Asc(tlVpf.sUsingFeatures1) And IMPORTAFFILIATESPOTS) = IMPORTAFFILIATESPOTS Then
                If rbcRemoteImport(1).Value <> True Then
                    ilAltered = True
                End If
            Else
                If ((Asc(tlVpf.sUsingFeatures1) And IMPORTINSERTION) <> IMPORTINSERTION) And ((Asc(tlVpf.sUsingFeatures1) And IMPORTAFFILIATESPOTS) <> IMPORTAFFILIATESPOTS) Then
                    If rbcRemoteImport(2).Value <> True Then
                        ilAltered = True
                    End If
                End If
            End If
        End If
    End If
    If (Asc(tlVpf.sUsingFeatures1) And EXPORTISCIBYPLEDGE) = EXPORTISCIBYPLEDGE Then
        If ckcAffExport(2).Value = vbUnchecked Then
            ilAltered = True
        End If
    Else
        If ckcAffExport(2).Value = vbChecked Then
            ilAltered = True
        End If
    End If
    
    '12/24/15: Affiliate Log tab not required
    If (tmVef.sType <> "R") And (tmVef.sType <> "N") Then
        If (Asc(tlVpf.sUsingFeatures1) And SUPPRESSWEBLOG) = SUPPRESSWEBLOG Then
            If udcVehOptTabs.WebInfo(3) = vbUnchecked Then
                ilAltered = True
            End If
        Else
            If udcVehOptTabs.WebInfo(3) = vbChecked Then
                ilAltered = True
            End If
        End If
        If (Asc(tlVpf.sUsingFeatures1) And EXPORTPOSTEDTIMES) = EXPORTPOSTEDTIMES Then
            If udcVehOptTabs.WebInfo(4) = vbUnchecked Then
                ilAltered = True
            End If
        Else
            If udcVehOptTabs.WebInfo(4) = vbChecked Then
                ilAltered = True
            End If
        End If
    End If
    
    If (Asc(tlVpf.sUsingFeatures2) And XDSAPPLYMERGE) = XDSAPPLYMERGE Then
        If ckcXDSave(3).Value = vbUnchecked Then
            ilAltered = True
        End If
    Else
        If ckcXDSave(3).Value = vbChecked Then
            ilAltered = True
        End If
    End If

    If (Asc(tlVpf.sUsingFeatures2) And INVOICEAIRDATEWO) = INVOICEAIRDATEWO Then
        If rbcShowAirDate(1).Value = False Then
            ilAltered = True
        End If
    Else
        If rbcShowAirDate(1).Value = True Then
            ilAltered = True
        End If
    End If

    Select Case tlVpf.sMoveLLD  'Move between today and lld
        Case "Y"
            If rbcSMoveLLD(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcSMoveLLD(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    Select Case tlVpf.sBillSA  'Bill Selling or Airing
        Case "Y"
            If rbcBillSA(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If (rbcBillSA(1).Value <> True) Or (tlVpf.sBillSA <> "N") Then
                ilAltered = True
                Exit Sub
            End If
    End Select

    If imInventoryAltered = True Then       '7-21-05
        ilAltered = True
        Exit Sub
    End If

    If tlVpf.sUnsoldBlank = "N" Then
        If ckcUnsoldBlank.Value = vbChecked Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If ckcUnsoldBlank.Value = vbUnchecked Then
            ilAltered = True
            Exit Sub
        End If
    End If

    '12/24/15: Affiliate Log tab not required
    If (tmVef.sType <> "R") And (tmVef.sType <> "N") Then
        If tlVpf.sAvailNameOnWeb = "N" Then
            If udcVehOptTabs.WebInfo(0) = vbChecked Then
                ilAltered = True
                Exit Sub
            End If
        Else
            If udcVehOptTabs.WebInfo(0) = vbUnchecked Then
                ilAltered = True
                Exit Sub
            End If
        End If
        If (tlVpf.sWebLogFeedTime = "N") Or (Trim$(tlVpf.sWebLogFeedTime) = "") Then
            If udcVehOptTabs.WebInfo(1) = vbChecked Then
                ilAltered = True
                Exit Sub
            End If
        Else
            If udcVehOptTabs.WebInfo(1) = vbUnchecked Then
                ilAltered = True
                Exit Sub
            End If
        End If
        If (tlVpf.sWebLogSummary = "N") Or (Trim$(tlVpf.sWebLogSummary) = "") Then
            If udcVehOptTabs.WebInfo(2) = vbChecked Then
                ilAltered = True
                Exit Sub
            End If
        Else
            If udcVehOptTabs.WebInfo(2) = vbUnchecked Then
                ilAltered = True
                Exit Sub
            End If
        End If
        'slStr = Trim$(edcLog(0).Text)
        slStr = Trim$(udcVehOptTabs.FTPInfo())
        If StrComp(slStr, smFTP, vbTextCompare) <> 0 Then
            ilAltered = True
            Exit Sub
        End If
        slStr = Trim$(udcVehOptTabs.LiveWindow())
        If StrComp(slStr, smLiveWindow, vbTextCompare) <> 0 Then
            ilAltered = True
            Exit Sub
        End If
        '2/28/19: Add Cart on Web
        If udcVehOptTabs.WebInfo(5) = vbChecked Then
            slStr = "Y"
        Else
            slStr = "N"
        End If
        If StrComp(slStr, smCartOnWeb, vbTextCompare) <> 0 Then
            ilAltered = True
            Exit Sub
        End If
    End If
    slStr = Trim$(edcLog(1).Text)
    If StrComp(slStr, smAutoExpt, vbTextCompare) <> 0 Then
        ilAltered = True
        Exit Sub
    End If
    slStr = Trim$(edcLog(2).Text)
    If StrComp(slStr, smAutoImpt, vbTextCompare) <> 0 Then
        ilAltered = True
        Exit Sub
    End If
    ' tlToRec.iUrfGCode = tlFromRec.iUrfGCode     'Counterpoint
    'If edcGPast.Text <> Trim$(Str$(tlVpf.iGMoFull)) Then      'month retain unpacked
    '    ilAltered = True
    '    Exit Sub
    'End If
    'If edcGHistory.Text <> Trim$(Str$(tlVpf.iGMoPack)) Then   'months retain packed
    '    ilAltered = True
    '    Exit Sub
    'End If
    'Select Case tlVpf.sGPriceStat  'Spot Price statistics
    '    Case "Y"
    '        If rbcSPrice(0).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case Else
    '        If rbcSPrice(1).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    'End Select
    Select Case tlVpf.sOwnership
        Case "B"
            If rbcGMedium(10).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "C"
            If rbcGMedium(11).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "D"
            If rbcGMedium(12).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcGMedium(9).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    Select Case tlVpf.sGenLog  'Scripts
        Case "N"
            If rbcGenLog(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "L"
            If rbcGenLog(2).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "M"
            If rbcGenLog(3).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "A"
            If rbcGenLog(4).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcGenLog(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    Select Case tlVpf.sGGridRes  'Grid resolution
        Case "F"
            If rbcLGrid(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "H"
            If rbcLGrid(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcLGrid(2).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    'Select Case tlVpf.sGScript  'Scripts
    '    Case "Y"
    '        If rbcLScripts(0).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case Else
    '        If rbcLScripts(1).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    'End Select
    Select Case tlVpf.sExpHiCorp
        Case "N"
            If ckcIFExport(5).Value <> vbUnchecked Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If ckcIFExport(5).Value <> vbChecked Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    'If edcSAGroupNo.Text <> Trim$(Str$(tlVpf.iSAGroupNo)) Then   'Vehicle Group Number
    If tlVpf.iSAGroupNo = 0 Then
        If edcGen(SAGROUPNO).Text = "" Then
            slStr = ""
        Else
            slStr = Trim$(Str$(tlVpf.iSAGroupNo))
        End If
    Else
        slStr = Trim$(Str$(tlVpf.iSAGroupNo))
    End If

    If edcGen(SAGROUPNO).Text <> slStr Then   'Vehicle Group Number
        ilAltered = True
        Exit Sub
    End If

    If StrComp(Trim$(tlVpf.sAccruedRevenue), Trim$(edcGen(0).Text), vbTextCompare) <> 0 Then
        ilAltered = True
        Exit Sub
    End If

    If StrComp(Trim$(tlVpf.sAccruedTrade), Trim$(edcGen(1).Text), vbTextCompare) <> 0 Then
        ilAltered = True
        Exit Sub
    End If

    If StrComp(Trim$(tlVpf.sBilledRevenue), Trim$(edcGen(2).Text), vbTextCompare) <> 0 Then
        ilAltered = True
        Exit Sub
    End If

    If StrComp(Trim$(tlVpf.sBilledTrade), Trim$(edcGen(3).Text), vbTextCompare) <> 0 Then
        ilAltered = True
        Exit Sub
    End If

    slStr = Trim$(edcGen(EDICALLLETTERS).Text)  'Trim$(edcGEDI(0).Text)
    If StrComp(slStr, Trim$(tlVpf.sEDICallLetters), vbTextCompare) <> 0 Then
        ilAltered = True
        Exit Sub
    End If
    slStr = Trim$(edcGen(EDIBAND).Text)  'Trim$(edcGEDI(1).Text)
    If StrComp(slStr, Trim$(tlVpf.sEDIBand), vbTextCompare) <> 0 Then
        ilAltered = True
        Exit Sub
    End If
    For ilLoop = LBound(tlVpf.sGZone) To UBound(tlVpf.sGZone) Step 1     'Log Time zones
        If tlVpf.sGZone(ilLoop) <> tgVpf(ilVpfIndex).sGZone(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iGLocalAdj(ilLoop) <> tgVpf(ilVpfIndex).iGLocalAdj(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iGFeedAdj(ilLoop) <> tgVpf(ilVpfIndex).iGFeedAdj(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iGV1Z(ilLoop) <> tgVpf(ilVpfIndex).iGV1Z(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iGV2Z(ilLoop) <> tgVpf(ilVpfIndex).iGV2Z(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iGV3Z(ilLoop) <> tgVpf(ilVpfIndex).iGV3Z(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iGV4Z(ilLoop) <> tgVpf(ilVpfIndex).iGV4Z(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.sGCSVer(ilLoop) <> tgVpf(ilVpfIndex).sGCSVer(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.sGFed(ilLoop) <> tgVpf(ilVpfIndex).sGFed(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iGMnfNCode(ilLoop) <> tgVpf(ilVpfIndex).iGMnfNCode(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.sGBus(ilLoop) <> tgVpf(ilVpfIndex).sGBus(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.sGSked(ilLoop) <> tgVpf(ilVpfIndex).sGSked(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
    Next ilLoop
    Select Case tlVpf.sShowTime    'Agency commission fixed
        Case "S"
            If rbcShowAirTime(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "H"
            If rbcShowAirTime(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "D"
            If rbcShowAirTime(2).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "A"
            If rbcShowAirTime(3).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If (rbcShowAirTime(0).Value) Or (rbcShowAirTime(1).Value) Or (rbcShowAirTime(2).Value) Or (rbcShowAirTime(3).Value) Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    Select Case tlVpf.sSVarComm     'Agency commission fixed
        Case "Y"
            If rbcSCommission(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "N"
            If rbcSCommission(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    Select Case tlVpf.sAdvtSep  'Separate advertiser by time or break
        Case "T"
            If rbcSAdvtSep(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "B"
            If rbcSAdvtSep(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            ilAltered = True
            Exit Sub
    End Select
    Select Case tlVpf.sSCompType  'Separate competitive by break
        Case "T"
            If rbcSCompetitive(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "B"
            If rbcSCompetitive(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcSCompetitive(2).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    If tlVpf.sSCompType = "T" Then
        gUnpackLength tlVpf.iSCompLen(0), tlVpf.iSCompLen(1), "2", True, slStr
    Else
        slStr = ""
    End If
    If gLengthToLong(edcSCompSepLen.Text) <> gLengthToLong(slStr) Then
        ilAltered = True
        Exit Sub
    End If
    Select Case tlVpf.sSSellOut  'Sellout
        Case "U"
            If rbcSSellout(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "B"
            If rbcSSellout(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "T"
            If rbcSSellout(2).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcSSellout(3).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    Select Case tlVpf.sSOverBook  'Overbook
        Case "Y"
            If rbcSOverbook(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcSOverbook(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    Select Case tlVpf.sSForceMG  'Force MG
        Case "W"
            If rbcSMove(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcSMove(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    'Select Case tlVpf.sSPlaceNet  'Network spot placement
    '    Case "F"
    '        If rbcSMixed(0).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case "L"
    '        If rbcSMixed(1).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case Else
    '        If rbcSMixed(2).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    'End Select
    'Select Case tlVpf.sSIntoRC  'Place programs into Rate card
    '    Case "Y"
    '        If rbcSIntoRC(0).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case Else
    '        If rbcSIntoRC(1).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    'End Select
    Select Case tlVpf.sSAvailOrder  'spot placement
        Case "1"
            If rbcSBreak(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "2"
            If rbcSBreak(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "3"
            If rbcSBreak(2).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "4"
            If rbcSBreak(3).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "5"
            If rbcSBreak(4).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "6"
            If rbcSBreak(5).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcSBreak(6).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    For ilLoop = 0 To 9     'Length and group #
        If tlVpf.iSLen(ilLoop) <> tgVpf(ilVpfIndex).iSLen(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iSLenGroup(ilLoop) <> tgVpf(ilVpfIndex).iSLenGroup(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
    Next ilLoop
    Select Case tlVpf.sSCommCalc  'Salesperson commission
        Case "B"
            If rbcSSalesperson(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "C"
            If rbcSSalesperson(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    'For ilLoop = 1 To 24     'Psa/promo
    For ilLoop = 0 To 23     'Psa/promo
        If tlVpf.iMMFPSA(ilLoop) <> tgVpf(ilVpfIndex).iMMFPSA(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iMSAPSA(ilLoop) <> tgVpf(ilVpfIndex).iMSAPSA(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iMSUPSA(ilLoop) <> tgVpf(ilVpfIndex).iMSUPSA(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iMMFPromo(ilLoop) <> tgVpf(ilVpfIndex).iMMFPromo(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iMSAPromo(ilLoop) <> tgVpf(ilVpfIndex).iMSAPromo(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iMSUPromo(ilLoop) <> tgVpf(ilVpfIndex).iMSUPromo(ilLoop) Then
            ilAltered = True
            Exit Sub
        End If
    Next ilLoop
    'Log
    gUnpackDate tlVpf.iLLastDateCpyAsgn(0), tlVpf.iLLastDateCpyAsgn(1), slStr    'Last log date
    slDate = edcLLDCpyAsgn.Text
    If gValidDate(slDate) Then
        If gValidDate(slStr) Then
            If gDateValue(slDate) <> gDateValue(slStr) Then
                ilAltered = True
                Exit Sub
            End If
        Else
            ilAltered = True
            Exit Sub
        End If
    Else
        If gValidDate(slStr) Then
            ilAltered = True
            Exit Sub
        End If
    End If
    gUnpackDate tlVpf.iLLD(0), tlVpf.iLLD(1), slStr    'Last log date
    slDate = edcLDate(0).Text
    If gValidDate(slDate) Then
        If gValidDate(slStr) Then
            If gDateValue(slDate) <> gDateValue(slStr) Then
                ilAltered = True
                Exit Sub
            End If
        Else
            ilAltered = True
            Exit Sub
        End If
    Else
        If gValidDate(slStr) Then
            ilAltered = True
            Exit Sub
        End If
    End If
    gUnpackDate tlVpf.iLPD(0), tlVpf.iLPD(1), slStr    'Last preliminary date
    slDate = edcLDate(1).Text
    If gValidDate(slDate) Then
        If gValidDate(slStr) Then
            If gDateValue(slDate) <> gDateValue(slStr) Then
                ilAltered = True
                Exit Sub
            End If
        Else
            ilAltered = True
            Exit Sub
        End If
    Else
        If gValidDate(slStr) Then
            ilAltered = True
            Exit Sub
        End If
    End If
    gUnpackDate tlVpf.iLastLog(0), tlVpf.iLastLog(1), slStr    'Last log date
    slDate = edcLAffDate(0).Text
    If gValidDate(slDate) Then
        If gValidDate(slStr) Then
            If gDateValue(slDate) <> gDateValue(slStr) Then
                ilAltered = True
                Exit Sub
            End If
        Else
            ilAltered = True
            Exit Sub
        End If
    Else
        If gValidDate(slStr) Then
            ilAltered = True
            Exit Sub
        End If
    End If
    gUnpackDate tlVpf.iLastCP(0), tlVpf.iLastCP(1), slStr    'Last C.P. date
    slDate = edcLAffDate(1).Text
    If gValidDate(slDate) Then
        If gValidDate(slStr) Then
            If gDateValue(slDate) <> gDateValue(slStr) Then
                ilAltered = True
                Exit Sub
            End If
        Else
            ilAltered = True
            Exit Sub
        End If
    Else
        If gValidDate(slStr) Then
            ilAltered = True
            Exit Sub
        End If
    End If
    Select Case tlVpf.slTimeZone  'Grid resolution
        Case "E"
            If rbcLZone(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "C"
            If rbcLZone(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case "M"
            If rbcLZone(2).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcLZone(3).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    Select Case tlVpf.sLDayLight  'Daylight saving
        Case "Y"
            If rbcLDaylight(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcLDaylight(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    Select Case tlVpf.sLTiming  'Log Timing
        Case "Y"
            If rbcLTiming(0).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
        Case Else
            If rbcLTiming(1).Value <> True Then
                ilAltered = True
                Exit Sub
            End If
    End Select
    'Select Case tlVpf.sLAvailLen  'Show lengths on unsold avails
    '    Case "Y"
    '        If rbcLLen(0).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case Else
    '        If rbcLLen(1).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    'End Select
    If Val(edcSLen.Text) <> tlVpf.iSDLen Then      'Default spot length
        ilAltered = True
        Exit Sub
    End If
    'Select Case tlVpf.sAffCPs  'Create Affiliate Custom CPs
    '    Case "Y"
    '        If rbcLAffCPs(0).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case Else
    '        If rbcLAffCPs(1).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    'End Select
    'Select Case tlVpf.sAffTimes  'Post Affiliate Exact Times
    '    Case "Y"
    '        If rbcLAffTimes(0).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case Else
    '        If rbcLAffTimes(1).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    'End Select
    'Select Case tlVpf.sLShowCut  'Show cut/instruction #
    '    Case "C"
    '        If rbcLCut(0).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case "I"
    '        If rbcLCut(1).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case "B"
    '        If rbcLCut(2).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case Else
    '        If rbcLCut(3).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    'End Select
    'Select Case tlVpf.slTimeFormat  'Time format
    '    Case "A"
    '        If rbcLTime(0).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Case Else
    '        If rbcLTime(1).Value <> True Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    'End Select
    If tmVef.sType = "A" Then
        Select Case tlVpf.sCopyOnAir  'Allow Copy on Airing Vehicle
            Case "Y"
                If rbcLCopyOnAir(0).Value <> True Then
                    ilAltered = True
                    Exit Sub
                End If
            Case Else
                If rbcLCopyOnAir(1).Value <> True Then
                    ilAltered = True
                    Exit Sub
                End If
        End Select
    End If
'    gUnpackLength tlVpf.iBCal(0), tlVpf.iBCal(1), slStr    'Last log date
'    lacBInvDate(0).Text = slStr
'    gUnpackLength tlVpf.iBStd(0), tlVpf.iBStd(1), slStr    'Last preliminary date
'    lacBinvDate(1).Text = slStr
    If Trim$(edcIFGroupNo.Text) <> Trim$(tlVpf.sGGroupNo) Then   'Vehicle Group Number
        ilAltered = True
        Exit Sub
    End If
    If Trim$(edcExpCntrVehNo.Text) <> Trim$(tlVpf.sExpVehNo) Then   'Vehicle Group Number
        ilAltered = True
        Exit Sub
    End If
    If Trim$(edcARBCode.Text) <> Trim$(tlVpf.sARBCode) Then   'Arbitron code
        ilAltered = True
        Exit Sub
    End If
    If Trim$(edcRadarCode.Text) <> Trim$(tlVpf.sRadarCode) Then   'Arbitron code
        ilAltered = True
        Exit Sub
    End If
    If Trim$(edcStnFdCode.Text) <> Trim$(tlVpf.sStnFdCode) Then   'Station Feed Code
        ilAltered = True
        Exit Sub
    End If
    If ckcStnFdInfo(0).Value = vbChecked Then
        If tlVpf.sStnFdCart <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sStnFdCart <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If ckcStnFdInfo(1).Value = vbChecked Then
        If tlVpf.sStnFdXRef <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sStnFdXRef <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    'If StrComp(edcLgVehNm.Text, smInitLgVehNm, 1) <> 0 Then
    '    ilAltered = True
    '    Exit Sub
    'End If
    'If StrComp(edcLgHd1.Text, smInitLgHd1, 1) <> 0 Then
    '    ilAltered = True
    '    Exit Sub
    'End If
    'If StrComp(edcLgFt1.Text, smInitLgFt1, 1) <> 0 Then
    '    ilAltered = True
    '    Exit Sub
    'End If
    'If StrComp(edcLgFt2.Text, smInitLgFt2, 1) <> 0 Then
    '    ilAltered = True
    '    Exit Sub
    'End If
    
    slStr = Trim$(edcLog(0).Text)
    If StrComp(slStr, smLogExpt, vbTextCompare) <> 0 Then
        ilAltered = True
        Exit Sub
    End If

    'Interface
    For ilLoop = 1 To 5 Step 1
        slTime = Trim$(edcIFEST(ilLoop - 1).Text)
        If slTime <> "" Then
            slTime = gConvertTime(slTime)
            If slTime <> "12:00AM" Then
                If tlVpf.iESTEndTime(ilLoop - 1) <> Minute(slTime) + 60 * Hour(slTime) Then
                    ilAltered = True
                    Exit Sub
                End If
            Else
                If tlVpf.iESTEndTime(ilLoop - 1) <> 24 * 60 Then
                    ilAltered = True
                    Exit Sub
                End If
            End If
        Else
            If tlVpf.iESTEndTime(ilLoop - 1) <> 0 Then
                ilAltered = True
                Exit Sub
            End If
        End If
        slTime = Trim$(edcIFCST(ilLoop - 1).Text)
        If slTime <> "" Then
            slTime = gConvertTime(slTime)
            If slTime <> "12:00AM" Then
                If tlVpf.iCSTEndTime(ilLoop - 1) <> Minute(slTime) + 60 * Hour(slTime) Then
                    ilAltered = True
                    Exit Sub
                End If
            Else
                If tlVpf.iCSTEndTime(ilLoop - 1) <> 24 * 60 Then
                    ilAltered = True
                    Exit Sub
                End If
           End If
        Else
            If tlVpf.iCSTEndTime(ilLoop - 1) <> 0 Then
                ilAltered = True
                Exit Sub
            End If
        End If
        slTime = Trim$(edcIFMST(ilLoop - 1).Text)
        If slTime <> "" Then
            slTime = gConvertTime(slTime)
            If slTime <> "12:00AM" Then
                If tlVpf.iMSTEndTime(ilLoop - 1) <> Minute(slTime) + 60 * Hour(slTime) Then
                    ilAltered = True
                    Exit Sub
                End If
            Else
                If tlVpf.iMSTEndTime(ilLoop - 1) <> 24 * 60 Then
                    ilAltered = True
                    Exit Sub
                End If
            End If
        Else
            If tlVpf.iMSTEndTime(ilLoop - 1) <> 0 Then
                ilAltered = True
                Exit Sub
            End If
        End If
        slTime = Trim$(edcIFPST(ilLoop - 1).Text)
        If slTime <> "" Then
            slTime = gConvertTime(slTime)
            If slTime <> "12:00AM" Then
                If tlVpf.iPSTEndTime(ilLoop - 1) <> Minute(slTime) + 60 * Hour(slTime) Then
                    ilAltered = True
                    Exit Sub
                End If
            Else
                If tlVpf.iPSTEndTime(ilLoop - 1) <> 24 * 60 Then
                    ilAltered = True
                    Exit Sub
                End If
            End If
        Else
            If tlVpf.iPSTEndTime(ilLoop - 1) <> 0 Then
                ilAltered = True
                Exit Sub
            End If
        End If
    Next ilLoop
    For ilLoop = 1 To 4 Step 1
        If Trim$(tlVpf.sMapZone(ilLoop - 1)) <> Trim$(edcIFZone(ilLoop - 1).Text) Then
            ilAltered = True
            Exit Sub
        End If
        If Trim$(tlVpf.sMapProgCode(ilLoop - 1)) <> Trim$(edcIFProgCode(ilLoop - 1).Text) Then
            ilAltered = True
            Exit Sub
        End If
        If tlVpf.iMapDPNo(ilLoop - 1) <> Val(edcIFDPNo(ilLoop - 1).Text) Then
            ilAltered = True
            Exit Sub
        End If
    Next ilLoop
    If ckcIFExport(0).Value = vbChecked Then
        If tlVpf.sExpHiClear <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sExpHiClear <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If ckcIFExport(1).Value = vbChecked Then
        If tlVpf.sExpHiDallas <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sExpHiDallas <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If ckcIFExport(2).Value = vbChecked Then
        If tlVpf.sExpHiPhoenix <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sExpHiPhoenix <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If ckcIFExport(3).Value = vbChecked Then
        If tlVpf.sExpHiNY <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sExpHiNY <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If ckcIFExport(6).Value = vbChecked Then
        If tlVpf.sExpHiNYISCI <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If (tlVpf.sExpHiNYISCI <> "N") And (Trim$(tlVpf.sExpHiNYISCI) <> "") Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If ckcIFExport(4).Value = vbChecked Then
        If tlVpf.sExpHiCmmlChg <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sExpHiCmmlChg <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If rbcExpBkCpyCart(0).Value Then
        If tlVpf.sExpBkCpyCart <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sExpBkCpyCart <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If rbcIFBulk(0).Value Then
        If tlVpf.sBulkXFer <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sBulkXFer <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If rbcIFSelling(0).Value Then
        If tlVpf.sClearAsSell <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sClearAsSell <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If rbcIFTime(0).Value Then
        If tlVpf.sClearChgTime <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sClearChgTime <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    slStr = Trim$(edcEDASWindow.Text)
    If tlVpf.lEDASWindow <> Val(slStr) Then
        ilAltered = True
        Exit Sub
    End If
    If ckcKCGenRot.Value = vbChecked Then
        If tlVpf.sKCGenRot <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sKCGenRot <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If ckcExportSQL.Value = vbChecked Then
        If tlVpf.sExportSQL <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If tlVpf.sExportSQL <> "N" Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If tmVef.sType = "V" Then
        If imVirtChgVeh Then
            ilAltered = True
            Exit Sub
        End If
    End If
    ilValue = Asc(tgSpf.sUsingFeatures2)
    ''If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G")) And (ilValue And BARTER) = BARTER Then
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S")) And (ilValue And BARTER) = BARTER Then
    If (tmVef.sType = "R") And ((ilValue And BARTER) = BARTER) Then
'        If tlVpf.lDefAcqCost <> Val(edcBarter(0).Text) Then
'            ilAltered = True
'            Exit Sub
'        End If
'        If tlVpf.lActAcqCost <> Val(edcBarter(1).Text) Then
'            ilAltered = True
'            Exit Sub
'        End If
'        Select Case Trim$(tlVpf.sBarterMethod)
'            Case "A"
'                If rbcBarterMethod(0).Value <> True Then
'                    ilAltered = True
'                    Exit Sub
'                End If
'            Case "M"
'                If rbcBarterMethod(1).Value <> True Then
'                    ilAltered = True
'                    Exit Sub
'                End If
'                If tlVpf.iBarterThreshold <> Val(edcBarterMethod(0).Text) Then
'                    ilAltered = True
'                    Exit Sub
'                End If
'            Case "U"
'                If rbcBarterMethod(2).Value <> True Then
'                    ilAltered = True
'                    Exit Sub
'                End If
'                If tlVpf.iBarterThreshold <> Val(edcBarterMethod(1).Text) Then
'                    ilAltered = True
'                    Exit Sub
'                End If
'            Case Else
'                If rbcBarterMethod(3).Value <> True Then
'                    ilAltered = True
'                    Exit Sub
'                End If
'                If tlVpf.iBarterXFree <> Val(edcBarterMethod(2).Text) Then
'                    ilAltered = True
'                    Exit Sub
'                End If
'                If tlVpf.iBarterYSold <> Val(edcBarterMethod(3).Text) Then
'                    ilAltered = True
'                    Exit Sub
'                End If
'        End Select
    End If
    ilValue = Asc(tgSpf.sUsingFeatures2)
    '6/19/07:  Jim:  Allow definition for Conventional, Airing and Game.  Always set no for Packages and Selling
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or ((tmVef.sType = "A") And (rbcLCopyOnAir(0).Value)) Or (tmVef.sType = "G") Or (tmVef.sType = "P")) And ((ilValue And SPLITCOPY) = SPLITCOPY) Then
    '5/11/11: Allow selling to be set to No
    'If ((tmVef.sType = "C") Or (tmVef.sType = "A") Or (tmVef.sType = "G")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
    If ((tmVef.sType = "C") Or (tmVef.sType = "A") Or (tmVef.sType = "G") Or (tmVef.sType = "S")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
        If rbcSplitCopy(0).Value Then
            If tlVpf.sAllowSplitCopy <> "Y" Then
                ilAltered = True
                Exit Sub
            End If
        Else
            If (tlVpf.sAllowSplitCopy <> "N") And (Trim$(tlVpf.sAllowSplitCopy) <> "") Then
                ilAltered = True
                Exit Sub
            End If
        End If
    End If
    If rbcShowRateOnInsertion(0).Value Then
        If tlVpf.sShowRateOnInsert <> "Y" Then
            ilAltered = True
            Exit Sub
        End If
    Else
        If (tlVpf.sShowRateOnInsert <> "N") And (Trim$(tlVpf.sShowRateOnInsert) <> "") Then
            ilAltered = True
            Exit Sub
        End If
    End If
    slStr = Trim$(edcEMail.Text)
    If StrComp(slStr, smEMail, vbTextCompare) <> 0 Then
        ilAltered = True
        Exit Sub
    End If
    'If smXDXMLForm = "ISCI" Then
    If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) Then
        If Val(edcInterfaceID(1).Text) <> tlVpf.iInterfaceID Then
            ilAltered = True
            Exit Sub
        End If
    End If
    If ((Asc(tgSpf.sUsingFeatures7) And WEGENEREXPORT) = WEGENEREXPORT) Then
        If ckcAffExport(0).Value = vbChecked Then
            If tlVpf.sWegenerExport <> "Y" Then
                ilAltered = True
                Exit Sub
            End If
        Else
            If tlVpf.sWegenerExport = "Y" Then
                ilAltered = True
                Exit Sub
            End If
        End If
    End If
    'If ((Asc(tgSpf.sUsingFeatures7) And OLAEXPORT) = OLAEXPORT) Then
    '    If ckcAffExport(1).Value = vbChecked Then
    '        If tlVpf.sOLAExport <> "Y" Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    Else
    '        If tlVpf.sOLAExport = "Y" Then
    '            ilAltered = True
    '            Exit Sub
    '        End If
    '    End If
    'End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
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
    mSetChg tmVpf, imVpfIndex, imIgnoreChg, imAltered
    If imAltered Then
        If tmVef.sType = "V" Then
            'Update button set if all mandatory fields have data and any field altered
            If (mVirtTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (imVirtChgVeh) And (UBound(smVirtSave, 2) > 1) Then
                If imUpdateAllowed Then
                    cmcUpdate.Enabled = True
                Else
                    cmcUpdate.Enabled = False
                End If
            Else
                cmcUpdate.Enabled = False
            End If
        Else
            cmcUpdate.Enabled = True
        End If
    Else
        If tmVef.sType = "V" Then
            'Update button set if all mandatory fields have data and any field altered
            If (mVirtTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (imVirtChgVeh) And (UBound(smVirtSave, 2) > 1) Then
                If imUpdateAllowed Then
                    cmcUpdate.Enabled = True
                Else
                    cmcUpdate.Enabled = False
                End If
            Else
                cmcUpdate.Enabled = False
            End If
        Else
            cmcUpdate.Enabled = False
        End If
    End If
    'Revert button set if any field changed
    If (imAltered) Or (imVirtChgVeh) Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    If (Not imAltered) And (Not imVirtChgVeh) Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
    If tmVef.sType = "A" Then
        plcLCopyOnAir.Visible = True
    Else
        plcLCopyOnAir.Visible = False
    End If
    'If ((tmVef.sType = "C") And (tmVef.iVefCode <= 0)) Or (tmVef.sType = "A") Or (tmVef.sType = "L") Then
    '    lacLgVehNm.Enabled = True
    '    lacLgHd1.Enabled = True
    '    lacLgFt1.Enabled = True
    '    lacLgFt2.Enabled = True
    '    edcLgVehNm.Enabled = True
    '    edcLgHd1.Enabled = True
    '    edcLgFt1.Enabled = True
    '    edcLgFt2.Enabled = True
    'Else
    '    lacLgVehNm.Enabled = False
    '    lacLgHd1.Enabled = False
    '    lacLgFt1.Enabled = False
    '    lacLgFt2.Enabled = False
    '    edcLgVehNm.Enabled = False
    '    edcLgHd1.Enabled = False
    '    edcLgFt1.Enabled = False
    '    edcLgFt2.Enabled = False
    'End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetGTZ                         *
'*                                                     *
'*             Created:5/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set Time zone area             *
'*                                                     *
'*******************************************************
Private Sub mSetGTZ()
    Dim flLeft As Single
    Dim flTop As Single
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    If (imGTZBoxNo < 1) Or (imGTZBoxNo > imTZMaxCtrls * (UBound(tmVpf.sGZone))) Then
        Exit Sub
    End If
    ilIndex = ((imGTZBoxNo - 1) \ imTZMaxCtrls) + 1
    If (imGTZBoxNo Mod imTZMaxCtrls) = GNAMEINDEX Then  'Time zone
        edcGTZDropDown.MaxLength = 3
        edcGTZDropDown.Width = tmTZCtrls(GNAMEINDEX).fBoxW
        edcGTZDropDown.Text = Trim$(tmVpf.sGZone(ilIndex - 1))
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
        flLeft = tmTZCtrls(GNAMEINDEX).fBoxX
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEDZONEINDEX Then  'Fed (events transmitted)
        edcGTZDropDown.MaxLength = 1
        edcGTZDropDown.Width = tmTZCtrls(GFEDZONEINDEX).fBoxW
        edcGTZDropDown.Text = Trim$(tmVpf.sGFed(ilIndex - 1))
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
        flLeft = tmTZCtrls(GFEDZONEINDEX).fBoxX
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GLOCALADJINDEX Then  'Local time adjustment
        edcGTZDropDown.MaxLength = 2
        edcGTZDropDown.Width = tmTZCtrls(GLOCALADJINDEX).fBoxW
        If Trim$(tmVpf.sGZone(ilIndex - 1)) <> "" Then
            edcGTZDropDown.Text = Trim$(Str$(tmVpf.iGLocalAdj(ilIndex - 1)))
        Else
            edcGTZDropDown.Text = ""
        End If
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
'        flLeft = imTZX + imTZW + 15
        flLeft = tmTZCtrls(GLOCALADJINDEX).fBoxX
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDADJINDEX Then  'Feed Time Adjustment
        edcGTZDropDown.MaxLength = 2
        edcGTZDropDown.Width = tmTZCtrls(GFEEDADJINDEX).fBoxW
        If Trim$(tmVpf.sGZone(ilIndex - 1)) <> "" Then
            edcGTZDropDown.Text = Trim$(Str$(tmVpf.iGFeedAdj(ilIndex - 1)))
        Else
            edcGTZDropDown.Text = ""
        End If
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
'        flLeft = imTZX + 2 * (imTZW + 15)
        flLeft = tmTZCtrls(GFEEDADJINDEX).fBoxX
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GVERDISPLINDEX Then  'Versions displacement
        edcGTZDropDown.MaxLength = 3
        edcGTZDropDown.Width = tmTZCtrls(GVERDISPLINDEX).fBoxW
        If Trim$(tmVpf.sGZone(ilIndex - 1)) <> "" Then
            edcGTZDropDown.Text = Trim$(Str$(tmVpf.iGV1Z(ilIndex - 1)))
        Else
            edcGTZDropDown.Text = ""
        End If
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
'        flLeft = imTZX + 2 * (imTZW + 15)
        flLeft = tmTZCtrls(GVERDISPLINDEX).fBoxX
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GVERDISPLINDEX + 1 Then 'Versions displacement
        edcGTZDropDown.MaxLength = 3
        edcGTZDropDown.Width = tmTZCtrls(GVERDISPLINDEX + 1).fBoxW
        If Trim$(tmVpf.sGZone(ilIndex - 1)) <> "" Then
            edcGTZDropDown.Text = Trim$(Str$(tmVpf.iGV2Z(ilIndex - 1)))
        Else
            edcGTZDropDown.Text = ""
        End If
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
'        flLeft = imTZX + 2 * (imTZW + 15)
        flLeft = tmTZCtrls(GVERDISPLINDEX + 1).fBoxX
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GVERDISPLINDEX + 2 Then 'Versions displacement
        edcGTZDropDown.MaxLength = 3
        edcGTZDropDown.Width = tmTZCtrls(GVERDISPLINDEX + 2).fBoxW
        If Trim$(tmVpf.sGZone(ilIndex - 1)) <> "" Then
            edcGTZDropDown.Text = Trim$(Str$(tmVpf.iGV3Z(ilIndex - 1)))
        Else
            edcGTZDropDown.Text = ""
        End If
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
'        flLeft = imTZX + 2 * (imTZW + 15)
        flLeft = tmTZCtrls(GVERDISPLINDEX + 2).fBoxX
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GVERDISPLINDEX + 3 Then 'Versions displacement
        edcGTZDropDown.MaxLength = 3
        edcGTZDropDown.Width = tmTZCtrls(GVERDISPLINDEX + 3).fBoxW
        If Trim$(tmVpf.sGZone(ilIndex - 1)) <> "" Then
            edcGTZDropDown.Text = Trim$(Str$(tmVpf.iGV4Z(ilIndex - 1)))
        Else
            edcGTZDropDown.Text = ""
        End If
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
'        flLeft = imTZX + 2 * (imTZW + 15)
        flLeft = tmTZCtrls(GVERDISPLINDEX + 3).fBoxX
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GCMMLSCHINDEX Then  'Commercial schd
        If (tmVpf.sGCSVer(ilIndex - 1) <> "A") Or (tmVpf.sGCSVer(ilIndex - 1) <> "O") Then
            tmVpf.sGCSVer(ilIndex - 1) = "O"
            mSetCommands
        End If
        pbcGTZToggle.Width = tmTZCtrls(GCMMLSCHINDEX).fBoxW
        pbcGTZToggle_Paint
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
        flLeft = tmTZCtrls(GCMMLSCHINDEX).fBoxX
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEDDELIVERYINDEX Then  'Fed (events transmitted)
        If (tmVff.sFedDelivery(ilIndex - 1) <> "Y") Or (tmVff.sFedDelivery(ilIndex - 1) <> "N") Then
            tmVff.sFedDelivery(ilIndex - 1) = "Y"
            imVffChg = True
            mSetCommands
        End If
        pbcGTZToggle.Width = tmTZCtrls(GFEDDELIVERYINDEX).fBoxW
        pbcGTZToggle_Paint
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
        flLeft = tmTZCtrls(GFEDDELIVERYINDEX).fBoxX
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDINDEX Then  'Feed
        mFeedPop
        edcGTZDropDown.MaxLength = 20
        lbcFeed.Height = gListBoxHeight(lbcFeed.ListCount, 6)
        edcGTZDropDown.Width = tmTZCtrls(GFEEDINDEX).fBoxW - cmcGTZDropDown.Width
        imChgMode = True
        If Trim$(tmVpf.sGZone(ilIndex - 1)) <> "" Then
            'Find matching feed code
            ilFound = False
            For ilLoop = 0 To UBound(tmFeedCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
                slNameCode = tmFeedCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet = CP_MSG_NONE Then
                    If Val(slCode) = tmVpf.iGMnfNCode(ilIndex - 1) Then
                        ilFound = True
                        Exit For
                    End If
                Else
                    ilFound = False
                End If
            Next ilLoop
            If ilFound Then
                lbcFeed.ListIndex = ilLoop + 2
            Else
                lbcFeed.ListIndex = 1   '[None]
            End If
        Else
            lbcFeed.ListIndex = 1   '[None]
        End If
        If lbcFeed.ListIndex >= 0 Then
            edcGTZDropDown.Text = lbcFeed.List(lbcFeed.ListIndex)
        Else
            edcGTZDropDown.Text = ""
        End If
        imChgMode = False
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
'        flLeft = imTZX + 2 * (imTZW + 15)
        flLeft = tmTZCtrls(GFEEDINDEX).fBoxX
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GBUSINDEX Then  'Bus
        edcGTZDropDown.MaxLength = 2
        edcGTZDropDown.Width = tmTZCtrls(GBUSINDEX).fBoxW
        If Trim$(tmVpf.sGZone(ilIndex - 1)) <> "" Then
            edcGTZDropDown.Text = Trim$(tmVpf.sGBus(ilIndex - 1))
        Else
            edcGTZDropDown.Text = ""
        End If
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
'        flLeft = imTZX + 2 * (imTZW + 15)
        flLeft = tmTZCtrls(GBUSINDEX).fBoxX
    Else    'Schedulet
        edcGTZDropDown.MaxLength = 2
        edcGTZDropDown.Width = tmTZCtrls(GSCHDINDEX).fBoxW
        If Trim$(tmVpf.sGZone(ilIndex - 1)) <> "" Then
            edcGTZDropDown.Text = Trim$(tmVpf.sGSked(ilIndex - 1))
        Else
            edcGTZDropDown.Text = ""
        End If
        flTop = imTZY + (ilIndex - 1) * (imTZH - 15)
'        flLeft = imTZX + 3 * (imTZW + 15)
        flLeft = tmTZCtrls(GSCHDINDEX).fBoxX
    End If
    If (imGTZBoxNo Mod imTZMaxCtrls) = GCMMLSCHINDEX Then  'Commercial schd
        lbcFeed.Visible = False
        cmcGTZDropDown.Visible = False
        cmcGTZDropDown.Visible = False
        edcGTZDropDown.Visible = False
        pbcGTZToggle.Move flLeft, flTop
        pbcGTZToggle.Visible = True
        pbcGTZToggle.SetFocus
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEDZONEINDEX Then  'Fed (events transmitted)
        pbcGTZToggle.Visible = False
        lbcFeed.Visible = False
        cmcGTZDropDown.Visible = False
        edcGTZDropDown.Move flLeft, flTop
        edcGTZDropDown.SelStart = 0
        edcGTZDropDown.SelLength = Len(edcGTZDropDown.Text)
        edcGTZDropDown.Visible = True
        edcGTZDropDown.SetFocus
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEDDELIVERYINDEX Then  'Fed (events transmitted)
        lbcFeed.Visible = False
        cmcGTZDropDown.Visible = False
        cmcGTZDropDown.Visible = False
        edcGTZDropDown.Visible = False
        pbcGTZToggle.Move flLeft, flTop
        pbcGTZToggle.Visible = True
        pbcGTZToggle.SetFocus
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDINDEX Then  'Feed
        pbcGTZToggle.Visible = False
        edcGTZDropDown.Move flLeft, flTop
        cmcGTZDropDown.Move edcGTZDropDown.Left + edcGTZDropDown.Width, edcGTZDropDown.Top
        lbcFeed.Move pbcGTZ.Left + edcGTZDropDown.Left, pbcGTZ.Top + edcGTZDropDown.Top + edcGTZDropDown.Height
        edcGTZDropDown.SelStart = 0
        edcGTZDropDown.SelLength = Len(edcGTZDropDown.Text)
        cmcGTZDropDown.Visible = True
        edcGTZDropDown.Visible = True
        edcGTZDropDown.SetFocus
    Else
        pbcGTZToggle.Visible = False
        lbcFeed.Visible = False
        cmcGTZDropDown.Visible = False
        edcGTZDropDown.Move flLeft, flTop
        edcGTZDropDown.SelStart = 0
        edcGTZDropDown.SelLength = Len(edcGTZDropDown.Text)
        edcGTZDropDown.Visible = True
        edcGTZDropDown.SetFocus
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetMPromo                      *
'*                                                     *
'*             Created:5/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set promo area                 *
'*                                                     *
'*******************************************************
Private Sub mSetMPromo()
    Dim flLeft As Single
    Dim flTop As Single
    If (imMPromoBoxNo < 1) Or (imMPromoBoxNo > 3 * (UBound(tmVpf.iMMFPromo) + LBONE)) Then
        Exit Sub
    End If
    If imMPromoBoxNo - LBONE <= UBound(tmVpf.iMMFPromo) Then
        edcMPromo.Text = Trim$(Str$(tmVpf.iMMFPromo(imMPromoBoxNo - LBONE)))
        flTop = imPromoY + (imMPromoBoxNo - 1) * (imPromoH - 15)
        flLeft = imPromoX
    ElseIf (imMPromoBoxNo - LBONE >= (UBound(tmVpf.iMMFPromo) + 1)) And (imMPromoBoxNo <= 2 * (UBound(tmVpf.iMMFPromo) + LBONE)) Then
        edcMPromo.Text = Trim$(Str$(tmVpf.iMSAPromo(imMPromoBoxNo - (UBound(tmVpf.iMMFPromo) + LBONE) - 1)))
        flTop = imPromoY + (imMPromoBoxNo - (UBound(tmVpf.iMMFPromo) + LBONE) - 1) * (imPromoH - 15)
        flLeft = imPromoX + imPromoW + 15
    Else
        edcMPromo.Text = Trim$(Str$(tmVpf.iMSUPromo(imMPromoBoxNo - 2 * (UBound(tmVpf.iMMFPromo) + LBONE) - 1)))
        flTop = imPromoY + (imMPromoBoxNo - 2 * (UBound(tmVpf.iMMFPromo) + LBONE) - 1) * (imPromoH - 15)
        flLeft = imPromoX + 2 * (imPromoW + 15)
    End If
    edcMPromo.Move flLeft, flTop
    edcMPromo.SelStart = 0
    edcMPromo.SelLength = Len(edcMPromo.Text)
    edcMPromo.Visible = True
    edcMPromo.SetFocus
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetMPsa                        *
'*                                                     *
'*             Created:5/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set PSA area                   *
'*                                                     *
'*******************************************************
Private Sub mSetMPsa()
    Dim flLeft As Single
    Dim flTop As Single
    If (imMPSABoxNo < 1) Or (imMPSABoxNo > 3 * (UBound(tmVpf.iMMFPSA) + LBONE)) Then
        Exit Sub
    End If
    If imMPSABoxNo - LBONE <= UBound(tmVpf.iMMFPSA) Then
        edcMPSA.Text = Trim$(Str$(tmVpf.iMMFPSA(imMPSABoxNo - LBONE)))
        flTop = imPsaY + (imMPSABoxNo - 1) * (imPsaH - 15)
        flLeft = imPsaX
    ElseIf (imMPSABoxNo - LBONE >= UBound(tmVpf.iMMFPSA) + 1) And (imMPSABoxNo <= 2 * (UBound(tmVpf.iMMFPSA) + LBONE)) Then
        edcMPSA.Text = Trim$(Str$(tmVpf.iMSAPSA(imMPSABoxNo - (UBound(tmVpf.iMMFPSA) + LBONE) - 1)))
        flTop = imPsaY + (imMPSABoxNo - (UBound(tmVpf.iMMFPSA) + LBONE) - 1) * (imPsaH - 15)
        flLeft = imPsaX + imPsaW + 15
    Else
        edcMPSA.Text = Trim$(Str$(tmVpf.iMSUPSA(imMPSABoxNo - 2 * (UBound(tmVpf.iMMFPSA) + LBONE) - 1)))
        flTop = imPsaY + (imMPSABoxNo - 2 * (UBound(tmVpf.iMMFPSA) + LBONE) - 1) * (imPsaH - 15)
        flLeft = imPsaX + 2 * (imPsaW + 15)
    End If
    edcMPSA.Move flLeft, flTop
    edcMPSA.SelStart = 0
    edcMPSA.SelLength = Len(edcMPSA.Text)
    edcMPSA.Visible = True
    edcMPSA.SetFocus
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetSSpotLen                    *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set spot length box            *
'*                                                     *
'*******************************************************
Private Sub mSetSSpotLen()
    Dim flLeft As Single
    Dim flTop As Single
    If (imSSpotLenBoxNo < 0) Or (imSSpotLenBoxNo > 2 * (UBound(tmVpf.iSLen) + 1) - 1) Then
        Exit Sub
    End If
    If imSSpotLenBoxNo <= UBound(tmVpf.iSLen) Then
        edcSSpotLG.Text = Trim$(Str$(tmVpf.iSLen(imSSpotLenBoxNo)))
        flLeft = 30 + imSSpotLenBoxNo * (edcSSpotLG.Width + 15)
        flTop = 15
    Else
        edcSSpotLG.Text = Trim$(Str$(tmVpf.iSLenGroup(imSSpotLenBoxNo - UBound(tmVpf.iSLen) - 1)))
        flLeft = 30 + (imSSpotLenBoxNo - UBound(tmVpf.iSLen) - 1) * (edcSSpotLG.Width + 15)
        flTop = edcSSpotLG.Height
    End If
    edcSSpotLG.Move flLeft, flTop
    edcSSpotLG.SelStart = 0
    edcSSpotLG.SelLength = Len(edcSSpotLG.Text)
    edcSSpotLG.Visible = True
    edcSSpotLG.SetFocus
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSSpotLenPaint                  *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint sales spot length table  *
'*                                                     *
'*******************************************************
Private Sub mSSpotLenPaint(tlVpf As VPF)
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim flX As Single
    Dim flY As Single
    flX = imSpotLGX + fgBoxInsetX
    flY = imSpotLGY
    For ilRow = 0 To 1 Step 1
        For ilCol = 0 To 9 Step 1
            gPaintArea pbcSSpotLen, flX, flY + 15, imSpotLGW - fgBoxInsetX - 15, imSpotLGH - 45, WHITE
            pbcSSpotLen.CurrentX = flX
            pbcSSpotLen.CurrentY = flY - 15
            If ilRow = 0 Then
                pbcSSpotLen.Print Trim$(Str$(tlVpf.iSLen(ilCol)))
            Else
                pbcSSpotLen.Print Trim$(Str$(tlVpf.iSLenGroup(ilCol)))
            End If
            flX = flX + imSpotLGW + 15
        Next ilCol
        flX = imSpotLGX + fgBoxInsetX
        flY = flY + imSpotLGH - 15
    Next ilRow
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
    If imVpfChanged Then
'        sgVpfStamp = "~"    'Force read
'        ilRet = gVpfRead()
    End If

    Screen.MousePointer = vbDefault
    Unload VehOpt
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVirtClearDrag                  *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear drag when drop on illegal*
'*                      object                         *
'*                                                     *
'*******************************************************
Private Sub mVirtClearDrag()
    imDragIndexSrce = -1
    imDragSrce = -1
    lacVehFrame.Visible = False
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVirtEnableBox                  *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mVirtEnableBox(ilBoxNo As Integer)
'
'   mVirtEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBVirtCtrls) Or (ilBoxNo > UBound(tmVirtCtrls)) Then
        lacVehFrame.Visible = False
        Exit Sub
    End If
    If (imVirtRowNo < vbcVehicle.Value) Or (imVirtRowNo >= vbcVehicle.Value + vbcVehicle.LargeChange + 1) Then
        lacVehFrame.Visible = False
        Exit Sub
    End If
    lacVehFrame.Move 0, tmVirtCtrls(VEHINDEX).fBoxY + (imVirtRowNo - vbcVehicle.Value) * (fgBoxGridH + 15) - 30
    lacVehFrame.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case NOSPOTSINDEX '# of spots
            imIgnoreChg = True
            edcDropDown.Width = tmVirtCtrls(NOSPOTSINDEX).fBoxW
            edcDropDown.MaxLength = 3
            gMoveTableCtrl pbcVehicle, edcDropDown, tmVirtCtrls(NOSPOTSINDEX).fBoxX, tmVirtCtrls(NOSPOTSINDEX).fBoxY + (imVirtRowNo - vbcVehicle.Value) * (fgBoxGridH + 15)
            edcDropDown.Text = smVirtSave(2, imVirtRowNo)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
            imIgnoreChg = False
        Case PERCENTINDEX '% of $'s
            imIgnoreChg = True
            edcDropDown.Width = tmVirtCtrls(PERCENTINDEX).fBoxW
            edcDropDown.MaxLength = 9
            gMoveTableCtrl pbcVehicle, edcDropDown, tmVirtCtrls(PERCENTINDEX).fBoxX, tmVirtCtrls(PERCENTINDEX).fBoxY + (imVirtRowNo - vbcVehicle.Value) * (fgBoxGridH + 15)
            edcDropDown.Text = smVirtSave(3, imVirtRowNo)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
            imIgnoreChg = False
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVirtMoveCtrlToRec              *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mVirtMoveCtrlToRec()
'
'   mVirtMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
    Dim ilLoop As Integer
    Dim slStr As String

    For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
        tmVsf.iFSCode(ilLoop) = 0
        tmVsf.iNoSpots(ilLoop) = 0
        'slStr = ""
        'gStrToPDN slStr, 4, 4, tmVsf.sFSComm(ilLoop)
        tmVsf.lFSComm(ilLoop) = 0
    Next ilLoop
    For ilLoop = LBONE To UBound(smVirtSave, 2) - 1 Step 1
        tmVsf.iFSCode(ilLoop - 1) = imVirtSave(1, ilLoop)
        tmVsf.iNoSpots(ilLoop - 1) = Val(smVirtSave(2, ilLoop))
        slStr = smVirtSave(3, ilLoop)
        'gStrToPDN slStr, 4, 4, tmVsf.sFSComm(ilLoop - 1)
        tmVsf.lFSComm(ilLoop - 1) = gStrDecToLong(slStr, 4)
    Next ilLoop
    Exit Sub

    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVirtMoveRecToCtrl              *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mVirtMoveRecToCtrl()
'
'   mVirtMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilRecCode As Integer
    Dim slNameCode As String  'Name and code
    Dim ilRet As Integer    'Return call status
    Dim slName As String
    Dim slCode As String    'Sales source code number
    Dim ilCode As Integer
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilUpper As Integer
    ReDim smVirtSave(0 To 3, 0 To 1) As String
    ReDim imVirtSave(0 To 1, 0 To 1) As Integer
    ReDim smVirtShow(0 To 3, 0 To 1) As String
    lacVehMsg.Visible = False
    ilUpper = 1
    imOrigNoVehicles = 0
    For ilIndex = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
        If tmVsf.iFSCode(ilIndex) > 0 Then
            ilRecCode = tmVsf.iFSCode(ilIndex)
            For ilLoop = 0 To UBound(tmVehNamesCode) - 1 Step 1 'lbcVehNamesCode.ListCount - 1 Step 1
                slNameCode = tmVehNamesCode(ilLoop).sKey   'lbcVehNamesCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mVirtMoveRecToCtrlErr
                gCPErrorMsg ilRet, "mVirtMoveRecToCtrl (gParseItem field 2)", VehOpt
                On Error GoTo 0
                If ilRecCode = Val(slCode) Then
                    imVirtSave(1, ilUpper) = tmVsf.iFSCode(ilIndex)
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    ilRet = gParseItem(slName, 3, "|", slName)
                    smVirtSave(1, ilUpper) = slName
                    gSetShow pbcVehicle, slName, tmVirtCtrls(VEHINDEX)
                    smVirtShow(1, ilUpper) = tmVirtCtrls(VEHINDEX).sShow
                    smVirtSave(2, ilUpper) = Trim$(Str$(tmVsf.iNoSpots(ilIndex)))
                    gSetShow pbcVehicle, smVirtSave(2, ilUpper), tmVirtCtrls(NOSPOTSINDEX)
                    smVirtShow(2, ilUpper) = tmVirtCtrls(NOSPOTSINDEX).sShow
                    'gPDNToStr tmVsf.sFSComm(ilIndex), 4, slStr
                    slStr = gLongToStrDec(tmVsf.lFSComm(ilIndex), 4)
                    smVirtSave(3, ilUpper) = slStr
                    gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
                    gSetShow pbcVehicle, slStr, tmVirtCtrls(PERCENTINDEX)
                    smVirtShow(3, ilUpper) = tmVirtCtrls(PERCENTINDEX).sShow
                    lbcVehNames.RemoveItem ilLoop
                    'lbcVehNamesCode.RemoveItem ilLoop
                    gRemoveItemFromSortCode ilLoop, tmVehNamesCode()
                    ilUpper = ilUpper + 1
                    ReDim Preserve smVirtSave(0 To 3, 0 To ilUpper) As String
                    ReDim Preserve imVirtSave(0 To 1, 0 To ilUpper) As Integer
                    ReDim Preserve smVirtShow(0 To 3, 0 To ilUpper) As String
                    imOrigNoVehicles = imOrigNoVehicles + 1
                    Exit For
                End If
            Next ilLoop
        End If
    Next ilIndex
    'Determine vehicle associated with virtual vehicle so it can be determined
    'if virtual vehicle referenced
    If imOrigNoVehicles > 0 Then
        ilCode = tmVef.iCode
        ilRet = gIICodeRefExist(VehOpt, ilCode, "Chf.Btr", "ChfVefCode")    'chfvefCode
        If Not ilRet Then
            imOrigNoVehicles = 0
        Else
            'lacVehMsg.Caption = "Virtual Vehicle can't be altered as it is referenced by a Contract"
            lacVehMsg.Caption = "Virtual Vehicle shouldn't be altered as it is referenced by a Contract without Counterpoint Approval"
            imOrigNoVehicles = 0    'This line should be removed when vehicles can't be altered
            lacVehMsg.Visible = True
        End If
    End If
    lbcVehNames.TopIndex = 0
    imVirtSettingValue = True
    vbcVehicle.Min = LBONE  'LBound(smVirtShow, 2)
    If UBound(smVirtShow, 2) - 1 <= vbcVehicle.LargeChange Then
        vbcVehicle.Max = LBONE  'LBound(smVirtShow, 2)
    Else
        vbcVehicle.Max = UBound(smVirtShow, 2) - vbcVehicle.LargeChange '- 1
    End If
    vbcVehicle.Value = vbcVehicle.Min
    If imVirtSettingValue Then
        pbcVehicle.Cls
        pbcVehicle_Paint
        imVirtSettingValue = False
    End If
    imVirtChgVeh = False
    Exit Sub
mVirtMoveRecToCtrlErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVirtSaveRec                    *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mVirtSaveRec() As Integer
'
'   iRet = mVirtSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRet As Integer
    Dim slMsg As String
    Do  'Loop until record updated or added
        If Not mVsfReadRec(tmVef.lVsfCode, SETFORWRITE) Then
            GoTo mVirtSaveRecErr
        End If
        mVirtMoveCtrlToRec
        ilRet = btrUpdate(hmVsf, tmVsf, imVsfRecLen)
        slMsg = "mVirtSaveRec (btrUpdate)"
    Loop While ilRet = BTRV_ERR_CONFLICT
    imVirtChgVeh = False
    mVirtSaveRec = True
    Exit Function
mVirtSaveRecErr:
    On Error GoTo 0
    MousePointer = vbDefault
    imTerminate = True
    mVirtSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVirtSetFocus                       *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mVirtSetFocus(ilBoxNo As Integer)
'
'   mVirtSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    lacVehFrame.Visible = False
    If (ilBoxNo < imLBVirtCtrls) Or (ilBoxNo > UBound(tmVirtCtrls)) Then
        Exit Sub
    End If
    If (imVirtRowNo < vbcVehicle.Value) Or (imVirtRowNo >= vbcVehicle.Value + vbcVehicle.LargeChange + 1) Then
        Exit Sub
    End If
    lacVehFrame.Move 0, tmVirtCtrls(VEHINDEX).fBoxY + (imVirtRowNo - vbcVehicle.Value) * (fgBoxGridH + 15) - 30
    lacVehFrame.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case NOSPOTSINDEX '# of spots
            edcDropDown.SetFocus
        Case PERCENTINDEX '% of $'s
            edcDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVirtSetShow                    *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mVirtSetShow(ilBoxNo As Integer)
'
'   mVirtSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    lacVehFrame.Visible = False
    If (ilBoxNo < imLBVirtCtrls) Or (ilBoxNo > UBound(tmVirtCtrls)) Then
        Exit Sub
    End If
    If (imVirtRowNo < vbcVehicle.Value) Or (imVirtRowNo >= vbcVehicle.Value + vbcVehicle.LargeChange + 1) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NOSPOTSINDEX
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smVirtSave(2, imVirtRowNo) = slStr
            gSetShow pbcVehicle, slStr, tmVirtCtrls(ilBoxNo)
            smVirtShow(2, imVirtRowNo) = tmVirtCtrls(ilBoxNo).sShow
        Case PERCENTINDEX
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smVirtSave(3, imVirtRowNo) = slStr
            gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            gSetShow pbcVehicle, slStr, tmVirtCtrls(ilBoxNo)
            smVirtShow(3, imVirtRowNo) = tmVirtCtrls(ilBoxNo).sShow
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVirtTestFields                 *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if field defined           *
'*                                                     *
'*******************************************************
Private Function mVirtTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mVirtTestFields(iTest, iState)
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
    Dim slMess As String    'Message string
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilLoop As Integer
    Dim slTotalPct As String
    Dim slStr As String
    Dim ilSRow As Integer
    Dim ilERow As Integer
    If (ilCtrlNo = TESTALLCTRLS) Then
        ilSRow = LBONE  'LBound(smVirtSave, 2)
        ilERow = UBound(smVirtSave, 2) - 1
    Else
        ilSRow = imVirtRowNo
        ilERow = imVirtRowNo
    End If
    slTotalPct = "0"
    For ilLoop = ilSRow To ilERow Step 1
        If (ilCtrlNo = NOSPOTSINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedStr(smVirtSave(2, ilLoop), "", slMess, tmVirtCtrls(NOSPOTSINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imVirtBoxNo = NOSPOTSINDEX
                    imVirtRowNo = ilLoop
                End If
                mVirtTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = PERCENTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedStr(smVirtSave(3, ilLoop), "", slMess, tmVirtCtrls(PERCENTINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imVirtBoxNo = PERCENTINDEX
                    imVirtRowNo = ilLoop
                End If
                mVirtTestFields = NO
                Exit Function
            End If
            slStr = smVirtSave(3, ilLoop)
            slTotalPct = gAddStr(slStr, slTotalPct)
        End If
    Next ilLoop
    slMess = "Percent must be specified"
    If (ilCtrlNo = TESTALLCTRLS) And (ilState = ALLMANDEFINED + SHOWMSG) Then
        If Val(slTotalPct) <> 100 Then
            If (ilState And SHOWMSG) = SHOWMSG Then
                ilRes = MsgBox("Percent total must equal 100", vbOKOnly + vbExclamation, "Error")
            End If
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imVirtBoxNo = PERCENTINDEX
                imVirtRowNo = 1
            End If
            mVirtTestFields = NO
            Exit Function
        End If
    End If
    mVirtTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVirtVehPop                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the Vehicle list      *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVirtVehPop()
    Dim ilRet As Integer
    'ilRet = gPopUserVehicleBox(VehOpt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH, lbcVehNames, lbcVehNamesCode)
    ilRet = gPopUserVehicleBox(VehOpt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH, lbcVehNames, tmVehNamesCode(), smVehNamesCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVirtVehPopErr
        gCPErrorMsg ilRet, "mVirtVehPop (gPopUserVehicleBox)", VehOpt
        On Error GoTo 0
    End If
    Exit Sub
mVirtVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVofReadRec                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mVofReadRec(ilVefCode As Integer, slType As String) As Integer
'
'   iRet = mVofReadRec()
'   Where:
'       ilVefCode(I) - Vehicle Code
'       slType (I) - L=Log; C=CP and O= Other
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    tmVofSrchKey.iVefCode = ilVefCode
    tmVofSrchKey.sType = slType
    ilRet = btrGetEqual(hmVof, tmVof, imVofRecLen, tmVofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        'Add Record
        tmVof.iVefCode = ilVefCode  'tmVef.iCode
        tmVof.sType = slType
        tmVof.lHd1CefCode = 0
        tmVof.lFt1CefCode = 0
        tmVof.lFt2CefCode = 0
        tmVof.iNoDaysCP = 30
        tmVof.sShowLen = "Y"
        tmVof.sShowProduct = "Y"
        tmVof.sShowCreative = "N"
        tmVof.sShowISCI = "Y"
        tmVof.sShowDP = "N"
        tmVof.sShowAirTime = "Y"
        tmVof.sShowAirLine = "N"
        tmVof.sShowHour = "N"
        tmVof.sSkipPage = "Y"
        tmVof.iLoadFactor = 1
        ilRet = btrInsert(hmVof, tmVof, imVofRecLen, INDEXKEY0)
        On Error GoTo mVofReadRecErr
        gBtrvErrorMsg ilRet, "mVofReadRec (btrInsert)", VehOpt
        On Error GoTo 0
    End If
    If slType = "L" Then
        tmLVof = tmVof
    ElseIf slType = "C" Then
        tmCVof = tmVof
    ElseIf slType = "O" Then
        tmOVof = tmVof
    End If
    mVofReadRec = True
    Exit Function
mVofReadRecErr:
    On Error GoTo 0
    mVofReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVffReadRec                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mVffReadRec(ilVefCode As Integer) As Integer
'
'   iRet = mVffReadRec()
'   Where:
'       ilVefCode(I) - Vehicle Code
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    imVbfChg = False
    If (tmVef.sType <> "C") And (tmVef.sType <> "S") And (tmVef.sType <> "A") And (tmVef.sType <> "P") And (tmVef.sType <> "G") And (tmVef.sType <> "L") And (tmVef.sType <> "R") And (tmVef.sType <> "N") Then
        mVffReadRec = False
        Exit Function
    End If
    If ilVefCode > 0 Then
        tmVffSrchKey1.iCode = ilVefCode
        ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, tmVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Else
        mVffReadRec = False
        Exit Function
    End If
    If ilRet <> BTRV_ERR_NONE Then
        'Add Record
        tmVff.iCode = 0
        tmVff.iVefCode = ilVefCode  'tmVef.iCode
        tmVff.sGroupName = ""
        tmVff.sWegenerExportID = ""
        tmVff.sOLAExportID = ""
        tmVff.iLiveCompliantAdj = 5
        tmVff.iUstCode = 0
        tmVff.iUrfCode = tgUrf(0).iCode
        'tmVff.sXDXMLForm = "S"
        If ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
            tmVff.sXDXMLForm = "S"
        Else
            If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) Then
                tmVff.sXDXMLForm = "P"
            Else
                tmVff.sXDXMLForm = ""
            End If
        End If
        tmVff.sXDISCIPrefix = ""
        tmVff.sXDProgCodeID = ""
        tmVff.sXDSaveCF = "Y"
        tmVff.sXDSaveHDD = "N"
        tmVff.sXDSaveNAS = "N"
        'tmVff.sUnused = ""
        tmVff.iCwfCode = 0
        tmVff.sAirWavePrgID = ""
        tmVff.sExportAirWave = ""
        tmVff.sExportNYESPN = ""
        tmVff.sPledgeVsAir = "N"
        tmVff.sFedDelivery(0) = ""
        tmVff.sFedDelivery(1) = ""
        tmVff.sFedDelivery(2) = ""
        tmVff.sFedDelivery(3) = ""
        tmVff.sFedDelivery(4) = ""
        'tmVff.sFedDelivery(5) = ""
        gPackDate "1/1/1990", tmVff.iLastAffExptDate(0), tmVff.iLastAffExptDate(1)
        tmVff.sMoveSportToNon = "N"
        tmVff.sMoveSportToSport = "N"
        tmVff.sMoveNonToSport = "N"
        tmVff.sMergeTraffic = "S"
        tmVff.sMergeAffiliate = "S"
        tmVff.sMergeWeb = "S"
        tmVff.sPledgeByEvent = "N"
        tmVff.lPledgeHdVtfCode = 0
        tmVff.lPledgeFtVtfCode = 0
        tmVff.iPledgeClearance = 0
        tmVff.sExportEncoESPN = "N"
        tmVff.sWebName = ""
        tmVff.lSeasonGhfCode = 0
        tmVff.iMcfCode = 0
        tmVff.sExportAudio = "N"
        tmVff.sExportMP2 = "N"
        tmVff.sExportCnCSpot = "N"
        tmVff.sExportEnco = "N"
        tmVff.sExportCnCNetInv = "N"
        tmVff.sIPumpEventTypeOV = ""
        tmVff.sExportIPump = "N"
        tmVff.sAddr4 = ""
        tmVff.lBBOpenCefCode = 0
        tmVff.lBBCloseCefCode = 0
        tmVff.lBBBothCefCode = 0
        tmVff.sXDSISCIPrefix = ""
        tmVff.sXDSSaveCF = "Y"
        tmVff.sXDSSaveHDD = "N"
        tmVff.sXDSSaveNAS = "N"
        tmVff.sMGsOnWeb = "N"   '"Y"
        tmVff.sReplacementOnWeb = "N"   'Y"
        tmVff.sExportMatrix = "N"
        tmVff.sSentToXDSStatus = "N"
        tmVff.sStationComp = "N"
        tmVff.sExportSalesForce = "N"
        tmVff.sExportEfficio = "N"
        tmVff.sExportJelli = "N"
        tmVff.sOnXMLInsertion = "N"
        tmVff.sOnInsertions = "N"
        tmVff.sPostLogSource = "N"
        tmVff.sExportTableau = "N"
        tmVff.sStationPassword = ""
        tmVff.sHonorZeroUnits = "N"
        tmVff.sHideCommOnLog = "N"
        tmVff.sHideCommOnWeb = "N"
        tmVff.iConflictWinLen = 0
        tmVff.sACT1LineupCode = ""
        tmVff.sPrgmmaticAllow = "N"
        tmVff.sSalesBrochure = ""
        tmVff.sCartOnWeb = "N"
        tmVff.sDefaultAudioType = "R"
        tmVff.iLogExptArfCode = 0
        'tmVff.sUnused = ""
        tmVff.sASICallLetters = ""
        tmVff.sASIBand = ""
        tmVff.sExportCustom = "" 'TTP 9992
        '10933
        tmVff.sXDEventZone = ""
        ilRet = btrInsert(hmVff, tmVff, imVffRecLen, INDEXKEY0)
        On Error GoTo mVffReadRecErr
        gBtrvErrorMsg ilRet, "mVffReadRec (btrInsert)", VehOpt
        On Error GoTo 0
    End If
    mVffReadRec = True
    Exit Function
mVffReadRecErr:
    On Error GoTo 0
    mVffReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVsfReadRec                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mVsfReadRec(llVsfCode As Long, ilForUpdate As Integer) As Integer
'
'   iRet = mVsfReadRec()
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    tmVsfSrchKey.lCode = llVsfCode
    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mVsfReadRecErr
    gBtrvErrorMsg ilRet, "mVsfReadRec (btrGetEqual)", VehOpt
    On Error GoTo 0
    mVsfReadRec = True
    Exit Function
mVsfReadRecErr:
    On Error GoTo 0
    mVsfReadRec = False
    Exit Function
End Function




Private Sub pbcAcq_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mRemoveFocus
End Sub

Private Sub pbcAcq_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilIndex As Integer
    
    mSetCommands
    If Index = 0 Then
        If imVBFIndex < 0 Then
            Exit Sub
        End If
        ilIndex = (X - imAcqCostX) \ (edcBarter(3).Width + 15)
        If Y < imAcqCostY + edcBarter(3).Height Then
            imAcqCostBoxNo = ilIndex
        Else
            imAcqCostBoxNo = ilIndex + (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD)
        End If
        mSetAcqCost
    ElseIf Index = 1 Then
        ilIndex = (X - imAcqIndexX) \ (edcBarter(4).Width + 15)
        If Y < imAcqIndexY + edcBarter(4).Height Then
            imAcqIndexBoxNo = ilIndex
        Else
            imAcqIndexBoxNo = -1
        End If
        mSetAcqIndex
    End If
End Sub

Private Sub pbcAcq_Paint(Index As Integer)
    If Index = 0 Then
        mAcqCostPaint
    ElseIf Index = 1 Then
        mAcqIndexPaint
    End If
End Sub

Private Sub pbcAcqSTab_GotFocus(Index As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    If GetFocus() <> pbcAcqSTab(Index).HWnd Then
        Exit Sub
    End If
    If tmVef.sType = "R" Then
        If imAcqCostBoxNo = -1 Then
            edcBarter(3).Visible = True
            imAcqCostBoxNo = 0
        Else
            mSetCommands
            If imAcqCostBoxNo = 0 Then
                edcBarter(3).Visible = False
                imAcqCostBoxNo = -1
                edcBarter(0).SetFocus
                Exit Sub
            End If
            imAcqCostBoxNo = imAcqCostBoxNo - 1
        End If
        mSetAcqCost
    ElseIf (tmVef.sType = "C") Or (tmVef.sType = "S") Then
        If imAcqIndexBoxNo = -1 Then
            edcBarter(4).Visible = True
            imAcqIndexBoxNo = 0
        Else
            mSetCommands
            If imAcqIndexBoxNo = 0 Then
                edcBarter(4).Visible = False
                imAcqIndexBoxNo = -1
                'edcBarter(0).SetFocus
                Exit Sub
            End If
            imAcqIndexBoxNo = imAcqIndexBoxNo - 1
        End If
        mSetAcqIndex
    End If
End Sub

Private Sub pbcAcqTab_GotFocus(Index As Integer)
    If GetFocus() <> pbcAcqTab(Index).HWnd Then
        Exit Sub
    End If
    If Index = 0 Then
        If imAcqCostBoxNo = -1 Then
            edcBarter(3).Visible = True
            imAcqCostBoxNo = 2 * (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD) - 1
        Else
            mSetCommands
            If imAcqCostBoxNo = 2 * (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD) - 1 Then
                edcBarter(3).Visible = False
                imAcqCostBoxNo = -1
                edcInsertComment.SetFocus
                Exit Sub
            End If
            imAcqCostBoxNo = imAcqCostBoxNo + 1
        End If
        mSetAcqCost
    ElseIf Index = 1 Then
        If imAcqIndexBoxNo = -1 Then
            edcBarter(4).Visible = True
            imAcqIndexBoxNo = (UBound(tmVbfIndex.lDefAcqCost) + ADJBD) - 1
        Else
            mSetCommands
            If imAcqIndexBoxNo = (UBound(tmVbfIndex.lDefAcqCost) + ADJBD) - 1 Then
                edcBarter(4).Visible = False
                imAcqIndexBoxNo = -1
                'edcInsertComment.SetFocus
                Exit Sub
            End If
            imAcqIndexBoxNo = imAcqIndexBoxNo + 1
        End If
        mSetAcqIndex
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    mRemoveFocus
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcCommEmbedded_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcCommEmbedded_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        If tmVpf.sEmbeddedComm <> "Y" Then
            imProducerAltered = True
        End If
        tmVpf.sEmbeddedComm = "Y"
        pbcCommEmbedded_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If tmVpf.sEmbeddedComm <> "N" Then
            imProducerAltered = True
        End If
        tmVpf.sEmbeddedComm = "N"
        pbcCommEmbedded_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If tmVpf.sEmbeddedComm = "Y" Then
            imProducerAltered = True
            tmVpf.sEmbeddedComm = "N"
            pbcCommEmbedded_Paint
        ElseIf tmVpf.sEmbeddedComm = "N" Then
            imProducerAltered = True
            tmVpf.sEmbeddedComm = "Y"
            pbcCommEmbedded_Paint
        Else
            imProducerAltered = True
            tmVpf.sEmbeddedComm = "Y"
            pbcCommEmbedded_Paint
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcCommEmbedded_LostFocus()
    pbcCommEmbedded.Visible = False
End Sub

Private Sub pbcCommEmbedded_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmVpf.sEmbeddedComm = "Y" Then
        imProducerAltered = True
        tmVpf.sEmbeddedComm = "N"
        pbcCommEmbedded_Paint
    ElseIf tmVpf.sEmbeddedComm = "N" Then
        imProducerAltered = True
        tmVpf.sEmbeddedComm = "Y"
        pbcCommEmbedded_Paint
    Else
        imProducerAltered = True
        tmVpf.sEmbeddedComm = "Y"
        pbcCommEmbedded_Paint
    End If
    mSetCommands
End Sub

Private Sub pbcCommEmbedded_Paint()
    pbcCommEmbedded.Cls
    pbcCommEmbedded.CurrentX = fgBoxInsetX
    pbcCommEmbedded.CurrentY = 0 'fgBoxInsetY
    If tmVpf.sEmbeddedComm = "Y" Then
        pbcCommEmbedded.Print "Yes"
    Else
        pbcCommEmbedded.Print "No"
    End If

End Sub

Private Sub pbcGTZ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    flX = imTZX
    flY = imTZY
    '3/1/13: Allow zone table to be filled in so that the BR will adjust the times to ET when overrides defined
    'If (tmVef.sType = "S") Then
    '    ilMaxCol = 1
    'Else
        ilMaxCol = imTZMaxCtrls
    'End If
    For ilCol = 1 To ilMaxCol Step 1
        If (X >= tmTZCtrls(ilCol).fBoxX) And (X <= tmTZCtrls(ilCol).fBoxX + tmTZCtrls(ilCol).fBoxW) Then
            For ilRow = 0 To 4 Step 1
                If (Y >= flY) And (Y <= flY + imTZH - 15) Then
                    If (ilCol > 1) And (Trim$(tmVpf.sGZone(ilRow)) = "") Then
                        Beep
                        Exit Sub
                    End If
                    imGTZBoxNo = imTZMaxCtrls * ilRow + ilCol
                    mSetGTZ
                    Exit Sub
                End If
                flY = flY + imTZH - 15
            Next ilRow
        End If
        flY = imTZY
    Next ilCol
End Sub
Private Sub pbcGTZ_Paint()
    mGTZPaint tmVpf, tmVff
End Sub
Private Sub pbcGTZSTab_GotFocus()
    Dim ilIndex As Integer
    If GetFocus() <> pbcGTZSTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to Left
    If imGTZBoxNo = -1 Then
        imGTZBoxNo = 1
        imTabDirection = 0  'Set-Left to right
    Else
        pbcGTZToggle.Visible = False
        lbcFeed.Visible = False
        cmcGTZDropDown.Visible = False
        edcGTZDropDown.Visible = False
        DoEvents
        If (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDINDEX Then
            If mFeedBranch() Then
                Exit Sub
            End If
        End If
        If imGTZBoxNo = 1 Then
            imGTZBoxNo = -1
            If rbcGMedium(0).Value Then
                rbcGMedium(0).SetFocus
            ElseIf rbcGMedium(1).Value Then
                rbcGMedium(1).SetFocus
            ElseIf rbcGMedium(2).Value Then
                rbcGMedium(2).SetFocus
            ElseIf rbcGMedium(3).Value Then
                rbcGMedium(3).SetFocus
            ElseIf rbcGMedium(4).Value Then
                rbcGMedium(4).SetFocus
            ElseIf rbcGMedium(5).Value Then
                rbcGMedium(5).SetFocus
            ElseIf rbcGMedium(PODCASTRBC).Value Then         '1-16-14
                rbcGMedium(PODCASTRBC).SetFocus
            Else
                'edcGSignOn.SetFocus
                edcGen(Signon).SetFocus
            End If
            Exit Sub
        End If
        ilIndex = ((imGTZBoxNo - 1) \ imTZMaxCtrls) + 1
        '3/1/13: Allow table table to be fill in for selling vehicles
        'If (tmVef.sType = "S") Or (Trim$(tmVpf.sGZone(ilIndex)) = "") Then
        If (Trim$(tmVpf.sGZone(ilIndex - 1)) = "") Then
            imGTZBoxNo = imGTZBoxNo - imTZMaxCtrls
        Else
            imGTZBoxNo = imGTZBoxNo - 1
        End If
    End If
    mSetGTZ
End Sub
Private Sub pbcGTZTab_GotFocus()
    Dim ilIndex As Integer
    If GetFocus() <> pbcGTZTab.HWnd Then
        Exit Sub
    End If
        imTabDirection = 0  'Set-Left to right
    If imGTZBoxNo = -1 Then
        '3/1/13: Allow table table to be fill in for selling vehicles
        'If tmVef.sType = "S" Then
        '    imGTZBoxNo = imTZMaxCtrls * (UBound(tmVpf.sGZone) - 1) + 1
        'Else
            If Trim$(tmVpf.sGZone(UBound(tmVpf.sGZone))) <> "" Then
                imGTZBoxNo = imTZMaxCtrls * (UBound(tmVpf.sGZone))
            Else
                imGTZBoxNo = imTZMaxCtrls * (UBound(tmVpf.sGZone) - 1) + 1
            End If
        'End If
        imTabDirection = -1  'Set-Right to left
    Else
        If (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDINDEX Then
            If mFeedBranch() Then
                Exit Sub
            End If
        End If
        ilIndex = ((imGTZBoxNo - 1) \ imTZMaxCtrls) + 1
        If ((imGTZBoxNo = imTZMaxCtrls * (UBound(tmVpf.sGZone))) And (tmVef.sType <> "S")) Or ((imGTZBoxNo = imTZMaxCtrls * (UBound(tmVpf.sGZone) - 1) + 1) And (tmVef.sType = "S")) Or ((imGTZBoxNo = imTZMaxCtrls * (UBound(tmVpf.sGZone) - 1) + 1) And (tmVef.sType <> "S") And Trim$(tmVpf.sGZone(ilIndex - 1)) = "") Then
            pbcGTZToggle.Visible = False
            lbcFeed.Visible = False
            cmcGTZDropDown.Visible = False
            edcGTZDropDown.Visible = False
            imGTZBoxNo = -1
            '10064
'            If (tmVef.sType <> "S") And (tmVef.sType <> "A") Then
'                cmcDone.SetFocus
'            Else
'                edcGen(SAGROUPNO).SetFocus
'            End If
            Exit Sub
        End If
        '3/1/13: Allow table table to be fill in for selling vehicles
        'If (tmVef.sType = "S") Or (Trim$(tmVpf.sGZone(ilIndex)) = "") Then
        If (Trim$(tmVpf.sGZone(ilIndex - 1)) = "") Then
            imGTZBoxNo = imGTZBoxNo + imTZMaxCtrls
        Else
            imGTZBoxNo = imGTZBoxNo + 1
        End If
    End If
    mSetGTZ
End Sub
Private Sub pbcGTZToggle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcGTZToggle_KeyPress(KeyAscii As Integer)
    Dim ilIndex As Integer
    ilIndex = ((imGTZBoxNo - 1) \ imTZMaxCtrls) + 1
    If (imGTZBoxNo Mod imTZMaxCtrls) = GCMMLSCHINDEX Then  'Version for cmml schd
        If (KeyAscii = Asc("A")) Or (KeyAscii = Asc("a")) Then
            tmVpf.sGCSVer(ilIndex - 1) = "A"
            pbcGTZToggle_Paint
        ElseIf KeyAscii = Asc("O") Or (KeyAscii = Asc("o")) Then
            tmVpf.sGCSVer(ilIndex - 1) = "O"
            pbcGTZToggle_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If tmVpf.sGCSVer(ilIndex - 1) = "A" Then
                tmVpf.sGCSVer(ilIndex - 1) = "O"
                pbcGTZToggle_Paint
            ElseIf tmVpf.sGCSVer(ilIndex - 1) = "O" Then
                tmVpf.sGCSVer(ilIndex - 1) = "A"
                pbcGTZToggle_Paint
            End If
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEDDELIVERYINDEX Then  'Fed (Yes or No)
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            tmVff.sFedDelivery(ilIndex - 1) = "Y"
            pbcGTZToggle_Paint
            imVffChg = True
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            tmVff.sFedDelivery(ilIndex - 1) = "N"
            pbcGTZToggle_Paint
            imVffChg = True
        End If
        If KeyAscii = Asc(" ") Then
            If tmVff.sFedDelivery(ilIndex - 1) = "Y" Then
                tmVff.sFedDelivery(ilIndex - 1) = "N"
                pbcGTZToggle_Paint
                imVffChg = True
            ElseIf tmVff.sFedDelivery(ilIndex - 1) = "N" Then
                tmVff.sFedDelivery(ilIndex - 1) = "Y"
                pbcGTZToggle_Paint
                imVffChg = True
            End If
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcGTZToggle_LostFocus()
    pbcGTZToggle.Visible = False
End Sub
Private Sub pbcGTZToggle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilIndex As Integer
    ilIndex = ((imGTZBoxNo - 1) \ imTZMaxCtrls) + 1
    If (imGTZBoxNo Mod imTZMaxCtrls) = GCMMLSCHINDEX Then  'Version for cmml schd
        If tmVpf.sGCSVer(ilIndex - 1) = "A" Then
            tmVpf.sGCSVer(ilIndex - 1) = "O"
            pbcGTZToggle_Paint
        ElseIf tmVpf.sGCSVer(ilIndex - 1) = "O" Then
            tmVpf.sGCSVer(ilIndex - 1) = "A"
            pbcGTZToggle_Paint
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEDDELIVERYINDEX Then  'Fed (Yes or No)
        If tmVff.sFedDelivery(ilIndex - 1) = "Y" Then
            tmVff.sFedDelivery(ilIndex - 1) = "N"
            pbcGTZToggle_Paint
            imVffChg = True
        ElseIf tmVff.sFedDelivery(ilIndex - 1) = "N" Then
            tmVff.sFedDelivery(ilIndex - 1) = "Y"
            pbcGTZToggle_Paint
            imVffChg = True
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcGTZToggle_Paint()
    Dim ilIndex
    pbcGTZToggle.Cls
    pbcGTZToggle.CurrentX = fgBoxInsetX
    pbcGTZToggle.CurrentY = 0 'fgBoxInsetY
    ilIndex = ((imGTZBoxNo - 1) \ imTZMaxCtrls) + 1
    If (imGTZBoxNo Mod imTZMaxCtrls) = GCMMLSCHINDEX Then  'Version for cmml schd
        If tmVpf.sGCSVer(ilIndex - 1) = "A" Then
            pbcGTZToggle.Print "All Versions"
        Else
            pbcGTZToggle.Print "Original"
        End If
    ElseIf (imGTZBoxNo Mod imTZMaxCtrls) = GFEDDELIVERYINDEX Then  'Fed (Yes or No)
        If tmVff.sFedDelivery(ilIndex - 1) = "Y" Then
            pbcGTZToggle.Print "Yes"
        Else
            pbcGTZToggle.Print "No"
        End If
    End If
End Sub

Private Sub pbcLevelSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcLevelSTab.HWnd Then
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
                cmcDone.SetFocus
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
    If GetFocus() <> pbcLevelTab.HWnd Then
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
                cmcDone.SetFocus
                Exit Sub
            Case Else
                ilBox = imSBoxNo + 1
                ilFound = True
        End Select
    Loop While Not ilFound
    imSBoxNo = ilBox
    mSEnableBox ilBox
End Sub

Private Sub pbcLogForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim ilBox As Integer

    For ilRow = 1 To 3 Step 1
        For ilBox = imLBLCtrls To UBound(tmLCtrls) Step 1
            If (X >= tmLCtrls(ilBox).fBoxX) And (X <= (tmLCtrls(ilBox).fBoxX + tmLCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmLCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmLCtrls(ilBox).fBoxY + tmLCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow
                    mLSetShow imLBoxNo
                    imLRowNo = ilRowNo
                    imLBoxNo = ilBox
                    mLEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mLSetFocus imLBoxNo
End Sub
Private Sub pbcLogForm_Paint()
    Dim ilRow As Integer
    Dim ilBox As Integer
    Dim slStr As String
    For ilRow = 1 To 3 Step 1
        For ilBox = LBONE To UBound(smLShow, 1) Step 1
            slStr = smLShow(ilBox, ilRow)
            'gPaintArea pbcVehicle, tmVirtCtrls(ilBox).fBoxX, tmVirtCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmVirtCtrls(ilBox).fBoxW - 15, tmVirtCtrls(ilBox).fBoxH - 15, WHITE
            pbcLogForm.CurrentX = tmLCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcLogForm.CurrentY = tmLCtrls(ilBox).fBoxY + (ilRow - 1) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            pbcLogForm.Print slStr
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcLSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcLSTab.HWnd Then
        Exit Sub
    End If
    mLSetShow imLBoxNo
    ilBox = imLBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imLRowNo = 1
                ilBox = LNODAYSINDEX
                imLBoxNo = ilBox
                mLEnableBox ilBox
                Exit Sub
            Case LNODAYSINDEX 'Time (first control within header)
                If (imLRowNo <= 1) Then
                    imLBoxNo = -1
                    imLRowNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imLRowNo = imLRowNo - 1
                ilBox = LFOOT2INDEX
                ilFound = True
            Case Else
                ilBox = imLBoxNo - 1
                ilFound = True
        End Select
    Loop While Not ilFound
    imLBoxNo = ilBox
    mLEnableBox ilBox
End Sub
Private Sub pbcLTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcLTab.HWnd Then
        Exit Sub
    End If
    mLSetShow imLBoxNo
    ilBox = imLBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imLRowNo = 3
                ilBox = LFOOT2INDEX
                imLBoxNo = ilBox
                mLEnableBox ilBox
                Exit Sub
            Case LFOOT2INDEX 'Time (first control within header)
                If (imLRowNo >= 3) Then
                    imLBoxNo = -1
                    imLRowNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imLRowNo = imLRowNo + 1
                ilBox = LNODAYSINDEX
                ilFound = True
            Case Else
                ilBox = imLBoxNo + 1
                ilFound = True
        End Select
    Loop While Not ilFound
    imLBoxNo = ilBox
    mLEnableBox ilBox
End Sub
Private Sub pbcLYN_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcLYN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        If imLSave(imLBoxNo, imLRowNo) <> 1 Then
            imLogAltered = True
        End If
        imLSave(imLBoxNo, imLRowNo) = 1
        pbcLYN_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imLSave(imLBoxNo, imLRowNo) <> 0 Then
            imLogAltered = True
        End If
        imLSave(imLBoxNo, imLRowNo) = 0
        pbcLYN_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imLSave(imLBoxNo, imLRowNo) = 1 Then
            imLogAltered = True
            imLSave(imLBoxNo, imLRowNo) = 0
            pbcLYN_Paint
        ElseIf imLSave(imLBoxNo, imLRowNo) = 0 Then
            imLogAltered = True
            imLSave(imLBoxNo, imLRowNo) = 1
            pbcLYN_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcLYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imLSave(imLBoxNo, imLRowNo) = 1 Then
        imLogAltered = True
        imLSave(imLBoxNo, imLRowNo) = 0
        pbcLYN_Paint
    ElseIf imLSave(imLBoxNo, imLRowNo) = 0 Then
        imLogAltered = True
        imLSave(imLBoxNo, imLRowNo) = 1
        pbcLYN_Paint
    End If
    mSetCommands
End Sub
Private Sub pbcLYN_Paint()
    pbcLYN.Cls
    pbcLYN.CurrentX = fgBoxInsetX
    pbcLYN.CurrentY = 0 'fgBoxInsetY
    If imLSave(imLBoxNo, imLRowNo) = 1 Then
        pbcLYN.Print "Yes"
    ElseIf imLSave(imLBoxNo, imLRowNo) = 0 Then
        pbcLYN.Print "No"
    End If
End Sub

Private Sub pbcMerge_Paint(Index As Integer)
    pbcMerge(Index).CurrentX = 0
    pbcMerge(Index).CurrentY = 0
    If Index = 0 Then
        pbcMerge(Index).Print "Traffic"
    ElseIf Index = 1 Then
        pbcMerge(Index).Print "Affiliate"
    ElseIf Index = 2 Then
        pbcMerge(Index).Print "Affidavit"
    End If
End Sub

Private Sub pbcMPromoSTab_GotFocus()
    If GetFocus() <> pbcMPromoSTab.HWnd Then
        Exit Sub
    End If
    imMPSABoxNo = -1
    If imMPromoBoxNo = -1 Then
        imMPromoBoxNo = 1
    Else
        mSetCommands
        If imMPromoBoxNo = 1 Then
            edcMPromo.Visible = False
            imMPromoBoxNo = -1
            pbcMPSATab.SetFocus
            Exit Sub
        End If
        imMPromoBoxNo = imMPromoBoxNo - 1
    End If
    mSetMPromo
End Sub
Private Sub pbcMPromoTab_GotFocus()
    If GetFocus() <> pbcMPromoTab.HWnd Then
        Exit Sub
    End If
    imMPSABoxNo = -1
    If imMPromoBoxNo = -1 Then
        imMPromoBoxNo = 3 * (UBound(tmVpf.iMMFPromo))
    Else
        mSetCommands
        If imMPromoBoxNo = 3 * (UBound(tmVpf.iMMFPromo)) Then
            edcMPromo.Visible = False
            imMPromoBoxNo = -1
            cmcDone.SetFocus
            Exit Sub
        End If
        imMPromoBoxNo = imMPromoBoxNo + 1
    End If
    mSetMPromo
End Sub
Private Sub pbcMPSASTab_GotFocus()
    If GetFocus() <> pbcMPSASTab.HWnd Then
        Exit Sub
    End If
    imMPromoBoxNo = -1
    If imMPSABoxNo = -1 Then
        imMPSABoxNo = 1
    Else
        mSetCommands
        If imMPSABoxNo = 1 Then
            edcMPSA.Visible = False
            imMPSABoxNo = -1
            pbcMPromoSTab.SetFocus
            Exit Sub
        End If
        imMPSABoxNo = imMPSABoxNo - 1
    End If
    mSetMPsa
End Sub
Private Sub pbcMPSATab_GotFocus()
    If GetFocus() <> pbcMPSATab.HWnd Then
        Exit Sub
    End If
    imMPromoBoxNo = -1
    If imMPSABoxNo = -1 Then
        imMPSABoxNo = 3 * (UBound(tmVpf.iMMFPSA))
    Else
        mSetCommands
        If imMPSABoxNo = 3 * (UBound(tmVpf.iMMFPSA)) Then
            edcMPSA.Visible = False
            imMPSABoxNo = -1
            pbcMPromoSTab.SetFocus
            Exit Sub
        End If
        imMPSABoxNo = imMPSABoxNo + 1
    End If
    mSetMPsa
End Sub

Private Sub pbcParticipantSTab_GotFocus()
    Dim ilPrev As Integer

    If GetFocus() <> pbcParticipantSTab.HWnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mPartEnableBox
        Exit Sub
    End If
    If imCtrlVisible Then
        imTabDirection = -1 'Set- right to left
        If grdParticipant.Col = SSOURCEINDEX Then
            If mSSourceBranch() Then
                Exit Sub
            End If
        End If
        If grdParticipant.Col = PARTINDEX Then
            If mVehGpBranch() Then
                Exit Sub
            End If
        End If
        mPartSetShow
        Do
            ilPrev = False
            If grdParticipant.Col = SSOURCEINDEX Then
                If grdParticipant.Row > grdParticipant.FixedRows Then
                    lmPartTopRow = -1
                    grdParticipant.Row = grdParticipant.Row - 1
                    If Not grdParticipant.RowIsVisible(grdParticipant.Row) Then
                        grdParticipant.TopRow = grdParticipant.TopRow - 1
                    End If
                    grdParticipant.Col = PRODPCTINDEX
                    mPartEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                grdParticipant.Col = grdParticipant.Col - 1
                If mPartColOk(grdParticipant.Row, grdParticipant.Col) Then
                    mPartEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        imTabDirection = 0  'Set-Left to right
        lmPartTopRow = -1
        grdParticipant.TopRow = grdParticipant.FixedRows
        grdParticipant.Col = SSOURCEINDEX
        grdParticipant.Row = grdParticipant.FixedRows
        If mPartColOk(grdParticipant.Row, grdParticipant.Col) Then
            mPartEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If

End Sub

Private Sub pbcParticipantTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcParticipantTab.HWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        imTabDirection = 0 'Set- Left to right
        If grdParticipant.Col = SSOURCEINDEX Then
            If mSSourceBranch() Then
                Exit Sub
            End If
        End If
        If grdParticipant.Col = PARTINDEX Then
            If mVehGpBranch() Then
                Exit Sub
            End If
        End If
        llEnableRow = lmPartEnableRow
        mPartSetShow
        Do
            ilNext = False
            If grdParticipant.Col = PRODPCTINDEX Then
                llRow = grdParticipant.Rows
                Do
                    llRow = llRow - 1
                Loop While grdParticipant.TextMatrix(llRow, SSOURCEINDEX) = ""
                llRow = llRow + 1
                If (grdParticipant.Row + 1 < llRow) Then
                    lmPartTopRow = -1
                    grdParticipant.Row = grdParticipant.Row + 1
                    If Not grdParticipant.RowIsVisible(grdParticipant.Row) Or (grdParticipant.Row - (grdParticipant.TopRow - grdParticipant.FixedRows) >= imInitNoRows) Then
                        imIgnoreScroll = True
                        grdParticipant.TopRow = grdParticipant.TopRow + 1
                    End If
                    grdParticipant.Col = SSOURCEINDEX
                    If Trim$(grdParticipant.TextMatrix(grdParticipant.Row, SSOURCEINDEX)) <> "" Then
                        If mPartColOk(grdParticipant.Row, grdParticipant.Col) Then
                            mPartEnableBox
                        Else
                            cmcCancel.SetFocus
                        End If
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdParticipant.Left - pbcArrow.Width - 30, grdParticipant.Top + grdParticipant.RowPos(grdParticipant.Row) + (grdParticipant.RowHeight(grdParticipant.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    If Trim$(grdParticipant.TextMatrix(llEnableRow, SSOURCEINDEX)) <> "" Then
                        lmPartTopRow = -1
                        If grdParticipant.Row + 1 >= grdParticipant.Rows Then
                            grdParticipant.AddItem ""
                            grdParticipant.RowHeight(grdParticipant.Row + 1) = fgBoxGridH + 15
                            grdParticipant.TextMatrix(grdParticipant.Row + 1, PIFCODEINDEX) = 0
                        End If
                        grdParticipant.Row = grdParticipant.Row + 1
                        If (Not grdParticipant.RowIsVisible(grdParticipant.Row)) Or (grdParticipant.Row - (grdParticipant.TopRow - grdParticipant.FixedRows) >= imInitNoRows) Then
                            imIgnoreScroll = True
                            grdParticipant.TopRow = grdParticipant.TopRow + 1
                        End If
                        grdParticipant.Col = SSOURCEINDEX
                        grdParticipant.TextMatrix(grdParticipant.Row, PIFCODEINDEX) = 0
                        'mEnableBox
                        imFromArrow = True
                        pbcArrow.Move grdParticipant.Left - pbcArrow.Width - 30, grdParticipant.Top + grdParticipant.RowPos(grdParticipant.Row) + (grdParticipant.RowHeight(grdParticipant.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                grdParticipant.Col = grdParticipant.Col + 1
                If mPartColOk(grdParticipant.Row, grdParticipant.Col) Then
                    mPartEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        imTabDirection = -1  'Set-Right to left
        lmPartTopRow = -1
        grdParticipant.TopRow = grdParticipant.FixedRows
        grdParticipant.Col = SSOURCEINDEX
        grdParticipant.Row = grdParticipant.FixedRows
        If mPartColOk(grdParticipant.Row, grdParticipant.Col) Then
            mPartEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If

End Sub

Private Sub pbcProducer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMax As Integer

    If tgSpf.sGUseAffSys = "Y" Then
        '1/17/10: Embedded moved to Affiliate Export ISCI
        'ilMax = COMMEMBEDDEDINDEX
        ilMax = EXPCOMMAUDIOINDEX
    Else
        ilMax = PRODUCERINDEX
    End If
    For ilBox = imLBPCtrls To ilMax Step 1
        If (X >= tmPCtrls(ilBox).fBoxX) And (X <= tmPCtrls(ilBox).fBoxX + tmPCtrls(ilBox).fBoxW) Then
            If (Y >= tmPCtrls(ilBox).fBoxY) And (Y <= tmPCtrls(ilBox).fBoxY + tmPCtrls(ilBox).fBoxH) Then
                mPSetShow imPBoxNo
                imPBoxNo = ilBox
                mPEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mPSetFocus imPBoxNo

End Sub

Private Sub pbcProducer_Paint()
    Dim ilBox As Integer
    Dim ilMax As Integer

    If tgSpf.sGUseAffSys = "Y" Then
        '1/17/10: Embedded moved to Affiliate Export ISCI
        'ilMax = COMMEMBEDDEDINDEX
        ilMax = EXPCOMMAUDIOINDEX
    Else
        ilMax = PRODUCERINDEX
    End If
    For ilBox = imLBPCtrls To ilMax Step 1
        pbcProducer.CurrentX = tmPCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcProducer.CurrentY = tmPCtrls(ilBox).fBoxY + fgBoxInsetY
        If ilBox <> COMMEMBEDDEDINDEX Then
            pbcProducer.Print tmPCtrls(ilBox).sShow
        Else
            If tmVpf.sEmbeddedComm = "Y" Then
                pbcProducer.Print "Yes"
            ElseIf tmVpf.sEmbeddedComm = "N" Then
                pbcProducer.Print "No"
            End If
        End If
    Next ilBox
End Sub

Private Sub pbcPromo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim flX As Single
    Dim flY As Single
    flX = imPromoX
    flY = imPromoY
    For ilCol = 0 To 2 Step 1
        If (X >= flX) And (X <= flX + imPromoW) Then
            For ilRow = 1 To 24 Step 1
                If (Y >= flY) And (Y <= flY + imPromoH - 15) Then
                    imMPromoBoxNo = ilRow + ilCol * 24
                    mSetMPromo
                    Exit Sub
                End If
                flY = flY + imPromoH - 15
            Next ilRow
        End If
        flX = flX + imPromoW + 15
        flY = imPromoY
    Next ilCol
End Sub
Private Sub pbcPromo_Paint()
    mMPromoPaint tmVpf
End Sub
Private Sub pbcPSA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim flX As Single
    Dim flY As Single
    flX = imPsaX
    flY = imPsaY
    For ilCol = 0 To 2 Step 1
        If (X >= flX) And (X <= flX + imPsaW) Then
            For ilRow = 1 To 24 Step 1
                If (Y >= flY) And (Y <= flY + imPsaH - 15) Then
                    imMPSABoxNo = ilRow + ilCol * 24
                    mSetMPsa
                    Exit Sub
                End If
                flY = flY + imPsaH - 15
            Next ilRow
        End If
        flX = flX + imPsaW + 15
        flY = imPsaY
    Next ilCol
End Sub
Private Sub pbcPSA_Paint()
    mMPsaPaint tmVpf
End Sub

Private Sub pbcPSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcPSTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to Left
    If imPBoxNo = PRODUCERINDEX Then
        If mProducerBranch() Then
            Exit Sub
        End If
    End If
'    If imPBoxNo = CONTENTPROVIDERINDEX Then
'        If mContentProviderBranch() Then
'            Exit Sub
'        End If
'    End If
    If imPBoxNo = EXPPROGAUDIOINDEX Then
        If mContentProviderBranch() Then
            Exit Sub
        End If
    End If
    If imPBoxNo = EXPCOMMAUDIOINDEX Then
        If mContentProviderBranch() Then
            Exit Sub
        End If
    End If
    mPSetShow imPBoxNo
    ilBox = imPBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                ilBox = PRODUCERINDEX
                imPBoxNo = ilBox
                mPEnableBox ilBox
                Exit Sub
            Case PRODUCERINDEX 'Time (first control within header)
                imPBoxNo = -1
                cmcDone.SetFocus
                Exit Sub
            Case Else
                ilBox = imPBoxNo - 1
                ilFound = True
        End Select
    Loop While Not ilFound
    imPBoxNo = ilBox
    mPEnableBox ilBox

End Sub

Private Sub pbcPTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcPTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = 0
    If imPBoxNo = PRODUCERINDEX Then
        If mProducerBranch() Then
            Exit Sub
        End If
    End If
'    If imPBoxNo = CONTENTPROVIDERINDEX Then
'        If mContentProviderBranch() Then
'            Exit Sub
'        End If
'    End If
    If imPBoxNo = EXPPROGAUDIOINDEX Then
        If mContentProviderBranch() Then
            Exit Sub
        End If
    End If
    If imPBoxNo = EXPCOMMAUDIOINDEX Then
        If mContentProviderBranch() Then
            Exit Sub
        End If
    End If
    mPSetShow imPBoxNo
    ilBox = imPBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                If tgSpf.sGUseAffSys = "Y" Then
                    '1/17/10: Embedded moved to Affiliate Export ISCI
                    ilBox = EXPCOMMAUDIOINDEX   'COMMEMBEDDEDINDEX
                Else
                    ilBox = PRODUCERINDEX
                End If
                imPBoxNo = ilBox
                mPEnableBox ilBox
                Exit Sub
            Case PRODUCERINDEX 'Time (first control within header)
                If tgSpf.sGUseAffSys <> "Y" Then
                    imPBoxNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
'                ilBox = CONTENTPROVIDERINDEX
                ilBox = EXPPROGAUDIOINDEX
                ilFound = True
            '1/17/10: Embedded moved to Affiliate Export ISCI
            'Case COMMEMBEDDEDINDEX
            Case EXPCOMMAUDIOINDEX
                imPBoxNo = -1
                cmcDone.SetFocus
                Exit Sub
            Case Else
                ilBox = imPBoxNo + 1
                ilFound = True
        End Select
    Loop While Not ilFound
    imPBoxNo = ilBox
    mPEnableBox ilBox
End Sub

Private Sub pbcSalesSTab_GotFocus()
    Dim ilLoop As Integer
    If GetFocus() <> pbcSalesSTab.HWnd Then
        Exit Sub
    End If
    If imSSpotLenBoxNo = -1 Then
        edcSSpotLG.Visible = True
        imSSpotLenBoxNo = 0
    Else
        mSetCommands
        If imSSpotLenBoxNo = 0 Then
            edcSSpotLG.Visible = False
            imSSpotLenBoxNo = -1
            For ilLoop = 0 To 6 Step 1
                If rbcSSalesperson(ilLoop).Value Then
                    rbcSSalesperson(ilLoop).SetFocus
                    Exit For
                End If
            Next ilLoop
            Exit Sub
        End If
        imSSpotLenBoxNo = imSSpotLenBoxNo - 1
    End If
    mSetSSpotLen
End Sub
Private Sub pbcSalesTab_GotFocus()
    If GetFocus() <> pbcSalesTab.HWnd Then
        Exit Sub
    End If
    If imSSpotLenBoxNo = -1 Then
        edcSSpotLG.Visible = True
        imSSpotLenBoxNo = 2 * (UBound(tmVpf.iSLen) + 1) - 1
    Else
        mSetCommands
        If imSSpotLenBoxNo = 2 * (UBound(tmVpf.iSLen) + 1) - 1 Then
            edcSSpotLG.Visible = False
            imSSpotLenBoxNo = -1
            edcSLen.SetFocus
            Exit Sub
        End If
        imSSpotLenBoxNo = imSSpotLenBoxNo + 1
    End If
    mSetSSpotLen
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
        For ilBox = LBONE To UBound(lmSSave) Step 1
            If ilRow = 1 Then
                If ilBox = LBONE Then
                    slStr = ".01"
                Else
                    slStr = Trim$(Str$(lmSSave(ilBox - 1) + 1))
                End If
            Else
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

Private Sub pbcSSpotLen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilIndex As Integer
    mSetCommands
    ilIndex = X \ (edcSSpotLG.Width + 15)
    If Y < edcSSpotLG.Height Then
        imSSpotLenBoxNo = ilIndex
    Else
        imSSpotLenBoxNo = ilIndex + UBound(tmVpf.iSLen) + 1
    End If
    mSetSSpotLen
End Sub
Private Sub pbcSSpotLen_Paint()
    mSSpotLenPaint tmVpf
End Sub
Private Sub pbcVehicle_DragDrop(Source As Control, X As Single, Y As Single)
    If imDragDest = -1 Then
        mVirtClearDrag
        Exit Sub
    End If
    Select Case imDragSrce
        Case DRAGVEHNAME
            cmcMoveToVehicle_Click
        Case DRAGVEHICLE
    End Select
    mVirtClearDrag
End Sub
Private Sub pbcVehicle_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    imDragIndexDest = 0
    If imDragSrce = DRAGVEHNAME Then
        If State = vbLeave Then
            lbcVehNames.DragIcon = IconTraf!imcIconDrag.DragIcon
            Exit Sub
        End If
        imDragDest = DRAGVEHICLE
        lbcVehNames.DragIcon = IconTraf!imcIconInsert.DragIcon
    End If
End Sub
Private Sub pbcVehicle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Exit Sub
    End If
    If imUpdateAllowed Then
        fmDragX = X
        fmDragY = Y
        imDragButton = Button
        imDragType = 0
        imDragShift = Shift
        imDragSrce = DRAGVEHICLE
        imDragIndexDest = 0
        tmcDrag.Enabled = True  'Start timer to see if drag or click
    End If
End Sub
Private Sub pbcVehicle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    ilCompRow = vbcVehicle.LargeChange + 1
    If UBound(smVirtSave, 2) - 1 > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smVirtSave, 2) - 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBVirtCtrls To UBound(tmVirtCtrls) Step 1
            If (X >= tmVirtCtrls(ilBox).fBoxX) And (X <= (tmVirtCtrls(ilBox).fBoxX + tmVirtCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmVirtCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmVirtCtrls(ilBox).fBoxY + tmVirtCtrls(ilBox).fBoxH)) Then
                    If ilBox = VEHINDEX Then
                        Beep
                        mVirtSetFocus imVirtBoxNo
                        If imVirtBoxNo = -1 Then
                            imVirtRowNo = ilRow + vbcVehicle.Value - 1
                            lacVehFrame.Move 0, tmVirtCtrls(VEHINDEX).fBoxY + (imVirtRowNo - vbcVehicle.Value) * (fgBoxGridH + 15) - 30
                            lacVehFrame.Visible = True
                        End If
                        Exit Sub
                    End If
                    If ilRow + vbcVehicle.Value - 1 <= imOrigNoVehicles Then
                        Beep
                        mVirtSetFocus imVirtBoxNo
                        Exit Sub
                    End If
                    ilRowNo = ilRow + vbcVehicle.Value - 1
                    mVirtSetShow imVirtBoxNo
                    imVirtRowNo = ilRowNo
                    imVirtBoxNo = ilBox
                    mVirtEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mVirtSetFocus imVirtBoxNo
End Sub
Private Sub pbcVehicle_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long
    ilStartRow = vbcVehicle.Value  'Top location
    ilEndRow = vbcVehicle.Value + vbcVehicle.LargeChange
    If ilEndRow > UBound(smVirtSave, 2) Then
        ilEndRow = UBound(smVirtSave, 2) 'include blank row as it might have data
    End If
    llColor = pbcVehicle.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = LBONE To UBound(smVirtShow, 1) Step 1
            slStr = smVirtShow(ilBox, ilRow)
            'gPaintArea pbcVehicle, tmVirtCtrls(ilBox).fBoxX, tmVirtCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmVirtCtrls(ilBox).fBoxW - 15, tmVirtCtrls(ilBox).fBoxH - 15, WHITE
            pbcVehicle.CurrentX = tmVirtCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcVehicle.CurrentY = tmVirtCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            pbcVehicle.Print slStr
        Next ilBox
    Next ilRow
    pbcVehicle.ForeColor = llColor
End Sub
Private Sub pbcVirtSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcVirtSTab.HWnd Then
        Exit Sub
    End If
    mVirtSetShow imVirtBoxNo
    ilBox = imVirtBoxNo
    If (imVirtBoxNo > imLBVirtCtrls) And (imVirtBoxNo <= UBound(tmVirtCtrls)) Then
        If mVirtTestFields(imVirtBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mVirtEnableBox imVirtBoxNo
            Exit Sub
        End If
    End If
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                If LBONE = UBound(smVirtSave, 2) Then
                    lbcVehNames.SetFocus
                    Exit Sub
                End If
                If imOrigNoVehicles > 0 Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imVirtSettingValue = True
                vbcVehicle.Value = vbcVehicle.Min
                imVirtSettingValue = False
                If UBound(smVirtShow, 2) - 1 <= vbcVehicle.LargeChange Then
                    vbcVehicle.Max = LBONE  'LBound(smVirtShow, 2)
                Else
                    vbcVehicle.Max = UBound(smVirtShow, 2) - vbcVehicle.LargeChange '- 1
                End If
                imVirtRowNo = 1
                imVirtSettingValue = False
                ilBox = NOSPOTSINDEX
                imVirtBoxNo = ilBox
                mVirtEnableBox ilBox
                Exit Sub
            Case VEHINDEX 'Time (first control within header)
                ilBox = PERCENTINDEX
                If (imVirtRowNo <= 1) Or (imVirtRowNo <= imOrigNoVehicles) Then
                    imVirtBoxNo = -1
                    imVirtRowNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imVirtRowNo = imVirtRowNo - 1
                If imVirtRowNo < vbcVehicle.Value Then
                    imVirtSettingValue = True
                    vbcVehicle.Value = vbcVehicle.Value - 1
                    imVirtSettingValue = False
                End If
                ilFound = True
            Case PERCENTINDEX
                ilBox = NOSPOTSINDEX
                ilFound = True
            Case NOSPOTSINDEX
                ilBox = VEHINDEX
                ilFound = False
        End Select
    Loop While Not ilFound
    imVirtBoxNo = ilBox
    mVirtEnableBox ilBox
End Sub
Private Sub pbcVirtTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcVirtTab.HWnd Then
        Exit Sub
    End If
    ilBox = imVirtBoxNo
    mVirtSetShow imVirtBoxNo
    If (imVirtBoxNo > imLBVirtCtrls) And (imVirtBoxNo <= UBound(tmVirtCtrls)) Then
        If mVirtTestFields(imVirtBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mVirtEnableBox imVirtBoxNo
            Exit Sub
        End If
    End If
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                If imOrigNoVehicles > 0 Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imVirtRowNo = UBound(smVirtSave, 2) - 1
                imVirtSettingValue = True
                If imVirtRowNo - 1 <= vbcVehicle.LargeChange Then
                    vbcVehicle.Value = vbcVehicle.Min
                Else
                    vbcVehicle.Value = imVirtRowNo - vbcVehicle.LargeChange '- 1
                End If
                imVirtSettingValue = False
                ilBox = PERCENTINDEX
                ilFound = True
            Case PERCENTINDEX
                imVirtRowNo = imVirtRowNo + 1
                If imVirtRowNo >= UBound(smVirtSave, 2) Then
                    imVirtBoxNo = -1
                    imVirtRowNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                If imVirtRowNo > vbcVehicle.Value + vbcVehicle.LargeChange Then
                    imVirtSettingValue = True
                    vbcVehicle.Value = vbcVehicle.Value + 1
                    imVirtSettingValue = False
                End If
                ilBox = VEHINDEX
                ilFound = False
            Case VEHINDEX
                ilBox = NOSPOTSINDEX
                ilFound = True
            Case NOSPOTSINDEX
                ilBox = PERCENTINDEX
                ilFound = True
        End Select
    Loop While Not ilFound
    imVirtBoxNo = ilBox
    mVirtEnableBox ilBox
End Sub

Private Sub pbcXDXMLForm_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcXDXMLForm_KeyPress(KeyAscii As Integer)
    If ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
        If (KeyAscii = Asc("H")) Or (KeyAscii = Asc("h")) Then
            'Each spot must be sent separately since the UnitID has the Affiliate Spot ID (Code)
            '9114
            If ((Asc(tgSpf.sUsingFeatures10) And UNITIDBYASTCODEFORBREAK) = UNITIDBYASTCODEFORBREAK) Then
                If (smXDXMLForm <> "H#B#P#") Then
                    imVffChg = True
                    smXDXMLForm = "H#B#P#"
                End If
            Else
                If smXDXMLForm = "H#B#" Then
                    imVffChg = True
                    smXDXMLForm = "H#B#P#"
                ElseIf smXDXMLForm = "H#B#P#" Then
                    imVffChg = True
                    smXDXMLForm = "H#B#"
                Else
                    imVffChg = True
                    smXDXMLForm = "H#B#"
                End If
            End If
            pbcXDXMLForm_Paint
        ElseIf (KeyAscii = Asc("N")) Or (KeyAscii = Asc("n")) Then
            If smXDXMLForm <> "" Then
                imVffChg = True
                smXDXMLForm = ""
            End If
            pbcXDXMLForm_Paint
        End If
    End If
    'If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) Then
    '    If (KeyAscii = Asc("I")) Or (KeyAscii = Asc("i")) Then
    '        If (smXDXMLForm <> "ISCI") Then
    '            imVffChg = True
    '            smXDXMLForm = "ISCI"
    '        End If
    '        pbcXDXMLForm_Paint
    '    End If
    'End If
    If KeyAscii = Asc(" ") Then
        'If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) And ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
        '    If smXDXMLForm = "H#B#" Then
        '        imVffChg = True
        '        smXDXMLForm = "H#B#P#"
        '    ElseIf smXDXMLForm = "H#B#P#" Then
        '        imVffChg = True
        '    '    smXDXMLForm = "ISCI"
        '    'ElseIf smXDXMLForm = "ISCI" Then
        '        imVffChg = True
        '        smXDXMLForm = "H#B#"
        '    Else
        '        imVffChg = True
        '        smXDXMLForm = "H#B#"
        '    End If
        '    'Each spot must be sent separately since the UnitID has the Affiliate Spot ID (Code)
        '    If (Mid(smXDXMLForm, 1, 1) = "H") And ((Asc(tgSpf.sUsingFeatures10) And UNITIDBYASTCODE) = UNITIDBYASTCODE) Then
        '        smXDXMLForm = "H#B#P#"
        '    End If
        'Else
            If ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
                'Each spot must be sent separately since the UnitID has the Affiliate Spot ID (Code)
            ' 9113 remove if 9114
            ' this if blocks break to only be hbp if 'astcode'.  Comment out the top part, so we always do the 'else' to turn on 9113
'            If ((Asc(tgSpf.sUsingFeatures10) And UNITIDBYASTCODEFORBREAK) = UNITIDBYASTCODEFORBREAK) Then
'                    If (smXDXMLForm = "") Then
'                        imVffChg = True
'                        smXDXMLForm = "H#B#P#"
'                    ElseIf (smXDXMLForm = "H#B#P#") Then
'                        imVffChg = True
'                        smXDXMLForm = ""
'                    Else
'                        imVffChg = True
'                        smXDXMLForm = ""
'                    End If
'                Else
                    If smXDXMLForm = "H#B#" Then
                        imVffChg = True
                        smXDXMLForm = "H#B#P#"
                    ElseIf smXDXMLForm = "H#B#P#" Then
                        imVffChg = True
                        smXDXMLForm = ""
                    ElseIf smXDXMLForm = "" Then
                        imVffChg = True
                        smXDXMLForm = "H#B#"
                   Else
                        imVffChg = True
                        smXDXMLForm = ""
                    End If
              '  End If
            'Else
            '    imVffChg = True
            '    smXDXMLForm = "ISCI"
            '9113 comment out when ready
            'End If
        End If
        pbcXDXMLForm_Paint
    End If
    mSetCommands
End Sub

Private Sub pbcXDXMLForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) And ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
    '    If smXDXMLForm = "H#B#" Then
    '        imVffChg = True
    '        smXDXMLForm = "H#B#P#"
    '    ElseIf smXDXMLForm = "H#B#P#" Then
    '        imVffChg = True
    '        smXDXMLForm = "ISCI"
    '    ElseIf smXDXMLForm = "ISCI" Then
    '        imVffChg = True
    '        smXDXMLForm = "H#B#"
    '    Else
    '        imVffChg = True
    '        smXDXMLForm = "H#B#"
    '    End If
    '    'Each spot must be sent separately since the UnitID has the Affiliate Spot ID (Code)
    '    If (Mid(smXDXMLForm, 1, 1) = "H") And ((Asc(tgSpf.sUsingFeatures10) And UNITIDBYASTCODE) = UNITIDBYASTCODE) Then
    '        smXDXMLForm = "H#B#P#"
    '    End If
    'Else
        If ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
            ' 9113 remove if 9114
            ' this if blocks break to only be hbp if 'astcode'.  Comment out the top part, so we always do the 'else' to turn on 9113
'            If ((Asc(tgSpf.sUsingFeatures10) And UNITIDBYASTCODEFORBREAK) = UNITIDBYASTCODEFORBREAK) Then
'                'Each spot must be sent separately since the UnitID has the Affiliate Spot ID (Code)
'                If (smXDXMLForm = "") Then
'                    imVffChg = True
'                    smXDXMLForm = "H#B#P#"
'                ElseIf (smXDXMLForm = "H#B#P#") Then
'                    imVffChg = True
'                    smXDXMLForm = ""
'                Else
'                    imVffChg = True
'                    smXDXMLForm = ""
'                End If
'            Else
                If smXDXMLForm = "H#B#" Then
                    imVffChg = True
                    smXDXMLForm = "H#B#P#"
                ElseIf smXDXMLForm = "H#B#P#" Then
                    imVffChg = True
                    smXDXMLForm = ""
                ElseIf smXDXMLForm = "" Then
                    imVffChg = True
                    smXDXMLForm = "H#B#"
                Else
                    imVffChg = True
                    smXDXMLForm = ""
                End If
            '9113 comment out to enable
          '  End If
    '    Else
    '        imVffChg = True
    '        smXDXMLForm = "ISCI"
        End If
    'End If
    pbcXDXMLForm_Paint
    mSetCommands
End Sub

Private Sub pbcXDXMLForm_Paint()
    pbcXDXMLForm.Cls
    pbcXDXMLForm.CurrentX = fgBoxInsetX
    pbcXDXMLForm.CurrentY = 0 'fgBoxInsetY
    'If smXDXMLForm = "ISCI" Then
    '    lacCode(3).Caption = smISCIAvailForm
    '    pbcXDXMLForm.Print smXDXMLForm
    'ElseIf smXDXMLForm = "H#B#P#" Then
    If smXDXMLForm = "H#B#P#" Then
        'lacCode(3).Caption = smHBHBPAvailForm
        pbcXDXMLForm.Print smXDXMLForm
    ElseIf smXDXMLForm = "H#B#" Then
        'lacCode(3).Caption = smHBHBPAvailForm
        pbcXDXMLForm.Print smXDXMLForm
    ElseIf smXDXMLForm = "" Then
        edcXDISCIPrefix(0).Text = ""
        edcInterfaceID(0).Text = ""
    End If
End Sub

Private Sub plcAccounting_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcBarter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mRemoveFocus
End Sub

Private Sub plcGeneral_Click()
    'edcLinkDropDown_LostFocus
    'lbcLink.Visible = False
    'cmcLinkDropDown.Visible = False
    'edcLinkDropDown.Visible = False
    'imLinkBoxNo = -1
    'edcGTZDropDown_LostFocus
    'pbcGTZToggle.Visible = False
    'lbcFeed.Visible = False
    'cmcGTZDropDown.Visible = False
    'edcGTZDropDown.Visible = False
    'imGTZBoxNo = -1
    pbcClickFocus.SetFocus  'The above code is in mRemoveFocus which is called be pbcClickFocus
End Sub

 Private Sub plcInvBy_Paint()
    plcInvBy.Cls
    plcInvBy.CurrentX = 0
    plcInvBy.CurrentY = 0
    plcInvBy.Print "By"
End Sub

Private Sub plcLog_Click()
    pbcClickFocus.SetFocus
End Sub


Private Sub plcParticipant_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcProducer_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcPSAPromo_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSales_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcShowRateOnInsertion_Paint()
    plcShowRateOnInsertion.CurrentX = 0
    plcShowRateOnInsertion.CurrentY = 0
    plcShowRateOnInsertion.Print "Show Spot Rate On Insertion Orders"
End Sub

Private Sub plcSplitCopy_Paint()
    Dim ilValue As Integer

    plcSplitCopy.CurrentX = 0
    plcSplitCopy.CurrentY = 0
    ilValue = Asc(tgSpf.sUsingFeatures2)
    If (((ilValue And SPLITCOPY) = SPLITCOPY) And ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
        plcSplitCopy.Print "Allow Split Network/Copy"
    ElseIf ((ilValue And SPLITCOPY) = SPLITCOPY) Then
        plcSplitCopy.Print "Allow Split Copy"
    Else
        plcSplitCopy.Print "Allow Split Network"
    End If
End Sub

Private Sub rbcAudioType_Click(Index As Integer)
    Dim Value As Integer
    Value = rbcAudioType(Index).Value
    If Value Then
        imVffChg = True
        mSetCommands
    End If
End Sub

Private Sub rbcAudioType_GotFocus(Index As Integer)
    gCtrlGotFocus rbcAudioType(Index)
End Sub

Private Sub rbcBarterMethod_Click(Index As Integer)
    If rbcBarterMethod(Index).Value Then
        'lacBarter(8).Enabled = False
        'edcBarter(5).Enabled = False
        If Index = 0 Then
            edcBarterMethod(0).Text = ""
            edcBarterMethod(1).Text = ""
            edcBarterMethod(2).Text = ""
            edcBarterMethod(3).Text = ""
            edcBarterMethod(0).Enabled = False
            edcBarterMethod(1).Enabled = False
            edcBarterMethod(2).Enabled = False
            edcBarterMethod(3).Enabled = False
            cbcPerPeriod(0).Enabled = False
            cbcPerPeriod(1).Enabled = False
            cbcPerPeriod(2).Enabled = False
            cbcPerPeriod(0).ListIndex = -1
            cbcPerPeriod(1).ListIndex = -1
            cbcPerPeriod(2).ListIndex = -1
        ElseIf Index = 1 Then
            edcBarterMethod(1).Text = ""
            edcBarterMethod(2).Text = ""
            edcBarterMethod(3).Text = ""
            edcBarterMethod(0).Enabled = True
            edcBarterMethod(1).Enabled = False
            edcBarterMethod(2).Enabled = False
            edcBarterMethod(3).Enabled = False
            cbcPerPeriod(0).Enabled = True
            cbcPerPeriod(1).Enabled = False
            cbcPerPeriod(2).Enabled = False
            cbcPerPeriod(1).ListIndex = -1
            cbcPerPeriod(2).ListIndex = -1
        ElseIf Index = 2 Then
            edcBarterMethod(0).Text = ""
            edcBarterMethod(2).Text = ""
            edcBarterMethod(3).Text = ""
            edcBarterMethod(0).Enabled = False
            edcBarterMethod(1).Enabled = True
            edcBarterMethod(2).Enabled = False
            edcBarterMethod(3).Enabled = False
            cbcPerPeriod(0).Enabled = False
            cbcPerPeriod(1).Enabled = True
            cbcPerPeriod(2).Enabled = False
            cbcPerPeriod(0).ListIndex = -1
            cbcPerPeriod(2).ListIndex = -1
        ElseIf Index = 3 Then
            edcBarterMethod(0).Text = ""
            edcBarterMethod(1).Text = ""
            edcBarterMethod(0).Enabled = False
            edcBarterMethod(1).Enabled = False
            edcBarterMethod(2).Enabled = True
            edcBarterMethod(3).Enabled = True
            cbcPerPeriod(0).Enabled = False
            cbcPerPeriod(1).Enabled = False
            cbcPerPeriod(2).Enabled = True
            cbcPerPeriod(0).ListIndex = -1
            cbcPerPeriod(1).ListIndex = -1
        ElseIf Index = 4 Then
            edcBarterMethod(0).Text = ""
            edcBarterMethod(1).Text = ""
            edcBarterMethod(2).Text = ""
            edcBarterMethod(3).Text = ""
            edcBarterMethod(0).Enabled = False
            edcBarterMethod(1).Enabled = False
            edcBarterMethod(2).Enabled = False
            edcBarterMethod(3).Enabled = False
            cbcPerPeriod(0).Enabled = False
            cbcPerPeriod(1).Enabled = False
            cbcPerPeriod(2).Enabled = False
            cbcPerPeriod(0).ListIndex = -1
            cbcPerPeriod(1).ListIndex = -1
            cbcPerPeriod(2).ListIndex = -1
        '8032
'        ElseIf Index = 5 Then
'            imVffChg = True
'            mSetCommands
'        ElseIf Index = 6 Then
'            imVffChg = True
'            mSetCommands
        'dan no longer needed
'        ElseIf Index = STATIONXMLWIDEORBIT Or Index = STATIONXMLNONE Or Index = STATIONXMLMARKETRON Then
'            imVffChg = True
'            mSetCommands
        ElseIf Index = 5 Then
            lacBarter(8).Enabled = False
            edcBarter(5).Enabled = False
            edcBarter(5).Text = ""
        ElseIf Index = 6 Then
            lacBarter(8).Enabled = True
            edcBarter(5).Enabled = True
        End If
    End If
    imVffChg = True
    mSetCommands
End Sub

Private Sub rbcBarterMethod_GotFocus(Index As Integer)
    imAcqCostBoxNo = -1
    imAcqIndexBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub rbcBillSA_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcBillSA(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcBillSA_GotFocus(Index As Integer)
    gCtrlGotFocus rbcBillSA(Index)
End Sub
Private Sub rbcExpBkCpyCart_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcExpBkCpyCart(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcExpBkCpyCart_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub rbcMerge_Click(Index As Integer)
    imVffChg = True
    If Index = 0 Then   'Traffic Merge
        If rbcMerge(Index).Value Then
            rbcMerge(2).Enabled = True
        End If
    End If
    If Index = 1 Then   'Traffic Separate
        If rbcMerge(Index).Value Then
            rbcMerge(3).Value = True
            rbcMerge(2).Enabled = False
        End If
    End If
    If Index = 2 Then   'Affiliate Merge
        If rbcMerge(Index).Value Then
            rbcMerge(4).Value = True
            rbcMerge(5).Enabled = False
        End If
    End If
    If Index = 3 Then   'Affiliate Separate
        If rbcMerge(Index).Value Then
            rbcMerge(5).Enabled = True
        End If
    End If
    mSetCommands
End Sub

Private Sub rbcMerge_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub rbcRemoteExport_Click(Index As Integer)
    Dim ilPasswordOk As Integer

    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If (rbcRemoteExport(Index).Value) And (Index <> 2) Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "RE-"
                CSPWord.Show vbModal
                If Not igPasswordOk Then
                    rbcRemoteExport(2).Value = True
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            rbcRemoteExport(2).Value = True
        End If
    End If
    If rbcRemoteExport(Index).Value Then
        mSetCommands
    End If
End Sub

Private Sub rbcRemoteExport_GotFocus(Index As Integer)
    edcGTZDropDown_LostFocus
    pbcGTZToggle.Visible = False
    lbcFeed.Visible = False
    cmcGTZDropDown.Visible = False
    edcGTZDropDown.Visible = False
    imGTZBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub rbcGenLog_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcGenLog(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcGenLog_GotFocus(Index As Integer)
    gCtrlGotFocus rbcGenLog(Index)
End Sub
Private Sub rbcGMedium_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    
    '10050 podcast
    mVendorEnableOptions
    Value = rbcGMedium(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcGMedium_GotFocus(Index As Integer)
    edcGTZDropDown_LostFocus
    pbcGTZToggle.Visible = False
    lbcFeed.Visible = False
    cmcGTZDropDown.Visible = False
    edcGTZDropDown.Visible = False
    imGTZBoxNo = -1
    gCtrlGotFocus rbcGMedium(Index)
End Sub
Private Sub rbcIFBulk_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcIFBulk(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcIFBulk_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub rbcIFSelling_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcIFSelling(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcIFSelling_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub rbcIFTime_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcIFTime(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcIFTime_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub rbcInvBy_Click(Index As Integer)
    If Index = 1 Then               'network inventory by year, must be rollover within year
        ckcRollover.Value = vbChecked
        ckcRollover.Enabled = False
    Else
        ckcRollover.Enabled = True
    End If
End Sub

Private Sub rbcLCopyOnAir_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcLCopyOnAir(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcLCopyOnAir_GotFocus(Index As Integer)
    mLSetShow imLBoxNo
    imLBoxNo = -1
    imLRowNo = -1
    gCtrlGotFocus rbcLCopyOnAir(Index)
End Sub
Private Sub rbcLDaylight_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcLDaylight(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcLDaylight_GotFocus(Index As Integer)
    gCtrlGotFocus rbcLDaylight(Index)
End Sub
Private Sub rbcLGrid_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcLGrid(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcLGrid_GotFocus(Index As Integer)
    mLSetShow imLBoxNo
    imLBoxNo = -1
    imLRowNo = -1
    gCtrlGotFocus rbcLGrid(Index)
End Sub
Private Sub rbcLTiming_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcLTiming(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcLTiming_GotFocus(Index As Integer)
    gCtrlGotFocus rbcLTiming(Index)
End Sub
Private Sub rbcLZone_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcLZone(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcLZone_GotFocus(Index As Integer)
    gCtrlGotFocus rbcLZone(Index)
End Sub

Private Sub rbcRemoteImport_Click(Index As Integer)
    Dim ilPasswordOk As Integer
    If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
        If igPasswordOk Then
            If (rbcRemoteImport(Index).Value) And (Index <> 2) Then
                ilPasswordOk = igPasswordOk
                sgPasswordAddition = "RI-"
                CSPWord.Show vbModal
                If Not igPasswordOk Then
                    rbcRemoteImport(2).Value = True
                End If
                sgPasswordAddition = ""
                igPasswordOk = ilPasswordOk
            End If
        Else
            rbcRemoteImport(2).Value = True
        End If
    End If
    If rbcRemoteImport(Index).Value Then
        mSetCommands
    End If
End Sub

Private Sub rbcRemoteImport_GotFocus(Index As Integer)
    edcGTZDropDown_LostFocus
    pbcGTZToggle.Visible = False
    lbcFeed.Visible = False
    cmcGTZDropDown.Visible = False
    edcGTZDropDown.Visible = False
    imGTZBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub rbcSAdvtSep_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSAdvtSep(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcSAdvtSep_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub rbcSBreak_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSBreak(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcSBreak_GotFocus(Index As Integer)
    imSSpotLenBoxNo = -1
    gCtrlGotFocus rbcSBreak(Index)
End Sub
Private Sub rbcSCommission_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSCommission(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcSCommission_GotFocus(Index As Integer)
    imSSpotLenBoxNo = -1
    gCtrlGotFocus rbcSCommission(Index)
End Sub
Private Sub rbcSCompetitive_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSCompetitive(Index).Value
    'End of coded added
    If Value Then
        If Index = 0 Then   'Time
            edcSCompSepLen.Enabled = True
        Else
            edcSCompSepLen.Enabled = False
            edcSCompSepLen.Text = ""
        End If
        mSetCommands
    End If
End Sub
Private Sub rbcSCompetitive_GotFocus(Index As Integer)
    imSSpotLenBoxNo = -1
    gCtrlGotFocus rbcSCompetitive(Index)
End Sub

Private Sub rbcShowAirDate_Click(Index As Integer)
    Dim Value As Integer
    Value = rbcShowAirDate(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub

Private Sub rbcShowAirDate_GotFocus(Index As Integer)
    edcGTZDropDown_LostFocus
    pbcGTZToggle.Visible = False
    lbcFeed.Visible = False
    cmcGTZDropDown.Visible = False
    edcGTZDropDown.Visible = False
    imGTZBoxNo = -1
    gCtrlGotFocus rbcShowAirDate(Index)
End Sub

Private Sub rbcShowAirTime_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcShowAirTime(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcShowAirTime_GotFocus(Index As Integer)
    edcGTZDropDown_LostFocus
    pbcGTZToggle.Visible = False
    lbcFeed.Visible = False
    cmcGTZDropDown.Visible = False
    edcGTZDropDown.Visible = False
    imGTZBoxNo = -1
    gCtrlGotFocus rbcShowAirTime(Index)
End Sub

Private Sub rbcShowRateOnInsertion_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcShowRateOnInsertion(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub

Private Sub rbcShowRateOnInsertion_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub rbcSMove_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSMove(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcSMove_GotFocus(Index As Integer)
    imSSpotLenBoxNo = -1
    gCtrlGotFocus rbcSMove(Index)
End Sub
Private Sub rbcSMoveLLD_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSMoveLLD(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcSMoveLLD_GotFocus(Index As Integer)
    gCtrlGotFocus rbcSMoveLLD(Index)
End Sub
Private Sub rbcSOverbook_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSOverbook(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcSOverbook_GotFocus(Index As Integer)
    imSSpotLenBoxNo = -1
    gCtrlGotFocus rbcSOverbook(Index)
End Sub

Private Sub rbcSplitCopy_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Dim ilValue As Integer
    Value = rbcSplitCopy(Index).Value
    'End of coded added
    If Value Then
        ilValue = Asc(tgSpf.sUsingFeatures2)
        'If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or ((tmVef.sType = "A") And (rbcLCopyOnAir(0).Value)) Or (tmVef.sType = "G") Or (tmVef.sType = "P")) And ((ilValue And SPLITCOPY) = SPLITCOPY) Then
        'Else
        '5/11/11: Allow Selling vehicles to be defined as No for split copy
        'If ((tmVef.sType = "C") Or (tmVef.sType = "A") Or (tmVef.sType = "G")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
        If ((tmVef.sType = "C") Or (tmVef.sType = "A") Or (tmVef.sType = "G") Or (tmVef.sType = "S")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
        'ElseIf ((tmVef.sType = "S") Or (tmVef.sType = "P")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
        ElseIf ((tmVef.sType = "P")) And (((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And SPLITNETWORKS) = SPLITNETWORKS)) Then
            rbcSplitCopy(0).Value = True
        Else
            rbcSplitCopy(Index).Value = False
        End If
        mSetCommands
    End If
End Sub

Private Sub rbcSplitCopy_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub rbcSSalesperson_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSSalesperson(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcSSalesperson_GotFocus(Index As Integer)
    imSSpotLenBoxNo = -1
    gCtrlGotFocus rbcSSalesperson(Index)
End Sub
Private Sub rbcSSellout_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSSellout(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcSSellout_GotFocus(Index As Integer)
    imSSpotLenBoxNo = -1
    gCtrlGotFocus rbcSSellout(Index)
End Sub

Private Sub tbcSelection_Click()
    Dim ilIndex As Integer
    Dim ilValue As Integer

    ilIndex = tbcSelection.SelectedItem.Index - 1
    plcGeneral.Visible = False
    plcSales.Visible = False
    plcSchedule(0).Visible = False
    plcSchedule(1).Visible = False
    plcSports.Visible = False
    plcPSAPromo.Visible = False
    plcLog.Visible = False
    plcAccounting.Visible = False
    plcVirtual.Visible = False
    plcProducer.Visible = False
    plcBarter.Visible = False
    cbcBarter.Visible = False
    frcBarterEnable(0).Visible = False
    frcBarterEnable(1).Visible = False
    frcBarterEnable(2).Visible = False
    lacBarter(4).Visible = False
    rbcBarterMethod(5).Visible = False
    rbcBarterMethod(6).Visible = False
    plcParticipant.Visible = False
    plcGreatPlains.Visible = False
    plcExport.Visible = False
    frcBarter(1).Visible = False
    If udcVehOptTabs.Visible Then
        udcVehOptTabs.Action 5, True
        udcVehOptTabs.Visible = False
    End If
    Select Case ilIndex
        Case 0  'General,...
            plcGeneral.Visible = True
        Case 1  'Sales
            If (tmVef.sType <> "N") Then
                plcSales.Visible = True
            End If
        Case 2  'Schedule
            If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Then
                plcSchedule(0).Visible = True
            End If
            If (tmVef.sType = "A") Then
                plcSchedule(1).Visible = True
            End If
        Case 3  'Sports
            If tmVef.sType = "G" Then
                plcSports.Visible = True
            End If
        Case 4  'PSA/Promo
            'If (tmVef.sType <> "R") And (tmVef.sType <> "N") Then
            If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Then
                plcPSAPromo.Visible = True
            End If
        Case 5  'Log
            '2/13/03- allow selling because of Last Log Date is used to know about adding spots to stf (tracking file)
            If ((tmVef.sType = "C") Or (tmVef.sType = "L") Or (tmVef.sType = "A") Or (tmVef.sType = "S") Or (tmVef.sType = "G")) And (igVpfType <> 1) Then
                plcLog.Visible = True
                imLRowNo = -1
                imLBoxNo = -1
            End If
        Case 6  'Producter
            If (tmVef.sType <> "N") Then
                If igVpfType <> 1 Then
                    plcProducer.Visible = True
                    imPBoxNo = -1
                End If
            End If
        Case 7  'Export
            If ((tmVef.sType = "C") Or (tmVef.sType = "L") Or (tmVef.sType = "S") Or (tmVef.sType = "A") Or (tmVef.sType = "G")) Then
                plcExport.ZOrder 0     ' Dan M 6-16-08
                plcExport.Visible = True
            End If
        Case 8  '6  'Accounting
            If (tmVef.sType <> "R") And (tmVef.sType <> "N") Then
                If igVpfType <> 1 Then
                    plcAccounting.Visible = True
                End If
            End If
        Case 9  '7  'Barter
            ilValue = Asc(tgSpf.sUsingFeatures2)
            ''If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G")) And (ilValue And BARTER) = BARTER Then
            'If ((tmVef.sType = "C") Or (tmVef.sType = "S")) And (ilValue And BARTER) = BARTER Then
            If (tmVef.sType = "R") And ((ilValue And BARTER) = BARTER) Then
                frcBarterEnable(1).Visible = False
                frcBarterEnable(2).Visible = False
                lncBarter(1).Visible = False
                frcBarter(1).Visible = True
                cbcBarter.Move 5670, 720
                frcBarterEnable(0).Move 120, 1440
                lncBarter(0).Visible = True
                lncBarter(0).Visible = True
                cbcBarter.Visible = True
                frcBarterEnable(0).Visible = True
                plcBarter.Visible = True
            ElseIf (tmVef.sType = "R") Then
                frcBarterEnable(1).Visible = False
                frcBarterEnable(2).Visible = False
                lncBarter(1).Visible = False
                frcBarter(1).Visible = True
                cbcBarter.Move 5670, 720
                frcBarterEnable(0).Move 120, 1440
                lncBarter(0).Visible = False
                cbcBarter.Visible = False
                frcBarterEnable(0).Visible = False
                plcBarter.Visible = True
            ElseIf ((tmVef.sType = "C") Or (tmVef.sType = "S")) And ((ilValue And BARTER) = BARTER) Then
                '8132
                If tmVef.sType = "C" Then
                    frcBarter(1).Visible = True
                End If
                If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then 'Acquisition Commissionable
                    frcBarterEnable(0).Visible = False
                    frcBarterEnable(2).Move 120, 645
                    lncBarter(1).X1 = 285
                    lncBarter(1).Y1 = frcBarterEnable(2).Top + frcBarterEnable(2).Height + 120
                    lncBarter(1).Y2 = lncBarter(1).Y1
                    cbcBarter.Move 5670, frcBarterEnable(2).Top + frcBarterEnable(2).Height + 240
                    frcBarterEnable(1).Move 120, cbcBarter.Top + cbcBarter.Height + 360
                    lncBarter(0).Visible = True
                    cbcBarter.Visible = True
                    lacBarter(4).Visible = True
                    rbcBarterMethod(5).Visible = True
                    rbcBarterMethod(6).Visible = True
                    frcBarterEnable(1).Visible = True
                    frcBarterEnable(2).Visible = True
                    lncBarter(1).Visible = True
                    plcBarter.Visible = True
                Else
                    lncBarter(1).Visible = False
                    frcBarterEnable(0).Visible = False
                    frcBarterEnable(1).Visible = False
                    frcBarterEnable(2).Move 120, 645
                    lncBarter(0).Visible = True
                    cbcBarter.Visible = False
                    lacBarter(4).Visible = True
                    rbcBarterMethod(5).Visible = True
                    rbcBarterMethod(6).Visible = True
                    frcBarterEnable(2).Visible = True
                    plcBarter.Visible = True
                End If
            ElseIf (tmVef.sType = "N") And ((ilValue And BARTER) = BARTER) Then
                lncBarter(0).Visible = False
                lncBarter(1).Visible = False
                cbcBarter.Visible = False
                plcBarter.Visible = True
            End If
        Case 10  '8  'Participant
            'If (tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G") Or (tmVef.sType = "T") Or (tmVef.sType = "R") Or (tmVef.sType = "N") Then
            If (igVpfType <> 1) And (tmVef.sType <> "L") Then
                If imCbcParticipantListIndex = -1 Then
                    If UBound(tmPifRec) <= LBound(tmPifRec) Then
                        cbcParticipant.ListIndex = 0
                    End If
                End If
                plcParticipant.Visible = True
            End If
        Case 11 '9  'Great Plains G/L 'allowed packages 5/7/08
            If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "N") Or (tmVef.sType = "R") Or (tmVef.sType = "G") Or (tmVef.sType = "P")) And ((Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS) Then
                plcGreatPlains.ZOrder 0     ' Dan M 6-16-08
                plcGreatPlains.Visible = True
            End If
        Case 12 'Affiliate Log
            '12/24/15: Bypass if Rep or NTR
            If (tmVef.sType <> "R") And (tmVef.sType <> "N") Then
                igUpdateAllowed = imUpdateAllowed
                udcVehOptTabs.Action 2, True
                udcVehOptTabs.Visible = True
                Screen.MousePointer = vbDefault
            End If
    End Select

End Sub

Private Sub tbcSelection_GotFocus()
    mRemoveFocus
End Sub

Private Sub tbcSelection_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mRemoveFocus
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If plcGeneral.Visible = True Then
        If (imGTZBoxNo Mod imTZMaxCtrls) = GFEEDINDEX Then  'Feed
            imLbcArrowSetting = False
            gProcessLbcClick lbcFeed, edcGTZDropDown, imChgMode, imLbcArrowSetting
        End If
    End If
    If plcProducer.Visible = True Then
        Select Case imPBoxNo
            Case PRODUCERINDEX
                imLbcArrowSetting = False
                gProcessLbcClick lbcProducer, edcPDropdown, imChgMode, imLbcArrowSetting
'            Case CONTENTPROVIDERINDEX
'                imLbcArrowSetting = False
'                gProcessLbcClick lbcContentProvider, edcPDropdown, imChgMode, imLbcArrowSetting
            Case EXPPROGAUDIOINDEX
                imLbcArrowSetting = False
                gProcessLbcClick lbcExpProgAudio, edcPDropdown, imChgMode, imLbcArrowSetting
            Case EXPCOMMAUDIOINDEX
                imLbcArrowSetting = False
                gProcessLbcClick lbcExpCommAudio, edcPDropdown, imChgMode, imLbcArrowSetting
        End Select
    End If
    If plcParticipant.Visible = True Then
        Select Case lmPartEnableCol
            Case SSOURCEINDEX
                imLbcArrowSetting = False
                gProcessLbcClick lbcSSource, edcSSDropDown, imSSChgMode, imLbcArrowSetting
            Case PARTINDEX
                imLbcArrowSetting = False
                gProcessLbcClick lbcVehGp, edcVehGpDropDown, imVehGpChgMode, imLbcArrowSetting
        End Select
    End If
End Sub
Private Sub tmcDrag_Timer()
    Dim ilListIndex As Integer
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            If imDragButton <> 1 Then
                Exit Sub
            End If
            Select Case imDragSrce
                Case DRAGVEHNAME
                    'ilListIndex = fmDragY \ fgListHtArial825 + lbcVehName.TopIndex
                    ilListIndex = lbcVehNames.ListIndex
                    If (ilListIndex >= 0) And (ilListIndex <= lbcVehNames.ListCount - 1) Then
                        'lbcVehicle.ListIndex = ilListIndex
                        'Test length- must equal rotation length
                        lbcVehNames.DragIcon = IconTraf!imcIconDrag.DragIcon
                        imDragIndexSrce = ilListIndex
                        'lacL1Frame.DragIcon = IconTraf!imcIconMove.DragIcon
                        lbcVehNames.Drag vbBeginDrag
                    Else
                        lbcVehNames.ListIndex = -1
                    End If
                Case DRAGVEHICLE
                    imDragType = -1
                    tmcDrag.Enabled = False
                    ilCompRow = vbcVehicle.LargeChange + 1
                    If UBound(smVirtSave, 2) - 1 > ilCompRow Then
                        ilMaxRow = ilCompRow
                    Else
                        ilMaxRow = UBound(smVirtSave, 2) - 1
                    End If
                    For ilRow = 1 To ilMaxRow Step 1
                        If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmVirtCtrls(VEHINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmVirtCtrls(VEHINDEX).fBoxY + tmVirtCtrls(VEHINDEX).fBoxH)) Then
                            mVirtSetShow imVirtBoxNo
                            imVirtBoxNo = -1
                            imVirtRowNo = -1
                            imVirtRowNo = ilRow + vbcVehicle.Value - 1
                            lacVehFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                            lacVehFrame.Move 0, tmVirtCtrls(VEHINDEX).fBoxY + (imVirtRowNo - vbcVehicle.Value) * (fgBoxGridH + 15) - 30
                            'If gInvertArea call then remove visible setting
                            lacVehFrame.Visible = True
                            lacVehFrame.Drag vbBeginDrag
                            lacVehFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                            Exit Sub
                        End If
                    Next ilRow
            End Select
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub

Private Sub udcVehOptTabs_SetSave(slType As String)
    If slType = "WebInfo" Then
        mSetCommands
    ElseIf slType = "FTPInfo" Then
        mSetCommands
    ElseIf slType = "LiveWindow" Then
        mSetCommands
    ElseIf slType = "AffLog" Then
        imVffChg = True
        mSetCommands
    End If
End Sub

Private Sub udcVehOptTabs_Validate(Cancel As Boolean)
    '9213 temp remove when no longer needed
    If udcVehOptTabs.AffLog(0) = vbChecked Then
        ckcSch(3).Value = vbUnchecked
    End If
End Sub


Private Sub vbcVehicle_Change()
    If imVirtSettingValue Then
        pbcVehicle.Cls
        pbcVehicle_Paint
        imVirtSettingValue = False
    Else
        mVirtSetShow imVirtBoxNo
        pbcVehicle.Cls
        pbcVehicle_Paint
        mVirtEnableBox imVirtBoxNo
    End If
End Sub
Private Sub plcBillSA_Paint()
    plcBillSA.CurrentX = 0
    plcBillSA.CurrentY = 0
    plcBillSA.Print "Invoice Airing Vehicle"
End Sub
Private Sub plcSAdvtSep_Paint()
    plcSAdvtSep.CurrentX = 0
    plcSAdvtSep.CurrentY = 0
    plcSAdvtSep.Print "Separate Advertisers by"
End Sub
Private Sub plcSMoveLLD_Paint()
    plcSMoveLLD.CurrentX = 0
    plcSMoveLLD.CurrentY = 0
    plcSMoveLLD.Print "Allow Spot Moves Between Todays Date and Last Log Date"
End Sub
Private Sub plcSSalesperson_Paint()
    plcSSalesperson.CurrentX = 0
    plcSSalesperson.CurrentY = 0
    plcSSalesperson.Print "Method of Calculating Salesperson Commission"
End Sub
Private Sub plcSMove_Paint()
    plcSMove.CurrentX = 0
    plcSMove.CurrentY = 0
    plcSMove.Print "On Spot Screen, Moves that Violate Contract are"
End Sub
Private Sub plcSOverbook_Paint()
    plcSOverbook.CurrentX = 0
    plcSOverbook.CurrentY = 0
    plcSOverbook.Print "Allow Overbooking of Avails"
End Sub
Private Sub plcSSellout_Paint()
    plcSSellout.CurrentX = 0
    plcSSellout.CurrentY = 0
    plcSSellout.Print "Sellout defined by"
End Sub
Private Sub plcSCompetitive_Paint()
    plcSCompetitive.CurrentX = 0
    plcSCompetitive.CurrentY = 0
    plcSCompetitive.Print "Separate Product Protection by"
End Sub
Private Sub plcSCommission_Paint()
    plcSCommission.CurrentX = 0
    plcSCommission.CurrentY = 0
    plcSCommission.Print "Variable Agency Commission"
End Sub
Private Sub plcShowAirTime_Paint(Index As Integer)
    plcShowAirTime(Index).CurrentX = 0
    plcShowAirTime(Index).CurrentY = 0
    If Index = 0 Then
        plcShowAirTime(Index).Print "On Invoices Show Air Time as"
    Else
        plcShowAirTime(Index).Print "On Invoices Show Air Date as"
    End If
End Sub
Private Sub plcGSAGroupNo_Paint()
    plcGSAGroupNo.CurrentX = 0
    plcGSAGroupNo.CurrentY = 0
    plcGSAGroupNo.Print "Selling/Airing Group #"
End Sub
Private Sub plcGMedium_Paint()
    plcGMedium.CurrentX = 0
    plcGMedium.CurrentY = -15
    plcGMedium.Print "Medium"
End Sub
Private Sub plcGenLog_Paint()
    plcGenLog.CurrentX = 0
    plcGenLog.CurrentY = 0
    plcGenLog.Print "Generating Logs for Vehicle"
End Sub
Private Sub plcLCopyOnAir_Paint()
    plcLCopyOnAir.CurrentX = 0
    plcLCopyOnAir.CurrentY = 0
    plcLCopyOnAir.Print "Allow Copy on Airing Vehicle"
End Sub
Private Sub plcLAffTimes_Paint()
    plcLAffTimes.CurrentX = 0
    plcLAffTimes.CurrentY = 0
    plcLAffTimes.Print "File Report Name:    Log                          C.P.                         Other"
End Sub
Private Sub plcLGrid_Paint()
    plcLGrid.CurrentX = 0
    plcLGrid.CurrentY = 0
    plcLGrid.Print "On Program Screen, default Hour Grid Resolution to"
End Sub
Private Sub plcLAffCPs_Paint()
    plcLAffCPs.CurrentX = 0
    plcLAffCPs.CurrentY = 0
    plcLAffCPs.Print "Print Report Name:  Log                          C.P.                         Other"
End Sub
Private Sub plcLTiming_Paint()
    plcLTiming.CurrentX = 0
    plcLTiming.CurrentY = 0
    plcLTiming.Print "Using Log Timing"
End Sub
Private Sub plcLDaylight_Paint()
    plcLDaylight.CurrentX = 0
    plcLDaylight.CurrentY = 0
    plcLDaylight.Print "On Daylight Savings"
End Sub
Private Sub plcLZone_Paint()
    plcLZone.CurrentX = 0
    plcLZone.CurrentY = 0
    plcLZone.Print "Time Zone"
End Sub
Private Sub plcExpBkCpyCart_Paint()
    plcExpBkCpyCart.CurrentX = 0
    plcExpBkCpyCart.CurrentY = 0
    plcExpBkCpyCart.Print "Show Cart # with Bulk Feed Export"
End Sub
Private Sub plcIFTime_Paint()
    plcIFTime.CurrentX = 0
    plcIFTime.CurrentY = 0
    plcIFTime.Print "In Export of Clearance Spots Show All Events after 11:59:55PM as 12:00:01AM"
End Sub
Private Sub plcIFSelling_Paint()
    plcIFSelling.CurrentX = 0
    plcIFSelling.CurrentY = 0
    plcIFSelling.Print "In Export of Clearance Spots Treat Vehicle as if Selling Vehicle"
End Sub
Private Sub plcIFBulk_Paint()
    plcIFBulk.CurrentX = 0
    plcIFBulk.CurrentY = 0
    plcIFBulk.Print "Generate Cross Reference Information in Bulk Feed Export"
End Sub
Private Sub plcIFExport_Paint()
    'plcIFExport.CurrentX = 0
    'plcIFExport.CurrentY = 0
    'plcIFExport.Print "Select for Export in"
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPEnableBox                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mPEnableBox(ilBoxNo As Integer)


    If (ilBoxNo < imLBPCtrls) Or (ilBoxNo > UBound(tmPCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case PRODUCERINDEX 'Producer
            mProducerPop
            If imTerminate Then
                Exit Sub
            End If
            lbcProducer.Height = gListBoxHeight(lbcProducer.ListCount, 6)
            edcPDropdown.Width = tmPCtrls(ilBoxNo).fBoxW - cmcPDropdown.Width
            edcPDropdown.MaxLength = 40
            gMoveFormCtrl pbcProducer, edcPDropdown, tmPCtrls(ilBoxNo).fBoxX - pbcProducer.Left, tmPCtrls(ilBoxNo).fBoxY - pbcProducer.Top
            cmcPDropdown.Move edcPDropdown.Left + edcPDropdown.Width, edcPDropdown.Top
            lbcProducer.Move edcPDropdown.Left + pbcProducer.Left, edcPDropdown.Top + edcPDropdown.Height + pbcProducer.Top
            imChgMode = True
            If lbcProducer.ListIndex < 0 Then
                If lbcProducer.ListCount > 2 Then
                    lbcProducer.ListIndex = 2  'Pick first name
                Else
                    lbcProducer.ListIndex = 1   '[New]
                End If
            End If
            If lbcProducer.ListIndex < 0 Then
                edcPDropdown.Text = ""
            Else
                edcPDropdown.Text = lbcProducer.List(lbcProducer.ListIndex)
            End If
            imChgMode = False
            edcPDropdown.SelStart = 0
            edcPDropdown.SelLength = Len(edcPDropdown.Text)
            edcPDropdown.Visible = True
            cmcPDropdown.Visible = True
            edcPDropdown.SetFocus
'        Case CONTENTPROVIDERINDEX 'Content Provider
'            mContentProviderPop
'            If imTerminate Then
'                Exit Sub
'            End If
'            lbcContentProvider.Height = gListBoxHeight(lbcContentProvider.ListCount, 6)
'            edcPDropdown.Width = tmPCtrls(ilBoxNo).fBoxW - cmcPDropdown.Width
'            edcPDropdown.MaxLength = 40
'            gMoveFormCtrl pbcProducer, edcPDropdown, tmPCtrls(ilBoxNo).fBoxX - pbcProducer.Left, tmPCtrls(ilBoxNo).fBoxY - pbcProducer.Top
'            cmcPDropdown.Move edcPDropdown.Left + edcPDropdown.Width, edcPDropdown.Top
'            lbcContentProvider.Move edcPDropdown.Left + pbcProducer.Left, edcPDropdown.Top + edcPDropdown.Height + pbcProducer.Top
'            imChgMode = True
'            If lbcContentProvider.ListIndex < 0 Then
'                If lbcContentProvider.ListCount > 2 Then
'                    lbcContentProvider.ListIndex = 2  'Pick first name
'                Else
'                    lbcContentProvider.ListIndex = 1   '[New]
'                End If
'            End If
'            If lbcContentProvider.ListIndex < 0 Then
'                edcPDropdown.Text = ""
'            Else
'                edcPDropdown.Text = lbcContentProvider.List(lbcContentProvider.ListIndex)
'            End If
'            imChgMode = False
'            edcPDropdown.SelStart = 0
'            edcPDropdown.SelLength = Len(edcPDropdown.Text)
'            edcPDropdown.Visible = True
'            cmcPDropdown.Visible = True
'            edcPDropdown.SetFocus
        Case EXPPROGAUDIOINDEX 'Export Program Audio
            mContentProviderPop
            If imTerminate Then
                Exit Sub
            End If
            lbcExpProgAudio.Height = gListBoxHeight(lbcExpProgAudio.ListCount, 6)
            edcPDropdown.Width = tmPCtrls(ilBoxNo).fBoxW - cmcPDropdown.Width
            edcPDropdown.MaxLength = 40
            gMoveFormCtrl pbcProducer, edcPDropdown, tmPCtrls(ilBoxNo).fBoxX - pbcProducer.Left, tmPCtrls(ilBoxNo).fBoxY - pbcProducer.Top
            cmcPDropdown.Move edcPDropdown.Left + edcPDropdown.Width, edcPDropdown.Top
            lbcExpProgAudio.Move edcPDropdown.Left + pbcProducer.Left, edcPDropdown.Top + edcPDropdown.Height + pbcProducer.Top
            imChgMode = True
            If lbcExpProgAudio.ListIndex < 0 Then
                If lbcExpProgAudio.ListCount > 1 Then
                    lbcExpProgAudio.ListIndex = 1  'Pick first name
                Else
                    lbcExpProgAudio.ListIndex = 0   '[New]
                End If
            End If
            If lbcExpProgAudio.ListIndex < 0 Then
                edcPDropdown.Text = ""
            Else
                edcPDropdown.Text = lbcExpProgAudio.List(lbcExpProgAudio.ListIndex)
            End If
            imChgMode = False
            edcPDropdown.SelStart = 0
            edcPDropdown.SelLength = Len(edcPDropdown.Text)
            edcPDropdown.Visible = True
            cmcPDropdown.Visible = True
            edcPDropdown.SetFocus
        Case EXPCOMMAUDIOINDEX '# of Days
            mContentProviderPop
            If imTerminate Then
                Exit Sub
            End If
            lbcExpCommAudio.Height = gListBoxHeight(lbcExpCommAudio.ListCount, 6)
            edcPDropdown.Width = tmPCtrls(ilBoxNo).fBoxW - cmcPDropdown.Width
            edcPDropdown.MaxLength = 40
            gMoveFormCtrl pbcProducer, edcPDropdown, tmPCtrls(ilBoxNo).fBoxX - pbcProducer.Left, tmPCtrls(ilBoxNo).fBoxY - pbcProducer.Top
            cmcPDropdown.Move edcPDropdown.Left + edcPDropdown.Width, edcPDropdown.Top
            lbcExpCommAudio.Move edcPDropdown.Left + pbcProducer.Left, edcPDropdown.Top + edcPDropdown.Height + pbcProducer.Top
            imChgMode = True
            If lbcExpCommAudio.ListIndex < 0 Then
                If lbcExpCommAudio.ListCount > 1 Then
                    lbcExpCommAudio.ListIndex = 1  'Pick first name
                Else
                    lbcExpCommAudio.ListIndex = 0   '[New]
                End If
            End If
            If lbcExpCommAudio.ListIndex < 0 Then
                edcPDropdown.Text = ""
            Else
                edcPDropdown.Text = lbcExpCommAudio.List(lbcExpCommAudio.ListIndex)
            End If
            imChgMode = False
            edcPDropdown.SelStart = 0
            edcPDropdown.SelLength = Len(edcPDropdown.Text)
            edcPDropdown.Visible = True
            cmcPDropdown.Visible = True
            edcPDropdown.SetFocus
        Case COMMEMBEDDEDINDEX
            pbcCommEmbedded.Width = tmPCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcProducer, pbcCommEmbedded, tmPCtrls(ilBoxNo).fBoxX - pbcProducer.Left, tmPCtrls(ilBoxNo).fBoxY - pbcProducer.Top
            If (tmVpf.sEmbeddedComm <> "Y") And (tmVpf.sEmbeddedComm <> "N") Then
                tmVpf.sEmbeddedComm = "N"
                imProducerAltered = True
            End If
            pbcCommEmbedded.Visible = True  'Set visibility
            pbcCommEmbedded.SetFocus

    End Select

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPSetShow                       *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mPSetShow(ilBoxNo As Integer)

    Dim slStr As String

    If (ilBoxNo < imLBPCtrls) Or (ilBoxNo > UBound(tmPCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case PRODUCERINDEX
            lbcProducer.Visible = False  'Set visibility
            edcPDropdown.Visible = False
            cmcPDropdown.Visible = False
            If lbcProducer.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcProducer.List(lbcProducer.ListIndex)
            End If
            gSetShow pbcProducer, slStr, tmPCtrls(ilBoxNo)
'        Case CONTENTPROVIDERINDEX
'            lbcContentProvider.Visible = False  'Set visibility
'            edcPDropdown.Visible = False
'            cmcPDropdown.Visible = False
'            If lbcContentProvider.ListIndex <= 0 Then
'                slStr = ""
'            Else
'                slStr = lbcContentProvider.List(lbcContentProvider.ListIndex)
'            End If
'            gSetShow pbcProducer, slStr, tmPCtrls(ilBoxNo)
        Case EXPPROGAUDIOINDEX
            lbcExpProgAudio.Visible = False  'Set visibility
            edcPDropdown.Visible = False
            cmcPDropdown.Visible = False
            If lbcExpProgAudio.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcExpProgAudio.List(lbcExpProgAudio.ListIndex)
            End If
            gSetShow pbcProducer, slStr, tmPCtrls(ilBoxNo)
        Case EXPCOMMAUDIOINDEX
            lbcExpCommAudio.Visible = False  'Set visibility
            edcPDropdown.Visible = False
            cmcPDropdown.Visible = False
            If lbcExpCommAudio.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcExpCommAudio.List(lbcExpCommAudio.ListIndex)
            End If
            gSetShow pbcProducer, slStr, tmPCtrls(ilBoxNo)
        Case COMMEMBEDDEDINDEX
            pbcCommEmbedded.Visible = False
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPSetFocus                      *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mPSetFocus(ilBoxNo As Integer)
    If (ilBoxNo < imLBPCtrls) Or (ilBoxNo > UBound(tmPCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case PRODUCERINDEX 'Producer
            edcPDropdown.Visible = True
            cmcPDropdown.Visible = True
            edcPDropdown.SetFocus
'        Case CONTENTPROVIDERINDEX 'Content Provider
'            edcPDropdown.Visible = True
'            cmcPDropdown.Visible = True
'            edcPDropdown.SetFocus
        Case EXPPROGAUDIOINDEX 'Export Program Audio
            edcPDropdown.Visible = True
            cmcPDropdown.Visible = True
            edcPDropdown.SetFocus
        Case EXPCOMMAUDIOINDEX '# of Days
            edcPDropdown.Visible = True
            cmcPDropdown.Visible = True
            edcPDropdown.SetFocus
        Case COMMEMBEDDEDINDEX
            pbcCommEmbedded.Visible = True  'Set visibility
            pbcCommEmbedded.SetFocus
    End Select

End Sub

Private Sub mInitArf()
    tmArf.iCode = 0
    tmArf.sType = "F" 'L=Lock box; A=Agency DP; S=Sales Office (not Used); C=Content Provider; P = Producer; F=FTP (Used in vehicle option)
    tmArf.sID = "FTP" 'Identification name (Type L and A); Abbreviation (Type C and P)
    tmArf.sNmAd(0) = "" 'Name and address
    tmArf.sNmAd(1) = "" 'Name and address
    tmArf.sNmAd(2) = "" 'Name and address
    tmArf.sNmAd(3) = "" 'Name and address
    tmArf.iMerge = 0    'Merge code number
    tmArf.sName = ""    'For Type C and P
    tmArf.sEMail = ""   'E-Mail address for Type C and P
    tmArf.sWebSite = ""  'Web Site address for Type C and P
    tmArf.sSendISCITo = ""   'Send ISCI Export to Producer or Content Provider, Type C and P
    tmArf.sContactName = "" 'Contact Name
    tmArf.sContactPhone = ""   'Contact Phone number
    tmArf.sUnused = ""
End Sub


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
            imIgnoreChg = True
            edcLevelPrice.Width = tmSCtrls(ilBoxNo).fBoxW
            edcLevelPrice.MaxLength = 0
            gMoveTableCtrl pbcSchedule, edcLevelPrice, tmSCtrls(ilBoxNo).fBoxX, tmSCtrls(ilBoxNo).fBoxY + (imSRowNo - 1) * (fgBoxGridH + 15)
            edcLevelPrice.Text = Trim$(Str$(lmSSave(ilBoxNo - LEVEL2INDEX + 1)))
            edcLevelPrice.Enabled = True
            edcLevelPrice.Visible = True  'Set visibility
            edcLevelPrice.SetFocus
            imIgnoreChg = False
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
                imLevelAltered = True
                lmSSave(ilBoxNo - LEVEL2INDEX + 1) = Val(slStr)
                If ilBoxNo = LEVEL2INDEX Then
                    edcSchedule(0).Text = slStr
                End If
                If ilBoxNo = LEVEL14INDEX Then
                    edcSchedule(1).Text = slStr
                End If
                pbcSchedule.Cls
                pbcSchedule_Paint
                mSetCommands
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
'*      Procedure Name:mSafReadRec                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mSafReadRec(ilVefCode As Integer) As Integer
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

    If (tmVef.sType <> "C") And (tmVef.sType <> "S") And (tmVef.sType <> "G") Then
        mSafReadRec = True
        Exit Function
    End If
    If ilVefCode <> 0 Then
        tmSafSrchKey1.iVefCode = ilVefCode
        ilRet = btrGetEqual(hmSaf, tmSaf, imSafRecLen, tmSafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Else
        ilRet = Not BTRV_ERR_NONE
    End If
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
        tmSaf.iVefCode = tmVef.iCode
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
        tmSaf.lGGEMailSentDate = 0
        tmSaf.iRetainTrafProj = 24
        tmSaf.iRetainCntr = 60
        tmSaf.iRetainPayRevHist = 60
        tmSaf.iNoDaysRetainUAF = 5
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

'*******************************************************
'*                                                     *
'*      Procedure Name:mVafReadRec                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mVafReadRec(ilVefCode As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mVafReadRecErr                                                                        *
'******************************************************************************************

'
'   iRet = mVafReadRec()
'   Where:
'       ilVefCode(I) - Vehicle Code
'       slType (I) - L=Log; C=CP and O= Other
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "N") Or (tmVef.sType = "R") Or (tmVef.sType = "G")) And ((Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS) Then
        If ilVefCode <> 0 Then
            tmVafSrchKey1.iVefCode = ilVefCode
            ilRet = btrGetEqual(hmVaf, tmVaf, imVafRecLen, tmVafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        Else
            ilRet = Not BTRV_ERR_NONE
        End If
        If ilRet <> BTRV_ERR_NONE Then
            'Add Record
            tmVaf.lCode = 0
            tmVaf.iVefCode = ilVefCode
            tmVaf.sDivisionCode = ""
            tmVaf.sBranchCodeCash = ""
            tmVaf.sBranchCodeTrade = ""
            tmVaf.sPCGrossSalesCash = ""
            tmVaf.sPCAgyCommCash = ""
            tmVaf.sPCRecvCash = ""
            tmVaf.sPCGrossSalesTrade = ""
            tmVaf.sPCRecvTrade = ""
            tmVaf.sVendorID = ""                '8-25-10
            tmVaf.sUnused = ""
        End If
    Else
        tmVaf.sDivisionCode = ""
        tmVaf.sBranchCodeCash = ""
        tmVaf.sBranchCodeTrade = ""
        tmVaf.sPCGrossSalesCash = ""
        tmVaf.sPCAgyCommCash = ""
        tmVaf.sPCRecvCash = ""
        tmVaf.sPCGrossSalesTrade = ""
        tmVaf.sPCRecvTrade = ""
        tmVaf.sVendorID = ""                '8-25-10
    End If
    mVafReadRec = True
    Exit Function
mVafReadRecErr: 'VBC NR
    On Error GoTo 0
    mVafReadRec = False
    Exit Function
End Function
'
'           mGetWeeksInYear - given the year, detrmine the # of weeks in
'           the standard broadcast year
'
Public Function mGetWeeksInYear(slYear As String)
Dim slStart As String
Dim slEnd As String
Dim llStart As Long
Dim llEnd As Long
Dim ilWeek As Integer

    ilWeek = 0
    If slYear <> "" Then
        slStart = "1/15/" & Trim$(slYear)
        slEnd = "12/15/" & Trim$(slYear)
        llStart = gDateValue(gObtainStartStd(slStart))    'get the std start date for the year
        llEnd = gDateValue(gObtainEndStd(slEnd))       'get the std end date for theyear
       ilWeek = ((llEnd - llStart) + 1) / 7
    End If
    mGetWeeksInYear = ilWeek


End Function
'
'                   mUpdateNIFCounts - update the weekly buckets by either
'                   the value entered for the week, or if the value is
'                   by the year, distribute the amount across all weeks in theyear
'
Public Sub mUpdateNifCounts()
Dim slStr As String
Dim ilWeeks As Integer
Dim ilLoop As Integer
Dim ilInvPerWeek As Integer

    slStr = Str$(imYear)
    ilWeeks = mGetWeeksInYear(slStr)
    If ilWeeks = 0 Then
        ilWeeks = 52                'default to 52 week year
    End If

    If smByWeekOrYear = "Y" Then    'the entered # minutes is for the entire year
        lmInventory = Val(edcInventory.Text) * 60
        lmInventory = lmInventory + (Val(edcSec.Text))
        ilInvPerWeek = lmInventory \ ilWeeks
        For ilLoop = 1 To 53 'ilWeeks
            If ilLoop > ilWeeks Then
                tmNif.lInvCount(ilLoop - 1) = 0
            Else
                tmNif.lInvCount(ilLoop - 1) = ilInvPerWeek
                If (lmInventory - ilInvPerWeek) >= 0 Then
                    lmInventory = lmInventory - ilInvPerWeek        'keep track of whats been allocated so far
                End If
            End If
        Next ilLoop
        tmNif.lInvCount(ilWeeks - 1) = tmNif.lInvCount(ilWeeks - 1) + lmInventory
    Else
        'entered value if for each week
        lmInventory = Val(edcInventory.Text) * 60
        lmInventory = lmInventory + (Val(edcSec.Text))
        For ilLoop = 1 To ilWeeks
            tmNif.lInvCount(ilLoop - 1) = lmInventory
        Next ilLoop
    End If
    Exit Sub
End Sub

Public Sub mObtainNIFYear()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLen                                                                                 *
'******************************************************************************************

Dim ilRet As Integer
Dim ilLoop As Integer
Dim slMinutes As String
Dim slSeconds As String
        lmInventory = 0
        'see if anything exists for the year
        tmNifSrchKey1.iVefCode = tmVef.iCode
        tmNifSrchKey1.iYear = imYear
        ilRet = btrGetEqual(hmNif, tmNif, imNifRecLen, tmNifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            'got a record, show the values
            edcYear.Text = Trim$(Str$(imYear))
            If tmNif.sInvWkYear <> "W" Then         'always assume Yearly invntory
                rbcInvBy(1).Value = True
                For ilLoop = 1 To 53
                    lmInventory = lmInventory + tmNif.lInvCount(ilLoop - 1)
                Next ilLoop
                ckcRollover.Enabled = False
            Else
                rbcInvBy(0).Value = True               'weekly avails
                'lmInventory = tmNif.lInvCount(1)        'each week is the same for now
                lmInventory = tmNif.lInvCount(0)        'each week is the same for now
                ckcRollover.Enabled = True
            End If
            If tmNif.sAllowRollover <> "N" Then         'assume yearly and rollover applicable
                ckcRollover.Value = vbChecked

            Else
                ckcRollover.Value = vbUnchecked

            End If


            slMinutes = lmInventory \ 60
            slSeconds = lmInventory - (Val(slMinutes) * 60)

            edcInventory.Text = Trim$(slMinutes)

            edcSec.Text = Trim$(slSeconds)

        Else
            edcInventory.Text = ""
            edcSec.Text = ""
            ckcRollover.Enabled = False
            For ilLoop = 1 To 53
                tmNif.lInvCount(ilLoop - 1) = 0
            Next ilLoop
            tmNif.iVefCode = tmVef.iCode
            tmNif.iYear = Val(edcYear.Text)
            tmNif.lCode = 0
        End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mVbfReadRec                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mVbfReadRec(ilVefCode As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mVbfReadRecErr                                                                        *
'******************************************************************************************

'
'   iRet = mVbfReadRec()
'   Where:
'       ilVefCode(I) - Vehicle Code
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilValue As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilLoop As Integer
    Dim tlVbf As VBF

    ReDim tmVbf(0 To 0) As VBF
    ReDim smVbfComment(0 To 0) As String
    imVbfChg = False
    imVBFIndex = -1
    cbcBarter.Clear
    pbcAcq(0).Cls
    pbcAcq(1).Cls
    frcBarterEnable(0).Enabled = False
    frcBarterEnable(1).Enabled = False
    frcBarterEnable(2).Enabled = False
    ilValue = Asc(tgSpf.sUsingFeatures2)
    ''If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G")) And (ilValue And BARTER) = BARTER Then
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S")) And (ilValue And BARTER) = BARTER Then
    If (tmVef.sType = "R") And ((ilValue And BARTER) = BARTER) Then
        If ilVefCode <> 0 Then
            frcBarterEnable(0).Enabled = True
            tmVbfSrchKey1.iVefCode = ilVefCode
            tmVbfSrchKey1.iStartDate(0) = 0
            tmVbfSrchKey1.iStartDate(1) = 0
            ilRet = btrGetGreaterOrEqual(hmVbf, tlVbf, imVBfRecLen, tmVbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tlVbf.iVefCode = ilVefCode)
                tmVbf(UBound(tmVbf)) = tlVbf
                gUnpackDate tlVbf.iStartDate(0), tlVbf.iStartDate(1), slStartDate
                gUnpackDate tlVbf.iEndDate(0), tlVbf.iEndDate(1), slEndDate
                If slEndDate <> "" Then
                    cbcBarter.AddItem slStartDate & "-" & slEndDate
                Else
                    cbcBarter.AddItem slStartDate & "-" & "TFN"
                End If
                cbcBarter.ItemData(cbcBarter.NewIndex) = tlVbf.lCode
                ReDim Preserve tmVbf(0 To UBound(tmVbf) + 1) As VBF
                smVbfComment(UBound(smVbfComment)) = ""
                If tlVbf.lInsertionCefCode > 0 Then
                    tmCefSrchKey.lCode = tlVbf.lInsertionCefCode
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        'If tmCef.iStrLen > 0 Then
                        '    smVbfComment(UBound(smVbfComment)) = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                        'End If
                        smVbfComment(UBound(smVbfComment)) = gStripChr0(tmCef.sComment)
                    End If
                End If
                ReDim Preserve smVbfComment(0 To UBound(smVbfComment) + 1) As String
                ilRet = btrGetNext(hmVbf, tlVbf, imVBfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
            Loop
            cbcBarter.AddItem "[New]", 0
            tlVbf.lCode = 0
            tlVbf.iVefCode = ilVefCode
            cbcBarter.ItemData(cbcBarter.NewIndex) = 0
            gPackDate "", tlVbf.iStartDate(0), tlVbf.iStartDate(1)
            gPackDate "", tlVbf.iEndDate(0), tlVbf.iEndDate(1)
            'For ilLoop = 1 To 10 Step 1
            For ilLoop = 0 To 9 Step 1
                tlVbf.iSpotLen(ilLoop) = 0
                tlVbf.lDefAcqCost(ilLoop) = 0
                tlVbf.lActAcqCost(ilLoop) = 0
            Next ilLoop
            tlVbf.iThreshold = 0
            tlVbf.iXFree = 0
            tlVbf.iYSold = 0
            tlVbf.sMethod = "N"
            tlVbf.lInsertionCefCode = 0
            tlVbf.iAcqCommPct = 0
            tlVbf.lBalance = 0
            gPackDate "", tlVbf.iBalanceDate(0), tlVbf.iBalanceDate(1)
            tlVbf.sPerPeriod = "W"
            tmVbf(UBound(tmVbf)) = tlVbf
            ReDim Preserve tmVbf(0 To UBound(tmVbf) + 1) As VBF
            smVbfComment(UBound(smVbfComment)) = ""
            ReDim Preserve smVbfComment(0 To UBound(smVbfComment) + 1) As String
        End If
    ElseIf ((tmVef.sType = "C") Or (tmVef.sType = "S")) And ((ilValue And BARTER) = BARTER) Then
        If ilVefCode <> 0 Then
            frcBarterEnable(2).Enabled = True
            tmVbfSrchKey1.iVefCode = ilVefCode
            gPackDate "1/1/1970", tmVbfSrchKey1.iStartDate(0), tmVbfSrchKey1.iStartDate(1)
            ilRet = btrGetEqual(hmVbf, tlVbf, imVBfRecLen, tmVbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If (ilRet <> BTRV_ERR_NONE) Then
                mInitVbf ilVefCode, tlVbf
                gPackDate "1/1/1970", tlVbf.iStartDate(0), tlVbf.iStartDate(1)
                gPackDate "12/31/2069", tlVbf.iEndDate(0), tlVbf.iEndDate(1)
                tlVbf.sMethod = "I"
            End If
            tmVbfIndex = tlVbf
            mSetAcqIndexLen
            If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then 'Acquisition Commissionable
                frcBarterEnable(1).Enabled = True
                tmVbfSrchKey1.iVefCode = ilVefCode
                gPackDate "1/2/1970", tmVbfSrchKey1.iStartDate(0), tmVbfSrchKey1.iStartDate(1)
                ilRet = btrGetGreaterOrEqual(hmVbf, tlVbf, imVBfRecLen, tmVbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tlVbf.iVefCode = ilVefCode)
                    tmVbf(UBound(tmVbf)) = tlVbf
                    gUnpackDate tlVbf.iStartDate(0), tlVbf.iStartDate(1), slStartDate
                    gUnpackDate tlVbf.iEndDate(0), tlVbf.iEndDate(1), slEndDate
                    If slEndDate <> "" Then
                        cbcBarter.AddItem slStartDate & "-" & slEndDate
                    Else
                        cbcBarter.AddItem slStartDate & "-" & "TFN"
                    End If
                    cbcBarter.ItemData(cbcBarter.NewIndex) = tlVbf.lCode
                    ReDim Preserve tmVbf(0 To UBound(tmVbf) + 1) As VBF
                    ilRet = btrGetNext(hmVbf, tlVbf, imVBfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                Loop
                mInitVbf ilVefCode, tlVbf
                cbcBarter.AddItem "[New]", 0
                cbcBarter.ItemData(cbcBarter.NewIndex) = 0
                tmVbf(UBound(tmVbf)) = tlVbf
                ReDim Preserve tmVbf(0 To UBound(tmVbf) + 1) As VBF
            End If
        End If
    End If
    mVbfReadRec = True
    Exit Function
mVbfReadRecErr: 'VBC NR
    On Error GoTo 0
    mVbfReadRec = False
    Exit Function
End Function

Private Sub mMoveCtrlToVbf()
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    

    If imVBFIndex <> -1 Then
        If tmVef.sType = "R" Then
            slStr = edcBarter(0).Text
            If gValidDate(slStr) Then
                gPackDate slStr, tmVbf(imVBFIndex).iStartDate(0), tmVbf(imVBFIndex).iStartDate(1)
            Else
                Beep
            End If
            If rbcBarterMethod(0).Value = True Then
                tmVbf(imVBFIndex).sMethod = "A"
                tmVbf(imVBFIndex).sPerPeriod = "W"
            ElseIf rbcBarterMethod(1).Value = True Then
                tmVbf(imVBFIndex).sMethod = "M"
                tmVbf(imVBFIndex).iThreshold = Val(edcBarterMethod(0).Text)
                If Left(cbcPerPeriod(0).Text, 1) = "M" Then
                    tmVbf(imVBFIndex).sPerPeriod = "M"
                ElseIf Left(cbcPerPeriod(0).Text, 1) = "Y" Then
                    tmVbf(imVBFIndex).sPerPeriod = "Y"
                Else
                    tmVbf(imVBFIndex).sPerPeriod = "W"
                End If
            ElseIf rbcBarterMethod(2).Value = True Then
                tmVbf(imVBFIndex).sMethod = "U"
                tmVbf(imVBFIndex).iThreshold = Val(edcBarterMethod(1).Text)
                If Left(cbcPerPeriod(1).Text, 1) = "M" Then
                    tmVbf(imVBFIndex).sPerPeriod = "M"
                ElseIf Left(cbcPerPeriod(1).Text, 1) = "Y" Then
                    tmVbf(imVBFIndex).sPerPeriod = "Y"
                Else
                    tmVbf(imVBFIndex).sPerPeriod = "W"
                End If
            ElseIf rbcBarterMethod(3).Value = True Then
                tmVbf(imVBFIndex).sMethod = "X"
                tmVbf(imVBFIndex).iXFree = Val(edcBarterMethod(2).Text)
                tmVbf(imVBFIndex).iYSold = Val(edcBarterMethod(3).Text)
                If Left(cbcPerPeriod(2).Text, 1) = "M" Then
                    tmVbf(imVBFIndex).sPerPeriod = "M"
                ElseIf Left(cbcPerPeriod(2).Text, 1) = "Y" Then
                    tmVbf(imVBFIndex).sPerPeriod = "Y"
                Else
                    tmVbf(imVBFIndex).sPerPeriod = "W"
                End If
            ElseIf rbcBarterMethod(4).Value = True Then
                tmVbf(imVBFIndex).sMethod = "N"
                tmVbf(imVBFIndex).iThreshold = 0
                tmVbf(imVBFIndex).iXFree = 0
                tmVbf(imVBFIndex).iYSold = 0
                tmVbf(imVBFIndex).sPerPeriod = "W"
            End If
            tmVbf(imVBFIndex).iAcqCommPct = 0
            smVbfComment(imVBFIndex) = edcInsertComment.Text
        ElseIf (tmVef.sType = "C") Or (tmVef.sType = "S") Then
            If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then 'Acquisition Commissionable
                slStr = edcBarter(1).Text
                If gValidDate(slStr) Then
                    gPackDate slStr, tmVbf(imVBFIndex).iStartDate(0), tmVbf(imVBFIndex).iStartDate(1)
                Else
                    Beep
                End If
                slStr = edcBarter(2).Text
                tmVbf(imVBFIndex).iAcqCommPct = gStrDecToInt(slStr, 2)
            Else
                gPackDate "1/1/1970", tmVbf(imVBFIndex).iStartDate(0), tmVbf(imVBFIndex).iStartDate(1)
                tmVbf(imVBFIndex).iAcqCommPct = 0
            End If
            'For ilLoop = 1 To 10 Step 1
            For ilLoop = 0 To 9 Step 1
                'tmVbf(imVBFIndex).iSpotLen(ilLoop) = 0
                'tmVbf(imVBFIndex).lDefAcqCost(ilLoop) = 0
                tmVbf(imVBFIndex).lActAcqCost(ilLoop) = 0
            Next ilLoop
            tmVbf(imVBFIndex).sMethod = "N"
            tmVbf(imVBFIndex).iThreshold = 0
            tmVbf(imVBFIndex).iXFree = 0
            tmVbf(imVBFIndex).iYSold = 0
            tmVbf(imVBFIndex).sPerPeriod = "W"
        End If
    End If
End Sub

Private Sub mMoveVbfToCtrl()
    Dim slStr As String
    Dim ilVbfChg As Integer
    Dim ilLen As Integer
    Dim llMinutes As Long
    Dim llSeconds As Long

    ilVbfChg = imVbfChg
    If imVBFIndex <> -1 Then
        mClearBarterCtrl False
        If tmVef.sType = "R" Then
            If tmVbf(imVBFIndex).lCode > 0 Then
                edcBarter(0).BackColor = &H80FFFF
                edcBarter(0).Enabled = False
                mSetAcqLen
            Else
                edcBarter(0).BackColor = &HFFFF00
                edcBarter(0).Enabled = True
                mSetAcqLen
                For ilLen = LBound(tmVbf(imVBFIndex).iSpotLen) To LBound(tmVbf(imVBFIndex).iSpotLen) Step 1
                    tmVbf(imVBFIndex).lDefAcqCost(ilLen) = 0
                    tmVbf(imVBFIndex).lActAcqCost(ilLen) = 0
                Next ilLen
            End If
            gUnpackDate tmVbf(imVBFIndex).iStartDate(0), tmVbf(imVBFIndex).iStartDate(1), slStr
            edcBarter(0).Text = slStr
            Select Case Trim$(tmVbf(imVBFIndex).sMethod)
                Case "A"
                    rbcBarterMethod(0).Value = True
                Case "M"
                    rbcBarterMethod(1).Value = True
                    edcBarterMethod(0).Text = Trim$(Str$(tmVbf(imVBFIndex).iThreshold))
                    edcBarterMethod(0).Enabled = True
                    cbcPerPeriod(0).Enabled = True
                    If tmVbf(imVBFIndex).sPerPeriod = "M" Then
                        cbcPerPeriod(0).ListIndex = 1
                    ElseIf tmVbf(imVBFIndex).sPerPeriod = "Y" Then
                        cbcPerPeriod(0).ListIndex = 2
                        If tmVbf(imVBFIndex).iBalanceDate(0) = 0 And tmVbf(imVBFIndex).iBalanceDate(1) = 0 Then
                        tmVbf(imVBFIndex).lBalance = CLng(tmVbf(imVBFIndex).iThreshold) * 60   'never updated, show full threshold available
                        End If
                        'convert seconds to minutes, seconds
                        llSeconds = tmVbf(imVBFIndex).lBalance
                        llMinutes = (llSeconds / 60)
                        llSeconds = llSeconds - (llMinutes * 60)
                        edcBarterMethod(4).Text = Str$(llMinutes) & "M" & Str$(llSeconds) & "S"
                    Else
                        cbcPerPeriod(0).ListIndex = 0
                    End If
                Case "U"
                    rbcBarterMethod(2).Value = True
                    edcBarterMethod(1).Text = Trim$(Str$(tmVbf(imVBFIndex).iThreshold))
                    edcBarterMethod(1).Enabled = True
                    cbcPerPeriod(1).Enabled = True
                    If tmVbf(imVBFIndex).sPerPeriod = "M" Then
                        cbcPerPeriod(1).ListIndex = 1
                    ElseIf tmVbf(imVBFIndex).sPerPeriod = "Y" Then
                        cbcPerPeriod(1).ListIndex = 2
                        edcBarterMethod(5).Text = Trim$(Str$(tmVbf(imVBFIndex).lBalance))
                    Else
                        cbcPerPeriod(1).ListIndex = 0
                    End If
                Case "X"
                    rbcBarterMethod(3).Value = True
                    edcBarterMethod(2).Text = Trim$(Str$(tmVbf(imVBFIndex).iXFree))
                    edcBarterMethod(3).Text = Trim$(Str$(tmVbf(imVBFIndex).iYSold))
                    edcBarterMethod(2).Enabled = True
                    edcBarterMethod(3).Enabled = True
                    cbcPerPeriod(2).Enabled = True
                    If tmVbf(imVBFIndex).sPerPeriod = "M" Then
                        cbcPerPeriod(2).ListIndex = 1
                    ElseIf tmVbf(imVBFIndex).sPerPeriod = "Y" Then
                        cbcPerPeriod(2).ListIndex = 2
                    Else
                        cbcPerPeriod(2).ListIndex = 0
                    End If
                Case Else
                    rbcBarterMethod(4).Value = True
            End Select
            edcInsertComment.Text = Trim$(smVbfComment(imVBFIndex))
            mAcqCostPaint
        ElseIf (tmVef.sType = "C") Or (tmVef.sType = "S") Then
            If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then 'Acquisition Commissionable
                If tmVbf(imVBFIndex).lCode > 0 Then
                    edcBarter(1).BackColor = &H80FFFF
                    edcBarter(1).Enabled = False
                    mSetAcqLen
                Else
                    edcBarter(1).BackColor = &HFFFF00
                    edcBarter(1).Enabled = True
                End If
                gUnpackDate tmVbf(imVBFIndex).iStartDate(0), tmVbf(imVBFIndex).iStartDate(1), slStr
                edcBarter(1).Text = slStr
                If tmVbf(imVBFIndex).iAcqCommPct > 0 Then
                    edcBarter(2).Text = gIntToStrDec(tmVbf(imVBFIndex).iAcqCommPct, 2)
                Else
                    edcBarter(2).Text = ""
                End If
            End If
            mAcqIndexPaint
        End If
    End If
    imVbfChg = ilVbfChg
End Sub

Private Sub mClearBarterCtrl(blClearAll As Boolean)
    'edcBarter(0).Text = ""
    'edcBarter(1).Text = ""
    If (tmVef.sType = "R") Or (tmVef.sType = "C") Or (tmVef.sType = "S") Then
        edcBarter(3).Text = ""
        edcBarter(4).Text = ""
        edcBarter(0).Text = ""
        rbcBarterMethod(4).Value = True
        edcBarterMethod(0).Text = ""
        edcBarterMethod(1).Text = ""
        edcBarterMethod(2).Text = ""
        edcBarterMethod(3).Text = ""
        edcInsertComment.Text = ""
        edcBarter(1).Text = ""
        edcBarter(2).Text = ""
        If blClearAll Then
            rbcBarterMethod(5).Value = True
            ckcBarter(0).Value = vbUnchecked
            edcBarter(5).Text = ""
            lacBarter(8).Enabled = False
            edcBarter(5).Enabled = False
        End If
        cbcPerPeriod(0).ListIndex = -1
        cbcPerPeriod(1).ListIndex = -1
        cbcPerPeriod(2).ListIndex = -1
    End If
End Sub

Private Sub mAdjVbfDates()
    Dim ilValue As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilVbf As Integer

    ilValue = Asc(tgSpf.sUsingFeatures2)
    ''If ((tmVef.sType = "C") Or (tmVef.sType = "S") Or (tmVef.sType = "G")) And (ilValue And BARTER) = BARTER Then
    'If ((tmVef.sType = "C") Or (tmVef.sType = "S")) And (ilValue And BARTER) = BARTER Then
    If ((tmVef.sType = "R") Or (tmVef.sType = "C") Or (tmVef.sType = "S")) And ((ilValue And BARTER) = BARTER) Then
        For ilLoop = 0 To UBound(tmVbf) - 1 Step 1
            If tmVbf(ilLoop).lCode = 0 Then
                gUnpackDate tmVbf(ilLoop).iStartDate(0), tmVbf(ilLoop).iStartDate(1), slStr
                llDate = gDateValue(slStr)
                If slStr <> "" Then
                    For ilVbf = 0 To UBound(tmVbf) - 1 Step 1
                        If ilVbf <> ilLoop Then
                            gUnpackDateLong tmVbf(ilVbf).iStartDate(0), tmVbf(ilVbf).iStartDate(1), llStartDate
                            gUnpackDateLong tmVbf(ilVbf).iEndDate(0), tmVbf(ilVbf).iEndDate(1), llEndDate
                            If llEndDate = 0 Then
                                llEndDate = 999999
                            End If
                            If llStartDate <= llEndDate Then
                                If llDate < llEndDate Then
                                    If llDate < llStartDate Then
                                        llEndDate = llStartDate - 1
                                    Else
                                        llEndDate = llDate - 1
                                    End If
                                End If
                                gPackDateLong llEndDate, tmVbf(ilVbf).iEndDate(0), tmVbf(ilVbf).iEndDate(1)
                            End If
                        End If
                    Next ilVbf
                End If
                Exit For
            End If
        Next ilLoop
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetSSpotLen                    *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set spot length box            *
'*                                                     *
'*******************************************************
Private Sub mSetAcqCost()

    Dim flLeft As Single
    Dim flTop As Single

    If imVBFIndex < 0 Then
        Exit Sub
    End If
    If (imAcqCostBoxNo < 0) Or (imAcqCostBoxNo > 2 * (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD) - 1) Then
        Exit Sub
    End If
    If imAcqCostBoxNo <= (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD) - 1 Then
        edcBarter(3).Text = gLongToStrDec(tmVbf(imVBFIndex).lDefAcqCost(imAcqCostBoxNo + 1 - ADJBD), 2)
        flLeft = imAcqCostX + imAcqCostBoxNo * (edcBarter(3).Width + 15)
        flTop = imAcqCostY
    Else
        edcBarter(3).Text = gLongToStrDec(tmVbf(imVBFIndex).lActAcqCost(imAcqCostBoxNo - (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD) + 1 - ADJBD), 2)
        'flLeft = imAcqCostX + (imAcqCostBoxNo - (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD)-ADJBD) * (edcBarter(3).Width + 15)
        flLeft = imAcqCostX + (imAcqCostBoxNo - (UBound(tmVbf(imVBFIndex).lDefAcqCost) + ADJBD)) * (edcBarter(3).Width + 15)
        flTop = imAcqCostY + edcBarter(3).Height
    End If
    edcBarter(3).Move flLeft, flTop
    edcBarter(3).SelStart = 0
    edcBarter(3).SelLength = Len(edcBarter(3).Text)
    edcBarter(3).Visible = True
    edcBarter(3).SetFocus
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAcqCostPaint                   *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint sales spot length table  *
'*                                                     *
'*******************************************************
Private Sub mAcqCostPaint()
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim flX As Single
    Dim flY As Single
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    If imVBFIndex < 0 Then
        pbcAcq(0).Cls
        Exit Sub
    End If
    llColor = pbcAcq(0).ForeColor
    slFontName = pbcAcq(0).FontName
    flFontSize = pbcAcq(0).FontSize
    pbcAcq(0).ForeColor = BLUE
    pbcAcq(0).FontBold = False
    pbcAcq(0).FontSize = 7
    pbcAcq(0).FontName = "Arial"
    pbcAcq(0).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    flX = imAcqCostX + fgBoxInsetX
    flY = 15
    For ilCol = 1 To 10 Step 1
        gPaintArea pbcAcq(0), flX, flY + 15, imAcqCostW - fgBoxInsetX - 15, imAcqCostH - 45, WHITE
        pbcAcq(0).CurrentX = flX
        pbcAcq(0).CurrentY = flY '- 15
        pbcAcq(0).Print Trim$(Str$(tmVbf(imVBFIndex).iSpotLen(ilCol - ADJBD)))
        flX = flX + imAcqCostW + 15
    Next ilCol
    pbcAcq(0).FontSize = flFontSize
    pbcAcq(0).FontName = slFontName
    pbcAcq(0).FontSize = flFontSize
    pbcAcq(0).ForeColor = llColor
    pbcAcq(0).FontBold = True
    flX = imAcqCostX + fgBoxInsetX
    flY = imAcqCostY
    For ilRow = 0 To 1 Step 1
        For ilCol = 1 To 10 Step 1
            gPaintArea pbcAcq(0), flX, flY + 15, imAcqCostW - fgBoxInsetX - 15, imAcqCostH - 45, WHITE
            pbcAcq(0).CurrentX = flX
            pbcAcq(0).CurrentY = flY - 15
            If ilRow = 0 Then
                pbcAcq(0).Print gLongToStrDec(tmVbf(imVBFIndex).lDefAcqCost(ilCol - ADJBD), 2)
            Else
                pbcAcq(0).Print gLongToStrDec(tmVbf(imVBFIndex).lActAcqCost(ilCol - ADJBD), 2)
            End If
            flX = flX + imAcqCostW + 15
        Next ilCol
        flX = imAcqCostX + fgBoxInsetX
        flY = flY + imAcqCostH - 15
    Next ilRow
End Sub

Private Sub mSetAcqLen()
    Dim ilVpf As Integer
    Dim ilLen As Integer
    Dim ilFound As Integer
    Dim ilNext As Integer

    If imVBFIndex <= -1 Then
        Exit Sub
    End If
    For ilVpf = LBound(tmVpf.iSLen) To UBound(tmVpf.iSLen) Step 1
        ilFound = False
        For ilLen = LBound(tmVbf(imVBFIndex).iSpotLen) To UBound(tmVbf(imVBFIndex).iSpotLen) Step 1
            If tmVpf.iSLen(ilVpf) = tmVbf(imVBFIndex).iSpotLen(ilLen) Then
                ilFound = True
                Exit For
            End If
        Next ilLen
        If Not ilFound Then
            For ilLen = LBound(tmVbf(imVBFIndex).iSpotLen) To UBound(tmVbf(imVBFIndex).iSpotLen) Step 1
                If tmVbf(imVBFIndex).iSpotLen(ilLen) <= 0 Then
                    tmVbf(imVBFIndex).iSpotLen(ilLen) = tmVpf.iSLen(ilVpf)
                    imVbfChg = True
                    Exit For
                End If
            Next ilLen
        End If
    Next ilVpf
    ilLen = LBound(tmVbf(imVBFIndex).iSpotLen)
    Do
        If tmVbf(imVBFIndex).iSpotLen(ilLen) > 0 Then
            ilFound = False
            For ilVpf = LBound(tmVpf.iSLen) To UBound(tmVpf.iSLen) Step 1
                If tmVpf.iSLen(ilVpf) = tmVbf(imVBFIndex).iSpotLen(ilLen) Then
                    ilFound = True
                    Exit For
                End If
            Next ilVpf
            If Not ilFound Then
                For ilNext = ilLen + 1 To UBound(tmVbf(imVBFIndex).iSpotLen) Step 1
                    tmVbf(imVBFIndex).iSpotLen(ilNext - 1) = tmVbf(imVBFIndex).iSpotLen(ilNext)
                    tmVbf(imVBFIndex).lDefAcqCost(ilNext - 1) = tmVbf(imVBFIndex).lDefAcqCost(ilNext)
                    tmVbf(imVBFIndex).lActAcqCost(ilNext - 1) = tmVbf(imVBFIndex).lActAcqCost(ilNext)
                Next ilNext
                If ilLen + 1 <= UBound(tmVbf(imVBFIndex).iSpotLen) Then
                    tmVbf(imVBFIndex).iSpotLen(UBound(tmVbf(imVBFIndex).lDefAcqCost)) = 0
                    tmVbf(imVBFIndex).lDefAcqCost(UBound(tmVbf(imVBFIndex).lDefAcqCost)) = 0
                    tmVbf(imVBFIndex).lActAcqCost(UBound(tmVbf(imVBFIndex).lDefAcqCost)) = 0
                End If
            Else
                ilLen = ilLen + 1
            End If
        Else
            ilLen = ilLen + 1
        End If
    Loop While ilLen <= UBound(tmVbf(imVBFIndex).iSpotLen)
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetAcqIndex                    *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set spot length box            *
'*                                                     *
'*******************************************************
Private Sub mSetAcqIndex()

    Dim flLeft As Single
    Dim flTop As Single

    If (imAcqIndexBoxNo < 0) Or (imAcqIndexBoxNo > (UBound(tmVbfIndex.lDefAcqCost) + ADJBD) - 1) Then
        Exit Sub
    End If
    If tmVbfIndex.iSpotLen(imAcqIndexBoxNo + 1 - ADJBD) = 0 Then
        Exit Sub
    End If
    edcBarter(4).Text = gLongToStrDec(tmVbfIndex.lDefAcqCost(imAcqIndexBoxNo + 1 - ADJBD), 2)
    flLeft = imAcqIndexX + imAcqIndexBoxNo * (edcBarter(4).Width + 15)
    flTop = imAcqIndexY
    edcBarter(4).Move flLeft, flTop
    edcBarter(4).SelStart = 0
    edcBarter(4).SelLength = Len(edcBarter(4).Text)
    edcBarter(4).Visible = True
    edcBarter(4).SetFocus
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAcqIndexPaint                   *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint sales spot length table  *
'*                                                     *
'*******************************************************
Private Sub mAcqIndexPaint()
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim flX As Single
    Dim flY As Single
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    'If imVBFIndex < 0 Then
    '    pbcAcq(1).Cls
    '    Exit Sub
    'End If
    llColor = pbcAcq(1).ForeColor
    slFontName = pbcAcq(1).FontName
    flFontSize = pbcAcq(1).FontSize
    pbcAcq(1).ForeColor = BLUE
    pbcAcq(1).FontBold = False
    pbcAcq(1).FontSize = 7
    pbcAcq(1).FontName = "Arial"
    pbcAcq(1).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    flX = imAcqIndexX + fgBoxInsetX
    flY = 15
    For ilCol = 1 To 10 Step 1
        gPaintArea pbcAcq(1), flX, flY + 15, imAcqIndexW - fgBoxInsetX - 15, imAcqIndexH - 45, WHITE
        If tmVbfIndex.iSpotLen(ilCol - ADJBD) > 0 Then
            pbcAcq(1).CurrentX = flX
            pbcAcq(1).CurrentY = flY '- 15
            pbcAcq(1).Print Trim$(Str$(tmVbfIndex.iSpotLen(ilCol - ADJBD)))
            flX = flX + imAcqIndexW + 15
        End If
    Next ilCol
    pbcAcq(1).FontSize = flFontSize
    pbcAcq(1).FontName = slFontName
    pbcAcq(1).FontSize = flFontSize
    pbcAcq(1).ForeColor = llColor
    pbcAcq(1).FontBold = True
    flX = imAcqIndexX + fgBoxInsetX
    flY = imAcqIndexY
    For ilCol = 1 To 10 Step 1
        If tmVbfIndex.iSpotLen(ilCol - ADJBD) > 0 Then
            gPaintArea pbcAcq(1), flX, flY + 15, imAcqIndexW - fgBoxInsetX - 15, imAcqIndexH - 45, WHITE
            pbcAcq(1).CurrentX = flX
            pbcAcq(1).CurrentY = flY - 15
            pbcAcq(1).Print gLongToStrDec(tmVbfIndex.lDefAcqCost(ilCol - ADJBD), 2)
            flX = flX + imAcqIndexW + 15
        End If
    Next ilCol
End Sub

Private Sub mSetAcqIndexLen()
    Dim ilVpf As Integer
    Dim ilLen As Integer
    Dim ilFound As Integer
    Dim ilNext As Integer

    For ilVpf = LBound(tmVpf.iSLen) To UBound(tmVpf.iSLen) Step 1
        ilFound = False
        For ilLen = LBound(tmVbfIndex.iSpotLen) To UBound(tmVbfIndex.iSpotLen) Step 1
            If tmVpf.iSLen(ilVpf) = tmVbfIndex.iSpotLen(ilLen) Then
                ilFound = True
                Exit For
            End If
        Next ilLen
        If Not ilFound Then
            For ilLen = LBound(tmVbfIndex.iSpotLen) To UBound(tmVbfIndex.iSpotLen) Step 1
                If tmVbfIndex.iSpotLen(ilLen) <= 0 Then
                    tmVbfIndex.iSpotLen(ilLen) = tmVpf.iSLen(ilVpf)
                    imVbfChg = True
                    Exit For
                End If
            Next ilLen
        End If
    Next ilVpf
    ilLen = LBound(tmVbfIndex.iSpotLen)
    Do
        If tmVbfIndex.iSpotLen(ilLen) > 0 Then
            ilFound = False
            For ilVpf = LBound(tmVpf.iSLen) To UBound(tmVpf.iSLen) Step 1
                If tmVpf.iSLen(ilVpf) = tmVbfIndex.iSpotLen(ilLen) Then
                    ilFound = True
                    Exit For
                End If
            Next ilVpf
            If Not ilFound Then
                For ilNext = ilLen + 1 To UBound(tmVbfIndex.iSpotLen) Step 1
                    tmVbfIndex.iSpotLen(ilNext - 1) = tmVbfIndex.iSpotLen(ilNext)
                    tmVbfIndex.lDefAcqCost(ilNext - 1) = tmVbfIndex.lDefAcqCost(ilNext)
                    tmVbfIndex.lActAcqCost(ilNext - 1) = tmVbfIndex.lActAcqCost(ilNext)
                Next ilNext
                If ilLen + 1 <= UBound(tmVbfIndex.iSpotLen) Then
                    tmVbfIndex.iSpotLen(UBound(tmVbfIndex.lDefAcqCost)) = 0
                    tmVbfIndex.lDefAcqCost(UBound(tmVbfIndex.lDefAcqCost)) = 0
                    tmVbfIndex.lActAcqCost(UBound(tmVbfIndex.lDefAcqCost)) = 0
                End If
            Else
                ilLen = ilLen + 1
            End If
        Else
            ilLen = ilLen + 1
        End If
    Loop While ilLen <= UBound(tmVbfIndex.iSpotLen)
End Sub

Private Sub mGridParticipantLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdParticipant.Rows - 1 Step 1
        grdParticipant.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdParticipant.Cols - 1 Step 1
        grdParticipant.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridParticipantColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdParticipant.Row = grdParticipant.FixedRows - 1
    grdParticipant.Col = SSOURCEINDEX
    grdParticipant.CellFontBold = False
    grdParticipant.CellFontName = "Arial"
    grdParticipant.CellFontSize = 6.75
    grdParticipant.CellForeColor = vbBlue
    'grdParticipant.CellBackColor = LIGHTBLUE
    grdParticipant.TextMatrix(grdParticipant.Row, grdParticipant.Col) = "Sales Source"
    grdParticipant.Col = PARTINDEX
    grdParticipant.CellFontBold = False
    grdParticipant.CellFontName = "Arial"
    grdParticipant.CellFontSize = 6.75
    grdParticipant.CellForeColor = vbBlue
    'grdParticipant.CellBackColor = LIGHTBLUE
    grdParticipant.TextMatrix(grdParticipant.Row, grdParticipant.Col) = "Participant"
    grdParticipant.Col = INTUPDATEINDEX
    grdParticipant.CellFontBold = False
    grdParticipant.CellFontName = "Arial"
    grdParticipant.CellFontSize = 6.75
    grdParticipant.CellForeColor = vbBlue
    'grdParticipant.CellBackColor = LIGHTBLUE
    grdParticipant.TextMatrix(grdParticipant.Row, grdParticipant.Col) = "Internal Inv Update"
    grdParticipant.Col = EXTUPDATEINDEX
    grdParticipant.CellFontBold = False
    grdParticipant.CellFontName = "Arial"
    grdParticipant.CellFontSize = 6.75
    grdParticipant.CellForeColor = vbBlue
    'grdParticipant.CellBackColor = LIGHTBLUE
    grdParticipant.TextMatrix(grdParticipant.Row, grdParticipant.Col) = "External Inv Update"
    grdParticipant.Col = PRODPCTINDEX
    grdParticipant.CellFontBold = False
    grdParticipant.CellFontName = "Arial"
    grdParticipant.CellFontSize = 6.75
    grdParticipant.CellForeColor = vbBlue
    'grdParticipant.CellBackColor = LIGHTBLUE
    grdParticipant.TextMatrix(grdParticipant.Row, grdParticipant.Col) = "Revenue %"
    grdParticipant.Col = PIFCODEINDEX
    grdParticipant.CellFontBold = False
    grdParticipant.CellFontName = "Arial"
    grdParticipant.CellFontSize = 6.75
    grdParticipant.CellForeColor = vbBlue
    'grdParticipant.CellBackColor = LIGHTBLUE
    grdParticipant.TextMatrix(grdParticipant.Row, grdParticipant.Col) = "PDFCode"
End Sub

Private Sub mGridParticipantColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdParticipant.ColWidth(PIFCODEINDEX) = 0
    grdParticipant.ColWidth(SSOURCEINDEX) = 0.27 * grdParticipant.Width
    grdParticipant.ColWidth(PARTINDEX) = 0.27 * grdParticipant.Width
    grdParticipant.ColWidth(INTUPDATEINDEX) = 0.16 * grdParticipant.Width
    grdParticipant.ColWidth(EXTUPDATEINDEX) = 0.16 * grdParticipant.Width
    grdParticipant.ColWidth(PRODPCTINDEX) = 0.1 * grdParticipant.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdParticipant.Width
    For ilCol = 0 To grdParticipant.Cols - 1 Step 1
        llWidth = llWidth + grdParticipant.ColWidth(ilCol)
        If (grdParticipant.ColWidth(ilCol) > 15) And (grdParticipant.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdParticipant.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdParticipant.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdParticipant.Width
            For ilCol = 0 To grdParticipant.Cols - 1 Step 1
                If (grdParticipant.ColWidth(ilCol) > 15) And (grdParticipant.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdParticipant.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdParticipant.FixedCols To grdParticipant.Cols - 1 Step 1
                If grdParticipant.ColWidth(ilCol) > 15 Then
                    ilColInc = grdParticipant.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdParticipant.ColWidth(ilCol) = grdParticipant.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub


Private Sub mPopParticipantDates()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ll12312069 As Long
    ReDim tlPif(0 To 0) As PIF

    imPifChg = False
    cbcParticipant.Clear
    mClearParticipantCtrl True
    ll12312069 = gDateValue("12/31/2069")
    ilRet = gObtainPIF_ForVef(hmPif, tmVef.iCode, tlPif())
    ReDim tmPifRec(0 To UBound(tlPif)) As PIFREC
    For ilLoop = 0 To UBound(tlPif) - 1 Step 1
        tmPifRec(ilLoop).tPif = tlPif(ilLoop)
        tmPifRec(ilLoop).iStatus = 1
        If tmPifRec(ilLoop).tPif.iSeqNo = 1 Then
            gUnpackDate tmPifRec(ilLoop).tPif.iStartDate(0), tmPifRec(ilLoop).tPif.iStartDate(1), slStartDate
            gUnpackDate tmPifRec(ilLoop).tPif.iEndDate(0), tmPifRec(ilLoop).tPif.iEndDate(1), slEndDate
            If gDateValue(slEndDate) = ll12312069 Then
                If igVehNewToVehOpt Then
                    slEndDate = "New"
                Else
                    slEndDate = "TFN"
                End If
            End If
            cbcParticipant.AddItem slStartDate & "-" & Trim$(slEndDate)
        End If
    Next ilLoop
    cbcParticipant.AddItem "[New]", 0
    For ilLoop = 0 To cbcParticipant.ListCount - 1 Step 1
        If InStr(1, cbcParticipant.List(ilLoop), "TFN", vbTextCompare) > 0 Then
            cbcParticipant.ListIndex = ilLoop
            Exit For
        End If
    Next ilLoop
End Sub

Private Sub mClearParticipantCtrl(ilClearDate As Integer)
    Dim ilRow As Integer
    Dim ilCol As Integer

    If ilClearDate Then
        'edcParticipantDate.Text = ""
        'edcParticipantDate.Enabled = True
        'edcParticipantDate.BackColor = CYAN
        csiParticipantDate.Text = ""
        csiParticipantDate.SetEnabled True
        csiParticipantDate.BackColor = CYAN
    End If
    grdParticipant.Redraw = False
    For ilRow = grdParticipant.FixedRows To grdParticipant.Rows - 1 Step 1
        grdParticipant.Row = ilRow
        For ilCol = grdParticipant.FixedCols To grdParticipant.Cols - 1 Step 1
            grdParticipant.Col = ilCol
            grdParticipant.CellBackColor = vbWhite
            grdParticipant.CellForeColor = vbBlack
            If ilCol <> PIFCODEINDEX Then
                grdParticipant.TextMatrix(ilRow, ilCol) = ""
            Else
                grdParticipant.TextMatrix(ilRow, ilCol) = "0"
            End If
        Next ilCol
    Next ilRow
    grdParticipant.Redraw = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSSourceBranch                  *
'*                                                     *
'*             Created:5/8/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to sales  *
'*                      source and process             *
'*                      communication back from sales  *
'*                      source                         *
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
Private Function mSSourceBranch() As Integer
'
'   ilRet = mSSourceBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcSSDropDown, lbcSSource, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mSSourceBranch = False
        Exit Function
    End If
    sgMnfCallType = "S"
    igMNmCallSource = CALLSOURCEVEHICLE
    If lbcSSource.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Vehicle^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Vehicle^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Vehicle.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Vehicle.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mSSourceBranch = True
    imUpdateAllowed = ilUpdateAllowed
'    gShowBranner
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcSSource.Clear
        smSSourceCodeTag = ""
        mSSourcePop
        If imTerminate Then
            mSSourceBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcSSource
        sgMNmName = ""
        If gLastFound(lbcSSource) > 0 Then
            imSSChgMode = True
            lbcSSource.ListIndex = gLastFound(lbcSSource)
            edcSSDropDown.Text = lbcSSource.List(lbcSSource.ListIndex)
            imSSChgMode = False
            mSSourceBranch = False
            'mSetChg imBoxNo
        Else
            imSSChgMode = True
            lbcSSource.ListIndex = 0
            edcSSDropDown.Text = lbcSSource.List(0)
            imSSChgMode = False
            'mSetChg imBoxNo
            edcSSDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mPartEnableBox
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mPartEnableBox
        Exit Function
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSSourcePop                     *
'*                                                     *
'*             Created:5/8/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Sales source list box *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mSSourcePop()
'
'   mSSourcePop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffset(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilSS As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    ilIndex = lbcSSource.ListIndex
    If ilIndex > 0 Then
        slName = lbcSSource.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "S"
    ilOffset(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(VehOpt, lbcSSource, lbcSSourceCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(VehOpt, lbcSSource, tmSSourceCode(), smSSourceCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffset())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSSourcePopErr
        gCPErrorMsg ilRet, "mSSourcePop (gIMoveListBox)", VehOpt
        On Error GoTo 0
        lbcSSource.AddItem "[None]", 0  'Force as first item on list
        lbcSSource.AddItem "[New]", 0  'Force as first item on list
        imSSChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcSSource
            If gLastFound(lbcSSource) > 0 Then
                lbcSSource.ListIndex = gLastFound(lbcSSource)
            Else
                lbcSSource.ListIndex = -1
            End If
        Else
            lbcSSource.ListIndex = ilIndex
        End If
        ReDim smSUpdateRvf(LBound(tmSSourceCode) To UBound(tmSSourceCode)) As String * 1
        For ilLoop = LBound(tmSSourceCode) To UBound(tmSSourceCode) - 1 Step 1
            slNameCode = tmSSourceCode(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilCode = Val(slCode)
            smSUpdateRvf(ilLoop) = "Y"
            For ilSS = LBound(tmSMnf) To UBound(tmSMnf) - 1 Step 1
                If tmSMnf(ilSS).iCode = ilCode Then
                    smSUpdateRvf(ilLoop) = tmSMnf(ilSS).sUnitType
                    Exit For
                End If
            Next ilSS
        Next ilLoop
        imSSChgMode = False
    End If
    Exit Sub
mSSourcePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGpBranch                    *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Vehicle group and process      *
'*                      communication back from        *
'*                      invoice sorting                *
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
Private Function mVehGpBranch() As Integer
'
'   ilRet = mVehGpBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcVehGpDropDown, lbcVehGp, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcVehGpDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mVehGpBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mVehGpBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "H"
    igMNmCallSource = CALLSOURCEVEHICLE
    If edcVehGpDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Vehicle^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\1"
        Else
            slStr = "Vehicle^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\1"
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\1"
    '    Else
    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\1"
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Vehicle.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Vehicle.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mVehGpBranch = True
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
        lbcVehGp.Clear
        smVehGpCodeTag = ""
        mVehGpPop
        If imTerminate Then
            mVehGpBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcVehGp
        sgMNmName = ""
        If gLastFound(lbcVehGp) > 0 Then
            imVehGpChgMode = True
            lbcVehGp.ListIndex = gLastFound(lbcVehGp)
            edcVehGpDropDown.Text = lbcVehGp.List(lbcVehGp.ListIndex)
            imVehGpChgMode = False
            mVehGpBranch = False
            'mSetChg imBoxNo
        Else
            imVehGpChgMode = True
            lbcVehGp.ListIndex = 1
            edcVehGpDropDown.Text = lbcVehGp.List(1)
            imVehGpChgMode = False
            'mSetChg imBoxNo
            edcVehGpDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mPartEnableBox
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mPartEnableBox
        Exit Function
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGpPop                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Vehicle Group Code    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mVehGpPop()
'
'   mVehGpPop
'   Where:
'
    'ReDim ilFilter(0) As Integer
    'ReDim slFilter(0) As String
    'ReDim ilOffset(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcVehGp.ListIndex
    If ilIndex > 0 Then
        slName = lbcVehGp.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilFilter(0) = CHARFILTER
    'slFilter(0) = "H"
    'ilOffset(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(VehOpt, lbcVehGp, lbcVehGpCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    'ilRet = gIMoveListBox(VehOpt, lbcVehGp, tmVehGpCode(), smVehGpCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gPopMnfPlusFieldsBox(VehOpt, lbcVehGp, tmVehGpCode(), smVehGpCodeTag, "H1")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehGpPopErr
        gCPErrorMsg ilRet, "mVehGpPop (gPopMnfPlusFieldsBox)", VehOpt
        On Error GoTo 0
        lbcVehGp.AddItem "[New]", 0  'Force as first item on list
        imVehGpChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcVehGp
            If gLastFound(lbcVehGp) >= 1 Then
                lbcVehGp.ListIndex = gLastFound(lbcVehGp)
            Else
                lbcVehGp.ListIndex = -1
            End If
        Else
            lbcVehGp.ListIndex = ilIndex
        End If
        imVehGpChgMode = False
    End If
    Exit Sub
mVehGpPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPartEnableBox                  *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mPartEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    Dim llRow As Long

    If (grdParticipant.Row < grdParticipant.FixedRows) Or (grdParticipant.Row >= grdParticipant.Rows) Or (grdParticipant.Col < grdParticipant.FixedCols) Or (grdParticipant.Col >= grdParticipant.Cols - 1) Then
        Exit Sub
    End If
    lmPartEnableRow = grdParticipant.Row
    lmPartEnableCol = grdParticipant.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdParticipant.Left - pbcArrow.Width - 30, grdParticipant.Top + grdParticipant.RowPos(grdParticipant.Row) + (grdParticipant.RowHeight(grdParticipant.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    imCtrlVisible = True
    slStr = grdParticipant.Text
    If slStr = "Missing" Then
        grdParticipant.Text = ""
        grdParticipant.CellForeColor = vbBlack
    End If
    Select Case grdParticipant.Col
        Case SSOURCEINDEX
            mSSourcePop
            If imTerminate Then
                Exit Sub
            End If
            lbcSSource.Height = gListBoxHeight(lbcSSource.ListCount, 6)
            edcSSDropDown.MaxLength = 20
            imSSChgMode = True
            slStr = grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol)
            gFindMatch slStr, 1, lbcSSource
            If gLastFound(lbcSSource) >= 1 Then
                lbcSSource.ListIndex = gLastFound(lbcSSource)
                edcSSDropDown.Text = lbcSSource.List(lbcSSource.ListIndex)
            Else
                lbcSSource.ListIndex = 1
                edcSSDropDown.Text = lbcSSource.List(1)
            End If
            imSSChgMode = False
        Case PARTINDEX
            mVehGpPop
            If imTerminate Then
                Exit Sub
            End If
            lbcVehGp.Height = gListBoxHeight(lbcVehGp.ListCount, 6)
            'TTP 10466 - Vehicle Group Names (mnfName): need to expand character limit to 50
            edcVehGpDropDown.MaxLength = 50
            imVehGpChgMode = True
            slStr = grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol)
            gFindMatch slStr, 1, lbcVehGp
            If gLastFound(lbcVehGp) >= 1 Then
                lbcVehGp.ListIndex = gLastFound(lbcVehGp)
                edcVehGpDropDown.Text = lbcVehGp.List(lbcVehGp.ListIndex)
            Else
                If lbcVehGp.ListCount > 1 Then
                    lbcVehGp.ListIndex = 1
                    edcVehGpDropDown.Text = lbcVehGp.List(1)
                Else
                    lbcVehGp.ListIndex = 0
                    edcVehGpDropDown.Text = lbcVehGp.List(0)
                End If
            End If
            'TTP 10466 - Vehicle Group Names (mnfName): need to expand character limit to 50
            lbcVehGp.Width = 4400
            imVehGpChgMode = False
        Case INTUPDATEINDEX
            If Trim$(grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol)) = "" Then
                grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol) = "Receivables"
            End If
        Case EXTUPDATEINDEX
            If Trim$(grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol)) = "" Then
                grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol) = "Receivables"
            End If
        Case PRODPCTINDEX
            edcProdPct.MaxLength = 6
            slStr = grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol)
            If slStr <> "" Then
                edcProdPct.Text = slStr
            Else
                slStr = "100.00"
                For llRow = grdParticipant.FixedRows To grdParticipant.Rows - 1 Step 1
                    If grdParticipant.TextMatrix(llRow, SSOURCEINDEX) <> "" Then
                        If llRow <> lmPartEnableRow Then
                            If grdParticipant.TextMatrix(llRow, SSOURCEINDEX) = grdParticipant.TextMatrix(lmPartEnableRow, SSOURCEINDEX) Then
                                slStr = gSubStr(slStr, grdParticipant.TextMatrix(llRow, PRODPCTINDEX))
                            End If
                        End If
                    End If
                Next llRow
                edcProdPct.Text = slStr
            End If
    End Select
    mPartSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mPartSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer

    If (grdParticipant.Row < grdParticipant.FixedRows) Or (grdParticipant.Row >= grdParticipant.Rows) Or (grdParticipant.Col < grdParticipant.FixedCols) Or (grdParticipant.Col >= grdParticipant.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdParticipant.Left - pbcArrow.Width - 30, grdParticipant.Top + grdParticipant.RowPos(grdParticipant.Row) + (grdParticipant.RowHeight(grdParticipant.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    llColPos = 0
    For ilCol = 0 To grdParticipant.Col - 1 Step 1
        llColPos = llColPos + grdParticipant.ColWidth(ilCol)
    Next ilCol
    Select Case grdParticipant.Col
        Case SSOURCEINDEX
            edcSSDropDown.Move grdParticipant.Left + llColPos + 30, grdParticipant.Top + grdParticipant.RowPos(grdParticipant.Row) + 30, grdParticipant.ColWidth(grdParticipant.Col) - 30, grdParticipant.RowHeight(grdParticipant.Row) - 15
            cmcDropDown.Move edcSSDropDown.Left + edcSSDropDown.Width, edcSSDropDown.Top
            lbcSSource.Move edcSSDropDown.Left, edcSSDropDown.Top - lbcSSource.Height '+ edcVehGpDropDown.Height
            lbcSSource.ZOrder vbBringToFront
            edcSSDropDown.Visible = True
            cmcDropDown.Visible = True
            edcSSDropDown.SetFocus
        Case PARTINDEX
            edcVehGpDropDown.Move grdParticipant.Left + llColPos + 30, grdParticipant.Top + grdParticipant.RowPos(grdParticipant.Row) + 30, grdParticipant.ColWidth(grdParticipant.Col) - 30, grdParticipant.RowHeight(grdParticipant.Row) - 15
            cmcDropDown.Move edcVehGpDropDown.Left + edcVehGpDropDown.Width, edcVehGpDropDown.Top
            lbcVehGp.Move edcVehGpDropDown.Left, edcVehGpDropDown.Top - lbcVehGp.Height '+ edcVehGpDropDown.Height
            lbcVehGp.ZOrder vbBringToFront
            edcVehGpDropDown.Visible = True
            cmcDropDown.Visible = True
            edcVehGpDropDown.SetFocus
        Case INTUPDATEINDEX
            pbcIntUpdateRvf.Move grdParticipant.Left + llColPos + 30, grdParticipant.Top + grdParticipant.RowPos(grdParticipant.Row) + 30, grdParticipant.ColWidth(grdParticipant.Col) - 30, grdParticipant.RowHeight(grdParticipant.Row) - 15
            pbcIntUpdateRvf.Visible = True
            pbcIntUpdateRvf.SetFocus
        Case EXTUPDATEINDEX
            pbcExtUpdateRvf.Move grdParticipant.Left + llColPos + 30, grdParticipant.Top + grdParticipant.RowPos(grdParticipant.Row) + 30, grdParticipant.ColWidth(grdParticipant.Col) - 30, grdParticipant.RowHeight(grdParticipant.Row) - 15
            pbcExtUpdateRvf.Visible = True
            pbcExtUpdateRvf.SetFocus
        Case PRODPCTINDEX
            edcProdPct.Move grdParticipant.Left + llColPos + 30, grdParticipant.Top + grdParticipant.RowPos(grdParticipant.Row) + 30, grdParticipant.ColWidth(grdParticipant.Col) - 30, grdParticipant.RowHeight(grdParticipant.Row) - 15
            edcProdPct.Visible = True  'Set visibility
            edcProdPct.SetFocus
    End Select
End Sub

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
Private Sub mPartSetShow()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPos                                                                                 *
'******************************************************************************************

'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim ilSSourceLastFd As Integer

    pbcArrow.Visible = False
    If (lmPartEnableRow >= grdParticipant.FixedRows) And (lmPartEnableRow < grdParticipant.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmPartEnableCol
            Case SSOURCEINDEX
                edcSSDropDown.Visible = False  'Set visibility
                lbcSSource.Visible = False
                slStr = edcSSDropDown.Text
                If StrComp(grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol), slStr, vbTextCompare) <> 0 Then
                    imPifChg = True
                    grdParticipant.TextMatrix(lmPartEnableRow, INTUPDATEINDEX) = ""
                    grdParticipant.TextMatrix(lmPartEnableRow, EXTUPDATEINDEX) = ""
                    grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol) = slStr
                    gFindMatch slStr, 2, lbcSSource
                    If gLastFound(lbcSSource) >= 2 Then
                        ilSSourceLastFd = gLastFound(lbcSSource)
                        If smSUpdateRvf(ilSSourceLastFd - 2) <> "A" Then
                            slStr = smSUpdateRvf(ilSSourceLastFd - 2)
                            Select Case slStr
                                Case "Y"
                                    grdParticipant.TextMatrix(lmPartEnableRow, INTUPDATEINDEX) = "Receivables"
                                Case "N"
                                    grdParticipant.TextMatrix(lmPartEnableRow, INTUPDATEINDEX) = "History"
                                Case "E"
                                    grdParticipant.TextMatrix(lmPartEnableRow, INTUPDATEINDEX) = "Export+History"
                                Case "F"
                                    grdParticipant.TextMatrix(lmPartEnableRow, INTUPDATEINDEX) = "Export+A/R"
                                Case Else
                                    grdParticipant.TextMatrix(lmPartEnableRow, INTUPDATEINDEX) = ""
                            End Select
                        End If
                    End If
                End If

             Case PARTINDEX
                edcVehGpDropDown.Visible = False  'Set visibility
                lbcVehGp.Visible = False
                slStr = edcVehGpDropDown.Text
                If StrComp(grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol), slStr, vbTextCompare) <> 0 Then
                    imPifChg = True
                End If
                grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol) = slStr
            Case INTUPDATEINDEX
                pbcIntUpdateRvf.Visible = False  'Set visibility
            Case EXTUPDATEINDEX
                pbcExtUpdateRvf.Visible = False  'Set visibility
             Case PRODPCTINDEX
                edcProdPct.Visible = False  'Set visibility
                slStr = edcProdPct.Text
                If gStrDecToLong(grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol), 2) <> gStrDecToLong(slStr, 2) Then
                    imPifChg = True
                End If
                grdParticipant.TextMatrix(lmPartEnableRow, lmPartEnableCol) = slStr
        End Select
    End If
    pbcArrow.Visible = False
    cmcDropDown.Visible = False
    lmPartEnableRow = -1
    lmPartEnableCol = -1
    imCtrlVisible = False
    mSetCommands
End Sub

Private Sub pbcIntUpdateRvf_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcIntUpdateRvf_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("R") Or (KeyAscii = Asc("r")) Then
        If grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) <> "Receivables" Then
            imPifChg = True
        End If
        grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Receivables"
        pbcIntUpdateRvf_Paint
    ElseIf KeyAscii = Asc("H") Or (KeyAscii = Asc("h")) Then
        If grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) <> "History" Then
            imPifChg = True
        End If
        grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "History"
        pbcIntUpdateRvf_Paint
    ElseIf KeyAscii = Asc("E") Or (KeyAscii = Asc("e")) Then
        If grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+History" Then
            imPifChg = True
            grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+A/R"
            pbcIntUpdateRvf_Paint
        ElseIf grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+A/R" Then
            imPifChg = True
            grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+History"
            pbcIntUpdateRvf_Paint
        Else
            If grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) <> "Export+History" Then
                imPifChg = True
            End If
            grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+History"
            pbcIntUpdateRvf_Paint
        End If
    End If
    If KeyAscii = Asc(" ") Then
        If grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Receivables" Then
            imPifChg = True
            grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "History"
            pbcIntUpdateRvf_Paint
        ElseIf grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "History" Then
            imPifChg = True
            grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+History"
            pbcIntUpdateRvf_Paint
        ElseIf grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+History" Then
            imPifChg = True
            grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+A/R"
            pbcIntUpdateRvf_Paint
        ElseIf grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+A/R" Then
            imPifChg = True
            grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Receivables"
            pbcIntUpdateRvf_Paint
        Else
            imPifChg = True
            grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Receivables"
            pbcIntUpdateRvf_Paint
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcIntUpdateRvf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Receivables" Then
        imPifChg = True
        grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "History"
        pbcIntUpdateRvf_Paint
    ElseIf grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "History" Then
        imPifChg = True
        grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+History"
        pbcIntUpdateRvf_Paint
    ElseIf grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+History" Then
        imPifChg = True
        grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+A/R"
        pbcIntUpdateRvf_Paint
    ElseIf grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Export+A/R" Then
        imPifChg = True
        grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX) = "Receivables"
        pbcIntUpdateRvf_Paint
    End If
    mSetCommands
End Sub

Private Sub pbcIntUpdateRvf_Paint()
    pbcIntUpdateRvf.Cls
    pbcIntUpdateRvf.CurrentX = fgBoxInsetX
    pbcIntUpdateRvf.CurrentY = 0 'fgBoxInsetY
    pbcIntUpdateRvf.Print grdParticipant.TextMatrix(grdParticipant.Row, INTUPDATEINDEX)
End Sub

Private Sub pbcExtUpdateRvf_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcExtUpdateRvf_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("R") Or (KeyAscii = Asc("r")) Then
        If grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) <> "Receivables" Then
            imPifChg = True
        End If
        grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) = "Receivables"
        pbcExtUpdateRvf_Paint
    ElseIf KeyAscii = Asc("H") Or (KeyAscii = Asc("h")) Then
        If grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) <> "History" Then
            imPifChg = True
        End If
        grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) = "History"
        pbcExtUpdateRvf_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) = "Receivables" Then
            imPifChg = True
            grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) = "History"
            pbcExtUpdateRvf_Paint
        ElseIf grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) = "History" Then
            imPifChg = True
            grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) = "Receivables"
            pbcExtUpdateRvf_Paint
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcExtUpdateRvf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) = "Receivables" Then
        imPifChg = True
        grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) = "History"
        pbcExtUpdateRvf_Paint
    ElseIf grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) = "History" Then
        imPifChg = True
        grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX) = "Receivables"
        pbcExtUpdateRvf_Paint
    End If
    mSetCommands
End Sub

Private Sub pbcExtUpdateRvf_Paint()
    pbcExtUpdateRvf.Cls
    pbcExtUpdateRvf.CurrentX = fgBoxInsetX
    pbcExtUpdateRvf.CurrentY = 0 'fgBoxInsetY
    pbcExtUpdateRvf.Print grdParticipant.TextMatrix(grdParticipant.Row, EXTUPDATEINDEX)
End Sub

Private Function mPartColOk(ilRow As Integer, ilCol As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************

    Dim slStr As String

    mPartColOk = True
    If grdParticipant.CellBackColor = LIGHTYELLOW Then
        mPartColOk = False
        Exit Function
    End If
    If (ilCol >= PARTINDEX) And (grdParticipant.TextMatrix(ilRow, SSOURCEINDEX) = "") Then
        mPartColOk = False
        Exit Function
    End If
    If (ilCol = INTUPDATEINDEX) Then
        slStr = grdParticipant.TextMatrix(ilRow, SSOURCEINDEX)
        gFindMatch slStr, 2, lbcSSource
        If gLastFound(lbcSSource) >= 2 Then
            If smSUpdateRvf(gLastFound(lbcSSource) - 2) <> "A" Then
                mPartColOk = False
                Exit Function
            End If
        Else
            mPartColOk = False
            Exit Function
        End If
    End If
    If (ilCol = EXTUPDATEINDEX) Then
        slStr = grdParticipant.TextMatrix(ilRow, SSOURCEINDEX)
        gFindMatch slStr, 2, lbcSSource
        If gLastFound(lbcSSource) >= 2 Then
            If smSUpdateRvf(gLastFound(lbcSSource) - 2) <> "A" Then
                mPartColOk = False
                Exit Function
            End If
        Else
            mPartColOk = False
            Exit Function
        End If
        If tmVef.sType <> "R" Then
            mPartColOk = False
            Exit Function
        End If
    End If
End Function


Private Sub mMoveCtrlToPif()
    Dim ilIndex As Integer
    Dim llRow As Long
    Dim slStartDate As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilPif As Integer
    Dim llStartDate As Long
    Dim llTestDate As Long
    Dim ilSSourceLastFd As Integer
    Dim ilStartIndex As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilSeqNo As Integer
    Dim ilNew As Integer
    Dim llOrigStartDate As Long

    If imCbcParticipantListIndex = -1 Then
        Exit Sub
    End If
    'slStartDate = edcParticipantDate.Text
    slStartDate = csiParticipantDate.Text
    If Not gValidDate(slStartDate) Then
        'Beep
        Exit Sub
    End If
    ilNew = True
    If imCbcParticipantListIndex > 0 Then
        If InStr(1, cbcParticipant.List(imCbcParticipantListIndex), "-New", vbTextCompare) <= 0 Then
            ilNew = False
        End If
    End If
    llStartDate = gDateValue(slStartDate)
    llOrigStartDate = -1
    If smOrigPartStartDate <> "" Then
        If gValidDate(smOrigPartStartDate) Then
            llOrigStartDate = gDateValue(smOrigPartStartDate)
        End If
    End If
    ilIndex = -1
    For ilPif = LBound(tmPifRec) To UBound(tmPifRec) - 1 Step 1
        gUnpackDateLong tmPifRec(ilPif).tPif.iStartDate(0), tmPifRec(ilPif).tPif.iStartDate(1), llTestDate
        If (llTestDate = llOrigStartDate) Or (llOrigStartDate = -1) Then
            If (imCbcParticipantListIndex > 0) And (Not ilNew) Then
                If tmPifRec(ilPif).tPif.lCode <> 0 Then
                    ilIndex = ilPif
                    Exit For
                ElseIf (tmPifRec(ilPif).tPif.lCode = 0) And (tmPifRec(ilPif).iStatus = 2) Then
                    ilIndex = ilPif
                    Exit For
                End If
            Else
                If (tmPifRec(ilPif).tPif.lCode = 0) And (tmPifRec(ilPif).iStatus = 3) Then
                    ilIndex = ilPif
                    Exit For
                End If
            End If
        End If
    Next ilPif
    ilSeqNo = 1
    For llRow = grdParticipant.FixedRows To grdParticipant.Rows - 1 Step 1
        'If grdParticipant.TextMatrix(llRow, SSOURCEINDEX) <> "" Then
        slStr = grdParticipant.TextMatrix(llRow, SSOURCEINDEX)
        If (slStr <> "") And (slStr <> "[None]") Then
            If ilIndex = -1 Then
                ilIndex = UBound(tmPifRec)
                tmPifRec(ilIndex).tPif.lCode = 0
                If (imCbcParticipantListIndex > 0) And (Not ilNew) Then
                    tmPifRec(ilIndex).iStatus = 2
                Else
                    tmPifRec(ilIndex).iStatus = 3
                End If
                ReDim Preserve tmPifRec(0 To UBound(tmPifRec) + 1) As PIFREC
            End If
            tmPifRec(ilIndex).tPif.iMnfSSCode = 0
            tmPifRec(ilIndex).tPif.iMnfGroup = 0
            tmPifRec(ilIndex).tPif.iProdPct = 0
            tmPifRec(ilIndex).tPif.sUpdateRVF = ""
            tmPifRec(ilIndex).tPif.sExtUpdateRvf = ""
            slStr = grdParticipant.TextMatrix(llRow, SSOURCEINDEX)
            gFindMatch slStr, 2, lbcSSource
            If gLastFound(lbcSSource) >= 2 Then
                ilSSourceLastFd = gLastFound(lbcSSource)
                slNameCode = tmSSourceCode(ilSSourceLastFd - 2).sKey  'lbcSSourceCode.List(gLastFound(lbcSSource) - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmPifRec(ilIndex).tPif.iMnfSSCode = Val(slCode)
                slStr = grdParticipant.TextMatrix(llRow, PARTINDEX)
                gFindMatch slStr, 1, lbcVehGp
                If gLastFound(lbcVehGp) >= 1 Then
                    slNameCode = tmVehGpCode(gLastFound(lbcVehGp) - 1).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tmPifRec(ilIndex).tPif.iMnfGroup = Val(slCode)
                End If
                If smSUpdateRvf(ilSSourceLastFd - 2) = "A" Then
                    slStr = grdParticipant.TextMatrix(llRow, INTUPDATEINDEX)
                    If StrComp(slStr, "Receivables", vbTextCompare) = 0 Then
                        tmPifRec(ilIndex).tPif.sUpdateRVF = "Y"
                    ElseIf StrComp(slStr, "History", vbTextCompare) = 0 Then
                        tmPifRec(ilIndex).tPif.sUpdateRVF = "N"
                    ElseIf StrComp(slStr, "Export+History", vbTextCompare) = 0 Then
                        tmPifRec(ilIndex).tPif.sUpdateRVF = "E"
                    ElseIf StrComp(slStr, "Export+A/R", vbTextCompare) = 0 Then
                        tmPifRec(ilIndex).tPif.sUpdateRVF = "F"
                    End If
                    If tmVef.sType = "R" Then
                        slStr = grdParticipant.TextMatrix(llRow, EXTUPDATEINDEX)
                        If StrComp(slStr, "Receivables", vbTextCompare) = 0 Then
                            tmPifRec(ilIndex).tPif.sExtUpdateRvf = "Y"
                        ElseIf StrComp(slStr, "History", vbTextCompare) = 0 Then
                            tmPifRec(ilIndex).tPif.sExtUpdateRvf = "N"
                        End If
                    End If
                End If
                slStr = grdParticipant.TextMatrix(llRow, PRODPCTINDEX)
                tmPifRec(ilIndex).tPif.iProdPct = gStrDecToInt(slStr, 2)
                gPackDate slStartDate, tmPifRec(ilIndex).tPif.iStartDate(0), tmPifRec(ilIndex).tPif.iStartDate(1)
                tmPifRec(ilIndex).tPif.iSeqNo = ilSeqNo
                ilSeqNo = ilSeqNo + 1
                tmPifRec(ilIndex).tPif.iVefCode = tmVef.iCode
                'Get Next Index
                ilStartIndex = ilIndex + 1
                ilIndex = -1
                For ilPif = ilStartIndex To UBound(tmPifRec) - 1 Step 1
                    gUnpackDateLong tmPifRec(ilPif).tPif.iStartDate(0), tmPifRec(ilPif).tPif.iStartDate(1), llTestDate
                    If llTestDate = llOrigStartDate Then
                        If (imCbcParticipantListIndex > 0) And (Not ilNew) Then
                            If tmPifRec(ilPif).tPif.lCode > 0 Then
                                ilIndex = ilPif
                                Exit For
                            ElseIf (tmPifRec(ilPif).tPif.lCode = 0) And (tmPifRec(ilPif).iStatus = 2) Then
                                ilIndex = ilPif
                                Exit For
                            End If
                        Else
                            If (tmPifRec(ilPif).tPif.lCode = 0) And (tmPifRec(ilPif).iStatus = 3) Then
                                ilIndex = ilPif
                                Exit For
                            End If
                        End If
                    End If
                Next ilPif
            End If
        End If
    Next llRow
    'Remove unused
    If ilIndex <> -1 Then
        ilPif = ilIndex
        Do While ilPif <= UBound(tmPifRec) - 1
            gUnpackDateLong tmPifRec(ilPif).tPif.iStartDate(0), tmPifRec(ilPif).tPif.iStartDate(1), llTestDate
            If llTestDate = llOrigStartDate Then
                If (imCbcParticipantListIndex > 0) And (Not ilNew) Then
                    If tmPifRec(ilPif).tPif.lCode > 0 Then
                        tmPifRec(ilPif).tPif.lCode = -tmPifRec(ilPif).tPif.lCode
                        ilPif = ilPif + 1
                    ElseIf (tmPifRec(ilPif).tPif.lCode = 0) And (tmPifRec(ilPif).iStatus = 2) Then
                        For ilLoop = ilPif To UBound(tmPifRec) - 1 Step 1
                            tmPifRec(ilLoop) = tmPifRec(ilLoop + 1)
                        Next ilLoop
                        ReDim Preserve tmPifRec(0 To UBound(tmPifRec) - 1) As PIFREC
                    Else
                        ilPif = ilPif + 1
                    End If
                Else
                    If (tmPifRec(ilPif).tPif.lCode = 0) And (tmPifRec(ilPif).iStatus = 3) Then
                        For ilLoop = ilPif To UBound(tmPifRec) - 1 Step 1
                            tmPifRec(ilLoop) = tmPifRec(ilLoop + 1)
                        Next ilLoop
                        ReDim Preserve tmPifRec(0 To UBound(tmPifRec) - 1) As PIFREC
                    Else
                        ilPif = ilPif + 1
                    End If
                End If
            Else
                ilPif = ilPif + 1
            End If
        Loop
    End If
End Sub

Private Sub mMovePifToCtrl()
    Dim slDates As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilLoop As Integer
    Dim slSalesSource As String
    Dim slParticipant As String
    Dim slUpdateRvf As String
    Dim slExtUpdateRvf As String
    Dim slPercent As String
    Dim ilSS As Integer
    Dim ilVef As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String
    Dim llTestDate As Long

    If imCbcParticipantListIndex > 0 Then
        slDates = cbcParticipant.List(imCbcParticipantListIndex)
        If InStr(1, slDates, "-New", vbTextCompare) <= 0 Then
            'edcParticipantDate.Enabled = False
            'edcParticipantDate.BackColor = LIGHTYELLOW
            csiParticipantDate.SetEnabled False
            csiParticipantDate.BackColor = LIGHTYELLOW
            cmcClear.Enabled = False
        Else
            'edcParticipantDate.Enabled = True
            'edcParticipantDate.BackColor = CYAN
            csiParticipantDate.SetEnabled True
            csiParticipantDate.BackColor = CYAN
            cmcClear.Enabled = True
        End If
        ilPos = InStr(1, slDates, "-", vbTextCompare)
    Else
        'edcParticipantDate.Enabled = True
        'edcParticipantDate.BackColor = CYAN
        csiParticipantDate.SetEnabled True
        csiParticipantDate.BackColor = CYAN
        If cbcParticipant.ListCount > 1 Then
            slDates = cbcParticipant.List(cbcParticipant.ListCount - 1)
            ilPos = InStr(1, slDates, "-", vbTextCompare)
        Else
            ilPos = 0
        End If
        cmcClear.Enabled = True
    End If
    If ilPos > 0 Then
        slStartDate = Left$(slDates, ilPos - 1)
        llStartDate = gDateValue(slStartDate)
        If imCbcParticipantListIndex > 0 Then
            'edcParticipantDate.Text = slStartDate
            csiParticipantDate.Text = slStartDate
            smOrigPartStartDate = slStartDate
        End If
        'ilRet = gObtainPIFForDate(hmPif, tmVef.iCode, slStartDate, tmPifRec())
        ilRow = grdParticipant.FixedRows
        For ilLoop = 0 To UBound(tmPifRec) - 1 Step 1
            gUnpackDateLong tmPifRec(ilLoop).tPif.iStartDate(0), tmPifRec(ilLoop).tPif.iStartDate(1), llTestDate
            If (llTestDate = llStartDate) And (tmPifRec(ilLoop).tPif.lCode >= 0) Then
                If ilRow >= grdParticipant.Rows Then
                    grdParticipant.AddItem ""
                    grdParticipant.RowHeight(ilRow) = fgBoxGridH + 15
                End If
                grdParticipant.Row = ilRow
                slSalesSource = ""
                slParticipant = ""
                slUpdateRvf = ""
                slExtUpdateRvf = ""
                slPercent = ""
                If (tmPifRec(ilLoop).tPif.iMnfSSCode > 0) And (tmPifRec(ilLoop).tPif.iMnfGroup > 0) Then
                    For ilSS = 0 To UBound(tmSSourceCode) - 1 Step 1  'lbcSSourceCode.ListCount - 1 Step 1
                        slNameCode = tmSSourceCode(ilSS).sKey    'lbcSSourceCode.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = tmPifRec(ilLoop).tPif.iMnfSSCode Then
                            ilRet = gParseItem(slNameCode, 1, "\", slSalesSource)
                            For ilVef = 0 To UBound(tmVehGpCode) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
                                slNameCode = tmVehGpCode(ilVef).sKey   'lbcVehGpCode.List(ilVef)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                If Val(slCode) = tmPifRec(ilLoop).tPif.iMnfGroup Then
                                    ilRet = gParseItem(slNameCode, 1, "\", slParticipant)
                                    Exit For
                                End If
                            Next ilVef
                            If smSUpdateRvf(ilSS) = "A" Then
                                slStr = tmPifRec(ilLoop).tPif.sUpdateRVF
                            Else
                                slStr = smSUpdateRvf(ilSS)
                            End If
                            Select Case slStr
                                Case "Y"
                                    slUpdateRvf = "Receivables"
                                Case "N"
                                    slUpdateRvf = "History"
                                Case "E"
                                    slUpdateRvf = "Export+History"
                                Case "F"
                                    slUpdateRvf = "Export+A/R"
                                Case Else
                                    slUpdateRvf = ""
                            End Select
                            If tmVef.sType = "R" Then
                                If smSUpdateRvf(ilSS) = "A" Then
                                    slStr = tmPifRec(ilLoop).tPif.sExtUpdateRvf
                                Else
                                    slStr = smSUpdateRvf(ilSS)
                                End If
                                Select Case slStr
                                    Case "Y"
                                        slExtUpdateRvf = "Receivables"
                                    Case "N"
                                        slExtUpdateRvf = "History"
                                    Case "E"
                                        slExtUpdateRvf = "History"
                                    Case "F"
                                        slExtUpdateRvf = "Receivables"
                                    Case Else
                                        slExtUpdateRvf = ""
                                End Select
                            Else
                                slExtUpdateRvf = ""
                            End If
                            If (smSUpdateRvf(ilSS) = "A") And (llStartDate <= lmLastBilledDate) And (tmPifRec(ilLoop).tPif.lCode > 0) Then
                                ilRet = True
                                'If edcParticipantDate.BackColor = LIGHTYELLOW Then
                                If csiParticipantDate.BackColor = LIGHTYELLOW Then
                                    For ilCol = grdParticipant.FixedCols To grdParticipant.Cols - 1 Step 1
                                        grdParticipant.Col = ilCol
                                        grdParticipant.CellBackColor = LIGHTYELLOW
                                    Next ilCol
                                End If
                            End If
                            slPercent = gIntToStrDec(tmPifRec(ilLoop).tPif.iProdPct, 2)
                            Exit For
                        End If
                    Next ilSS
                End If
                grdParticipant.TextMatrix(ilRow, SSOURCEINDEX) = slSalesSource
                grdParticipant.TextMatrix(ilRow, PARTINDEX) = slParticipant
                grdParticipant.TextMatrix(ilRow, INTUPDATEINDEX) = slUpdateRvf
                grdParticipant.TextMatrix(ilRow, EXTUPDATEINDEX) = slExtUpdateRvf
                grdParticipant.TextMatrix(ilRow, PRODPCTINDEX) = slPercent
                If imCbcParticipantListIndex > 0 Then
                    grdParticipant.TextMatrix(ilRow, PIFCODEINDEX) = tmPifRec(ilLoop).tPif.lCode
                Else
                    grdParticipant.TextMatrix(ilRow, PIFCODEINDEX) = "0"
                End If
                ilRow = ilRow + 1
            End If
        Next ilLoop
    End If
    imPifChg = False
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGridFieldsOk                   *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mGridFieldsOk() As Integer
'
'   iRet = mGridFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim slStr As String
    Dim llRow As Long
    Dim llCol As Long
    Dim ilError As Integer
    Dim ilSSourceLastFd As Integer
    Dim slTotalPct As String
    Dim llTest As Long
    Dim slDate As String

    ilError = False
    grdParticipant.Redraw = False
    'slStr = edcParticipantDate.Text
    slStr = csiParticipantDate.Text
    If slStr <> "" Then
        If Not gValidDate(slStr) Then
            MsgBox "Participant Date form not valid", vbOKOnly + vbExclamation, "Error"
            ilError = True
        End If
        slDate = gObtainStartStd(slStr)
        If gDateValue(slDate) <> gDateValue(slStr) Then
            MsgBox "Participant Start date must be on the start of a Standard Broadcast month", vbOKOnly + vbExclamation, "Error"
            ilError = True
        End If
    End If
    For llRow = grdParticipant.FixedRows To grdParticipant.Rows - 1 Step 1
        grdParticipant.Row = llRow
        For llCol = grdParticipant.FixedCols To grdParticipant.Cols - 1 Step 1
            grdParticipant.Col = llCol
            grdParticipant.CellForeColor = vbBlack
        Next llCol
    Next llRow
    For llRow = grdParticipant.FixedRows To grdParticipant.Rows - 1 Step 1
        slStr = Trim$(grdParticipant.TextMatrix(llRow, SSOURCEINDEX))
        If (slStr <> "") And (slStr <> "[None]") Then
            If (Trim$(grdParticipant.TextMatrix(llRow, PARTINDEX)) = "") Or (grdParticipant.TextMatrix(llRow, PARTINDEX) = "Missing") Then
                grdParticipant.TextMatrix(llRow, PARTINDEX) = "Missing"
                ilError = True
                grdParticipant.Row = llRow
                grdParticipant.Col = PARTINDEX
                grdParticipant.CellForeColor = vbRed
            End If
            slStr = grdParticipant.TextMatrix(llRow, SSOURCEINDEX)
            gFindMatch slStr, 2, lbcSSource
            If gLastFound(lbcSSource) >= 2 Then
                ilSSourceLastFd = gLastFound(lbcSSource)
                If smSUpdateRvf(ilSSourceLastFd - 2) = "A" Then
                    If (Trim$(grdParticipant.TextMatrix(llRow, INTUPDATEINDEX)) = "") Or (grdParticipant.TextMatrix(llRow, INTUPDATEINDEX) = "Missing") Then
                        grdParticipant.TextMatrix(llRow, INTUPDATEINDEX) = "Missing"
                        ilError = True
                        grdParticipant.Row = llRow
                        grdParticipant.Col = INTUPDATEINDEX
                        grdParticipant.CellForeColor = vbRed
                    End If
                    If tmVef.sType = "R" Then
                        If (Trim$(grdParticipant.TextMatrix(llRow, EXTUPDATEINDEX)) = "") Or (grdParticipant.TextMatrix(llRow, EXTUPDATEINDEX) = "Missing") Then
                            grdParticipant.TextMatrix(llRow, EXTUPDATEINDEX) = "Missing"
                            ilError = True
                            grdParticipant.Row = llRow
                            grdParticipant.Col = EXTUPDATEINDEX
                            grdParticipant.CellForeColor = vbRed
                        End If
                    End If
                End If
            End If
            If (Trim$(grdParticipant.TextMatrix(llRow, PRODPCTINDEX)) = "") Or (grdParticipant.TextMatrix(llRow, PRODPCTINDEX) = "Missing") Then
                grdParticipant.TextMatrix(llRow, PRODPCTINDEX) = "Missing"
                ilError = True
                grdParticipant.Row = llRow
                grdParticipant.Col = PRODPCTINDEX
                grdParticipant.CellForeColor = vbRed
            End If
        Else
            If llRow = grdParticipant.FixedRows Then
                'If Trim$(edcParticipantDate.Text) <> "" Then
                If Trim$(csiParticipantDate.Text) <> "" Then
                    grdParticipant.TextMatrix(llRow, SSOURCEINDEX) = "Missing"
                    ilError = True
                    grdParticipant.Row = llRow
                    grdParticipant.Col = SSOURCEINDEX
                    grdParticipant.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    If Not ilError Then
        For llRow = grdParticipant.FixedRows To grdParticipant.Rows - 1 Step 1
            slStr = Trim$(grdParticipant.TextMatrix(llRow, SSOURCEINDEX))
            If (slStr <> "") And (slStr <> "[None]") Then
                slTotalPct = ".00"
                For llTest = grdParticipant.FixedRows To grdParticipant.Rows - 1 Step 1
                    slStr = Trim$(grdParticipant.TextMatrix(llTest, SSOURCEINDEX))
                    If (slStr <> "") And (slStr <> "[None]") Then
                        If Trim$(grdParticipant.TextMatrix(llRow, SSOURCEINDEX)) = Trim$(grdParticipant.TextMatrix(llTest, SSOURCEINDEX)) Then
                            slTotalPct = gAddStr(slTotalPct, grdParticipant.TextMatrix(llTest, PRODPCTINDEX))
                        End If
                    End If
                Next llTest
                If gCompNumberStr(slTotalPct, "100.00") <> 0 Then
                    ilError = True
                    grdParticipant.Row = llRow
                    grdParticipant.Col = PRODPCTINDEX
                    grdParticipant.CellForeColor = vbRed
                End If
            End If
        Next llRow
    End If
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
    grdParticipant.Redraw = True
End Function

Private Sub mAdjPifDates()
    Dim ilPif As Integer
    Dim ilIndex As Integer
    Dim llStartDate As Long
    Dim llTestStartDate As Long
    Dim llTestEndDate As Long
    Dim llSetEndDate As Long
    Dim ilPrevChg As Integer
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slEndDate As String
    Dim slDates As String

    For ilPif = 0 To UBound(tmPifRec) - 1 Step 1
        If (tmPifRec(ilPif).tPif.lCode = 0) And (tmPifRec(ilPif).iStatus = 3) Then
            gUnpackDateLong tmPifRec(ilPif).tPif.iStartDate(0), tmPifRec(ilPif).tPif.iStartDate(1), llStartDate
            ilPrevChg = False
            For ilIndex = 0 To ilPif - 1 Step 1
                gUnpackDateLong tmPifRec(ilIndex).tPif.iStartDate(0), tmPifRec(ilIndex).tPif.iStartDate(1), llTestStartDate
                If (llTestStartDate = llStartDate) And (tmPifRec(ilIndex).tPif.lCode = 0) Then
                    ilPrevChg = True
                    Exit For
                End If
            Next ilIndex
            If Not ilPrevChg Then
                llSetEndDate = 0
                ilIndex = 0
                Do While ilIndex <= UBound(tmPifRec) - 1
                    If tmPifRec(ilIndex).tPif.lCode > 0 Then
                        gUnpackDateLong tmPifRec(ilIndex).tPif.iStartDate(0), tmPifRec(ilIndex).tPif.iStartDate(1), llTestStartDate
                        gUnpackDateLong tmPifRec(ilIndex).tPif.iEndDate(0), tmPifRec(ilIndex).tPif.iEndDate(1), llTestEndDate
                        If (llStartDate >= llTestStartDate) And llStartDate <= llTestEndDate Then
                            If llSetEndDate = 0 Then
                                llSetEndDate = llTestEndDate
                            End If
                            If llStartDate = llTestStartDate Then
                                tmPifRec(ilIndex).tPif.lCode = -tmPifRec(ilIndex).tPif.lCode
                            Else
                                gPackDateLong llStartDate - 1, tmPifRec(ilIndex).tPif.iEndDate(0), tmPifRec(ilIndex).tPif.iEndDate(1)
                            End If
                        End If
                        ilIndex = ilIndex + 1
                    ElseIf (tmPifRec(ilIndex).tPif.lCode = 0) And (tmPifRec(ilIndex).iStatus = 2) Then
                        gUnpackDateLong tmPifRec(ilIndex).tPif.iStartDate(0), tmPifRec(ilIndex).tPif.iStartDate(1), llTestStartDate
                        gUnpackDateLong tmPifRec(ilIndex).tPif.iEndDate(0), tmPifRec(ilIndex).tPif.iEndDate(1), llTestEndDate
                        If (llStartDate >= llTestStartDate) And llStartDate <= llTestEndDate Then
                            If llSetEndDate = 0 Then
                                llSetEndDate = llTestEndDate
                            End If
                            If llStartDate = llTestStartDate Then
                                For ilLoop = ilIndex To UBound(tmPifRec) - 1 Step 1
                                    tmPifRec(ilLoop) = tmPifRec(ilLoop + 1)
                                Next ilLoop
                                ReDim Preserve tmPifRec(0 To UBound(tmPifRec) - 1) As PIFREC
                            Else
                                gPackDateLong llStartDate - 1, tmPifRec(ilIndex).tPif.iEndDate(0), tmPifRec(ilIndex).tPif.iEndDate(1)
                                ilIndex = ilIndex + 1
                            End If
                        Else
                            ilIndex = ilIndex + 1
                        End If
                    Else
                        ilIndex = ilIndex + 1
                    End If
                Loop
                If llSetEndDate = 0 Then
                    llSetEndDate = gDateValue("12/31/2069")
                End If
                For ilIndex = 0 To UBound(tmPifRec) - 1 Step 1
                    gUnpackDateLong tmPifRec(ilIndex).tPif.iStartDate(0), tmPifRec(ilIndex).tPif.iStartDate(1), llTestStartDate
                    If (llTestStartDate = llStartDate) And (tmPifRec(ilIndex).tPif.lCode = 0) And (tmPifRec(ilPif).iStatus = 3) Then
                        gPackDateLong llSetEndDate, tmPifRec(ilIndex).tPif.iEndDate(0), tmPifRec(ilIndex).tPif.iEndDate(1)
                    End If
                Next ilIndex

            End If
        ElseIf (tmPifRec(ilPif).tPif.lCode = 0) And (tmPifRec(ilPif).iStatus = 2) Then
            slDates = cbcParticipant.List(imCbcParticipantListIndex)
            If InStr(1, slDates, "-New", vbTextCompare) <= 0 Then
                ilPos = InStr(1, slDates, "-", vbTextCompare)
                slEndDate = Trim$(Mid$(slDates, ilPos + 1))
                If StrComp(slEndDate, "TFN", vbTextCompare) <> 0 Then
                    gPackDate slEndDate, tmPifRec(ilPif).tPif.iEndDate(0), tmPifRec(ilPif).tPif.iEndDate(1)
                Else
                    gPackDate "12/31/2069", tmPifRec(ilPif).tPif.iEndDate(0), tmPifRec(ilPif).tPif.iEndDate(1)
                End If
            End If
        End If
        gUnpackDateLong tmPifRec(ilPif).tPif.iEndDate(0), tmPifRec(ilPif).tPif.iEndDate(1), llSetEndDate
        If llSetEndDate = 0 Then
            gPackDate "12/31/2069", tmPifRec(ilPif).tPif.iEndDate(0), tmPifRec(ilPif).tPif.iEndDate(1)
        End If
    Next ilPif
End Sub

Private Function mCheckPartStartDate() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llStartDate                                                                           *
'******************************************************************************************

    Dim ilPif As Integer
    Dim slStartDate As String
    Dim slDate As String

    mCheckPartStartDate = 0
    'Need to allow new sales sources to be added even in the pasted
    'because of the case that a salesperson sales source has not been
    'defined within the vehicle.
    'In the reports, it does not look if the sales source is set to ask,
    'it only looks at the receivables to determine if the split was
    'preformed or not.
    'Since they can't alter any row in the past, only different sales sources will be added
    'For ilPif = 0 To UBound(tmPifRec) - 1 Step 1
    '    If tmPifRec(ilPif).tPif.lCode = 0 Then
    '        gUnpackDateLong tmPifRec(ilPif).tPif.iStartDate(0), tmPifRec(ilPif).tPif.iStartDate(1), llStartDate
    '        If llStartDate <= lmLastBilledDate Then
    '            mCheckPartStartDate = 1
    '            Exit Function
    '        End If
    '    End If
    'Next ilPif
    For ilPif = 0 To UBound(tmPifRec) - 1 Step 1
        If tmPifRec(ilPif).tPif.lCode = 0 Then
            gUnpackDate tmPifRec(ilPif).tPif.iStartDate(0), tmPifRec(ilPif).tPif.iStartDate(1), slStartDate
            slDate = gObtainStartStd(slStartDate)
            If gDateValue(slDate) <> gDateValue(slStartDate) Then
                mCheckPartStartDate = 2
                Exit Function
            End If
        End If
    Next ilPif
End Function


Private Sub mVtfReadRec()
    Dim ilRet As Integer
    Dim slStr As String
    
    tmVtfSrchKey.lCode = tmVff.lPledgeHdVtfCode
    If tmVff.lPledgeHdVtfCode <> 0 Then
        tmVtf.sText = ""
        imVtfRecLen = Len(tmVtf) '5011
        ilRet = btrGetEqual(hmVtf, tmVtf, imVtfRecLen, tmVtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmVtf.sText = ""
        End If
    Else
        tmVtf.sText = ""
    End If
    slStr = gStripChr0(tmVtf.sText)
    edcTextHd.MaxLength = 5000
    edcTextHd.SetText (slStr)

    tmVtfSrchKey.lCode = tmVff.lPledgeFtVtfCode
    If tmVff.lPledgeFtVtfCode <> 0 Then
        tmVtf.sText = ""
        imVtfRecLen = Len(tmVtf) '5011
        ilRet = btrGetEqual(hmVtf, tmVtf, imVtfRecLen, tmVtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmVtf.sText = ""
        End If
    Else
        tmVtf.sText = ""
    End If
    edcTextFt.MaxLength = 5000
    slStr = gStripChr0(tmVtf.sText)
    edcTextFt.SetText (slStr)
    
    lmT1VtfCode = 0
    tmVtfSrchKey1.iVefCode = tmVef.iCode
    tmVtfSrchKey1.sType = "1"
    imVtfRecLen = Len(tmVtf) '5011
    ilRet = btrGetEqual(hmVtf, tmVtf, imVtfRecLen, tmVtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        tmVtf.sText = ""
    Else
        lmT1VtfCode = tmVtf.lCode
    End If
    slStr = gStripChr0(tmVtf.sText)
    edcSchedule(2).Text = slStr

    lmT2VtfCode = 0
    tmVtfSrchKey1.iVefCode = tmVef.iCode
    tmVtfSrchKey1.sType = "2"
    imVtfRecLen = Len(tmVtf) '5011
    ilRet = btrGetEqual(hmVtf, tmVtf, imVtfRecLen, tmVtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        tmVtf.sText = ""
    Else
        lmT2VtfCode = tmVtf.lCode
    End If
    slStr = gStripChr0(tmVtf.sText)
    edcSchedule(3).Text = slStr

End Sub

Private Function mSaveVtf(slType As String) As Long
    Dim llVtfCode As Long
    Dim slStr As String
    Dim slTextOnly As String
    Dim ilRet As Integer
    
    imVtfRecLen = Len(tmVtf) '- Len(tmCsf.sComment) + Len(Trim$(tmCsf.sComment)) ' + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
    If slType = "H" Then
        slStr = edcTextHd.Text
        slTextOnly = edcTextHd.TextOnly
        llVtfCode = tmVff.lPledgeHdVtfCode
    ElseIf slType = "1" Then
        slStr = edcSchedule(2).Text
        slTextOnly = edcSchedule(2).Text
        llVtfCode = lmT1VtfCode
    ElseIf slType = "2" Then
        slStr = edcSchedule(3).Text
        slTextOnly = edcSchedule(3).Text
        llVtfCode = lmT2VtfCode
    Else
        slStr = edcTextFt.Text
        slTextOnly = edcTextFt.TextOnly
        llVtfCode = tmVff.lPledgeFtVtfCode
    End If
    If llVtfCode > 0 Then
        tmVtfSrchKey.lCode = llVtfCode
        ilRet = btrGetEqual(hmVtf, tmVtf, imVtfRecLen, tmVtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If Len(slTextOnly) > 0 Then
                tmVtf.sText = Trim$(slStr) & Chr$(0)
                ilRet = btrUpdate(hmVtf, tmVtf, imVtfRecLen)
                mSaveVtf = tmVtf.lCode
            Else
                ilRet = btrDelete(hmVtf)
                mSaveVtf = 0
            End If
        Else
            mSaveVtf = 0
        End If
    Else
        If Len(slTextOnly) > 0 Then
            tmVtf.lCode = 0
            tmVtf.iVefCode = tmVef.iCode
            tmVtf.sType = slType
            tmVtf.sText = Trim$(slStr) & Chr$(0)
            ilRet = btrInsert(hmVtf, tmVtf, imVtfRecLen, INDEXKEY0)
            If ilRet = BTRV_ERR_NONE Then
                mSaveVtf = tmVtf.lCode
            Else
                mSaveVtf = 0
            End If
        Else
            mSaveVtf = 0
        End If
    End If
End Function

Private Sub mSeasonPop()
    Dim llStartDate As Long
    Dim slStartDate As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    lbcSeason.Clear
    ReDim tmSeasonInfo(0 To 0) As SEASONINFO
    If tmVef.sType <> "G" Then
        Exit Sub
    End If
    tmGhfSrchKey1.iVefCode = tmVef.iCode
    ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.iVefCode = tmVef.iCode)
        gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llStartDate
        slStartDate = Trim$(Str$(llStartDate))
        Do While Len(slStartDate) < 6
            slStartDate = "0" & slStartDate
        Loop
        tmSeasonInfo(UBound(tmSeasonInfo)).sKey = slStartDate
        tmSeasonInfo(UBound(tmSeasonInfo)).sSeasonName = tmGhf.sSeasonName
        tmSeasonInfo(UBound(tmSeasonInfo)).lCode = tmGhf.lCode
        ReDim Preserve tmSeasonInfo(0 To UBound(tmSeasonInfo) + 1) As SEASONINFO
        ilRet = btrGetNext(hmGhf, tmGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    If UBound(tmSeasonInfo) > 1 Then
        'Sort descending
        ArraySortTyp fnAV(tmSeasonInfo(), 0), UBound(tmSeasonInfo), 1, LenB(tmSeasonInfo(0)), 0, LenB(tmSeasonInfo(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tmSeasonInfo) - 1 Step 1
        lbcSeason.AddItem Trim$(tmSeasonInfo(ilLoop).sSeasonName)
        lbcSeason.ItemData(lbcSeason.NewIndex) = tmSeasonInfo(ilLoop).lCode
    Next ilLoop
    For ilLoop = 0 To lbcSeason.ListCount - 1 Step 1
        If lbcSeason.ItemData(ilLoop) = tmVff.lSeasonGhfCode Then
            lbcSeason.ListIndex = ilLoop
            Exit For
        End If
    Next ilLoop
    
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName()
    Dim slStr As String
    Dim ilVef As Integer
    Dim ilVff As Integer

    If (tmVef.sType = "P") Or (tmVef.sType = "S") Then
        mOKName = True
        Exit Function
    End If
    If edcLog(3).Text <> "" Then    'Test name
        slStr = UCase(Trim$(edcLog(3).Text))
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If tmVef.iCode <> tgMVef(ilVef).iCode Then
                If (tgMVef(ilVef).sType <> "P") And (tgMVef(ilVef).sType <> "S") Then
                    If StrComp(UCase(Trim$(tgMVef(ilVef).sName)), slStr, vbTextCompare) = 0 Then
                        Beep
                        MsgBox "Override Affidavit Name already defined as Vehicle Name, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                        mOKName = False
                        Exit Function
                    Else
                        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                            If tgMVef(ilVef).iCode = tgVff(ilVff).iVefCode Then
                                If StrComp(UCase(Trim$(tgVff(ilVff).sWebName)), slStr, vbTextCompare) = 0 Then
                                    Beep
                                    MsgBox "Override Affidavit Name already defined as Override Name, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                                    mOKName = False
                                    Exit Function
                                End If
                            End If
                        Next ilVff
                    End If
                End If
            End If
        Next ilVef
    End If
    mOKName = True
End Function

Private Sub mMediaPop()
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffset(0) As Integer
    Dim slName As String
    Dim ilIndex As Integer
    If sgUseCartNo = "N" Then
        cbcMedia.AddItem "[None]", 0  'Force as first item on list
        ReDim tmMediaCode(0 To 0) As SORTCODE
        Exit Sub
    End If
    ilIndex = cbcMedia.ListIndex
    If ilIndex >= 0 Then
        slName = cbcMedia.List(ilIndex)
    End If
    ilfilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffset(0) = 0
    'ilRet = gIMoveListBox(CopyInv, cbcMedia, lbcMediaCode, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilFilter(), slFilter(), ilOffset())
    '10071
    ilRet = gIMoveListBox(Me, cbcMedia, tmMediaCode(), smMediaCodeTag, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilfilter(), slFilter(), ilOffset())
'    ilRet = gIMoveListBox(CopyInv, cbcMedia, tmMediaCode(), smMediaCodeTag, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mMediaPopErr
        '10071
        gCPErrorMsg ilRet, "mMediaPop (gIMoveListBox)", Me
       ' gCPErrorMsg ilRet, "mMediaPop (gIMoveListBox)", CopyInv
        On Error GoTo 0
        cbcMedia.AddItem "[None]", 0  'Force as first item on list
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
    Exit Sub
mMediaPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mInitVbf(ilVefCode, tlVbf As VBF)
    Dim ilLoop As Integer
    
    tlVbf.lCode = 0
    tlVbf.iVefCode = ilVefCode
    gPackDate "", tlVbf.iStartDate(0), tlVbf.iStartDate(1)
    gPackDate "", tlVbf.iEndDate(0), tlVbf.iEndDate(1)
    'For ilLoop = 1 To 10 Step 1
    For ilLoop = 0 To 9 Step 1
        tlVbf.iSpotLen(ilLoop) = 0
        tlVbf.lDefAcqCost(ilLoop) = 0
        tlVbf.lActAcqCost(ilLoop) = 0
    Next ilLoop
    tlVbf.iThreshold = 0
    tlVbf.iXFree = 0
    tlVbf.iYSold = 0
    tlVbf.sMethod = "N"
    tlVbf.lInsertionCefCode = 0
    tlVbf.iAcqCommPct = 0
    tlVbf.lBalance = 0
    gPackDate "", tlVbf.iBalanceDate(0), tlVbf.iBalanceDate(1)
    tlVbf.sPerPeriod = "W"
End Sub
Private Sub mGetAffiliateSite()
    Dim ilRet As Integer
    Dim hlSite As Integer
    Dim tlSiteSrchKey As LONGKEY0    'Vef key record image
    Dim ilSiteRecLen As Integer
    
    hlSite = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlSite, "", sgDBPath & "Site.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        smAllowMGSpots = "Y"
        smAllowReplSpots = "N"
        smNoMissedReason = "N"
        ilRet = btrClose(hlSite)
        btrDestroy hlSite
        Exit Sub
    End If
    ilSiteRecLen = Len(tgSite)
    tlSiteSrchKey.lCode = 1
    ilRet = btrGetEqual(hlSite, tgSite, ilSiteRecLen, tlSiteSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        smAllowMGSpots = "Y"
        smAllowReplSpots = "N"
        smNoMissedReason = "N"
        ilRet = btrClose(hlSite)
        btrDestroy hlSite
        Exit Sub
    End If
    smAllowMGSpots = tgSite.sAllowMGSpots
    'temporarily disallow Replacement until ready to code on the Web
    smAllowReplSpots = "N"  'tgSite.sAllowReplSpots
    smNoMissedReason = tgSite.sNoMissedReason
    ilRet = btrClose(hlSite)
    btrDestroy hlSite
End Sub
Private Sub mXdsVendorChange()
    Dim bmXdsHonorMergeNow As Boolean
    Dim blIsChange As Boolean
    Dim blIsNationalToo As Boolean
    
    bmXdsHonorMergeNow = False
    If ckcXDSave(3).Value = vbChecked Then
        bmXdsHonorMergeNow = True
    End If
    If bmOriginalXdsHonorMerge <> bmXdsHonorMergeNow Then
        blIsChange = True
        blIsNationalToo = True
        bmOriginalXdsHonorMerge = bmXdsHonorMergeNow
    End If
    If Trim(smOriginalXdsProgramCode) <> Trim(edcInterfaceID(0).Text) Then
        If UCase(Trim(smOriginalXdsProgramCode)) = "MERGE" Or UCase(Trim(edcInterfaceID(0).Text)) = "MERGE" Then
            blIsChange = True
            smOriginalXdsProgramCode = edcInterfaceID(0).Text
        End If
    End If
    If blIsChange Then
        mVatSetToGoToWeb blIsNationalToo, tmVff.iVefCode
    End If
End Sub
Private Sub mVatSetToGoToWeb(blIsNationalToo As Boolean, ilVefCode As Integer)
    '7942
    Dim rst_pw As ADODB.Recordset
    Dim slSql As String
    On Error GoTo ErrHand
        
    slSql = "UPDATE VAT_Vendor_Agreement SET vatSentToWeb = '' WHERE vatattcode in (Select vatattcode from VAT_Vendor_Agreement inner join Att on vatattcode = attcode where attvefcode = " & ilVefCode & ") "
    If blIsNationalToo Then
        slSql = slSql & " AND (vatWvtVendorId = 114 OR vatwvtvendorid = 115) "
    Else
        slSql = slSql & " AND vatWvtVendorId = 114"
    End If
    Set rst_pw = gSQLSelectCall(slSql)
    Exit Sub
    
ErrHand:
End Sub

Private Function mUpdateArf(ilArfCode As Integer, slFTP As String) As Integer
    Dim ilRet As Integer
    
    tmArfSrchKey.iCode = ilArfCode
    If tmArfSrchKey.iCode <> 0 Then
        ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            If Trim$(slFTP) <> "" Then
                tmArf.sFTP = slFTP
                ilRet = btrUpdate(hmArf, tmArf, imArfRecLen)
                mUpdateArf = tmArf.iCode
            Else
                ilRet = btrDelete(hmArf)
                mUpdateArf = 0
            End If
        Else
            If Trim$(slFTP) <> "" Then
                mInitArf
                tmArf.sFTP = slFTP
                ilRet = btrInsert(hmArf, tmArf, imArfRecLen, INDEXKEY0)
                mUpdateArf = tmArf.iCode
            Else
                mUpdateArf = 0
            End If
        End If
    Else
        If Trim$(slFTP) <> "" Then
            mInitArf
            tmArf.sFTP = slFTP
            ilRet = btrInsert(hmArf, tmArf, imArfRecLen, INDEXKEY0)
            mUpdateArf = tmArf.iCode
        Else
            mUpdateArf = 0
        End If
    End If
End Function
Private Sub mVendorsLoadAndSelect(blLoad As Boolean, ilAvfCode As Integer)
    Dim ilLoop As Integer
    Dim slSql As String
    
    imCurrentVendorInfoIndex = -1
    If blLoad Then
        cbcCsiGeneric(ADVENDOR).Clear
        cbcCsiGeneric(ADVENDOR).AddItem "[New]"
        cbcCsiGeneric(ADVENDOR).SetItemData = -1
        cbcCsiGeneric(ADVENDOR).AddItem "[None]"  'Force as first item on list
        cbcCsiGeneric(ADVENDOR).SetItemData = 0
        tmVendorInfo = gGetDigitalVendors(False)
        For ilLoop = 0 To UBound(tmVendorInfo) - 1 Step 1
            cbcCsiGeneric(ADVENDOR).AddItem Trim$(tmVendorInfo(ilLoop).sName)
            cbcCsiGeneric(ADVENDOR).SetItemData = tmVendorInfo(ilLoop).iCode
        Next
    End If
    If ilAvfCode > 0 Then
        For ilLoop = 0 To cbcCsiGeneric(ADVENDOR).ListCount - 1 Step 1
            If cbcCsiGeneric(ADVENDOR).GetItemData(ilLoop) = ilAvfCode Then
                imChgMode = True
                cbcCsiGeneric(ADVENDOR).SetListIndex = ilLoop
                mVendorSetIndex ilAvfCode
                imChgMode = False
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub mVendorSetIndex(ilAvfCode As Integer)
    Dim ilLoop As Integer
    For ilLoop = 0 To UBound(tmVendorInfo) - 1 Step 1
        If tmVendorInfo(ilLoop).iCode = ilAvfCode Then
            imCurrentVendorInfoIndex = ilLoop
            Exit For
        End If
    Next
End Sub
'Private Sub mGetPodcastInfo()
''10050
'    Dim slSql As String
'    Dim ilLoop As Integer
'    Dim ilRet As Integer
'
'    smAdVendorVehNameOriginal = ""
'    If tmVff.iAvfCode > 0 Then
'        For ilLoop = 0 To cbcCsiGeneric(ADVENDOR).ListCount - 1 Step 1
'            If cbcCsiGeneric(ADVENDOR).GetItemData(ilLoop) = tmVff.iAvfCode Then
'                cbcCsiGeneric(ADVENDOR).SetListIndex = ilLoop
'                Exit For
'            End If
'        Next ilLoop
'    Else
'        cbcCsiGeneric(ADVENDOR).SetListIndex = 1
'    End If
'    If tmVff.lAdVehNameCefCode > 0 Then
'        tmCefSrchKey.lCode = tmVff.lAdVehNameCefCode
'        tmCef.sComment = ""
'        imCefRecLen = Len(tmCef)    '1009
'        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'        If ilRet <> BTRV_ERR_NONE Then
'            'what does this do?
'            tmCef.lCode = 0
'        Else
'            smAdVendorVehNameOriginal = gStripChr0(tmCef.sComment)
'        End If
'    End If
'    edcGen(ADVENDORVEHICLENAME).Text = smAdVendorVehNameOriginal
'
'End Sub
Private Sub mVendorSetAndEnableInfoOriginal()
    '10981
    Dim ilLoop As Integer
    
    smVendorExternalIDOriginal = ""
    imCurrentVendorInfoIndex = -1
    If tmVff.iAvfCode > 0 Then
        For ilLoop = 0 To cbcCsiGeneric(ADVENDOR).ListCount - 1 Step 1
            If cbcCsiGeneric(ADVENDOR).GetItemData(ilLoop) = tmVff.iAvfCode Then
                imChgMode = True
                cbcCsiGeneric(ADVENDOR).SetListIndex = ilLoop
                mVendorSetIndex tmVff.iAvfCode
                'sets smVendorExternalIDOriginal
                mVendorSetExtID
                imChgMode = False
                Exit For
            End If
        Next ilLoop
    Else
        imChgMode = True
        cbcCsiGeneric(ADVENDOR).SetListIndex = 1
        imChgMode = False
    End If
    mVendorEnableOptions
End Sub
Private Function mVendorSetExtID() As String
    Dim slRet As String
    
    slRet = ""
    If imCurrentVendorInfoIndex > -1 Then
        If tmVendorInfo(imCurrentVendorInfoIndex).oExternalVehicleIDType <> ExternalVehicleIDType.None Then
            slRet = mVendorGetExtId(tmVendorInfo(imCurrentVendorInfoIndex).iCode)
        End If
    End If
    edcGen(ADVENDOREXTERNALIDINDEX).Text = slRet
    'going to database? Then we want to reset original
    smVendorExternalIDOriginal = slRet
    mVendorSetExtID = slRet
End Function
Private Function mVendorGetExtId(ilAvfCode As Integer) As String
    Dim slSql As String
    Dim slRet As String
    
    slSql = "Select vvcExternalID from VVC_Vendor_Vef where vvcavfcode = " & ilAvfCode & " and vvcvefcode = " & tmVff.iVefCode
    Set cmt_rst = gSQLSelectCall(slSql)
    If Not cmt_rst.EOF Then
        slRet = cmt_rst!vvcExternalID
    End If
    mVendorGetExtId = Trim$(slRet)
End Function
'Private Function mAdVendorNameToCef() As Long
'    Dim llRet As Long
'    Dim ilBtrRet As Integer
'    llRet = 0
'
'    ' update if name changed
'    If smAdVendorVehNameOriginal <> Trim$(edcGen(ADVENDORVEHICLENAME).Text) Then
'        If Len(Trim$(edcGen(ADVENDORVEHICLENAME).Text)) > 0 Then
'            smAdVendorVehNameOriginal = edcGen(ADVENDORVEHICLENAME).Text
'            tmCef.sComment = smAdVendorVehNameOriginal & Chr$(0)
'            imCefRecLen = Len(tmCef)
'            tmCef.lCode = tmVff.lAdVehNameCefCode
'            If tmVff.lAdVehNameCefCode > 0 Then
'                ilBtrRet = btrUpdate(hmCef, tmCef, imCefRecLen)
'                llRet = tmVff.lAdVehNameCefCode
'            Else
'                ilBtrRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
'                llRet = tmCef.lCode
'            End If
'            If ilBtrRet <> BTRV_ERR_NONE Then
'                'should I do more?
'                smAdVendorVehNameOriginal = ""
'                llRet = 0
'            End If
'        ElseIf tmVff.lAdVehNameCefCode > 0 Then
'            tmCefSrchKey.lCode = tmVff.lAdVehNameCefCode
'            imCefRecLen = Len(tmCef)
'            ilBtrRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'            If ilBtrRet <> BTRV_ERR_NONE Then
'                'what does this do?
'                tmCef.lCode = 0
'            Else
'                btrDelete hmCef
'            End If
'        End If
'    Else
'        llRet = tmVff.lAdVehNameCefCode
'    End If
'    mAdVendorNameToCef = llRet
'End Function
Public Function mWriteToVVC()
    Dim blRet As Boolean
    Dim slNewID As String
    Dim slSql As String
    
    blRet = True
    slNewID = Trim$(edcGen(ADVENDOREXTERNALIDINDEX).Text)
    If tmVff.iAvfCode > 0 Then
        If smVendorExternalIDOriginal <> slNewID Then
            If Len(smVendorExternalIDOriginal) = 0 Then
                slSql = "Insert Into VVC_Vendor_Vef (vvcAvfCode,vvcVefCode,vvcExternalID) Values (" & tmVff.iAvfCode & "," & tmVff.iVefCode & ",'" & gSqlSafeAndTrim(slNewID) & "')"
            ElseIf Len(slNewID) = 0 Then
                slSql = "Delete from VVC_Vendor_Vef where vvcavfcode = " & tmVff.iAvfCode & " AND vvcvefcode = " & tmVff.iVefCode
            Else
                slSql = "Update VVC_Vendor_Vef set vvcExternalID = '" & gSqlSafeAndTrim(slNewID) & "' WHERE vvcavfcode = " & tmVff.iAvfCode & " AND vvcVefCode = " & tmVff.iVefCode
            End If
            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                'handle error
                blRet = False
            Else
                smVendorExternalIDOriginal = slNewID
            End If
        End If
    End If

    mWriteToVVC = blRet
End Function
Private Sub mVendorEnableOptions()
    Dim blEnable As Boolean
    Dim ilLoop As Integer
    
    blEnable = False
    If rbcGMedium(PODCASTRBC).Value Then
        blEnable = True
    Else
        imChgMode = True
        edcGen(ADVENDOREXTERNALIDINDEX).Text = ""
        cbcCsiGeneric(ADVENDOR).SetListIndex = 1
        imChgMode = False
        imCurrentVendorInfoIndex = -1
        smVendorExternalIDOriginal = ""
    End If
    cbcCsiGeneric(ADVENDOR).Enabled blEnable
    If blEnable And imCurrentVendorInfoIndex > -1 Then
        blEnable = False
        If tmVendorInfo(imCurrentVendorInfoIndex).oExternalVehicleIDType = ExternalVehicleIDType.Allowed Then
            blEnable = True
        End If
    End If
    edcGen(ADVENDOREXTERNALIDINDEX).Enabled = blEnable
End Sub
Private Sub mGoToAdVendorForm()
    Dim myForm As New AdServerVendor
    Dim ilPasswordOk As Integer
    Dim ilPreviouslyPassedAdVendor As Integer
    Dim ilLoop As Integer
    
    '10467 stop coming in again after double click
    If Not bmAirPodDoubleClick Then
        If cbcCsiGeneric(ADVENDOR).ListIndex <> 1 Then
            If (Trim$(tgUrf(0).sName) <> sgCPName) And (imIgnoreClickEvent = False) Then
                If igPasswordOk Then
                    bmAirPodDoubleClick = True
                    ilPasswordOk = igPasswordOk
                    sgPasswordAddition = "ADV-"
                    CSPWord.Show vbModal
                    If Not igPasswordOk Then
                        'new' to 'none'
                        If cbcCsiGeneric(ADVENDOR).ListIndex = 0 Then
                            cbcCsiGeneric(ADVENDOR).SetListIndex = 1
                        End If
                    Else
                        imChgMode = True
                        imVffChg = True
                        Set myForm = New AdServerVendor
                        myForm.imPassedID = -1
                        If cbcCsiGeneric(ADVENDOR).GetItemData(cbcCsiGeneric(ADVENDOR).ListIndex) > -1 Then
                            ilPreviouslyPassedAdVendor = cbcCsiGeneric(ADVENDOR).GetItemData(cbcCsiGeneric(ADVENDOR).ListIndex)
                            myForm.imPassedID = ilPreviouslyPassedAdVendor
                        End If
                        myForm.Show vbModal
                        If myForm.bmNeedRefresh Then
                            mVendorsLoadAndSelect True, myForm.imPassedID
                        ElseIf ilPreviouslyPassedAdVendor <> myForm.imPassedID And myForm.imPassedID > 0 Then
                            For ilLoop = 0 To cbcCsiGeneric(ADVENDOR).ListCount - 1 Step 1
                                If cbcCsiGeneric(ADVENDOR).GetItemData(ilLoop) = myForm.imPassedID Then
                                    cbcCsiGeneric(ADVENDOR).SetListIndex = ilLoop
                                    Exit For
                                End If
                            Next ilLoop
                        ElseIf cbcCsiGeneric(ADVENDOR).ListIndex = 0 Then
                            cbcCsiGeneric(ADVENDOR).SetListIndex = 1
                        End If
                        Unload myForm
                        imChgMode = False
                    End If
                    sgPasswordAddition = ""
                    igPasswordOk = ilPasswordOk
                Else
                    cbcCsiGeneric(ADVENDOR).SetListIndex = 1
                End If
            End If
            bmAirPodDoubleClick = False
        End If
    End If
End Sub
Private Function mIsUsedInDigitalLine(ilVefCode As Integer) As Boolean
    Dim slSql As String
    Dim blRet As Boolean
    
    blRet = False
    slSql = "select count(*)as amount from pcf_Pod_cpm_cntr where pcfvefcode = " & ilVefCode
    Set cmt_rst = gSQLSelectCall(slSql)
    If Not cmt_rst.EOF Then
        If cmt_rst!amount > 0 Then
            blRet = True
        End If
    End If
    mIsUsedInDigitalLine = blRet
End Function
Private Sub mVendorLegacyAdjustBoostr()
    Dim slSql As String
    If imCurrentVendorInfoIndex > -1 Then
        If tmVendorInfo(imCurrentVendorInfoIndex).sName = "Boostr" And tmVendorInfo(imCurrentVendorInfoIndex).iCode = tmVff.iAvfCode Then
            If tmVef.lExtId > 0 Then
                slSql = "Select count(*)as amount From VVC_Vendor_Vef where vvcvefcode = 23 and vvcAvfcode = " & tmVff.iAvfCode
                Set cmt_rst = gSQLSelectCall(slSql)
                If Not cmt_rst.EOF Then
                    If cmt_rst!amount = 0 Then
                        slSql = "Insert Into VVC_Vendor_Vef (vvcAvfCode,vvcVefCode,vvcExternalID) Values (" & tmVff.iAvfCode & "," & tmVff.iVefCode & ",'" & tmVef.lExtId & "')"
                        If gSQLWaitNoMsgBox(slSql, False) = 0 Then
                            mVendorSetExtID
                            mVendorEnableOptions
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub
