VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ShowFix 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Show/Fix"
   ClientHeight    =   5970
   ClientLeft      =   300
   ClientTop       =   1770
   ClientWidth     =   9390
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
   LinkTopic       =   "DoneMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5970
   ScaleWidth      =   9390
   Begin VB.PictureBox plcDDF 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4680
      Left            =   8265
      ScaleHeight     =   4650
      ScaleWidth      =   8820
      TabIndex        =   0
      Top             =   2025
      Width           =   8850
      Begin VB.DriveListBox cbcDataDrive 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6000
         TabIndex        =   7
         Top             =   285
         Width           =   2565
      End
      Begin VB.DirListBox lbcDataPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   3465
         Left            =   5985
         TabIndex        =   8
         Top             =   765
         Width           =   2565
      End
      Begin VB.ListBox lbcFile 
         Appearance      =   0  'Flat
         Height          =   3930
         Left            =   3165
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   285
         Width           =   2565
      End
      Begin VB.DirListBox lbcPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   3465
         Left            =   315
         TabIndex        =   3
         Top             =   765
         Width           =   2565
      End
      Begin VB.DriveListBox cbcDrive 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   315
         TabIndex        =   2
         Top             =   285
         Width           =   2565
      End
      Begin VB.Label lacPath 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "DDF Path"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1185
         TabIndex        =   1
         Top             =   45
         Width           =   1065
      End
      Begin VB.Label lacLinkFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "DATA Path"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6750
         TabIndex        =   6
         Top             =   45
         Width           =   1290
      End
      Begin VB.Label lacFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "File"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3750
         TabIndex        =   4
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcData 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5205
      Left            =   8475
      ScaleHeight     =   5175
      ScaleWidth      =   9000
      TabIndex        =   51
      Top             =   1755
      Visible         =   0   'False
      Width           =   9030
      Begin VB.TextBox edcValue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1140
         MaxLength       =   10
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   930
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gdcSub 
         Height          =   2190
         Left            =   30
         TabIndex        =   115
         Top             =   2475
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   3863
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin ComctlLib.ProgressBar plcGauge 
         Height          =   225
         Left            =   150
         TabIndex        =   112
         Top             =   4890
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton cmcGDone 
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
         Left            =   5355
         TabIndex        =   54
         Top             =   4770
         Width           =   1050
      End
      Begin VB.CommandButton cmcGCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
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
         Left            =   6495
         TabIndex        =   55
         Top             =   4770
         Width           =   1050
      End
      Begin VB.CommandButton cmcGSave 
         Appearance      =   0  'Flat
         Caption         =   "Sa&ve..."
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
         Left            =   7650
         TabIndex        =   57
         Top             =   4770
         Width           =   1050
      End
      Begin VB.CommandButton cmcSort 
         Appearance      =   0  'Flat
         Caption         =   "&Sort..."
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
         Index           =   0
         Left            =   8025
         TabIndex        =   56
         Top             =   4770
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.VScrollBar vbcData 
         Height          =   2910
         LargeChange     =   17
         Left            =   8745
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   30
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gdcData 
         Height          =   2430
         Left            =   30
         TabIndex        =   114
         Top             =   30
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   4286
         _Version        =   393216
         AllowBigSelection=   0   'False
         ScrollBars      =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.PictureBox pbcTab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   75
         Left            =   8490
         ScaleHeight     =   45
         ScaleWidth      =   0
         TabIndex        =   111
         Top             =   45
         Width           =   15
      End
      Begin VB.PictureBox pbcSTab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   8475
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   109
         Top             =   45
         Width           =   15
      End
      Begin VB.Label plcChg 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4110
         TabIndex        =   117
         Top             =   2445
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lacChg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   2490
         TabIndex        =   107
         Top             =   2220
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label lacTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         TabIndex        =   53
         Top             =   4665
         Width           =   5055
      End
   End
   Begin VB.PictureBox plcLinks 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5280
      Left            =   8625
      ScaleHeight     =   5250
      ScaleWidth      =   9090
      TabIndex        =   9
      Top             =   1530
      Visible         =   0   'False
      Width           =   9120
      Begin VB.CommandButton cmcLink 
         Appearance      =   0  'Flat
         Caption         =   "C&lear All"
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
         Index           =   3
         Left            =   7335
         TabIndex        =   108
         Top             =   4530
         Width           =   1050
      End
      Begin VB.CommandButton cmcLink 
         Appearance      =   0  'Flat
         Caption         =   "Re&move"
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
         Index           =   2
         Left            =   7335
         TabIndex        =   20
         Top             =   4110
         Width           =   1050
      End
      Begin VB.CommandButton cmcLink 
         Appearance      =   0  'Flat
         Caption         =   "&Add"
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
         Index           =   1
         Left            =   7335
         TabIndex        =   19
         Top             =   3705
         Width           =   1050
      End
      Begin VB.CommandButton cmcLink 
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
         Index           =   0
         Left            =   7335
         TabIndex        =   18
         Top             =   3285
         Width           =   1050
      End
      Begin VB.ListBox lbcLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   345
         TabIndex        =   17
         Top             =   3270
         Width           =   6660
      End
      Begin VB.ListBox lbcToIndex 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   3165
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   330
         Width           =   2565
      End
      Begin VB.ListBox lbcToField 
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
         Height          =   1080
         Left            =   3180
         TabIndex        =   16
         Top             =   1605
         Width           =   2565
      End
      Begin VB.ListBox lbcFromField 
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
         Height          =   2550
         Left            =   6000
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   330
         Width           =   2565
      End
      Begin VB.ListBox lbcToFile 
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
         Height          =   2550
         Left            =   330
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   330
         Width           =   2565
      End
      Begin VB.Label lacLinks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "To Field"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   3795
         TabIndex        =   14
         Top             =   90
         Width           =   1305
      End
      Begin VB.Label lacLinks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "To File"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   945
         TabIndex        =   12
         Top             =   75
         Width           =   1305
      End
      Begin VB.Label lacLinks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Fields"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   6015
         TabIndex        =   10
         Top             =   75
         Width           =   2535
      End
   End
   Begin VB.PictureBox plcSelection 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   8700
      ScaleHeight     =   4665
      ScaleWidth      =   8805
      TabIndex        =   21
      Top             =   1200
      Width           =   8835
      Begin VB.PictureBox pbcLbcSelection1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   2100
         Left            =   3315
         ScaleHeight     =   2100
         ScaleWidth      =   5145
         TabIndex        =   119
         Top             =   405
         Width           =   5145
      End
      Begin VB.TextBox edcExample 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Height          =   960
         Left            =   5955
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   61
         Top             =   2835
         Width           =   2535
      End
      Begin VB.CommandButton cmcClearAll 
         Appearance      =   0  'Flat
         Caption         =   "&Clear All"
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
         Left            =   7335
         TabIndex        =   62
         Top             =   4275
         Width           =   1050
      End
      Begin VB.CommandButton cmcLinks 
         Appearance      =   0  'Flat
         Caption         =   "&Links..."
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
         Index           =   0
         Left            =   6135
         TabIndex        =   95
         Top             =   4275
         Width           =   1050
      End
      Begin VB.TextBox edcSelection 
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
         Height          =   1665
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   2835
         Width           =   5475
      End
      Begin VB.ListBox lbcSelection 
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
         Height          =   2130
         Index           =   1
         ItemData        =   "SHOWFIX.frx":0000
         Left            =   3300
         List            =   "SHOWFIX.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   390
         Width           =   5190
      End
      Begin VB.ListBox lbcSelection 
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
         Height          =   2130
         Index           =   0
         Left            =   375
         TabIndex        =   23
         Top             =   390
         Width           =   2580
      End
      Begin VB.Label lacSelMsg 
         Appearance      =   0  'Flat
         Caption         =   "One Operator per Line. Enclose Values within Quotes."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   0
         Left            =   5970
         TabIndex        =   60
         Top             =   3795
         Width           =   2220
      End
      Begin VB.Line lncLine 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   1
         X1              =   2265
         X2              =   4365
         Y1              =   -15
         Y2              =   0
      End
      Begin VB.Label lacSelection 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Operators"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   135
         Width           =   2595
      End
      Begin VB.Label lacSelection 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Fields"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   3285
         TabIndex        =   24
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.PictureBox plcFOpen 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4680
      Left            =   8820
      ScaleHeight     =   4650
      ScaleWidth      =   7965
      TabIndex        =   96
      Top             =   1020
      Visible         =   0   'False
      Width           =   7995
      Begin VB.TextBox edcFDescription 
         Appearance      =   0  'Flat
         Height          =   1620
         Left            =   1395
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   2970
         Width           =   5715
      End
      Begin VB.CommandButton cmcFOpen 
         Appearance      =   0  'Flat
         Caption         =   "&Open"
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
         Left            =   6420
         TabIndex        =   105
         Top             =   525
         Width           =   1050
      End
      Begin VB.CommandButton cmcFCancel 
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
         Left            =   6435
         TabIndex        =   106
         Top             =   1080
         Width           =   1050
      End
      Begin VB.TextBox edcFFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   405
         TabIndex        =   98
         Top             =   525
         Width           =   2505
      End
      Begin VB.FileListBox lbcFFile 
         Appearance      =   0  'Flat
         Height          =   1785
         Left            =   405
         TabIndex        =   99
         Top             =   1005
         Width           =   2505
      End
      Begin VB.DriveListBox cbcFDrive 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3450
         TabIndex        =   101
         Top             =   525
         Width           =   2565
      End
      Begin VB.DirListBox lbcFPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1830
         Left            =   3450
         TabIndex        =   102
         Top             =   1005
         Width           =   2565
      End
      Begin VB.Label lacFOpen 
         Height          =   195
         Left            =   15
         TabIndex        =   118
         Top             =   0
         Width           =   2610
      End
      Begin VB.Label lacFDescription 
         Appearance      =   0  'Flat
         Caption         =   "Description"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   180
         TabIndex        =   103
         Top             =   2970
         Width           =   1080
      End
      Begin VB.Label lacFFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "File Name"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   405
         TabIndex        =   97
         Top             =   285
         Width           =   2490
      End
      Begin VB.Label lacFPath 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "File Path"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3465
         TabIndex        =   100
         Top             =   285
         Width           =   2550
      End
   End
   Begin VB.PictureBox plcSave 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5205
      Left            =   8865
      ScaleHeight     =   5175
      ScaleWidth      =   9075
      TabIndex        =   71
      Top             =   735
      Visible         =   0   'False
      Width           =   9105
      Begin VB.PictureBox plcPAS 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4440
         ScaleHeight     =   225
         ScaleWidth      =   4155
         TabIndex        =   90
         Top             =   4380
         Width           =   4155
         Begin VB.OptionButton rbcPAS 
            Caption         =   "Yes"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2715
            TabIndex        =   91
            Top             =   0
            Width           =   675
         End
         Begin VB.OptionButton rbcPAS 
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3420
            TabIndex        =   92
            Top             =   0
            Width           =   660
         End
      End
      Begin VB.PictureBox plcRun 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         ScaleHeight     =   225
         ScaleWidth      =   6315
         TabIndex        =   78
         Top             =   3930
         Width           =   6315
         Begin VB.OptionButton rbcRun 
            Caption         =   "Allow User Interaction"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3600
            TabIndex        =   80
            Top             =   0
            Width           =   2250
         End
         Begin VB.OptionButton rbcRun 
            Caption         =   "Run Immediately"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1800
            TabIndex        =   79
            Top             =   0
            Width           =   1725
         End
      End
      Begin VB.TextBox edcRecordNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Index           =   1
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   84
         Top             =   3645
         Width           =   825
      End
      Begin VB.TextBox edcRecordNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Index           =   0
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   82
         Top             =   3645
         Width           =   825
      End
      Begin VB.PictureBox plcInclude 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         ScaleHeight     =   225
         ScaleWidth      =   5160
         TabIndex        =   87
         Top             =   4380
         Width           =   5160
         Begin VB.CheckBox cbcInclude 
            Caption         =   "File && Path"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1890
            TabIndex        =   89
            Top             =   0
            Width           =   1260
         End
         Begin VB.CheckBox cbcInclude 
            Caption         =   "Selective"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   735
            TabIndex        =   88
            Top             =   0
            Width           =   1125
         End
      End
      Begin VB.TextBox edcWrap 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   2730
         MaxLength       =   4
         TabIndex        =   86
         Top             =   4005
         Width           =   720
      End
      Begin VB.CommandButton cmcSCancel 
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
         Left            =   4590
         TabIndex        =   94
         Top             =   4740
         Width           =   1050
      End
      Begin VB.CommandButton cmcSSave 
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
         Left            =   3240
         TabIndex        =   93
         Top             =   4740
         Width           =   1050
      End
      Begin VB.TextBox edcDescription 
         Appearance      =   0  'Flat
         Height          =   2490
         Left            =   1320
         TabIndex        =   77
         Top             =   1050
         Width           =   7185
      End
      Begin VB.PictureBox plcSaveTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1335
         ScaleHeight     =   345
         ScaleWidth      =   5745
         TabIndex        =   73
         Top             =   495
         Width           =   5775
         Begin VB.TextBox edcSaveTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   45
            TabIndex        =   74
            Top             =   90
            Width           =   5655
         End
      End
      Begin VB.CommandButton cmcSBrowse 
         Appearance      =   0  'Flat
         Caption         =   "&Browse..."
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
         Left            =   7425
         TabIndex        =   75
         Top             =   510
         Width           =   1050
      End
      Begin VB.Label lacSave 
         Caption         =   "Save Grid Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   116
         Top             =   45
         Width           =   2085
      End
      Begin VB.Label lacRecordNo 
         Appearance      =   0  'Flat
         Caption         =   "End Display #"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   2415
         TabIndex        =   83
         Top             =   3690
         Width           =   1200
      End
      Begin VB.Label lacRecordNo 
         Appearance      =   0  'Flat
         Caption         =   "Start Display #"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   81
         Top             =   3690
         Width           =   1380
      End
      Begin VB.Label lacWrap 
         Appearance      =   0  'Flat
         Caption         =   "Max # of Characters per Line"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   105
         TabIndex        =   85
         Top             =   4050
         Width           =   2610
      End
      Begin VB.Label lacDescription 
         Appearance      =   0  'Flat
         Caption         =   "Description"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   105
         TabIndex        =   76
         Top             =   1050
         Width           =   1080
      End
      Begin VB.Label lacSaveTo 
         Appearance      =   0  'Flat
         Caption         =   "Save To"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   105
         TabIndex        =   72
         Top             =   555
         Width           =   810
      End
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   45
      Top             =   3315
   End
   Begin VB.Timer tmcScroll 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   45
      Top             =   2730
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   585
      Top             =   5535
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   7785
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   65
      Top             =   5460
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
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
         Height          =   1440
         Left            =   45
         Picture         =   "SHOWFIX.frx":0004
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   69
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
         TabIndex        =   67
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
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   70
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.TextBox edcField 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1305
      MaxLength       =   10
      TabIndex        =   64
      Top             =   1125
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmcField 
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
      Left            =   2250
      Picture         =   "SHOWFIX.frx":2E1E
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   1095
      Visible         =   0   'False
      Width           =   195
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
      Left            =   3510
      TabIndex        =   58
      Top             =   5565
      Width           =   1050
   End
   Begin VB.CommandButton cmcProcess 
      Appearance      =   0  'Flat
      Caption         =   "&Process"
      Enabled         =   0   'False
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
      Left            =   4875
      TabIndex        =   59
      Top             =   5565
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   45
      Top             =   4575
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".Txt"
      Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
      FilterIndex     =   1
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   45
      Top             =   3840
   End
   Begin VB.PictureBox plcEditReplace 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   9060
      ScaleHeight     =   4665
      ScaleWidth      =   8805
      TabIndex        =   27
      Top             =   540
      Width           =   8835
      Begin VB.OptionButton rbcEditReplace 
         Caption         =   "Mass Removal"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3645
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   45
         Width           =   1620
      End
      Begin VB.TextBox edcSearchTo 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   6165
         TabIndex        =   37
         Top             =   300
         Width           =   1305
      End
      Begin VB.TextBox edcSearchFrom 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   4110
         TabIndex        =   35
         Top             =   300
         Width           =   1305
      End
      Begin VB.TextBox edcSearchKey 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   33
         Text            =   "0"
         Top             =   300
         Width           =   330
      End
      Begin VB.PictureBox pbcLbcEditReplace2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   1470
         Left            =   150
         ScaleHeight     =   1470
         ScaleWidth      =   8535
         TabIndex        =   121
         Top             =   2595
         Width           =   8535
      End
      Begin VB.PictureBox pbcLbcEditReplace0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   1455
         Left            =   150
         ScaleHeight     =   1455
         ScaleWidth      =   5790
         TabIndex        =   120
         Top             =   765
         Width           =   5790
      End
      Begin VB.CommandButton cmcSort 
         Appearance      =   0  'Flat
         Caption         =   "&Sort..."
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
         Index           =   2
         Left            =   8160
         TabIndex        =   49
         Top             =   4290
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmcLinks 
         Appearance      =   0  'Flat
         Caption         =   "&Links..."
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
         Index           =   1
         Left            =   6450
         TabIndex        =   50
         Top             =   4305
         Width           =   1050
      End
      Begin VB.OptionButton rbcEditReplace 
         Caption         =   "Fix by Mass Substitution"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1125
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   45
         Width           =   2400
      End
      Begin VB.CommandButton cmcEditReplace 
         Appearance      =   0  'Flat
         Caption         =   "Add Al&l "
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
         Index           =   1
         Left            =   2595
         TabIndex        =   46
         Top             =   4305
         Width           =   1050
      End
      Begin VB.CommandButton cmcEditReplace 
         Appearance      =   0  'Flat
         Caption         =   "&Clear All"
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
         Index           =   3
         Left            =   5160
         TabIndex        =   48
         Top             =   4305
         Width           =   1050
      End
      Begin VB.CommandButton cmcEditReplace 
         Appearance      =   0  'Flat
         Caption         =   "Re&move"
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
         Index           =   2
         Left            =   3900
         TabIndex        =   47
         Top             =   4305
         Width           =   1050
      End
      Begin VB.CommandButton cmcEditReplace 
         Appearance      =   0  'Flat
         Caption         =   "&Add"
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
         Index           =   0
         Left            =   1350
         TabIndex        =   45
         Top             =   4305
         Width           =   1050
      End
      Begin VB.ListBox lbcEditReplace 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   2
         Left            =   135
         TabIndex        =   43
         Top             =   2580
         Width           =   8565
      End
      Begin VB.TextBox edcReplaceValue 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6135
         TabIndex        =   41
         Top             =   780
         Width           =   2565
      End
      Begin VB.OptionButton rbcEditReplace 
         Caption         =   "Fix by Typing"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   6585
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   45
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.OptionButton rbcEditReplace 
         Caption         =   "Show"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   28
         Top             =   45
         Width           =   810
      End
      Begin VB.ListBox lbcEditReplace 
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
         Height          =   870
         Index           =   1
         Left            =   6135
         TabIndex        =   44
         Top             =   1350
         Width           =   2565
      End
      Begin VB.ListBox lbcEditReplace 
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
         Index           =   0
         Left            =   135
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   750
         Width           =   5820
      End
      Begin VB.Label lacSearchTo 
         Appearance      =   0  'Flat
         Caption         =   "To:"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5700
         TabIndex        =   36
         Top             =   300
         Width           =   375
      End
      Begin VB.Label lacSearchFrom 
         Appearance      =   0  'Flat
         Caption         =   "From "
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1905
         TabIndex        =   34
         Top             =   300
         Width           =   2115
      End
      Begin VB.Label lacSearchKey 
         Appearance      =   0  'Flat
         Caption         =   "Search Key #"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   210
         TabIndex        =   32
         Top             =   300
         Width           =   1095
      End
      Begin VB.Image imcEditReplace 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   0
         Left            =   150
         Top             =   2355
         Width           =   8535
      End
      Begin VB.Image imcEditReplace 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   1
         Left            =   165
         Top             =   4080
         Width           =   8520
      End
      Begin VB.Label lacEditReplace 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Replace Value or Field Name"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   6135
         TabIndex        =   40
         Top             =   555
         Width           =   2550
      End
      Begin VB.Label lacEditReplace 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Conversion Data Type"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   6135
         TabIndex        =   42
         Top             =   1125
         Width           =   2550
      End
      Begin VB.Label lacEditReplace 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Fields"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   435
         TabIndex        =   38
         Top             =   540
         Width           =   4980
      End
   End
   Begin ComctlLib.TabStrip plcTabSelection 
      Height          =   5205
      Left            =   180
      TabIndex        =   113
      Top             =   45
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   9181
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "DDF P&ath and Files"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "S&elections"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "S&how or Fix Fields"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lacMode 
      Height          =   255
      Left            =   210
      TabIndex        =   122
      Top             =   5400
      Width           =   3105
   End
   Begin VB.Image imcIconInsert 
      Appearance      =   0  'Flat
      DragIcon        =   "SHOWFIX.frx":2F18
      Height          =   180
      Left            =   1245
      Top             =   5745
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcIconDwnArrow 
      Appearance      =   0  'Flat
      DragIcon        =   "SHOWFIX.frx":3222
      Height          =   195
      Left            =   1395
      Top             =   5565
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imcIconDrag 
      Appearance      =   0  'Flat
      DragIcon        =   "SHOWFIX.frx":352C
      Height          =   180
      Left            =   1110
      Top             =   5565
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imcTrashClosed 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   6240
      Picture         =   "SHOWFIX.frx":3836
      Top             =   5430
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcTrashOpened 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   7020
      Picture         =   "SHOWFIX.frx":3B40
      Top             =   5415
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8805
      Picture         =   "SHOWFIX.frx":3E4A
      Top             =   5430
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Spec..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Spec"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save Spec &As..."
      End
      Begin VB.Menu mnuFileSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFilePSetup 
         Caption         =   "P&rinter Setup..."
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Ex&it"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On ..."
      End
      Begin VB.Menu mnuHelpSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About ..."
      End
   End
End
Attribute VB_Name = "ShowFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ShowFix.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Show/Fix screen code
Option Explicit
Option Compare Text
Dim imIgnoreChg As Integer
Dim smDBPath As String
Dim smDDFPath As String
Dim smExportPath As String
Dim smImportPath As String
Dim lmTotalNoRec As Long
Dim lmProcessedNoRec As Long
Dim hmDB As Integer
Dim imEditCol As Integer
Dim imEditRow As Integer
Dim imMouseUpX As Integer
Dim imRecCls As Integer
Dim hmTo As Integer   'To file hanle
Dim hmFrom As Integer   'From file hanle
Dim smCmmdFileName As String
Dim imPasswordAsked As Integer
Dim smSpecDescript As String
Dim hmFile As Integer
Dim imFileID As Integer
Dim lmActualRecLen As Long
Dim imBaseRecLen As Integer
Dim imLVarLengthError As Integer

Dim imListFieldR0(1 To 5) As Integer
Dim imListFieldR2(1 To 6) As Integer
Dim imListFieldS(1 To 5) As Integer

Dim tmRec As BASEREC
Dim tmRecInfo() As RECINFO
Dim tmRowLinkList() As ROWLINKLIST
Dim tmGrid1Data() As GRID1DATA
Const GRID1SIZE = 10    '8
Dim tmGrid2Data() As GRID2DATA
Const GRID2SIZE = 16
Dim tmGrid3Data() As GRID3DATA
Const GRID3SIZE = 128
Dim tmGrid4Data() As GRID4DATA
Const GRID4SIZE = 1024
Dim tmGrid5Data() As GRID5DATA
Const GRID5SIZE = 4096
Dim tmGrid6Data() As GRID6DATA
Const GRID6SIZE = 32000
Dim tmRecSel() As SELCOMP
Dim smFieldValue As String
Dim smLogicSel As String
Dim tmLink() As LINK
Dim tmKeyFile() As KEYFILE
Dim tmDDFFile() As DDFFILE    'DDF File- File Names
Dim tmDDFField() As DDFFIELD    'DDF File- Field Names
Dim tmDDFLinkField() As DDFFIELD    'DDF File- Field Names
Dim tmFieldSrchKey1 As DDFFIELD1
Dim tmDDFLinkIndex() As DDFINDEX    'DDF File- Index Names
Dim tmIndexSrchKey0 As DDFINDEX0
Dim tmSearchDDFField As DDFFIELD
Dim tmSearchDDFIndex As DDFINDEX
Dim smHeading As String
Dim imExporting As Integer
Dim imTerminate As Integer
Dim imTabSelection As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imReadingFile As Integer
Dim imKeyDownRow As Integer
Dim imFieldSizeCount(1 To 6) As Integer
Dim imVarLenRec As Integer
Dim imVarType As Integer    '12 or 13
Dim imVarBaseRecLen As Integer  'Record length without variable string
Dim lmMaxNoRec As Long
Dim imVBCAdjust As Integer
Dim imEnableOnCancel As Integer 'Enable cmcDone and cmcProcess on Cancel
Dim imNoRowLbc As Integer   'Number of row in lbcEditReplace(2)
Dim fmBoxGridH As Single
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
Dim imDragScroll As Integer 'Object to be scrolled (same values as below)
Dim imScrollDir As Integer  '0=Page up; 1=Page down
Const DRAGFIELD = 1
Const DRAGEDIT = 2
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imRecUpper1 As Integer
Dim imRecUpper2 As Integer
Dim imRecMaxUpper1 As Integer
Dim imRecMaxUpper2 As Integer
'Dim imRowUpper1 As Integer
'Dim imRowUpper2 As Integer
Dim imGrid1Upper1 As Integer
Dim imGrid1Upper2 As Integer
Dim imGrid2Upper1 As Integer
Dim imGrid2Upper2 As Integer
Dim imGrid3Upper1 As Integer
Dim imGrid3Upper2 As Integer
Dim imGrid4Upper1 As Integer
Dim imGrid4Upper2 As Integer
Dim imGrid5Upper1 As Integer
Dim imGrid5Upper2 As Integer
Dim imGrid6Upper1 As Integer
Dim imGrid6Upper2 As Integer
Dim imClsGrid As Integer
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmBBSpot As BBSPOTSS
Dim tmProgTest As PROGRAMSS
Dim tmAvailTest As AVAILSS
Dim tmSpotTest As CSPOTSS
Dim tmBBSpotTest As BBSPOTSS
Dim lmPercent As Long
Dim tmGetDDFIndex As DDFINDEX
Dim tmGetDDFFile As DDFFILE
Dim tmGetDDFFIELD As DDFFIELD

Public Sub mGetFieldFields()
    Dim tlGDDFField As DDFFIELD
    
    tlGDDFField.iOffset = 0
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField, 0
    tmGetDDFFIELD.iFieldID = Val(smFieldValue)
    tlGDDFField.iOffset = 2
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField, 0
    tmGetDDFFIELD.iFileID = Val(smFieldValue)
    tlGDDFField.iOffset = 4
    tlGDDFField.iSize = 20
    tlGDDFField.sDataType = Chr(0)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField, 0
    tmGetDDFFIELD.sName = smFieldValue
    'tlGDDFField.iOffset = 24
    'tlGDDFField.iSize = 1
    'tlGDDFField.sDataType = Chr(1)
    'tlGDDFField.sDec = ""
    'mGetField tlGDDFField, 0
    'tmGetDDFFIELD.sDataType = smFieldValue
    tmGetDDFFIELD.sDataType = Chr(tmRec.sChar(25))
    tlGDDFField.iOffset = 25
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField, 0
    tmGetDDFFIELD.iOffset = Val(smFieldValue)
    tlGDDFField.iOffset = 27
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField, 0
    tmGetDDFFIELD.iSize = Val(smFieldValue)
    tlGDDFField.iOffset = 29
    tlGDDFField.iSize = 1
    tlGDDFField.sDataType = Chr(0)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField, 0
    tmGetDDFFIELD.sDec = smFieldValue
    tlGDDFField.iOffset = 30
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField, 0
    tmGetDDFFIELD.iFlags = Val(smFieldValue)
End Sub

Public Sub mGetIndexFields()
    Dim tlDDFField As DDFFIELD
    
    tlDDFField.iOffset = 0
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField, 0
    tmGetDDFIndex.iFileID = Val(smFieldValue)
    tlDDFField.iOffset = 2
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField, 0
    tmGetDDFIndex.iFieldID = Val(smFieldValue)
    tlDDFField.iOffset = 4
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField, 0
    tmGetDDFIndex.iNumber = Val(smFieldValue)
    tlDDFField.iOffset = 6
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField, 0
    tmGetDDFIndex.iPart = Val(smFieldValue)
    tlDDFField.iOffset = 8
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField, 0
    tmGetDDFIndex.iFlag = Val(smFieldValue)
End Sub

Private Sub cbcDataDrive_Change()
    'lbcFromField.Clear
    'lbcSelection(1).Clear
    'edcSelection.Text = ""
    'lbcEditReplace(0).Clear
    'lbcEditReplace(2).Clear
    'lbcToIndex.Clear
    'lbcToField.Clear
    'lbcLink.Clear
    Screen.MousePointer = vbHourglass
    lbcDataPath.Path = cbcDataDrive.Drive
    Screen.MousePointer = vbDefault
End Sub
Private Sub cbcDrive_Change()
    Screen.MousePointer = vbHourglass
    lbcPath.Path = cbcDrive.Drive
    'cbcDataDrive.Drive = cbcDrive.Drive
    lbcFile.Clear
    lbcToFile.Clear
    lbcFromField.Clear
    lbcSelection(1).Clear
    pbcLbcSelection1_Paint
    edcSelection.Text = ""
    lbcEditReplace(0).Clear
    lbcEditReplace(2).Clear
    lbcEditReplace(2).AddItem " "
    pbcLbcEditReplace0_Paint
    pbcLbcEditReplace2_Paint
    lbcToIndex.Clear
    lbcToField.Clear
    lbcLink.Clear
    ReDim tmLink(0 To 0) As LINK
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub
Private Sub cbcFDrive_Change()
    Screen.MousePointer = vbHourglass
    lbcFPath.Path = cbcFDrive.Drive
    lbcFFile.Clear
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcField.SelStart = 0
    edcField.SelLength = Len(edcField.Text)
    edcField.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcField.SelStart = 0
    edcField.SelLength = Len(edcField.Text)
    edcField.SetFocus
End Sub
Private Sub cmcClearAll_Click()
    edcSelection.Text = ""
End Sub
Private Sub cmcDone_Click()
    mTerminate
End Sub
Private Sub cmcDone_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcEditReplace_Click(Index As Integer)
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slFieldInfo As String
    Dim slFileName As String
    Dim slFieldName As String
    Dim slFieldDataType As String
    Dim slNewDataType As String
    Dim ilField As Integer
    Dim ilFound As Integer
    Dim ilDataOK As Integer
    Dim ilDataType As Integer
    Dim slValue As String
    Dim slSvValue As String
    Dim ilNoDecPt As Integer
    Dim ilSize As Integer
    Dim ilFileID As Integer
    Dim ilPos As Integer
    Dim slPrimaryName As String
    Dim ilPrimaryFileID As Integer
    
    If Index = 0 Then   'Add
        If lbcEditReplace(0).ListIndex >= 0 Then
            If imDragIndexDest < 0 Then
                imDragIndexDest = lbcEditReplace(2).ListCount - 1
            End If
            If rbcEditReplace(0).Value Or rbcEditReplace(1).Value Then
                lbcEditReplace(2).AddItem lbcEditReplace(0).List(lbcEditReplace(0).ListIndex), imDragIndexDest
            Else
                slValue = Trim$(edcReplaceValue.Text)
                If (Left$(slValue, 1) = """") Or (Left$(slValue, 1) = "'") Then
                    slValue = Mid$(slValue, 2)
                End If
                If (right$(slValue, 1) = """") Or (right$(slValue, 1) = "'") Then
                    slValue = Left$(slValue, Len(slValue) - 1)
                End If
                If Len(slValue) = 0 Then
                    Beep
                    edcReplaceValue.SetFocus
                    Exit Sub
                Else
                    slFieldInfo = lbcEditReplace(0).List(lbcEditReplace(0).ListIndex)
                    ilRet = gParseItem(slFieldInfo, 1, "|", slFileName)
                    ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
                    ilRet = gParseItem(slFieldInfo, 3, "|", slFieldDataType)
                    ilRet = gParseItem(slFieldInfo, 4, "| ", slNewDataType)
                    If ilRet <> 0 Then
                        slNewDataType = "" 'Leave with one blank so display works
                    End If
                    slPrimaryName = lbcFile.List(lbcFile.ListIndex)
                    ilPrimaryFileID = -1
                    For ilLoop = 0 To UBound(tmDDFFile) - 1 Step 1
                        If StrComp(Trim$(tmDDFFile(ilLoop).sName), slPrimaryName, 1) = 0 Then
                            ilPrimaryFileID = tmDDFFile(ilLoop).iFileID
                            Exit For
                        End If
                    Next ilLoop
                    'Find file ID
                    ilFileID = mGetFileID(slFileName)
                    If ilPrimaryFileID <> ilFileID Then
                        Beep
                        MsgBox "Can only replace fields from file " & Trim$(lbcFile.List(lbcFile.ListIndex)), vbOkOnly + vbExclamation + vbApplicationModal, "Search Key Error"
                        lbcEditReplace(0).ListIndex = -1
                        Exit Sub
                    End If
                    ilFound = False
                    For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                        If (StrComp(Trim$(slFieldName), Trim$(tmDDFField(ilField).sName), 1) = 0) And (ilFileID = tmDDFField(ilField).iFileID) Then
                            ilFound = True
                            ilSize = tmDDFField(ilField).iSize
                            ilDataType = Asc(tmDDFField(ilField).sDataType)
                            If slNewDataType = lbcEditReplace(1).List(1) Then
                                ilDataType = 100
                            ElseIf slNewDataType = lbcEditReplace(1).List(2) Then
                                ilDataType = 101
                            End If
                        End If
                    Next ilField
                    If Not ilFound Then
                        Beep
                        lbcEditReplace(0).SetFocus
                        Exit Sub
                    End If
                    ilFound = False
                    slSvValue = slValue
                    If mTestOper(ilDataType) Then
                        If InStr(1, slValue, "+", 1) > 0 Then
                            ilPos = InStr(1, slValue, "+", 1)
                            slValue = Left$(slValue, ilPos - 1)
                        ElseIf InStr(1, slValue, "-", 1) > 1 Then
                            ilPos = InStr(1, slValue, "-", 1)
                            slValue = Left$(slValue, ilPos - 1)
                        ElseIf InStr(1, slValue, "*", 1) > 0 Then
                            ilPos = InStr(1, slValue, "*", 1)
                            slValue = Left$(slValue, ilPos - 1)
                        ElseIf InStr(1, slValue, "/", 1) > 0 Then
                            ilPos = InStr(1, slValue, "/", 1)
                            slValue = Left$(slValue, ilPos - 1)
                        ElseIf InStr(1, slValue, "\", 1) > 0 Then
                            ilPos = InStr(1, slValue, "\", 1)
                            slValue = Left$(slValue, ilPos - 1)
                        End If
                    End If
                    For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                        If (StrComp(Trim$(slValue), Trim$(tmDDFField(ilField).sName), 1) = 0) And (ilFileID = tmDDFField(ilField).iFileID) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilField
                    If Not ilFound Then
                        slValue = slSvValue
                        If Not mCheckValue(slValue, ilDataType, ilSize) Then
                            Beep
                            edcReplaceValue.SetFocus
                            Exit Sub
                        End If
                    End If
                    If slNewDataType = "" Then 'Leave with one blank so display works
                        slNewDataType = " "
                    End If
                    lbcEditReplace(2).AddItem slFileName & "|" & slFieldName & "|" & slFieldDataType & "|" & slNewDataType & "|" & Trim$(edcReplaceValue.Text), imDragIndexDest
                End If
            End If
            lbcEditReplace(0).ListIndex = -1
        End If
    ElseIf Index = 1 Then 'Add All
        lbcEditReplace(2).Clear
        For ilLoop = 0 To lbcEditReplace(0).ListCount - 1 Step 1
            lbcEditReplace(2).AddItem lbcEditReplace(0).List(ilLoop)
        Next ilLoop
        lbcEditReplace(2).AddItem " "
        lbcEditReplace(0).ListIndex = -1
    ElseIf Index = 2 Then   'Remove
        If lbcEditReplace(2).ListIndex >= 0 Then
            lbcEditReplace(2).RemoveItem lbcEditReplace(2).ListIndex
        End If
    ElseIf Index = 3 Then   'Clear All
        lbcEditReplace(2).Clear
        lbcEditReplace(2).AddItem " "
    End If
    pbcLbcEditReplace0_Paint
    pbcLbcEditReplace2_Paint
    imDragIndexDest = -1
    mSetCommands
End Sub
Private Sub cmcEditReplace_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cmcEditReplace_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcFCancel_Click()
    cmcDone.Enabled = True
    plcFOpen.Visible = False
    mSetCommands
End Sub
Private Sub cmcField_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcField.SelStart = 0
    edcField.SelLength = Len(edcField.Text)
    edcField.SetFocus
End Sub
Private Sub cmcFOpen_Click()
    Dim ilPos As Integer
    Dim slName As String
    Dim slFromFile As String
    Dim ilRet As Integer
    slName = Trim$(edcFFile.Text)
    If Len(slName) <= 0 Then
        edcFFile.SetFocus
        Exit Sub
    End If
    ilPos = InStr(slName, "*")
    If ilPos > 0 Then
        lbcFFile.FileName = lbcFPath.Path & "\" & slName
        edcFFile.SetFocus
        Exit Sub
    End If
    ilPos = InStr(slName, "?")
    If ilPos > 0 Then
        lbcFFile.FileName = lbcFPath.Path & "\" & slName
        edcFFile.SetFocus
        Exit Sub
    End If
    'Read file
    slFromFile = slName
    If InStr(slFromFile, ":") = 0 Then
        slFromFile = lbcFPath.Path & "\" & slFromFile
    End If
    ilRet = 0
    On Error GoTo cmcFOpenErr:
    hmFrom = FreeFile
    Open slFromFile For Input As hmFrom
    If ilRet <> 0 Then
        If InStr(slFromFile, ".") = 0 Then
            ilRet = 0
            slFromFile = slFromFile & ".sfi"
            Open slFromFile For Input As hmFrom
            If ilRet <> 0 Then
                MsgBox "Open " & slFromFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                edcFFile.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Open " & slFromFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            edcFFile.SetFocus
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    mReadSpec False
    Close hmFrom
    smCmmdFileName = slName
    Screen.MousePointer = vbDefault
    cmcFCancel_Click
    On Error GoTo 0
    Exit Sub
cmcFOpenErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub cmcGCancel_Click()
    edcValue.Visible = False
    imEditRow = -1
    imEditCol = -1
    If imReadingFile Then
        imTerminate = True
        Exit Sub
    End If
    mnuFileOpen.Enabled = True
    mnuFilePrint.Enabled = False
    'plcTabSelection(0).ZOrder vbBringToFront
    'plcTabSelection.SelectedItem.Index = 1
    gdcData.Rows = 2
    gdcData.Cols = 2
    gdcData.Clear
    gdcSub.Rows = 2
    gdcSub.Cols = 2
    gdcSub.Clear
    
    'SendKeys "%a"
    plcTabSelection_Click
    'plcSelection.Visible = False
    'plcLinks.Visible = False
    'plcEditReplace.Visible = False
    plcData.Visible = False
    'plcDDF.Visible = True
    'plcDDF.ZOrder vbBringToFront
    cmcDone.Enabled = True
    cmcProcess.Enabled = True
    mSetCommands
End Sub
Private Sub cmcGDone_Click()
    mUpdateField
    imReadingFile = True    'Set to avoid error in gdcData
'    gdcData.Rows = 2
'    gdcData.Cols = 2
'    gdcData.Col = 0
'    gdcData.Row = 0
'    gdcData.Text = ""
'    gdcData.Col = 1
'    gdcData.Row = 0
'    gdcData.Text = ""
'    gdcData.Col = 0
'    gdcData.Row = 1
'    gdcData.Text = ""
'    gdcData.Col = 1
'    gdcData.Row = 1
'    gdcData.Text = ""
    mnuFilePrint.Enabled = False
    mnuFileOpen.Enabled = True
    'plcTabSelection(0).ZOrder vbBringToFront
    'plcTabSelection.SelectedItem.Index = 1
    gdcData.Rows = 2
    gdcData.Cols = 2
    gdcData.Clear
    gdcSub.Rows = 2
    gdcSub.Cols = 2
    gdcSub.Clear
    'SendKeys "%a"
    plcTabSelection_Click
    'plcSelection.Visible = False
    'plcLinks.Visible = False
    'plcEditReplace.Visible = False
    plcData.Visible = False
    'plcDDF.Visible = True
    'plcDDF.ZOrder vbBringToFront
    cmcDone.Enabled = True
    cmcProcess.Enabled = True
    mSetCommands
    imReadingFile = False
End Sub
Private Sub cmcGSave_Click()
    mnuFilePrint.Enabled = False
    cmcDone.Enabled = False
    cmcProcess.Enabled = False
    imEnableOnCancel = False
    lacSave.Caption = "Save Grid Data"
    plcRun.Visible = False
    lacSaveTo.Visible = True
    plcSaveTo.Visible = True
    cmcSBrowse.Visible = True
    cmcSSave.Caption = "&Save"
    lacDescription.Top = lacSaveTo.Top + 505
    edcDescription.Top = plcSaveTo.Top + 505
    lacRecordNo(0).Visible = True
    lacRecordNo(1).Visible = True
    edcRecordNo(0).Visible = True
    edcRecordNo(1).Visible = True
    cbcInclude(0).Value = vbChecked
    cbcInclude(1).Value = vbChecked
    edcRecordNo(0).Text = "1"
    edcRecordNo(1).Text = Trim$(Str$(lmMaxNoRec))
    lacWrap.Visible = True
    edcWrap.Visible = True
    edcWrap.Text = ""
    If (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary") Then
        plcPAS.Visible = True
        rbcPAS(0).Value = True
    Else
        plcPAS.Visible = False
    End If
    plcInclude.Visible = True
    plcSave.Visible = True
    plcSave.ZOrder vbBringToFront
End Sub
Private Sub cmcLink_Click(Index As Integer)
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slToFile As String
    Dim slFromFile As String
    Dim slName As String
    Dim ilToDataType As Integer
    Dim ilFromDataType As Integer
    Dim ilToSize As Integer
    Dim ilFromSize As Integer
    Dim slIndex As String
    Dim ilKeyNo As Integer
    Dim ilToIndex As Integer
    Dim ilFromIndex As Integer
    Dim ilListIndex As Integer
    Select Case Index
        Case 0  'Done
            'Add fields for files linked into lbcSelection(1) and
            'lbcEditReplace(0) if rbcEditReplace(0) & (1) = True
            mPopFieldBoxes
            cmcDone.Enabled = True
            cmcProcess.Enabled = True
            plcLinks.Visible = False
        Case 1  'Add
            If (lbcToField.ListIndex < 0) Or (lbcFromField.ListIndex < 0) Then
                Beep
                Exit Sub
            End If
            slName = lbcToIndex.List(lbcToIndex.ListIndex)
            ilPos = InStr(slName, ":")
            If ilPos = 0 Then
                Beep
                Exit Sub
            End If
            slIndex = Mid$(slName, 7, ilPos - 7)
            ilKeyNo = Val(slIndex)
            'Form link file name: field links from file Name: field name
            slName = lbcToFile.List(lbcToFile.ListIndex)
            For ilLoop = 0 To UBound(tmDDFFile) - 1 Step 1
                If StrComp(Trim$(tmDDFFile(ilLoop).sName), slName, 1) = 0 Then
                    slToFile = Trim$(gRemoveZero(tmDDFFile(ilLoop).sLocation))
                    ilPos = InStr(slToFile, ":")
                    If ilPos > 0 Then
                        slToFile = Mid$(slToFile, ilPos + 1)
                    End If
                    ilPos = InStr(slToFile, "\")
                    Do While ilPos > 0
                        slToFile = Mid$(slToFile, ilPos + 1)
                        ilPos = InStr(slToFile, "\")
                    Loop
                    Exit For
                End If
            Next ilLoop
            For ilLoop = 0 To UBound(tmDDFLinkField) - 1 Step 1
                If Trim$(tmDDFLinkField(ilLoop).sName) = lbcToField.List(lbcToField.ListIndex) Then
                    ilToDataType = Asc(tmDDFLinkField(ilLoop).sDataType)
                    ilToSize = tmDDFLinkField(ilLoop).iSize
                    ilToIndex = ilLoop
                    Exit For
                End If
            Next ilLoop
            slName = lbcFile.List(lbcFile.ListIndex)
            For ilLoop = 0 To UBound(tmDDFFile) - 1 Step 1
                If StrComp(Trim$(tmDDFFile(ilLoop).sName), slName, 1) = 0 Then
                    slFromFile = Trim$(gRemoveZero(tmDDFFile(ilLoop).sLocation))
                    ilPos = InStr(slFromFile, ":")
                    If ilPos > 0 Then
                        slFromFile = Mid$(slFromFile, ilPos + 1)
                    End If
                    ilPos = InStr(slFromFile, "\")
                    Do While ilPos > 0
                        slFromFile = Mid$(slFromFile, ilPos + 1)
                        ilPos = InStr(slFromFile, "\")
                    Loop
                    Exit For
                End If
            Next ilLoop
            For ilLoop = 0 To UBound(tmDDFField) - 1 Step 1
                If Trim$(tmDDFField(ilLoop).sName) = lbcFromField.List(lbcFromField.ListIndex) Then
                    ilFromDataType = Asc(tmDDFField(ilLoop).sDataType)
                    ilFromSize = tmDDFField(ilLoop).iSize
                    ilFromIndex = ilLoop
                    Exit For
                End If
            Next ilLoop
            If ilToDataType <> ilFromDataType Then
                If ((ilToDataType = 1) Or (ilToDataType = 15)) And ((ilFromDataType = 1) Or (ilFromDataType = 15)) Then
                    If ilToSize <> ilFromSize Then
                        Beep
                        Exit Sub
                    End If
                Else
                    Beep
                    Exit Sub
                End If
            End If
            lbcLink.AddItem slToFile & ":" & lbcToField.List(lbcToField.ListIndex) & " <<== " & slFromFile & ":" & lbcFromField.List(lbcFromField.ListIndex)
            tmLink(UBound(tmLink)).iKeyNo = ilKeyNo
            tmLink(UBound(tmLink)).iToFileID = tmDDFLinkField(ilToIndex).iFileID
            tmLink(UBound(tmLink)).iToFieldID = tmDDFLinkField(ilToIndex).iFieldID
            tmLink(UBound(tmLink)).iKeyOffset = 0
            For ilLoop = 0 To ilToIndex - 1 Step 1
                tmLink(UBound(tmLink)).iKeyOffset = tmLink(UBound(tmLink)).iKeyOffset + mFieldSize(tmDDFLinkField(ilToIndex))
            Next ilLoop
            tmLink(UBound(tmLink)).iFromFileID = tmDDFField(ilFromIndex).iFileID
            tmLink(UBound(tmLink)).iFromFieldID = tmDDFField(ilFromIndex).iFieldID
            tmLink(UBound(tmLink)).iFromOffset = tmDDFField(ilFromIndex).iOffset
            ReDim Preserve tmLink(0 To UBound(tmLink) + 1) As LINK
        Case 2  'Remove
            If lbcLink.ListIndex >= 0 Then
                ilListIndex = lbcLink.ListIndex
                lbcLink.RemoveItem lbcLink.ListIndex
                If ilListIndex > 0 Then
                    For ilLoop = UBound(tmLink) - 1 To ilListIndex Step -1
                        tmLink(ilLoop - 1) = tmLink(ilLoop)
                    Next ilLoop
                End If
                ReDim Preserve tmLink(0 To UBound(tmLink) - 1) As LINK
            End If
        Case 3  'Clear All
            lbcLink.Clear
            ReDim tmLink(0 To 0) As LINK
    End Select
    mSetLinkCommands
End Sub
Private Sub cmcLinks_Click(Index As Integer)
    cmcDone.Enabled = False
    cmcProcess.Enabled = False
    plcLinks.Visible = True
    plcLinks.ZOrder vbBringToFront
End Sub
Private Sub cmcLinks_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cmcLinks_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcProcess_Click()
    Dim ilNoRows As Integer
    Dim ilNoCols As Integer
    Dim ilLoop As Integer
    Dim ilCol As Integer
    Dim ilPos As Integer
    Dim slFieldName As String
    Dim slFieldInfo As String
    Dim ilRet As Integer
    Dim slStr As String
    
    slStr = Trim$(lacSearchFrom.Caption)
    If InStr(slStr, "From") = 1 Then
        Beep
        MsgBox "Invalid Search Key", vbOkOnly + vbExclamation + vbApplicationModal, "Search Key Error"
        Exit Sub
    End If
    If Not mCheckSearchValues() Then
        Beep
        MsgBox "Invalid Search Value", vbOkOnly + vbExclamation + vbApplicationModal, "Search Value Error"
        Exit Sub
    End If
    If Not mMakeRecSel() Then
        ''plcTabSelection_Click 1
        ''plcTabSelection.SelectedItem.Index = 2
        'SendKeys "%e"
        mnuFilePrint.Enabled = False
        Exit Sub
    End If
    
    If rbcEditReplace(2).Value Then
        ilRet = MsgBox("Backup of database must be done before continuing with Mass Substitution as the operation can not be reversed, has it been done", vbYesNo + vbQuestion, "Mass Replace")
        If ilRet = vbNo Then
            Exit Sub
        End If
    End If
    If rbcEditReplace(3).Value Then
        ilRet = MsgBox("Backup of database must be done before continuing with Mass Removal as the operation can not be reversed, has it been done", vbYesNo + vbQuestion, "Mass Delete")
        If ilRet = vbNo Then
            Exit Sub
        End If
    End If
    cmcDone.Enabled = False
    cmcProcess.Enabled = False
    mnuFilePrint.Enabled = True
    imReadingFile = True
    imClsGrid = True
    cmcGCancel.Enabled = True
    cmcGDone.Enabled = False
    cmcSort(0).Enabled = False
    cmcGSave.Enabled = False
    lacChg.Visible = False
    plcChg.Visible = False
    If rbcEditReplace(0).Value Or rbcEditReplace(1).Value Then 'Show or Fix by Typing
        mnuFileOpen.Enabled = False
        plcSelection.Visible = False
        plcLinks.Visible = False
        plcEditReplace.Visible = False
        plcDDF.Visible = False
        If (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary") Then
            gdcData.Move 30, 30, plcData.Width - 60 - vbcData.Width, 2430
            vbcData.Move gdcData.Left + gdcData.Width, 30, vbcData.Width, gdcData.Height
            gdcSub.Move 30, 2475, plcData.Width - 60, 2190
            gdcSub.Visible = True
            vbcData.LargeChange = 8
            gdcData.Rows = 10
        Else
            gdcData.Move 30, 30, plcData.Width - 60 - vbcData.Width, 4590
            vbcData.Move gdcData.Left + gdcData.Width - 15, 30, vbcData.Width, gdcData.Height
            gdcSub.Visible = False
            vbcData.LargeChange = 17
            gdcData.Rows = 20
        End If
        plcData.Visible = True
        gdcData.Visible = True
        vbcData.Visible = True
        plcData.ZOrder vbBringToFront
        ilNoCols = lbcEditReplace(2).ListCount ' + 1
        gdcData.Cols = ilNoCols
        gdcData.FixedAlignment(0) = 1   'Right Align
        gdcData.Col = 0
        gdcData.Row = 0
        gdcData.ColWidth(0) = ShowFix.TextWidth(" Display #")
        gdcData.Text = "Display #"
        ilCol = 1
        For ilLoop = 0 To lbcEditReplace(2).ListCount - 2 Step 1
            slFieldInfo = lbcEditReplace(2).List(ilLoop)
            ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
            gdcData.Col = ilCol
            gdcData.Row = 0
            gdcData.ColWidth(ilCol) = ShowFix.TextWidth(Trim$(slFieldName) & "   ")
            gdcData.Text = Trim$(slFieldName)
            ilCol = ilCol + 1
        Next ilLoop
        DoEvents
        mObtainRecords False
        cmcGCancel.Enabled = False
        cmcGDone.Enabled = True
        cmcSort(0).Enabled = True
        cmcGSave.Enabled = True
    ElseIf rbcEditReplace(2).Value Then 'Fix by Mass Substitution
        mnuFilePrint.Enabled = False
        mnuFileOpen.Enabled = False
        plcSelection.Visible = False
        plcLinks.Visible = False
        plcEditReplace.Visible = False
        plcDDF.Visible = False
        plcData.Visible = True
        gdcData.Visible = False
        gdcSub.Visible = False
        vbcData.Visible = False
        plcData.ZOrder vbBringToFront
        DoEvents
        mObtainRecords True
        cmcGCancel.Enabled = False
        cmcGDone.Enabled = True
    ElseIf rbcEditReplace(3).Value Then 'Mass Removal
        mnuFilePrint.Enabled = False
        mnuFileOpen.Enabled = False
        plcSelection.Visible = False
        plcLinks.Visible = False
        plcEditReplace.Visible = False
        plcDDF.Visible = False
        plcData.Visible = True
        gdcData.Visible = False
        gdcSub.Visible = False
        vbcData.Visible = False
        plcData.ZOrder vbBringToFront
        DoEvents
        mObtainRecords True
        cmcGCancel.Enabled = False
        cmcGDone.Enabled = True
    End If
    imClsGrid = False
    imReadingFile = False
    imTerminate = False
End Sub
Private Sub cmcProcess_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cmcProcess_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcSBrowse_Click()
    CMDialogBox.DialogTitle = "Save To File"
    If lacSave.Caption = "Save Specifications" Then
        CMDialogBox.Filter = "Show/Fix Spec|*.SFI|All|*.*"
        CMDialogBox.DefaultExt = ".SFI"
        If (Len(smImportPath) > 0) Then
            CMDialogBox.InitDir = Left$(smImportPath, Len(smImportPath) - 1)
        End If
    Else
        CMDialogBox.Filter = "ASC|*.Asc|Text|*.Txt|All|*.*"
        CMDialogBox.DefaultExt = ".Txt"
        If (Len(smExportPath) > 0) Then
            CMDialogBox.InitDir = Left$(smExportPath, Len(smExportPath) - 1)
        End If
    End If
    CMDialogBox.Flags = cdlOFNCreatePrompt
    CMDialogBox.Action = 1 'Open dialog
    edcSaveTo.Text = CMDialogBox.FileName
End Sub
Private Sub cmcSCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    If imEnableOnCancel Then
        cmcDone.Enabled = True
        cmcProcess.Enabled = True
    End If
    mnuFilePrint.Enabled = True
    plcSave.Visible = False
End Sub
Private Sub cmcSort_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub cmcSort_GotFocus(Index As Integer)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcSSave_Click()
    Dim slToFile
    Dim ilRet As Integer
    If imExporting Then
        Exit Sub
    End If
    If lacSave.Caption = "Print Grid Data" Then
        Screen.MousePointer = vbHourglass
        imExporting = True
        mSaveGrid 1
        Screen.MousePointer = vbDefault
        imExporting = False
    Else
        slToFile = edcSaveTo.Text
        If InStr(slToFile, ".") = 0 Then
            slToFile = slToFile & ".SFI"
        End If
        If InStr(slToFile, ":") = 0 Then
            If lacSave.Caption = "Save Specifications" Then
                slToFile = smImportPath & slToFile
            Else
                slToFile = smExportPath & slToFile
            End If
        End If
        ilRet = 0
        On Error GoTo cmcSSaveErr:
        hmTo = FreeFile
        Open slToFile For Output As hmTo
        If ilRet <> 0 Then
            MsgBox "Open " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            edcSaveTo.SetFocus
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        imExporting = True
        If lacSave.Caption = "Save Specifications" Then
            mSaveSpec
        Else
            mSaveGrid 0
        End If
        Close hmTo
        smCmmdFileName = edcSaveTo.Text
        Screen.MousePointer = vbDefault
        imExporting = False
        cmcSCancel_Click
    End If
    Exit Sub
cmcSSaveErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub edcExample_GotFocus()
    If plcSelection.Visible Then
        edcSelection.SetFocus
    End If
End Sub
Private Sub edcExample_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub edcFDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub edcFFile_Change()
    Dim ilPos As Integer
    Dim slName As String
    Dim slFromFile As String
    Dim ilRet As Integer
    If Len(edcFFile.Text) <= 0 Then
        Exit Sub
    End If
    slName = Trim$(edcFFile.Text)
    'Read file
    slFromFile = slName
    If InStr(slFromFile, ":") = 0 Then
        slFromFile = lbcFPath.Path & "\" & slFromFile
    End If
    ilRet = 0
    On Error GoTo edcFFileErr:
    hmFrom = FreeFile
    Open slFromFile For Input As hmFrom
    If ilRet <> 0 Then
        edcFDescription.Text = ""
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mReadSpec True
    Close hmFrom
    edcFFile.Text = slName
    edcFDescription.Text = smSpecDescript
    Screen.MousePointer = vbDefault
    Exit Sub
edcFFileErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub edcField_Change()
    Dim slStr As String
    slStr = edcField.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcField_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcField_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcField.SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcField_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcField.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcField.Text = slDate
            End If
        End If
        edcField.SelStart = 0
        edcField.SelLength = Len(edcField.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcField.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcField.Text = slDate
            End If
        End If
        edcField.SelStart = 0
        edcField.SelLength = Len(edcField.Text)
    End If
End Sub
Private Sub edcRecordNo_KeyPress(Index As Integer, KeyAscii As Integer)
    'Filter characters (numbers 0 thru 9)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcReplaceValue_Change()
    mSetCommands
End Sub
Private Sub edcReplaceValue_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub edcReplaceValue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub edcSearchKey_Change()
    mGetKeyFieldName
End Sub

Private Sub edcWrap_KeyPress(KeyAscii As Integer)
    'Filter characters (numbers 0 thru 9)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    'Me.Refresh
    If imTabSelection = 0 Then
        plcDDF.Visible = False
        plcDDF.Visible = True
    ElseIf imTabSelection = 1 Then
        plcSelection.Visible = False
        plcSelection.Visible = True
    ElseIf imTabSelection = 2 Then
        plcEditReplace.Visible = False
        plcEditReplace.Visible = True
    End If
    DoEvents
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        Exit Sub
    End If
    If (Len(smCmmdFileName) > 0) And (rbcEditReplace(2).Value) Then
        Exit Sub
    End If
    If InStr(1, Command$, "BugFix", vbTextCompare) > 0 Then
        igPasswordOk = True
        imPasswordAsked = True
    Else
        imPasswordAsked = False
        CSIPass.Show vbModal
        imPasswordAsked = True
        plcDDF.Visible = False
        plcDDF.Visible = True
        DoEvents
    End If
    If Not igPasswordOk Then
        rbcEditReplace(1).Enabled = False
        rbcEditReplace(2).Enabled = False
        rbcEditReplace(3).Enabled = False
        lacMode.ForeColor = RED
        lacMode.Caption = "Mode: View Only"
    Else
        rbcEditReplace(1).Enabled = False
        rbcEditReplace(2).Enabled = True
        rbcEditReplace(3).Enabled = True
        lacMode.ForeColor = DARKGREEN
        lacMode.Caption = "Mode: View and Update"
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
    Erase tmRecInfo
    Erase tmRowLinkList
    Erase tmGrid1Data
    Erase tmGrid2Data
    Erase tmGrid3Data
    Erase tmGrid4Data
    Erase tmGrid5Data
    Erase tmGrid6Data
    Erase tmDDFFile
    Erase tmDDFField
    Erase tmDDFLinkField
    Erase tmDDFLinkIndex
    Erase tmLink
    Erase tmKeyFile
    'Reset used instead of Close to cause # Clients on network to be decrement
    Set ShowFix = Nothing   'Remove data segment
    btrStopAppl
    End
End Sub
Private Sub gdcData_GotFocus()
    Dim i As Integer
    i = 0
End Sub
Private Sub gdcData_KeyDown(KeyCode As Integer, Shift As Integer)
    imKeyDownRow = gdcData.Row
End Sub
Private Sub gdcData_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilCol As Integer
    Dim ilRet As Integer
    Dim llRec As Long
    Dim ilRecIndex1 As Integer
    Dim ilRecIndex2 As Integer
    Dim ilField As Integer
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim slStr As String
    Dim slStr1 As String
    Dim slInt2 As INT2
    Dim slInt4 As INT4
    If (KeyCode = KEYUP) And (imKeyDownRow = 1) Then
        If vbcData.Value > 1 Then
            ilCol = gdcData.Col
            vbcData.Value = vbcData.Value - 1
            gdcData.Row = 1
            gdcData.Col = ilCol
        End If
    End If
    If (KeyCode = KEYDOWN) And (imKeyDownRow = vbcData.LargeChange + 1) Then
        If vbcData.Value < UBound(tmRecInfo) - vbcData.LargeChange Then
            ilCol = gdcData.Col
            vbcData.Value = vbcData.Value + 1
            gdcData.Row = vbcData.LargeChange + 1
            gdcData.Col = ilCol
        End If
    End If
    If KeyCode = KEYDELETE Then
        'If (imMouseUpX < gdcData.ColWidth(0)) And (gdcData.SelStartRow = gdcData.SelEndRow) And (gdcData.SelStartRow > 0) Then
        If (imMouseUpX < gdcData.ColWidth(0)) And (gdcData.RowSel = gdcData.Row) And (gdcData.RowSel > 0) Then
            If (Not imPasswordAsked) Then
                sgDBPath = lbcDataPath.Path & "\"
                CSIPass.Show vbModal
                If Not igPasswordOk Then
                    Exit Sub
                End If
                imPasswordAsked = True
            Else
                If Not igPasswordOk Then
                    Exit Sub
                End If
            End If
            ilRet = MsgBox("Delete the Record", vbYesNo + vbDefaultButton2 + vbQuestion + vbApplicationModal, "Erase")
            If ilRet = vbYes Then
                llRec = imVBCAdjust * (vbcData.Value - 1) + gdcData.RowSel
                ilRecIndex1 = 1
                ilRecIndex2 = 1
                Do While llRec > 32000
                    ilRecIndex1 = ilRecIndex1 + 1
                    llRec = llRec - 32000
                Loop
                ilRecIndex2 = llRec
                lmActualRecLen = imBaseRecLen
                ilRet = btrGetDirect(hmFile, tmRec, lmActualRecLen, tmRecInfo(ilRecIndex1, ilRecIndex2).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmFile)
                End If
                tmRecInfo(ilRecIndex1, ilRecIndex2).lRecNo = -tmRecInfo(ilRecIndex1, ilRecIndex2).lRecNo
                gdcData.Col = 0
                gdcData.Text = "Deleted"
            End If
        End If
    End If
    If KeyCode = KEYINSERT Then
        gdcData.Col = 0
        gdcData.Row = gdcData.RowSel   'gdcData.SelStartRow
        slStr1 = Trim$(gdcData.Text)
        If (slStr1 = "") Or (slStr1 = "Deleted") Then
            mUpdateField
            If (Not imPasswordAsked) Then
                sgDBPath = lbcDataPath.Path & "\"
                CSIPass.Show vbModal
                If Not igPasswordOk Then
                    Exit Sub
                End If
                imPasswordAsked = True
            Else
                If Not igPasswordOk Then
                    Exit Sub
                End If
            End If
            ilRet = MsgBox("Insert the Record", vbYesNo + vbDefaultButton2 + vbQuestion + vbApplicationModal, "Add")
            If ilRet = vbYes Then
                llRec = imVBCAdjust * (vbcData.Value - 1) + gdcData.Row
                ilRecIndex1 = 1
                ilRecIndex2 = 1
                Do While llRec > 32000
                    ilRecIndex1 = ilRecIndex1 + 1
                    llRec = llRec - 32000
                Loop
                ilRecIndex2 = llRec
                lmActualRecLen = imBaseRecLen
                'Update record
                If imVarLenRec Then
                    If imVarType = 12 Then
                        For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                            If (Asc(tmDDFField(ilField).sDataType) = 12) Then
                                If (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary") Then
                                    tmSsf = tmRec
                                    lmActualRecLen = imVarBaseRecLen + Len(tmProg) * tmSsf.iCount
                                Else
                                    lmActualRecLen = imVarBaseRecLen + tmDDFField(ilField).iSize
                                End If
                                'ilRet = btrUpdate(hmFile, tmRec, lmActualRecLen)
                                ilRet = btrInsert(hmFile, tmRec, lmActualRecLen, 0)
                                Exit For
                            End If
                        Next ilField
                    Else
                        ilRet = BTRV_ERR_NONE
                        For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                            If (Asc(tmDDFField(ilField).sDataType) = 13) Then
                                ilIndex = tmDDFField(ilField).iOffset + 1
                                'slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                                'HMemCpy ilValue, slInt2, 2
                                ilValue = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
                                lmActualRecLen = imVarBaseRecLen + ilValue + 4
                                'ilRet = btrUpdate(hmFile, tmRec, lmActualRecLen)
                                ilRet = btrInsert(hmFile, tmRec, lmActualRecLen, 0)
                                Exit For
                            End If
                        Next ilField
                    End If
                Else
                    lmActualRecLen = imBaseRecLen
                    'ilRet = btrUpdate(hmFile, tmRec, lmActualRecLen)
                    ilRet = btrInsert(hmFile, tmRec, lmActualRecLen, 0)
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    MsgBox "Insert Failed, Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal
                    imTerminate = True
                    Exit Sub
                End If
                ilRet = btrGetPosition(hmFile, tmRecInfo(ilRecIndex1, ilRecIndex2).lRecPos)
                If slStr1 = "Deleted" Then
                    tmRecInfo(ilRecIndex1, ilRecIndex2).iRecLen = lmActualRecLen
                    'tmRecInfo(ilRecIndex1, ilRecIndex2).iRecIndex1 = ilRecIndex1
                    'tmRecInfo(ilRecIndex1, ilRecIndex2).iRecIndex2 = ilRecIndex2
                    tmRecInfo(ilRecIndex1, ilRecIndex2).lRecNo = 0
                Else
                    tmRecInfo(ilRecIndex1, ilRecIndex2).iRecLen = lmActualRecLen
                    tmRecInfo(ilRecIndex1, ilRecIndex2).iRecIndex1 = ilRecIndex1
                    tmRecInfo(ilRecIndex1, ilRecIndex2).iRecIndex2 = ilRecIndex2
                    tmRecInfo(ilRecIndex1, ilRecIndex2).lRecNo = 0
                    imRecMaxUpper1 = UBound(tmRecInfo, 1)
                    imRecMaxUpper2 = UBound(tmRecInfo, 2)
                    lmMaxNoRec = lmMaxNoRec + 1
                End If
                gdcData.Col = 0
                gdcData.Text = "Inserted"
                imRecCls = False
                vbcData_Change
            End If
        End If
    End If
End Sub
Private Sub gdcData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If edcValue.Visible Then
        mUpdateField
    End If
End Sub
Private Sub gdcData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRec As Long
    Dim ilRecIndex1 As Integer
    Dim ilRecIndex2 As Integer
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim slFieldName As String
    Dim slFieldInfo As String
    Dim slFieldDataType As String
    Dim ilFileID As Integer
    Dim slFileName As String
    Dim ilField As Integer
    Dim slSvDataType As String
    Dim ilSize As Integer
    Dim ilLeft As Integer
    Dim ilTop As Integer
    Dim ilLoop As Integer
    Dim ilInsertRow As Integer
    Dim ilSvRow As Integer
    Dim ilSvCol As Integer
    Dim slStr1 As String
    If (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary") Then
        Exit Sub
    End If
    If (Not imPasswordAsked) Then
        sgDBPath = lbcDataPath.Path & "\"
        CSIPass.Show vbModal
        If Not igPasswordOk Then
            Exit Sub
        End If
        imPasswordAsked = True
    Else
        If Not igPasswordOk Then
            Exit Sub
        End If
    End If
    imMouseUpX = X
    mUpdateField
    imEditCol = -1
    imEditRow = -1
    'If (Y >= gdcData.RowHeight(0)) And (X >= gdcData.ColWidth(0)) And (gdcData.SelStartRow = gdcData.SelEndRow) And (gdcData.SelStartRow > 0) And (gdcData.SelStartCol = gdcData.SelEndCol) And (gdcData.SelStartCol > 0) Then
    If (Y >= gdcData.RowHeight(0)) And (X >= gdcData.ColWidth(0)) And (gdcData.RowSel = gdcData.Row) And (gdcData.RowSel > 0) And (gdcData.ColSel = gdcData.Col) And (gdcData.ColSel > 0) Then
        ilSvRow = gdcData.Row
        ilSvCol = gdcData.Col
        gdcData.Col = 0
        gdcData.Row = gdcData.RowSel    'gdcData.SelStartRow
        slStr1 = Trim$(gdcData.Text)
        If (slStr1 = "") Or (slStr1 = "Deleted") Then
            ilInsertRow = True
        Else
            ilInsertRow = False
        End If
        gdcData.Row = ilSvRow
        gdcData.Col = ilSvCol
        imEditCol = gdcData.ColSel 'gdcData.SelStartCol
        imEditRow = gdcData.RowSel 'gdcData.SelStartRow
        llRec = imVBCAdjust * (vbcData.Value - 1) + imEditRow
        ilRecIndex1 = 1
        ilRecIndex2 = 1
        Do While llRec > 32000
            ilRecIndex1 = ilRecIndex1 + 1
            llRec = llRec - 32000
        Loop
        ilRecIndex2 = llRec
        If Not ilInsertRow Then
            lmActualRecLen = imBaseRecLen
            ilRet = btrGetDirect(hmFile, tmRec, lmActualRecLen, tmRecInfo(ilRecIndex1, ilRecIndex2).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        End If
        'If ilRet = BTRV_ERR_NONE Then
            slFieldInfo = lbcEditReplace(2).List(imEditCol - 1)
            ilRet = gParseItem(slFieldInfo, 1, "|", slFileName)
            ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
            ilRet = gParseItem(slFieldInfo, 4, "|", slFieldDataType)
            If ilRet <> CP_MSG_NONE Then
                slFieldDataType = ""
                If (StrComp(Trim$(slFileName), "Mnf.Btr", 1) = 0) And (StrComp(Trim$(slFieldName), "mnfRPU", 1) = 0) Then
                    slFieldDataType = Chr(101)
                End If
                If (StrComp(Trim$(slFileName), "Mnf.Btr", 1) = 0) And (StrComp(Trim$(slFieldName), "mnfSSComm", 1) = 0) Then
                    slFieldDataType = Chr(101)
                End If
            Else
                If StrComp(slFieldDataType, "Length of Time", 1) = 0 Then
                    slFieldDataType = Chr(100)
                End If
                If StrComp(slFieldDataType, "Hex Field Dump", 1) = 0 Then
                    slFieldDataType = Chr(101)
                End If
            End If
            ilFileID = mGetFileID(slFileName)
            For ilLoop = 0 To UBound(tmLink) - 1 Step 1
                If tmLink(ilLoop).iToFileID = ilFileID Then
                    imEditCol = -1
                    imEditRow = -1
                    Exit Sub
                End If
            Next ilLoop
            If (Not imRecCls) And (ilInsertRow) Then
                For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                    If (ilFileID = tmDDFField(ilField).iFileID) Then
                        smFieldValue = ""
                        mPutField tmDDFField(ilField), 0
                    End If
                Next ilField
                imRecCls = True
            End If
            For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                If (StrComp(Trim$(slFieldName), Trim$(tmDDFField(ilField).sName), 1) = 0) And (ilFileID = tmDDFField(ilField).iFileID) Then
                    'mGetField tmDDFField(tmRecSel(gdcData.SelStartCol).iFieldIndex1), 0
                    slSvDataType = tmDDFField(ilField).sDataType
                    If slFieldDataType <> "" Then
                        tmDDFField(ilField).sDataType = slFieldDataType
                    End If
                    If Not ilInsertRow Then
                        mGetField tmDDFField(ilField), 0
                    Else
                        smFieldValue = gdcData.Text
                    End If
                    ilSize = mFieldSize(tmDDFField(ilField))
                    tmDDFField(ilField).sDataType = slSvDataType
                    ilLeft = 45 + gdcData.ColWidth(0)
                    ilTop = 60
                    For ilLoop = gdcData.TopRow - 1 To imEditRow - 1 Step 1
                        ilTop = ilTop + gdcData.RowHeight(ilLoop) '+ 15
                    Next ilLoop
                    'For ilLoop = 0 To imEditCol - 1 Step 1
                    For ilLoop = gdcData.LeftCol To imEditCol - 1 Step 1
                        ilLeft = ilLeft + gdcData.ColWidth(ilLoop) '+ 15
                    Next ilLoop
                    edcValue.Move ilLeft, ilTop, gdcData.ColWidth(imEditCol) - 15
                    edcValue.MaxLength = ilSize
                    edcValue.Text = Trim$(smFieldValue)
                    edcValue.SelStart = 0
                    edcValue.SelLength = Len(edcValue.Text)
                    edcValue.ZOrder vbBringToFront
                    edcValue.Visible = True
                    edcValue.SetFocus
                    Exit For
                End If
            Next ilField
        'End If
    End If
End Sub
Private Sub gdcData_SelChange()
    Dim ilCol As Integer
    Dim ilRow As Integer
    Dim llRec As Long
    Dim ilRecIndex1 As Integer
    Dim ilRecIndex2 As Integer
    Dim ilLoop As Integer
    Dim slValue As String
    Dim llValue As Long
    Dim ilRet As Integer
    Dim ilRank As Integer
    Dim ilPriceLevel As Integer
    Dim slStr As String
    
    If imReadingFile Then    'Getting Data
        Exit Sub
    End If
    If (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary") Then
        'If (gdcData.SelStartRow = gdcData.SelEndRow) And (gdcData.SelStartRow > 0) Then
        If (gdcData.RowSel = gdcData.Row) And (gdcData.RowSel > 0) Then
            llRec = imVBCAdjust * (vbcData.Value - 1) + gdcData.RowSel 'gdcData.SelStartRow
            ilRecIndex1 = 1
            ilRecIndex2 = 1
            Do While llRec > 32000
                ilRecIndex1 = ilRecIndex1 + 1
                llRec = llRec - 32000
            Loop
            ilRecIndex2 = llRec
            lmActualRecLen = imBaseRecLen
            ilRet = gSSFGetDirect(hmFile, tmSsf, lmActualRecLen, tmRecInfo(ilRecIndex1, ilRecIndex2).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                gdcSub.Cols = 11
                gdcSub.Rows = tmSsf.iCount + 1
                gdcSub.RowHeight(0) = 705   '3 * gdcSub.RowHeight(1)
                gdcSub.Col = 0
                gdcSub.Row = 0
                gdcSub.Text = Trim$(Str$(tmRecInfo(ilRecIndex1, ilRecIndex2).lRecNo))
                gdcSub.ColWidth(0) = ShowFix.TextWidth("100000 ")
                gdcSub.Col = 1
                gdcSub.Row = 0
                gdcSub.Text = "Start Time" & Chr(13) & Chr(10) & "Time" & Chr(13) & Chr(10) & "sdfCode"
                gdcSub.ColWidth(1) = ShowFix.TextWidth("Start Time  ")
                gdcSub.Col = 2
                gdcSub.Row = 0
                gdcSub.Text = "End Time" & Chr(13) & Chr(10) & "ltfCode" & Chr(13) & Chr(10) & "Line Entered"
                gdcSub.ColWidth(2) = ShowFix.TextWidth("Line Entered  ")
                gdcSub.Col = 3
                gdcSub.Row = 0
                gdcSub.Text = "mnfExcl" & Chr(13) & Chr(10) & "Units" & Chr(13) & Chr(10) & "Hour >="
                gdcSub.ColWidth(3) = ShowFix.TextWidth("Hour >=  ")
                gdcSub.Col = 4
                gdcSub.Row = 0
                gdcSub.Text = "mnfExcl" & Chr(13) & Chr(10) & "Length" & Chr(13) & Chr(10) & "Hour <="
                gdcSub.ColWidth(4) = ShowFix.TextWidth("Hour <= ")
                gdcSub.Col = 5
                gdcSub.Row = 0
                gdcSub.Text = "ltfCode" & Chr(13) & Chr(10) & "anfCode" & Chr(13) & Chr(10) & "mnfComp"
                gdcSub.ColWidth(5) = ShowFix.TextWidth("mnfCode  ")
                gdcSub.Col = 6
                gdcSub.Row = 0
                gdcSub.Text = "lvfCode" & Chr(13) & Chr(10) & "# Spots" & Chr(13) & Chr(10) & "mnfComp"
                gdcSub.ColWidth(6) = ShowFix.TextWidth("# Spots  ")
                gdcSub.Col = 7
                gdcSub.Row = 0
                gdcSub.Text = "----" & Chr(13) & Chr(10) & "Sustaining" & Chr(13) & Chr(10) & "Length"
                gdcSub.ColWidth(7) = ShowFix.TextWidth("Sustaining  ")
                gdcSub.Col = 8
                gdcSub.Row = 0
                gdcSub.Text = "----" & Chr(13) & Chr(10) & "Sponsorship" & Chr(13) & Chr(10) & "adfCode"
                gdcSub.ColWidth(8) = ShowFix.TextWidth("Sponsorship  ")
                gdcSub.Col = 9
                gdcSub.Row = 0
                gdcSub.Text = "----" & Chr(13) & Chr(10) & "Avail" & Chr(13) & Chr(10) & "Rank"
                gdcSub.ColWidth(9) = ShowFix.TextWidth("Unlock  ")
                gdcSub.Col = 10
                gdcSub.Row = 0
                gdcSub.Text = "----" & Chr(13) & Chr(10) & "Spot" & Chr(13) & Chr(10) & "----"
                gdcSub.ColWidth(10) = ShowFix.TextWidth("Unlock  ")
                For ilLoop = 1 To tmSsf.iCount Step 1
                    ilRow = ilLoop
                   LSet tmProg = tmSsf.tPAS(ilLoop)
                    If tmProg.iRecType = 1 Then    'Program
                        gdcSub.Col = 0
                        gdcSub.Row = ilRow
                        gdcSub.Text = "Prog"
                        gUnpackTime tmProg.iStartTime(0), tmProg.iStartTime(1), "A", "1", slValue
                        gdcSub.Col = 1
                        gdcSub.Row = ilRow
                        gdcSub.Text = slValue
                        gUnpackTime tmProg.iEndTime(0), tmProg.iEndTime(1), "A", "1", slValue
                        gdcSub.Col = 2
                        gdcSub.Row = ilRow
                        gdcSub.Text = slValue
                        gdcSub.Col = 3
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmProg.iMnfExcl(0)))
                        gdcSub.Col = 4
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmProg.iMnfExcl(1)))
                        gdcSub.Col = 5
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmProg.iLtfCode))
                        gdcSub.Col = 6
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmProg.lLvfCode))
                        gdcSub.Col = 7
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 8
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 9
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 10
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                    ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then 'Avail
                       LSet tmAvail = tmSsf.tPAS(ilLoop)
                        gdcSub.Col = 0
                        gdcSub.Row = ilRow
                        gdcSub.Text = "Avail"
                        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slValue
                        gdcSub.Col = 1
                        gdcSub.Row = ilRow
                        gdcSub.Text = slValue
                        gdcSub.Col = 2
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmAvail.iLtfCode))
                        gdcSub.Col = 3
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmAvail.iAvInfo And &H1F))
                        gdcSub.Col = 4
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmAvail.iLen))
                        gdcSub.Col = 5
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmAvail.ianfCode))
                        gdcSub.Col = 6
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmAvail.iNoSpotsThis))
                        gdcSub.Col = 7
                        gdcSub.Row = ilRow
                        If (tmAvail.iAvInfo And SSSUSTAINING) = SSSUSTAINING Then
                            gdcSub.Text = "Allow"
                        Else
                            gdcSub.Text = "Disallow"
                        End If
                        gdcSub.Col = 8
                        gdcSub.Row = ilRow
                        If (tmAvail.iAvInfo And SSSPONSORSHIP) = SSSPONSORSHIP Then
                            gdcSub.Text = "Allow"
                        Else
                            gdcSub.Text = "Disallow"
                        End If
                        gdcSub.Col = 9
                        gdcSub.Row = ilRow
                        If (tmAvail.iAvInfo And SSLOCK) = SSLOCK Then
                            gdcSub.Text = "Lock"
                        Else
                            gdcSub.Text = "Unlock"
                        End If
                        gdcSub.Col = 10
                        gdcSub.Row = ilRow
                        If (tmAvail.iAvInfo And SSLOCKSPOT) = SSLOCKSPOT Then
                            gdcSub.Text = "Lock"
                        Else
                            gdcSub.Text = "Unlock"
                        End If
                    ElseIf (tmProg.iRecType And &HF) = 10 Then    'Spot
                       LSet tmSpot = tmSsf.tPAS(ilLoop)
                        gdcSub.Col = 0
                        gdcSub.Row = ilRow
                        gdcSub.Text = "Spot"
                        gdcSub.Col = 1
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmSpot.lSdfCode))
                        gdcSub.Col = 2
                        gdcSub.Row = ilRow
                        slStr = ""
                        If (tmSpot.lBkInfo And SS1STPOSITION) = SS1STPOSITION Then
                            slStr = "F,"
                            If (tmSpot.lBkInfo And SSSOLOAVAIL) = SSSOLOAVAIL Then
                                slStr = slStr & "S,"
                            End If
                        Else
                            If (tmSpot.lBkInfo And SSSOLOAVAIL) = SSSOLOAVAIL Then
                                slStr = "S,"
                            End If
                        End If
                        llValue = tmSpot.lBkInfo And &H1FFFF    'Book Data
                        gdcSub.Text = slStr & Format$(llValue, "m/d/yy")
                        gdcSub.Col = 3
                        gdcSub.Row = ilRow
                        llValue = (tmSpot.lBkInfo And &H7E0000) \ SHIFT17  'Start Hour in Minutes
                        gdcSub.Text = Trim$(Str$(llValue)) & "Min"
                        gdcSub.Col = 4
                        gdcSub.Row = ilRow
                        llValue = (tmSpot.lBkInfo And &H1F800000) \ SHIFT23  'Start Hour in Minutes
                        gdcSub.Text = Trim$(Str$(llValue)) & "Min"
                        gdcSub.Col = 5
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmSpot.iMnfComp(0)))
                        gdcSub.Col = 6
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmSpot.iMnfComp(1)))
                        gdcSub.Col = 7
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmSpot.iPosLen And &HFFF))
                        gdcSub.Col = 8
                        gdcSub.Row = ilRow
                        gdcSub.Text = Trim$(Str$(tmSpot.iAdfCode))
                        gdcSub.Col = 9
                        gdcSub.Row = ilRow
                        'gdcSub.Text = Trim$(Str$(tmSpot.iRank))
                        ilRank = tmSpot.iRank And RANKMASK
                        ilPriceLevel = (tmSpot.iRank And PRICELEVELMASK) / SHIFT11
                        gdcSub.Text = Trim$(Str$(ilPriceLevel)) & "-" & Trim$(Str$(ilRank))
                        gdcSub.Col = 10
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                    ElseIf (tmProg.iRecType And &HF) = 11 Then    'BBSpot
                        tmBBSpot = tmSsf.tPAS(ilLoop)
                        gdcSub.Col = 0
                        gdcSub.Row = ilRow
                        gdcSub.Text = "BBSpot"
                        gdcSub.Col = 1
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 2
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 3
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 4
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 5
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 6
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 7
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 8
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 9
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                        gdcSub.Col = 10
                        gdcSub.Row = ilRow
                        gdcSub.Text = ""
                    End If
                Next ilLoop
            Else
                For ilRow = 0 To gdcSub.Rows - 1 Step 1
                    gdcSub.Row = ilRow
                    'Blank columns
                    For ilCol = 0 To gdcSub.Cols - 1 Step 1
                        gdcSub.Col = ilCol
                        gdcSub.Text = ""
                    Next ilCol
                Next ilRow
            End If
        Else
            For ilRow = 0 To gdcSub.Rows - 1 Step 1
                gdcSub.Row = ilRow
                'Blank columns
                For ilCol = 0 To gdcSub.Cols - 1 Step 1
                    gdcSub.Col = ilCol
                    gdcSub.Text = ""
                Next ilCol
            Next ilRow
        End If
    End If
End Sub

Private Sub imcEditReplace_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub imcEditReplace_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    mDragOver DRAGEDIT, State, Index    '0=Top; 1= botton
End Sub
Private Sub imcTrash_DragDrop(Source As Control, X As Single, Y As Single)
    Select Case imDragSrce
        Case DRAGFIELD
        Case DRAGEDIT
            lbcEditReplace(2).RemoveItem imDragIndexSrce
            DoEvents
    End Select
    mClearDrag
End Sub
Private Sub imcTrash_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Select Case imDragSrce
        Case DRAGFIELD
        Case DRAGEDIT
            If State = vbEnter Then    'Enter drag over
                lbcEditReplace(2).DragIcon = imcIconDwnArrow.DragIcon
                imcTrash.Picture = imcTrashOpened.Picture
            ElseIf State = vbLeave Then
                lbcEditReplace(2).DragIcon = imcIconDrag.DragIcon
                imcTrash.Picture = imcTrashClosed.Picture
            End If
    End Select
End Sub
Private Sub lbcDataPath_Change()
    'lbcFromField.Clear
    'lbcSelection(1).Clear
    'edcSelection.Text = ""
    'lbcEditReplace(0).Clear
    'lbcEditReplace(2).Clear
    'lbcToIndex.Clear
    'lbcToField.Clear
    'lbcLink.Clear
End Sub
Private Sub lbcEditReplace_Click(Index As Integer)
    Dim ilRet As Integer
    Dim slFieldInfo As String
    Dim slFileName As String
    Dim slFieldName As String
    Dim slFieldDataType As String
    Dim slNewDataType As String
    If Index = 1 Then   'Add or remove conversion type
        If lbcEditReplace(0).ListIndex >= 0 Then
            If lbcEditReplace(1).ListIndex = 0 Then
                slFieldInfo = lbcEditReplace(0).List(lbcEditReplace(0).ListIndex)
                ilRet = gParseItem(slFieldInfo, 1, "|", slFileName)
                ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
                ilRet = gParseItem(slFieldInfo, 3, "|", slFieldDataType)
                lbcEditReplace(0).List(lbcEditReplace(0).ListIndex) = slFileName & "|" & slFieldName & "|" & slFieldDataType
            Else
                slNewDataType = lbcEditReplace(1).List(lbcEditReplace(1).ListIndex)
                slFieldInfo = lbcEditReplace(0).List(lbcEditReplace(0).ListIndex)
                ilRet = gParseItem(slFieldInfo, 1, "|", slFileName)
                ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
                ilRet = gParseItem(slFieldInfo, 3, "|", slFieldDataType)
                lbcEditReplace(0).List(lbcEditReplace(0).ListIndex) = slFileName & "|" & slFieldName & "|" & slFieldDataType & "|" & slNewDataType
            End If
        End If
    End If
    If Index = 0 Then
        pbcLbcEditReplace0_Paint
    ElseIf Index = 2 Then
        pbcLbcEditReplace2_Paint
    End If

    mSetCommands
End Sub
Private Sub lbcEditReplace_DblClick(Index As Integer)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If Index = 0 Then
        cmcEditReplace_Click 0
    ElseIf Index = 0 Then
        pbcLbcEditReplace0_Paint
    ElseIf Index = 2 Then
        pbcLbcEditReplace2_Paint
    End If
End Sub
Private Sub lbcEditReplace_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Dim ilRet As Integer
    Dim slName As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRes As Integer
    If Index = 2 Then
        If imDragDest = -1 Then
            mClearDrag
            Exit Sub
        End If
        Select Case imDragSrce
            Case DRAGFIELD
                lbcEditReplace(0).ListIndex = imDragIndexSrce
                cmcEditReplace_Click 0
                imDragIndexDest = -1
        End Select
    End If
    mClearDrag
End Sub
Private Sub lbcEditReplace_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Dim ilListIndex As Integer
    If Index = 2 Then
    
        imDragIndexDest = -1
        If imDragSrce = DRAGFIELD Then
            If State = vbLeave Then
                lbcEditReplace(0).DragIcon = imcIconDrag.DragIcon
                Exit Sub
            End If
            ilListIndex = Y \ fgListHTArial825 + lbcEditReplace(2).TopIndex
            If (ilListIndex <= lbcEditReplace(2).ListCount - 1) Then
                lbcEditReplace(0).DragIcon = imcIconInsert.DragIcon
                imDragIndexDest = ilListIndex
                imDragDest = DRAGEDIT
                If ilListIndex <> lbcEditReplace(2).ListCount - 2 Then
                    If lbcEditReplace(2).ListIndex <> ilListIndex Then
                        lbcEditReplace(2).ListIndex = ilListIndex
                    End If
                End If
                'DoEvents   'This is causing a stack overflow
            Else
                lbcEditReplace(0).DragIcon = imcIconDrag.DragIcon
            End If
        End If
    End If
End Sub
Private Sub lbcEditReplace_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        If (rbcEditReplace(0).Value) Or (rbcEditReplace(2).Value And (Len(edcReplaceValue.Text) <> 0)) Then
            fmDragX = X
            fmDragY = Y
            imDragButton = Button
            imDragType = 0
            imDragShift = Shift
            imDragSrce = DRAGFIELD
            imDragIndexDest = -1
            tmcDrag.Enabled = True  'Start timer to see if drag or click
        End If
    ElseIf (Index = 2) And (rbcEditReplace(3).Value = False) Then
        fmDragX = X
        fmDragY = Y
        imDragButton = Button
        imDragType = 0
        imDragShift = Shift
        imDragSrce = DRAGEDIT
        imDragIndexDest = -1
        tmcDrag.Enabled = True  'Start timer to see if drag or click
    End If
End Sub
Private Sub lbcEditReplace_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub lbcEditReplace_Scroll(Index As Integer)
    If Index = 0 Then
        pbcLbcEditReplace0_Paint
    ElseIf Index = 2 Then
        pbcLbcEditReplace2_Paint
    End If
End Sub

Private Sub lbcFFile_Click()
    Dim ilPos As Integer
    Dim slName As String
    Dim slFromFile As String
    Dim ilRet As Integer
    If lbcFFile.ListIndex < 0 Then
        Exit Sub
    End If
    slName = Trim$(lbcFFile.List(lbcFFile.ListIndex))
    'Read file
    slFromFile = slName
    If InStr(slFromFile, ":") = 0 Then
        slFromFile = lbcFPath.Path & "\" & slFromFile
    End If
    ilRet = 0
    On Error GoTo lbcFFileErr:
    hmFrom = FreeFile
    Open slFromFile For Input As hmFrom
    If ilRet <> 0 Then
        edcFDescription.Text = ""
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mReadSpec True
    Close hmFrom
    edcFFile.Text = slName
    edcFDescription.Text = smSpecDescript
    Screen.MousePointer = vbDefault
    Exit Sub
lbcFFileErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub lbcFFile_DblClick()
    cmcFOpen_Click
End Sub
Private Sub lbcFile_Click()
    Dim slName As String
    Dim ilLoop As Integer
    edcSelection.Text = ""
    lbcEditReplace(2).Clear
    lbcEditReplace(2).AddItem " "
    pbcLbcEditReplace2_Paint
    lbcLink.Clear
    ReDim tmLink(0 To 0) As LINK
    lbcFromField.Clear
    lbcSelection(1).Clear
    pbcLbcSelection1_Paint
    lbcEditReplace(0).Clear
    pbcLbcEditReplace0_Paint
    edcSearchKey.Text = "0"
    imFileID = -1
    ReDim tmDDFField(0 To 0) As DDFFIELD
    If lbcFile.ListIndex >= 0 Then
        slName = lbcFile.List(lbcFile.ListIndex)
        For ilLoop = 0 To UBound(tmDDFFile) - 1 Step 1
            If StrComp(Trim$(tmDDFFile(ilLoop).sName), slName, 1) = 0 Then
                imFileID = tmDDFFile(ilLoop).iFileID
                mFieldBuild tmDDFFile(ilLoop).iFileID, tmDDFField()
                mFieldPop tmDDFField(), 0, UBound(tmDDFField) - 1, lbcFromField, False
                mFieldPop tmDDFField(), 0, UBound(tmDDFField) - 1, lbcEditReplace(0), True
                mFieldPop tmDDFField(), 0, UBound(tmDDFField) - 1, lbcSelection(1), True
                Exit For
            End If
        Next ilLoop
    End If
    mGetKeyFieldName
    mSetCommands
End Sub
Private Sub lbcFPath_Change()
    edcFFile.Text = "*.SFI"
    lbcFFile.FileName = lbcFPath.Path & "\*.SFI"
End Sub
Private Sub lbcFromField_Click()
    mSetLinkCommands
End Sub
Private Sub lbcLink_Click()
    mSetLinkCommands
End Sub

Private Sub lbcPath_Change()
    Dim slStr As String
    If Not imIgnoreChg Then
        'csiClearInit
        slStr = lbcPath.Path
        If right$(slStr, 1) <> "\" Then
            smDDFPath = lbcPath.Path & "\"
        Else
            smDDFPath = lbcPath.Path
        End If
        'hmDB = CBtrvMngrInit(0) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        mFileBuild tmDDFFile()
        mFilePop tmDDFFile(), lbcFile
        mFilePop tmDDFFile(), lbcToFile
    End If
    'If cbcDrive.Drive = cbcDataDrive.Drive Then
    '    lbcDataPath.Path = lbcPath.Path
    'End If
    lbcFromField.Clear
    lbcSelection(1).Clear
    pbcLbcSelection1_Paint
    edcSelection.Text = ""
    lbcEditReplace(0).Clear
    lbcEditReplace(2).Clear
    lbcEditReplace(2).AddItem " "
    pbcLbcEditReplace0_Paint
    pbcLbcEditReplace2_Paint
    lbcToIndex.Clear
    lbcToField.Clear
    lbcLink.Clear
    ReDim tmLink(0 To 0) As LINK
End Sub
Private Sub lbcSelection_Click(Index As Integer)
    Dim slOper As String
    Dim slFieldName As String
    Dim ilPos As Integer
    Dim ilStart As Integer
    Dim ilLen As Integer
    Dim ilPosX As Integer
    Dim ilPosY As Integer
    Dim slFieldInfo As String
    Dim slFileName As String
    Dim ilRet As Integer
    If Index = 0 Then   'Operators
        If lbcSelection(0).ListIndex >= 0 Then
            slOper = Trim$(lbcSelection(0).List(lbcSelection(0).ListIndex))
            If StrComp(slOper, "X = Y Equal", 1) = 0 Then
                slOper = "X = Y"
                ilLen = 1
                ilStart = edcSelection.SelStart
            ElseIf StrComp(slOper, "X <> Y Not Equal", 1) = 0 Then
                slOper = "X <> Y"
                ilLen = 1
                ilStart = edcSelection.SelStart
            ElseIf StrComp(slOper, "X < Y Less Than", 1) = 0 Then
                slOper = "X < Y"
                ilLen = 1
                ilStart = edcSelection.SelStart
            ElseIf StrComp(slOper, "X > Y Greater Than", 1) = 0 Then
                slOper = "X > Y"
                ilLen = 1
                ilStart = edcSelection.SelStart
            ElseIf StrComp(slOper, "X <= Y Less Than or Equal", 1) = 0 Then
                slOper = "X <= Y"
                ilLen = 1
                ilStart = edcSelection.SelStart
            ElseIf StrComp(slOper, "X >= Y Greater Than Or Equal", 1) = 0 Then
                slOper = "X >= Y"
                ilLen = 1
                ilStart = edcSelection.SelStart
            ElseIf InStr(slOper, "And") > 0 Then
                slOper = "And"
                ilLen = -1
                ilStart = -1
            ElseIf InStr(slOper, "Or") > 0 Then
                slOper = "Or"
                ilLen = -1
                ilStart = -1
            ElseIf InStr(slOper, "Parentheses") > 0 Then
                slOper = "(X)"
                ilLen = 1
                ilStart = edcSelection.SelStart + 1
            Else
                Beep
                Exit Sub
            End If
            If edcSelection.SelStart = Len(edcSelection.Text) Then
                If Len(edcSelection.Text) <> 0 Then
                    If Asc(right$(edcSelection.Text, 1)) = 10 Then
                        slOper = slOper
                    Else
                        slOper = Chr(13) & Chr(10) & slOper
                        ilStart = ilStart + 2
                    End If
                Else
                    If (slOper = "And") Or (slOper = "Or") Then
                        Beep
                        Exit Sub
                    Else
                        slOper = slOper
                    End If
                End If
            Else
                If edcSelection.SelStart = 0 Then
                    slOper = slOper & Chr(13) & Chr(10)
                    ilStart = ilStart + 2
                ElseIf Asc(Mid(edcSelection.Text, edcSelection.SelStart, 1)) = 10 Then
                    slOper = slOper & Chr(13) & Chr(10)
                    ilStart = ilStart + 2
                ElseIf Mid(edcSelection.Text, edcSelection.SelStart + 1, 1) = ")" Then
                    slOper = Chr(13) & Chr(10) & slOper
                    ilStart = ilStart + 2
                ElseIf Asc(Mid(edcSelection.Text, edcSelection.SelStart + 1, 1)) = 13 Then
                    slOper = Chr(13) & Chr(10) & slOper
                    ilStart = ilStart + 2
                ElseIf Mid(edcSelection.Text, edcSelection.SelStart, 1) = "(" Then
                    slOper = slOper
                Else
                    Beep
                    Exit Sub
                End If
            End If
            Clipboard.SetText slOper
            edcSelection.SelText = Clipboard.GetText()
            If ilLen <> -1 Then
                edcSelection.SelStart = ilStart
                edcSelection.SelLength = ilLen
                edcSelection.SetFocus
            End If
        End If
    ElseIf Index = 1 Then   'Field
        If lbcSelection(1).ListIndex >= 0 Then
            slFieldInfo = lbcSelection(1).List(lbcSelection(1).ListIndex)
            ilRet = gParseItem(slFieldInfo, 1, "|", slFileName)
            ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
            Clipboard.SetText slFileName & ":" & slFieldName
            edcSelection.SelText = Clipboard.GetText()
            ilPosX = InStr(edcSelection.Text, "(X ")
            If ilPosX = 0 Then
                ilPosX = InStr(edcSelection.Text, "X ")
            Else
                ilPosX = ilPosX + 1
            End If
            ilPosY = InStr(edcSelection.Text, " Y")
            If ilPosY > 0 Then
                ilPosY = ilPosY + 1
            End If
            If (ilPosX <> 0) And (ilPosY <> 0) Then
                If ilPosX < ilPosY Then
                    edcSelection.SelStart = ilPosX - 1
                    edcSelection.SelLength = 1
                Else
                    edcSelection.SelStart = ilPosY - 1
                    edcSelection.SelLength = 1
                End If
            ElseIf ilPosX > 0 Then
                edcSelection.SelStart = ilPosX - 1
                edcSelection.SelLength = 1
            ElseIf ilPosY > 0 Then
                edcSelection.SelStart = ilPosY - 1
                edcSelection.SelLength = 1
            Else
                edcSelection.SelStart = Len(edcSelection.Text)
                edcSelection.SelLength = 0
            End If
            edcSelection.SetFocus
        End If
        pbcLbcSelection1_Paint
    End If
End Sub

Private Sub lbcSelection_Scroll(Index As Integer)
    pbcLbcSelection1_Paint
End Sub

Private Sub lbcToField_Click()
    mSetLinkCommands
End Sub
Private Sub lbcToFile_Click()
    Dim ilLoop As Integer
    Dim slName As String
    lbcToField.Clear
    lbcToIndex.Clear
    ReDim tmDDFLinkField(0 To 0) As DDFFIELD
    If lbcToFile.ListIndex >= 0 Then
        slName = lbcToFile.List(lbcToFile.ListIndex)
        For ilLoop = 0 To UBound(tmDDFFile) - 1 Step 1
            If StrComp(Trim$(tmDDFFile(ilLoop).sName), slName, 1) = 0 Then
                mFieldBuild tmDDFFile(ilLoop).iFileID, tmDDFLinkField()
                'mFieldPop tmDDFLinkField(), lbcToField, False
                mIndexBuild tmDDFFile(ilLoop).iFileID, tmDDFLinkIndex()
                mIndexPop tmDDFLinkIndex(), tmDDFLinkField(), lbcToIndex
                Exit For
            End If
        Next ilLoop
    End If
    mSetLinkCommands
End Sub
Private Sub lbcToIndex_Click()
    Dim ilIndex As Integer
    Dim slName As String
    Dim slIndex As String
    Dim ilMaxSeg As Integer
    Dim ilLoop As Integer
    Dim ilSeg As Integer
    Dim ilField As Integer
    Dim ilPos As Integer
    lbcToField.Clear
    mSetLinkCommands
    If lbcToIndex.ListIndex < 0 Then
        Exit Sub
    End If
    slName = lbcToIndex.List(lbcToIndex.ListIndex)
    ilPos = InStr(slName, ":")
    If ilPos = 0 Then
        Exit Sub
    End If
    slIndex = Mid$(slName, 7, ilPos - 7)
    ilIndex = Val(slIndex)
    ilMaxSeg = 0
    For ilLoop = 0 To UBound(tmDDFLinkIndex) - 1 Step 1
        If (tmDDFLinkIndex(ilLoop).iNumber = ilIndex) And (tmDDFLinkIndex(ilLoop).iPart > ilMaxSeg) Then
            ilMaxSeg = tmDDFLinkIndex(ilLoop).iPart
        End If
    Next ilLoop
    For ilSeg = 0 To ilMaxSeg Step 1
        For ilLoop = 0 To UBound(tmDDFLinkIndex) - 1 Step 1
            If (tmDDFLinkIndex(ilLoop).iNumber = ilIndex) And (tmDDFLinkIndex(ilLoop).iPart = ilSeg) Then
                For ilField = 0 To UBound(tmDDFLinkField) - 1 Step 1
                    If tmDDFLinkField(ilField).iFieldID = tmDDFLinkIndex(ilLoop).iFieldID Then
                        lbcToField.AddItem Trim$(tmDDFLinkField(ilField).sName)
                    End If
                Next ilField
            End If
        Next ilLoop
    Next ilSeg
    mSetLinkCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddGrid1Data                   *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add data to Grid column table   *
'*                                                     *
'*******************************************************
Private Sub mAddGrid1Data(ilField As Integer)
    tmGrid1Data(imGrid1Upper1, imGrid1Upper2).iNextIndex1 = -1
    tmGrid1Data(imGrid1Upper1, imGrid1Upper2).iNextIndex2 = -1
    tmGrid1Data(imGrid1Upper1, imGrid1Upper2).iDDFFieldIndex = ilField
    tmGrid1Data(imGrid1Upper1, imGrid1Upper2).sValue = smFieldValue
    imGrid1Upper2 = imGrid1Upper2 + 1
    If imGrid1Upper2 > 32000 Then
        imGrid1Upper1 = imGrid1Upper1 + 1
        imGrid1Upper2 = 1
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddGrid2Data                   *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add data to Grid column table   *
'*                                                     *
'*******************************************************
Private Sub mAddGrid2Data(ilField As Integer)
    tmGrid2Data(imGrid2Upper1, imGrid2Upper2).iNextIndex1 = -1
    tmGrid2Data(imGrid2Upper1, imGrid2Upper2).iNextIndex2 = -1
    tmGrid2Data(imGrid2Upper1, imGrid2Upper2).iDDFFieldIndex = ilField
    tmGrid2Data(imGrid2Upper1, imGrid2Upper2).sValue = smFieldValue
    imGrid2Upper2 = imGrid2Upper2 + 1
    If imGrid2Upper2 > 32000 Then
        imGrid2Upper1 = imGrid2Upper1 + 1
        imGrid2Upper2 = 1
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddGrid3Data                   *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add data to Grid column table   *
'*                                                     *
'*******************************************************
Private Sub mAddGrid3Data(ilField As Integer)
    tmGrid3Data(imGrid3Upper1, imGrid3Upper2).iNextIndex1 = -1
    tmGrid3Data(imGrid3Upper1, imGrid3Upper2).iNextIndex2 = -1
    tmGrid3Data(imGrid3Upper1, imGrid3Upper2).iDDFFieldIndex = ilField
    tmGrid3Data(imGrid3Upper1, imGrid3Upper2).sValue = smFieldValue
    imGrid3Upper2 = imGrid3Upper2 + 1
    If imGrid3Upper2 > 32000 Then
        imGrid3Upper1 = imGrid3Upper1 + 1
        imGrid3Upper2 = 1
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddGrid4Data                   *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add data to Grid column table   *
'*                                                     *
'*******************************************************
Private Sub mAddGrid4Data(ilField As Integer)
    tmGrid4Data(imGrid4Upper1, imGrid4Upper2).iNextIndex1 = -1
    tmGrid4Data(imGrid4Upper1, imGrid4Upper2).iNextIndex2 = -1
    tmGrid4Data(imGrid4Upper1, imGrid4Upper2).iDDFFieldIndex = ilField
    tmGrid4Data(imGrid4Upper1, imGrid4Upper2).sValue = smFieldValue
    imGrid4Upper2 = imGrid4Upper2 + 1
    If imGrid4Upper2 > 32000 Then
        imGrid4Upper1 = imGrid4Upper1 + 1
        imGrid4Upper2 = 1
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddGrid5Data                   *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add data to Grid column table   *
'*                                                     *
'*******************************************************
Private Sub mAddGrid5Data(ilField As Integer)
    tmGrid5Data(imGrid5Upper1, imGrid5Upper2).iNextIndex1 = -1
    tmGrid5Data(imGrid5Upper1, imGrid5Upper2).iNextIndex2 = -1
    tmGrid5Data(imGrid5Upper1, imGrid5Upper2).iDDFFieldIndex = ilField
    tmGrid5Data(imGrid5Upper1, imGrid5Upper2).sValue = smFieldValue
    imGrid5Upper2 = imGrid5Upper2 + 1
    If imGrid5Upper2 > 32000 Then
        imGrid5Upper1 = imGrid5Upper1 + 1
        imGrid5Upper2 = 1
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddGrid6Data                   *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add data to Grid column table   *
'*                                                     *
'*******************************************************
Private Sub mAddGrid6Data(ilField As Integer)
    tmGrid6Data(imGrid6Upper1, imGrid6Upper2).iNextIndex1 = -1
    tmGrid6Data(imGrid6Upper1, imGrid6Upper2).iNextIndex2 = -1
    tmGrid6Data(imGrid6Upper1, imGrid6Upper2).iDDFFieldIndex = ilField
    tmGrid6Data(imGrid6Upper1, imGrid6Upper2).sValue = smFieldValue
    imGrid6Upper2 = imGrid6Upper2 + 1
    If imGrid6Upper2 > 32000 Then
        imGrid6Upper1 = imGrid6Upper1 + 1
        imGrid6Upper2 = 1
    End If
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
    Dim ilCalType As Integer
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = edcField.Text
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
'*      Procedure Name:gBtrvErrorMsg                   *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Show error btrieve message      *
'*                                                     *
'*******************************************************
Private Sub mBtrvErrorMsg(ilErrCode As Integer, slMess As String)
'
'   slMess = "Open Arf"
'   gBtrvErrorMsg ilRet, slMess
'   Where:
'       ilErrCode (I)- Error number from btrieve
'       slMess (I) - Message to be added to the btrieve error message
'       MainForm (I)- Name of Form to unload if error exists
'
    Dim slErrMsg As String
    If ilErrCode = BTRV_ERR_NONE Then
        Exit Sub
    End If
    slErrMsg = "Btrieve: "
    Select Case ilErrCode
        Case BTRV_ERR_INVALID_OP    '1
            slErrMsg = slErrMsg & "Invalid Operation (1)"
        Case BTRV_ERR_IO_ERR        '2
            slErrMsg = slErrMsg & "I/O Error (2)"
        Case BTRV_ERR_NOT_OPEN      '3
            slErrMsg = slErrMsg & "File not Opened (3)"
        Case BTRV_ERR_KEY_NOT_FOUND '4
            slErrMsg = slErrMsg & "Key Value not Found (4)"
        Case BTRV_ERR_DUPLICATE_KEY '5
            slErrMsg = slErrMsg & "Duplicate Key Value (5)"
        Case BTRV_ERR_INVALID_KEY   '6
            slErrMsg = slErrMsg & "Invalid Key Number (6)"
        Case BTRV_ERR_DIFF_KEY      '7
            slErrMsg = slErrMsg & "Different Key Number (7)"
        Case BTRV_ERR_INVALID_POS '8
            slErrMsg = slErrMsg & "Invalid Positioning (8)"
        Case BTRV_ERR_END_OF_FILE '9
            slErrMsg = slErrMsg & "End of File (9)"
        Case BTRV_ERR_MOD_KEY_VALUE '10
            slErrMsg = slErrMsg & "Modifiable Key Value Error (10)"
        Case BTRV_ERR_INVALID_FNAME '11
            slErrMsg = slErrMsg & "Invalid File Name (11)"
        Case BTRV_ERR_FILE_NOT_FOUND    '12
            slErrMsg = slErrMsg & "File Not Found (12)"
        Case BTRV_ERR_EXT_FILE      '13
            slErrMsg = slErrMsg & "Extended File Error (13)"
        Case BTRV_ERR_PREIMAGE_OPEN '14
            slErrMsg = slErrMsg & "Per-image Open Error (14)"
        Case BTRV_ERR_PREIMAGE_IO   '15
            slErrMsg = slErrMsg & "Pre-image I/O Error (15)"
        Case BTRV_ERR_EXPANSION '16
            slErrMsg = slErrMsg & "Expansion Error (16)"
        Case BTRV_ERR_CLOSE '17
            slErrMsg = slErrMsg & "Close Error (17)"
        Case BTRV_ERR_DISK_FULL '18
            slErrMsg = slErrMsg & "Disk Full (18)"
        Case BTRV_ERR_UNRECOVERABLE '19
            slErrMsg = slErrMsg & "Unrecoverable Error (19)"
        Case BTRV_ERR_MGR_INACTIVE  '20
            slErrMsg = slErrMsg & "Record Manager Inactive (20)"
        Case BTRV_ERR_KEYBUF_LENGTH '21
            slErrMsg = slErrMsg & "Key Buffer Too Short (21)"
        Case BTRV_ERR_DATABUF_LENGTH    '22
            slErrMsg = slErrMsg & "Data Buffer Length (22)"
        Case BTRV_ERR_POSBLK_LENGTH '23
            slErrMsg = slErrMsg & "Position Block Length (23)"
        Case BTRV_ERR_PAGESIZE  '24
            slErrMsg = slErrMsg & "Page Size Error (24)"
        Case BTRV_ERR_CREATE_IO '25
            slErrMsg = slErrMsg & "Create I/O Error"
        Case BTRV_ERR_NUMKEYS   '26
            slErrMsg = slErrMsg & "Number of Keys (26)"
        Case BTRV_ERR_INVALID_KEYPOS    '27
            slErrMsg = slErrMsg & "Invalid Key Position (27)"
        Case BTRV_ERR_REC_LENGTH    '28
            slErrMsg = slErrMsg & "Invalid Record Length (28)"
        Case BTRV_ERR_KEY_LENGTH    '29
            slErrMsg = slErrMsg & "Invalid Key Length (29)"
        Case BTRV_ERR_NOT_BTRV_FILE '30
            slErrMsg = slErrMsg & "Not a Btrieve File (30)"
        Case BTRV_ERR_ALREADY_EXTD  '31
            slErrMsg = slErrMsg & "File Already Extended (31)"
        Case BTRV_ERR_EXTD_IO   '32
            slErrMsg = slErrMsg & "Extend I/O Error (32)"
        Case BTRV_ERR_INVALID_EXT_NAME  '34
            slErrMsg = slErrMsg & "Invalid Extension Name (34)"
        Case BTRV_ERR_DIRECTORY '35
            slErrMsg = slErrMsg & "Directory Error (35)"
        Case BTRV_ERR_TRANSACTION   '36
            slErrMsg = slErrMsg & "Transaction Error"
        Case BTRV_ERR_TRANS_ACTIVE  '37
            slErrMsg = slErrMsg & "Transaction in Active (37)"
        Case BTRV_ERR_TRANS_FILE_IO '38
            slErrMsg = slErrMsg & "Transaction Control File I/O Error (38)"
        Case BTRV_ERR_END_ABORT_TRANS   '39
            slErrMsg = slErrMsg & "End/Abort Transaction Error (39)"
        Case BTRV_ERR_TRANS_MAX_FILES   '40
            slErrMsg = slErrMsg & "Transaction Max Files (40)"
        Case BTRV_ERR_OP_NOT_ALLOWED    '41
            slErrMsg = slErrMsg & "Operation Not allowed (41)"
        Case BTRV_ERR_ACCEL_ACCESS  '42
            slErrMsg = slErrMsg & "Incomplete Accelerated Access (42)"
        Case BTRV_ERR_INVALID_REC_ADDR  '43
            slErrMsg = slErrMsg & "Invalid Record Address (43)"
        Case BTRV_ERR_NULL_KEY_PATH '44
            slErrMsg = slErrMsg & "Null Key Path (44)"
        Case BTRV_ERR_INCON_KEY_FLAGS   '45
            slErrMsg = slErrMsg & "Inconsistent Key Flags (45)"
        Case BTRV_ERR_ACCESS_DENIED '46
            slErrMsg = slErrMsg & "Access to File Denied (46)"
        Case BTRV_ERR_MAX_OPEN_FILES    '47
            slErrMsg = slErrMsg & "Maximum Open Files (47)"
        Case BTRV_ERR_INVALID_ALT_SEQ   '48
            slErrMsg = slErrMsg & "Invalid Alternate Sequence Definition (48)"
        Case BTRV_ERR_KEY_TYPE  '49
            slErrMsg = slErrMsg & "Key Type Error"
        Case BTRV_ERR_OWNER_SET '50
            slErrMsg = slErrMsg & "Owner Already Set (50)"
        Case BTRV_ERR_INVALID_OWNER '51
            slErrMsg = slErrMsg & "Invalid Owner (51)"
        Case BTRV_ERR_WRITE_CACHE   '52
            slErrMsg = slErrMsg & "Error Writing Cache (52)"
        Case BTRV_ERR_INVALID_INTF  '53
            slErrMsg = slErrMsg & "Invalid Interface (53)"
        Case BTRV_ERR_VARIABLE_PAGE '54
            slErrMsg = slErrMsg & "Variable Page Error (54)"
        Case 55
            slErrMsg = slErrMsg & "Autoincrement Error (55)"
        Case BTRV_ERR_INCOMPLT_INDEX    '56
            slErrMsg = slErrMsg & "Incomplete Index"
        Case BTRV_ERR_EXPAND_MEM    '57
            slErrMsg = slErrMsg & "Expanded Memory Error (57)"
        Case BTRV_ERR_COMPBUF_SIZE  '58
            slErrMsg = slErrMsg & "Compression Buffer Too Short (58)"
        Case BTRV_ERR_FILE_EXISTS   '59
            slErrMsg = slErrMsg & "File Already Exists (59)"
        Case BTRV_ERR_REJECT_COUNT  '60
            slErrMsg = slErrMsg & "Reject Count Reached (60)"
        Case BTRV_ERR_WORK_SPACE_SIZE   '61
            slErrMsg = slErrMsg & "Work Space Too Small (61)"
        Case BTRV_ERR_INCORRECT_DESCP   '62
            slErrMsg = slErrMsg & "Incorrect Description (62)"
        Case BTRV_ERR_INVALID_EXTDBUF   '63
            slErrMsg = slErrMsg & "Invalid Extended Insert Buffer (63)"
        Case BTRV_ERR_FILTER_LIMIT  '64
            slErrMsg = slErrMsg & "Filter Limit Reached (64)"
        Case BTRV_ERR_INCOR_FLD_OFFSET  '65
            slErrMsg = slErrMsg & "Incorrect Field Offset (65)"
        Case BTRV_ERR_AUTO_TRANS_ABORT  '74
            slErrMsg = slErrMsg & "Automatic Transaction Abort (74)"
        Case BTRV_ERR_DEADLOCK_DETECT   '78
            slErrMsg = slErrMsg & "Deadlock Detected (78)"
        Case BTRV_ERR_CONFLICT  '80
            slErrMsg = slErrMsg & "Conflict Error (80)"
        Case BTRV_ERR_LOCK  '81
            slErrMsg = slErrMsg & "Lock Error (81)"
        Case BTRV_ERR_LOST_POS  '82
            slErrMsg = slErrMsg & "Lost Position (82)"
        Case BTRV_ERR_READ_TRANS    '83
            slErrMsg = slErrMsg & "Read Outside Transaction (83)"
        Case BTRV_ERR_REC_LOCKED    '84
            slErrMsg = slErrMsg & "Record Locked (84)"
        Case BTRV_ERR_FILE_LOCKED   '85
            slErrMsg = slErrMsg & "File Locked (85)"
        Case BTRV_ERR_FILE_TBL_FULL '86
            slErrMsg = slErrMsg & "File Table Full (86)"
        Case BTRV_ERR_HNDL_TBL_FULL '87
            slErrMsg = slErrMsg & "Handle Table Full (87)"
        Case BTRV_ERR_INCOM_MODE    '88
            slErrMsg = slErrMsg & "Incompatible Mode Error (88)"
        Case BTRV_ERR_REDIR_DEV_FULL    '90
            slErrMsg = slErrMsg & "Redirection Device Table Full (90)"
        Case BTRV_ERR_SERVER    '91
            slErrMsg = slErrMsg & "Serve Error (91)"
        Case BTRV_ERR_TRANS_TBL_FULL    '92
            slErrMsg = slErrMsg & "Transaction Table Full (92)"
        Case BTRV_ERR_INCOM_LOCK    '93
            slErrMsg = slErrMsg & "Incompatible Lock Type (93)"
        Case BTRV_ERR_PERMISSION    '94
            slErrMsg = slErrMsg & "Permission Error (94)"
        Case BTRV_ERR_SESSION   '95
            slErrMsg = slErrMsg & "Session No Longer Valid (95)"
        Case BTRV_ERR_COMM_ENV  '96
            slErrMsg = slErrMsg & "Communications Environment Error (96)"
        Case BTRV_ERR_DATA_MSGSIZE  '97
            slErrMsg = slErrMsg & "Data Message too Small (97)"
        Case BTRV_ERR_INTERNAL_TRANS    '98
            slErrMsg = slErrMsg & "Internal Transaction Error (98)"
    End Select
    slErrMsg = slErrMsg & ": " & slMess
    MsgBox slErrMsg, vbOkOnly + vbCritical + vbApplicationModal, "Btrieve Error"
    mTerminate
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCenterStdAlone                 *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Center form                     *
'*                                                     *
'*******************************************************
Private Sub mCenterStdAlone(Frm As Form)
    Frm.Move (Screen.Width - Frm.Width) \ 2, (Screen.Height - Frm.Height) \ 2 + 175 '+ Screen.Height \ 10
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
    tmcScroll.Enabled = False   'incase dropped within scroll area
    imcTrash.Picture = imcTrashClosed.Picture
    imcTrash.Enabled = False
    imcTrash.Visible = False
    imDragIndexSrce = -1
    imDragIndexDest = -1
    imDragSrce = -1
    imDragScroll = -1
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDragOver                       *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process drag over              *
'*                                                     *
'*******************************************************
Private Sub mDragOver(ilObject As Integer, State As Integer, Index As Integer)
    If imDragSrce = DRAGEDIT Then
        Exit Sub
    End If
    imDragScroll = ilObject
    If State = vbEnter Then
        tmcScroll.Enabled = True
        imScrollDir = Index
        Exit Sub
    End If
    If State = vbLeave Then
        tmcScroll.Enabled = False
        imDragScroll = -1
        Exit Sub
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mExtractFields                  *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Extract fields from record      *
'*                                                     *
'*******************************************************
Private Sub mExtractFields()
    Dim ilCol As Integer
    Dim ilLoop As Integer
    Dim ilField As Integer
    Dim ilColIndex1 As Integer
    Dim ilColIndex2 As Integer
    Dim ilTColIndex1 As Integer
    Dim ilTColIndex2 As Integer
    Dim ilGridType As Integer
    Dim ilPos As Integer
    Dim slFieldName As String
    Dim slFieldInfo As String
    Dim slFieldDataType As String
    Dim slSvDataType As String
    Dim ilRet As Integer
    Dim ilSize As Integer
    Dim ilFileID As Integer
    Dim slFileName As String
    ilCol = 1
    For ilLoop = 0 To lbcEditReplace(2).ListCount - 2 Step 1
        slFieldInfo = lbcEditReplace(2).List(ilLoop)
        ilRet = gParseItem(slFieldInfo, 1, "|", slFileName)
        ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
        ilRet = gParseItem(slFieldInfo, 4, "|", slFieldDataType)
        If ilRet <> CP_MSG_NONE Then
            slFieldDataType = ""
            If (StrComp(Trim$(slFileName), "Mnf.Btr", 1) = 0) And (StrComp(Trim$(slFieldName), "mnfRPU", 1) = 0) Then
                slFieldDataType = Chr(101)
            End If
            If (StrComp(Trim$(slFileName), "Mnf.Btr", 1) = 0) And (StrComp(Trim$(slFieldName), "mnfSSComm", 1) = 0) Then
                slFieldDataType = Chr(101)
            End If
        Else
            If StrComp(slFieldDataType, "Length of Time", 1) = 0 Then
                slFieldDataType = Chr(100)
            End If
            If StrComp(slFieldDataType, "Hex Field Dump", 1) = 0 Then
                slFieldDataType = Chr(101)
            End If
        End If
        ilFileID = mGetFileID(slFileName)
        For ilField = 0 To UBound(tmDDFField) - 1 Step 1
            If (StrComp(Trim$(slFieldName), Trim$(tmDDFField(ilField).sName), 1) = 0) And (ilFileID = tmDDFField(ilField).iFileID) Then
                slSvDataType = tmDDFField(ilField).sDataType
                If slFieldDataType <> "" Then
                    tmDDFField(ilField).sDataType = slFieldDataType
                End If
                If tmDDFField(ilField).iFileID = imFileID Then  'Key file
                    mGetField tmDDFField(ilField), 0
                Else
                    mGetFieldByKey tmDDFField(ilField)
                End If
                ilSize = mFieldSize(tmDDFField(ilField))
                tmDDFField(ilField).sDataType = slSvDataType
                If ilCol = 1 Then
                    tmRowLinkList(imRecUpper1, imRecUpper2).iRecIndex1 = tmRecInfo(imRecUpper1, imRecUpper2).iRecIndex1
                    tmRowLinkList(imRecUpper1, imRecUpper2).iRecIndex2 = tmRecInfo(imRecUpper1, imRecUpper2).iRecIndex2
                    tmRecInfo(imRecUpper1, imRecUpper2).iRowLinkIndex1 = imRecUpper1
                    tmRecInfo(imRecUpper1, imRecUpper2).iRowLinkIndex2 = imRecUpper2
                    If ilSize <= GRID1SIZE Then
                        tmRowLinkList(imRecUpper1, imRecUpper2).iGridType = 1
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex1 = imGrid1Upper1
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex2 = imGrid1Upper2
                        mAddGrid1Data ilField
                    ElseIf ilSize <= GRID2SIZE Then
                        tmRowLinkList(imRecUpper1, imRecUpper2).iGridType = 2
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex1 = imGrid2Upper1
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex2 = imGrid2Upper2
                        mAddGrid2Data ilField
                    ElseIf ilSize <= GRID3SIZE Then
                        tmRowLinkList(imRecUpper1, imRecUpper2).iGridType = 3
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex1 = imGrid3Upper1
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex2 = imGrid3Upper2
                        mAddGrid3Data ilField
                    ElseIf ilSize <= GRID4SIZE Then
                        tmRowLinkList(imRecUpper1, imRecUpper2).iGridType = 4
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex1 = imGrid4Upper1
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex2 = imGrid4Upper2
                        mAddGrid4Data ilField
                    ElseIf ilSize <= GRID5SIZE Then
                        tmRowLinkList(imRecUpper1, imRecUpper2).iGridType = 5
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex1 = imGrid5Upper1
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex2 = imGrid5Upper2
                        mAddGrid5Data ilField
                    ElseIf ilSize <= GRID6SIZE Then
                        tmRowLinkList(imRecUpper1, imRecUpper2).iGridType = 61
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex1 = imGrid6Upper1
                        tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex2 = imGrid6Upper2
                        mAddGrid6Data ilField
                    End If
                Else
                    ilColIndex1 = tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex1
                    ilColIndex2 = tmRowLinkList(imRecUpper1, imRecUpper2).iFirstIndex2
                    ilGridType = tmRowLinkList(imRecUpper1, imRecUpper2).iGridType
                    Do While ilColIndex1 > 0
                        Select Case ilGridType
                            Case 1
                                If tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex1 > 0 Then
                                    ilGridType = tmGrid1Data(ilColIndex1, ilColIndex2).iNextGridType
                                    ilTColIndex1 = tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex1
                                    ilTColIndex2 = tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex2
                                Else
                                    If ilSize <= GRID1SIZE Then
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextGridType = 1
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid1Upper1
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid1Upper2
                                        mAddGrid1Data ilField
                                    ElseIf ilSize <= GRID2SIZE Then
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextGridType = 2
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid2Upper1
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid2Upper2
                                        mAddGrid2Data ilField
                                    ElseIf ilSize <= GRID3SIZE Then
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextGridType = 3
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid3Upper1
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid3Upper2
                                        mAddGrid3Data ilField
                                    ElseIf ilSize <= GRID4SIZE Then
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextGridType = 4
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid4Upper1
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid4Upper2
                                        mAddGrid4Data ilField
                                    ElseIf ilSize <= GRID5SIZE Then
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextGridType = 5
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid5Upper1
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid5Upper2
                                        mAddGrid5Data ilField
                                    ElseIf ilSize <= GRID6SIZE Then
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextGridType = 6
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid6Upper1
                                        tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid6Upper2
                                        mAddGrid6Data ilField
                                    End If
                                    Exit Do
                                End If
                            Case 2
                                If tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex1 > 0 Then
                                    ilGridType = tmGrid2Data(ilColIndex1, ilColIndex2).iNextGridType
                                    ilTColIndex1 = tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex1
                                    ilTColIndex2 = tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex2
                                Else
                                    If ilSize <= GRID1SIZE Then
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextGridType = 1
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid1Upper1
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid1Upper2
                                        mAddGrid1Data ilField
                                    ElseIf ilSize <= GRID2SIZE Then
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextGridType = 2
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid2Upper1
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid2Upper2
                                        mAddGrid2Data ilField
                                    ElseIf ilSize <= GRID3SIZE Then
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextGridType = 3
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid3Upper1
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid3Upper2
                                        mAddGrid3Data ilField
                                    ElseIf ilSize <= GRID4SIZE Then
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextGridType = 4
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid4Upper1
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid4Upper2
                                        mAddGrid4Data ilField
                                    ElseIf ilSize <= GRID5SIZE Then
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextGridType = 5
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid5Upper1
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid5Upper2
                                        mAddGrid5Data ilField
                                    ElseIf ilSize <= GRID6SIZE Then
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextGridType = 6
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid6Upper1
                                        tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid6Upper2
                                        mAddGrid6Data ilField
                                    End If
                                    Exit Do
                                End If
                            Case 3
                                If tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex1 > 0 Then
                                    ilGridType = tmGrid3Data(ilColIndex1, ilColIndex2).iNextGridType
                                    ilTColIndex1 = tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex1
                                    ilTColIndex2 = tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex2
                                Else
                                    If ilSize <= GRID1SIZE Then
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextGridType = 1
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid1Upper1
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid1Upper2
                                        mAddGrid1Data ilField
                                    ElseIf ilSize <= GRID2SIZE Then
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextGridType = 2
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid2Upper1
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid2Upper2
                                        mAddGrid2Data ilField
                                    ElseIf ilSize <= GRID3SIZE Then
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextGridType = 3
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid3Upper1
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid3Upper2
                                        mAddGrid3Data ilField
                                    ElseIf ilSize <= GRID4SIZE Then
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextGridType = 4
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid4Upper1
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid4Upper2
                                        mAddGrid4Data ilField
                                    ElseIf ilSize <= GRID5SIZE Then
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextGridType = 5
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid5Upper1
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid5Upper2
                                        mAddGrid5Data ilField
                                    ElseIf ilSize <= GRID6SIZE Then
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextGridType = 6
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid6Upper1
                                        tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid6Upper2
                                        mAddGrid6Data ilField
                                    End If
                                    Exit Do
                                End If
                            Case 4
                                If tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex1 > 0 Then
                                    ilGridType = tmGrid4Data(ilColIndex1, ilColIndex2).iNextGridType
                                    ilTColIndex1 = tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex1
                                    ilTColIndex2 = tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex2
                                Else
                                    If ilSize <= GRID1SIZE Then
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextGridType = 1
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid1Upper1
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid1Upper2
                                        mAddGrid1Data ilField
                                    ElseIf ilSize <= GRID2SIZE Then
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextGridType = 2
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid2Upper1
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid2Upper2
                                        mAddGrid2Data ilField
                                    ElseIf ilSize <= GRID3SIZE Then
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextGridType = 3
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid3Upper1
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid3Upper2
                                        mAddGrid3Data ilField
                                    ElseIf ilSize <= GRID4SIZE Then
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextGridType = 4
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid4Upper1
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid4Upper2
                                        mAddGrid4Data ilField
                                    ElseIf ilSize <= GRID5SIZE Then
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextGridType = 5
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid5Upper1
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid5Upper2
                                        mAddGrid5Data ilField
                                    ElseIf ilSize <= GRID6SIZE Then
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextGridType = 6
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid6Upper1
                                        tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid6Upper2
                                        mAddGrid6Data ilField
                                    End If
                                    Exit Do
                                End If
                            Case 5
                                If tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex1 > 0 Then
                                    ilGridType = tmGrid5Data(ilColIndex1, ilColIndex2).iNextGridType
                                    ilTColIndex1 = tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex1
                                    ilTColIndex2 = tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex2
                                Else
                                    If ilSize <= GRID1SIZE Then
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextGridType = 1
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid1Upper1
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid1Upper2
                                        mAddGrid1Data ilField
                                    ElseIf ilSize <= GRID2SIZE Then
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextGridType = 2
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid2Upper1
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid2Upper2
                                        mAddGrid2Data ilField
                                    ElseIf ilSize <= GRID3SIZE Then
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextGridType = 3
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid3Upper1
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid3Upper2
                                        mAddGrid3Data ilField
                                    ElseIf ilSize <= GRID4SIZE Then
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextGridType = 4
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid4Upper1
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid4Upper2
                                        mAddGrid4Data ilField
                                    ElseIf ilSize <= GRID5SIZE Then
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextGridType = 5
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid5Upper1
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid5Upper2
                                        mAddGrid5Data ilField
                                    ElseIf ilSize <= GRID6SIZE Then
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextGridType = 6
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid6Upper1
                                        tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid6Upper2
                                        mAddGrid6Data ilField
                                    End If
                                    Exit Do
                                End If
                            Case 6
                                If tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex1 > 0 Then
                                    ilGridType = tmGrid6Data(ilColIndex1, ilColIndex2).iNextGridType
                                    ilTColIndex1 = tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex1
                                    ilTColIndex2 = tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex2
                                Else
                                    If ilSize <= GRID1SIZE Then
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextGridType = 1
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid1Upper1
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid1Upper2
                                        mAddGrid1Data ilField
                                    ElseIf ilSize <= GRID2SIZE Then
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextGridType = 2
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid2Upper1
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid2Upper2
                                        mAddGrid2Data ilField
                                    ElseIf ilSize <= GRID3SIZE Then
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextGridType = 3
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid3Upper1
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid3Upper2
                                        mAddGrid3Data ilField
                                    ElseIf ilSize <= GRID4SIZE Then
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextGridType = 4
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid4Upper1
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid4Upper2
                                        mAddGrid4Data ilField
                                    ElseIf ilSize <= GRID5SIZE Then
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextGridType = 5
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid5Upper1
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid5Upper2
                                        mAddGrid5Data ilField
                                    ElseIf ilSize <= GRID6SIZE Then
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextGridType = 6
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex1 = imGrid6Upper1
                                        tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex2 = imGrid6Upper2
                                        mAddGrid6Data ilField
                                    End If
                                    Exit Do
                                End If
                        End Select
                        ilColIndex1 = ilTColIndex1
                        ilColIndex2 = ilTColIndex2
                    Loop
                End If
                ilCol = ilCol + 1
                Exit For
            End If
        Next ilField
    Next ilLoop
    'imRecUpper2 = imRecUpper2 + 1
    'If imRecUpper2 > 32000 Then
    '    imRecUpper1 = imRecUpper1 + 1
    '    imRecUpper2 = 1
    'End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFieldBuild                     *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list box with field    *
'*                     names                           *
'*                                                     *
'*******************************************************
Private Sub mFieldBuild(ilFileID As Integer, tlDDFField() As DDFFIELD)
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilMax As Integer
    Dim ilOffset As Integer
    Dim ilUpper As Integer
    
    'smDDFPath = lbcPath.Path & "\"
    Screen.MousePointer = vbHourglass
    ReDim tlField(0 To 0) As DDFFIELD
    ilUpper = UBound(tlField)
    hmFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hmFile, "", smDDFPath & "Field.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Screen.MousePointer = 0 'DEFAULT
        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
        Exit Sub
    End If
    lmActualRecLen = Len(tlField(ilUpper)) 'Len(tmRec)
    imBaseRecLen = lmActualRecLen
    tmFieldSrchKey1.iFileID = ilFileID
    'ilRet = btrGetGreaterOrEqual(hmFile, tlField(ilUpper), lmActualRecLen, tmFieldSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    ilRet = btrGetGreaterOrEqual(hmFile, tmRec, lmActualRecLen, tmFieldSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = 0 'DEFAULT
        mBtrvErrorMsg ilRet, "btrGetGreaterOrEqual: " & "FIELD.DDF"
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Exit Sub
    End If
    Do While (ilRet = BTRV_ERR_NONE)
        If imTerminate Then
            Screen.MousePointer = 0 'DEFAULT
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            mTerminate
            Exit Sub
        End If
        mGetFieldFields
        tlField(ilUpper) = tmGetDDFFIELD
        If (tlField(ilUpper).iFileID <> ilFileID) Then
            Exit Do
        End If
        If Asc(tlField(ilUpper).sDataType) <= 17 Then
            ilUpper = ilUpper + 1
            ReDim Preserve tlField(0 To ilUpper) As DDFFIELD
        End If
        lmActualRecLen = imBaseRecLen
        ilRet = btrGetNext(hmFile, tmRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hmFile)
    btrDestroy hmFile
    'Sort by offset
    ilMax = 0
    For ilLoop = 0 To UBound(tlField) - 1 Step 1
        If tlField(ilLoop).iOffset > ilMax Then
            ilMax = tlField(ilLoop).iOffset
        End If
    Next ilLoop
    ilUpper = UBound(tlDDFField)
    ilOffset = 0
    Do
        For ilLoop = 0 To UBound(tlField) - 1 Step 1
            If tlField(ilLoop).iOffset = ilOffset Then
                tlDDFField(ilUpper) = tlField(ilLoop)
                ilUpper = ilUpper + 1
                ReDim Preserve tlDDFField(0 To ilUpper) As DDFFIELD
                Exit For
            End If
        Next ilLoop
        ilOffset = ilOffset + 1
    Loop While ilOffset <= ilMax
    Screen.MousePointer = 0 'DEFAULT
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFieldPop                       *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list box with field    *
'*                     names                           *
'*                                                     *
'*******************************************************
Private Sub mFieldPop(tlDDFField() As DDFFIELD, ilFromIndex As Integer, ilToIndex As Integer, lbcField As Control, ilAddInfo As Integer)
    Dim ilLoop As Integer
    Dim slFileName As String
    Dim ilFileName As Integer
    Dim slFile As String
    Dim ilPos As Integer
    Dim ilFileID As Integer
    ilFileID = -1
    If ilFromIndex > ilToIndex Then
        Exit Sub
    End If
    'lbcField.Clear
    For ilLoop = ilFromIndex To ilToIndex Step 1
        If ilFileID <> tlDDFField(ilLoop).iFileID Then
            For ilFileName = 0 To UBound(tmDDFFile) - 1 Step 1
                If tlDDFField(ilLoop).iFileID = tmDDFFile(ilFileName).iFileID Then
                    'Extract File Name C:\xxx\xxx.btr or c:xxx.btr
                    slFile = Trim$(gRemoveZero(tmDDFFile(ilFileName).sLocation))
                    ilPos = InStr(slFile, ":")
                    If ilPos > 0 Then
                        slFile = Mid$(slFile, ilPos + 1)
                    End If
                    ilPos = InStr(slFile, "\")
                    Do While ilPos > 0
                        slFile = Mid$(slFile, ilPos + 1)
                        ilPos = InStr(slFile, "\")
                    Loop
                    ilFileID = tlDDFField(ilLoop).iFileID
                    Exit For
                End If
            Next ilFileName
        End If
        'Don't add ssfPas to field definitions
        If (Asc(tlDDFField(ilLoop).sDataType) <> 12) Or (Trim$(tlDDFField(ilLoop).sName) <> "ssfPAS") Then
            If ilAddInfo Then
                Select Case Asc(tlDDFField(ilLoop).sDataType)
                    Case 0  'String
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "String"
                    Case 1  'Integer
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "Integer"
                    Case 2  'IEEE Float
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "IEEE Float"
                    Case 3  'Btrieve Date
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "Btrieve Date"
                    Case 4  'Btrieve Time
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "Btrieve Time"
                    Case 5  'COBOL Decimal COMP-3
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "COBOL Dec COMP-3"
                    Case 6  'COBOL Money
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "COBOL Money"
                    Case 7  'Logical
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "Logical"
                    Case 8  'COBOL Numeric
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "COBOL Numeric"
                    Case 9  'BASIC Float
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "BASIC Float"
                    Case 10 'Pascal LString
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "Pascal LString"
                    Case 11 'C ZString
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "C ZString"
                    Case 12 'Variable Length Note
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "Variable Length"
                    Case 13 'LVar
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "LVar"
                    Case 14 'Unsigned Binary
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "Unsigned Binary"
                    Case 15 'AutoIncrement
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "AutoIncrement"
                    Case 16 'Bit
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "Bit"
                    Case 17 'COBOL Numeric STS
                        lbcField.AddItem slFile & "|" & Trim$(tlDDFField(ilLoop).sName) & "|" & "COBOL Numeric STS"
                End Select
            Else
                lbcField.AddItem Trim$(tlDDFField(ilLoop).sName)
            End If
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFieldSize                      *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine field size            *
'*                                                     *
'*******************************************************
Private Function mFieldSize(tlDDFField As DDFFIELD) As Integer
    Dim ilSize As Integer
    Select Case Asc(tlDDFField.sDataType)
        Case 0  'String
            ilSize = tlDDFField.iSize
        Case 1  'Integer
            If tlDDFField.iSize = 1 Then
                ilSize = 3
            ElseIf tlDDFField.iSize = 2 Then
                ilSize = 5
            ElseIf tlDDFField.iSize = 4 Then
                ilSize = 10
            End If
        Case 2  'IEEE Float
        Case 3  'Btrieve Date
            ilSize = 10
        Case 4  'Btrieve Time
            ilSize = 10
        Case 5  'COBOL Decimal COMP-3
            ilSize = 12
        Case 6  'COBOL Money
            ilSize = 12
        Case 7  'Logical
            ilSize = 5
        Case 8  'COBOL Numeric
        Case 9  'BASIC Float
        Case 10 'Pascal LString
        Case 11 'C ZString
            ilSize = tlDDFField.iSize
        Case 12 'Variable Length Note
            ilSize = tlDDFField.iSize
        Case 13 'LVar
            ilSize = tlDDFField.iSize
        Case 14 'Unsigned Binary
        Case 15 'AutoIncrement
            If ilSize = 2 Then
                ilSize = 5
            ElseIf ilSize = 4 Then
                ilSize = 10
            End If
        Case 16 'Bit
        Case 17 'COBOL Numeric STS
        Case 100  'Length
            ilSize = 9
        Case 101    'Hex Dump
            ilSize = 3 * tlDDFField.iSize - 1 'lmActualRecLen - 7
    End Select
    mFieldSize = ilSize
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFilePop                        *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build Table of file names       *
'*                                                     *
'*******************************************************
Private Sub mFileBuild(tlDDFFile() As DDFFILE)
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim slStr As String
    
    
    'slStr = lbcPath.Path
    'If right$(slStr, 1) <> "\" Then
    '    smDDFPath = lbcPath.Path & "\"
    'Else
    '    smDDFPath = lbcPath.Path
    'End If
    Screen.MousePointer = vbHourglass
    ReDim tlDDFFile(0 To 0) As DDFFILE
    ilUpper = UBound(tlDDFFile)
    hmFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hmFile, "", smDDFPath & "File.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Screen.MousePointer = 0 'DEFAULT
        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
        Exit Sub
    End If
    lmActualRecLen = Len(tlDDFFile(ilUpper)) 'Len(tmRec)
    imBaseRecLen = lmActualRecLen
    'ilRet = btrGetFirst(hmFile, tlDDFFile(ilUpper), lmActualRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    ilRet = btrGetFirst(hmFile, tmRec, lmActualRecLen, INDEXKEY0, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = 0 'DEFAULT
        mBtrvErrorMsg ilRet, "btrGetFirst: " & "FILE.DDF"
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Exit Sub
    End If
    Do While ilRet = BTRV_ERR_NONE
        If imTerminate Then
            Screen.MousePointer = 0 'DEFAULT
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            mTerminate
            Exit Sub
        End If
        mGetFileFields
        tlDDFFile(ilUpper) = tmGetDDFFile
        'If (StrComp(Trim$(tlDDFFile(ilUpper).sName), "X$Field", 1) <> 0) And (StrComp(Trim$(tlDDFFile(ilUpper).sName), "X$File", 1) <> 0) And (StrComp(Trim$(tlDDFFile(ilUpper).sName), "X$Index", 1) <> 0) Then
        If (InStr(1, Trim$(tlDDFFile(ilUpper).sName), "X$", 1) <> 1) Then
            ilUpper = ilUpper + 1
            ReDim Preserve tlDDFFile(0 To ilUpper) As DDFFILE
        End If
        lmActualRecLen = imBaseRecLen
        'ilRet = btrGetNext(hmFile, tlDDFFile(ilUpper), lmActualRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilRet = btrGetNext(hmFile, tmRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hmFile)
    btrDestroy hmFile
    Screen.MousePointer = 0 'DEFAULT
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFilePop                        *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list box with file     *
'*                     names                           *
'*                                                     *
'*******************************************************
Private Sub mFilePop(tlDDFFile() As DDFFILE, lbcFile As Control)
    Dim ilLoop As Integer
    lbcFile.Clear
    For ilLoop = 0 To UBound(tlDDFFile) - 1 Step 1
        lbcFile.AddItem Trim$(tlDDFFile(ilLoop).sName)
    Next ilLoop
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetField                       *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get field from record           *
'*                                                     *
'*******************************************************
Private Sub mGetField(tlDDFField As DDFFIELD, ilRecType As Integer)
'
'   ilRecType: 0= tmRec; 1=tgKeyRec
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilSize As Integer
    Dim ilValue As Integer
    Dim llValue As Long
    Dim ilDyMn As Integer
    Dim ilYear As Integer
    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim slMoney As String
    Dim slHex As String
    Dim blByte As BYTE4
    Dim blByte2 As Byte
    Dim slInt2 As INT2
    Dim slInt4 As INT4
    'Used with ePDNToStr
    'Dim slResult As String * 20
    'Dim ilPos As Integer
    ilIndex = tlDDFField.iOffset + 1
    ilSize = tlDDFField.iSize
    smFieldValue = "Error"
    Select Case Asc(tlDDFField.sDataType)
        Case 0  'String
            smFieldValue = ""
            For ilLoop = 1 To ilSize Step 1
                If ilRecType = 0 Then
                    'If Asc(tmRec.sChar(ilIndex)) = 0 Then
                    If (tmRec.sChar(ilIndex)) = 0 Then
                        smFieldValue = smFieldValue & " "
                    Else
                        'If Asc(tmRec.sChar(ilIndex)) > 127 Then
                        '    smFieldValue = smFieldValue & Chr(Asc(tmRec.sChar(ilIndex)) - 128)
                        'Else
                            smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                        'End If
                    End If
                Else
                    'If Asc(tgKeyRec.sChar(ilIndex)) = 0 Then
                    If (tgKeyRec.sChar(ilIndex)) = 0 Then
                        smFieldValue = smFieldValue & " "
                    Else
                        'smFieldValue = smFieldValue & tgKeyRec.sChar(ilIndex)
                        smFieldValue = smFieldValue & Chr(tgKeyRec.sChar(ilIndex))
                    End If
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 1  'Integer
            If ilSize = 1 Then
                If ilRecType = 0 Then
                    ilValue = tmRec.sChar(ilIndex) 'Asc(tmRec.sChar(ilIndex))
                Else
                    ilValue = tgKeyRec.sChar(ilIndex) 'Asc(tgKeyRec.sChar(ilIndex))
                End If
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 2 Then
                'If ilRecType = 0 Then
                '    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                'Else
                '    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
                'End If
                'HMemCpy ilValue, slInt2, 2
                ''ilValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                If ilRecType = 0 Then
                    'ilValue = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
                    HMemCpy ilValue, tmRec.sChar(ilIndex), 2
                Else
                    'ilValue = 256 * (tgKeyRec.sChar(ilIndex + 1)) + (tgKeyRec.sChar(ilIndex))
                    HMemCpy ilValue, tgKeyRec.sChar(ilIndex), 2
                End If
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 4 Then
                'If ilRecType = 0 Then
                '    slInt4.sInt4 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1) & tmRec.sChar(ilIndex + 2) & tmRec.sChar(ilIndex + 3)
                'Else
                '    slInt4.sInt4 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1) & tgKeyRec.sChar(ilIndex + 2) & tgKeyRec.sChar(ilIndex + 3)
                'End If
                'HMemCpy llValue, slInt4, 4
                ''llValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 3))) + Asc(tmRec.sChar(ilIndex + 2))) And &HFFFF&
                ''llValue = 32768 * llValue + (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                If ilRecType = 0 Then
                    ''llValue = 256 * CLng(tmRec.sChar(ilIndex + 3)) + tmRec.sChar(ilIndex + 2)
                    ''llValue = 65536 * llValue + 256 * CLng(tmRec.sChar(ilIndex + 1)) + tmRec.sChar(ilIndex)
                    'blByte.bByte(2) = tmRec.sChar(ilIndex + 2)
                    'blByte.bByte(3) = tmRec.sChar(ilIndex + 3)
                    'blByte.bByte(0) = tmRec.sChar(ilIndex + 0)
                    'blByte.bByte(1) = tmRec.sChar(ilIndex + 1)
                    'HMemCpy llValue, blByte, 4
                    HMemCpy llValue, tmRec.sChar(ilIndex), 4
                Else
                    'lValue = 256 * CLng(tgKeyRec.sChar(ilIndex + 3)) + tgKeyRec.sChar(ilIndex + 2)
                    'llValue = 65536 * llValue + 256 * CLng(tgKeyRec.sChar(ilIndex + 1)) + tgKeyRec.sChar(ilIndex)
                    HMemCpy llValue, tgKeyRec.sChar(ilIndex), 4
                End If
                smFieldValue = Trim$(Str$(llValue))
            End If
        Case 2  'IEEE Float
        Case 3  'Btrieve Date
            If ilRecType = 0 Then
                'ilDyMn = 256 * Asc(tmRec.sChar(ilIndex + 1)) + Asc(tmRec.sChar(ilIndex))
                'ilYear = 256 * Asc(tmRec.sChar(ilIndex + 3)) + Asc(tmRec.sChar(ilIndex + 2))
                ilDyMn = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
                ilYear = 256 * (tmRec.sChar(ilIndex + 3)) + (tmRec.sChar(ilIndex + 2))
            Else
                'ilDyMn = 256 * Asc(tgKeyRec.sChar(ilIndex + 1)) + Asc(tgKeyRec.sChar(ilIndex))
                'ilYear = 256 * Asc(tgKeyRec.sChar(ilIndex + 3)) + Asc(tgKeyRec.sChar(ilIndex + 2))
                ilDyMn = 256 * (tgKeyRec.sChar(ilIndex + 1)) + (tgKeyRec.sChar(ilIndex))
                ilYear = 256 * (tgKeyRec.sChar(ilIndex + 3)) + (tgKeyRec.sChar(ilIndex + 2))
            End If
            gUnpackDate ilDyMn, ilYear, smFieldValue
        Case 4  'Btrieve Time
            If ilRecType = 0 Then
                'ilHsSec = 256 * Asc(tmRec.sChar(ilIndex + 1)) + Asc(tmRec.sChar(ilIndex))
                'ilMinHr = 256 * Asc(tmRec.sChar(ilIndex + 3)) + Asc(tmRec.sChar(ilIndex + 2))
                ilHsSec = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
                ilMinHr = 256 * (tmRec.sChar(ilIndex + 3)) + (tmRec.sChar(ilIndex + 2))
            Else
                'ilHsSec = 256 * Asc(tgKeyRec.sChar(ilIndex + 1)) + Asc(tgKeyRec.sChar(ilIndex))
                'ilMinHr = 256 * Asc(tgKeyRec.sChar(ilIndex + 3)) + Asc(tgKeyRec.sChar(ilIndex + 2))
                ilHsSec = 256 * (tgKeyRec.sChar(ilIndex + 1)) + (tgKeyRec.sChar(ilIndex))
                ilMinHr = 256 * (tgKeyRec.sChar(ilIndex + 3)) + (tgKeyRec.sChar(ilIndex + 2))
            End If
            gUnpackTime ilHsSec, ilMinHr, "A", "1", smFieldValue
        Case 5  'COBOL Decimal COMP-3
            slMoney = ""
            For ilLoop = 1 To ilSize Step 1
                If ilRecType = 0 Then
                    'slMoney = slMoney & tmRec.sChar(ilIndex)
                    slMoney = slMoney & Chr(tmRec.sChar(ilIndex))
                Else
                    'slMoney = slMoney & tgKeyRec.sChar(ilIndex)
                    slMoney = slMoney & Chr(tgKeyRec.sChar(ilIndex))
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
            'ePDNToStr slMoney, ilSize, Asc(tlDDFField.sDec), slResult
            'ilPos = InStr(slResult, Chr(0))
            'If ilPos > 0 Then
            '    smFieldValue = Left$(slResult, ilPos - 1)
            'Else
            '    smFieldValue = Trim$(slResult)
            'End If
            gPDNToStr slMoney, Asc(tlDDFField.sDec), smFieldValue
        Case 6  'COBOL Money
            slMoney = ""
            For ilLoop = 1 To ilSize Step 1
                If ilRecType = 0 Then
                    'slMoney = slMoney & tmRec.sChar(ilIndex)
                    slMoney = slMoney & Chr(tmRec.sChar(ilIndex))
                Else
                    'slMoney = slMoney & tgKeyRec.sChar(ilIndex)
                    slMoney = slMoney & Chr(tgKeyRec.sChar(ilIndex))
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
            'ePDNToStr slMoney, ilSize, Asc(tlDDFField.sDec), slResult
            'ilPos = InStr(slResult, Chr(0))
            'If ilPos > 0 Then
            '    smFieldValue = Left$(slResult, ilPos - 1)
            'Else
            '    smFieldValue = Trim$(slResult)
            'End If
            gPDNToStr slMoney, Asc(tlDDFField.sDec), smFieldValue
        Case 7  'Logical
            If ilSize = 1 Then
                If ilRecType = 0 Then
                    'ilValue = Asc(tmRec.sChar(ilIndex))
                    ilValue = (tmRec.sChar(ilIndex))
                Else
                    'ilValue = Asc(tgKeyRec.sChar(ilIndex))
                    ilValue = (tgKeyRec.sChar(ilIndex))
                End If
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 2 Then
                'If ilRecType = 0 Then
                '    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                'Else
                '    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
                'End If
                'HMemCpy ilValue, slInt2, 2
                ''ilValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                If ilRecType = 0 Then
                    'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
                    HMemCpy ilValue, tmRec.sChar(ilIndex), 2
                Else
                    'ilValue = (256 * (tgKeyRec.sChar(ilIndex + 1))) + (tgKeyRec.sChar(ilIndex))
                    HMemCpy ilValue, tgKeyRec.sChar(ilIndex), 2
                End If
                smFieldValue = Trim$(Str$(ilValue))
            End If
        Case 8  'COBOL Numeric
        Case 9  'BASIC Float
        Case 10 'Pascal LString
            If ilRecType = 0 Then
                'ilValue = Asc(tmRec.sChar(ilIndex))
                ilValue = tmRec.sChar(ilIndex)
            Else
                'ilValue = Asc(tgKeyRec.sChar(ilIndex))
                ilValue = tgKeyRec.sChar(ilIndex)
            End If
            smFieldValue = ""
            ilIndex = ilIndex + 1
            For ilLoop = 1 To ilValue Step 1
                If ilRecType = 0 Then
                    'If Asc(tmRec.sChar(ilIndex)) = 0 Then
                    If (tmRec.sChar(ilIndex)) = 0 Then
                        smFieldValue = smFieldValue & " "
                    Else
                        'smFieldValue = smFieldValue & tmRec.sChar(ilIndex)
                        smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                    End If
                Else
                    'If Asc(tgKeyRec.sChar(ilIndex)) = 0 Then
                    If (tgKeyRec.sChar(ilIndex)) = 0 Then
                        smFieldValue = smFieldValue & " "
                    Else
                        'smFieldValue = smFieldValue & tgKeyRec.sChar(ilIndex)
                        smFieldValue = smFieldValue & Chr(tgKeyRec.sChar(ilIndex))
                    End If
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 11 'C ZString
            smFieldValue = ""
            If ilRecType = 0 Then
                'Do While Asc(tmRec.sChar(ilIndex)) <> 0
                '    smFieldValue = smFieldValue & tmRec.sChar(ilIndex)
                '    ilIndex = ilIndex + 1
                'Loop
                Do While (tmRec.sChar(ilIndex)) <> 0
                    smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                    ilIndex = ilIndex + 1
                Loop
            Else
                'Do While Asc(tgKeyRec.sChar(ilIndex)) <> 0
                '    smFieldValue = smFieldValue & tgKeyRec.sChar(ilIndex)
                '    ilIndex = ilIndex + 1
                'Loop
                Do While (tgKeyRec.sChar(ilIndex)) <> 0
                    smFieldValue = smFieldValue & Chr(tgKeyRec.sChar(ilIndex))
                    ilIndex = ilIndex + 1
                Loop
            End If
        Case 12 'Variable Length Note
        Case 13 'LVar (length, Characters, 0, 0); Length does not include zeros
            'If ilRecType = 0 Then
            '    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
            'Else
            '    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
            'End If
            'HMemCpy ilValue, slInt2, 2
            If ilRecType = 0 Then
                'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
                HMemCpy ilValue, tmRec.sChar(ilIndex), 2
            Else
                'ilValue = (256 * (tgKeyRec.sChar(ilIndex + 1))) + (tgKeyRec.sChar(ilIndex))
                HMemCpy ilValue, tgKeyRec.sChar(ilIndex), 2
            End If
            If ilValue > tlDDFField.iSize - 4 Then
                imLVarLengthError = True
                ilValue = tlDDFField.iSize - 4
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = ilValue And &HFF
                    'tmRec.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                    HMemCpy tmRec.sChar(ilIndex), ilValue, 2
                Else
                    'tgKey.sChar(ilIndex) = ilValue And &HFF
                    'tgKey.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                    HMemCpy tgKey.sChar(ilIndex), ilValue, 2
                End If
                tmRec.sChar(ilIndex + tlDDFField.iSize - 2) = Asc(Chr(0))
                tmRec.sChar(ilIndex + tlDDFField.iSize - 1) = Asc(Chr(0))
                
            End If
            smFieldValue = ""
            ilIndex = ilIndex + 2
            For ilLoop = 1 To ilValue Step 1
                If ilRecType = 0 Then
                    'If Asc(tmRec.sChar(ilIndex)) = 0 Then
                    '    smFieldValue = smFieldValue & " "
                    'Else
                    '    smFieldValue = smFieldValue & tmRec.sChar(ilIndex)
                    'End If
                    If (tmRec.sChar(ilIndex)) = 0 Then
                        smFieldValue = smFieldValue & " "
                    Else
                        smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                    End If
                Else
                    'If Asc(tgKeyRec.sChar(ilIndex)) = 0 Then
                    '    smFieldValue = smFieldValue & " "
                    'Else
                    '    smFieldValue = smFieldValue & tgKeyRec.sChar(ilIndex)
                    'End If
                    If (tgKeyRec.sChar(ilIndex)) = 0 Then
                        smFieldValue = smFieldValue & " "
                    Else
                        smFieldValue = smFieldValue & Chr(tgKeyRec.sChar(ilIndex))
                    End If
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 14 'Unsigned Binary
            'If ilRecType = 0 Then
            '    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
            'Else
            '    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
            'End If
            'HMemCpy ilValue, slInt2, 2
            If ilRecType = 0 Then
                'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
                HMemCpy ilValue, tmRec.sChar(ilIndex), 2
            Else
                'ilValue = (256 * (tgKeyRec.sChar(ilIndex + 1))) + (tgKeyRec.sChar(ilIndex))
                HMemCpy ilValue, tgKeyRec.sChar(ilIndex), 2
            End If
            llValue = gUnsignIntToLong(ilValue)
            smFieldValue = Trim$(Str$(llValue))
        Case 15 'AutoIncrement
            If ilSize = 2 Then
                'If ilRecType = 0 Then
                '    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                'Else
                '    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
                'End If
                'HMemCpy ilValue, slInt2, 2
                ''ilValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                If ilRecType = 0 Then
                    'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
                    HMemCpy ilValue, tmRec.sChar(ilIndex), 2
                Else
                    'ilValue = (256 * (tgKeyRec.sChar(ilIndex + 1))) + (tgKeyRec.sChar(ilIndex))
                    HMemCpy ilValue, tgKeyRec.sChar(ilIndex), 2
                End If
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 4 Then
                'If ilRecType = 0 Then
                '    slInt4.sInt4 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1) & tmRec.sChar(ilIndex + 2) & tmRec.sChar(ilIndex + 3)
                'Else
                '    slInt4.sInt4 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1) & tgKeyRec.sChar(ilIndex + 2) & tgKeyRec.sChar(ilIndex + 3)
                'End If
                'HMemCpy llValue, slInt4, 4
                ''llValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 3))) + Asc(tmRec.sChar(ilIndex + 2))) And &HFFFF&
                ''llValue = 32768 * llValue + (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                If ilRecType = 0 Then
                    'llValue = 256 * CLng(tmRec.sChar(ilIndex + 3)) + tmRec.sChar(ilIndex + 2)
                    'llValue = 65536 * llValue + 256 * CLng(tmRec.sChar(ilIndex + 1)) + tmRec.sChar(ilIndex)
                    HMemCpy llValue, tmRec.sChar(ilIndex), 4
                Else
                    'llValue = 256 * CLng(tgKeyRec.sChar(ilIndex + 3)) + tgKeyRec.sChar(ilIndex + 2)
                    'llValue = 65536 * llValue + 256 * CLng(tgKeyRec.sChar(ilIndex + 1)) + tgKeyRec.sChar(ilIndex)
                    HMemCpy llValue, tgKeyRec.sChar(ilIndex), 4
                End If
                smFieldValue = Trim$(Str$(llValue))
            Else
                smFieldValue = "Error"
            End If
        Case 16 'Bit
        Case 17 'COBOL Numeric STS
        Case 100  'Length
            If ilRecType = 0 Then
                'ilHsSec = 256 * Asc(tmRec.sChar(ilIndex + 1)) + Asc(tmRec.sChar(ilIndex))
                'ilMinHr = 256 * Asc(tmRec.sChar(ilIndex + 3)) + Asc(tmRec.sChar(ilIndex + 2))
                ilHsSec = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
                ilMinHr = 256 * (tmRec.sChar(ilIndex + 3)) + (tmRec.sChar(ilIndex + 2))
            Else
                'ilHsSec = 256 * Asc(tgKeyRec.sChar(ilIndex + 1)) + Asc(tgKeyRec.sChar(ilIndex))
                'ilMinHr = 256 * Asc(tgKeyRec.sChar(ilIndex + 3)) + Asc(tgKeyRec.sChar(ilIndex + 2))
                ilHsSec = 256 * (tgKeyRec.sChar(ilIndex + 1)) + (tgKeyRec.sChar(ilIndex))
                ilMinHr = 256 * (tgKeyRec.sChar(ilIndex + 3)) + (tgKeyRec.sChar(ilIndex + 2))
            End If
            gUnpackLength ilHsSec, ilMinHr, "3", False, smFieldValue
        Case 101
            smFieldValue = ""
            For ilLoop = 1 To ilSize Step 1
                If ilRecType = 0 Then
                    'ilValue = Asc(tmRec.sChar(ilIndex))
                    ilValue = (tmRec.sChar(ilIndex))
                Else
                    'ilValue = Asc(tgKeyRec.sChar(ilIndex))
                    ilValue = (tgKeyRec.sChar(ilIndex))
                End If
                slHex = Hex$(ilValue)
                Do While Len(slHex) < 2
                    slHex = "0" & slHex
                Loop
                smFieldValue = smFieldValue & slHex & " "
                ilIndex = ilIndex + 1
            Next ilLoop
            smFieldValue = Trim$(smFieldValue)
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetKeyData                     *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Key field data from key file    *
'*                                                     *
'*******************************************************
Private Sub mGetFieldByKey(tlDDFField As DDFFIELD)
    Dim ilLoop As Integer
    Dim tlKeyDDFField As DDFFIELD
    Dim ilField As Integer
    Dim ilRet As Integer
    Dim ilKey As Integer
    Dim ilRecLen As Integer
    Dim llRecLen As Long
    Dim ilMatch As Integer
    Dim slTestValue As String
    Dim ilTest As Integer
    For ilLoop = 1 To Len(tgKey) Step 1
        tgKey.sChar(ilLoop) = 0 'Chr$(0)
    Next ilLoop
    For ilLoop = 0 To UBound(tmLink) - 1 Step 1
        If tmLink(ilLoop).iToFileID = tlDDFField.iFileID Then
            ilKey = tmLink(ilLoop).iKeyNo
            For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                If (tmDDFField(ilField).iFileID = tmLink(ilLoop).iFromFileID) And (tmDDFField(ilField).iFieldID = tmLink(ilLoop).iFromFieldID) Then
                    mGetField tmDDFField(ilField), 0
                    Exit For
                End If
            Next ilField
            tlKeyDDFField.iOffset = tmLink(ilLoop).iKeyOffset
            For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                If (tmDDFField(ilField).iFileID = tmLink(ilLoop).iToFileID) And (tmDDFField(ilField).iFieldID = tmLink(ilLoop).iToFieldID) Then
                    tlKeyDDFField.iOffset = tmDDFField(ilField).iOffset
                    tlKeyDDFField.iSize = tmDDFField(ilField).iSize
                    tlKeyDDFField.sDataType = tmDDFField(ilField).sDataType
                    tlKeyDDFField.sDec = tmDDFField(ilField).sDec
                    mPutField tlKeyDDFField, 1   'Store smFieldValue into record (tgKey)
                    Exit For
                End If
            Next ilField
        End If
    Next ilLoop
    'Key read, then extract
    For ilLoop = 0 To UBound(tmKeyFile) - 1 Step 1
        If tmKeyFile(ilLoop).iFileID = tlDDFField.iFileID Then
            llRecLen = CLng(Len(tgKeyRec))
            ilRet = btrGetGreaterOrEqual(tmKeyFile(ilLoop).hFile, tgKeyRec, llRecLen, tgKey, ilKey, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If ilRet = 0 Then
                'Test that key field match- if not blank smFieldValue
                ilMatch = True
                For ilTest = 0 To UBound(tmLink) - 1 Step 1
                    If tmLink(ilTest).iToFileID = tlDDFField.iFileID Then
                        For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                            If (tmDDFField(ilField).iFileID = tmLink(ilTest).iFromFileID) And (tmDDFField(ilField).iFieldID = tmLink(ilTest).iFromFieldID) Then
                                mGetField tmDDFField(ilField), 0
                                Exit For
                            End If
                        Next ilField
                        slTestValue = smFieldValue
                        For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                            If (tmDDFField(ilField).iFileID = tmLink(ilTest).iToFileID) And (tmDDFField(ilField).iFieldID = tmLink(ilTest).iToFieldID) Then
                                mGetField tmDDFField(ilField), 1
                                If StrComp(smFieldValue, slTestValue, 1) <> 0 Then
                                    ilMatch = False
                                End If
                                Exit For
                            End If
                        Next ilField
                    End If
                    If Not ilMatch Then
                        Exit For
                    End If
                Next ilTest
                If ilMatch Then
                    mGetField tlDDFField, 1
                Else
                    smFieldValue = ""
                End If
            Else
                smFieldValue = ""
            End If
            Exit For
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetFileID                      *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Translate File Name into ID     *
'*                                                     *
'*******************************************************
Private Function mGetFileID(slFileName As String) As Integer
    Dim ilFileName As Integer
    Dim slFile As String
    Dim ilPos As Integer
    For ilFileName = 0 To UBound(tmDDFFile) - 1 Step 1
        'Extract File Name C:\xxx\xxx.btr or c:xxx.btr
        slFile = Trim$(gRemoveZero(tmDDFFile(ilFileName).sLocation))
        ilPos = InStr(slFile, ":")
        If ilPos > 0 Then
            slFile = Mid$(slFile, ilPos + 1)
        End If
        ilPos = InStr(slFile, "\")
        Do While ilPos > 0
            slFile = Mid$(slFile, ilPos + 1)
            ilPos = InStr(slFile, "\")
        Loop
        If StrComp(slFile, slFileName, 1) = 0 Then
            mGetFileID = tmDDFFile(ilFileName).iFileID
            Exit Function
        End If
    Next ilFileName
    mGetFileID = -1
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mIndexBuild                     *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build Index table               *
'*                                                     *
'*******************************************************
Private Sub mIndexBuild(ilFileID As Integer, tlDDFIndex() As DDFINDEX)
    Dim ilRet As Integer
    Dim ilUpper As Integer
    'smDDFPath = lbcPath.Path & "\"
    Screen.MousePointer = vbHourglass
    ReDim tlDDFIndex(0 To 0) As DDFINDEX
    ilUpper = UBound(tlDDFIndex)
    hmFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hmFile, "", smDDFPath & "Index.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Screen.MousePointer = 0 'DEFAULT
        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
        Exit Sub
    End If
    lmActualRecLen = Len(tlDDFIndex(ilUpper)) 'Len(tmRec)
    imBaseRecLen = lmActualRecLen
    tmIndexSrchKey0.iFileID = ilFileID
    'ilRet = btrGetGreaterOrEqual(hmFile, tlDDFIndex(ilUpper), lmActualRecLen, tmIndexSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
    ilRet = btrGetGreaterOrEqual(hmFile, tmRec, lmActualRecLen, tmIndexSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = 0 'DEFAULT
        mBtrvErrorMsg ilRet, "btrGetGreaterOrEqual: " & "INDEX.DDF"
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Exit Sub
    End If
    Do While (ilRet = BTRV_ERR_NONE)
        If imTerminate Then
            Screen.MousePointer = 0 'DEFAULT
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            mTerminate
            Exit Sub
        End If
        mGetIndexFields
        tlDDFIndex(ilUpper) = tmGetDDFIndex
        If (tlDDFIndex(ilUpper).iFileID <> ilFileID) Then
            Exit Do
        End If
        ilUpper = ilUpper + 1
        ReDim Preserve tlDDFIndex(0 To ilUpper) As DDFINDEX
        lmActualRecLen = imBaseRecLen
        ilRet = btrGetNext(hmFile, tmRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hmFile)
    btrDestroy hmFile
    Screen.MousePointer = 0 'DEFAULT
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mIndexPop                       *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list box with Index    *
'*                     names                           *
'*                                                     *
'*******************************************************
Private Sub mIndexPop(tlDDFIndex() As DDFINDEX, tlDDFField() As DDFFIELD, lbcIndex As Control)
    Dim ilLoop As Integer
    Dim ilMaxIndex As Integer
    Dim ilMaxSeg As Integer
    Dim ilIndex As Integer
    Dim ilSeg As Integer
    Dim slName As String
    Dim ilField As Integer
    lbcIndex.Clear
    ilMaxIndex = 0
    For ilLoop = 0 To UBound(tlDDFIndex) - 1 Step 1
        If tlDDFIndex(ilLoop).iNumber > ilMaxIndex Then
            ilMaxIndex = tlDDFIndex(ilLoop).iNumber
        End If
    Next ilLoop
    For ilIndex = 0 To ilMaxIndex Step 1
        ilMaxSeg = 0
        For ilLoop = 0 To UBound(tlDDFIndex) - 1 Step 1
            If (tlDDFIndex(ilLoop).iNumber = ilIndex) And (tlDDFIndex(ilLoop).iPart > ilMaxSeg) Then
                ilMaxSeg = tlDDFIndex(ilLoop).iPart
            End If
        Next ilLoop
        slName = "Index" & Str$(ilIndex) & ":"
        For ilSeg = 0 To ilMaxSeg Step 1
            For ilLoop = 0 To UBound(tlDDFIndex) - 1 Step 1
                If (tlDDFIndex(ilLoop).iNumber = ilIndex) And (tlDDFIndex(ilLoop).iPart = ilSeg) Then
                    For ilField = 0 To UBound(tlDDFField) - 1 Step 1
                        If tlDDFField(ilField).iFieldID = tlDDFIndex(ilLoop).iFieldID Then
                            If Len(slName) = 0 Then
                                slName = Trim$(tlDDFField(ilField).sName)
                            Else
                                slName = slName & " " & Trim$(tlDDFField(ilField).sName)
                            End If
                        End If
                    Next ilField
                End If
            Next ilLoop
        Next ilSeg
        lbcIndex.AddItem slName
    Next ilIndex
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
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim slDrive As String
    Dim slFromFile As String
    Dim slCommand As String
    Dim ilSpace As Integer
    Dim slStr As String
    Dim slPath As String
    ReDim tmLink(0 To 0) As LINK
    ReDim tmKeyFile(0 To 0) As KEYFILE
    ReDim tmDDFLinkField(0 To 0) As DDFFIELD
    ReDim tmDDFField(0 To 0) As DDFFIELD
    fgListHTArial825 = 210
    igAlignCharWidth = 90
    igScrollBarWidth = 255
    imLBCDCtrls = 1
    If Command$ <> "Debug" Then
        'smCmmdFileName = Command$   'Selection File
        slCommand = Command$
        ilPos = InStr(1, slCommand, "/C:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(Left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    slCommand = Left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
            smCmmdFileName = slStr   'Selection File
        End If
        ilPos = InStr(1, slCommand, "/P:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(Left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    slCommand = Left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
            smDDFPath = slStr   'Selection File
            If right$(smDDFPath, 1) = "\" Then
                smDDFPath = Left$(smDDFPath, Len(smDDFPath) - 1)    'Remove \
            End If
        End If
        ilPos = InStr(1, slCommand, "/F:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(Left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    slCommand = Left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
            smDBPath = slStr   'Selection File
            If right$(smDBPath, 1) = "\" Then
                smDBPath = Left$(smDBPath, Len(smDBPath) - 1)   'Remove \
            End If
            sgDBPath = slStr
            If right$(sgDBPath, 1) <> "\" Then  'Add \ if missing
                sgDBPath = sgDBPath & "\"
            End If
        End If
        ilPos = InStr(1, slCommand, "/I:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(Left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    slCommand = Left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
            smImportPath = slStr   'Selection File
            If right$(smImportPath, 1) = "\" Then
                smImportPath = Left$(smImportPath, Len(smImportPath) - 1)   'Remove \
            End If
        End If
        ilPos = InStr(1, slCommand, "/E:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(Left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    slCommand = Left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
            smExportPath = slStr   'Selection File
        End If
        'ilTestSystem = False
        'ilPos = InStr(1, slCommand, "/S:Test", 1)
        'If ilPos > 0 Then
        '    ilTestSystem = True
        'End If
        'ilPos = InStr(1, slCommand, "/S:Prod", 1)
        'If ilPos > 0 Then
        '    ilTestSystem = False
        'End If
    Else
        ilPos = InStr(1, slCommand, "/F:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(Left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    slCommand = Left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
            smDBPath = slStr   'Selection File
            If right$(smDBPath, 1) = "\" Then
                smDBPath = Left$(smDBPath, Len(smDBPath) - 1)   'Remove \
            End If
            sgDBPath = slStr
            If right$(sgDBPath, 1) <> "\" Then  'Add \ if missing
                sgDBPath = sgDBPath & "\"
            End If
        End If
        smCmmdFileName = ""   'Selection File
        'ilTestSystem = False
    End If
    hmDB = CBtrvMngrInit(0) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    'gInitStdAlone ShowFix, "Counterpoint", ilTestSystem
    gdcSub.CellAlignment = flexAlignLeftTop

    imListFieldR0(1) = 15
    imListFieldR0(2) = 8 * igAlignCharWidth 'igAvgCharWidth * 5
    imListFieldR0(3) = 27 * igAlignCharWidth 'igAvgCharWidth * 12
    imListFieldR0(4) = 42 * igAlignCharWidth 'igAvgCharWidth * 12
    imListFieldR0(5) = 100 * igAlignCharWidth 'igAvgCharWidth * 12
    imListFieldR2(1) = 15
    imListFieldR2(2) = 10 * igAlignCharWidth 'igAvgCharWidth * 5
    imListFieldR2(3) = 33 * igAlignCharWidth 'igAvgCharWidth * 12
    imListFieldR2(4) = 51 * igAlignCharWidth 'igAvgCharWidth * 12
    imListFieldR2(5) = 80 * igAlignCharWidth 'igAvgCharWidth * 12
    imListFieldR2(6) = 200 * igAlignCharWidth 'igAvgCharWidth * 12
    imListFieldS(1) = 15
    imListFieldS(2) = 8 * igAlignCharWidth 'igAvgCharWidth * 5
    imListFieldS(3) = 27 * igAlignCharWidth 'igAvgCharWidth * 12
    imListFieldS(4) = 47 * igAlignCharWidth 'igAvgCharWidth * 12
    imListFieldS(5) = 100 * igAlignCharWidth 'igAvgCharWidth * 12
    
    
    mInitVar
    mInitBox
    imDragIndexDest = -1
    lbcEditReplace(2).AddItem " "
    pbcLbcEditReplace2_Paint
    imNoRowLbc = lbcEditReplace(2).Height \ fgListHTArial825
    imcTrash.Visible = False
    imcTrash.Picture = imcTrashClosed.Picture
    
    
    'plcGauge.Move ShowFix.Width / 2 - plcGauge.Width / 2
    'cmcFix.Move ShowFix.Width / 2 - cmcFix.Width / 2
    'cmcCancel.Move ShowFix.Width / 2 - cmcCancel.Width / 2
    plcTabSelection.Move 180, 60
    'plcTabSelection(0).Move 285, 300
    'plcTabSelection(1).Move plcTabSelection(0).left + plcTabSelection(0).Width + 30, plcTabSelection(0).top
    'plcTabSelection(2).Move plcTabSelection(1).left + plcTabSelection(1).Width + 30, plcTabSelection(0).top
    'plcTabSelection(3).Move plcTabSelection(2).Left + plcTabSelection(2).Width + 30, plcTabSelection(0).Top
    plcData.Move 180, 60
    plcLinks.Move 255, 480, plcLinks.Width, plcData.Height
    plcSave.Move 180, 60, plcSave.Width, plcData.Height
    plcFOpen.Move 255, 480, plcFOpen.Width, plcData.Height
    plcDDF.Move 255, 480
    plcSelection.Move plcDDF.Left, plcDDF.Top, plcDDF.Width
    plcEditReplace.Move plcDDF.Left, plcDDF.Top, plcDDF.Width
    'lncLine(0).X1 = plcTabSelection.left - plcDDF.left + 45
    'lncEdge(0).X1 = lncLine(0).X2 '+ 15
    'lncEdge(0).X2 = lncLine(0).X2 '+ 15
    'lncEdge(0).Y1 = 0
    'lncEdge(0).Y2 = 30
    'lncEdge(1).X1 = lncLine(0).X2 + 15 '+ 30
    'lncEdge(1).X2 = lncLine(0).X2 + 15 '+ 30
    'lncEdge(1).Y1 = 0
    'lncEdge(1).Y2 = 15
    'lncLine(1).X1 = plcTabSelection(1).left - plcDDF.left + 45
    'lncEdge(2).X1 = lncLine(1).X2 + 15
    'lncEdge(2).X2 = lncLine(1).X2 + 15
    'lncEdge(2).Y1 = 0
    'lncEdge(2).Y2 = 30
    'lncEdge(3).X1 = lncLine(1).X2 + 30
    'lncEdge(3).X2 = lncLine(1).X2 + 30
    'lncEdge(3).Y1 = 0
    'lncEdge(3).Y2 = 15
    'lncLine(2).X1 = plcTabSelection(2).left - plcDDF.left + 45
    'lncEdge(4).X1 = lncLine(2).X2 '+ 15
    'lncEdge(4).X2 = lncLine(2).X2 '+ 15
    'lncEdge(4).Y1 = 0
    'lncEdge(4).Y2 = 30
    'lncEdge(5).X1 = lncLine(2).X2 + 15 '+ 30
    'lncEdge(5).X2 = lncLine(2).X2 + 15 '+ 30
    'lncEdge(5).Y1 = 0
    'lncEdge(5).Y2 = 15
    ''lncLine(3).X1 = plcTabSelection(3).Left - plcDDF.Left + 45
    ''lncEdge(6).X1 = lncLine(3).X2 + 15
    ''lncEdge(6).X2 = lncLine(3).X2 + 15
    ''lncEdge(6).Y1 = 0
    ''lncEdge(6).Y2 = 30
    ''lncEdge(7).X1 = lncLine(3).X2 + 30
    ''lncEdge(7).X2 = lncLine(3).X2 + 30
    ''lncEdge(7).Y1 = 0
    ''lncEdge(7).Y2 = 15
    lbcSelection(0).AddItem "Comparisons"
    lbcSelection(0).AddItem "  X = Y Equal"
    lbcSelection(0).AddItem "  X <> Y Not Equal"
    lbcSelection(0).AddItem "  X < Y Less Than"
    lbcSelection(0).AddItem "  X > Y Greater Than"
    lbcSelection(0).AddItem "  X <= Y Less Than or Equal"
    lbcSelection(0).AddItem "  X >= Y Greater Than Or Equal"
    lbcSelection(0).AddItem "Boolean"
    lbcSelection(0).AddItem "  X and Y"
    lbcSelection(0).AddItem "  X or Y"
    lbcSelection(0).AddItem "Other"
    lbcSelection(0).AddItem "  (X) Parentheses"
    
    lbcEditReplace(1).AddItem "[DDF Spec]"
    lbcEditReplace(1).AddItem "Length of Time"
    lbcEditReplace(1).AddItem "Hex Field Dump"
    edcExample.Text = "1st Example:" & Chr(13) & Chr(10) & "file.btr:Status = " & """S""" & Chr(13) & Chr(10) & "And" & Chr(13) & Chr(10) & "file.btr:Date <= " & """6/1/95""" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "2nd Example: only B's" & Chr(13) & Chr(10) & "file.btr:Name >= " & """B""" & Chr(13) & Chr(10) & "And" & Chr(13) & Chr(10) & "file.btr:Name < " & """C"""
    mnuFilePrint.Enabled = False
    rbcEditReplace(0).Value = True
    lacEditReplace(1).Enabled = False
    edcReplaceValue.Enabled = False
    If Len(smDDFPath) = 0 Then
        If Len(smDBPath) = 0 Then
            smDBPath = CurDir$
            smDDFPath = smDBPath
            sgDBPath = smDBPath & "\"
        Else
            smDDFPath = smDBPath
        End If
    Else
        If Len(smDBPath) = 0 Then
            smDBPath = smDDFPath
            sgDBPath = smDBPath & "\"
        End If
    End If
    ilPos = InStr(smDDFPath, ":")
    If ilPos > 0 Then
        slDrive = Left(smDDFPath, ilPos)
        slPath = Mid$(smDDFPath, ilPos + 1)
        imIgnoreChg = True
        cbcDrive.Drive = slDrive
        imIgnoreChg = False
        lbcPath.Path = smDDFPath 'slPath
    End If
    ilPos = InStr(smDBPath, ":")
    If ilPos > 0 Then
        slDrive = Left(smDBPath, ilPos)
        slPath = Mid$(smDBPath, ilPos + 1)
        imIgnoreChg = True
        cbcDataDrive.Drive = slDrive
        imIgnoreChg = False
        lbcDataPath.Path = smDBPath 'slPath
    End If
    'smDDFPath = smDBPath
    ilPos = InStr(smImportPath, ":")
    If ilPos > 0 Then
        slDrive = Left(smImportPath, ilPos)
        slPath = Mid$(smImportPath, ilPos + 1)
        cbcFDrive.Drive = slDrive
        lbcFPath.Path = slPath
        lbcFFile.FileName = lbcFPath.Path & "\*.SFI"
    End If
    If Len(smImportPath) > 0 Then
        If right$(smImportPath, 1) <> "\" Then
            smImportPath = smImportPath & "\"
        End If
    End If
    If Len(smExportPath) > 0 Then
        If right$(smExportPath, 1) <> "\" Then
            smExportPath = smExportPath & "\"
        End If
    End If
    If Len(smCmmdFileName) > 0 Then
        slFromFile = smCmmdFileName
        If InStr(slFromFile, ":") = 0 Then
            slFromFile = smImportPath & slFromFile
        End If
        ilRet = 0
        On Error GoTo mInitErr:
        hmFrom = FreeFile
        Open slFromFile For Input As hmFrom
        If ilRet = 0 Then
            mReadSpec False
        Else
            ilRet = 0
            If InStr(slFromFile, ".") = 0 Then
                slFromFile = slFromFile & ".sfi"
                Open slFromFile For Input As hmFrom
                If ilRet = 0 Then
                    mReadSpec False
                Else
                    MsgBox "Unable to find command file " & smCmmdFileName, vbOkOnly, "File Missing"
                End If
            Else
                MsgBox "Unable to find command file " & smCmmdFileName, vbOkOnly, "File Missing"
            End If
        End If
        Close hmFrom
        On Error GoTo 0
        mSetCommands
        If cmcProcess.Enabled Then
            If rbcRun(0).Value Then
                tmcStart.Enabled = True
            End If
        End If
    End If
    ShowFix.Height = 6750
    mCenterStdAlone ShowFix
    ShowFix.Show
    'plcTabSelection_Click 0
    If plcTabSelection.SelectedItem.Index <> 1 Then
        ''plcTabSelection.SelectedItem.Index = 1
        'SendKeys "%a"
    Else
        plcTabSelection_Click
    End If
    imReadingFile = False
    imTerminate = False
    Exit Sub
mInitErr:
    ilRet = Err
    Resume Next
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
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fmBoxGridH
    Next ilLoop
    plcCalendar.Move edcField.Left, edcField.Top + edcField.Height
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gInitGlobalVar                  *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Initialize global variables     *
'*                                                     *
'*******************************************************
Private Sub mInitVar()
'
'   gInitGlobalVar
'
    'smDBPath = sgDBPath
    'smDDFPath = sgDDFPath
    'smExportPath = sgExportPath
    'smImportPath = sgImportPath
    imCalType = 1   'Calendar
    fgBoxGridH = 180      'Height of grid area (distance from bottom of form letter to bottom of form box)
    'fgListHTArial825 = 225  'Height of list item (arial 8.25 bold; DivideStyle=1-Black line; ListStyle=0-2D)
    'fgListHtSerif825 = 255  'Height of list item (san serif 8.25; Divide style=3=Raised; ListStyle=1-3D)
    fgBoxInsetX = 30    'X Margin from box outline to text
    lmTotalNoRec = 0
    lmProcessedNoRec = 0
    imEditRow = -1
    imEditCol = -1
    imRecCls = False
    ReDim tgUrf(0 To 0) As URF
    igBkgdProg = 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeRecSel                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain selection rules          *
'*                                                     *
'*******************************************************
Private Function mMakeRecSel() As Integer
    Dim slText As String
    Dim slChar As String
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilParNo As Integer
    Dim ilParType As Integer
    Dim ilField As Integer
    Dim ilFound As Integer
    Dim ilCurrIndex As Integer
    Dim ilCount As Integer
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slFileName As String
    Dim slFieldName As String
    Dim ilRet As Integer
    Dim ilFileID As Integer
    Dim ilFieldType1 As Integer
    Dim ilFieldType2 As Integer
    Dim ilFirstQuote As Integer '0=None; 1=Single; 2=Double
    ReDim tmRecSel(1 To 1) As SELCOMP
    smLogicSel = ""
    slText = Trim$(edcSelection.Text) & " "
    If Len(slText) = 1 Then
        mMakeRecSel = True
        Exit Function
    End If
    ilIndex = 1
    slName = ""
    ilParNo = 1
    Do While ilIndex <= Len(slText)
        ilCurrIndex = ilIndex
        slChar = Mid$(slText, ilIndex, 1)
        If ilParNo = 1 Then
            If slName = "" Then
                ilFirstQuote = 0
                If (slChar <> "(") And (slChar <> " ") And (slChar <> ")") And (slChar <> Chr(10)) And (slChar <> Chr(13)) Then
                    slName = slName & slChar
                    If (slChar = """") Then
                        ilParType = 2   'Value
                        ilFirstQuote = 2
                    ElseIf (slChar = "'") Then
                        ilParType = 2   'Value
                        ilFirstQuote = 1
                    Else
                        ilParType = 1   'Field
                    End If
                End If
                ilIndex = ilIndex + 1
            Else
                If ilParType = 1 Then   'Field
                    If (slChar = " ") Or (slChar = "=") Or (slChar = "<") Or (slChar = ">") Or (slChar = "(") Or (slChar = Chr(10)) Or (slChar = Chr(13)) Then
                        If (StrComp(slName, "And", 1) = 0) And ((slChar = " ") Or (slChar = "(") Or (slChar = Chr(10)) Or (slChar = Chr(13))) Then
                            smLogicSel = smLogicSel & "And"
                            slName = ""
                            ilParNo = 1
                        ElseIf (StrComp(slName, "Or", 1) = 0) And ((slChar = " ") Or (slChar = "(") Or (slChar = Chr(10)) Or (slChar = Chr(13))) Then
                            smLogicSel = smLogicSel & "Or"
                            slName = ""
                            ilParNo = 1
                        Else
                            'Test if field name exist
                            ilPos = InStr(slName, ":")
                            If ilPos = 0 Then
                                slFileName = ""
                                slFieldName = Trim$(slName)
                                ilFileID = -2
                            Else
                                ilRet = gParseItem(slName, 1, ":", slFileName)
                                ilRet = gParseItem(slName, 2, ":", slFieldName)
                                slFileName = Trim(slFileName)
                                slFieldName = Trim$(slFieldName)
                                ilFileID = mGetFileID(slFileName)
                            End If
                            ilFound = False
                            For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                                If (StrComp(slFieldName, Trim$(tmDDFField(ilField).sName), 1) = 0) And ((ilFileID = -2) Or (ilFileID = tmDDFField(ilField).iFileID)) Then
                                    ilFound = True
                                    tmRecSel(UBound(tmRecSel)).iFieldIndex1 = ilField
                                    tmRecSel(UBound(tmRecSel)).sValue1 = ""
                                    tmRecSel(UBound(tmRecSel)).iOper = -1
                                    tmRecSel(UBound(tmRecSel)).iFieldIndex2 = -1
                                    tmRecSel(UBound(tmRecSel)).sValue2 = ""
                                    slName = ""
                                    Exit For
                                End If
                            Next ilField
                            If Not ilFound Then
                                If (Left(slFieldName, 1) >= "0") And (Left$(slFieldName, 1) <= "9") Then
                                    ilFound = True
                                    tmRecSel(UBound(tmRecSel)).iFieldIndex1 = -1
                                    tmRecSel(UBound(tmRecSel)).sValue1 = slFieldName
                                    tmRecSel(UBound(tmRecSel)).iOper = -1
                                    tmRecSel(UBound(tmRecSel)).iFieldIndex2 = -1
                                    tmRecSel(UBound(tmRecSel)).sValue2 = ""
                                    slName = ""
                                End If
                            End If
                            If Not ilFound Then
                                MsgBox "Field Not Found- " & slName, vbOkOnly + vbExclamation + vbApplicationModal, "Selection Formula Error"
                                mMakeRecSel = False
                                Exit Function
                            End If
                            ilParNo = 2
                        End If
                    Else
                        slName = slName & slChar
                        ilIndex = ilIndex + 1
                    End If
                Else    'String
                    If ((Left$(slName, 1) = """") And (ilFirstQuote = 2)) Or ((Left$(slName, 1) = "'") And (ilFirstQuote = 1)) Then
                        slName = ""
                    End If
                    If ((slChar = """") And (ilFirstQuote = 2)) Or ((slChar = "'") And (ilFirstQuote = 1)) Then
                        tmRecSel(UBound(tmRecSel)).iFieldIndex1 = -1
                        tmRecSel(UBound(tmRecSel)).sValue1 = slName
                        tmRecSel(UBound(tmRecSel)).iOper = -1
                        tmRecSel(UBound(tmRecSel)).iFieldIndex2 = -1
                        tmRecSel(UBound(tmRecSel)).sValue2 = ""
                        slName = ""
                        ilParNo = 2
                    Else
                        slName = slName & slChar
                    End If
                    ilIndex = ilIndex + 1
                End If
            End If
        ElseIf ilParNo = 2 Then
            If slName = "" Then
                If (slChar <> " ") Then
                    slName = slName & slChar
                End If
                ilIndex = ilIndex + 1
            Else
                If (slChar = "=") Or (slChar = "<") Or (slChar = ">") Then
                    slName = slName & slChar
                    ilIndex = ilIndex + 1
                Else
                    If slName = "=" Then
                        tmRecSel(UBound(tmRecSel)).iOper = 0
                    ElseIf slName = "<>" Then
                        tmRecSel(UBound(tmRecSel)).iOper = 1
                    ElseIf slName = "<" Then
                        tmRecSel(UBound(tmRecSel)).iOper = 2
                    ElseIf slName = ">" Then
                        tmRecSel(UBound(tmRecSel)).iOper = 3
                    ElseIf slName = "<=" Then
                        tmRecSel(UBound(tmRecSel)).iOper = 4
                    ElseIf slName = ">=" Then
                        tmRecSel(UBound(tmRecSel)).iOper = 5
                    Else
                        MsgBox "Operator Not Found: " & slName, vbOkOnly + vbExclamation + vbApplicationModal, "Selection Formula Error"
                        mMakeRecSel = False
                        Exit Function
                    End If
                    slName = ""
                    ilParNo = 3
                End If
            End If
        ElseIf ilParNo = 3 Then
            If slName = "" Then
                ilFirstQuote = 0
                If (slChar <> " ") Then
                    slName = slName & slChar
                    If (slChar = """") Then
                        ilParType = 2   'Value
                        ilFirstQuote = 2
                    ElseIf (slChar = "'") Then
                        ilParType = 2   'Value
                        ilFirstQuote = 1
                    Else
                        ilParType = 1   'Field
                    End If
                End If
                ilIndex = ilIndex + 1
            Else
                If ilParType = 1 Then   'Field
                    If (slChar = " ") Or (slChar = "(") Or (slChar = ")") Or (slChar = Chr(10)) Or (slChar = Chr(13)) Then
                        'Test if field name exist
                        ilPos = InStr(slName, ":")
                        If ilPos = 0 Then
                            slFileName = ""
                            slFieldName = Trim$(slName)
                            ilFileID = -2
                        Else
                            ilRet = gParseItem(slName, 1, ":", slFileName)
                            ilRet = gParseItem(slName, 2, ":", slFieldName)
                            slFileName = Trim(slFileName)
                            slFieldName = Trim$(slFieldName)
                            ilFileID = mGetFileID(slFileName)
                        End If
                        ilFound = False
                        For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                            If (StrComp(slFieldName, Trim$(tmDDFField(ilField).sName), 1) = 0) And ((ilFileID = -2) Or (ilFileID = tmDDFField(ilField).iFileID)) Then
                                ilFound = True
                                tmRecSel(UBound(tmRecSel)).iFieldIndex2 = ilField
                                tmRecSel(UBound(tmRecSel)).sValue2 = ""
                                slName = ""
                                Exit For
                            End If
                        Next ilField
                        If Not ilFound Then
                            If (Left(slFieldName, 1) >= "0") And (Left$(slFieldName, 1) <= "9") Then
                                ilFound = True
                                tmRecSel(UBound(tmRecSel)).iFieldIndex2 = -1
                                tmRecSel(UBound(tmRecSel)).sValue2 = slFieldName
                                slName = ""
                            End If
                        End If
                        If Not ilFound Then
                            MsgBox "Field Not Found: " & slName, vbOkOnly + vbExclamation + vbApplicationModal, "Selection Formula Error"
                            mMakeRecSel = False
                            Exit Function
                        End If
                        ilParNo = 1
                    Else
                        slName = slName & slChar
                        ilIndex = ilIndex + 1
                    End If
                Else    'String
                    If ((Left$(slName, 1) = """") And (ilFirstQuote = 2)) Or ((Left$(slName, 1) = "'") And (ilFirstQuote = 1)) Then
                        slName = ""
                    End If
                    If ((slChar = """") And (ilFirstQuote = 2)) Or ((slChar = "'") And (ilFirstQuote = 1)) Then
                        tmRecSel(UBound(tmRecSel)).iFieldIndex2 = -1
                        tmRecSel(UBound(tmRecSel)).sValue2 = slName
                        slName = ""
                        ilParNo = 1
                    Else
                        slName = slName & slChar
                    End If
                    ilIndex = ilIndex + 1
                End If
                If ilParNo = 1 Then
                    If (tmRecSel(UBound(tmRecSel)).iFieldIndex1 = -1) And (tmRecSel(UBound(tmRecSel)).iFieldIndex2 = -1) Then
                        MsgBox "Two Value Defined, One Must Be A Field", vbOkOnly + vbExclamation + vbApplicationModal, "Selection Formula Error"
                        mMakeRecSel = False
                        Exit Function
                    End If
                        If (tmRecSel(UBound(tmRecSel)).iFieldIndex1 >= 0) And (tmRecSel(UBound(tmRecSel)).iFieldIndex2 >= 0) Then
                            ilFieldType1 = Asc(tmDDFField(tmRecSel(UBound(tmRecSel)).iFieldIndex1).sDataType)
                            ilFieldType2 = Asc(tmDDFField(tmRecSel(UBound(tmRecSel)).iFieldIndex2).sDataType)
                            If ilFieldType1 = 15 Then
                                ilFieldType1 = 1
                            End If
                            If ilFieldType2 = 15 Then
                                ilFieldType2 = 1
                            End If
                            If ilFieldType1 <> ilFieldType2 Then
                                MsgBox "Field Data Types Don't Match", vbOkOnly + vbExclamation + vbApplicationModal, "Selection Formula Error"
                                mMakeRecSel = False
                                Exit Function
                            End If
                        End If
                    smLogicSel = smLogicSel & "#" & Trim$(Str$(UBound(tmRecSel)))
                    ReDim Preserve tmRecSel(1 To UBound(tmRecSel) + 1) As SELCOMP
                End If
            End If
        End If
        If ((slChar = "(") Or (slChar = ")")) And (ilIndex <> ilCurrIndex) Then
            smLogicSel = smLogicSel & slChar
        End If
    Loop
    If ilParNo <> 1 Then
        MsgBox "Not All Fields And/Or Values Defined", vbOkOnly + vbExclamation + vbApplicationModal, "Selection Formula Error"
        mMakeRecSel = False
        Exit Function
    End If
    'Count Parentheses
    ilCount = 0
    For ilLoop = 1 To Len(smLogicSel) Step 1
        slChar = Mid(smLogicSel, ilLoop, 1)
        If slChar = "(" Then
            ilCount = ilCount + 1
        End If
        If slChar = ")" Then
            If ilLoop = 1 Then
                MsgBox "Parentheses Not Matching", vbOkOnly + vbExclamation + vbApplicationModal, "Selection Formula Error"
                mMakeRecSel = False
                Exit Function
            Else
                If Mid(smLogicSel, ilLoop - 1, 1) = "(" Then
                    MsgBox "Parentheses Not Matching", vbOkOnly + vbExclamation + vbApplicationModal, "Selection Formula Error"
                    mMakeRecSel = False
                    Exit Function
                End If
            End If
            ilCount = ilCount - 1
        End If
    Next ilLoop
    If ilCount <> 0 Then
        MsgBox "Parentheses Missmatch", vbOkOnly + vbExclamation + vbApplicationModal, "Selection Formula Error"
        Exit Function
    End If
    mMakeRecSel = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMassUpdateRecords              *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Mass pdate records              *
'*                                                     *
'*******************************************************
Private Sub mMassUpdateRecords()
    Dim ilRet As Integer
    Dim llRecNo As Long
    Dim llStartRecNo As Long
    Dim llEndRecNo As Long
    Dim ilRecIndex1 As Integer
    Dim ilRecIndex2 As Integer
    Dim llRec As Long
    Dim ilLoop As Integer
    Dim ilField As Integer
    Dim ilDataType As Integer
    Dim slSvDataType As String
    Dim slFieldDataType As String
    Dim slFieldInfo As String
    Dim slFieldName As String
    Dim ilSize As Integer
    Dim ilCol As Integer
    Dim ilPos As Integer
    Dim slValue As String
    Dim llValue As Long
    Dim slInt2 As INT2
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim llPercent As Long
    'Dim slSvFieldValue As String
    Dim ilOper As Integer
    Dim slAdjustValue As String
    llStartRecNo = 1
    llEndRecNo = lmMaxNoRec
    llRec = llStartRecNo
    ilRecIndex1 = 1
    ilRecIndex2 = 1
    Do While llRec > 32000
        ilRecIndex1 = ilRecIndex1 + 1
        llRec = llRec - 32000
    Loop
    ilRecIndex2 = llRec
    ilRet = 0
    On Error GoTo mMassUpdateRecordsErr:
    DoEvents
    If imTerminate Then
        Exit Sub
    End If
    imVarBaseRecLen = 0
    For ilField = 0 To UBound(tmDDFField) - 1 Step 1
        If (Asc(tmDDFField(ilField).sDataType) = 12) Then
            imVarLenRec = True
            imVarType = 12
        ElseIf (Asc(tmDDFField(ilField).sDataType) = 13) Then
            imVarLenRec = True
            imVarType = 13
        Else
            imVarBaseRecLen = imVarBaseRecLen + tmDDFField(ilField).iSize
        End If
    Next ilField
    For llRecNo = llStartRecNo To llEndRecNo Step 1
        If ((ilRecIndex1 = imRecMaxUpper1) And (ilRecIndex2 < imRecMaxUpper2)) Or ((ilRecIndex1 < imRecMaxUpper1) And (ilRecIndex2 <= 32000)) Then
            Do
                ilRet = btrBeginTrans(hmFile, 1000)
                lmActualRecLen = imBaseRecLen
                ilRet = btrGetDirect(hmFile, tmRec, lmActualRecLen, tmRecInfo(ilRecIndex1, ilRecIndex2).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                DoEvents
                If imTerminate Then
                    Exit Sub
                End If
                ilRet = btrDelete(hmFile)
                If ilRet <> BTRV_ERR_NONE Then
                    MsgBox "Delete Failed, Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal
                    imTerminate = True
                    ilRet = btrAbortTrans(hmFile)
                    Exit Sub
                End If
                'Get fields
                For ilLoop = 0 To lbcEditReplace(2).ListCount - 2 Step 1
                    slFieldInfo = lbcEditReplace(2).List(ilLoop)
                    ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
                    ilRet = gParseItem(slFieldInfo, 4, "|", slFieldDataType)    'New data type
                    If ilRet <> CP_MSG_NONE Then
                        slFieldDataType = ""
                    Else
                        If StrComp(slFieldDataType, "Length of Time", 1) = 0 Then
                            slFieldDataType = Chr(100)
                        End If
                        If StrComp(slFieldDataType, "Hex Field Dump", 1) = 0 Then
                            slFieldDataType = Chr(101)
                        End If
                    End If
                    ilRet = gParseItem(slFieldInfo, 5, "|", smFieldValue)
                    If (Left$(smFieldValue, 1) = """") Or (Left$(smFieldValue, 1) = "'") Then
                        smFieldValue = Mid$(smFieldValue, 2)
                    End If
                    If (right$(smFieldValue, 1) = """") Or (right$(smFieldValue, 1) = "'") Then
                        smFieldValue = Left$(smFieldValue, Len(smFieldValue) - 1)
                    End If
                    'slSvFieldValue = smFieldValue
                    For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                        If StrComp(Trim$(slFieldName), Trim$(tmDDFField(ilField).sName), 1) = 0 Then
                            ilDataType = Asc(tmDDFField(ilField).sDataType)
                        End If
                    Next ilField
                    ilOper = -1
                    If mTestOper(ilDataType) Then
                        If InStr(1, smFieldValue, "+", 1) > 0 Then
                            ilPos = InStr(1, smFieldValue, "+", 1)
                            slAdjustValue = Mid$(smFieldValue, ilPos + 1)
                            smFieldValue = Left$(smFieldValue, ilPos - 1)
                            ilOper = 1
                        ElseIf InStr(1, smFieldValue, "-", 1) > 1 Then
                            ilPos = InStr(1, smFieldValue, "-", 1)
                            slAdjustValue = Mid$(smFieldValue, ilPos + 1)
                            smFieldValue = Left$(smFieldValue, ilPos - 1)
                            ilOper = 2
                        ElseIf InStr(1, smFieldValue, "*", 1) > 0 Then
                            ilPos = InStr(1, smFieldValue, "*", 1)
                            slAdjustValue = Mid$(smFieldValue, ilPos + 1)
                            smFieldValue = Left$(smFieldValue, ilPos - 1)
                           ilOper = 3
                        ElseIf InStr(1, smFieldValue, "/", 1) > 0 Then
                            ilPos = InStr(1, smFieldValue, "/", 1)
                            slAdjustValue = Mid$(smFieldValue, ilPos + 1)
                            smFieldValue = Left$(smFieldValue, ilPos - 1)
                            ilOper = 4
                        ElseIf InStr(1, smFieldValue, "\", 1) > 0 Then
                            ilPos = InStr(1, smFieldValue, "\", 1)
                            slAdjustValue = Mid$(smFieldValue, ilPos + 1)
                            smFieldValue = Left$(smFieldValue, ilPos - 1)
                            ilOper = 5
                        End If
                    End If
                    ''smFieldValue = slSvFieldValue
                    For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                        If StrComp(Trim$(smFieldValue), Trim$(tmDDFField(ilField).sName), 1) = 0 Then
                            mGetField tmDDFField(ilField), 0    'This will store the value into smFieldValue
                            If ilOper = 1 Then
                                smFieldValue = gAddStr(smFieldValue, slAdjustValue)
                            ElseIf ilOper = 2 Then
                                smFieldValue = gSubStr(smFieldValue, slAdjustValue)
                            ElseIf ilOper = 3 Then
                                smFieldValue = gMulStr(smFieldValue, slAdjustValue)
                            ElseIf ilOper = 4 Then
                                smFieldValue = gDivStr(smFieldValue, slAdjustValue)
                            ElseIf ilOper = 5 Then
                                smFieldValue = gTDivStr(smFieldValue, slAdjustValue)
                            End If
                            Exit For
                        End If
                    Next ilField
                    For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                        If StrComp(Trim$(slFieldName), Trim$(tmDDFField(ilField).sName), 1) = 0 Then
                            slSvDataType = tmDDFField(ilField).sDataType
                            If slFieldDataType <> "" Then
                                tmDDFField(ilField).sDataType = slFieldDataType
                            End If
                            mPutField tmDDFField(ilField), 0   'Store smFieldValue into record
                            tmDDFField(ilField).sDataType = slSvDataType
                            Exit For
                        End If
                    Next ilField
                Next ilLoop
                'Update record
                If imVarLenRec Then
                    If imVarType = 12 Then
                        For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                            If (Asc(tmDDFField(ilField).sDataType) = 12) Then
                                If (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary") Then
                                    tmSsf = tmRec
                                    lmActualRecLen = imVarBaseRecLen + Len(tmProg) * tmSsf.iCount
                                Else
                                    lmActualRecLen = imVarBaseRecLen + tmDDFField(ilField).iSize
                                End If
                                'ilRet = btrUpdate(hmFile, tmRec, lmActualRecLen)
                                ilRet = btrInsert(hmFile, tmRec, lmActualRecLen, 0)
                                Exit For
                            End If
                        Next ilField
                    Else
                        ilRet = BTRV_ERR_NONE
                        For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                            If (Asc(tmDDFField(ilField).sDataType) = 13) Then
                                ilIndex = tmDDFField(ilField).iOffset + 1
                                'slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                                'HMemCpy ilValue, slInt2, 2
                                ilValue = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
                                'lmActualRecLen = imVarBaseRecLen + Len(smFieldValue) + 4
                                lmActualRecLen = imVarBaseRecLen + ilValue + 4
                                'ilRet = btrUpdate(hmFile, tmRec, lmActualRecLen)
                                ilRet = btrInsert(hmFile, tmRec, lmActualRecLen, 0)
                                Exit For
                            End If
                        Next ilField
                    End If
                Else
                    lmActualRecLen = imBaseRecLen
                    'ilRet = btrUpdate(hmFile, tmRec, lmActualRecLen)
                    ilRet = btrInsert(hmFile, tmRec, lmActualRecLen, 0)
                End If
                DoEvents
                If imTerminate Then
                    ilRet = btrAbortTrans(hmFile)
                    Exit Sub
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                MsgBox "Insert Failed, Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal
                ilRet = btrAbortTrans(hmFile)
                imTerminate = True
                Exit Sub
            End If
            ilRet = btrEndTrans(hmFile)
            ilRecIndex2 = ilRecIndex2 + 1
            If ilRecIndex2 > 32000 Then
                ilRecIndex1 = ilRecIndex1 + 1
                ilRecIndex2 = 1
            End If
            lacChg.Caption = "Updated" & Str$(llRecNo) & " of" & Str$(lmMaxNoRec)
            llPercent = 100 * llRecNo / lmMaxNoRec
            If llPercent > 100 Then
                llPercent = 100
            End If
            If lmPercent <> llPercent Then
                plcChg.Caption = Trim$(Str$(llPercent)) & "%"
                lmPercent = llPercent
            End If
        End If
    Next llRecNo
    lacChg.Caption = "Updated" & Str$(lmMaxNoRec) & " of" & Str$(lmMaxNoRec)
    plcChg.Caption = "100%"
    On Error GoTo 0
    Exit Sub
mMassUpdateRecordsErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub mnuFileExit_Click()
    mTerminate
End Sub


Private Sub mnuFileOpen_Click()
    cmcDone.Enabled = False
    cmcProcess.Enabled = False
    lacFOpen.Caption = "Open File"
    plcFOpen.Visible = True
    plcFOpen.ZOrder vbBringToFront
End Sub
Private Sub mnuFilePrint_Click()
    Dim ilColumns As Integer
    cmcDone.Enabled = False
    cmcProcess.Enabled = False
    imEnableOnCancel = False
    lacSave.Caption = "Print Grid Data"
    plcRun.Visible = False
    lacSaveTo.Visible = False
    plcSaveTo.Visible = False
    cmcSBrowse.Visible = False
    lacDescription.Top = lacSaveTo.Top
    edcDescription.Top = plcSaveTo.Top
    cmcSSave.Caption = "&Print"
    lacRecordNo(0).Visible = True
    lacRecordNo(1).Visible = True
    edcRecordNo(0).Visible = True
    edcRecordNo(1).Visible = True
    cbcInclude(0).Value = vbChecked
    cbcInclude(1).Value = vbChecked
    edcRecordNo(0).Text = "1"
    edcRecordNo(1).Text = Trim$(Str$(lmMaxNoRec))
    lacWrap.Visible = True
    edcWrap.Visible = True
    edcWrap.Text = ""
    If (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary") Then
        plcPAS.Visible = True
        rbcPAS(0).Value = True
    Else
        plcPAS.Visible = False
    End If
    Printer.FontName = "Courier New"
    Printer.FontSize = 8
    Printer.FontBold = True
    '1440 twip per inch- 1/4 margin on each side
    ilColumns = (Printer.Width - (1440 \ 2)) \ Printer.TextWidth("M") - 10
    edcWrap.Text = Trim$(Str$(ilColumns))
    'If ilColumns > 132 Then
    '    edcWrap.Text = 132
    'Else
    '    edcWrap.Text = 80
    'End If
    plcInclude.Visible = True
    plcSave.Visible = True
    plcSave.ZOrder vbBringToFront
End Sub
Private Sub mnuFilePSetup_Click()
    cdcSetup.Flags = cdlPDPrintSetup
    cdcSetup.Action = 5    'DLG_PRINT
End Sub
Private Sub mnuFileSave_Click()
    If Len(smCmmdFileName) <= 0 Then
        mnuFileSaveAs_Click
    Else
        lacSave.Caption = "Save Specifications"
        edcSaveTo.Text = smCmmdFileName
        cmcSSave_Click
    End If
End Sub
Private Sub mnuFileSaveAs_Click()
    imEnableOnCancel = cmcProcess.Enabled   'If menu selected while in Grid, then cmcProcess.Enabled is false
                                            'otherwise it is true
    cmcDone.Enabled = False
    cmcProcess.Enabled = False
    lacSave.Caption = "Save Specifications"
    lacRecordNo(0).Visible = False
    lacRecordNo(1).Visible = False
    edcRecordNo(0).Visible = False
    edcRecordNo(1).Visible = False
    lacWrap.Visible = False
    edcWrap.Visible = False
    plcInclude.Visible = False
    plcPAS.Visible = False
    plcRun.Visible = True
    plcSave.Visible = True
    plcSave.ZOrder vbBringToFront
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainRecords                  *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain records                  *
'*                                                     *
'*******************************************************
Private Sub mObtainRecords(ilForMassChg As Integer)
    Dim ilRet As Integer
    Dim slFile As String
    Dim slName As String
    Dim slFName As String
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slFieldName As String
    Dim ilField As Integer
    Dim llNoRec As Long         'Number of records in Sof
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilNo32000 As Integer
    Dim llNoFields As Long
    Dim llRecNo As Long
    Dim llNoProcessed As Long
    Dim llNoShown As Long
    Dim ilSize As Integer
    Dim slFileName As String
    Dim ilFileID As Integer
    Dim slFieldInfo As String
    Dim slSvDataType As String
    Dim slFieldDataType As String
    Dim ilFound As Integer
    Dim ilTest As Integer
    Dim llPercent As Long
    Dim ilKeyIndex As Integer
    Dim slSearchFrom As String
    Dim slSearchTo As String
    Dim ilSearchDDFOffset As Integer
    
    ilKeyIndex = Val(edcSearchKey.Text)
    slSearchFrom = Trim$(edcSearchFrom.Text)
    slSearchTo = Trim$(edcSearchTo.Text)
    imRecUpper1 = 1
    imRecUpper2 = 1
    imRecMaxUpper1 = 1
    imRecMaxUpper2 = 1
    'imRowUpper1 = 1
    'imRowUpper2 = 1
    imGrid1Upper1 = 1
    imGrid1Upper2 = 1
    imGrid2Upper1 = 1
    imGrid2Upper2 = 1
    imGrid3Upper1 = 1
    imGrid3Upper2 = 1
    imGrid4Upper1 = 1
    imGrid4Upper2 = 1
    imGrid5Upper1 = 1
    imGrid5Upper2 = 1
    imGrid6Upper1 = 1
    imGrid6Upper2 = 1
    ReDim tmGrid1Data(1 To 1, 1 To 1) As GRID1DATA
    ReDim tmGrid2Data(1 To 1, 1 To 1) As GRID2DATA
    ReDim tmGrid3Data(1 To 1, 1 To 1) As GRID3DATA
    ReDim tmGrid4Data(1 To 1, 1 To 1) As GRID4DATA
    ReDim tmGrid5Data(1 To 1, 1 To 1) As GRID5DATA
    ReDim tmGrid6Data(1 To 1, 1 To 1) As GRID6DATA
    If Not ilForMassChg Then
        lmMaxNoRec = 1
        vbcData.Min = 1
        vbcData.Max = 1
        vbcData_Change
    End If
    If lbcFile.ListIndex < 0 Then
        Exit Sub
    End If
    llRecNo = 0
    smDBPath = lbcDataPath.Path & "\"
    sgDBPath = smDBPath
    Screen.MousePointer = vbHourglass
    ilRet = btrClose(hmFile)
    btrDestroy hmFile
    hmFile = CBtrvTable()    'CBtrvObj()
    slFile = lbcFile.List(lbcFile.ListIndex)
    slName = lbcFile.List(lbcFile.ListIndex)
    For ilLoop = 0 To UBound(tmDDFFile) - 1 Step 1
        If StrComp(Trim$(tmDDFFile(ilLoop).sName), slName, 1) = 0 Then
            'Extract File Name C:\xxx\xxx.btr or c:xxx.btr
            slFile = Trim$(gRemoveZero(tmDDFFile(ilLoop).sLocation))
            ilPos = InStr(slFile, ":")
            If ilPos > 0 Then
                slFile = Mid$(slFile, ilPos + 1)
            End If
            ilPos = InStr(slFile, "\")
            Do While ilPos > 0
                slFile = Mid$(slFile, ilPos + 1)
                ilPos = InStr(slFile, "\")
            Loop
            imFileID = tmDDFFile(ilLoop).iFileID
            Exit For
        End If
    Next ilLoop
    ilPos = InStr(slFile, ".")
    If ilPos > 0 Then
        slFName = Left$(slFile, ilPos - 1)
    Else
        slFName = slFile
    End If
    ilRet = btrOpen(hmFile, "", smDBPath & slFile, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Screen.MousePointer = 0 'DEFAULT
        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
        Exit Sub
    End If
    If (slSearchFrom <> "") Or (slSearchTo <> "") Then
        ilRet = mSearchBuildIndex(imFileID, ilKeyIndex, tmSearchDDFIndex)
        If Not ilRet Then
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            Screen.MousePointer = 0 'DEFAULT
            Exit Sub
        End If
        ilRet = mSearchBuildField(imFileID, tmSearchDDFIndex.iFieldID, tmSearchDDFField)
        If Not ilRet Then
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            Screen.MousePointer = 0 'DEFAULT
            Exit Sub
        End If
        ilSearchDDFOffset = tmSearchDDFField.iOffset
    End If
    
    'Open key files
    For ilLoop = LBound(tmKeyFile) To UBound(tmKeyFile) - 1 Step 1
        ilRet = btrClose(tmKeyFile(ilLoop).hFile)
        btrDestroy tmKeyFile(ilLoop).hFile
    Next ilLoop
    ReDim tmKeyFile(0 To 0) As KEYFILE
    For ilLoop = LBound(tmLink) To UBound(tmLink) - 1 Step 1
        ilFound = False
        For ilTest = LBound(tmKeyFile) To UBound(tmKeyFile) - 1 Step 1
            ilFound = False
            If tmLink(ilLoop).iToFileID = tmKeyFile(ilTest).iFileID Then
                ilFound = True
                Exit For
            End If
        Next ilTest
        If Not ilFound Then
            For ilTest = 0 To UBound(tmDDFFile) - 1 Step 1
                If tmLink(ilLoop).iToFileID = tmDDFFile(ilTest).iFileID Then
                    'Extract File Name C:\xxx\xxx.btr or c:xxx.btr
                    slFile = Trim$(gRemoveZero(tmDDFFile(ilTest).sLocation))
                    ilPos = InStr(slFile, ":")
                    If ilPos > 0 Then
                        slFile = Mid$(slFile, ilPos + 1)
                    End If
                    ilPos = InStr(slFile, "\")
                    Do While ilPos > 0
                        slFile = Mid$(slFile, ilPos + 1)
                        ilPos = InStr(slFile, "\")
                    Loop
                    tmKeyFile(UBound(tmKeyFile)).iFileID = tmDDFFile(ilTest).iFileID
                    tmKeyFile(UBound(tmKeyFile)).hFile = CBtrvTable()    'CBtrvObj()
                    ilRet = btrOpen(tmKeyFile(UBound(tmKeyFile)).hFile, "", smDBPath & slFile, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilRet = btrClose(tmKeyFile(UBound(tmKeyFile)).hFile)
                        btrDestroy tmKeyFile(UBound(tmKeyFile)).hFile
                        Screen.MousePointer = 0 'DEFAULT
                        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
                        Exit Sub
                    End If
                    ReDim Preserve tmKeyFile(0 To UBound(tmKeyFile) + 1) As KEYFILE
                    Exit For
                End If
            Next ilTest
        End If
    Next ilLoop
    If Not ilForMassChg Then
        'Determine if Fix (use Ext operations) or variable rcords (use GetFirst, GetNext)
        For ilLoop = LBound(imFieldSizeCount) To UBound(imFieldSizeCount) Step 1
            imFieldSizeCount(ilLoop) = 0
        Next ilLoop
    End If
    imVarLenRec = False
    For ilLoop = 0 To lbcEditReplace(2).ListCount - 2 Step 1
        slFieldInfo = lbcEditReplace(2).List(ilLoop)
        ilRet = gParseItem(slFieldInfo, 1, "|", slFileName)
        ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
        ilRet = gParseItem(slFieldInfo, 4, "|", slFieldDataType)
        If ilRet <> CP_MSG_NONE Then
            slFieldDataType = ""
        Else
            'If Not ilForMassChg Then
            '    If (StrComp(Trim$(slFileName), "Mnf.Btr", 1) = 0) And (StrComp(Trim$(slFieldName), "mnfRPU", 1) = 0) Then
            '        slFieldDataType = "Hex Field Dump"
            '    End If
            '    If (StrComp(Trim$(slFileName), "Mnf.Btr", 1) = 0) And (StrComp(Trim$(slFieldName), "mnfSSComm", 1) = 0) Then
            '        slFieldDataType = "Hex Field Dump"
            '    End If
            'End If
            If StrComp(slFieldDataType, "Length of Time", 1) = 0 Then
                slFieldDataType = Chr(100)
            End If
            If StrComp(slFieldDataType, "Hex Field Dump", 1) = 0 Then
                slFieldDataType = Chr(101)
            End If
        End If
        ilFileID = mGetFileID(slFileName)
        For ilField = 0 To UBound(tmDDFField) - 1 Step 1
            If (StrComp(Trim$(slFieldName), Trim$(tmDDFField(ilField).sName), 1) = 0) And (ilFileID = tmDDFField(ilField).iFileID) Then
                If Not ilForMassChg Then
                    slSvDataType = tmDDFField(ilField).sDataType
                    If slFieldDataType <> "" Then
                        tmDDFField(ilField).sDataType = slFieldDataType
                    End If
                    ilSize = mFieldSize(tmDDFField(ilField))
                    tmDDFField(ilField).sDataType = slSvDataType
                    If ilSize <= GRID1SIZE Then
                        imFieldSizeCount(1) = imFieldSizeCount(1) + 1
                    ElseIf ilSize <= GRID2SIZE Then
                        imFieldSizeCount(2) = imFieldSizeCount(2) + 1
                    ElseIf ilSize <= GRID3SIZE Then
                        imFieldSizeCount(3) = imFieldSizeCount(3) + 1
                    ElseIf ilSize <= GRID4SIZE Then
                        imFieldSizeCount(4) = imFieldSizeCount(4) + 1
                    ElseIf ilSize <= GRID5SIZE Then
                        imFieldSizeCount(5) = imFieldSizeCount(5) + 1
                    ElseIf ilSize <= GRID6SIZE Then
                        imFieldSizeCount(6) = imFieldSizeCount(6) + 1
                    End If
                End If
                Exit For
            End If
        Next ilField
    Next ilLoop
    imVarBaseRecLen = 0
    For ilField = 0 To UBound(tmDDFField) - 1 Step 1
        If (Asc(tmDDFField(ilField).sDataType) = 12) Then
            imVarLenRec = True
            imVarType = 12
        ElseIf (Asc(tmDDFField(ilField).sDataType) = 13) Then
            imVarLenRec = True
            imVarType = 13
        Else
            imVarBaseRecLen = imVarBaseRecLen + tmDDFField(ilField).iSize
        End If
    Next ilField
    lmMaxNoRec = btrRecords(hmFile)
    llNoProcessed = 0
    llNoShown = 0
    If (slSearchFrom = "") And (slSearchTo = "") Then
        lacTotal.Caption = "Read 0, Selected 0:" & Str(lmMaxNoRec) & " records"
        plcGauge.Visible = True
    Else
        lacTotal.Caption = "Read 0, Selected 0" & " records"
        plcGauge.Visible = False
    End If
    plcGauge.Value = 0
    lmPercent = 0
    If lmMaxNoRec = 0 Then
        plcGauge.Value = 100
        ReDim tmRecInfo(1 To 1, 1 To 1) As RECINFO
        imRecMaxUpper1 = 1
        imRecMaxUpper2 = 1
        vbcData.Min = 1
        vbcData.Max = 1
        Screen.MousePointer = 0 'DEFAULT
        Exit Sub
    End If
    If lmMaxNoRec Mod 32000 <> 0 Then
        ilNo32000 = lmMaxNoRec \ 32000 + 1
    Else
        ilNo32000 = lmMaxNoRec \ 32000
    End If
    If lmMaxNoRec <= 32000 Then
        ReDim tmRecInfo(1 To 1, 1 To lmMaxNoRec) As RECINFO
        If Not ilForMassChg Then
            ReDim tmRowLinkList(1 To 1, 1 To lmMaxNoRec) As ROWLINKLIST
        End If
    Else
        On Error GoTo mObtainRecordsErr:
        Do
            ilRet = 0
            ReDim tmRecInfo(1 To ilNo32000, 1 To 32000) As RECINFO
            If (Not ilForMassChg) And (ilRet = 0) Then
                ReDim tmRowLinkList(1 To ilNo32000, 1 To 32000) As ROWLINKLIST
            End If
            If ilRet <> 0 Then
                ilNo32000 = ilNo32000 - 5
                If ilNo32000 <= 0 Then
                    ilNo32000 = 1
                End If
            End If
        Loop While ilRet <> 0
        On Error GoTo 0
    End If
    If Not ilForMassChg Then
        For ilLoop = LBound(imFieldSizeCount) To UBound(imFieldSizeCount) Step 1
            If imFieldSizeCount(ilLoop) > 0 Then
                llNoFields = (vbcData.LargeChange + 1) * imFieldSizeCount(ilLoop) 'lmMaxNoRec * imFieldSizeCount(ilLoop)
                If llNoFields Mod 32000 <> 0 Then
                    ilNo32000 = llNoFields \ 32000 + 1
                Else
                    ilNo32000 = llNoFields \ 32000
                End If
                Select Case ilLoop
                    Case 1
                        If llNoFields <= 32000 Then
                            ReDim tmGrid1Data(1 To 1, 1 To llNoFields) As GRID1DATA
                        Else
                            ReDim tmGrid1Data(1 To ilNo32000, 1 To 32000) As GRID1DATA
                        End If
                    Case 2
                        If llNoFields <= 32000 Then
                            ReDim tmGrid2Data(1 To 1, 1 To llNoFields) As GRID2DATA
                        Else
                            ReDim tmGrid2Data(1 To ilNo32000, 1 To 32000) As GRID2DATA
                        End If
                    Case 3
                        If llNoFields <= 32000 Then
                            ReDim tmGrid3Data(1 To 1, 1 To llNoFields) As GRID3DATA
                        Else
                            ReDim tmGrid3Data(1 To ilNo32000, 1 To 32000) As GRID3DATA
                        End If
                    Case 4
                        If llNoFields <= 32000 Then
                            ReDim tmGrid4Data(1 To 1, 1 To llNoFields) As GRID4DATA
                        Else
                            ReDim tmGrid4Data(1 To ilNo32000, 1 To 32000) As GRID4DATA
                        End If
                    Case 5
                        If llNoFields <= 32000 Then
                            ReDim tmGrid5Data(1 To 1, 1 To llNoFields) As GRID5DATA
                        Else
                            ReDim tmGrid5Data(1 To ilNo32000, 1 To 32000) As GRID5DATA
                        End If
                    Case 6
                        If llNoFields <= 32000 Then
                            ReDim tmGrid6Data(1 To 1, 1 To llNoFields) As GRID6DATA
                        Else
                            ReDim tmGrid6Data(1 To ilNo32000, 1 To 32000) As GRID6DATA
                        End If
                End Select
            End If
        Next ilLoop
    End If
    If (imVarLenRec) Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary") Then
        If (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary") Then
            lmActualRecLen = Len(tmSsf)
        Else
            lmActualRecLen = Len(tmRec)
        End If
        imBaseRecLen = lmActualRecLen
        'ilRet = btrGetFirst(hmFile, tmRec, lmActualRecLen, ilKeyIndex, BTRV_LOCK_NONE)
        If (slSearchFrom = "") And (slSearchTo = "") Then
            ilRet = btrGetFirst(hmFile, tmRec, lmActualRecLen, ilKeyIndex, BTRV_LOCK_NONE)
        Else
            If slSearchFrom = "" Then
                ilRet = btrGetFirst(hmFile, tmRec, lmActualRecLen, ilKeyIndex, BTRV_LOCK_NONE)
            Else
                For ilLoop = LBound(tgKey.sChar) To UBound(tgKey.sChar) Step 1
                    tgKey.sChar(ilLoop) = 0
                Next ilLoop
                smFieldValue = slSearchFrom
                '7/15/10: Handle case where second field is descending
                If (StrComp(UCase$(Trim$(slFileName)), "CHF.BTR", vbTextCompare) = 0) And (ilKeyIndex = 1) Then
                    smFieldValue = Trim$(Str$(Val(slSearchFrom) - 1))
                ElseIf (StrComp(UCase$(Trim$(slFileName)), "CPTT.MKD", vbTextCompare) = 0) And ((ilKeyIndex = 1) Or (ilKeyIndex = 2)) Then
                    smFieldValue = Trim$(Str$(Val(slSearchFrom) - 1))
                End If
                tmSearchDDFField.iOffset = 0
                mPutField tmSearchDDFField, 1
                tmSearchDDFField.iOffset = ilSearchDDFOffset
                ilRet = btrGetGreaterOrEqual(hmFile, tmRec, lmActualRecLen, tgKey, ilKeyIndex, BTRV_LOCK_NONE)
            End If
        End If
        If ilRet <> BTRV_ERR_NONE Then
            Screen.MousePointer = 0 'DEFAULT
            mBtrvErrorMsg ilRet, "btrGetfirst: " & slFile
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            Exit Sub
        End If
        Do While (ilRet = BTRV_ERR_NONE)
            DoEvents
            If imTerminate Then
                Exit Do
                'Screen.MousePointer = 0 'DEFAULT
                'ilRet = btrClose(hmFile)
                'btrDestroy hmFile
                'Exit Sub
            End If
            If slSearchTo <> "" Then
                mGetField tmSearchDDFField, 0
                If mSearchCompleted(tmSearchDDFField, slSearchTo) Then
                    Exit Do
                End If
            End If
            If mRecordSelected() Then
                llNoShown = llNoShown + 1
                ilRet = btrGetPosition(hmFile, tmRecInfo(imRecUpper1, imRecUpper2).lRecPos)
                tmRecInfo(imRecUpper1, imRecUpper2).iRecLen = lmActualRecLen
                tmRecInfo(imRecUpper1, imRecUpper2).iRecIndex1 = imRecUpper1
                tmRecInfo(imRecUpper1, imRecUpper2).iRecIndex2 = imRecUpper2
                llRecNo = llRecNo + 1
                tmRecInfo(imRecUpper1, imRecUpper2).lRecNo = llRecNo
                imRecUpper2 = imRecUpper2 + 1
                If imRecUpper2 > 32000 Then
                    imRecUpper1 = imRecUpper1 + 1
                    imRecUpper2 = 1
                End If
            End If
            llNoProcessed = llNoProcessed + 1
            If (slSearchFrom = "") And (slSearchTo = "") Then
                lacTotal.Caption = "Read" & Str(llNoProcessed) & ", Selected" & Str(llNoShown) & ":" & Str(lmMaxNoRec) & " records"
            Else
                lacTotal.Caption = "Read" & Str(llNoProcessed) & ", Selected" & Str(llNoShown) & " records"
            End If
            llPercent = 100 * llNoProcessed / lmMaxNoRec
            If llPercent > 100 Then
                llPercent = 100
            End If
            If lmPercent <> llPercent Then
                plcGauge.Value = llPercent
                lmPercent = llPercent
            End If
            lmActualRecLen = imBaseRecLen
            ilRet = btrGetNext(hmFile, tmRec, lmActualRecLen, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                mFixLVarRecords
            End If
        Loop
    Else
        lmActualRecLen = 0
        For ilField = 0 To UBound(tmDDFField) - 1 Step 1
            lmActualRecLen = lmActualRecLen + tmDDFField(ilField).iSize
        Next ilField
        imBaseRecLen = lmActualRecLen
        ilExtLen = lmActualRecLen
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        btrExtClear hmFile   'Clear any previous extend operation
        If (slSearchFrom = "") And (slSearchTo = "") Then
            ilRet = btrGetFirst(hmFile, tmRec, lmActualRecLen, ilKeyIndex, BTRV_LOCK_NONE)
        Else
            If slSearchFrom = "" Then
                ilRet = btrGetFirst(hmFile, tmRec, lmActualRecLen, ilKeyIndex, BTRV_LOCK_NONE)
            Else
                For ilLoop = LBound(tgKey.sChar) To UBound(tgKey.sChar) Step 1
                    tgKey.sChar(ilLoop) = 0
                Next ilLoop
                smFieldValue = slSearchFrom
                '7/15/10: Handle case where second field is descending
                If (StrComp(UCase$(Trim$(slFileName)), "CHF.BTR", vbTextCompare) = 0) And (ilKeyIndex = 1) Then
                    smFieldValue = Trim$(Str$(Val(slSearchFrom) - 1))
                ElseIf (StrComp(UCase$(Trim$(slFileName)), "CPTT.MKD", vbTextCompare) = 0) And ((ilKeyIndex = 1) Or (ilKeyIndex = 2)) Then
                    smFieldValue = Trim$(Str$(Val(slSearchFrom) - 1))
                End If
                tmSearchDDFField.iOffset = 0
                mPutField tmSearchDDFField, 1
                tmSearchDDFField.iOffset = ilSearchDDFOffset
                ilRet = btrGetGreaterOrEqual(hmFile, tmRec, lmActualRecLen, tgKey, ilKeyIndex, BTRV_LOCK_NONE)
            End If
        End If
        If ilRet <> BTRV_ERR_NONE Then
            Screen.MousePointer = 0 'DEFAULT
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                mBtrvErrorMsg ilRet, "btrGetfirst: " & slFile
            End If
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            Exit Sub
        End If
        Call btrExtSetBounds(hmFile, llNoRec, -1, "UC") 'Set extract limits (all records)
        If (slSearchFrom <> "") Or (slSearchTo <> "") Then
            If slSearchFrom <> "" Then
                For ilLoop = LBound(tgKey.sChar) To UBound(tgKey.sChar) Step 1
                    tgKey.sChar(ilLoop) = 0
                Next ilLoop
                smFieldValue = slSearchFrom
                '7/15/10: Handle case where second field is descending
                If (StrComp(UCase$(Trim$(slFileName)), "CHF.BTR", vbTextCompare) = 0) And (ilKeyIndex = 1) Then
                    smFieldValue = Trim$(Str$(Val(slSearchFrom) - 1))
                ElseIf (StrComp(UCase$(Trim$(slFileName)), "CPTT.MKD", vbTextCompare) = 0) And ((ilKeyIndex = 1) Or (ilKeyIndex = 2)) Then
                    smFieldValue = Trim$(Str$(Val(slSearchFrom) - 1))
                End If
                tmSearchDDFField.iOffset = 0
                mPutField tmSearchDDFField, 1
                tmSearchDDFField.iOffset = ilSearchDDFOffset
                If slSearchTo = "" Then
                    ilRet = btrExtAddLogicConst(hmFile, Asc(tmSearchDDFField.sDataType), tmSearchDDFField.iOffset, tmSearchDDFField.iSize, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tgKey, tmSearchDDFField.iSize)
                Else
                    ilRet = btrExtAddLogicConst(hmFile, Asc(tmSearchDDFField.sDataType), tmSearchDDFField.iOffset, tmSearchDDFField.iSize, BTRV_EXT_GTE, BTRV_EXT_AND, tgKey, tmSearchDDFField.iSize)
                    For ilLoop = LBound(tgKey.sChar) To UBound(tgKey.sChar) Step 1
                        tgKey.sChar(ilLoop) = 0
                    Next ilLoop
                    smFieldValue = slSearchTo
                    tmSearchDDFField.iOffset = 0
                    mPutField tmSearchDDFField, 1
                    tmSearchDDFField.iOffset = ilSearchDDFOffset
                    ilRet = btrExtAddLogicConst(hmFile, Asc(tmSearchDDFField.sDataType), tmSearchDDFField.iOffset, tmSearchDDFField.iSize, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tgKey, tmSearchDDFField.iSize)
                End If
            Else
                For ilLoop = LBound(tgKey.sChar) To UBound(tgKey.sChar) Step 1
                    tgKey.sChar(ilLoop) = 0
                Next ilLoop
                smFieldValue = slSearchTo
                tmSearchDDFField.iOffset = 0
                mPutField tmSearchDDFField, 1
                tmSearchDDFField.iOffset = ilSearchDDFOffset
                ilRet = btrExtAddLogicConst(hmFile, Asc(tmSearchDDFField.sDataType), tmSearchDDFField.iOffset, tmSearchDDFField.iSize, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tgKey, tmSearchDDFField.iSize)
            End If
        End If
        ilRet = btrExtAddField(hmFile, 0, ilExtLen)  'Extract iCode field
        If ilRet <> BTRV_ERR_NONE Then
            Screen.MousePointer = 0 'DEFAULT
            mBtrvErrorMsg ilRet, "btrGetfirst: " & slFile
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            Exit Sub
        End If
        ilRet = btrExtGetNext(hmFile, tmRec, CLng(ilExtLen), llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Screen.MousePointer = 0 'DEFAULT
                mBtrvErrorMsg ilRet, "btrGetfirst: " & slFile
                ilRet = btrClose(hmFile)
                btrDestroy hmFile
                Exit Sub
            End If
            ilExtLen = lmActualRecLen  'Extract operation record size
            'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(imRowUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmFile, tmRec, CLng(ilExtLen), llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                DoEvents
                If imTerminate Then
                    Exit Do
                    'Screen.MousePointer = 0 'DEFAULT
                    'ilRet = btrClose(hmFile)
                    'btrDestroy hmFile
                    'Exit Sub
                End If
                If mRecordSelected() Then
                    llNoShown = llNoShown + 1
                    tmRecInfo(imRecUpper1, imRecUpper2).lRecPos = llRecPos
                    tmRecInfo(imRecUpper1, imRecUpper2).iRecLen = lmActualRecLen
                    tmRecInfo(imRecUpper1, imRecUpper2).iRecIndex1 = imRecUpper1
                    tmRecInfo(imRecUpper1, imRecUpper2).iRecIndex2 = imRecUpper2
                    llRecNo = llRecNo + 1
                    tmRecInfo(imRecUpper1, imRecUpper2).lRecNo = llRecNo
                    imRecUpper2 = imRecUpper2 + 1
                    If imRecUpper2 > 32000 Then
                        imRecUpper1 = imRecUpper1 + 1
                        imRecUpper2 = 1
                    End If
                End If
                llNoProcessed = llNoProcessed + 1
                If (slSearchFrom = "") And (slSearchTo = "") Then
                    lacTotal.Caption = "Read" & Str(llNoProcessed) & ", Selected" & Str(llNoShown) & ":" & Str(lmMaxNoRec) & " records"
                Else
                    lacTotal.Caption = "Read" & Str(llNoProcessed) & ", Selected" & Str(llNoShown) & " records"
                End If
                llPercent = 100 * llNoProcessed / lmMaxNoRec
                If llPercent > 100 Then
                    llPercent = 100
                End If
                If lmPercent <> llPercent Then
                    plcGauge.Value = llPercent
                    lmPercent = llPercent
                End If
                ilRet = btrExtGetNext(hmFile, tmRec, CLng(ilExtLen), llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmFile, tmRec, CLng(ilExtLen), llRecPos)
                Loop
            Loop
        End If
    End If
    plcGauge.Value = 100
    DoEvents
    imRecMaxUpper1 = imRecUpper1
    imRecMaxUpper2 = imRecUpper2
    If Not ilForMassChg Then
        vbcData.Min = 1
    End If
    lacTotal.Caption = Trim$(Str$(llNoShown)) & " Records Selected out of" & Str$(llNoProcessed)
    lmMaxNoRec = llNoShown
    If Not ilForMassChg Then
        If lmMaxNoRec <= 32000 Then
            imVBCAdjust = 1
            If imRecUpper2 - 1 <= vbcData.LargeChange + 1 Then
                vbcData.Max = 1
            Else
                vbcData.Max = imRecUpper2 - vbcData.LargeChange
            End If
        Else
            If lmMaxNoRec Mod 32000 <> 0 Then
                imVBCAdjust = lmMaxNoRec \ 32000 + 1
            Else
                imVBCAdjust = lmMaxNoRec \ 32000
            End If
            vbcData.Max = (lmMaxNoRec - vbcData.LargeChange) \ imVBCAdjust + imVBCAdjust - 1
        End If
        imClsGrid = False
        vbcData_Change
    Else
        imClsGrid = False
        If Not imTerminate Then
            'Update records
            lacChg.Visible = True
            plcChg.Visible = True
            plcChg.Caption = "0&"
            lmPercent = 0
            If rbcEditReplace(3).Value Then
                mMassRemoveRecords
            Else
                mMassUpdateRecords
            End If
        End If
    End If
    Screen.MousePointer = 0 'DEFAULT
    Exit Sub
mObtainRecordsErr:
    ilRet = Err
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopFieldBoxes                  *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate field list boxes       *
'*                                                     *
'*******************************************************
Private Sub mPopFieldBoxes()
    Dim slName As String
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilLink As Integer
    Dim ilFound As Integer
    Dim ilStartIndex As Integer
    Dim ilEndIndex As Integer
    'Dim ilCount As Integer
    ReDim ilFileID(0 To 0) As Integer
    
    
    If lbcFile.ListIndex >= 0 Then
        slName = lbcFile.List(lbcFile.ListIndex)
        For ilLoop = 0 To UBound(tmDDFFile) - 1 Step 1
            If StrComp(Trim$(tmDDFFile(ilLoop).sName), slName, 1) = 0 Then
                ilFileID(0) = tmDDFFile(ilLoop).iFileID
                ReDim Preserve ilFileID(0 To 1) As Integer
                Exit For
            End If
        Next ilLoop
        'Build all other IDs previousely processed
        For ilLoop = 0 To UBound(tmDDFField) - 1 Step 1
            ilFound = False
            For ilTest = 0 To UBound(ilFileID) - 1 Step 1
                If ilFileID(ilTest) = tmDDFField(ilLoop).iFileID Then
                    ilFound = True
                    Exit For
                End If
            Next ilTest
            If Not ilFound Then
                ilFileID(UBound(ilFileID)) = tmDDFField(ilLoop).iFileID
                ReDim Preserve ilFileID(0 To UBound(ilFileID) + 1) As Integer
            End If
        Next ilLoop
        'Remove any link files removed
        For ilLoop = 1 To UBound(ilFileID) - 1 Step 1
            ilFound = False
            For ilLink = 0 To UBound(tmLink) - 1 Step 1
                If tmLink(ilLink).iToFileID = ilFileID(ilLoop) Then
                    ilFound = True
                    Exit For
                End If
                If Not ilFound Then
                    'Remove values from tmDDFField, lbcSelection(1), lbcEditReplace(0) and (2)
                End If
            Next ilLink
        Next ilLoop
        'Add any new files
        ilStartIndex = UBound(tmDDFField)
        For ilLink = 0 To UBound(tmLink) - 1 Step 1
            ilFound = False
            For ilTest = 0 To UBound(ilFileID) - 1 Step 1
                If ilFileID(ilTest) = tmLink(ilLink).iToFileID Then
                    ilFound = True
                    Exit For
                End If
            Next ilTest
            If Not ilFound Then
                ilFileID(UBound(ilFileID)) = tmLink(ilLink).iToFileID
                ReDim Preserve ilFileID(0 To UBound(ilFileID) + 1) As Integer
                mFieldBuild tmLink(ilLink).iToFileID, tmDDFField()
            End If
        Next ilLink
        mFieldPop tmDDFField(), ilStartIndex, UBound(tmDDFField) - 1, lbcSelection(1), True
        If rbcEditReplace(2).Value Then
            ''Build lbcEditReplace(0) from base file only
            'ilCount = 0
            'For ilLoop = 0 To UBound(tmDDFField) Step 1
            '    If tmDDFField(ilLoop).iFileID = ilFileID(0) Then
            '        ilCount = ilCount + 1
            '    End If
            'Next ilLoop
            'For ilLoop = lbcEditReplace(0).ListCount - 1 To ilCount Step 1
            '    lbcEditReplace(0).RemoveItem ilLoop
            'Next ilLoop
            lbcEditReplace(0).Clear
            For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1
                lbcEditReplace(0).AddItem lbcSelection(1).List(ilLoop)
            Next ilLoop
        ElseIf rbcEditReplace(3).Value Then
            lbcEditReplace(0).Clear
        Else
            'For ilLoop = lbcEditReplace(0).ListCount To lbcSelection(1).ListCount - 1 Step 1
            '    lbcEditReplace(0).AddItem lbcSelection(1).List(ilLoop)
            'Next ilLoop
            lbcEditReplace(0).Clear
            For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1
                lbcEditReplace(0).AddItem lbcSelection(1).List(ilLoop)
            Next ilLoop
        End If
        pbcLbcEditReplace0_Paint
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPutField                       *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Put field into record           *
'*                                                     *
'*******************************************************
Private Sub mPutField(tlDDFField As DDFFIELD, ilRecType As Integer)
'
'   ilRecType: 0= tmRec; 1=tgKey
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilSize As Integer
    Dim ilValue As Integer
    Dim llValue As Long
    Dim ilDyMn As Integer
    Dim ilYear As Integer
    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim slMoney As String
    Dim slInt2 As INT2
    Dim slInt4 As INT4
    Dim slChar As String
    ilIndex = tlDDFField.iOffset + 1
    ilSize = tlDDFField.iSize
    Select Case Asc(tlDDFField.sDataType)
        Case 0  'String
            For ilLoop = 1 To ilSize Step 1
                If Len(smFieldValue) > 0 Then
                    If ilRecType = 0 Then
                        'tmRec.sChar(ilIndex) = left$(smFieldValue, 1)
                        tmRec.sChar(ilIndex) = Asc(Left$(smFieldValue, 1))
                    Else
                        'tgKey.sChar(ilIndex) = left$(smFieldValue, 1)
                        tgKey.sChar(ilIndex) = Asc(Left$(smFieldValue, 1))
                    End If
                    smFieldValue = Mid(smFieldValue, 2)
                Else
                    If ilRecType = 0 Then
                        'tmRec.sChar(ilIndex) = " "
                        tmRec.sChar(ilIndex) = Asc(" ")
                    Else
                        'tgKey.sChar(ilIndex) = " "
                        tgKey.sChar(ilIndex) = Asc(" ")
                    End If
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 1  'Integer
            If ilSize = 1 Then
                ilValue = Val(smFieldValue)
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = Chr(ilValue)
                    tmRec.sChar(ilIndex) = ilValue And &HFF
                Else
                    'tgKey.sChar(ilIndex) = Chr(ilValue)
                    tgKey.sChar(ilIndex) = ilValue And &HFF
                End If
            ElseIf ilSize = 2 Then
                ilValue = Val(smFieldValue)
                'HMemCpy slInt2, ilValue, 2
                'If ilRecType = 0 Then
                '    tmRec.sChar(ilIndex) = left$(slInt2.sInt2, 1)
                '    tmRec.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
                'Else
                '    tgKey.sChar(ilIndex) = left$(slInt2.sInt2, 1)
                '    tgKey.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
                'End If
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = ilValue And &HFF
                    'tmRec.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                    HMemCpy tmRec.sChar(ilIndex), ilValue, 2
                Else
                    'tgKey.sChar(ilIndex) = ilValue And &HFF
                    'tgKey.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                    HMemCpy tgKey.sChar(ilIndex), ilValue, 2
                End If
            ElseIf ilSize = 4 Then
                llValue = Val(smFieldValue)
                'HMemCpy slInt4, llValue, 4
                'If ilRecType = 0 Then
                '    tmRec.sChar(ilIndex) = left$(slInt4.sInt4, 1)
                '    tmRec.sChar(ilIndex + 1) = Mid$(slInt4.sInt4, 2, 1)
                '    tmRec.sChar(ilIndex + 2) = Mid$(slInt4.sInt4, 3, 1)
                '    tmRec.sChar(ilIndex + 3) = Mid$(slInt4.sInt4, 4, 1)
                'Else
                '    tgKey.sChar(ilIndex) = left$(slInt4.sInt4, 1)
                '    tgKey.sChar(ilIndex + 1) = Mid$(slInt4.sInt4, 2, 1)
                '    tgKey.sChar(ilIndex + 2) = Mid$(slInt4.sInt4, 3, 1)
                '    tgKey.sChar(ilIndex + 3) = Mid$(slInt4.sInt4, 4, 1)
                'End If
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex + 2) = (llValue \ 65536) And &HFF
                    'tmRec.sChar(ilIndex + 3) = ((llValue \ 65536) \ 256) And &HFF
                    'tmRec.sChar(ilIndex) = llValue And &HFF
                    'tmRec.sChar(ilIndex + 1) = (llValue \ 256) And &HFF
                    HMemCpy tmRec.sChar(ilIndex), llValue, 4
                Else
                    'tgKey.sChar(ilIndex + 2) = (llValue \ 65536) And &HFF
                    'tgKey.sChar(ilIndex + 3) = ((llValue \ 65536) \ 256) And &HFF
                    'tgKey.sChar(ilIndex) = llValue And &HFF
                    'tgKey.sChar(ilIndex + 1) = (llValue \ 256) And &HFF
                    HMemCpy tgKey.sChar(ilIndex), llValue, 4
                End If
            End If
        Case 2  'IEEE Float
        Case 3  'Btrieve Date
            gPackDate smFieldValue, ilDyMn, ilYear
            'HMemCpy slInt2, ilDyMn, 2
            'If ilRecType = 0 Then
            '    tmRec.sChar(ilIndex) = left$(slInt2.sInt2, 1)
            '    tmRec.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
            'Else
            '    tgKey.sChar(ilIndex) = left$(slInt2.sInt2, 1)
            '    tgKey.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
            'End If
            If ilRecType = 0 Then
                tmRec.sChar(ilIndex) = ilDyMn And &HFF
                tmRec.sChar(ilIndex + 1) = (ilDyMn \ 256) And &HFF
            Else
                tgKey.sChar(ilIndex) = ilDyMn And &HFF
                tgKey.sChar(ilIndex + 1) = (ilDyMn \ 256) And &HFF
            End If
            'HMemCpy slInt2, ilYear, 2
            'If ilRecType = 0 Then
            '    tmRec.sChar(ilIndex + 2) = left$(slInt2.sInt2, 1)
            '    tmRec.sChar(ilIndex + 3) = Mid$(slInt2.sInt2, 2, 1)
            'Else
            '    tgKey.sChar(ilIndex + 2) = left$(slInt2.sInt2, 1)
            '    tgKey.sChar(ilIndex + 3) = Mid$(slInt2.sInt2, 2, 1)
            'End If
            If ilRecType = 0 Then
                tmRec.sChar(ilIndex + 2) = ilYear And &HFF
                tmRec.sChar(ilIndex + 3) = (ilYear \ 256) And &HFF
            Else
                tgKey.sChar(ilIndex + 2) = ilYear And &HFF
                tgKey.sChar(ilIndex + 3) = (ilYear \ 256) And &HFF
            End If
        Case 4  'Btrieve Time
            gPackTime smFieldValue, ilHsSec, ilMinHr
            'HMemCpy slInt2, ilHsSec, 2
            'If ilRecType = 0 Then
            '    tmRec.sChar(ilIndex) = left$(slInt2.sInt2, 1)
            '    tmRec.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
            'Else
            '    tgKey.sChar(ilIndex) = left$(slInt2.sInt2, 1)
            '    tgKey.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
            'End If
            If ilRecType = 0 Then
                tmRec.sChar(ilIndex) = ilHsSec And &HFF
                tmRec.sChar(ilIndex + 1) = (ilHsSec \ 256) And &HFF
            Else
                tgKey.sChar(ilIndex) = ilHsSec And &HFF
                tgKey.sChar(ilIndex + 1) = (ilHsSec \ 256) And &HFF
            End If
            'HMemCpy slInt2, ilMinHr, 2
            'If ilRecType = 0 Then
            '    tmRec.sChar(ilIndex + 2) = left$(slInt2.sInt2, 1)
            '    tmRec.sChar(ilIndex + 3) = Mid$(slInt2.sInt2, 2, 1)
            'Else
            '    tgKey.sChar(ilIndex + 2) = left$(slInt2.sInt2, 1)
            '    tgKey.sChar(ilIndex + 3) = Mid$(slInt2.sInt2, 2, 1)
            'End If
            If ilRecType = 0 Then
                tmRec.sChar(ilIndex + 2) = ilMinHr And &HFF
                tmRec.sChar(ilIndex + 3) = (ilMinHr \ 256) And &HFF
            Else
                tgKey.sChar(ilIndex + 2) = ilMinHr And &HFF
                tgKey.sChar(ilIndex + 3) = (ilMinHr \ 256) And &HFF
            End If
        Case 5  'COBOL Decimal COMP-3
            gStrToPDN smFieldValue, Asc(tlDDFField.sDec), ilSize, slMoney
            For ilLoop = 1 To ilSize Step 1
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = left$(slMoney, 1)
                    tmRec.sChar(ilIndex) = Asc(Left$(slMoney, 1))
                Else
                    'tgKey.sChar(ilIndex) = left$(slMoney, 1)
                    tgKey.sChar(ilIndex) = Asc(Left$(slMoney, 1))
                End If
                slMoney = Mid(slMoney, 2)
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 6  'COBOL Money
            gStrToPDN smFieldValue, Asc(tlDDFField.sDec), ilSize, slMoney
            For ilLoop = 1 To ilSize Step 1
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = left$(slMoney, 1)
                    tmRec.sChar(ilIndex) = Asc(Left$(slMoney, 1))
                Else
                    'tgKey.sChar(ilIndex) = left$(slMoney, 1)
                    tgKey.sChar(ilIndex) = Asc(Left$(slMoney, 1))
                End If
                slMoney = Mid(slMoney, 2)
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 7  'Logical
            If ilSize = 1 Then
                ilValue = Val(smFieldValue)
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = Chr(ilValue)
                    tmRec.sChar(ilIndex) = ilValue And &HFF
                Else
                    'tgKey.sChar(ilIndex) = Chr(ilValue)
                    tgKey.sChar(ilIndex) = ilValue And &HFF
                End If
            ElseIf ilSize = 2 Then
                ilValue = Val(smFieldValue)
                'HMemCpy slInt2, ilValue, 2
                'If ilRecType = 0 Then
                '    tmRec.sChar(ilIndex) = left$(slInt2.sInt2, 1)
                '    tmRec.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
                'Else
                '    tgKey.sChar(ilIndex) = left$(slInt2.sInt2, 1)
                '    tgKey.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
                'End If
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = ilValue And &HFF
                    'tmRec.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                    HMemCpy tmRec.sChar(ilIndex), ilValue, 2
                Else
                    'tgKey.sChar(ilIndex) = ilValue And &HFF
                    'tgKey.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                    HMemCpy tgKey.sChar(ilIndex), ilValue, 2
                End If
            End If
        Case 8  'COBOL Numeric
        Case 9  'BASIC Float
        Case 10 'Pascal LString
            ilValue = Len(smFieldValue)
            If ilRecType = 0 Then
                'tmRec.sChar(ilIndex) = Chr(ilValue)
                tmRec.sChar(ilIndex) = ilValue And &HFF
            Else
                'tgKey.sChar(ilIndex) = Chr(ilValue)
                tgKey.sChar(ilIndex) = ilValue And &HFF
            End If
            ilIndex = ilIndex + 1
            For ilLoop = 1 To ilValue Step 1
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = left$(smFieldValue, 1)
                    tmRec.sChar(ilIndex) = Asc(Left$(smFieldValue, 1))
                Else
                    'tgKey.sChar(ilIndex) = left$(smFieldValue, 1)
                    tgKey.sChar(ilIndex) = Asc(Left$(smFieldValue, 1))
                End If
                smFieldValue = Mid(smFieldValue, 2)
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 11 'C ZString
            ilValue = Len(smFieldValue)
            For ilLoop = 1 To ilValue Step 1
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = left$(smFieldValue, 1)
                    tmRec.sChar(ilIndex) = Asc(Left$(smFieldValue, 1))
                Else
                    'tgKey.sChar(ilIndex) = left$(smFieldValue, 1)
                    tgKey.sChar(ilIndex) = Asc(Left$(smFieldValue, 1))
                End If
                smFieldValue = Mid(smFieldValue, 2)
                ilIndex = ilIndex + 1
            Next ilLoop
            If ilRecType = 0 Then
                'tmRec.sChar(ilIndex) = Chr(0)
                tmRec.sChar(ilIndex) = 0
            Else
                'tgKey.sChar(ilIndex) = Chr(0)
                tgKey.sChar(ilIndex) = 0
            End If
        Case 12 'Variable Length Note
        Case 13 'LVar (Length, Characters, 0, 0); Length does not include zeros
            ilValue = Len(smFieldValue)
            'HMemCpy slInt2, ilValue, 2
            'If ilRecType = 0 Then
            '    tmRec.sChar(ilIndex) = left$(slInt2.sInt2, 1)
            '    tmRec.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
            'Else
            '    tgKey.sChar(ilIndex) = left$(slInt2.sInt2, 1)
            '    tgKey.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
            'End If
            If ilRecType = 0 Then
                'tmRec.sChar(ilIndex) = ilValue And &HFF
                'tmRec.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                HMemCpy tmRec.sChar(ilIndex), ilValue, 2
            Else
                'tgKey.sChar(ilIndex) = ilValue And &HFF
                'tgKey.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                HMemCpy tgKey.sChar(ilIndex), ilValue, 2
            End If
            ilIndex = ilIndex + 2
            For ilLoop = 1 To ilValue Step 1
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = left$(smFieldValue, 1)
                    tmRec.sChar(ilIndex) = Asc(Left$(smFieldValue, 1))
                Else
                    'tgKey.sChar(ilIndex) = left$(smFieldValue, 1)
                    tgKey.sChar(ilIndex) = Asc(Left$(smFieldValue, 1))
                End If
                smFieldValue = Mid(smFieldValue, 2)
                ilIndex = ilIndex + 1
            Next ilLoop
            'If ilRecType = 0 Then
            '    tmRec.sChar(ilIndex) = Chr(0)
            '    tmRec.sChar(ilIndex + 1) = Chr(0)
            'Else
            '    tgKey.sChar(ilIndex) = Chr(0)
            '    tgKey.sChar(ilIndex + 1) = Chr(0)
            'End If
            If ilRecType = 0 Then
                tmRec.sChar(ilIndex) = 0
                tmRec.sChar(ilIndex + 1) = 0
            Else
                tgKey.sChar(ilIndex) = 0
                tgKey.sChar(ilIndex + 1) = 0
            End If
        Case 14 'Unsigned Binary
            llValue = Val(smFieldValue)
            ilValue = gLongToUnsignInt(llValue)
            'HMemCpy slInt2, ilValue, 2
            'If ilRecType = 0 Then
            '    tmRec.sChar(ilIndex) = left$(slInt2.sInt2, 1)
            '    tmRec.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
            'Else
            '    tgKey.sChar(ilIndex) = left$(slInt2.sInt2, 1)
            '    tgKey.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
            'End If
            If ilRecType = 0 Then
                'tmRec.sChar(ilIndex) = ilValue And &HFF
                'tmRec.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                HMemCpy tmRec.sChar(ilIndex), ilValue, 2
            Else
                'tgKey.sChar(ilIndex) = ilValue And &HFF
                'tgKey.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                HMemCpy tgKey.sChar(ilIndex), ilValue, 2
            End If
        Case 15 'AutoIncrement
            If ilSize = 2 Then
                ilValue = Val(smFieldValue)
                'HMemCpy slInt2, ilValue, 2
                'If ilRecType = 0 Then
                '    tmRec.sChar(ilIndex) = left$(slInt2.sInt2, 1)
                '    tmRec.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
                'Else
                '    tgKey.sChar(ilIndex) = left$(slInt2.sInt2, 1)
                '    tgKey.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
                'End If
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = ilValue And &HFF
                    'tmRec.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                    HMemCpy tmRec.sChar(ilIndex), ilValue, 2
                Else
                    'tgKey.sChar(ilIndex) = ilValue And &HFF
                    'tgKey.sChar(ilIndex + 1) = (ilValue \ 256) And &HFF
                    HMemCpy tgKey.sChar(ilIndex), ilValue, 2
                End If
            ElseIf ilSize = 4 Then
                llValue = Val(smFieldValue)
                'HMemCpy slInt4, llValue, 4
                'If ilRecType = 0 Then
                '    tmRec.sChar(ilIndex) = left$(slInt4.sInt4, 1)
                '    tmRec.sChar(ilIndex + 1) = Mid$(slInt4.sInt4, 2, 1)
                '    tmRec.sChar(ilIndex + 2) = Mid$(slInt4.sInt4, 3, 1)
                '    tmRec.sChar(ilIndex + 3) = Mid$(slInt4.sInt4, 4, 1)
                'Else
                '    tgKey.sChar(ilIndex) = left$(slInt4.sInt4, 1)
                '    tgKey.sChar(ilIndex + 1) = Mid$(slInt4.sInt4, 2, 1)
                '    tgKey.sChar(ilIndex + 2) = Mid$(slInt4.sInt4, 3, 1)
                '    tgKey.sChar(ilIndex + 3) = Mid$(slInt4.sInt4, 4, 1)
                'End If
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex + 2) = (llValue \ 65536) And &HFF
                    'tmRec.sChar(ilIndex + 3) = ((llValue \ 65536) \ 256) And &HFF
                    'tmRec.sChar(ilIndex) = llValue And &HFF
                    'tmRec.sChar(ilIndex + 1) = (llValue \ 256) And &HFF
                    HMemCpy tmRec.sChar(ilIndex), llValue, 4
                Else
                    'tgKey.sChar(ilIndex + 2) = (llValue \ 65536) And &HFF
                    'tgKey.sChar(ilIndex + 3) = ((llValue \ 65536) \ 256) And &HFF
                    'tgKey.sChar(ilIndex) = llValue And &HFF
                    'tgKey.sChar(ilIndex + 1) = (llValue \ 256) And &HFF
                    HMemCpy tgKey.sChar(ilIndex), llValue, 4
                End If
            End If
        Case 16 'Bit
        Case 17 'COBOL Numeric STS
        Case 100  'Length
            gPackLength smFieldValue, ilHsSec, ilMinHr
            'HMemCpy slInt2, ilHsSec, 2
            'If ilRecType = 0 Then
            '    tmRec.sChar(ilIndex) = left$(slInt2.sInt2, 1)
            '    tmRec.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
            'Else
            '    tgKey.sChar(ilIndex) = left$(slInt2.sInt2, 1)
            '    tgKey.sChar(ilIndex + 1) = Mid$(slInt2.sInt2, 2, 1)
            'End If
            If ilRecType = 0 Then
                tmRec.sChar(ilIndex) = ilHsSec And &HFF
                tmRec.sChar(ilIndex + 1) = (ilHsSec \ 256) And &HFF
            Else
                tgKey.sChar(ilIndex) = ilHsSec And &HFF
                tgKey.sChar(ilIndex + 1) = (ilHsSec \ 256) And &HFF
            End If
            'HMemCpy slInt2, ilMinHr, 2
            'If ilRecType = 0 Then
            '    tmRec.sChar(ilIndex + 2) = left$(slInt2.sInt2, 1)
            '    tmRec.sChar(ilIndex + 3) = Mid$(slInt2.sInt2, 2, 1)
            'Else
            '    tgKey.sChar(ilIndex + 2) = left$(slInt2.sInt2, 1)
            '    tgKey.sChar(ilIndex + 3) = Mid$(slInt2.sInt2, 2, 1)
            'End If
            If ilRecType = 0 Then
                tmRec.sChar(ilIndex + 2) = ilMinHr And &HFF
                tmRec.sChar(ilIndex + 3) = (ilMinHr \ 256) And &HFF
            Else
                tgKey.sChar(ilIndex + 2) = ilMinHr And &HFF
                tgKey.sChar(ilIndex + 3) = (ilMinHr \ 256) And &HFF
            End If
        Case 101
            ilLoop = 1
            Do While ilLoop <= Len(smFieldValue)
                slChar = Mid$(smFieldValue, ilLoop, 1)
                'Ignore leading blanks
                Do While slChar = " "
                    ilLoop = ilLoop + 1
                    If ilLoop > Len(smFieldValue) Then
                        Exit Sub
                    End If
                    slChar = Mid$(smFieldValue, ilLoop, 1)
                Loop
                If (slChar >= "0") And (slChar <= "9") Then
                    ilValue = Val(slChar)
                ElseIf (slChar >= "A") And (slChar <= "F") Then
                    ilValue = (Asc(slChar) - Asc("A") + 10)
                Else
                    ilValue = 0
                End If
                ilLoop = ilLoop + 1
                If ilLoop <= Len(smFieldValue) Then
                    slChar = Mid$(smFieldValue, ilLoop, 1)
                    If (slChar >= "0") And (slChar <= "9") Then
                        ilValue = 16 * ilValue + Val(slChar)
                    ElseIf (slChar >= "A") And (slChar <= "F") Then
                        ilValue = 16 * ilValue + (Asc(slChar) - Asc("A") + 10)
                    Else
                        ilValue = ilValue + 0   'Don't multiply by 16 as only one character
                    End If
                End If
                ilLoop = ilLoop + 1
                If ilRecType = 0 Then
                    'tmRec.sChar(ilIndex) = Chr(ilValue)
                    tmRec.sChar(ilIndex) = ilValue And &HFF
                Else
                    'tgKey.sChar(ilIndex) = Chr(ilValue)
                    tgKey.sChar(ilIndex) = ilValue And &HFF
                End If
                ilIndex = ilIndex + 1
            Loop
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadSpec                       *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Import specification            *
'*                                                     *
'*******************************************************
Private Sub mReadSpec(ilGetDescriptionOnly As Integer)
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilImportArea As Integer '1=Generate; 2=Description; 3=File; 4=Links; 5=Selections; 6=Task; 7=Fields; 8=Run
    Dim ilLoop As Integer
    Dim slSelection As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    Dim slText As String
    Dim ilFirstField As Integer
    Dim ilUpper As Integer
    Dim slStr As String
    smSpecDescript = ""
    slSelection = ""
    ilImportArea = -1
    ilEof = False
    ilRet = 0
    If Not ilGetDescriptionOnly Then
        lbcEditReplace(2).Clear
        ReDim tmLink(0 To 0) As LINK
    End If
    ilFirstField = True
    Do
        On Error GoTo mReadSpecErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                If slLine = "[Generated]" Then
                    If ilImportArea = 2 Then
                        edcFDescription.Text = smSpecDescript
                        edcDescription.Text = smSpecDescript
                    ElseIf ilImportArea = 5 Then
                        If Not ilGetDescriptionOnly Then
                            edcSelection.Text = slSelection
                        End If
                    End If
                    ilImportArea = 1
                ElseIf slLine = "[Description]" Then
                    If ilImportArea = 2 Then
                        edcFDescription.Text = smSpecDescript
                    ElseIf ilImportArea = 5 Then
                        If Not ilGetDescriptionOnly Then
                            edcSelection.Text = slSelection
                        End If
                    End If
                    ilImportArea = 2
                ElseIf slLine = "[File]" Then
                    If ilImportArea = 2 Then
                        edcFDescription.Text = smSpecDescript
                    ElseIf ilImportArea = 5 Then
                        If Not ilGetDescriptionOnly Then
                            edcSelection.Text = slSelection
                        End If
                    End If
                    ilImportArea = 3
                ElseIf slLine = "[Links]" Then
                    If ilImportArea = 2 Then
                        edcFDescription.Text = smSpecDescript
                    ElseIf ilImportArea = 5 Then
                        If Not ilGetDescriptionOnly Then
                            edcSelection.Text = slSelection
                        End If
                    End If
                    ilImportArea = 4
                ElseIf slLine = "[Selections]" Then
                    If ilImportArea = 2 Then
                        edcFDescription.Text = smSpecDescript
                    ElseIf ilImportArea = 5 Then
                        If Not ilGetDescriptionOnly Then
                            edcSelection.Text = slSelection
                        End If
                    End If
                    ilImportArea = 5
                ElseIf slLine = "[Task]" Then
                    If ilImportArea = 2 Then
                        edcFDescription.Text = smSpecDescript
                    ElseIf ilImportArea = 5 Then
                        If Not ilGetDescriptionOnly Then
                            edcSelection.Text = slSelection
                        End If
                    End If
                    ilImportArea = 6
                ElseIf slLine = "[Fields]" Then
                    If ilImportArea = 2 Then
                        edcFDescription.Text = smSpecDescript
                    ElseIf ilImportArea = 5 Then
                        If Not ilGetDescriptionOnly Then
                            edcSelection.Text = slSelection
                        End If
                    End If
                    ilImportArea = 7
                ElseIf slLine = "[Key]" Then
                    If ilImportArea = 2 Then
                        edcFDescription.Text = smSpecDescript
                    ElseIf ilImportArea = 5 Then
                        If Not ilGetDescriptionOnly Then
                            edcSelection.Text = slSelection
                        End If
                    End If
                    ilImportArea = 8
                ElseIf slLine = "[Run]" Then
                    If ilImportArea = 2 Then
                        edcFDescription.Text = smSpecDescript
                    ElseIf ilImportArea = 5 Then
                        If Not ilGetDescriptionOnly Then
                            edcSelection.Text = slSelection
                        End If
                    End If
                    ilImportArea = 9
                Else
                    If Len(slLine) > 0 Then
                        If ilImportArea = 1 Then    'Generate- ignore line
                        ElseIf ilImportArea = 2 Then    'Description
                            smSpecDescript = smSpecDescript & " " & slLine
                        ElseIf ilImportArea = 3 Then    'File
                            If Not ilGetDescriptionOnly Then
                                For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
                                    If lbcFile.List(ilLoop) = slLine Then
                                        lbcFile.ListIndex = ilLoop
                                        Exit For
                                    End If
                                Next ilLoop
                            End If
                        ElseIf ilImportArea = 4 Then    'Links
                            If Not ilGetDescriptionOnly Then
                                If StrComp("tmLink", Left$(slLine, 6), 1) = 0 Then
                                    ilUpper = UBound(tmLink)
                                    ilRet = gParseItem(slLine, 2, "|", slStr)
                                    tmLink(ilUpper).iKeyNo = Val(slStr)
                                    ilRet = gParseItem(slLine, 3, "|", slStr)
                                    tmLink(ilUpper).iToFileID = Val(slStr)
                                    ilRet = gParseItem(slLine, 4, "|", slStr)
                                    tmLink(ilUpper).iToFieldID = Val(slStr)
                                    ilRet = gParseItem(slLine, 5, "|", slStr)
                                    tmLink(ilUpper).iKeyOffset = Val(slStr)
                                    ilRet = gParseItem(slLine, 6, "|", slStr)
                                    tmLink(ilUpper).iFromFileID = Val(slStr)
                                    ilRet = gParseItem(slLine, 7, "|", slStr)
                                    tmLink(ilUpper).iFromFieldID = Val(slStr)
                                    ilRet = gParseItem(slLine, 8, "|", slStr)
                                    tmLink(ilUpper).iFromOffset = Val(slStr)
                                    ReDim Preserve tmLink(0 To ilUpper + 1) As LINK
                                Else
                                    lbcLink.AddItem slLine
                                End If
                            End If
                        ElseIf ilImportArea = 5 Then    'Selections
                            If Not ilGetDescriptionOnly Then
                                If Len(slSelection) > 0 Then
                                    slSelection = slSelection & Chr(13) & Chr(10) & slLine
                                Else
                                    slSelection = slLine
                                End If
                            End If
                        ElseIf ilImportArea = 6 Then    'Task
                            If Not ilGetDescriptionOnly Then
                                If slLine = rbcEditReplace(0).Caption Then
                                    rbcEditReplace(0).Value = True
                                ElseIf slLine = rbcEditReplace(1).Caption Then
                                    rbcEditReplace(1).Value = True
                                ElseIf slLine = rbcEditReplace(2).Caption Then
                                    rbcEditReplace(2).Value = True
                                    imPasswordAsked = True
                                ElseIf slLine = rbcEditReplace(3).Caption Then
                                    rbcEditReplace(3).Value = True
                                    imPasswordAsked = True
                                End If
                            End If
                        ElseIf ilImportArea = 7 Then    'Fields
                            If Not ilGetDescriptionOnly Then
                                If ilFirstField Then
                                    lbcEditReplace(2).Clear
                                    ilFirstField = False
                                End If
                                lbcEditReplace(2).AddItem slLine
                                pbcLbcEditReplace2_Paint
                                ilPos = InStr(slLine, "|")  'End Of File Name
                                If ilPos > 0 Then
                                    ilPos = InStr(ilPos + 1, slLine, "|")   'End Of Field Name
                                    If ilPos > 0 Then
                                        ilPos = InStr(ilPos + 1, slLine, "|")   'End of Type
                                        If ilPos > 0 Then
                                            slText = Left$(slLine, ilPos - 1)   'Get name without conversion type
                                            For ilLoop = 0 To lbcEditReplace(0).ListCount - 1 Step 1
                                                If slText = lbcEditReplace(0).List(ilLoop) Then 'Add conversion type
                                                    ilPos = InStr(ilPos + 1, slLine, "|")   'Remove replace value
                                                    If ilPos > 0 Then
                                                        slText = Left$(slLine, ilPos - 1)
                                                        lbcEditReplace(0).List(ilLoop) = slText
                                                    Else
                                                        lbcEditReplace(0).List(ilLoop) = slLine
                                                    End If
                                                    Exit For
                                                End If
                                            Next ilLoop
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf ilImportArea = 8 Then    'Key
                            If Not ilGetDescriptionOnly Then
                                ilRet = gParseItem(slLine, 1, "|", slStr)
                                edcSearchKey.Text = Val(slStr)
                                ilRet = gParseItem(slLine, 2, "|", slStr)
                                edcSearchFrom.Text = slStr
                                ilRet = gParseItem(slLine, 3, "|", slStr)
                                edcSearchTo.Text = slStr
                            End If
                        ElseIf ilImportArea = 9 Then    'Run
                            If Not ilGetDescriptionOnly Then
                                If slLine = rbcRun(0).Caption Then
                                    rbcRun(0).Value = True
                                ElseIf slLine = rbcRun(1).Caption Then
                                    rbcRun(1).Value = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop Until ilEof
    If Not ilGetDescriptionOnly Then
        lbcEditReplace(2).AddItem " "
        mPopFieldBoxes
    End If
    pbcLbcEditReplace2_Paint
    Exit Sub
mReadSpecErr:
    ilRet = Err
    Resume Next
    
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRecordSelected                 *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if record selected    *
'*                                                     *
'*******************************************************
Private Function mRecordSelected() As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilDataType As Integer
    Dim slLogic As String
    Dim slTempLogic As String
    Dim slChar As String * 1
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim slValue As String
    Dim ilDone As Integer
    Dim ilPass As Integer   'Pass zero- remove And; Pass 1 remove Or
    Dim ilLevel As Integer
    Dim ilMaxLevel As Integer
    Dim slTest As String
    Dim ilPartRemoved As Integer
    If UBound(tmRecSel) <= 1 Then
        mRecordSelected = True
        Exit Function
    End If
    For ilLoop = LBound(tmRecSel) To UBound(tmRecSel) - 1 Step 1
        If tmRecSel(ilLoop).iFieldIndex1 >= 0 Then
            If tmDDFField(tmRecSel(ilLoop).iFieldIndex1).iFileID = imFileID Then  'Key file
                mGetField tmDDFField(tmRecSel(ilLoop).iFieldIndex1), 0
            Else
                mGetFieldByKey tmDDFField(tmRecSel(ilLoop).iFieldIndex1)
            End If
            tmRecSel(ilLoop).sValue1 = Trim$(smFieldValue)
            ilDataType = Asc(tmDDFField(tmRecSel(ilLoop).iFieldIndex1).sDataType)
        End If
        If tmRecSel(ilLoop).iFieldIndex2 >= 0 Then
            If tmDDFField(tmRecSel(ilLoop).iFieldIndex2).iFileID = imFileID Then  'Key file
                mGetField tmDDFField(tmRecSel(ilLoop).iFieldIndex2), 0
            Else
                mGetFieldByKey tmDDFField(tmRecSel(ilLoop).iFieldIndex2)
            End If
            tmRecSel(ilLoop).sValue2 = Trim$(smFieldValue)
            ilDataType = Asc(tmDDFField(tmRecSel(ilLoop).iFieldIndex2).sDataType)
        End If
        Select Case ilDataType
            Case 0  'String
                ilRet = StrComp(tmRecSel(ilLoop).sValue1, tmRecSel(ilLoop).sValue2, 1)
            Case 1  'Integer
                If Val(tmRecSel(ilLoop).sValue1) = Val(tmRecSel(ilLoop).sValue2) Then
                    ilRet = 0
                ElseIf Val(tmRecSel(ilLoop).sValue1) < Val(tmRecSel(ilLoop).sValue2) Then
                    ilRet = -1
                Else
                    ilRet = 1
                End If
            Case 2  'IEEE Float
            Case 3  'Btrieve Date
                If gDateValue(tmRecSel(ilLoop).sValue1) = gDateValue(tmRecSel(ilLoop).sValue2) Then
                    ilRet = 0
                ElseIf gDateValue(tmRecSel(ilLoop).sValue1) < gDateValue(tmRecSel(ilLoop).sValue2) Then
                    ilRet = -1
                Else
                    ilRet = 1
                End If
            Case 4  'Btrieve Time
                If gTimeToCurrency(tmRecSel(ilLoop).sValue1, False) = gTimeToCurrency(tmRecSel(ilLoop).sValue2, False) Then
                    ilRet = 0
                ElseIf gTimeToCurrency(tmRecSel(ilLoop).sValue1, False) < gTimeToCurrency(tmRecSel(ilLoop).sValue2, False) Then
                    ilRet = -1
                Else
                    ilRet = 1
                End If
            Case 5  'COBOL Decimal COMP-3
                ilRet = gCompNumberStr(tmRecSel(ilLoop).sValue1, tmRecSel(ilLoop).sValue2)
            Case 6  'COBOL Money
                ilRet = gCompNumberStr(tmRecSel(ilLoop).sValue1, tmRecSel(ilLoop).sValue2)
            Case 7  'Logical
                If Val(tmRecSel(ilLoop).sValue1) = Val(tmRecSel(ilLoop).sValue2) Then
                    ilRet = 0
                ElseIf Val(tmRecSel(ilLoop).sValue1) < Val(tmRecSel(ilLoop).sValue2) Then
                    ilRet = -1
                Else
                    ilRet = 1
                End If
            Case 8  'COBOL Numeric
            Case 9  'BASIC Float
            Case 10 'Pascal LString
            Case 11 'C ZString
                ilRet = StrComp(Trim$(tmRecSel(ilLoop).sValue1), Trim$(tmRecSel(ilLoop).sValue2), 1)
            Case 12 'Variable Length Note
            Case 13 'LVar
                ilRet = StrComp(Trim$(tmRecSel(ilLoop).sValue1), Trim$(tmRecSel(ilLoop).sValue2), 1)
            Case 14 'Unsigned Binary
            Case 15 'AutoIncrement
                If Val(tmRecSel(ilLoop).sValue1) = Val(tmRecSel(ilLoop).sValue2) Then
                    ilRet = 0
                ElseIf Val(tmRecSel(ilLoop).sValue1) < Val(tmRecSel(ilLoop).sValue2) Then
                    ilRet = -1
                Else
                    ilRet = 1
                End If
            Case 16 'Bit
            Case 17 'COBOL Numeric STS
            Case 100  'Length
                If gLengthToCurrency(tmRecSel(ilLoop).sValue1) = gLengthToCurrency(tmRecSel(ilLoop).sValue2) Then
                    ilRet = 0
                ElseIf gLengthToCurrency(tmRecSel(ilLoop).sValue1) < gLengthToCurrency(tmRecSel(ilLoop).sValue2) Then
                    ilRet = -1
                Else
                    ilRet = 1
                End If
        End Select
        tmRecSel(ilLoop).iResult = False
        Select Case tmRecSel(ilLoop).iOper
            Case 0  '=
                If ilRet = 0 Then
                    tmRecSel(ilLoop).iResult = True
                End If
            Case 1  '<>
                If ilRet <> 0 Then
                    tmRecSel(ilLoop).iResult = True
                End If
            Case 2  '<
                If ilRet < 0 Then
                    tmRecSel(ilLoop).iResult = True
                End If
            Case 3  '>
                If ilRet > 0 Then
                    tmRecSel(ilLoop).iResult = True
                End If
            Case 4  '<=
                If ilRet <= 0 Then
                    tmRecSel(ilLoop).iResult = True
                End If
            Case 5  '>=
                If ilRet >= 0 Then
                    tmRecSel(ilLoop).iResult = True
                End If
        End Select
    Next ilLoop
    'Replace Index in smLogicSel with T or F
    ilIndex = 1
    slLogic = ""
    Do While ilIndex <= Len(smLogicSel)
        slChar = Mid(smLogicSel, ilIndex, 1)
        If slChar = "#" Then
            ilIndex = ilIndex + 1
            slValue = ""
            Do
                slChar = Mid(smLogicSel, ilIndex, 1)
                If (Asc(slChar) >= Asc("0")) And (Asc(slChar) <= Asc("9")) Then
                    slValue = slValue & slChar
                    ilIndex = ilIndex + 1
                    If ilIndex > Len(smLogicSel) Then
                        ilValue = Val(slValue)
                        If tmRecSel(ilValue).iResult = True Then
                            slLogic = slLogic & "T"
                        Else
                            slLogic = slLogic & "F"
                        End If
                        Exit Do
                    End If
                Else
                    ilValue = Val(slValue)
                    If tmRecSel(ilValue).iResult = True Then
                        slLogic = slLogic & "T"
                    Else
                        slLogic = slLogic & "F"
                    End If
                    ilIndex = ilIndex - 1
                    Exit Do
                End If
            Loop While ilIndex <= Len(smLogicSel)
        Else
            slLogic = slLogic & slChar
        End If
        ilIndex = ilIndex + 1
    Loop
    ilDone = False
    Do
        For ilPass = 0 To 1 Step 1
            'Remove extra parentheses=> ((T)) change to (T); ((F)) change to (F)
            Do
                ilIndex = 1
                slTempLogic = ""
                ilPartRemoved = False
                Do While ilIndex <= Len(slLogic)
                    slChar = Mid(slLogic, ilIndex, 1)
                    If slChar = "(" Then
                        If Mid(slLogic, ilIndex, 3) = "(T)" Then
                            slTempLogic = slTempLogic & "T"
                            ilIndex = ilIndex + 3
                            ilPartRemoved = True
                        ElseIf Mid(slLogic, ilIndex, 3) = "(F)" Then
                            slTempLogic = slTempLogic & "F"
                            ilIndex = ilIndex + 3
                            ilPartRemoved = True
                        Else
                            slTempLogic = slTempLogic & slChar
                            ilIndex = ilIndex + 1
                        End If
                    Else
                        slTempLogic = slTempLogic & slChar
                        ilIndex = ilIndex + 1
                    End If
                Loop
                slLogic = slTempLogic
            Loop While ilPartRemoved
            If Len(slLogic) = 1 Then
                If slLogic = "T" Then
                    mRecordSelected = True
                Else
                    mRecordSelected = False
                End If
                Exit Function
            ElseIf slLogic = "(T)" Then
                mRecordSelected = True
                Exit Function
            ElseIf slLogic = "(F)" Then
                mRecordSelected = True
                Exit Function
            End If
            'Determine max Parentheses level
            ilLevel = 0
            ilMaxLevel = 0
            ilIndex = 1
            Do While ilIndex <= Len(slLogic)
                slChar = Mid(slLogic, ilIndex, 1)
                If slChar = "(" Then
                    ilLevel = ilLevel + 1
                    If ilLevel > ilMaxLevel Then
                        ilMaxLevel = ilLevel
                    End If
                ElseIf slChar = ")" Then
                    ilLevel = ilLevel - 1
                End If
                ilIndex = ilIndex + 1
            Loop
            'Perform Logic operation for level
            ilIndex = 1
            slTempLogic = ""
            ilLevel = 0
            Do While ilIndex <= Len(slLogic)
                slChar = Mid(slLogic, ilIndex, 1)
                If (slChar <> "(") And (slChar <> ")") Then
                    If ilLevel = ilMaxLevel Then
                        If ilPass = 0 Then
                            slTest = Mid(slLogic, ilIndex, 5)
                            If slTest = "TAndT" Then
                                slTempLogic = slTempLogic & "T"
                                ilIndex = ilIndex + 4
                            ElseIf slTest = "TAndF" Then
                                slTempLogic = slTempLogic & "F"
                                ilIndex = ilIndex + 4
                            ElseIf slTest = "FAndT" Then
                                slTempLogic = slTempLogic & "F"
                                ilIndex = ilIndex + 4
                            ElseIf slTest = "FAndF" Then
                                slTempLogic = slTempLogic & "F"
                                ilIndex = ilIndex + 4
                            Else
                                slTempLogic = slTempLogic & slChar
                            End If
                        Else
                            slTest = Mid(slLogic, ilIndex, 4)
                            If slTest = "TOrT" Then
                                slTempLogic = slTempLogic & "T"
                                ilIndex = ilIndex + 3
                            ElseIf slTest = "TOrF" Then
                                slTempLogic = slTempLogic & "T"
                                ilIndex = ilIndex + 3
                            ElseIf slTest = "FOrT" Then
                                slTempLogic = slTempLogic & "T"
                                ilIndex = ilIndex + 3
                            ElseIf slTest = "FOrF" Then
                                slTempLogic = slTempLogic & "F"
                                ilIndex = ilIndex + 3
                            Else
                                slTempLogic = slTempLogic & slChar
                            End If
                        End If
                    Else
                        slTempLogic = slTempLogic & slChar
                    End If
                Else
                    slTempLogic = slTempLogic & slChar
                    If slChar = "(" Then
                        ilLevel = ilLevel + 1
                    Else
                        ilLevel = ilLevel - 1
                    End If
                End If
                ilIndex = ilIndex + 1
            Loop
            slLogic = slTempLogic
        Next ilPass
    Loop While Not ilDone
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveGrid                       *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Save information within gird to *
'*                     file                            *
'*                                                     *
'*******************************************************
Private Sub mSaveGrid(ilSaveOrPrint As Integer)
'
'   ilSaveOrPrint: 0=Save; 1=Print
'
    Dim slRecord As String
    Dim ilRet As Integer
    Dim llRecNo As Long
    Dim llStartRecNo As Long
    Dim llEndRecNo As Long
    Dim ilRecIndex1 As Integer
    Dim ilRecIndex2 As Integer
    Dim llRec As Long
    Dim ilLoop As Integer
    Dim ilField As Integer
    Dim slSvDataType As String
    Dim slFieldDataType As String
    Dim slFieldInfo As String
    Dim slFieldName As String
    Dim ilSize As Integer
    Dim ilCol As Integer
    Dim slTitle As String
    Dim ilColMax As Integer
    Dim ilCMax As Integer
    Dim ilPos As Integer
    Dim slValue As String
    Dim llValue As Long
    Dim ilLinesPerPage As Integer
    Dim ilCurrentLineNo As Integer
    Dim ilPageNo As Integer
    Dim slGen As String
    Dim slFileName As String
    Dim ilFileID As Integer
    Dim ilHLoop As Integer
    Dim slHFileName As String
    Dim slHFieldDataType As String
    Dim slHFieldInfo As String
    Dim slHFieldName As String
    Dim slHSvDataType As String
    Dim ilHFileID As Integer
    Dim ilHField As Integer
    Dim ilHSize As Integer
    Dim ilHCol As Integer
    Dim ilRank As Integer
    Dim ilPriceLevel As Integer
    
    llStartRecNo = Val(edcRecordNo(0).Text)
    llEndRecNo = Val(edcRecordNo(1).Text)
    ilColMax = Val(edcWrap.Text)
    llRec = llStartRecNo
    ilRecIndex1 = 1
    ilRecIndex2 = 1
    Do While llRec > 32000
        ilRecIndex1 = ilRecIndex1 + 1
        llRec = llRec - 32000
    Loop
    slGen = "Generated: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    ilRecIndex2 = llRec
    ilRet = 0
    If ilSaveOrPrint = 1 Then
        '1440 twip per inch-- 1/2 margin at top and bottom
        ilLinesPerPage = (Printer.Height - 1440) / Printer.TextHeight("TEST") - 1
    End If
    ilCurrentLineNo = 0
    ilPageNo = 0
    On Error GoTo mSaveGridErr:
    'Output heading line
    GoSub mHeading1
    If cbcInclude(1).Value = vbChecked Then 'Report information
        smFieldValue = Trim$(edcDescription.Text)
        smHeading = ""
        If Len(smFieldValue) > 0 Then
            smHeading = "Description: "
            Do While Len(smFieldValue) > 0
                ilPos = InStr(smFieldValue, " ")
                If ilPos > 0 Then
                    If Len(smHeading) + ilPos - 1 > 80 Then
                        If ilSaveOrPrint = 1 Then
                            Printer.Print smHeading
                            ilCurrentLineNo = ilCurrentLineNo + 1
                        Else
                            Print #hmTo, smHeading
                        End If
                        smHeading = ""
                    End If
                    smHeading = smHeading & Left$(smFieldValue, ilPos)
                    smFieldValue = Mid(smFieldValue, ilPos + 1)
                Else
                    smHeading = smHeading & smFieldValue
                    If ilSaveOrPrint = 1 Then
                        Printer.Print smHeading
                        ilCurrentLineNo = ilCurrentLineNo + 1
                    Else
                        Print #hmTo, smHeading
                    End If
                    smFieldValue = ""
                    Exit Do
                End If
            Loop
        End If
        smHeading = ""
        If ilSaveOrPrint = 1 Then
            Printer.Print smHeading
            ilCurrentLineNo = ilCurrentLineNo + 1
        Else
            Print #hmTo, smHeading
        End If
        DoEvents
        If imTerminate Then
            If ilSaveOrPrint = 1 Then
                On Error Resume Next
                ilRet = AbortDoc(Printer.hDC)
                Printer.EndDoc
            End If
            Exit Sub
        End If
    End If
    'Get fields name
    GoSub mHeading2
    DoEvents
    If imTerminate Then
        If ilSaveOrPrint = 1 Then
            On Error Resume Next
            ilRet = AbortDoc(Printer.hDC)
            Printer.EndDoc
        End If
        Exit Sub
    End If
    For llRecNo = llStartRecNo To llEndRecNo Step 1
        If ((ilRecIndex1 = imRecMaxUpper1) And (ilRecIndex2 < imRecMaxUpper2)) Or ((ilRecIndex1 < imRecMaxUpper1) And (ilRecIndex2 <= 32000)) Then
            lmActualRecLen = imBaseRecLen
            ilRet = btrGetDirect(hmFile, tmRec, lmActualRecLen, tmRecInfo(ilRecIndex1, ilRecIndex2).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            DoEvents
            If imTerminate Then
                If ilSaveOrPrint = 1 Then
                    On Error Resume Next
                    ilRet = AbortDoc(Printer.hDC)
                    Printer.EndDoc
                End If
                Exit Sub
            End If
            'Get fields
            slRecord = ""
            ilCol = 1
            For ilLoop = 0 To lbcEditReplace(2).ListCount - 2 Step 1
                slFieldInfo = lbcEditReplace(2).List(ilLoop)
                ilRet = gParseItem(slFieldInfo, 1, "|", slFileName)
                ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
                ilRet = gParseItem(slFieldInfo, 4, "|", slFieldDataType)
                If ilRet <> CP_MSG_NONE Then
                    slFieldDataType = ""
                Else
                    If StrComp(slFieldDataType, "Length of Time", 1) = 0 Then
                        slFieldDataType = Chr(100)
                    End If
                    If StrComp(slFieldDataType, "Hex Field Dump", 1) = 0 Then
                        slFieldDataType = Chr(101)
                    End If
                End If
                ilFileID = mGetFileID(slFileName)
                For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                    If (StrComp(Trim$(slFieldName), Trim$(tmDDFField(ilField).sName), 1) = 0) And (ilFileID = tmDDFField(ilField).iFileID) Then
                        slSvDataType = tmDDFField(ilField).sDataType
                        If slFieldDataType <> "" Then
                            tmDDFField(ilField).sDataType = slFieldDataType
                        End If
                        If tmDDFField(ilField).iFileID = imFileID Then  'Key file
                            mGetField tmDDFField(ilField), 0
                        Else
                            mGetFieldByKey tmDDFField(ilField)
                        End If
                        If (Asc(tmDDFField(ilField).sDataType) = 12) Or (Asc(tmDDFField(ilField).sDataType) = 13) Then
                            ilSize = Len(smFieldValue)
                        ElseIf (Asc(tmDDFField(ilField).sDataType) = 101) Then
                            ilSize = Len(smFieldValue)
                        Else
                            ilSize = mFieldSize(tmDDFField(ilField))
                        End If
                        gdcData.Row = 0
                        gdcData.Col = ilCol
                        slTitle = gdcData.Text
                        If Len(slTitle) > ilSize Then
                            ilSize = Len(slTitle)
                        End If
                        Do While Len(smFieldValue) < ilSize
                            smFieldValue = smFieldValue & " "
                        Loop
                        If (Asc(tmDDFField(ilField).sDataType) = 101) Or (Asc(tmDDFField(ilField).sDataType) = 12) Or (Asc(tmDDFField(ilField).sDataType) = 13) Then
                            If ilColMax > 0 Then
                                ilCMax = ilColMax
                            Else
                                ilCMax = 80
                            End If
                            If Len(slRecord & smFieldValue & " ") > ilCMax Then
                                GoSub mLineOutput
                                slRecord = "     "
                                Do While Len(smFieldValue) > 0
                                    ilPos = InStr(smFieldValue, " ")
                                    If ilPos > 0 Then
                                        If Len(slRecord) + ilPos - 1 > ilCMax Then
                                            GoSub mLineOutput
                                            slRecord = "     "
                                        End If
                                        slRecord = slRecord & Left$(smFieldValue, ilPos)
                                        smFieldValue = Mid(smFieldValue, ilPos + 1)
                                    Else
                                        Exit Do
                                    End If
                                Loop
                            End If
                        Else
                            If (ilColMax > 0) And (Len(slRecord & smFieldValue & " ") > ilColMax) Then
                                GoSub mLineOutput
                                slRecord = "     "
                            End If
                        End If
                        slRecord = slRecord & smFieldValue & " "
                        tmDDFField(ilField).sDataType = slSvDataType
                        ilCol = ilCol + 1
                        Exit For
                    End If
                Next ilField
            Next ilLoop
            GoSub mLineOutput
            DoEvents
            If imTerminate Then
                If ilSaveOrPrint = 1 Then
                    On Error Resume Next
                    ilRet = AbortDoc(Printer.hDC)
                    Printer.EndDoc
                End If
                Exit Sub
            End If
            If ((Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary")) And (rbcPAS(0).Value) Then
                lmActualRecLen = imBaseRecLen
                ilRet = gSSFGetDirect(hmFile, tmSsf, lmActualRecLen, tmRecInfo(ilRecIndex1, ilRecIndex2).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                DoEvents
                If imTerminate Then
                    If ilSaveOrPrint = 1 Then
                        On Error Resume Next
                        ilRet = AbortDoc(Printer.hDC)
                        Printer.EndDoc
                    End If
                    Exit Sub
                End If
                slRecord = " "
                GoSub mLineOutput
                slRecord = "Type   Start Time End Time   mnfExcl mnfExcl ltfCode lvfCode"
                GoSub mLineOutput
                slRecord = "       Time       ltfCode    Units   Length  anfCode # Spots Sustaining Sponsorship Avail  Spot"
                GoSub mLineOutput
                slRecord = "       sdfCode    Line Enter Hour >= Hour <= mnfComp mnfComp Length     adfCode     Rank"
                GoSub mLineOutput
                For ilLoop = 1 To tmSsf.iCount Step 1
                    DoEvents
                    If imTerminate Then
                        If ilSaveOrPrint = 1 Then
                            On Error Resume Next
                            ilRet = AbortDoc(Printer.hDC)
                            Printer.EndDoc
                        End If
                        Exit Sub
                    End If
                   LSet tmProg = tmSsf.tPAS(ilLoop)
                    If tmProg.iRecType = 1 Then    'Program
                        slRecord = "Prog   "
                        gUnpackTime tmProg.iStartTime(0), tmProg.iStartTime(1), "A", "1", slValue
                        Do While Len(slValue) < 11
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        gUnpackTime tmProg.iEndTime(0), tmProg.iEndTime(1), "A", "1", slValue
                        Do While Len(slValue) < 11
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmProg.iMnfExcl(0)))
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmProg.iMnfExcl(1)))
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmProg.iLtfCode))
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmProg.lLvfCode))
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        GoSub mLineOutput
                    ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then 'Avail
                       LSet tmAvail = tmSsf.tPAS(ilLoop)
                        slRecord = "Avail  "
                        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slValue
                        Do While Len(slValue) < 11
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmAvail.iLtfCode))
                        Do While Len(slValue) < 11
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmAvail.iAvInfo And &H1F))
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmAvail.iLen))
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmAvail.ianfCode))
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmAvail.iNoSpotsThis))
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        If (tmAvail.iAvInfo And SSSUSTAINING) = SSSUSTAINING Then
                            slRecord = slRecord & "Allowed    "
                        Else
                            slRecord = slRecord & "Disallowed "
                        End If
                        If (tmAvail.iAvInfo And SSSPONSORSHIP) = SSSPONSORSHIP Then
                            slRecord = slRecord & "Allowed     "
                        Else
                            slRecord = slRecord & "Disallowed  "
                        End If
                        If (tmAvail.iAvInfo And SSLOCK) = SSLOCK Then
                            slRecord = slRecord & "Lock   "
                        Else
                            slRecord = slRecord & "Unlock "
                        End If
                        If (tmAvail.iAvInfo And SSLOCKSPOT) = SSLOCKSPOT Then
                            slRecord = slRecord & "Lock   "
                        Else
                            slRecord = slRecord & "Unlock "
                        End If
                        GoSub mLineOutput
                    ElseIf (tmProg.iRecType And &HF) = 10 Then    'Spot
                       LSet tmSpot = tmSsf.tPAS(ilLoop)
                        slRecord = "Spot   "
                        slValue = Trim$(Str$(tmSpot.lSdfCode))
                        Do While Len(slValue) < 11
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        llValue = tmSpot.lBkInfo And &H1FFFF    'Book Data
                        slValue = Format$(llValue, "m/d/yy")
                        Do While Len(slValue) < 11
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        llValue = (tmSpot.lBkInfo And &H7E0000) \ SHIFT17  'Start Hour in Minutes
                        slValue = Trim$(Str$(llValue)) & "Min"
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        llValue = (tmSpot.lBkInfo And &H1F800000) \ SHIFT23  'Start Hour in Minutes
                        slValue = Trim$(Str$(llValue)) & "Min"
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmSpot.iMnfComp(0)))
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmSpot.iMnfComp(1)))
                        Do While Len(slValue) < 8
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmSpot.iPosLen And &HFFF))
                        Do While Len(slValue) < 11
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        slValue = Trim$(Str$(tmSpot.iAdfCode))
                        Do While Len(slValue) < 12
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        ilRank = tmSpot.iRank And RANKMASK
                        ilPriceLevel = (tmSpot.iRank And PRICELEVELMASK) / SHIFT11
                        slValue = Trim$(Str$(ilPriceLevel)) & "-" & Trim$(Str$(ilRank))
                        Do While Len(slValue) < 7
                            slValue = slValue & " "
                        Loop
                        slRecord = slRecord & slValue
                        GoSub mLineOutput
                    ElseIf (tmProg.iRecType And &HF) = 11 Then    'BBSpot
                        slRecord = "BBSpot"
                        GoSub mLineOutput
                    End If
                Next ilLoop
                'GoSub mLineOutput
            End If
            ilRecIndex2 = ilRecIndex2 + 1
            If ilRecIndex2 > 32000 Then
                ilRecIndex1 = ilRecIndex1 + 1
                ilRecIndex2 = 1
            End If
        End If
    Next llRecNo
    'Output Specifications
    If cbcInclude(0).Value = vbChecked Then 'Specifications information
        Printer.NewPage
        GoSub mHeading1
        smHeading = ""
        If ilSaveOrPrint = 1 Then
            Printer.Print smHeading
            ilCurrentLineNo = ilCurrentLineNo + 1
        Else
            Print #hmTo, smHeading
        End If
        smHeading = "File: " & lbcFile.List(lbcFile.ListIndex)
        If ilSaveOrPrint = 1 Then
            Printer.Print smHeading
        Else
            Print #hmTo, smHeading
        End If
        smHeading = "Selection:"
        If ilSaveOrPrint = 1 Then
            Printer.Print smHeading
        Else
            Print #hmTo, smHeading
        End If
        smHeading = edcSelection.Text
        If ilSaveOrPrint = 1 Then
            Printer.Print smHeading
        Else
            Print #hmTo, smHeading
        End If
        smHeading = "Key: " & Trim$(edcSearchKey.Text) & " " & Trim$(lacSearchFrom.Caption) & " " & Trim$(edcSearchFrom.Text) & " To " & Trim$(edcSearchTo.Text)
        If ilSaveOrPrint = 1 Then
            Printer.Print smHeading
        Else
            Print #hmTo, smHeading
        End If
    End If
    If ilSaveOrPrint = 1 Then
        Printer.EndDoc
    End If
    On Error GoTo 0
    Exit Sub
mSaveGridErr:
    ilRet = Err
    Resume Next
mHeading1:  'Output file name and date
    If cbcInclude(1).Value = vbChecked Then 'Report information
        smHeading = "File: " & lbcFile.List(lbcFile.ListIndex)
        If ilSaveOrPrint = 1 Then
            Printer.Print smHeading
            ilCurrentLineNo = ilCurrentLineNo + 1
            Printer.Print "Database Path: " & smDBPath
            ilCurrentLineNo = ilCurrentLineNo + 1
        Else
            Print #hmTo, smHeading
        End If
        smHeading = slGen
        If ilSaveOrPrint = 1 Then
            Do While Len(smHeading) < ilColMax - 10
                smHeading = smHeading & " "
            Loop
            ilPageNo = ilPageNo + 1
            smHeading = smHeading & "Page" & Str$(ilPageNo)
            Printer.Print smHeading
            ilCurrentLineNo = ilCurrentLineNo + 1
        Else
            Print #hmTo, smHeading
        End If
    Else
        If ilSaveOrPrint = 1 Then
            smHeading = slGen
            Do While Len(smHeading) < ilColMax - 10
                smHeading = smHeading & " "
            Loop
            ilPageNo = ilPageNo + 1
            smHeading = smHeading & "Page" & Str$(ilPageNo)
            Printer.Print smHeading
            ilCurrentLineNo = ilCurrentLineNo + 1
        End If
    End If
    Return
mHeading2:  'Output Column headings
    ilHCol = 1
    smHeading = ""
    For ilHLoop = 0 To lbcEditReplace(2).ListCount - 2 Step 1
        slHFieldInfo = lbcEditReplace(2).List(ilHLoop)
        ilRet = gParseItem(slHFieldInfo, 1, "|", slHFileName)
        ilRet = gParseItem(slHFieldInfo, 2, "|", slHFieldName)
        ilRet = gParseItem(slHFieldInfo, 4, "|", slHFieldDataType)
        If ilRet <> CP_MSG_NONE Then
            slHFieldDataType = ""
        Else
            If StrComp(slHFieldDataType, "Length of Time", 1) = 0 Then
                slHFieldDataType = Chr(100)
            End If
            If StrComp(slHFieldDataType, "Hex Field Dump", 1) = 0 Then
                slHFieldDataType = Chr(101)
            End If
        End If
        ilHFileID = mGetFileID(slHFileName)
        For ilHField = 0 To UBound(tmDDFField) - 1 Step 1
            If (StrComp(Trim$(slHFieldName), Trim$(tmDDFField(ilHField).sName), 1) = 0) And (ilHFileID = tmDDFField(ilHField).iFileID) Then
                slHSvDataType = tmDDFField(ilHField).sDataType
                If slHFieldDataType <> "" Then
                    tmDDFField(ilHField).sDataType = slHFieldDataType
                End If
                ilHSize = mFieldSize(tmDDFField(ilHField))
                gdcData.Row = 0
                gdcData.Col = ilHCol
                slTitle = gdcData.Text
                If (Asc(tmDDFField(ilHField).sDataType) = 101) Or (Asc(tmDDFField(ilHField).sDataType) = 12) Or (Asc(tmDDFField(ilHField).sDataType) = 13) Then
                    ilHSize = Len(slTitle)
                End If
                Do While Len(slTitle) < ilHSize
                    slTitle = slTitle & " "
                Loop
                If (ilColMax > 0) And (Len(smHeading & slTitle & " ") > ilColMax) Then
                    If ilSaveOrPrint = 1 Then
                        Printer.Print smHeading
                        ilCurrentLineNo = ilCurrentLineNo + 1
                    Else
                        Print #hmTo, smHeading
                    End If
                    smHeading = "     "
                End If
                smHeading = smHeading & slTitle & " "
                tmDDFField(ilHField).sDataType = slHSvDataType
                ilHCol = ilHCol + 1
                Exit For
            End If
        Next ilHField
    Next ilHLoop
    If ilSaveOrPrint = 1 Then
        Printer.Print smHeading
        ilCurrentLineNo = ilCurrentLineNo + 1
    Else
        Print #hmTo, smHeading
    End If
    Return
mLineOutput:
    If ilSaveOrPrint = 1 Then
        If ilCurrentLineNo >= ilLinesPerPage Then
            Printer.NewPage
            ilCurrentLineNo = 0
            GoSub mHeading1
            GoSub mHeading2
        End If
        Printer.Print slRecord
        ilCurrentLineNo = ilCurrentLineNo + 1
    Else
        Print #hmTo, slRecord
    End If
    Return
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveGrid                       *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Save information within gird to *
'*                     file                            *
'*                                                     *
'*******************************************************
Private Sub mSaveSpec()
    Dim slLine As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilLoop As Integer
    
    ilRet = 0
    On Error GoTo mSaveSpecErr:
    slLine = "[Generated]"
    Print #hmTo, slLine
    slLine = Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmTo, slLine
    slLine = "[Description]"
    Print #hmTo, slLine
    slStr = Trim$(edcDescription.Text)
    If Len(slStr) > 0 Then
        slLine = ""
        Do While Len(slStr) > 0
            ilPos = InStr(slStr, " ")
            If ilPos > 0 Then
                If Len(slLine) + ilPos - 1 > 80 Then
                    Print #hmTo, slLine
                    slLine = ""
                End If
                slLine = slLine & Left$(slStr, ilPos)
                slStr = Mid(slStr, ilPos + 1)
            Else
                slLine = slLine & slStr
                Print #hmTo, slLine
                slStr = ""
                Exit Do
            End If
        Loop
    End If
    'slLine = "[DDF Path]"
    'Print #hmTo, slLine
    'slLine = lbcPath.Path
    'Print #hmTo, slLine
    slLine = "[File]"
    Print #hmTo, slLine
    slLine = lbcFile.List(lbcFile.ListIndex)
    Print #hmTo, slLine
    'slLine = "[DATA Path]"
    'Print #hmTo, slLine
    'slLine = lbcDataPath.Path
    'Print #hmTo, slLine
    slLine = "[Links]"
    Print #hmTo, slLine
    For ilLoop = 0 To lbcLink.ListCount - 1 Step 1
        slLine = lbcLink.List(ilLoop)
        Print #hmTo, slLine
    Next ilLoop
    For ilLoop = 0 To UBound(tmLink) - 1 Step 1
        slLine = "tmLink|" & Trim$(Str$(tmLink(ilLoop).iKeyNo)) & "|" & Trim$(Str$(tmLink(ilLoop).iToFileID)) & "|" & Trim$(Str$(tmLink(ilLoop).iToFieldID)) & "|" & Trim$(Str$(tmLink(ilLoop).iKeyOffset)) & "|" & Trim$(Str$(tmLink(ilLoop).iFromFileID)) & "|" & Trim$(Str$(tmLink(ilLoop).iFromFieldID)) & "|" & Trim$(Str$(tmLink(ilLoop).iFromOffset)) & "|"
        Print #hmTo, slLine
    Next ilLoop
    slLine = "[Selections]"
    Print #hmTo, slLine
    slLine = Trim$(edcSelection.Text)
    If Len(slLine) > 0 Then
        Print #hmTo, slLine
    End If
    slLine = "[Task]"
    Print #hmTo, slLine
    If rbcEditReplace(1).Value Then
        slLine = rbcEditReplace(1).Caption
    ElseIf rbcEditReplace(2).Value Then
        slLine = rbcEditReplace(2).Caption
    ElseIf rbcEditReplace(3).Value Then
        slLine = rbcEditReplace(3).Caption
    Else
        slLine = rbcEditReplace(0).Caption
    End If
    Print #hmTo, slLine
    If rbcEditReplace(3).Value = False Then
        slLine = "[Fields]"
        Print #hmTo, slLine
        For ilLoop = 0 To lbcEditReplace(2).ListCount - 2 Step 1
            slLine = lbcEditReplace(2).List(ilLoop)
            Print #hmTo, slLine
        Next ilLoop
    End If
    
    slLine = "[Key]"
    Print #hmTo, slLine
    slLine = Trim$(edcSearchKey.Text) & "|" & Trim$(edcSearchFrom.Text) & "|" & Trim$(edcSearchTo.Text)
    Print #hmTo, slLine
    
    slLine = "[Run]"
    Print #hmTo, slLine
    If rbcRun(0).Value Then
        slLine = rbcRun(0).Caption
    Else
        slLine = rbcRun(1).Caption
    End If
    Print #hmTo, slLine
    On Error GoTo 0
    Return
mSaveSpecErr:
    ilRet = Err
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set Buttons                     *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
    If rbcEditReplace(0).Value Or rbcEditReplace(1).Value Then 'Show or Fix by Typing
        If lbcEditReplace(0).ListIndex >= 0 Then
            cmcEditReplace(0).Enabled = True
        Else
            cmcEditReplace(0).Enabled = False
        End If
        cmcEditReplace(1).Enabled = True
    ElseIf rbcEditReplace(2).Value Then 'Fix by Mass Substitution
        If (lbcEditReplace(0).ListIndex >= 0) And (edcReplaceValue.Text <> "") Then
            cmcEditReplace(0).Enabled = True
        Else
            cmcEditReplace(0).Enabled = False
        End If
        cmcEditReplace(1).Enabled = False
    End If
    If (lbcFile.ListIndex >= 0) And (lbcEditReplace(2).ListCount > 1) Then
        cmcProcess.Enabled = True
        mnuFileSave.Enabled = True
        mnuFileSaveAs.Enabled = True
    Else
        cmcProcess.Enabled = False
        mnuFileSave.Enabled = False
        mnuFileSaveAs.Enabled = False
    End If
    If lbcEditReplace(0).ListIndex >= 0 Then
        lbcEditReplace(1).Enabled = True
        lacEditReplace(2).Enabled = True
    Else
        lbcEditReplace(1).Enabled = False
        lacEditReplace(2).Enabled = False
    End If
    If (lbcEditReplace(2).ListIndex >= 0) And (lbcEditReplace(2).ListIndex <> (lbcEditReplace(2).ListCount - 1)) Then  'Remove
        cmcEditReplace(2).Enabled = True
    Else
        cmcEditReplace(2).Enabled = False
    End If
    If lbcEditReplace(2).ListCount > 1 Then     'Clear All
        cmcEditReplace(3).Enabled = True
    Else
        cmcEditReplace(3).Enabled = False
    End If
    If rbcEditReplace(3).Value Then  'Mass Removal
        cmcProcess.Enabled = True
        cmcEditReplace(0).Enabled = False
        cmcEditReplace(1).Enabled = False
        cmcEditReplace(2).Enabled = False
        cmcEditReplace(3).Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetLinkCommands                *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set Buttons                     *
'*                                                     *
'*******************************************************
Private Sub mSetLinkCommands()
    If (lbcToField.ListIndex >= 0) And (lbcFromField.ListIndex >= 0) Then 'Show or Fix by Typing
        cmcLink(1).Enabled = True
    Else
        cmcLink(1).Enabled = False
    End If
    If lbcLink.ListIndex >= 0 Then
        cmcLink(2).Enabled = True
    Else
        cmcLink(2).Enabled = False
    End If
    If lbcLink.ListCount > 0 Then
        cmcLink(3).Enabled = True
    Else
        cmcLink(3).Enabled = False
    End If
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
    Dim ilRet As Integer
    Dim ilLoop As Integer
    ilRet = btrClose(hmFile)
    btrDestroy hmFile
    For ilLoop = LBound(tmKeyFile) To UBound(tmKeyFile) - 1 Step 1
        ilRet = btrClose(tmKeyFile(ilLoop).hFile)
        btrDestroy tmKeyFile(ilLoop).hFile
    Next ilLoop
    Unload ShowFix
    'Set ShowFix = Nothing   'Remove data segment
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestOper                       *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if Replace field      *
'*                     should be checked for operator  *
'*                                                     *
'*******************************************************
Private Function mTestOper(ilDataType) As Integer
    mTestOper = False
    Select Case ilDataType
        Case 1  'Integer
            mTestOper = True
        Case 2  'IEEE Float
            mTestOper = True
        Case 5  'COBOL Decimal COMP-3
            mTestOper = True
        Case 6  'COBOL Money
            mTestOper = True
        Case 8  'COBOL Numeric
            mTestOper = True
        Case 9  'BASIC Float
            mTestOper = True
        Case 14 'Unsigned Binary
            mTestOper = True
        Case 15 'AutoIncrement
            mTestOper = True
        Case 17 'COBOL Numeric STS
            mTestOper = True
    End Select
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mUpdateField                    *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update record                   *
'*                                                     *
'*******************************************************
Private Sub mUpdateField()
    Dim llRec As Long
    Dim ilRecIndex1 As Integer
    Dim ilRecIndex2 As Integer
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim slFieldName As String
    Dim slFieldInfo As String
    Dim slFieldDataType As String
    Dim ilFileID As Integer
    Dim slFileName As String
    Dim ilField As Integer
    Dim ilField1 As Integer
    'Dim slSvDataType As String
    Dim ilSize As Integer
    Dim ilSvRow As Integer
    Dim ilSvCol As Integer
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim slStr As String
    Dim slStr1 As String
    Dim slInt2 As INT2
    Dim slInt4 As INT4
    Dim ilInsertRow As Integer
    If (edcValue.Visible) And (imEditCol > 0) And (imEditRow > 0) And (igPasswordOk) Then
        slStr = Trim$(edcValue.Text)
        If StrComp(Trim$(smFieldValue), slStr, 1) = 0 Then
            imEditRow = -1
            imEditCol = -1
            edcValue.Visible = False
            Exit Sub
        End If
        ilSvRow = gdcData.Row
        ilSvCol = gdcData.Col
        gdcData.Col = 0
        gdcData.Row = imEditRow
        slStr1 = Trim$(gdcData.Text)
        If (slStr1 = "") Or (slStr1 = "Deleted") Then
            ilInsertRow = True
        Else
            ilInsertRow = False
        End If
        'imEditCol = gdcData.SelStartCol
        'imEditRow = gdcData.SelStartRow
        llRec = imVBCAdjust * (vbcData.Value - 1) + imEditRow
        ilRecIndex1 = 1
        ilRecIndex2 = 1
        Do While llRec > 32000
            ilRecIndex1 = ilRecIndex1 + 1
            llRec = llRec - 32000
        Loop
        ilRecIndex2 = llRec
        If Not ilInsertRow Then
            lmActualRecLen = imBaseRecLen
            ilRet = btrGetDirect(hmFile, tmRec, lmActualRecLen, tmRecInfo(ilRecIndex1, ilRecIndex2).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                slFieldInfo = lbcEditReplace(2).List(imEditCol - 1)
                ilRet = gParseItem(slFieldInfo, 1, "|", slFileName)
                ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
                ilRet = gParseItem(slFieldInfo, 4, "|", slFieldDataType)
                If ilRet <> CP_MSG_NONE Then
                    slFieldDataType = ""
                    If (StrComp(Trim$(slFileName), "Mnf.Btr", 1) = 0) And (StrComp(Trim$(slFieldName), "mnfRPU", 1) = 0) Then
                        slFieldDataType = Chr(101)
                    End If
                    If (StrComp(Trim$(slFileName), "Mnf.Btr", 1) = 0) And (StrComp(Trim$(slFieldName), "mnfSSComm", 1) = 0) Then
                        slFieldDataType = Chr(101)
                    End If
                Else
                    If StrComp(slFieldDataType, "Length of Time", 1) = 0 Then
                        slFieldDataType = Chr(100)
                    End If
                    If StrComp(slFieldDataType, "Hex Field Dump", 1) = 0 Then
                        slFieldDataType = Chr(101)
                    End If
                End If
                ilFileID = mGetFileID(slFileName)
                For ilField1 = 0 To UBound(tmDDFField) - 1 Step 1
                    If (StrComp(Trim$(slFieldName), Trim$(tmDDFField(ilField1).sName), 1) = 0) And (ilFileID = tmDDFField(ilField1).iFileID) Then
                        ilRet = btrBeginTrans(hmFile, 1000)
                        'mGetField tmDDFField(tmRecSel(gdcData.SelStartCol).iFieldIndex1), 0
                        'slSvDataType = tmDDFField(ilField1).sDataType
                        If slFieldDataType <> "" Then
                            tmDDFField(ilField1).sDataType = slFieldDataType
                        End If
                        smFieldValue = slStr
                        mPutField tmDDFField(ilField1), 0
                        ilRet = btrDelete(hmFile)
                        If ilRet <> BTRV_ERR_NONE Then
                            MsgBox "Delete Failed, Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal
                            imTerminate = True
                            ilRet = btrAbortTrans(hmFile)
                            Exit Sub
                        End If
                        'Update record
                        If imVarLenRec Then
                            If imVarType = 12 Then
                                For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                                    If (Asc(tmDDFField(ilField).sDataType) = 12) Then
                                        If (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary") Then
                                            tmSsf = tmRec
                                            lmActualRecLen = imVarBaseRecLen + Len(tmProg) * tmSsf.iCount
                                        Else
                                            lmActualRecLen = imVarBaseRecLen + tmDDFField(ilField).iSize
                                        End If
                                        'ilRet = btrUpdate(hmFile, tmRec, lmActualRecLen)
                                        ilRet = btrInsert(hmFile, tmRec, lmActualRecLen, 0)
                                        Exit For
                                    End If
                                Next ilField
                            Else
                                ilRet = BTRV_ERR_NONE
                                For ilField = 0 To UBound(tmDDFField) - 1 Step 1
                                    If (Asc(tmDDFField(ilField).sDataType) = 13) Then
                                        ilIndex = tmDDFField(ilField).iOffset + 1
                                        slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                                        'HMemCpy ilValue, slInt2, 2
                                        lmActualRecLen = imVarBaseRecLen + ilValue + 4
                                        'ilRet = btrUpdate(hmFile, tmRec, lmActualRecLen)
                                        ilRet = btrInsert(hmFile, tmRec, lmActualRecLen, 0)
                                        Exit For
                                    End If
                                Next ilField
                            End If
                        Else
                            lmActualRecLen = imBaseRecLen
                            'ilRet = btrUpdate(hmFile, tmRec, lmActualRecLen)
                            ilRet = btrInsert(hmFile, tmRec, lmActualRecLen, 0)
                        End If
                        If ilRet <> BTRV_ERR_NONE Then
                            MsgBox "Insert Failed, Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal
                            ilRet = btrAbortTrans(hmFile)
                            imTerminate = True
                            Exit Sub
                        End If
                        ilRet = btrEndTrans(hmFile)
                        gdcData.Row = imEditRow
                        gdcData.Col = imEditCol
                        imEditRow = -1
                        imEditCol = -1
                        gdcData.Text = slStr
                        gdcData.Row = ilSvRow
                        gdcData.Col = ilSvCol
                        edcValue.Visible = False
                        Exit For
                    End If
                Next ilField1
            End If
        Else
            gdcData.Col = 0
            gdcData.Row = imEditRow
            slStr1 = Trim$(gdcData.Text)
            If (slStr1 = "") Or (slStr1 = "Deleted") Then
                slFieldInfo = lbcEditReplace(2).List(imEditCol - 1)
                ilRet = gParseItem(slFieldInfo, 1, "|", slFileName)
                ilRet = gParseItem(slFieldInfo, 2, "|", slFieldName)
                ilRet = gParseItem(slFieldInfo, 4, "|", slFieldDataType)
                If ilRet <> CP_MSG_NONE Then
                    slFieldDataType = ""
                    If (StrComp(Trim$(slFileName), "Mnf.Btr", 1) = 0) And (StrComp(Trim$(slFieldName), "mnfRPU", 1) = 0) Then
                        slFieldDataType = Chr(101)
                    End If
                    If (StrComp(Trim$(slFileName), "Mnf.Btr", 1) = 0) And (StrComp(Trim$(slFieldName), "mnfSSComm", 1) = 0) Then
                        slFieldDataType = Chr(101)
                    End If
                Else
                    If StrComp(slFieldDataType, "Length of Time", 1) = 0 Then
                        slFieldDataType = Chr(100)
                    End If
                    If StrComp(slFieldDataType, "Hex Field Dump", 1) = 0 Then
                        slFieldDataType = Chr(101)
                    End If
                End If
                ilFileID = mGetFileID(slFileName)
                For ilField1 = 0 To UBound(tmDDFField) - 1 Step 1
                    If (StrComp(Trim$(slFieldName), Trim$(tmDDFField(ilField1).sName), 1) = 0) And (ilFileID = tmDDFField(ilField1).iFileID) Then
                        'mGetField tmDDFField(tmRecSel(gdcData.SelStartCol).iFieldIndex1), 0
                        'slSvDataType = tmDDFField(ilField1).sDataType
                        If slFieldDataType <> "" Then
                            tmDDFField(ilField1).sDataType = slFieldDataType
                        End If
                        smFieldValue = slStr
                        mPutField tmDDFField(ilField1), 0
                        gdcData.Row = imEditRow
                        gdcData.Col = imEditCol
                        imEditRow = -1
                        imEditCol = -1
                        gdcData.Text = slStr
                        gdcData.Row = ilSvRow
                        gdcData.Col = ilSvCol
                        edcValue.Visible = False
                        Exit For
                    End If
                Next ilField1
            End If
            gdcData.Row = ilSvRow
            gdcData.Col = ilSvCol
        End If
    Else
        imEditRow = -1
        imEditCol = -1
        edcValue.Visible = False
        Exit Sub
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
        slDay = Trim$(Str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcField.Text = Format$(llDate, "m/d/yy")
                edcField.SelStart = 0
                edcField.SelLength = Len(edcField.Text)
                imBypassFocus = True
                edcField.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcField.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = Trim$(Str$(imCalMonth)) & "/15/" & Trim$(Str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub

Private Sub pbcLink_Click()

End Sub

Private Sub pbcLbcEditReplace0_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilEditEnd As Integer
    Dim ilField As Integer
    Dim slFields(1 To 4) As String
    Dim llFgColor As Long
    Dim llWidth As Long
    
    If rbcEditReplace(3).Value Then
        pbcLbcEditReplace0.Cls
        Exit Sub
    End If
    ilEditEnd = lbcEditReplace(0).TopIndex + lbcEditReplace(0).Height \ fgListHTArial825
    If ilEditEnd > lbcEditReplace(0).ListCount Then
        ilEditEnd = lbcEditReplace(0).ListCount
    End If
    If lbcEditReplace(0).ListCount <= lbcEditReplace(0).Height \ fgListHTArial825 Then
        llWidth = lbcEditReplace(0).Width - 30
    Else
        llWidth = lbcEditReplace(0).Width - igScrollBarWidth - 30
    End If
    pbcLbcEditReplace0.Width = llWidth
    pbcLbcEditReplace0.Cls
    llFgColor = pbcLbcEditReplace0.ForeColor
    For ilLoop = lbcEditReplace(0).TopIndex To ilEditEnd - 1 Step 1
        pbcLbcEditReplace0.ForeColor = llFgColor
        If lbcEditReplace(0).MultiSelect = 0 Then
            If lbcEditReplace(0).ListIndex = ilLoop Then
                gPaintArea pbcLbcEditReplace0, CSng(0), CSng((ilLoop - lbcEditReplace(0).TopIndex) * fgListHTArial825), CSng(pbcLbcEditReplace0.Width), CSng(fgListHTArial825) - 15, vbHighlight 'WHITE
                pbcLbcEditReplace0.ForeColor = vbWhite
            End If
        Else
            If lbcEditReplace(0).Selected(ilLoop) Then
                gPaintArea pbcLbcEditReplace0, CSng(0), CSng((ilLoop - lbcEditReplace(0).TopIndex) * fgListHTArial825), CSng(pbcLbcEditReplace0.Width), CSng(fgListHTArial825) - 15, vbHighlight 'WHITE
                pbcLbcEditReplace0.ForeColor = vbWhite
            End If
        End If
        slStr = lbcEditReplace(0).List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        For ilField = LBound(slFields) To UBound(slFields) Step 1
            pbcLbcEditReplace0.CurrentX = imListFieldR0(ilField)
            pbcLbcEditReplace0.CurrentY = (ilLoop - lbcEditReplace(0).TopIndex) * fgListHTArial825 + 15
            slStr = slFields(ilField)
            gAdjShowLen pbcLbcEditReplace0, slStr, imListFieldR0(ilField + 1) - imListFieldR0(ilField)
            pbcLbcEditReplace0.Print slStr
        Next ilField
        pbcLbcEditReplace0.ForeColor = llFgColor
    Next ilLoop

End Sub

Private Sub pbcLbcEditReplace2_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilEditEnd As Integer
    Dim ilField As Integer
    Dim slFields(1 To 5) As String
    Dim llFgColor As Long
    Dim llWidth As Long
    
    If rbcEditReplace(3).Value Then
        pbcLbcEditReplace2.Cls
        Exit Sub
    End If
    ilEditEnd = lbcEditReplace(2).TopIndex + lbcEditReplace(2).Height \ fgListHTArial825
    If ilEditEnd > lbcEditReplace(2).ListCount Then
        ilEditEnd = lbcEditReplace(2).ListCount
    End If
    If lbcEditReplace(2).ListCount <= lbcEditReplace(2).Height \ fgListHTArial825 Then
        llWidth = lbcEditReplace(2).Width - 30
    Else
        llWidth = lbcEditReplace(2).Width - igScrollBarWidth - 30
    End If
    pbcLbcEditReplace2.Width = llWidth
    pbcLbcEditReplace2.Cls
    llFgColor = pbcLbcEditReplace2.ForeColor
    For ilLoop = lbcEditReplace(2).TopIndex To ilEditEnd - 1 Step 1
        pbcLbcEditReplace2.ForeColor = llFgColor
        If lbcEditReplace(2).MultiSelect = 0 Then
            If lbcEditReplace(2).ListIndex = ilLoop Then
                gPaintArea pbcLbcEditReplace2, CSng(0), CSng((ilLoop - lbcEditReplace(2).TopIndex) * fgListHTArial825), CSng(pbcLbcEditReplace2.Width), CSng(fgListHTArial825) - 15, vbHighlight 'WHITE
                pbcLbcEditReplace2.ForeColor = vbWhite
            End If
        Else
            If lbcEditReplace(2).Selected(ilLoop) Then
                gPaintArea pbcLbcEditReplace2, CSng(0), CSng((ilLoop - lbcEditReplace(2).TopIndex) * fgListHTArial825), CSng(pbcLbcEditReplace2.Width), CSng(fgListHTArial825) - 15, vbHighlight 'WHITE
                pbcLbcEditReplace2.ForeColor = vbWhite
            End If
        End If
        slStr = lbcEditReplace(2).List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        For ilField = LBound(slFields) To UBound(slFields) Step 1
            pbcLbcEditReplace2.CurrentX = imListFieldR2(ilField)
            pbcLbcEditReplace2.CurrentY = (ilLoop - lbcEditReplace(2).TopIndex) * fgListHTArial825 + 15
            slStr = slFields(ilField)
            gAdjShowLen pbcLbcEditReplace2, slStr, imListFieldR2(ilField + 1) - imListFieldR2(ilField)
            pbcLbcEditReplace2.Print slStr
        Next ilField
        pbcLbcEditReplace2.ForeColor = llFgColor
    Next ilLoop
End Sub


Private Sub pbcLbcSelection1_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilSel1End As Integer
    Dim ilField As Integer
    Dim slFields(1 To 4) As String
    Dim llFgColor As Long
    Dim llWidth As Long
    
    ilSel1End = lbcSelection(1).TopIndex + lbcSelection(1).Height \ fgListHTArial825
    If ilSel1End > lbcSelection(1).ListCount Then
        ilSel1End = lbcSelection(1).ListCount
    End If
    If lbcSelection(1).ListCount <= lbcSelection(1).Height \ fgListHTArial825 Then
        llWidth = lbcSelection(1).Width - 30
    Else
        llWidth = lbcSelection(1).Width - igScrollBarWidth - 30
    End If
    pbcLbcSelection1.Width = llWidth
    pbcLbcSelection1.Cls
    llFgColor = pbcLbcSelection1.ForeColor
    For ilLoop = lbcSelection(1).TopIndex To ilSel1End - 1 Step 1
        pbcLbcSelection1.ForeColor = llFgColor
        If lbcSelection(1).MultiSelect = 0 Then
            If lbcSelection(1).ListIndex = ilLoop Then
                gPaintArea pbcLbcSelection1, CSng(0), CSng((ilLoop - lbcSelection(1).TopIndex) * fgListHTArial825), CSng(pbcLbcSelection1.Width), CSng(fgListHTArial825) - 15, vbHighlight 'WHITE
                pbcLbcSelection1.ForeColor = vbWhite
            End If
        Else
            If lbcSelection(1).Selected(ilLoop) Then
                gPaintArea pbcLbcSelection1, CSng(0), CSng((ilLoop - lbcSelection(1).TopIndex) * fgListHTArial825), CSng(pbcLbcSelection1.Width), CSng(fgListHTArial825) - 15, vbHighlight 'WHITE
                pbcLbcSelection1.ForeColor = vbWhite
            End If
        End If
        slStr = lbcSelection(1).List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        For ilField = LBound(slFields) To UBound(slFields) Step 1
            pbcLbcSelection1.CurrentX = imListFieldS(ilField)
            pbcLbcSelection1.CurrentY = (ilLoop - lbcSelection(1).TopIndex) * fgListHTArial825 + 15
            slStr = slFields(ilField)
            gAdjShowLen pbcLbcSelection1, slStr, imListFieldS(ilField + 1) - imListFieldS(ilField)
            pbcLbcSelection1.Print slStr
        Next ilField
        pbcLbcSelection1.ForeColor = llFgColor
    Next ilLoop

End Sub

Private Sub pbcSTab_GotFocus()
    mUpdateField
End Sub
Private Sub pbcTab_GotFocus()
    mUpdateField
End Sub
Private Sub plcEditReplace_DragDrop(Source As Control, X As Single, Y As Single)
    mClearDrag
End Sub
Private Sub plcEditReplace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcTabSelection_Click()
    Dim ilIndex As Integer
    
    ilIndex = plcTabSelection.SelectedItem.Index - 1
    If ilIndex = 0 Then   'DDF
        plcSelection.Visible = False
        'plcLinks.Visible = False
        plcEditReplace.Visible = False
        plcDDF.Visible = True
        plcDDF.ZOrder vbBringToFront
    'ElseIf ilIndex = 1 Then   'Links
    '    plcDDF.Visible = False
    '    plcEditReplace.Visible = False
    '    plcSelection.Visible = False
    '    plcLinks.Visible = True
    '    plcLinks.ZOrder vbBringToFront
    ElseIf ilIndex = 1 Then   'Selection
        plcDDF.Visible = False
        'plcLinks.Visible = False
        plcEditReplace.Visible = False
        plcSelection.Visible = True
        plcSelection.ZOrder vbBringToFront
    ElseIf ilIndex = 2 Then   'EditReplace
        plcDDF.Visible = False
        plcSelection.Visible = False
        'plcLinks.Visible = False
        plcEditReplace.Visible = True
        plcEditReplace.ZOrder vbBringToFront
    End If
    imTabSelection = ilIndex
End Sub
Private Sub rbcEditReplace_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcEditReplace(Index).Value
    'End of coded added
    If Value Then
        If ((Index = 2) Or (Index = 3)) And (Not imPasswordAsked) Then
            sgDBPath = lbcDataPath.Path & "\"
            CSIPass.Show vbModal
            If Not igPasswordOk Then
                lacMode.ForeColor = RED
                lacMode.Caption = "Mode: View Only"
                rbcEditReplace(2).Enabled = False
                rbcEditReplace(0).Value = True
                Exit Sub
            End If
            lacMode.ForeColor = DARKGREEN
            lacMode.Caption = "Mode: View and Update"
            imPasswordAsked = True
        Else
            If Not igPasswordOk Then
                rbcEditReplace(2).Enabled = False
                rbcEditReplace(0).Value = True
                Exit Sub
            End If
        End If
        lbcEditReplace(2).Clear
        lbcEditReplace(2).AddItem " "
        edcReplaceValue.Text = ""
        If Index = 0 Then
            lacEditReplace(0).Caption = "Show Fields"
            lacEditReplace(1).Enabled = False
            edcReplaceValue.Enabled = False
            cmcEditReplace(1).Enabled = True
            pbcLbcEditReplace0_Paint
        ElseIf Index = 1 Then
            lacEditReplace(0).Caption = "Show and Fix by Typing Fields"
            lacEditReplace(1).Enabled = False
            edcReplaceValue.Enabled = False
            cmcEditReplace(1).Enabled = True
        ElseIf Index = 2 Then
            lacEditReplace(0).Caption = "Fix by Mass Substitution Fields"
            lacEditReplace(1).Enabled = True
            edcReplaceValue.Enabled = True
            cmcEditReplace(1).Enabled = False   'Add All
        ElseIf Index = 3 Then
            lacEditReplace(0).Caption = ""
            lacEditReplace(1).Enabled = False
            lacEditReplace(2).Enabled = False
            lbcEditReplace(1).Enabled = False
            edcReplaceValue.Enabled = False
            cmcEditReplace(0).Enabled = False   'Add All
            cmcEditReplace(1).Enabled = False   'Add All
            cmcEditReplace(2).Enabled = False   'Add All
            cmcEditReplace(3).Enabled = False   'Add All
        End If
        pbcLbcEditReplace2_Paint
        mSetCommands
    End If
End Sub
Private Sub tmcDrag_Timer()
    Dim ilListIndex As Integer
    Dim llRg As Long
    Dim llRet As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilLen As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            If imDragButton <> 1 Then
                Exit Sub
            End If
            Select Case imDragSrce
                Case DRAGFIELD
                    ilListIndex = fmDragY \ fgListHtSerif825 + lbcEditReplace(0).TopIndex
                    If (ilListIndex >= 0) And (ilListIndex <= lbcEditReplace(0).ListCount - 1) Then
                        lbcEditReplace(0).ListIndex = ilListIndex
                        lbcEditReplace(0).DragIcon = imcIconDrag.DragIcon
                        imDragIndexSrce = ilListIndex
                        lbcEditReplace(0).Drag vbBeginDrag
                    Else
                        lbcEditReplace(0).ListIndex = -1
                    End If
                Case DRAGEDIT
                    ilListIndex = fmDragY \ fgListHTArial825 + lbcEditReplace(2).TopIndex
                    If (ilListIndex <= lbcEditReplace(2).ListCount - 2) Then
                        lbcEditReplace(2).ListIndex = ilListIndex
                        lbcEditReplace(2).DragIcon = imcIconDrag.DragIcon
                        imDragIndexSrce = ilListIndex
                        lbcEditReplace(2).Drag vbBeginDrag
                        imcTrash.Visible = True
                        imcTrash.Enabled = True
                    Else
                        lbcEditReplace(2).ListIndex = -1
                    End If
            End Select
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub tmcScroll_Timer()
    If imDragSrce = -1 Then
        tmcScroll.Enabled = False
        Exit Sub
    End If
    If imDragScroll = DRAGEDIT Then
        If imScrollDir = 0 Then 'Page up
           If lbcEditReplace(2).TopIndex > 0 Then
              lbcEditReplace(2).TopIndex = lbcEditReplace(2).TopIndex - 1
           End If
        Else    'Page down
           If lbcEditReplace(2).TopIndex < lbcEditReplace(2).ListCount - imNoRowLbc Then
              lbcEditReplace(2).TopIndex = lbcEditReplace(2).TopIndex + 1
           End If
        End If
    End If
End Sub
Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    cmcProcess_Click
    DoEvents
    If rbcEditReplace(2).Value Then
        mTerminate
        Exit Sub
    End If
End Sub
Private Sub vbcData_Change()
    Dim llStartRec As Long
    Dim llEndRec As Long
    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilRecIndex1 As Integer
    Dim ilRecIndex2 As Integer
    Dim ilRowLinkIndex1 As Integer
    Dim ilRowLinkIndex2 As Integer
    Dim ilColIndex1 As Integer
    Dim ilColIndex2 As Integer
    Dim ilTColIndex1 As Integer
    Dim ilTColIndex2 As Integer
    Dim ilGridType As Integer
    Dim ilRow As Integer
    Dim llRec As Long
    Dim ilRet As Integer
    Dim ilLoop1 As Integer
    Dim ilLoop2 As Integer
    Dim ilField As Integer
    Dim ilIndex As Integer
    Dim ilValue As Integer
    
    Screen.MousePointer = vbHourglass
    mUpdateField
    If imClsGrid Then
        For ilRow = 1 To gdcData.Rows - 1 Step 1
            gdcData.Row = ilRow
            'Blank columns
            For ilCol = 1 To gdcData.Cols Step 1
                gdcData.Col = ilCol - 1
                gdcData.Text = ""
            Next ilCol
        Next ilRow
        Exit Sub
    End If
    ilRow = 1
    llStartRec = CLng(imVBCAdjust) * (vbcData.Value - 1) + 1
    If (imRecMaxUpper1 = 1) And (llStartRec + vbcData.LargeChange > lmMaxNoRec) Then
        llEndRec = lmMaxNoRec
    Else
        llEndRec = llStartRec + vbcData.LargeChange
    End If
    llRec = llStartRec
    ilRecIndex1 = 1
    ilRecIndex2 = 1
    Do While llRec > 32000
        ilRecIndex1 = ilRecIndex1 + 1
        llRec = llRec - 32000
    Loop
    ilRecIndex2 = llRec
    imGrid1Upper1 = 1
    imGrid1Upper2 = 1
    imGrid2Upper1 = 1
    imGrid2Upper2 = 1
    imGrid3Upper1 = 1
    imGrid3Upper2 = 1
    imGrid4Upper1 = 1
    imGrid4Upper2 = 1
    imGrid5Upper1 = 1
    imGrid5Upper2 = 1
    imGrid6Upper1 = 1
    imGrid6Upper2 = 1
    imRecUpper1 = ilRecIndex1
    imRecUpper2 = ilRecIndex2
    For ilLoop1 = LBound(tmGrid1Data, 1) To UBound(tmGrid1Data, 1) Step 1
        For ilLoop2 = LBound(tmGrid1Data, 2) To UBound(tmGrid1Data, 2) Step 1
            tmGrid1Data(ilLoop1, ilLoop2).iNextIndex1 = -1
            tmGrid1Data(ilLoop1, ilLoop2).iNextIndex2 = -1
        Next ilLoop2
    Next ilLoop1
    For ilLoop1 = LBound(tmGrid2Data, 1) To UBound(tmGrid2Data, 1) Step 1
        For ilLoop2 = LBound(tmGrid2Data, 2) To UBound(tmGrid2Data, 2) Step 1
            tmGrid2Data(ilLoop1, ilLoop2).iNextIndex1 = -1
            tmGrid2Data(ilLoop1, ilLoop2).iNextIndex2 = -1
        Next ilLoop2
    Next ilLoop1
    For ilLoop1 = LBound(tmGrid3Data, 1) To UBound(tmGrid3Data, 1) Step 1
        For ilLoop2 = LBound(tmGrid3Data, 2) To UBound(tmGrid3Data, 2) Step 1
            tmGrid3Data(ilLoop1, ilLoop2).iNextIndex1 = -1
            tmGrid3Data(ilLoop1, ilLoop2).iNextIndex2 = -1
        Next ilLoop2
    Next ilLoop1
    For ilLoop1 = LBound(tmGrid4Data, 1) To UBound(tmGrid4Data, 1) Step 1
        For ilLoop2 = LBound(tmGrid4Data, 2) To UBound(tmGrid4Data, 2) Step 1
            tmGrid4Data(ilLoop1, ilLoop2).iNextIndex1 = -1
            tmGrid4Data(ilLoop1, ilLoop2).iNextIndex2 = -1
        Next ilLoop2
    Next ilLoop1
    For ilLoop1 = LBound(tmGrid5Data, 1) To UBound(tmGrid5Data, 1) Step 1
        For ilLoop2 = LBound(tmGrid5Data, 2) To UBound(tmGrid5Data, 2) Step 1
            tmGrid5Data(ilLoop1, ilLoop2).iNextIndex1 = -1
            tmGrid5Data(ilLoop1, ilLoop2).iNextIndex2 = -1
        Next ilLoop2
    Next ilLoop1
    For ilLoop1 = LBound(tmGrid6Data, 1) To UBound(tmGrid6Data, 1) Step 1
        For ilLoop2 = LBound(tmGrid6Data, 2) To UBound(tmGrid6Data, 2) Step 1
            tmGrid6Data(ilLoop1, ilLoop2).iNextIndex1 = -1
            tmGrid6Data(ilLoop1, ilLoop2).iNextIndex2 = -1
        Next ilLoop2
    Next ilLoop1
    For llRec = llStartRec To llEndRec Step 1
        If ((imRecUpper1 = imRecMaxUpper1) And (imRecUpper2 < imRecMaxUpper2)) Or ((imRecUpper1 < imRecMaxUpper1) And (imRecUpper2 <= 32000)) Then
            lmActualRecLen = imBaseRecLen
            ilRet = btrGetDirect(hmFile, tmRec, lmActualRecLen, tmRecInfo(imRecUpper1, imRecUpper2).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            mExtractFields
            imRecUpper2 = imRecUpper2 + 1
            If imRecUpper2 > 32000 Then
                imRecUpper1 = imRecUpper1 + 1
                imRecUpper2 = 1
            End If
        End If
    Next llRec
    
    For llRec = llStartRec To llEndRec Step 1
        If ((ilRecIndex1 = imRecMaxUpper1) And (ilRecIndex2 < imRecMaxUpper2)) Or ((ilRecIndex1 < imRecMaxUpper1) And (ilRecIndex2 <= 32000)) Then
            ilRowLinkIndex1 = tmRecInfo(ilRecIndex1, ilRecIndex2).iRowLinkIndex1
            ilRowLinkIndex2 = tmRecInfo(ilRecIndex1, ilRecIndex2).iRowLinkIndex2
            ilColIndex1 = tmRowLinkList(ilRowLinkIndex1, ilRowLinkIndex2).iFirstIndex1
            ilColIndex2 = tmRowLinkList(ilRowLinkIndex1, ilRowLinkIndex2).iFirstIndex2
            ilGridType = tmRowLinkList(ilRowLinkIndex1, ilRowLinkIndex2).iGridType
            gdcData.Row = ilRow
            gdcData.Col = 0
            gdcData.CellAlignment = flexAlignLeftTop
            gdcData.Text = Trim$(Str$(tmRecInfo(ilRecIndex1, ilRecIndex2).lRecNo))
            ilCol = 1
            Do While ilColIndex1 > 0
                gdcData.Col = ilCol
                ilCol = ilCol + 1
                gdcData.CellAlignment = flexAlignLeftTop
                Select Case ilGridType
                    Case 1
                        gdcData.Text = tmGrid1Data(ilColIndex1, ilColIndex2).sValue
                        ilGridType = tmGrid1Data(ilColIndex1, ilColIndex2).iNextGridType
                        ilTColIndex1 = tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex1
                        ilTColIndex2 = tmGrid1Data(ilColIndex1, ilColIndex2).iNextIndex2
                    Case 2
                        gdcData.Text = tmGrid2Data(ilColIndex1, ilColIndex2).sValue
                        ilGridType = tmGrid2Data(ilColIndex1, ilColIndex2).iNextGridType
                        ilTColIndex1 = tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex1
                        ilTColIndex2 = tmGrid2Data(ilColIndex1, ilColIndex2).iNextIndex2
                    Case 3
                        gdcData.Text = tmGrid3Data(ilColIndex1, ilColIndex2).sValue
                        ilGridType = tmGrid3Data(ilColIndex1, ilColIndex2).iNextGridType
                        ilTColIndex1 = tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex1
                        ilTColIndex2 = tmGrid3Data(ilColIndex1, ilColIndex2).iNextIndex2
                    Case 4
                        gdcData.Text = tmGrid4Data(ilColIndex1, ilColIndex2).sValue
                        ilGridType = tmGrid4Data(ilColIndex1, ilColIndex2).iNextGridType
                        ilTColIndex1 = tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex1
                        ilTColIndex2 = tmGrid4Data(ilColIndex1, ilColIndex2).iNextIndex2
                    Case 5
                        gdcData.Text = tmGrid5Data(ilColIndex1, ilColIndex2).sValue
                        ilGridType = tmGrid5Data(ilColIndex1, ilColIndex2).iNextGridType
                        ilTColIndex1 = tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex1
                        ilTColIndex2 = tmGrid5Data(ilColIndex1, ilColIndex2).iNextIndex2
                    Case 6
                        gdcData.Text = tmGrid6Data(ilColIndex1, ilColIndex2).sValue
                        ilGridType = tmGrid6Data(ilColIndex1, ilColIndex2).iNextGridType
                        ilTColIndex1 = tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex1
                        ilTColIndex2 = tmGrid6Data(ilColIndex1, ilColIndex2).iNextIndex2
                End Select
                ilColIndex1 = ilTColIndex1
                ilColIndex2 = ilTColIndex2
            Loop
            ilRecIndex2 = ilRecIndex2 + 1
            If ilRecIndex2 > 32000 Then
                ilRecIndex1 = ilRecIndex1 + 1
                ilRecIndex2 = 1
            End If
        Else
            gdcData.Row = ilRow
            'Blank columns
            For ilCol = 1 To gdcData.Cols Step 1
                gdcData.Col = ilCol - 1
                gdcData.Text = ""
            Next ilCol
        End If
        ilRow = ilRow + 1
    Next llRec
    Do While ilRow < gdcData.Rows
        gdcData.Row = ilRow
        'Blank columns
        For ilCol = 1 To gdcData.Cols Step 1
            gdcData.Col = ilCol - 1
            gdcData.Text = ""
        Next ilCol
        ilRow = ilRow + 1
    Loop
    Screen.MousePointer = vbDefault
End Sub
Private Sub vbcData_GotFocus()
    mUpdateField
End Sub
Private Sub plcPAS_Paint()
    plcPAS.CurrentX = 0
    plcPAS.CurrentY = 0
    plcPAS.Print "Include Prog/Avail/Spot Detail"
End Sub
Private Sub plcRun_Paint()
    plcRun.CurrentX = 0
    plcRun.CurrentY = 0
    plcRun.Print "When File Selected"
End Sub
Private Sub plcInclude_Paint()
    plcInclude.CurrentX = 0
    plcInclude.CurrentY = 0
    plcInclude.Print "Include"
End Sub

Public Sub mGetFileFields()
    Dim tlDDFField As DDFFIELD
    
    tlDDFField.iOffset = 0
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField, 0
    tmGetDDFFile.iFileID = Val(smFieldValue)
    tlDDFField.iOffset = 2
    tlDDFField.iSize = 20
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField, 0
    tmGetDDFFile.sName = smFieldValue
    tlDDFField.iOffset = 22
    tlDDFField.iSize = 64
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField, 0
    tmGetDDFFile.sLocation = smFieldValue
    tlDDFField.iOffset = 86
    tlDDFField.iSize = 1
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField, 0
    tmGetDDFFile.sFlags = smFieldValue
    'tmGetDDFFile.sLocation = smFieldValue
    tlDDFField.iOffset = 87
    tlDDFField.iSize = 10
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField, 0
    tmGetDDFFile.sReserved = smFieldValue
End Sub

Private Sub mFixLVarRecords()
    Dim ilField1 As Integer
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim ilRet As Integer
    
    If Not imVarLenRec Then
        Exit Sub
    End If
    If imVarType <> 13 Then
        Exit Sub
    End If
    If Not igPasswordOk Then
        Exit Sub
    End If
    If (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF Spot Summary") Or (Trim$(lbcFile.List(lbcFile.ListIndex)) = "SSF_Spot_Summary") Then
        Exit Sub
    End If
    imLVarLengthError = False
    For ilField1 = 0 To UBound(tmDDFField) - 1 Step 1
        If (Asc(tmDDFField(ilField1).sDataType) = 13) Then
            If tmDDFField(ilField1).iFileID = imFileID Then  'Key file
                mGetField tmDDFField(ilField1), 0
            Else
                mGetFieldByKey tmDDFField(ilField1)
            End If
            If (imLVarLengthError) Then
                ilIndex = tmDDFField(ilField1).iOffset + 1
                ilValue = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
                lmActualRecLen = imVarBaseRecLen + ilValue + 4
                ilRet = btrUpdate(hmFile, tmRec, lmActualRecLen)
            End If
            Exit For
        End If
    Next ilField1
End Sub

Private Function mSearchBuildIndex(ilFileID As Integer, ilKeyNo As Integer, tlDDFIndex As DDFINDEX) As Integer
    Dim ilRet As Integer
    Dim hlFile As Integer
    
    mSearchBuildIndex = False
    hlFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hlFile, "", smDDFPath & "Index.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Screen.MousePointer = 0 'DEFAULT
        Exit Function
    End If
    lmActualRecLen = Len(tlDDFIndex) 'Len(tmRec)
    imBaseRecLen = lmActualRecLen
    tmIndexSrchKey0.iFileID = ilFileID
    'ilRet = btrGetGreaterOrEqual(hlFile, tlDDFIndex(ilUpper), lmActualRecLen, tmIndexSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
    ilRet = btrGetGreaterOrEqual(hlFile, tmRec, lmActualRecLen, tmIndexSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = 0 'DEFAULT
        mBtrvErrorMsg ilRet, "btrGetGreaterOrEqual: " & "INDEX.DDF"
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If
    Do While (ilRet = BTRV_ERR_NONE)
        If imTerminate Then
            ilRet = btrClose(hlFile)
            btrDestroy hlFile
            mTerminate
            Exit Function
        End If
        mGetIndexFields
        tlDDFIndex = tmGetDDFIndex
        If (tlDDFIndex.iFileID <> ilFileID) Then
            Exit Do
        End If
        If (tlDDFIndex.iNumber = ilKeyNo) And (tlDDFIndex.iPart = 0) Then
            mSearchBuildIndex = True
            Exit Do
        End If
        lmActualRecLen = imBaseRecLen
        ilRet = btrGetNext(hlFile, tmRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
    
End Function

Private Function mSearchBuildField(ilFileID As Integer, ilFieldID As Integer, tlDDFField As DDFFIELD) As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilMax As Integer
    Dim ilOffset As Integer
    Dim hlFile As Integer
    
    mSearchBuildField = False
    hlFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hlFile, "", smDDFPath & "Field.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If
    lmActualRecLen = Len(tlDDFField) 'Len(tmRec)
    imBaseRecLen = lmActualRecLen
    tmFieldSrchKey1.iFileID = ilFileID
    'ilRet = btrGetGreaterOrEqual(hmFile, tlField(ilUpper), lmActualRecLen, tmFieldSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    ilRet = btrGetGreaterOrEqual(hlFile, tmRec, lmActualRecLen, tmFieldSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mBtrvErrorMsg ilRet, "btrGetGreaterOrEqual: " & "FIELD.DDF"
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If
    Do While (ilRet = BTRV_ERR_NONE)
        If imTerminate Then
            ilRet = btrClose(hlFile)
            btrDestroy hlFile
            mTerminate
            Exit Function
        End If
        mGetFieldFields
        tlDDFField = tmGetDDFFIELD
        If (tlDDFField.iFileID <> ilFileID) Then
            Exit Do
        End If
        If tlDDFField.iFieldID = ilFieldID Then
            mSearchBuildField = True
            Exit Do
        End If
        lmActualRecLen = imBaseRecLen
        ilRet = btrGetNext(hlFile, tmRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
End Function

Private Sub mGetKeyFieldName()
    Dim slFile As String
    Dim slName As String
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim ilKeyIndex As Integer
    Dim ilRet As Integer
    
    lacSearchFrom.Caption = "From:"
    If lbcFile.ListIndex < 0 Then
        Exit Sub
    End If
    If imFileID < 0 Then
        Exit Sub
    End If
    ilKeyIndex = Val(edcSearchKey.Text)
    ilRet = mSearchBuildIndex(imFileID, ilKeyIndex, tmSearchDDFIndex)
    If Not ilRet Then
        Exit Sub
    End If
    ilRet = mSearchBuildField(imFileID, tmSearchDDFIndex.iFieldID, tmSearchDDFField)
    If Not ilRet Then
        Exit Sub
    End If
    lacSearchFrom.Caption = Trim$(tmSearchDDFField.sName) & "- From:"
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMassRemoveRecords              *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Mass pdate records              *
'*                                                     *
'*******************************************************
Private Sub mMassRemoveRecords()
    Dim ilRet As Integer
    Dim llRecNo As Long
    Dim llStartRecNo As Long
    Dim llEndRecNo As Long
    Dim ilRecIndex1 As Integer
    Dim ilRecIndex2 As Integer
    Dim llRec As Long
    Dim ilLoop As Integer
    Dim ilField As Integer
    Dim ilDataType As Integer
    Dim slSvDataType As String
    Dim slFieldDataType As String
    Dim slFieldInfo As String
    Dim slFieldName As String
    Dim ilSize As Integer
    Dim ilCol As Integer
    Dim ilPos As Integer
    Dim slValue As String
    Dim llValue As Long
    Dim slInt2 As INT2
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim llPercent As Long
    'Dim slSvFieldValue As String
    Dim ilOper As Integer
    Dim slAdjustValue As String
    llStartRecNo = 1
    llEndRecNo = lmMaxNoRec
    llRec = llStartRecNo
    ilRecIndex1 = 1
    ilRecIndex2 = 1
    Do While llRec > 32000
        ilRecIndex1 = ilRecIndex1 + 1
        llRec = llRec - 32000
    Loop
    ilRecIndex2 = llRec
    ilRet = 0
    On Error GoTo mMassRemoveRecordsErr:
    DoEvents
    If imTerminate Then
        Exit Sub
    End If
    imVarBaseRecLen = 0
    For ilField = 0 To UBound(tmDDFField) - 1 Step 1
        If (Asc(tmDDFField(ilField).sDataType) = 12) Then
            imVarLenRec = True
            imVarType = 12
        ElseIf (Asc(tmDDFField(ilField).sDataType) = 13) Then
            imVarLenRec = True
            imVarType = 13
        Else
            imVarBaseRecLen = imVarBaseRecLen + tmDDFField(ilField).iSize
        End If
    Next ilField
    For llRecNo = llStartRecNo To llEndRecNo Step 1
        If ((ilRecIndex1 = imRecMaxUpper1) And (ilRecIndex2 < imRecMaxUpper2)) Or ((ilRecIndex1 < imRecMaxUpper1) And (ilRecIndex2 <= 32000)) Then
            Do
                ilRet = btrBeginTrans(hmFile, 1000)
                lmActualRecLen = imBaseRecLen
                ilRet = btrGetDirect(hmFile, tmRec, lmActualRecLen, tmRecInfo(ilRecIndex1, ilRecIndex2).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                DoEvents
                If imTerminate Then
                    Exit Sub
                End If
                ilRet = btrDelete(hmFile)
                DoEvents
                If imTerminate Then
                    ilRet = btrAbortTrans(hmFile)
                    Exit Sub
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                MsgBox "Delete Failed, Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal
                imTerminate = True
                ilRet = btrAbortTrans(hmFile)
                Exit Sub
            End If
            ilRet = btrEndTrans(hmFile)
            ilRecIndex2 = ilRecIndex2 + 1
            If ilRecIndex2 > 32000 Then
                ilRecIndex1 = ilRecIndex1 + 1
                ilRecIndex2 = 1
            End If
            lacChg.Caption = "Updated" & Str$(llRecNo) & " of" & Str$(lmMaxNoRec)
            llPercent = 100 * llRecNo / lmMaxNoRec
            If llPercent > 100 Then
                llPercent = 100
            End If
            If lmPercent <> llPercent Then
                plcChg.Caption = Trim$(Str$(llPercent)) & "%"
                lmPercent = llPercent
            End If
        End If
    Next llRecNo
    lacChg.Caption = "Removed" & Str$(lmMaxNoRec) & " of" & Str$(lmMaxNoRec)
    plcChg.Caption = "100%"
    On Error GoTo 0
    Exit Sub
mMassRemoveRecordsErr:
    ilRet = Err
    Resume Next
End Sub

Private Function mSearchCompleted(tlDDFField As DDFFIELD, slSearchTo As String) As Integer
    mSearchCompleted = False
    Select Case Asc(tlDDFField.sDataType)
        Case 0  'String
            If StrComp(smFieldValue, slSearchTo, vbTextCompare) > 0 Then
                mSearchCompleted = True
            End If
        Case 1  'Integer
            If Val(smFieldValue) > Val(slSearchTo) Then
                mSearchCompleted = True
            End If
        Case 2  'IEEE Float
        Case 3  'Btrieve Date
            If gDateValue(smFieldValue) > gDateValue(slSearchTo) Then
                mSearchCompleted = True
            End If
        Case 4  'Btrieve Time
            If gTimeToLong(smFieldValue, False) > gTimeToLong(slSearchTo, False) Then
                mSearchCompleted = True
            End If
        Case 5  'COBOL Decimal COMP-3
        Case 6  'COBOL Money
        Case 7  'Logical
        Case 8  'COBOL Numeric
        Case 9  'BASIC Float
        Case 10 'Pascal LString
        Case 11 'C ZString
            If StrComp(smFieldValue, slSearchTo, vbTextCompare) > 0 Then
                mSearchCompleted = True
            End If
        Case 12 'Variable Length Note
        Case 13 'LVar (length, Characters, 0, 0); Length does not include zeros
        Case 14 'Unsigned Binary
        Case 15 'AutoIncrement
            If Val(smFieldValue) > Val(slSearchTo) Then
                mSearchCompleted = True
            End If
        Case 16 'Bit
        Case 17 'COBOL Numeric STS
        Case 100  'Length
        Case 101
    End Select
End Function



Private Function mCheckSearchValues() As Integer
    Dim slSearchFrom As String
    Dim slSearchTo As String
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilKeyIndex As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilPass As Integer
    
    mCheckSearchValues = True
    ilKeyIndex = Val(edcSearchKey.Text)
    slSearchFrom = Trim$(edcSearchFrom.Text)
    slSearchTo = Trim$(edcSearchTo.Text)
    If (slSearchFrom <> "") Or (slSearchTo <> "") Then
        slName = lbcFile.List(lbcFile.ListIndex)
        For ilLoop = 0 To UBound(tmDDFFile) - 1 Step 1
            If StrComp(Trim$(tmDDFFile(ilLoop).sName), slName, 1) = 0 Then
                imFileID = tmDDFFile(ilLoop).iFileID
                Exit For
            End If
        Next ilLoop
        ilRet = mSearchBuildIndex(imFileID, ilKeyIndex, tmSearchDDFIndex)
        If Not ilRet Then
            mCheckSearchValues = False
            Exit Function
        End If
        ilRet = mSearchBuildField(imFileID, tmSearchDDFIndex.iFieldID, tmSearchDDFField)
        If Not ilRet Then
            mCheckSearchValues = False
            Exit Function
        End If
        For ilPass = 0 To 1 Step 1
            If ilPass = 0 Then
                slStr = slSearchFrom
            Else
                slStr = slSearchTo
            End If
            If slStr <> "" Then
                If Not mCheckValue(slStr, Asc(tmSearchDDFField.sDataType), tmSearchDDFField.iSize) Then
                    mCheckSearchValues = False
                    Exit Function
                End If
            End If
        Next ilPass
    End If
End Function

Private Function mCheckValue(slValue As String, ilDataType As Integer, ilSize As Integer) As Integer
    Dim ilLoop As Integer
    Dim ilNoDecPt As Integer
    
    mCheckValue = True
    Select Case ilDataType
        Case 0  'String
            If Len(slValue) > ilSize Then
                mCheckValue = False
                Exit Function
            End If
        Case 1  'Integer
            For ilLoop = 1 To Len(slValue) Step 1
                If (Mid$(slValue, ilLoop, 1) < "0") Or (Mid$(slValue, ilLoop, 1) > "9") Then
                    If (ilLoop <> 1) Or ((Mid$(slValue, ilLoop, 1) <> "-") And (Mid$(slValue, ilLoop, 1) <> "+")) Then
                        mCheckValue = False
                        Exit Function
                    End If
                End If
            Next ilLoop
            If ilSize = 1 Then
                If (Val(slValue) > 255) Or (Val(slValue) < -256) Then
                    mCheckValue = False
                    Exit Function
                End If
            ElseIf ilSize = 2 Then
                If (Val(slValue) > 32767) Or (Val(slValue) < -32768) Then
                    mCheckValue = False
                    Exit Function
                End If
            Else
                If (Val(slValue) > 2147483647) Or (Val(slValue) < -2147483648#) Then
                    mCheckValue = False
                    Exit Function
                End If
            End If
        Case 2  'IEEE Float
        Case 3  'Btrieve Date
            If Trim$(slValue) <> "" Then
                If Not gValidDate(slValue) Then
                    mCheckValue = False
                    Exit Function
                End If
            End If
        Case 4  'Btrieve Time
            If Not gValidTime(slValue) Then
                mCheckValue = False
                Exit Function
            End If
        Case 5  'COBOL Decimal COMP-3
            ilNoDecPt = 0
            For ilLoop = 1 To Len(slValue) Step 1
                If ((Mid$(slValue, ilLoop, 1) < "0") Or (Mid$(slValue, ilLoop, 1) > "9")) And (Mid$(slValue, ilLoop, 1) <> ".") Then
                    If (ilLoop <> 1) Or ((Mid$(slValue, ilLoop, 1) <> "-") And (Mid$(slValue, ilLoop, 1) <> "+")) Then
                        mCheckValue = False
                        Exit Function
                    End If
                End If
                If Mid$(slValue, ilLoop, 1) = "." Then
                    ilNoDecPt = ilNoDecPt + 1
                    If ilNoDecPt > 1 Then
                        mCheckValue = False
                        Exit Function
                    End If
                End If
            Next ilLoop
            If Len(slValue) > 2 * ilSize - 1 Then
                mCheckValue = False
                Exit Function
            End If
        Case 6  'COBOL Money
            ilNoDecPt = 0
            For ilLoop = 1 To Len(slValue) Step 1
                If ((Mid$(slValue, ilLoop, 1) < "0") Or (Mid$(slValue, ilLoop, 1) > "9")) And (Mid$(slValue, ilLoop, 1) <> ".") Then
                    If (ilLoop <> 1) Or ((Mid$(slValue, ilLoop, 1) <> "-") And (Mid$(slValue, ilLoop, 1) <> "+")) Then
                        mCheckValue = False
                        Exit Function
                    End If
                End If
                If Mid$(slValue, ilLoop, 1) = "." Then
                    ilNoDecPt = ilNoDecPt + 1
                    If ilNoDecPt > 1 Then
                        mCheckValue = False
                        Exit Function
                    End If
                End If
            Next ilLoop
            If Len(slValue) > 2 * ilSize - 1 Then
                mCheckValue = False
                Exit Function
            End If
        Case 7  'Logical
        Case 8  'COBOL Numeric
        Case 9  'BASIC Float
        Case 10 'Pascal LString
        Case 11 'C ZString
            If Len(slValue) > ilSize Then
                mCheckValue = False
                Exit Function
            End If
        Case 12 'Variable Length Note
        Case 13 'LVar (length, Characters, 0, 0); Length does not include zeros
        Case 14 'Unsigned Binary
        Case 15 'AutoIncrement
            For ilLoop = 1 To Len(slValue) Step 1
                If (Asc(Mid$(slValue, ilLoop, 1)) < Asc("0")) Or (Asc(Mid$(slValue, ilLoop, 1)) > Asc("9")) Then
                    mCheckValue = False
                    Exit Function
                End If
            Next ilLoop
            If ilSize = 2 Then
                If (Val(slValue) > 32767) Then
                    mCheckValue = False
                    Exit Function
                End If
            Else
                If (Val(slValue) > 2147483647) Then
                    mCheckValue = False
                    Exit Function
                End If
            End If
        Case 16 'Bit
        Case 17 'COBOL Numeric STS
        Case 100  'Length
            If Not gValidLength(slValue) Then
                mCheckValue = False
                Exit Function
            End If
        Case 101
    End Select
End Function
