VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl AffExportCriteria 
   Appearance      =   0  'Flat
   ClientHeight    =   2940
   ClientLeft      =   840
   ClientTop       =   2190
   ClientWidth     =   10170
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
   ScaleHeight     =   2940
   ScaleWidth      =   10170
   Begin VB.Frame frciPump 
      Caption         =   "iPump"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   8370
      TabIndex        =   113
      Top             =   1815
      Visible         =   0   'False
      Width           =   2130
      Begin VB.CheckBox ckciPOutput 
         Caption         =   "Generate Facts File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   114
         Top             =   315
         Width           =   1770
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   -30
      Top             =   2445
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2940
      FormDesignWidth =   10170
   End
   Begin VB.Frame frcXDS 
      Caption         =   "X-Digital"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   240
      TabIndex        =   90
      Top             =   1290
      Visible         =   0   'False
      Width           =   8610
      Begin VB.CheckBox ckcXReexport 
         Caption         =   "Re-Export all, not just new"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4920
         TabIndex        =   115
         Top             =   480
         Width           =   2610
      End
      Begin VB.Frame frcXProvider 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4335
         TabIndex        =   93
         Top             =   210
         Width           =   4455
         Begin VB.OptionButton rbcXProvider 
            Caption         =   "HeadEnd 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   2445
            TabIndex        =   96
            Top             =   0
            Width           =   1500
         End
         Begin VB.OptionButton rbcXProvider 
            Caption         =   "HeadEnd 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   870
            TabIndex        =   95
            Top             =   0
            Width           =   1500
         End
         Begin VB.Label lacXProvider 
            Caption         =   "Provider"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   94
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.TextBox edcXExportToPath 
         Height          =   225
         Left            =   1605
         TabIndex        =   105
         Top             =   1155
         Width           =   4650
      End
      Begin VB.CommandButton cmcXBrowse 
         Caption         =   "&Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6330
         TabIndex        =   106
         Top             =   1110
         Width           =   1170
      End
      Begin VB.CheckBox ckcXExportType 
         Caption         =   "File Delivery (Envelopes)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   92
         Top             =   225
         Width           =   2685
      End
      Begin VB.CheckBox ckcXExportType 
         Caption         =   "Spot Insertions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   91
         Top             =   225
         Width           =   1770
      End
      Begin VB.Frame frcXSuppressNotices 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   225
         Left            =   0
         TabIndex        =   97
         Top             =   495
         Width           =   4740
         Begin VB.OptionButton optXGenType 
            Caption         =   "Generate XML Test File"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2145
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   -15
            Width           =   2415
         End
         Begin VB.OptionButton optXGenType 
            Caption         =   "Send To X-Digital"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   1980
         End
      End
      Begin VB.Frame frmXVeh 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   0
         TabIndex        =   100
         Top             =   780
         Width           =   4410
         Begin VB.OptionButton rbcXSpots 
            Caption         =   "All Spots"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   810
            TabIndex        =   102
            Top             =   0
            Width           =   1305
         End
         Begin VB.OptionButton rbcXSpots 
            Caption         =   "Regional Spots"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   2115
            TabIndex        =   103
            Top             =   0
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.Label lacXSpots 
            Caption         =   "Export"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   101
            Top             =   0
            Width           =   660
         End
      End
      Begin VB.Label lacXExportTo 
         Caption         =   "Export To Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   104
         Top             =   1185
         Width           =   1635
      End
   End
   Begin VB.Frame frcWegener 
      Caption         =   "Wegener"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   8505
      TabIndex        =   77
      Top             =   1560
      Visible         =   0   'False
      Width           =   8610
      Begin VB.TextBox edcWExportToPath 
         Height          =   240
         Left            =   1680
         TabIndex        =   88
         Top             =   1035
         Width           =   4920
      End
      Begin VB.CommandButton cmcWExportTo 
         Caption         =   "&Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6720
         TabIndex        =   89
         Top             =   1005
         Width           =   1170
      End
      Begin VB.CheckBox ckcWGenerate 
         Caption         =   "Import Station Info"
         Height          =   195
         Index           =   0
         Left            =   8145
         TabIndex        =   86
         Top             =   630
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.CheckBox ckcWGenerate 
         Caption         =   "Generate Wegener Export"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   82
         Top             =   705
         Value           =   1  'Checked
         Width           =   2640
      End
      Begin VB.CheckBox ckcWGenCSV 
         Caption         =   "Generate CSV file along with XML file"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2640
         TabIndex        =   83
         Top             =   705
         Width           =   3885
      End
      Begin VB.TextBox txtWStationInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1680
         TabIndex        =   79
         Top             =   180
         Width           =   5355
      End
      Begin VB.TextBox txtWRunLetter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7710
         MaxLength       =   1
         TabIndex        =   85
         Top             =   660
         Width           =   405
      End
      Begin VB.CommandButton cmcWBrowse 
         Caption         =   "&Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7260
         TabIndex        =   80
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label lacWExportTo 
         Caption         =   "Export To Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   87
         Top             =   1080
         Width           =   1620
      End
      Begin VB.Label lacWFileInfo 
         Caption         =   "(Import information from: RX_Calls and JNS_RecGroup.Csv)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   60
         TabIndex        =   81
         Top             =   525
         Width           =   7320
      End
      Begin VB.Label lacWStationInfo 
         Caption         =   "Station Info Path"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   78
         Top             =   195
         Width           =   1575
      End
      Begin VB.Label lacWRunLetter 
         Caption         =   "Run Letter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6750
         TabIndex        =   84
         Top             =   705
         Width           =   990
      End
   End
   Begin VB.Frame frcIDC 
      Caption         =   "IDC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   8670
      TabIndex        =   69
      Top             =   1320
      Visible         =   0   'False
      Width           =   7590
      Begin VB.CommandButton cmcDBrowse 
         Caption         =   "&Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6300
         TabIndex        =   76
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox edcDExportToPath 
         Height          =   240
         Left            =   1635
         TabIndex        =   75
         Top             =   585
         Width           =   4500
      End
      Begin VB.CheckBox ckcDGenType 
         Caption         =   "Send to IDC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   70
         Top             =   240
         Width           =   1410
      End
      Begin VB.CheckBox ckcDGenType 
         Caption         =   "Send to Excel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4965
         TabIndex        =   73
         Top             =   270
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CheckBox ckcDGenType 
         Caption         =   "Generate Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1425
         TabIndex        =   71
         Top             =   240
         Width           =   1590
      End
      Begin VB.CheckBox ckcDGenType 
         Caption         =   "Generate Audio List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   3045
         TabIndex        =   72
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label lacDExportTo 
         Caption         =   "Export To Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   74
         Top             =   630
         Width           =   1635
      End
   End
   Begin VB.Frame frcStarGuide 
      Caption         =   "StarGuide"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   8865
      TabIndex        =   59
      Top             =   1050
      Visible         =   0   'False
      Width           =   7440
      Begin VB.TextBox edcSExportToPath 
         Height          =   240
         Left            =   1545
         TabIndex        =   67
         Top             =   615
         Width           =   4500
      End
      Begin VB.CommandButton cmcSBrowse 
         Caption         =   "&Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6180
         TabIndex        =   68
         Top             =   585
         Width           =   1170
      End
      Begin VB.TextBox txtSRunLetter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5400
         MaxLength       =   1
         TabIndex        =   65
         Top             =   180
         Width           =   405
      End
      Begin VB.Frame frmSVeh 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   0
         TabIndex        =   60
         Top             =   195
         Width           =   4200
         Begin VB.OptionButton rbcSSpots 
            Caption         =   "All Spots"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   810
            TabIndex        =   62
            Top             =   45
            Width           =   1380
         End
         Begin VB.OptionButton rbcSSpots 
            Caption         =   "Regional Spots"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   2190
            TabIndex        =   63
            Top             =   45
            Width           =   1845
         End
         Begin VB.Label lacSStarGuide 
            Caption         =   "Export"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            TabIndex        =   61
            Top             =   45
            Width           =   660
         End
      End
      Begin VB.Label lacSExportTo 
         Caption         =   "Export To Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   66
         Top             =   660
         Width           =   1770
      End
      Begin VB.Label lacSRunLetter 
         Caption         =   "Run Letter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4410
         TabIndex        =   64
         Top             =   225
         Width           =   1095
      End
   End
   Begin VB.Frame frcCnC 
      Caption         =   "Clearance n Compensation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   8160
      TabIndex        =   54
      Top             =   840
      Visible         =   0   'False
      Width           =   8280
      Begin VB.TextBox txtCFile 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         TabIndex        =   56
         Top             =   240
         Width           =   4665
      End
      Begin VB.CommandButton cmcCBrowse 
         Caption         =   "&Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5925
         TabIndex        =   55
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label lbcFile 
         Caption         =   "Export File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   58
         Top             =   225
         Width           =   1155
      End
      Begin VB.Label lacNote 
         Caption         =   "Note: Only dates marked as Completed will be exported"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   57
         Top             =   630
         Width           =   5055
      End
   End
   Begin VB.Frame frcISCIXRef 
      Caption         =   "ISCI Cross Reference"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8640
      TabIndex        =   49
      Top             =   720
      Visible         =   0   'False
      Width           =   7320
      Begin VB.Frame frcRPrefix 
         Appearance      =   0  'Flat
         Caption         =   "RISCIPrefix"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3240
         TabIndex        =   116
         Top             =   120
         Width           =   3855
         Begin VB.OptionButton rbcRPrefix 
            Caption         =   "Break"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   118
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton rbcRPrefix 
            Caption         =   "National ISCI Model"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   117
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lacRPrefix 
            Caption         =   "ISCI Prefix:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   119
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.TextBox edcRExportToPath 
         Height          =   240
         Left            =   1635
         TabIndex        =   52
         Top             =   540
         Width           =   4320
      End
      Begin VB.CommandButton cmcRBrowse 
         Caption         =   "&Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6090
         TabIndex        =   53
         Top             =   510
         Width           =   1170
      End
      Begin VB.CheckBox ckcRIncludeGeneric 
         Caption         =   "Include Generic copy lacking split copy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   50
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lacRExportTo 
         Caption         =   "Export To Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -15
         TabIndex        =   51
         Top             =   585
         Width           =   1680
      End
   End
   Begin VB.Frame frcRCS 
      Caption         =   "RCS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   9240
      TabIndex        =   40
      Top             =   270
      Visible         =   0   'False
      Width           =   7470
      Begin VB.TextBox edcRCSExportToPath 
         Height          =   240
         Left            =   1680
         TabIndex        =   47
         Top             =   555
         Width           =   4395
      End
      Begin VB.CommandButton cmcRCSBrowse 
         Caption         =   "&Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6225
         TabIndex        =   48
         Top             =   525
         Width           =   1170
      End
      Begin VB.CheckBox chkRCSZone 
         Caption         =   "PST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3345
         TabIndex        =   45
         Top             =   210
         Value           =   1  'Checked
         Width           =   780
      End
      Begin VB.CheckBox chkRCSZone 
         Caption         =   "MST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2490
         TabIndex        =   44
         Top             =   210
         Value           =   1  'Checked
         Width           =   885
      End
      Begin VB.CheckBox chkRCSZone 
         Caption         =   "CST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1665
         TabIndex        =   43
         Top             =   210
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chkRCSZone 
         Caption         =   "EST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   885
         TabIndex        =   42
         Top             =   210
         Value           =   1  'Checked
         Width           =   780
      End
      Begin VB.Label lacRCSZone 
         Caption         =   "Zone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   210
         Width           =   855
      End
      Begin VB.Label lacRCSExportTo 
         Caption         =   "Export To Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   46
         Top             =   600
         Width           =   1680
      End
   End
   Begin VB.Frame frcISCI 
      Caption         =   "ISCI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   9135
      TabIndex        =   20
      Top             =   105
      Visible         =   0   'False
      Width           =   8985
      Begin VB.Frame frcIVeh 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4185
         TabIndex        =   107
         Top             =   180
         Visible         =   0   'False
         Width           =   4980
         Begin VB.OptionButton rbcISpots 
            Caption         =   "Resend"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   4410
            TabIndex        =   110
            Top             =   60
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.OptionButton rbcISpots 
            Caption         =   "Unsent"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   810
            TabIndex        =   109
            Top             =   45
            Width           =   975
         End
         Begin VB.OptionButton rbcISpots 
            Caption         =   "All"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1980
            TabIndex        =   108
            Top             =   45
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.Label Label3 
            Caption         =   "ISCI"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   15
            TabIndex        =   111
            Top             =   45
            Width           =   660
         End
      End
      Begin VB.TextBox edcINoDaysResend 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6390
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "30"
         Top             =   795
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox edcIExportToPath 
         Height          =   270
         Left            =   1575
         TabIndex        =   38
         Top             =   1155
         Width           =   5370
      End
      Begin VB.CommandButton cmcIBrowse 
         Caption         =   "&Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7170
         TabIndex        =   39
         Top             =   1140
         Width           =   1050
      End
      Begin VB.Frame frcIISCIByBreak 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   0
         TabIndex        =   25
         Top             =   525
         Visible         =   0   'False
         Width           =   7320
         Begin VB.CheckBox ckcIIncludeCommands 
            Caption         =   "Timezone"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   5400
            TabIndex        =   120
            Top             =   0
            Width           =   990
         End
         Begin VB.CheckBox ckcIIncludeCommands 
            Caption         =   "Spot ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   4020
            TabIndex        =   28
            Top             =   0
            Width           =   990
         End
         Begin VB.CheckBox ckcIIncludeCommands 
            Caption         =   "Program Segments"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   2010
            TabIndex        =   27
            Top             =   0
            Value           =   1  'Checked
            Width           =   2115
         End
         Begin VB.Label lacIProgSeg 
            Caption         =   "ISCI By Break include"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.Frame frcIExportType 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   615
         TabIndex        =   22
         Top             =   225
         Width           =   3435
         Begin VB.OptionButton rbcIExportType 
            Caption         =   "ISCI by Break"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1860
            TabIndex        =   24
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton rbcIExportType 
            Caption         =   "Unique ISCI"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   23
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.CheckBox ckcIIncludeCommands 
         Caption         =   "ISCI Export Contact(s) E-Mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   885
         TabIndex        =   34
         Top             =   825
         Width           =   2985
      End
      Begin VB.Frame frcIUniqueISCI 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   5145
         TabIndex        =   29
         Top             =   525
         Visible         =   0   'False
         Width           =   4935
         Begin VB.OptionButton rbcIUniqueBy 
            Caption         =   "Vehicle"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   31
            Top             =   0
            Width           =   1275
         End
         Begin VB.OptionButton rbcIUniqueBy 
            Caption         =   "Vehicle and  Station"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   2700
            TabIndex        =   32
            Top             =   0
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.Label lacIUniqueISCI 
            Caption         =   "Unique ISCI By:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   1650
         End
      End
      Begin VB.Label lacINoDaysResend 
         Caption         =   "# Days Off Air to Resend ISCI"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3990
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.Label lacIExportToPath 
         Caption         =   "Export to Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   37
         Top             =   1185
         Width           =   1560
      End
      Begin VB.Label lacIEMail 
         Caption         =   "Include"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   33
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label lacIExportType 
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   21
         Top             =   210
         Width           =   1260
      End
   End
   Begin VB.Frame frcStationLog 
      Caption         =   "Station Log"
      Height          =   1770
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   7665
      Begin VB.Frame frcCumulus 
         Height          =   780
         Left            =   6240
         TabIndex        =   17
         Top             =   360
         Width           =   1800
         Begin VB.CheckBox ckcCUSendEmails 
            Caption         =   "Send Emails"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   360
            TabIndex        =   19
            Top             =   240
            Value           =   1  'Checked
            Width           =   1350
         End
         Begin VB.CheckBox ckcCumulus 
            Caption         =   "Cumulus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   18
            Top             =   15
            Value           =   1  'Checked
            Width           =   1050
         End
      End
      Begin VB.Frame frcCSIWeb 
         Height          =   780
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   3330
         Begin VB.CheckBox ckcCWRemoveISCI 
            Caption         =   "Remove ISCI from IDC Agreements"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   285
            TabIndex        =   112
            Top             =   480
            Width           =   3000
         End
         Begin VB.CheckBox ckcCWSendEmails 
            Caption         =   "Send Emails"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   285
            TabIndex        =   16
            Top             =   240
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox ckcCSIWeb 
            Caption         =   "CSI Electronic Affidavit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   135
            TabIndex        =   15
            Top             =   0
            Value           =   1  'Checked
            Width           =   2175
         End
      End
      Begin VB.Frame frcMarketron 
         Height          =   780
         Left            =   3600
         TabIndex        =   10
         Top             =   120
         Width           =   2355
         Begin VB.CheckBox ckcMOutput 
            Caption         =   "Generate File"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   330
            TabIndex        =   13
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox ckcMOutput 
            Caption         =   "Send to Marketron"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   330
            TabIndex        =   12
            Top             =   240
            Value           =   1  'Checked
            Width           =   1965
         End
         Begin VB.CheckBox ckcMarketron 
            Caption         =   "Marketron"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   135
            TabIndex        =   11
            Top             =   0
            Value           =   1  'Checked
            Width           =   1215
         End
      End
      Begin VB.Frame frcUnivision 
         Height          =   855
         Left            =   0
         TabIndex        =   1
         Top             =   810
         Width           =   7485
         Begin VB.Frame frmUVeh 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   105
            TabIndex        =   5
            Top             =   210
            Width           =   3780
            Begin VB.OptionButton rbcUSpots 
               Caption         =   "All Spots"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   825
               TabIndex        =   7
               Top             =   45
               Width           =   1155
            End
            Begin VB.OptionButton rbcUSpots 
               Caption         =   "Spot Changes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   2025
               TabIndex        =   6
               Top             =   45
               Width           =   1605
            End
            Begin VB.Label lacUUnivision 
               Caption         =   "Export"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   8
               Top             =   45
               Width           =   660
            End
         End
         Begin VB.TextBox txtUFile 
            Height          =   240
            Left            =   1020
            TabIndex        =   4
            Top             =   510
            Width           =   4800
         End
         Begin VB.CommandButton cmcUBrowse 
            Caption         =   "&Browse..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5940
            TabIndex        =   3
            Top             =   480
            Width           =   1170
         End
         Begin VB.CheckBox ckcUnivision 
            Caption         =   "Univision"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   2
            Top             =   0
            Value           =   1  'Checked
            Width           =   1005
         End
         Begin VB.Label lacUFile 
            Caption         =   "Export File"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   9
            Top             =   510
            Width           =   975
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   75
      Top             =   2055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "AffExportCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of AffExportCriteria.ctl on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  imPopReqd                     imSelectedIndex               imComboBoxIndex           *
'*  imBypassSetting               imTypeRowNo                                             *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mPopulate                                                                             *
'*                                                                                        *
'* Public Property Procedures (Marked)                                                    *
'*  Enabled(Let)                  Verify(Get)                                             *
'*                                                                                        *
'* Public User-Defined Events (Marked)                                                    *
'*  SetSave                                                                               *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AffExportCriteria.ctl
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text

'Private rst_Shtt As ADODB.Recordset
'Private rst_artt As ADODB.Recordset

'Event SetSave(ilStatus As Integer) 'VBC NR
'Event ContactFocus()
'Event PhoneChanged(slPhone As String)
'Event FaxChanged(slFax As String)
Event SetChgFlag()
Event IDCChg(ilValue As Integer)
Event WGenerate(ilIndex As Integer, ilValue As Integer)
Event ISCIChg()

'Program library dates Field Areas
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imLbcArrowSetting As Integer
Dim imBypassFocus As Integer
Dim imDoubleClickName As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imStartMode As Integer
Private imFromArrow As Integer
Private imOkToSetChgFlag As Integer
'9435
Private bmIsMarketronWebVendor As Boolean

Private rst_Eht As ADODB.Recordset
Private rst_Evt As ADODB.Recordset
Private rst_Ect As ADODB.Recordset

'ISCI
Private imEmbeddedAllowed As Integer

Dim smNowDate As String
Dim lmNowDate As Long



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
    Dim llHeight As Long
    Dim llWidth As Long

    Screen.MousePointer = vbHourglass
    imOkToSetChgFlag = False
    frcIDC.BorderStyle = 0
    frcWegener.BorderStyle = 0
    frcISCI.BorderStyle = 0
    frcXDS.BorderStyle = 0
    frcStarGuide.BorderStyle = 0
    frcCnC.BorderStyle = 0
    frcISCIXRef.BorderStyle = 0
    frcRCS.BorderStyle = 0
    frciPump.BorderStyle = 0
    frcStationLog.BorderStyle = 0
        
    frcStationLog.Left = 0  'lacName.Left
    frcStationLog.Top = 0   'lacName.Top + lacName.Height + 120
        
    llHeight = 0
    llWidth = 0
    frcIDC.Move frcStationLog.Left, frcStationLog.Top
    mCheckSize frcIDC, llHeight, llWidth
    frcWegener.Move frcStationLog.Left, frcStationLog.Top
    mCheckSize frcWegener, llHeight, llWidth
    frcISCI.Move frcStationLog.Left, frcStationLog.Top
    mCheckSize frcISCI, llHeight, llWidth
    frcXDS.Move frcStationLog.Left, frcStationLog.Top
    mCheckSize frcXDS, llHeight, llWidth
    frcStarGuide.Move frcStationLog.Left, frcStationLog.Top
    mCheckSize frcStarGuide, llHeight, llWidth
    frcCnC.Move frcStationLog.Left, frcStationLog.Top
    mCheckSize frcCnC, llHeight, llWidth
    frcISCIXRef.Move frcStationLog.Left, frcStationLog.Top
    mCheckSize frcISCIXRef, llHeight, llWidth
    frcRCS.Move frcStationLog.Left, frcStationLog.Top
    mCheckSize frcRCS, llHeight, llWidth
    'ISCI
    frcIUniqueISCI.Move frcIISCIByBreak.Left, frcIISCIByBreak.Top
    '6393 ipump
    frciPump.Move frcStationLog.Left, frcStationLog.Top
    mCheckSize frciPump, llHeight, llWidth
    mShowFrame
    Screen.MousePointer = vbDefault
    '9184
    frcCumulus.Visible = False
    frcMarketron.Top = frcCSIWeb.Top
    '9435
    If gIsWebVendor(22) Then
        bmIsMarketronWebVendor = True
    End If
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
'*  flTextHeight                  ilLoop                        ilCol                     *
'*                                                                                        *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    Dim llRow As Long
    'flTextHeight = pbcDates.TextHeight("1") - 35

    'grdContact.Move 180, 120, Width - pbcArrow.Width - 120
    'grdContact.Height = Height - grdContact.Top - 120
    'grdContact.Redraw = False
    'pbcSTab.Move -100, -100
    'pbcTab.Move -100, -100
    'pbcClickFocus.Move -100, -100
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

'
'   mTerminate
'   Where:
'


    Screen.MousePointer = vbDefault
End Sub



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

    'RaiseEvent SetSave(True)

End Sub


Public Sub Action(ilType As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilIndex As Integer
    Select Case ilType
        Case 1  'Show Frame
            mShowFrame
        Case 2  'Init function
            'Test if unloading control
            ilRet = 0
            On Error GoTo UserControlErr:
            mInit
        Case 3  'Populate
            mSetCtrls
            'dan M 12/18/13 6548
            If Trim$(sgExportTypeChar) = "X" Then
                mSetXProvider
            End If
            imOkToSetChgFlag = True
        Case 4  'Clear
            Screen.MousePointer = vbDefault
        Case 5  'Save
            mSaveCtrls
        Case 6  'Custom
            'DoEvents
            mCustomSetCtrl
            DoEvents
            'mSetFonts
    End Select
    Exit Sub
UserControlErr:
    ilRet = 1
    Resume Next
End Sub
'See page 512 in Component tools centers to learn about Enabling the UserControl
'In addition to defining the Get and Let, you need to set the Property ID:
'Menu item Tools->Procedure Attributes.  In Name box select Enabled.
'Click on Advance. In Procedure ID, select Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal blValue As Boolean)
    UserControl.Enabled = blValue
    PropertyChanged "Enabled" 'VBC NR
End Property
Public Property Let Embedded(ilEmbedded As Integer) 'VBC NR
    imEmbeddedAllowed = ilEmbedded 'VBC NR
    PropertyChanged "Embedded" 'VBC NR
End Property 'VBC NR

Public Property Let ckcGenType(ilIndex As Integer, ilValue As Integer) 'VBC NR
    ckcDGenType(ilIndex) = ilValue
    PropertyChanged "ckcGenType" 'VBC NR
End Property 'VBC NR

Public Property Get ckcGenType(ilIndex As Integer) As Integer 'VBC NR
    ckcGenType = ckcDGenType(ilIndex).Value
End Property 'VBC NR
Public Property Get rbcUSpot(ilIndex As Integer) As Integer
    rbcUSpot = rbcUSpots(ilIndex).Value
End Property 'VBC NR
Public Property Get edcUFile() As String
    edcUFile = Trim$(txtUFile.Text)
End Property 'VBC NR
Public Property Get MOutput(ilIndex As Integer, slProp As String) As Integer
    MOutput = ckcMOutput(ilIndex).Value
End Property 'VBC NR
Public Property Let MOutput(ilIndex As Integer, slProp As String, ilValue As Integer)
    If slProp = "E" Then
        ckcMOutput(ilIndex).Enabled = ilValue
    ElseIf slProp = "V" Then
        ckcMOutput(ilIndex).Value = ilValue
    End If
    PropertyChanged "MOutput" 'VBC NR
End Property
Public Property Get iPOutput(ilIndex As Integer) As Integer
    iPOutput = ckciPOutput(ilIndex).Value
End Property 'VBC NR
Public Property Let iPOutput(ilIndex As Integer, ilValue As Integer)
    ckciPOutput(ilIndex).Value = ilValue
End Property
Public Property Get GetCtrlBottom() As Long
    Dim llPos As Long
    Select Case igExportTypeNumber
        Case 1    'Marketron
        Case 2    'Univision
        Case 3    'Web
            llPos = ckcCWRemoveISCI.Top + ckcCWRemoveISCI.Height
        Case 4, 5    'RCS 4 or 5
        Case 6    'Clearance and Compensation
        Case 7    'IDC
        Case 8    'ISCI
        Case 9    'ISCI Cross Reference
        Case 10    'StarGuide
        Case 11    'Wegener
        Case 12    'X-Digital
            llPos = cmcXBrowse.Top + cmcXBrowse.Height
        Case 13     'Wegener iPump
    End Select
    GetCtrlBottom = llPos
End Property 'VBC NR
Public Property Get CSendEMails() As Integer
    CSendEMails = ckcCWSendEmails.Value
End Property 'VBC NR
Public Property Get CRemoveISCI() As Integer
    CRemoveISCI = ckcCWRemoveISCI.Value
End Property 'VBC NR
'9184
'Public Property Get CUSendEMails() As Integer
'    CUSendEMails = ckcCUSendEmails.Value
'End Property 'VBC NR
Public Property Get DGenType(ilIndex As Integer) As Integer
    DGenType = ckcDGenType(ilIndex).Value
End Property 'VBC NR
Public Property Let DGenType(ilIndex As Integer, ilValue As Integer) 'VBC NR
    ckcDGenType(ilIndex).Value = ilValue
    PropertyChanged "DGenType" 'VBC NR
End Property 'VBC NR
Public Property Get DExportToPath() As String
    DExportToPath = Trim$(edcDExportToPath.Text)
End Property 'VBC NR
Public Property Get XGenType(ilIndex As Integer, slProp As String) As Integer
    XGenType = optXGenType(ilIndex).Value
End Property 'VBC NR
Public Property Let XGenType(ilIndex As Integer, slProp As String, ilValue As Integer)
    If slProp = "E" Then
        optXGenType(ilIndex).Enabled = ilValue
    ElseIf slProp = "V" Then
        optXGenType(ilIndex).Value = ilValue
    End If
    PropertyChanged "XGenType" 'VBC NR
End Property
Public Property Get XSpots(ilIndex As Integer) As Integer
    XSpots = rbcXSpots(ilIndex).Value
End Property 'VBC NR
Public Property Get XProvider(ilIndex As Integer) As Integer
    XProvider = rbcXProvider(ilIndex).Value
End Property 'VBC NR

Public Property Get XExportType(ilIndex As Integer, slProp As String) As Integer
    XExportType = ckcXExportType(ilIndex).Value
End Property 'VBC NR
Public Property Let XExportType(ilIndex As Integer, slProp As String, ilValue As Integer) 'VBC NR
    'The Get and Let must have the same parameter except the last one
    'The last parameter must be on the right side of the calling setup
    'udcCriteria.XExportType(0, "V") = vbUnchecked
    If slProp = "E" Then
        ckcXExportType(ilIndex).Enabled = ilValue
    ElseIf slProp = "V" Then
        ckcXExportType(ilIndex).Value = ilValue
    End If
    PropertyChanged "XExportType" 'VBC NR
End Property 'VBC NR
Public Property Get XExportToPath() As String
    XExportToPath = Trim$(edcXExportToPath.Text)
End Property 'VBC NR
Public Property Let XReExportVisible(blValue As Boolean)
    If blValue Then
        ckcXReexport.Visible = True
    Else
        ckcXReexport.Visible = False
    End If
End Property
Public Property Get XReExport() As Boolean
    Dim blRet As Boolean
    If ckcXReexport.Value = vbChecked Then
        blRet = True
    Else
        blRet = False
    End If
    XReExport = blRet
End Property
Public Property Get WGenerate(ilIndex As Integer) As Integer
    WGenerate = ckcWGenerate(ilIndex).Value
End Property 'VBC NR
Public Property Get WGenCSV() As Integer
    WGenCSV = ckcWGenCSV.Value
End Property 'VBC NR
Public Property Get edcWStationInfo() As String
    edcWStationInfo = Trim$(txtWStationInfo.Text)
End Property 'VBC NR
Public Property Get WExportToPath() As String
    WExportToPath = Trim$(edcWExportToPath.Text)
End Property 'VBC NR
Public Property Get WRunLetter() As String
    WRunLetter = Trim$(txtWRunLetter.Text)
End Property 'VBC NR
Public Property Get IExportToPath() As String
    IExportToPath = Trim$(edcIExportToPath.Text)
End Property 'VBC NR
Public Property Get iExportType(ilIndex As Integer) As Integer
    iExportType = rbcIExportType(ilIndex).Value
End Property 'VBC NR
Public Property Get ISpots(ilIndex As Integer) As Integer
    ISpots = rbcISpots(ilIndex).Value
End Property 'VBC NR
Public Property Get IIncludeCommands(ilIndex As Integer) As Integer
    IIncludeCommands = ckcIIncludeCommands(ilIndex).Value
End Property 'VBC NR
Public Property Get IUniqueBy(ilIndex As Integer) As Integer
    IUniqueBy = rbcIUniqueBy(ilIndex).Value
End Property 'VBC NR
Public Property Get INoDaysResend() As String
    INoDaysResend = Trim$(edcINoDaysResend.Text)
End Property 'VBC NR
Public Property Get IVeh() As Integer
    IVeh = frcIVeh.Visible
End Property 'VBC NR
Public Property Get RIncludeGeneric() As Integer
    RIncludeGeneric = ckcRIncludeGeneric.Value
End Property 'VBC NR
Public Property Get RExportToPath() As String
    RExportToPath = Trim$(edcRExportToPath.Text)
End Property 'VBC NR
'7459
Public Property Get RPrefix(ilIndex As Integer) As Integer
    RPrefix = rbcRPrefix(ilIndex).Value
End Property
Public Property Let RPrefixVisible(blValue As Boolean)
    If blValue Then
        frcRPrefix.Visible = True
    Else
        frcRPrefix.Visible = False
    End If
End Property
Public Property Get RCSZone(ilIndex As Integer) As Integer
    RCSZone = chkRCSZone(ilIndex).Value
End Property 'VBC NR
Public Property Get RCSExportToPath() As String
    RCSExportToPath = Trim$(edcRCSExportToPath.Text)
End Property 'VBC NR

Public Property Get SSpots(ilIndex As Integer) As Integer
    SSpots = rbcSSpots(ilIndex).Value
End Property 'VBC NR
Public Property Get SRunLetter() As String
    SRunLetter = Trim$(txtSRunLetter.Text)
End Property 'VBC NR
Public Property Get SExportToPath() As String
    SExportToPath = Trim$(edcSExportToPath.Text)
End Property 'VBC NR
Public Property Get edcCFile() As String
    edcCFile = Trim$(txtCFile.Text)
End Property 'VBC NR


Private Sub chkZone_Click(Index As Integer)
    mSetChgFlag
End Sub

Private Sub ckcCSendEmails_Click()
    mSetChgFlag
End Sub

Private Sub ckcCSIWeb_Click()
    If ckcCSIWeb.Value = vbChecked Then
        ckcCWSendEmails.Enabled = True
        ckcCWRemoveISCI.Enabled = True
    Else
        ckcCWSendEmails.Value = vbUnchecked
        ckcCWSendEmails.Enabled = False
        ckcCWRemoveISCI.Enabled = False
    End If
    mSetChgFlag
End Sub

Private Sub ckcCumulus_Click()
    If ckcCumulus.Value = vbChecked Then
        ckcCUSendEmails.Enabled = True
    Else
        ckcCUSendEmails.Value = vbUnchecked
        ckcCUSendEmails.Enabled = False
    End If
    mSetChgFlag
End Sub

Private Sub ckcCUSendEmails_Click()
    mSetChgFlag
End Sub

Private Sub ckcDGenType_Click(Index As Integer)
    Dim ilValue As Integer
    If Index = 0 Then
        ilValue = ckcDGenType(Index).Value
        RaiseEvent IDCChg(ilValue)
    End If
    mSetChgFlag
End Sub

Private Sub ckcIIncludeCommands_Click(Index As Integer)
    RaiseEvent ISCIChg
    mSetChgFlag
End Sub

Private Sub ckcIncludeGeneric_Click()
    mSetChgFlag
End Sub

Private Sub ckciPOutput_Click(Index As Integer)
     mSetChgFlag
End Sub

Private Sub ckcMarketron_Click()
    If ckcMarketron.Value = vbChecked Then '  And Not bmIsMarketronWebVendor Then
        ckcMOutput(0).Enabled = True
        ckcMOutput(1).Enabled = True
    Else
        ckcMOutput(0).Value = vbUnchecked
        ckcMOutput(0).Enabled = False
        ckcMOutput(1).Value = vbUnchecked
        ckcMOutput(1).Enabled = False
    End If
    mSetChgFlag
End Sub

Private Sub ckcMOutput_Click(Index As Integer)
    mSetChgFlag
End Sub

Private Sub ckcUnivision_Click()
    If ckcUnivision.Value = vbChecked Then
        rbcUSpots(0).Enabled = True
        rbcUSpots(1).Enabled = True
        txtUFile.Enabled = True
        cmcUBrowse.Enabled = True
        txtUFile.BorderStyle = 1
    Else
        rbcUSpots(0).Enabled = False
        rbcUSpots(1).Enabled = False
        txtUFile.Enabled = False
        cmcUBrowse.Enabled = False
    End If
    mSetChgFlag
End Sub

Private Sub ckcWGenCSV_Click()
    '7342
    If ckcWGenCSV.Value = vbChecked Then
        edcWExportToPath.Enabled = True
        cmcWExportTo.Enabled = True
        lacWExportTo.Enabled = True
    Else
        edcWExportToPath.Enabled = False
        cmcWExportTo.Enabled = False
        lacWExportTo.Enabled = False
    End If
    mSetChgFlag
End Sub

Private Sub ckcWGenerate_Click(Index As Integer)
    If Index = 0 Then
        If ckcWGenerate(0) = vbChecked Then
            lacWStationInfo.Enabled = True
            txtWStationInfo.Enabled = True
            cmcWBrowse.Enabled = True
        Else
            lacWStationInfo.Enabled = False
            txtWStationInfo.Enabled = False
            cmcWBrowse.Enabled = False
        End If
    ElseIf Index = 1 Then
        RaiseEvent WGenerate(Index, ckcWGenerate(1).Value)
        If ckcWGenerate(1).Value = vbChecked Then
            'lacStartDate.Enabled = True
            'txtDate.Enabled = True
            'lacDays.Enabled = True
            'txtNumberDays.Enabled = True
            ckcWGenCSV.Enabled = True
            'lbcVehicles.Enabled = True
        Else
            'lacStartDate.Enabled = False
            'txtDate.Enabled = False
            'lacDays.Enabled = False
            'txtNumberDays.Enabled = False
            ckcWGenCSV.Enabled = False
            'lbcVehicles.Enabled = False
        End If
    End If
    mSetChgFlag
End Sub

Private Sub ckcXExportType_Click(Index As Integer)
    mSetChgFlag
End Sub

Private Sub cmcCBrowse_Click()
    'Clearance and Compensation Export
    Dim slCurDir As String
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    ' Set filters
    'CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    '"(*.txt)|*.txt|CSV Files (*.csv)|*.csv"
    CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    ' Display the Open dialog box
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    CommonDialog1.fileName = "CnCSpots.txt"
    CommonDialog1.ShowSave
    ' Display name of selected file
    txtCFile.Text = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub cmcDBrowse_Click()
    'Export IDC
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    'Dim slCurDir As String
    'slCurDir = CurDir
    'igPathType = 0
    'sgGetPath = edcDExportToPath.Text
    'frmGetPath.Show vbModal
    'If igGetPath = 0 Then
    '    edcDExportToPath.Text = sgGetPath
    'End If
    'ChDir slCurDir
    gBrowseForFolder CommonDialog1, edcDExportToPath
    Exit Sub
End Sub

Private Sub cmcIBrowse_Click()
    'ISCI Export
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    'Dim slCurDir As String
    'slCurDir = CurDir
    'igPathType = 0
    'sgGetPath = edcIExportToPath.Text
    'frmGetPath.Show vbModal
    'If igGetPath = 0 Then
    '    edcIExportToPath.Text = sgGetPath
    'End If
    'ChDir slCurDir
    gBrowseForFolder CommonDialog1, edcIExportToPath
    Exit Sub
End Sub

Private Sub cmcRBrowse_Click()
    'ISCI Cross Reference Export
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    'Dim slCurDir As String
    'slCurDir = CurDir
    'igPathType = 0
    'sgGetPath = edcRExportToPath.Text
    'frmGetPath.Show vbModal
    'If igGetPath = 0 Then
    '    edcRExportToPath.Text = sgGetPath
    'End If
    gBrowseForFolder CommonDialog1, edcRExportToPath
    Exit Sub
End Sub

Private Sub cmcRCSBrowse_Click()
    'RCS4 & RCS5 export
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    'Dim slCurDir As String
    'slCurDir = CurDir
    'igPathType = 0
    'sgGetPath = edcRCSExportToPath.Text
    'frmGetPath.Show vbModal
    'If igGetPath = 0 Then
    '    edcRCSExportToPath.Text = sgGetPath
    'End If
    'ChDir slCurDir
    gBrowseForFolder CommonDialog1, edcRCSExportToPath
    Exit Sub
End Sub

Private Sub cmcSBrowse_Click()
    'Starguide Export
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    'Dim slCurDir As String
    'slCurDir = CurDir
    'igPathType = 0
    'sgGetPath = edcSExportToPath.Text
    'frmGetPath.Show vbModal
    'If igGetPath = 0 Then
    '    edcSExportToPath.Text = sgGetPath
    'End If
    'ChDir slCurDir
    gBrowseForFolder CommonDialog1, edcSExportToPath
    Exit Sub
End Sub

Private Sub cmcUBrowse_Click()
    'Univision Scheduled Station Spots Export
    Dim slCurDir As String
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    ' Set filters
    'CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    '"(*.txt)|*.txt|CSV Files (*.csv)|*.csv"
    CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    ' Display the Open dialog box
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    CommonDialog1.fileName = "MktSpots.txt"
    CommonDialog1.ShowSave
    ' Display name of selected file

    txtUFile.Text = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmcWBrowse_Click()
    'Export wegener - compel
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    'Dim slCurDir As String
    'slCurDir = CurDir
    'igPathType = 0
    'sgGetPath = txtWStationInfo.Text
    'frmGetPath.Show vbModal
    'If igGetPath = 0 Then
    '    txtWStationInfo.Text = sgGetPath
    'End If
    'ChDir slCurDir
    gBrowseForFolder CommonDialog1, txtWStationInfo
    Exit Sub
End Sub

Private Sub cmcWExportTo_Click()
    'Export wegener - compel
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    'Dim slCurDir As String
    'slCurDir = CurDir
    'igPathType = 0
    'sgGetPath = edcWExportToPath.Text
    'frmGetPath.Show vbModal
    'If igGetPath = 0 Then
    '    edcWExportToPath.Text = sgGetPath
    'End If
    'ChDir slCurDir
    gBrowseForFolder CommonDialog1, edcWExportToPath
    Exit Sub
End Sub

Private Sub cmcXBrowse_Click()
    'X-DIGITAL Export
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    'Dim slCurDir As String
    'slCurDir = CurDir
    'igPathType = 0
    'sgGetPath = edcXExportToPath.Text
    'frmGetPath.Show vbModal
    'If igGetPath = 0 Then
    '    edcXExportToPath.Text = sgGetPath
    'End If
    'ChDir slCurDir
    gBrowseForFolder CommonDialog1, edcXExportToPath
    Exit Sub
End Sub

Private Sub edcDExportToPath_Change()
    mSetChgFlag
End Sub

Private Sub edcIExportToPath_Change()
    RaiseEvent ISCIChg
    mSetChgFlag
End Sub

Private Sub edcRCSExportToPath_Change()
    mSetChgFlag
End Sub

Private Sub edcRExportToPath_Change()
    mSetChgFlag
End Sub

Private Sub edcSExportToPath_Change()
    mSetChgFlag
End Sub

Private Sub edcWExportToPath_Change()
    mSetChgFlag
End Sub

Private Sub edcXExportToPath_Change()
    mSetChgFlag
End Sub

Private Sub optXGenType_Click(Index As Integer)
    mSetChgFlag
End Sub

Private Sub rbcIExportType_Click(Index As Integer)
    If rbcIExportType(Index).Value Then
        If Index = 0 Then
            frcIUniqueISCI.Visible = True
            If imEmbeddedAllowed Then
                rbcIUniqueBy(0).Enabled = True
            End If
            rbcIUniqueBy(1).Enabled = True
            frcIISCIByBreak.Visible = False
        Else
            frcIISCIByBreak.Visible = True
            frcIUniqueISCI.Visible = False
            ckcIIncludeCommands(0).Enabled = True
        End If
    Else
        If Index = 0 Then
            rbcIUniqueBy(0).Enabled = False
            rbcIUniqueBy(1).Enabled = False
        Else
            ckcIIncludeCommands(0).Enabled = False
        End If
    End If
    If rbcIExportType(Index).Value Then
        If Index = 0 Then
            frcIUniqueISCI.Visible = True
            frcIISCIByBreak.Visible = False
            If imEmbeddedAllowed Then
                rbcIUniqueBy(0).Enabled = True
            End If
            rbcIUniqueBy(1).Enabled = True
            frcIVeh.Enabled = True
            rbcISpots(0).Enabled = True
            rbcISpots(2).Enabled = True
            If (rbcISpots(0).Value = True) Then
                lacINoDaysResend.Enabled = True
                edcINoDaysResend.Enabled = True
            Else
                lacINoDaysResend.Enabled = False
                edcINoDaysResend.Enabled = False
            End If
        Else
            frcIISCIByBreak.Visible = True
            frcIUniqueISCI.Visible = False
            ckcIIncludeCommands(0).Enabled = True
        End If
    Else
        If Index = 0 Then
            rbcIUniqueBy(0).Enabled = False
            rbcIUniqueBy(1).Enabled = False
            frcIVeh.Enabled = False
            rbcISpots(0).Enabled = False
            rbcISpots(2).Enabled = False
            lacINoDaysResend.Enabled = False
            edcINoDaysResend.Enabled = False
        Else
            ckcIIncludeCommands(0).Enabled = False
        End If
    End If
    RaiseEvent ISCIChg

    mSetChgFlag
End Sub

Private Sub rbcIUniqueBy_Click(Index As Integer)
    RaiseEvent ISCIChg
    
    mSetChgFlag
End Sub

Private Sub rbcSSpots_Click(Index As Integer)
    mSetChgFlag
End Sub

Private Sub rbcUSpots_Click(Index As Integer)
    mSetChgFlag
End Sub

Private Sub rbcXProvider_Click(Index As Integer)
    '6765
    mSetChgFlag
End Sub

Private Sub rbcXSpots_Click(Index As Integer)
    mSetChgFlag
End Sub

Private Sub txtFile_Change()
    mSetChgFlag
End Sub

Private Sub txtSRunLetter_Change()
    mSetChgFlag
End Sub

Private Sub txtWRunLetter_Change()
    mSetChgFlag
End Sub

Private Sub txtWStationInfo_Change()
    mSetChgFlag
End Sub

Private Sub txtUFile_Change()
    mSetChgFlag
End Sub

Private Sub UserControl_Click()
    'pbcClickFocus.SetFocus
End Sub

Private Sub UserControl_GotFocus()
    'RaiseEvent ContactFocus
End Sub

Private Sub UserControl_Initialize()
    mSetFonts
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Form_MouseUp Button, Shift, X, Y
End Sub

Public Property Get Verify() As Integer 'VBC NR
    If mTestFields() Then 'VBC NR
        Verify = True 'VBC NR
    Else 'VBC NR
        Verify = False
    End If
End Property


Public Sub mSetFonts()
    Dim Ctrl As control
    Dim ilFontSize As Integer
    Dim ilColorFontSize As Integer
    Dim ilBold As Integer
    Dim ilChg As Integer
    Dim slStr As String
    Dim slFontName As String
    
    
    'On Error Resume Next
    ilFontSize = 14
    ilBold = True
    ilColorFontSize = 10
    slFontName = "Arial"
    If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
        ilFontSize = 8
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
        ilFontSize = 8
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 768 Then
        ilFontSize = 10
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 800 Then
        ilFontSize = 10
        ilBold = True
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 1024 Then
        ilFontSize = 12
        ilBold = True
    End If
    For Each Ctrl In UserControl.Controls
        If TypeOf Ctrl Is MSHFlexGrid Then
            Ctrl.Font.Name = slFontName
            Ctrl.FontFixed.Name = slFontName
            Ctrl.Font.Size = ilFontSize
            Ctrl.FontFixed.Size = ilFontSize
            Ctrl.Font.Bold = ilBold
            Ctrl.FontFixed.Bold = ilBold
        ElseIf TypeOf Ctrl Is TabStrip Then
            Ctrl.Font.Name = slFontName
            Ctrl.Font.Size = ilFontSize
            Ctrl.Font.Bold = ilBold
        ''ElseIf TypeOf Ctrl Is Resize Then
        ''ElseIf TypeOf Ctrl Is Timer Then
        ''ElseIf TypeOf Ctrl Is Image Then
        ''ElseIf TypeOf Ctrl Is ImageList Then
        ''ElseIf TypeOf Ctrl Is CommonDialog Then
        ''ElseIf TypeOf Ctrl Is AffExportCriteria Then
        ''ElseIf TypeOf Ctrl Is AffCommentGrid Then
        ''ElseIf TypeOf Ctrl Is AffContactGrid Then
        ''Else
        'ElseIf (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is ListBox) Or (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is PictureBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Label) Then
        ElseIf (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is ListBox) _
               Or (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is PictureBox) _
               Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Label) _
               Or (TypeOf Ctrl Is CSI_Calendar) Or (TypeOf Ctrl Is CSI_Calendar_UP) Or (TypeOf Ctrl Is CSI_ComboBoxList) Or (TypeOf Ctrl Is CSI_DayPicker) Then
            ilChg = 0
            If TypeOf Ctrl Is CommandButton Then
               ilChg = 1
            Else
                If (Ctrl.ForeColor = vbBlack) Or (Ctrl.ForeColor = &H80000008) Or (Ctrl.ForeColor = &H80000012) Or (Ctrl.ForeColor = &H8000000F) Then
                    ilChg = 1
                Else
                    ilChg = 2
                End If
            End If
            slStr = Ctrl.Name
            If (InStr(1, slStr, "Arrow", vbTextCompare) > 0) Or ((InStr(1, slStr, "Dropdown", vbTextCompare) > 0) And (TypeOf Ctrl Is CommandButton)) Then
                ilChg = 0
            End If
            If ilChg = 1 Then
                Ctrl.FontName = slFontName
                Ctrl.FontSize = ilFontSize
                Ctrl.FontBold = ilBold
            ElseIf ilChg = 2 Then
                Ctrl.FontName = slFontName
                Ctrl.FontSize = ilColorFontSize
                Ctrl.FontBold = False
            End If
        End If
    Next Ctrl
End Sub

Private Sub UserControl_Resize()
    mSetFonts
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    'rst_Shtt.Close
    'rst_artt.Close
End Sub

Private Sub mShowFrame()
    Dim slName As String
    
    frcStationLog.Visible = False
    frcIDC.Visible = False
    frcISCI.Visible = False
    frcWegener.Visible = False
    frcXDS.Visible = False
    frcStarGuide.Visible = False
    frcCnC.Visible = False
    frcISCIXRef.Visible = False
    frcRCS.Visible = False
    frciPump.Visible = False
    Select Case Trim$(sgExportTypeChar)
        Case "A"    '"Aff Logs"
            'frcStationLog.Height = 2 * frcStationLog.Height
            txtUFile.Height = lacUFile.Height
            frcStationLog.Visible = True
            ckcCSIWeb.Top = 0
            ckcCumulus.Top = 0
            ckcMarketron.Top = 0
            ckcUnivision.Top = 0
'            '9435
            If bmIsMarketronWebVendor Then
                ckcMarketron.Enabled = False
                frcMarketron.Enabled = False
                ckcMOutput(0).Enabled = False
                ckcMOutput(1).Enabled = False
            End If
        Case "D"    '"IDC"
            frcIDC.Visible = True
        Case "I"    '"ISCI"
            frcISCI.Visible = True
        Case "W"    '"Wegener"
            frcWegener.Visible = True
        Case "X"    '"X-Digital"
            frcXDS.Visible = True
        Case "S"    '"StarGd"
            frcStarGuide.Visible = True
        Case "C"    '"C & C"
            frcCnC.Visible = True
        Case "R"    '"ISCI C/R"
            mSetRPrefix
            frcISCIXRef.Visible = True
        Case "4"    '"RCS 4"
            frcRCS.Visible = True
        Case "5"    '"RCS 5"
            frcRCS.Visible = True
        Case "P"
            frciPump.Visible = True
    End Select
            
End Sub

Private Sub mCheckSize(frcCtrl As Frame, llHeight As Long, llWidth As Long)

    If frcCtrl.Height > llHeight Then
        llHeight = frcCtrl.Height
    End If
    If frcCtrl.Width > llWidth Then
        llWidth = frcCtrl.Width
    End If
End Sub

Private Sub mSetCtrls()
    Dim llNext As Long
    Dim slControlName As String
    Dim Ctrl As control
    Dim ilRet As Integer
    Dim slIndex As String
    Dim blVisible As Boolean
    
    On Error GoTo mSetCtrlsErr
    llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEct
    If llNext <> -1 Then
        Do While llNext <> -1
            For Each Ctrl In UserControl.Controls
                blVisible = Ctrl.Visible
                If blVisible Then
                    On Error GoTo IndexErr
                    ilRet = 0
                    slIndex = Ctrl.Index
                    If ilRet = 1 Then
                        slIndex = ""
                    End If
                    On Error GoTo mSetCtrlsErr
                    slControlName = Trim$(Ctrl.Name) & slIndex
                    If Trim$(tgEctInfo(llNext).sFieldName) = slControlName Then
                        If TypeOf Ctrl Is ListBox Then
                             'Build the Parameter Information record
                        ElseIf TypeOf Ctrl Is TextBox Then
                             'Build the Parameter Information record
                             Ctrl.Text = Trim$(tgEctInfo(llNext).sFieldString)
                        ElseIf TypeOf Ctrl Is ComboBox Then
                             'Build the Parameter Information record
                        ElseIf TypeOf Ctrl Is OptionButton Then
                            Ctrl.Value = tgEctInfo(llNext).lFieldValue
                        ElseIf TypeOf Ctrl Is CheckBox Then
                            Ctrl.Value = tgEctInfo(llNext).lFieldValue
                        End If
                        Exit For
                    End If
                End If
            Next Ctrl
            llNext = tgEctInfo(llNext).lNextEct
        Loop
    Else
        For Each Ctrl In UserControl.Controls
            blVisible = Ctrl.Visible
            If blVisible Then
                If TypeOf Ctrl Is ListBox Then
                     'Build the Parameter Information record
                ElseIf TypeOf Ctrl Is TextBox Then
                     'Build the Parameter Information record
                     Ctrl.Text = ""
                ElseIf TypeOf Ctrl Is ComboBox Then
                     'Build the Parameter Information record
                ElseIf TypeOf Ctrl Is OptionButton Then
                    Ctrl.Value = False
                ElseIf TypeOf Ctrl Is CheckBox Then
                    Ctrl.Value = vbUnchecked
                End If
            End If
        Next Ctrl
        'Set Defaults
        mSetDefaults
    End If
    ckcCSIWeb_Click
    ckcCumulus_Click
    '9435
    If Not bmIsMarketronWebVendor Then
        ckcMarketron_Click
    Else
        frcMarketron.Enabled = False
        ckcMarketron.Value = vbUnchecked
        ckcMarketron.Enabled = False
        ckcMOutput(0).Enabled = False
        ckcMOutput(1).Enabled = False
        ckcMOutput(0).Value = vbUnchecked
        ckcMOutput(1).Value = vbUnchecked
    End If
    ckcUnivision_Click
    Exit Sub
IndexErr:
    ilRet = 1
    Resume Next
mSetCtrlsErr:
    blVisible = False
    Resume Next
End Sub

Private Sub mSaveCtrls()
    Dim llNext As Long
    Dim slControlName As String
    Dim Ctrl As control
    Dim blSetValue As Boolean
    Dim slFieldString As String
    Dim llFieldValue As Long
    Dim slFieldType As String
    Dim llEctInfo As Long
    Dim ilRet As Integer
    Dim slIndex As String
    Dim blVisible As Boolean
    Dim llSvNext As Long
    Dim llCheck As Long
    
    On Error GoTo mSaveCtrlErr
    llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEct
    Do While llNext <> -1
        llSvNext = tgEctInfo(llNext).lNextEct
        tgEctInfo(llNext).lNextEct = -9999
        llNext = llSvNext
    Loop
    tgEhtInfo(lgExportEhtInfoIndex).lFirstEct = -1
    For Each Ctrl In UserControl.Controls
        blVisible = Ctrl.Visible
        If blVisible Then
            blSetValue = False
            On Error GoTo IndexErr
            ilRet = 0
            slIndex = Ctrl.Index
            If ilRet = 1 Then
                slIndex = ""
            End If
            On Error GoTo mSaveCtrlErr
            slControlName = Trim$(Ctrl.Name) & slIndex
            If TypeOf Ctrl Is ListBox Then
            ElseIf TypeOf Ctrl Is TextBox Then
                blSetValue = True
                slFieldString = Ctrl.Text
                llFieldValue = 0
                slFieldType = "S"
            ElseIf TypeOf Ctrl Is ComboBox Then
            ElseIf TypeOf Ctrl Is OptionButton Then
                blSetValue = True
                slFieldString = ""
                llFieldValue = Ctrl.Value
                slFieldType = "L"
            ElseIf TypeOf Ctrl Is CheckBox Then
                blSetValue = True
                slFieldString = ""
                llFieldValue = Ctrl.Value
                slFieldType = "L"
            End If
            If blSetValue Then
                If tgEhtInfo(lgExportEhtInfoIndex).lFirstEct = -1 Then
                    llNext = -1
                    'llEctInfo = UBound(tgEctInfo)
                    'For llCheck = 0 To UBound(tgEctInfo) - 1 Step 1
                    '    If tgEctInfo(llCheck).lNextEct = -9999 Then
                    '        llEctInfo = llCheck
                    '        Exit For
                    '    End If
                    'Next llCheck
                    'tgEhtInfo(lgExportEhtInfoIndex).lFirstEct = llEctInfo
                    'tgEctInfo(llEctInfo).lNextEct = -1
                Else
                    llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEct
                    'llEctInfo = UBound(tgEctInfo)
                    'For llCheck = 0 To UBound(tgEctInfo) - 1 Step 1
                    '    If tgEctInfo(llCheck).lNextEct = -9999 Then
                    '        llEctInfo = llCheck
                    '        Exit For
                    '    End If
                    'Next llCheck
                    'tgEhtInfo(lgExportEhtInfoIndex).lFirstEct = llEctInfo
                    'tgEctInfo(llEctInfo).lNextEct = llNext
                End If
                llEctInfo = UBound(tgEctInfo)
                For llCheck = 0 To UBound(tgEctInfo) - 1 Step 1
                    If tgEctInfo(llCheck).lNextEct = -9999 Then
                        llEctInfo = llCheck
                        Exit For
                    End If
                Next llCheck
                tgEhtInfo(lgExportEhtInfoIndex).lFirstEct = llEctInfo
                tgEctInfo(llEctInfo).sLogType = mSetLogType(slControlName)
                tgEctInfo(llEctInfo).sFieldType = slFieldType
                tgEctInfo(llEctInfo).sFieldName = slControlName
                tgEctInfo(llEctInfo).lFieldValue = llFieldValue
                tgEctInfo(llEctInfo).sFieldString = slFieldString
                If tgEctInfo(llEctInfo).lNextEct = -9999 Then
                    tgEctInfo(llEctInfo).lNextEct = llNext
                Else
                    tgEctInfo(llEctInfo).lNextEct = llNext
                    ReDim Preserve tgEctInfo(0 To UBound(tgEctInfo) + 1) As ECTINFO
                End If
            End If
        End If
    Next Ctrl
    Exit Sub
IndexErr:
    ilRet = 1
    Resume Next
mSaveCtrlErr:
    blVisible = False
    Resume Next

End Sub

Private Sub mSetChgFlag()
    If imOkToSetChgFlag Then
        RaiseEvent SetChgFlag
    End If
End Sub

Private Sub mCustomSetCtrl()
    ckcCSIWeb.Value = vbUnchecked
    ckcCumulus.Value = vbUnchecked
    ckcMarketron.Value = vbUnchecked
    ckcUnivision.Value = vbUnchecked
    '6393
    frciPump.Visible = False
    Select Case igExportTypeNumber
        Case 1    'Marketron
            'FrmExportMarketron.Show Modal
            frcStationLog.BorderStyle = 0
            frcMarketron.BorderStyle = 0
            frcStationLog.Move 0, 0
            frcMarketron.Move 0, 0
            ckcMOutput(0).Left = 0
            ckcMOutput(1).Left = 0
            DoEvents
            frcCSIWeb.Visible = False
            frcCumulus.Visible = False
            frcUnivision.Visible = False
            frcStationLog.Visible = True
            ckcMarketron.Visible = False
            frcMarketron.Visible = True
            ckcCumulus.Visible = False
            frcCumulus.Visible = False
            mSetFonts
            '9435 not really doing anything
            If Not bmIsMarketronWebVendor Then
                ckcMarketron.Value = vbChecked
                mCustomMLoad
            End If
        Case 2    'Univision
            'frmExportSchdSpot.Show Modal
            frcStationLog.BorderStyle = 0
            frcUnivision.BorderStyle = 0
            frcStationLog.Move 0, 0
            frcUnivision.Move 0, 0
            lacUFile.Left = 0
            frmUVeh.Left = 0
            frcCSIWeb.Visible = False
            frcCumulus.Visible = False
            frcMarketron.Visible = False
            frcStationLog.Visible = True
            ckcUnivision.Visible = False
            ckcUnivision.Value = vbChecked
            frcUnivision.Visible = True
            ckcCumulus.Visible = False
            frcCumulus.Visible = False
           mSetFonts
            mCustomULoad
            txtUFile.BorderStyle = 0
            DoEvents
            txtUFile.BorderStyle = 1
        Case 3    'Web
            'frmWebExportSchdSpot.Show Modal
            If (sgWebExport = "W") Or (sgWebExport = "B") Then
                frcStationLog.BorderStyle = 0
                frcCSIWeb.BorderStyle = 0
                frcStationLog.Move 0, 0
                frcCSIWeb.Move 0, 0
                ckcCWSendEmails.Left = 0
                ckcCWSendEmails.Top = 0
                ckcCWRemoveISCI.Left = ckcCWSendEmails.Left
                ckcCWRemoveISCI.Top = ckcCWSendEmails.Top + ckcCWSendEmails.Height
                ckcCWRemoveISCI.Width = 2 * ckcCWRemoveISCI.Width
                DoEvents
                frcCumulus.Visible = False
                frcMarketron.Visible = False
                frcUnivision.Visible = False
                frcStationLog.Visible = True
                ckcCSIWeb.Visible = False
                ckcCSIWeb.Value = vbChecked
                frcCSIWeb.Visible = True
                ckcCumulus.Visible = False
                frcCumulus.Visible = False
                mSetFonts
                mCustomCSILoad
            Else
                frcStationLog.BorderStyle = 0
                frcCumulus.BorderStyle = 0
                frcStationLog.Move 0, 0
                frcCumulus.Move 0, 0
                ckcCUSendEmails.Left = 0
                ckcCUSendEmails.Top = 0
                DoEvents
                frcCSIWeb.Visible = False
                frcMarketron.Visible = False
                frcUnivision.Visible = False
                frcStationLog.Visible = True
                ckcCumulus.Visible = False
                frcCumulus.Visible = False
                mSetFonts
                mCustomCULoad
            End If
        Case 6    'Clearance and Compensation
            'frmExportCnCSpots.Show Modal
            frcCnC.BorderStyle = 0
            frcCnC.Move 0, 0
            DoEvents
            frcCnC.Visible = True
            mSetFonts
            mCustomCLoad
        Case 7    'IDC
            'FrmExportIDC.Show Modal
            frcIDC.BorderStyle = 0
            frcIDC.Move 0, 0
            ckcDGenType(0).Top = 0
            ckcDGenType(1).Top = 0
            ckcDGenType(2).Top = 0
            ckcDGenType(3).Top = 0
            lacDExportTo.Top = ckcDGenType(0).Top + (3 * ckcDGenType(0).Height / 2)
            edcDExportToPath.Top = ckcDGenType(0).Top + (3 * ckcDGenType(0).Height / 2)
            edcDExportToPath.Height = ckcDGenType(0).Height
            cmcDBrowse.Top = ckcDGenType(0).Top + (3 * ckcDGenType(0).Height / 2)
            cmcDBrowse.Height = ckcDGenType(0).Height
            DoEvents
            frcIDC.Visible = True
            mSetFonts
            mCustomDLoad
        Case 8    'ISCI
            'frmExportISCI.Show Modal
            frcISCI.BorderStyle = 0
            frcISCI.Move 0, 0
            DoEvents
            frcISCI.Visible = True
            frcIExportType.Top = 120  'edcIExportToPath.Top + edcIExportToPath.Height + 120
            lacIExportType.Top = frcIExportType.Top
            frcIISCIByBreak.Top = lacIExportType.Top + lacIExportType.Height + 60
            ckcIIncludeCommands(1).Top = frcIISCIByBreak.Top + frcIISCIByBreak.Height
            lacIEMail.Top = ckcIIncludeCommands(1).Top
            'share location
            frcIUniqueISCI.Move frcIISCIByBreak.Left, frcIISCIByBreak.Top
            edcIExportToPath.Top = ckcIIncludeCommands(1).Top + (3 * ckcIIncludeCommands(1).Height / 2)
            lacIExportToPath.Top = edcIExportToPath.Top
            cmcIBrowse.Top = edcIExportToPath.Top
            mSetFonts
            mCustomILoad
            If imEmbeddedAllowed Then
                rbcIUniqueBy(0).Enabled = True
            Else
                rbcIUniqueBy(0).Enabled = False
            End If
        Case 9    'ISCI Cross Reference
            'frmExportISCIXRef.Show Modal
            'TTP 10457 - ISCI Cross Reference Export
            edcRExportToPath.Text = sgISCIxRefExportPath
            frcISCIXRef.BorderStyle = 0
            frcISCIXRef.Move 0, 0
            ckcRIncludeGeneric.Move 0, 120
            lacRExportTo.Top = ckcRIncludeGeneric.Top + (3 * ckcRIncludeGeneric.Height / 2)
            edcRExportToPath.Top = ckcRIncludeGeneric.Top + (3 * ckcRIncludeGeneric.Height / 2)
            edcRExportToPath.Height = ckcRIncludeGeneric.Height
            cmcRBrowse.Top = ckcRIncludeGeneric.Top + (3 * ckcRIncludeGeneric.Height / 2)
            cmcRBrowse.Height = ckcRIncludeGeneric.Height
            mSetFonts
            '7459
             mSetRPrefix
            DoEvents
            frcISCIXRef.Visible = True
            mCustomRLoad
        Case 4, 5    'RCS 4 or 5
            'igRCSExportBy = 4
            'igRCSExportBy = 5
            'frmExportRCS.Show Modal
            frcRCS.BorderStyle = 0
            frcRCS.Move 0, 0
            lacRCSZone.Top = 90
            chkRCSZone(0).Top = 0
            chkRCSZone(1).Top = chkRCSZone(0)
            chkRCSZone(2).Top = chkRCSZone(0)
            chkRCSZone(3).Top = chkRCSZone(0)
            lacRCSExportTo.Top = chkRCSZone(0).Top + (3 * chkRCSZone(0).Height / 2)
            edcRCSExportToPath.Top = lacRCSExportTo.Top
            edcRCSExportToPath.Height = lacRCSExportTo.Height
            cmcRCSBrowse.Top = lacRCSExportTo.Top
            cmcRCSBrowse.Height = lacRCSExportTo.Height
            DoEvents
            frcRCS.Visible = True
            mSetFonts
            mCustomRCSLoad Trim$(Str$(igExportTypeNumber))
        Case 10    'StarGuide
            'frmExportStarGuide.Show Modal
            frcStarGuide.BorderStyle = 0
            frcStarGuide.Move 0, 0
            lacSExportTo.Top = txtSRunLetter.Top + (3 * txtSRunLetter.Height / 2)
            edcSExportToPath.Top = lacSExportTo.Top
            edcSExportToPath.Height = txtSRunLetter.Height
            cmcSBrowse.Top = lacSExportTo.Top
            cmcSBrowse.Height = txtSRunLetter.Height
            DoEvents
            frcStarGuide.Visible = True
            mSetFonts
            mCustomSLoad
        Case 11    'Wegener
            'FrmExportWegener.Show Modal
            frcWegener.BorderStyle = 0
            frcWegener.Move 0, 0
            DoEvents
            frcWegener.Visible = True
            mSetFonts
            lacWFileInfo.FontSize = 7
            mCustomWLoad
        Case 12    'X-Digital
            'FrmExportXDigital.Show Modal
            frcXDS.BorderStyle = 0
            frcXDS.Move 0, 0
            'ckcXExportType(0).Top = 0
            'ckcXExportType(1).Top = 0
            frcXSuppressNotices.Top = ckcXExportType(0).Top + ckcXExportType(0).Height
            frmXVeh.Top = frcXSuppressNotices.Top + frcXSuppressNotices.Height
            lacXExportTo.Top = frmXVeh.Top + (3 * ckcXExportType(0).Height / 2)
            lacXSpots.Top = rbcXSpots(0).Top
            lacXProvider.Top = rbcXProvider(0).Top
            edcXExportToPath.Top = lacXExportTo.Top
            edcXExportToPath.Height = ckcXExportType(0).Height
            cmcXBrowse.Top = lacXExportTo.Top
            cmcXBrowse.Height = ckcXExportType(0).Height
            ckcXReexport.Top = frcXSuppressNotices.Top
            DoEvents
            frcXDS.Visible = True
            mSetFonts
            mCustomXLoad
        Case 13     'Wegener iPump
'            frcStationLog.BorderStyle = 0
'            frcMarketron.BorderStyle = 0
'            frcStationLog.Move 0, 0
'            frcMarketron.Move 0, 0
'            ckcMOutput(0).Left = 0
'            ckcMOutput(1).Left = 0
            DoEvents
            frcCSIWeb.Visible = False
            frcCumulus.Visible = False
            frcUnivision.Visible = False
            frcStationLog.Visible = False
            ckcMarketron.Visible = False
            ckcCumulus.Visible = False
            frcCumulus.Visible = False
            frcMarketron.Visible = False
            frciPump.BorderStyle = 0
            frciPump.Move 0, 0
            ckciPOutput(0).Left = 0
            frciPump.Visible = True
            mSetFonts
            mCustomPLoad
    End Select

End Sub

Private Sub mCustomULoad()
    Dim ilCount As Integer
    Dim blActive As Boolean
    Dim slTUFile As String
    Dim ilTUSpot0 As Integer
    Dim ilTUSpot1 As Integer
    Dim slUFile As String
    Dim ilUSpot0 As Integer
    Dim ilUSpot1 As Integer
    Dim llEhtCode As Long

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'A' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        blActive = False
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            If Trim$(rst_Ect!ectLogType) = "U" Then
                If Trim$(rst_Ect!ectFieldName) = "ckcUnivision" Then
                    If rst_Ect!ectFieldValue = vbChecked Then
                        blActive = True
                        ilCount = ilCount + 1
                        If ilCount > 1 Then
                            Exit Do
                        End If
                        llEhtCode = rst_Ect!ectehtCode
                    End If
                End If
                If Trim$(rst_Ect!ectFieldName) = "txtUFile" Then
                    If Trim$(rst_Ect!ectFieldString) <> "" Then
                        slTUFile = Trim$(rst_Ect!ectFieldString)
                    End If
                End If
                If Trim$(rst_Ect!ectFieldName) = "rbcUSpots0" Then
                    ilTUSpot0 = rst_Ect!ectFieldValue
                End If
                If Trim$(rst_Ect!ectFieldName) = "rbcUSpots1" Then
                    ilTUSpot1 = rst_Ect!ectFieldValue
                End If
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        ElseIf blActive Then
            slUFile = slTUFile
            ilUSpot0 = ilTUSpot0
            ilUSpot1 = ilTUSpot1
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        txtUFile.Text = slUFile
        rbcUSpots(0).Value = ilUSpot0
        rbcUSpots(1).Value = ilUSpot1
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    End If
    If Trim$(txtCFile.Text) = "" Then
        txtUFile.Text = sgExportDirectory & "MktSpots.txt"
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom Univision Load"
    Resume Next
    Exit Sub
End Sub

Private Sub mCustomMLoad()
    Dim ilCount As Integer
    Dim blActive As Boolean
    Dim ilTMOutput0 As Integer
    Dim ilTMOutput1 As Integer
    Dim ilMOutput0 As Integer
    Dim ilMOutput1 As Integer
    Dim llEhtCode As Long

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'A' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        blActive = False
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            If Trim$(rst_Ect!ectLogType) = "M" Then
                If Trim$(rst_Ect!ectFieldName) = "ckcMarketron" Then
                    If rst_Ect!ectFieldValue = vbChecked Then
                        '9435
                        If Not gIsWebVendor(22) Then
                            blActive = True
                            ilCount = ilCount + 1
                            If ilCount > 1 Then
                                Exit Do
                            End If
                            llEhtCode = rst_Ect!ectehtCode
                        End If
                    End If
                End If
                If Trim$(rst_Ect!ectFieldName) = "ckcMOutput0" Then
                    ilTMOutput0 = rst_Ect!ectFieldValue
                End If
                If Trim$(rst_Ect!ectFieldName) = "ckcMOutput1" Then
                    ilTMOutput1 = rst_Ect!ectFieldValue
                End If
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        ElseIf blActive Then
            ilMOutput0 = ilTMOutput0
            ilMOutput1 = ilTMOutput1
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        ckcMOutput(0).Value = ilMOutput0
        ckcMOutput(1).Value = ilMOutput1
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    Else
        ckcMOutput(0).Value = vbChecked
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom Marketron Load"
    Resume Next
    Exit Sub
End Sub

Private Sub mCustomCSILoad()
    Dim ilCount As Integer
    Dim blActive As Boolean
    Dim ilTCSendEmails As Integer
    Dim ilTCRemoveISCI As Integer
    Dim ilCSendEmails As Integer
    Dim ilCRemoveISCI As Integer
    Dim llEhtCode As Long

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'A' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        blActive = False
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            If Trim$(rst_Ect!ectLogType) = "W" Then
                If Trim$(rst_Ect!ectFieldName) = "ckcCSIWeb" Then
                    If rst_Ect!ectFieldValue = vbChecked Then
                        ilCount = ilCount + 1
                        If ilCount > 1 Then
                            Exit Do
                        End If
                        llEhtCode = rst_Ect!ectehtCode
                    End If
                End If
                If Trim$(rst_Ect!ectFieldName) = "ckcCSendEmails" Then
                    ilTCSendEmails = rst_Ect!ectFieldValue
                End If
                If Trim$(rst_Ect!ectFieldName) = "ckcCRemoveISCI" Then
                    ilTCRemoveISCI = rst_Ect!ectFieldValue
                End If
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        ElseIf blActive Then
            ilCSendEmails = ilTCSendEmails
            ilCRemoveISCI = ilTCRemoveISCI
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        ckcCWSendEmails.Value = ilCSendEmails
        ckcCWRemoveISCI.Value = ilCRemoveISCI
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    Else
        ckcCWSendEmails.Value = vbChecked
        ckcCWRemoveISCI.Value = vbUnchecked
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom CSI Web Load"
    Resume Next
    Exit Sub
End Sub

Private Sub mCustomCULoad()
    Dim ilCount As Integer
    Dim blActive As Boolean
    Dim ilTCUSendEmails As Integer
    Dim ilCUSendEmails As Integer
    Dim llEhtCode As Long

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'A' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        blActive = False
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            If Trim$(rst_Ect!ectLogType) = "C" Then
                If Trim$(rst_Ect!ectFieldName) = "ckcCumulus" Then
                    If rst_Ect!ectFieldValue = vbChecked Then
                        blActive = True
                        ilCount = ilCount + 1
                        If ilCount > 1 Then
                            Exit Do
                        End If
                        llEhtCode = rst_Ect!ectehtCode
                    End If
                End If
                If Trim$(rst_Ect!ectFieldName) = "ckcCUSendEmails" Then
                    ilTCUSendEmails = rst_Ect!ectFieldValue
                End If
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        ElseIf blActive Then
            ilCUSendEmails = ilTCUSendEmails
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        ckcCUSendEmails.Value = ilCUSendEmails
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    Else
        ckcCUSendEmails.Value = vbChecked
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom Cumulus Load"
    Resume Next
    Exit Sub
End Sub

Private Sub mCustomDLoad()
    Dim ilCount As Integer
    Dim ilDGenType0 As Integer
    Dim ilDGenType2 As Integer
    Dim ilDGenType3 As Integer
    Dim slDExportTo As String
    Dim llEhtCode As Long

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'D' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        ilCount = ilCount + 1
        If ilCount > 1 Then
            Exit Do
        End If
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            llEhtCode = rst_Ect!ectehtCode
            If Trim$(rst_Ect!ectFieldName) = "ckcDGenType0" Then
                ilDGenType0 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "ckcDGenType2" Then
                ilDGenType2 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "ckcDGenType3" Then
                ilDGenType3 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "edcDExportToPath" Then
                slDExportTo = Trim$(rst_Ect!ectFieldString)
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        ckcDGenType(0).Value = ilDGenType0
        ckcDGenType(2).Value = ilDGenType2
        ckcDGenType(3).Value = ilDGenType3
        edcDExportToPath.Text = slDExportTo
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    Else
        ckcDGenType(0).Value = vbUnchecked
        ckcDGenType(2).Value = vbUnchecked
        ckcDGenType(3).Value = vbUnchecked
    End If
    If Trim$(edcDExportToPath.Text) = "" Then
        edcDExportToPath.Text = sgExportDirectory
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom IDC Load"
    Resume Next
    Exit Sub
End Sub


Private Sub mCustomXLoad()
    Dim ilCount As Integer
    Dim ilXExportType0 As Integer
    Dim ilXExportType1 As Integer
    Dim ilXGenType0 As Integer
    Dim ilXGenType1 As Integer
    Dim ilXSpots0 As Integer
    Dim ilXSpots1 As Integer
    Dim ilXProvider0 As Integer
    Dim ilXProvider1 As Integer
    Dim slXExportTo As String
    Dim llEhtCode As Long

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'X' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        ilCount = ilCount + 1
        If ilCount > 1 Then
            Exit Do
        End If
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            llEhtCode = rst_Ect!ectehtCode
            If Trim$(rst_Ect!ectFieldName) = "ckcXExportType0" Then
                ilXExportType0 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "ckcXExportType1" Then
                ilXExportType1 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "optXGenType0" Then
                ilXGenType0 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "optXGenType1" Then
                ilXGenType1 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "rbcXSpots0" Then
                ilXSpots0 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "rbcXSpots1" Then
                ilXSpots1 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "rbcXProvider0" Then
                ilXProvider0 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "rbcXProvider1" Then
                ilXProvider1 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "edcXExportToPath" Then
                slXExportTo = Trim$(rst_Ect!ectFieldString)
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        ckcXExportType(0).Value = ilXExportType0
        ckcXExportType(1).Value = ilXExportType1
        optXGenType(0).Value = ilXGenType0
        optXGenType(1).Value = ilXGenType1
        rbcXSpots(0).Value = ilXSpots0
        rbcXSpots(1).Value = ilXSpots1
        rbcXProvider(0).Value = ilXProvider0
        rbcXProvider(1).Value = ilXProvider1
        edcXExportToPath.Text = slXExportTo
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    Else
        ckcXExportType(0).Value = vbUnchecked
        ckcXExportType(1).Value = vbUnchecked
        optXGenType(0).Value = True
        rbcXSpots(1).Value = True
    End If
    If Trim$(edcXExportToPath.Text) = "" Then
        edcXExportToPath.Text = sgExportDirectory
    End If
    mSetXProvider
    If (ilCount <> 1) Or (UBound(sgXDSSection) <= 1) Then
        rbcXProvider(0).Value = True
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom XDS Load"
    Resume Next
    Exit Sub
End Sub

Private Sub mCustomWLoad()
    Dim ilCount As Integer
    Dim slWStationInfo As String
    Dim ilWGenerate0 As Integer
    Dim ilWGenerate1 As Integer
    Dim ilWGenCSV As Integer
    Dim slWExportTo As String
    Dim llEhtCode As Long
    Dim slWRunLetter As String

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'W' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        ilCount = ilCount + 1
        If ilCount > 1 Then
            Exit Do
        End If
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            llEhtCode = rst_Ect!ectehtCode
            If Trim$(rst_Ect!ectFieldName) = "txtWStationInfo" Then
                If Trim$(rst_Ect!ectFieldString) <> "" Then
                    slWStationInfo = Trim$(rst_Ect!ectFieldString)
                End If
            End If
            If Trim$(rst_Ect!ectFieldName) = "ckcWGenerate0" Then
                ilWGenerate0 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "ckcWGenerate1" Then
                ilWGenerate1 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "ckcWGenCSV" Then
                ilWGenCSV = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "edcWExportToPath" Then
                slWExportTo = Trim$(rst_Ect!ectFieldString)
            End If
            If Trim$(rst_Ect!ectFieldName) = "txtWRunLetter" Then
                slWRunLetter = Trim$(rst_Ect!ectFieldString)
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        txtWStationInfo.Text = slWStationInfo
        ckcWGenerate(0).Value = ilWGenerate0
        ckcWGenerate(1).Value = ilWGenerate1
        ckcWGenCSV.Value = ilWGenCSV
        edcWExportToPath.Text = slWExportTo
        txtWRunLetter.Text = slWRunLetter
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    Else
        If Len(sgImportDirectory) > 0 Then
            txtWStationInfo.Text = Left$(sgImportDirectory, Len(sgImportDirectory) - 1)
        Else
            txtWStationInfo.Text = ""
        End If
        ckcWGenerate(0).Value = vbChecked
        ckcWGenerate(1).Value = vbChecked
        ckcWGenCSV.Value = vbUnchecked
        txtWRunLetter.Text = slWRunLetter
    End If
    If Trim$(edcWExportToPath.Text) = "" Then
        edcWExportToPath.Text = sgExportDirectory
    End If
    If Trim$(txtWRunLetter.Text) = "" Then
        txtWRunLetter.Text = "A"
    End If
    '7342
    ckcWGenCSV_Click
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom Wegener Load"
    Resume Next
    Exit Sub
End Sub

Private Sub mCustomILoad()
    Dim ilCount As Integer
    Dim slIExportToPath As String
    Dim ilIExportType0 As Integer
    Dim ilIExportType1 As Integer
    Dim ilIUniqueBy0 As Integer
    Dim ilIUniqueBy1 As Integer
    Dim ilIIncludeCommands0 As Integer
    Dim ilIIncludeCommands1 As Integer
    Dim slINoDaysResend As String
    Dim llEhtCode As Long

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'I' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        ilCount = ilCount + 1
        If ilCount > 1 Then
            Exit Do
        End If
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            llEhtCode = rst_Ect!ectehtCode
            If Trim$(rst_Ect!ectFieldName) = "edcIExportToPath" Then
                If Trim$(rst_Ect!ectFieldString) <> "" Then
                    slIExportToPath = Trim$(rst_Ect!ectFieldString)
                End If
            End If
            If Trim$(rst_Ect!ectFieldName) = "rbcIExportType0" Then
                ilIExportType0 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "rbcIExportType1" Then
                ilIExportType1 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "rbcIUniqueBy0" Then
                ilIUniqueBy0 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "rbcIUniqueBy1" Then
                ilIUniqueBy1 = rst_Ect!ectFieldValue
            End If
            
            If Trim$(rst_Ect!ectFieldName) = "ckcIIncludeCommands0" Then
                ilIIncludeCommands0 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "ckcIIncludeCommands1" Then
                ilIIncludeCommands1 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "edcINoDaysResend" Then
                If Trim$(rst_Ect!ectFieldString) <> "" Then
                    slINoDaysResend = Trim$(rst_Ect!ectFieldString)
                Else
                    slINoDaysResend = "30"
                End If
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        edcIExportToPath.Text = slIExportToPath
        rbcIExportType(0).Value = ilIExportType0
        rbcIExportType(1).Value = ilIExportType1
        rbcIUniqueBy(0).Value = ilIUniqueBy0
        rbcIUniqueBy(1).Value = ilIUniqueBy1
        ckcIIncludeCommands(0).Value = ilIIncludeCommands0
        ckcIIncludeCommands(1).Value = ilIIncludeCommands1
        edcINoDaysResend.Text = slINoDaysResend
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    Else
        If Len(sgExportDirectory) > 0 Then
            edcIExportToPath.Text = sgExportDirectory
        Else
            edcIExportToPath.Text = ""
        End If
        ckcIIncludeCommands(0).Value = vbChecked
        rbcIUniqueBy(1).Value = True
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom ISCI Load"
    Resume Next
    Exit Sub
End Sub

Private Sub mCustomRLoad()
    Dim ilCount As Integer
    Dim ilRIncludeGeneric As Integer
    Dim slRExportTo As String
    Dim llEhtCode As Long
    Dim ilPrefix As Integer
    
    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'R' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        ilCount = ilCount + 1
        If ilCount > 1 Then
            Exit Do
        End If
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            llEhtCode = rst_Ect!ectehtCode
            If Trim$(rst_Ect!ectFieldName) = "ckcRIncludeGeneric" Then
                ilRIncludeGeneric = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "edcRExportToPath" Then
                slRExportTo = Trim$(rst_Ect!ectFieldString)
            End If
            '7459.  I only check this value because the other is default
             If Trim$(rst_Ect!ectFieldName) = "rbcRPrefix1" Then
                ilPrefix = rst_Ect!ectFieldValue
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        ckcRIncludeGeneric.Value = ilRIncludeGeneric
        edcRExportToPath.Text = slRExportTo
        '7459
        If ilPrefix = -1 Then
            rbcRPrefix(1).Value = True
        Else
            rbcRPrefix(0).Value = True
        End If
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    End If
    If Trim$(edcRExportToPath.Text) = "" Then
        edcRExportToPath.Text = sgExportDirectory
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom Cross Reference Load"
    Resume Next
    Exit Sub
End Sub

Private Sub mCustomRCSLoad(slExportType As String)
    Dim ilCount As Integer
    Dim ilRCSZone0 As Integer
    Dim ilRCSZone1 As Integer
    Dim ilRCSZone2 As Integer
    Dim ilRCSZone3 As Integer
    Dim slRCSExportTo As String
    Dim llEhtCode As Long
    Dim ilLoop As Integer
    Dim ilZone As Integer

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = '" & slExportType & "' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        ilCount = ilCount + 1
        If ilCount > 1 Then
            Exit Do
        End If
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            llEhtCode = rst_Ect!ectehtCode
            If Trim$(rst_Ect!ectFieldName) = "chkRCSZone0" Then
                ilRCSZone0 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "chkRCSZone1" Then
                ilRCSZone1 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "chkRCSZone2" Then
                ilRCSZone2 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "chkRCSZone3" Then
                ilRCSZone3 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "edcRCSExportToPath" Then
                slRCSExportTo = Trim$(rst_Ect!ectFieldString)
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        chkRCSZone(0).Value = ilRCSZone0
        chkRCSZone(1).Value = ilRCSZone1
        chkRCSZone(2).Value = ilRCSZone2
        chkRCSZone(3).Value = ilRCSZone3
        edcRCSExportToPath.Text = slRCSExportTo
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    Else
        chkRCSZone(0).Value = vbUnchecked
        chkRCSZone(1).Value = vbUnchecked
        chkRCSZone(2).Value = vbUnchecked
        chkRCSZone(3).Value = vbUnchecked
        For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            For ilZone = LBound(tgVehicleInfo(ilLoop).sZone) To UBound(tgVehicleInfo(ilLoop).sZone) Step 1
                Select Case Left$(tgVehicleInfo(ilLoop).sZone(ilZone), 1)
                    Case "E"
                        If tgVehicleInfo(ilLoop).sFed(ilZone) = "*" Then
                            chkRCSZone(0).Value = vbChecked
                        End If
                    Case "C"
                        If tgVehicleInfo(ilLoop).sFed(ilZone) = "*" Then
                            chkRCSZone(1).Value = vbChecked
                        End If
                    Case "M"
                        If tgVehicleInfo(ilLoop).sFed(ilZone) = "*" Then
                            chkRCSZone(2).Value = vbChecked
                        End If
                    Case "P"
                        If tgVehicleInfo(ilLoop).sFed(ilZone) = "*" Then
                            chkRCSZone(3).Value = vbChecked
                        End If
                End Select
            Next ilZone
        Next ilLoop
    End If
    If Trim$(edcRCSExportToPath.Text) = "" Then
        edcRCSExportToPath.Text = sgExportDirectory
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom RCS Load"
    Resume Next
    Exit Sub
End Sub

Private Sub mCustomSLoad()
    Dim ilCount As Integer
    Dim ilSSpots0 As Integer
    Dim ilSSpots1 As Integer
    Dim slSRunLetter As String
    Dim slSExportTo As String
    Dim llEhtCode As Long

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'S' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        ilCount = ilCount + 1
        If ilCount > 1 Then
            Exit Do
        End If
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            llEhtCode = rst_Ect!ectehtCode
            If Trim$(rst_Ect!ectFieldName) = "rbcSSpots0" Then
                ilSSpots0 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "rbcSSpots1" Then
                ilSSpots1 = rst_Ect!ectFieldValue
            End If
            If Trim$(rst_Ect!ectFieldName) = "txtSRunLetter" Then
                slSRunLetter = Trim$(rst_Ect!ectFieldString)
            End If
            If Trim$(rst_Ect!ectFieldName) = "edcSExportToPath" Then
                slSExportTo = Trim$(rst_Ect!ectFieldString)
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        rbcSSpots(0).Value = ilSSpots0
        rbcSSpots(1).Value = ilSSpots1
        txtSRunLetter.Text = slSRunLetter
        edcSExportToPath.Text = slSExportTo
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    End If
    If Trim$(edcSExportToPath.Text) = "" Then
        edcSExportToPath.Text = sgExportDirectory
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom StarGuide Load"
    Resume Next
    Exit Sub
End Sub

Private Sub mCustomCLoad()
    Dim ilCount As Integer
    Dim slCFile As String
    Dim llEhtCode As Long

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'C' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        ilCount = ilCount + 1
        If ilCount > 1 Then
            Exit Do
        End If
        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
        Set rst_Ect = gSQLSelectCall(SQLQuery)
        Do While Not rst_Ect.EOF
            llEhtCode = rst_Ect!ectehtCode
            If Trim$(rst_Ect!ectFieldName) = "txtCFile" Then
                slCFile = Trim$(rst_Ect!ectFieldString)
            End If
            rst_Ect.MoveNext
        Loop
        If ilCount > 1 Then
            Exit Do
        End If
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        txtCFile.Text = slCFile
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    End If
    If Trim$(txtCFile.Text) = "" Then
        txtCFile.Text = sgExportDirectory & "CnCSpots.txt"
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom CnC Load"
    Resume Next
    Exit Sub
End Sub

Private Sub mCustomPLoad()
'wegener iPump
' trying to load the vehicles that are associated with this export.  Must be only one header; if there are 2, we
' won't know which vehicles to load.
    Dim ilCount As Integer
    Dim llEhtCode As Long

    On Error GoTo ErrHandler
    ilCount = 0
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    If igExportSource = 1 Then
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtExportType = 'P' And ehtSubType = 'S'"
    Else
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    End If
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    Do While Not rst_Eht.EOF
        ilCount = ilCount + 1
        If ilCount > 1 Then
            Exit Do
        End If
        llEhtCode = rst_Eht!ehtCode
        rst_Eht.MoveNext
    Loop
    If ilCount = 1 Then
        SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & llEhtCode
        Set rst_Evt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Evt.EOF
            tgEvtInfo(UBound(tgEvtInfo)).iVefCode = rst_Evt!evtVefCode
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
            rst_Evt.MoveNext
        Loop
    End If
    Exit Sub
ErrHandler:
    gHandleError "AffErrorLog.txt", "Custom Wegener iPump Load"
    Resume Next
    Exit Sub
End Sub

Private Function mTestFields() As Integer
    mTestFields = False
    Select Case Trim$(sgExportTypeChar)
        Case "A"    '"Aff Logs"
            If (ckcCSIWeb.Value = vbUnchecked) And (ckcCumulus.Value = vbUnchecked) And (ckcMarketron.Value = vbUnchecked) And (ckcUnivision.Value = vbUnchecked) Then
                Exit Function
            End If
            If ckcMarketron.Value = vbChecked Then
                If (ckcMOutput(0).Value = vbUnchecked) And (ckcMOutput(1).Value = vbUnchecked) Then
                    Exit Function
                End If
            End If
            If ckcUnivision.Value = vbChecked Then
                If Trim$(txtUFile.Text) = "" Then
                    Exit Function
                End If
                If (rbcUSpots(0).Value = False) And (rbcUSpots(1).Value = False) Then
                    Exit Function
                End If
            End If
        Case "D"    '"IDC"
            If (ckcDGenType(0).Value = vbUnchecked) And (ckcDGenType(2).Value = vbUnchecked) And (ckcDGenType(3).Value = vbUnchecked) Then
                Exit Function
            End If
            If Trim$(edcDExportToPath.Text) = "" Then
                Exit Function
            End If
        Case "I"    '"ISCI"
            If Trim$(edcIExportToPath.Text) = "" Then
                Exit Function
            End If
            If (rbcIExportType(0).Value = False) And (rbcIExportType(1).Value = False) Then
                Exit Function
            End If
            If rbcIExportType(0).Value Then
                If (rbcIUniqueBy(0).Value = False) And (rbcIUniqueBy(1).Value = False) Then
                    Exit Function
                End If
            End If
        Case "W"    '"Wegener"
            If Trim$(txtWStationInfo.Text) = "" Then
                Exit Function
            End If
            If (ckcWGenerate(1).Value = vbUnchecked) And (ckcWGenCSV.Value = vbUnchecked) Then
                Exit Function
            End If
        Case "X"    '"X-Digital"
            If (ckcXExportType(0).Value = vbUnchecked) And (ckcXExportType(1).Value = vbUnchecked) Then
                Exit Function
            End If
            If (optXGenType(0).Value = False) And (optXGenType(1).Value = False) Then
                Exit Function
            End If
            If (rbcXSpots(0).Value = False) And (rbcXSpots(1).Value = False) Then
                Exit Function
            End If
            '7314 added 'must be visible'
            If (rbcXProvider(0).Value = False) And (rbcXProvider(1).Value = False) And rbcXProvider(0).Visible = True Then
                Exit Function
            End If
        Case "S"    '"StarGuide"
            If (rbcSSpots(0).Value = False) And (rbcSSpots(1).Value = False) Then
                Exit Function
            End If
            If Trim$(txtSRunLetter.Text) = "" Then
                Exit Function
            End If
        Case "C"    '"C & C"
            If Trim$(txtCFile.Text) = "" Then
                Exit Function
            End If
        Case "R"    '"ISCI C/R"
        Case "4"    '"RCS 4"
            If (chkRCSZone(0).Value = vbUnchecked) And (chkRCSZone(1).Value = vbUnchecked) And (chkRCSZone(2).Value = vbUnchecked) And (chkRCSZone(3).Value = vbUnchecked) Then
                Exit Function
            End If
        Case "5"    '"RCS 5"
            If (chkRCSZone(0).Value = vbUnchecked) And (chkRCSZone(1).Value = vbUnchecked) And (chkRCSZone(2).Value = vbUnchecked) And (chkRCSZone(3).Value = vbUnchecked) Then
                Exit Function
            End If
    End Select
    mTestFields = True
End Function

Private Function mSetLogType(slCtrlName As String) As String
    mSetLogType = ""
    If (slCtrlName = "ckcCSIWeb") Or (slCtrlName = "ckcCSendEmails") Or (slCtrlName = "ckcCRemoveISCI") Then
        mSetLogType = "W"
    End If
    If (slCtrlName = "ckcCumulus") Or (slCtrlName = "ckcCUSendEmails") Then
        mSetLogType = "C"
    End If
    If (slCtrlName = "ckcMarketron") Or (slCtrlName = "ckcMOutput0") Or (slCtrlName = "ckcMOutput1") Then
        mSetLogType = "M"
    End If
    If (slCtrlName = "ckcUnivision") Or (slCtrlName = "txtUFiles") Or (slCtrlName = "rbcUSpots0") Or (slCtrlName = "rbcUSpots1") Then
        mSetLogType = "U"
    End If
End Function

Private Sub mSetDefaults()
    Dim ilLoop As Integer
    Dim ilZone As Integer
    Select Case Trim$(sgExportTypeChar)
        Case "A"    '"Aff Logs"
            'Counterpoint Affidavit
            ckcCWSendEmails.Value = vbChecked
            ckcCWRemoveISCI.Value = vbUnchecked
            'Cumulus
            ckcCUSendEmails.Value = vbChecked
            'Marketron
            ckcMOutput(0).Value = vbChecked
            'Univision
            txtUFile.Text = sgExportDirectory & "MktSpots.txt"
        Case "D"    '"IDC"
            ckcDGenType(0).Value = vbUnchecked
            ckcDGenType(2).Value = vbUnchecked
            ckcDGenType(3).Value = vbUnchecked
            edcDExportToPath.Text = sgExportDirectory
        Case "I"    '"ISCI"
            edcIExportToPath.Text = sgExportDirectory
            ckcIIncludeCommands(0).Value = vbChecked
            rbcIUniqueBy(1).Value = True
        Case "W"    '"Wegener"
            If Len(sgImportDirectory) > 0 Then
                txtWStationInfo.Text = Left$(sgImportDirectory, Len(sgImportDirectory) - 1)
            Else
                txtWStationInfo.Text = ""
            End If
            ckcWGenerate(0).Value = vbChecked
            ckcWGenerate(1).Value = vbChecked
            ckcWGenCSV.Value = vbUnchecked
            edcWExportToPath.Text = sgExportDirectory
        Case "X"    '"X-Digital"
            ckcXExportType(0).Value = vbUnchecked
            ckcXExportType(1).Value = vbUnchecked
            optXGenType(0).Value = True
            rbcXSpots(1).Value = True
            mSetXProvider
            rbcXProvider(0).Value = True
            edcXExportToPath.Text = sgExportDirectory
            ckcXReexport.Value = vbUnchecked
        Case "S"    '"StarGd"
            edcSExportToPath.Text = sgExportDirectory
        Case "C"    '"C & C"
            txtCFile.Text = sgExportDirectory & "CnCSpots.txt"
        Case "R"    '"ISCI C/R"
            edcRExportToPath.Text = sgExportDirectory
            '7459
             mSetRPrefix
        Case "4", "5"    '"RCS 4 and 5"
            chkRCSZone(0).Value = vbUnchecked
            chkRCSZone(1).Value = vbUnchecked
            chkRCSZone(2).Value = vbUnchecked
            chkRCSZone(3).Value = vbUnchecked
            For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                For ilZone = LBound(tgVehicleInfo(ilLoop).sZone) To UBound(tgVehicleInfo(ilLoop).sZone) Step 1
                    Select Case Left$(tgVehicleInfo(ilLoop).sZone(ilZone), 1)
                        Case "E"
                            If tgVehicleInfo(ilLoop).sFed(ilZone) = "*" Then
                                chkRCSZone(0).Value = vbChecked
                            End If
                        Case "C"
                            If tgVehicleInfo(ilLoop).sFed(ilZone) = "*" Then
                                chkRCSZone(1).Value = vbChecked
                            End If
                        Case "M"
                            If tgVehicleInfo(ilLoop).sFed(ilZone) = "*" Then
                                chkRCSZone(2).Value = vbChecked
                            End If
                        Case "P"
                            If tgVehicleInfo(ilLoop).sFed(ilZone) = "*" Then
                                chkRCSZone(3).Value = vbChecked
                            End If
                    End Select
                Next ilZone
            Next ilLoop
            edcRCSExportToPath.Text = sgExportDirectory
    End Select
End Sub

Private Sub mSetXProvider()
    Dim slName1 As String
    Dim slName2 As String
    Dim slSection As String
    Dim slSvIniPathFileName As String

    slSvIniPathFileName = sgIniPathFileName
    sgIniPathFileName = gXmlIniPath()
    If UBound(sgXDSSection) <= 1 Then
        If UBound(sgXDSSection) = 1 Then
            slSection = Mid(sgXDSSection(0), 2, Len(sgXDSSection(0)) - 2)
            If Not gLoadOption(slSection, "Provider", slName1) Then
                rbcXProvider(0).Caption = slSection
            Else
                rbcXProvider(0).Caption = slName1
            End If
        Else
            rbcXProvider(0).Caption = "XDigital"
        End If
        frcXProvider.Visible = False
    Else
        slSection = Mid(sgXDSSection(0), 2, Len(sgXDSSection(0)) - 2)
        If Not gLoadOption(slSection, "Provider", slName1) Then
            rbcXProvider(0).Caption = slSection
        Else
            rbcXProvider(0).Caption = slName1
        End If
        slSection = Mid(sgXDSSection(1), 2, Len(sgXDSSection(1)) - 2)
        If Not gLoadOption(slSection, "Provider", slName2) Then
            rbcXProvider(1).Caption = slSection
        Else
            rbcXProvider(1).Caption = slName2
        End If
        frcXProvider.Visible = True
    End If
    sgIniPathFileName = slSvIniPathFileName
End Sub

Private Sub mSetRPrefix()
    '7459
    With frcRPrefix
        .BorderStyle = 0
        .Top = ckcRIncludeGeneric.Top
    End With
    lacRPrefix.Top = 0
    rbcRPrefix(0).Top = 0
    rbcRPrefix(1).Top = 0
    frcRPrefix.Height = lacRPrefix.Height + 10
End Sub

