VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmStation 
   Caption         =   "Station Information"
   ClientHeight    =   7065
   ClientLeft      =   7005
   ClientTop       =   7890
   ClientWidth     =   11790
   Icon            =   "AffStation.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   11790
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Sister Stations"
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   5
      Left            =   180
      TabIndex        =   107
      Top             =   7440
      Visible         =   0   'False
      Width           =   10245
      Begin VB.TextBox edcMasterStation 
         Height          =   285
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   161
         Top             =   1845
         Width           =   1230
      End
      Begin VB.Frame frcMarketClusterMarket 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   150
         TabIndex        =   155
         Top             =   1125
         Width           =   6585
         Begin VB.OptionButton rbcMarketClusterMarket 
            Caption         =   "Stations in any Market"
            Height          =   195
            Index           =   1
            Left            =   3660
            TabIndex        =   158
            Top             =   90
            Width           =   2145
         End
         Begin VB.OptionButton rbcMarketClusterMarket 
            Caption         =   "Stations within Same Market"
            Height          =   195
            Index           =   0
            Left            =   1065
            TabIndex        =   157
            Top             =   90
            Width           =   2610
         End
         Begin VB.Label lacMarketClusterMarket 
            Caption         =   "DMA Market"
            Height          =   240
            Left            =   0
            TabIndex        =   156
            Top             =   75
            Width           =   945
         End
      End
      Begin VB.Frame frcMarketClusterAction 
         Appearance      =   0  'Flat
         Caption         =   "Action"
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   135
         TabIndex        =   151
         Top             =   465
         Width           =   7065
         Begin VB.OptionButton rbcMarketCluster 
            Caption         =   "Remove from Sister Stations"
            Height          =   225
            Index           =   2
            Left            =   4170
            TabIndex        =   154
            Top             =   240
            Width           =   2625
         End
         Begin VB.OptionButton rbcMarketCluster 
            Caption         =   "Add to Sister Stations"
            Height          =   225
            Index           =   1
            Left            =   2145
            TabIndex        =   153
            Top             =   225
            Width           =   2190
         End
         Begin VB.OptionButton rbcMarketCluster 
            Caption         =   "Create Sister Stations"
            Height          =   225
            Index           =   0
            Left            =   210
            TabIndex        =   152
            Top             =   210
            Width           =   2385
         End
      End
      Begin VB.PictureBox pbcPicture6 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   315
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   -30
         Width           =   60
      End
      Begin VB.PictureBox pbcPicture5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   90
         Left            =   30
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   109
         Top             =   3090
         Width           =   60
      End
      Begin VB.PictureBox pbcPicture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   30
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   108
         Top             =   180
         Width           =   60
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSisterStations 
         Height          =   2385
         Left            =   135
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   2265
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   4207
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
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
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lacMasterStation 
         Caption         =   "Primary Station Call Letters:"
         Height          =   255
         Left            =   120
         TabIndex        =   160
         Top             =   1875
         Width           =   2700
      End
      Begin VB.Label lacMarketClusterNote 
         Caption         =   "Multicast XXX with:"
         Height          =   240
         Left            =   135
         TabIndex        =   159
         Top             =   1500
         Width           =   4350
      End
      Begin VB.Label lacMarketCluster 
         Caption         =   "DMA Market:"
         Height          =   195
         Left            =   120
         TabIndex        =   150
         Top             =   180
         Width           =   7035
      End
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Main"
      ForeColor       =   &H80000008&
      Height          =   5310
      Index           =   0
      Left            =   14205
      TabIndex        =   11
      Top             =   1410
      Width           =   11085
      Begin V81Affiliate.CSI_ComboBoxList cbcMoniker 
         Height          =   315
         Left            =   5670
         TabIndex        =   18
         Top             =   150
         Width           =   2205
         _ExtentX        =   3307
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_Calendar edcHistStartDate 
         Height          =   285
         Left            =   7665
         TabIndex        =   27
         Top             =   540
         Width           =   1245
         _ExtentX        =   2064
         _ExtentY        =   529
         Text            =   "04/26/2024"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BorderStyle     =   1
         CSI_ShowDropDownOnFocus=   0   'False
         CSI_InputBoxBoxAlignment=   0
         CSI_CalBackColor=   16777130
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
         CSI_CurDayForeColor=   51200
         CSI_ForceMondaySelectionOnly=   0   'False
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin VB.TextBox edcWatts 
         Height          =   285
         Left            =   8535
         MaxLength       =   13
         TabIndex        =   20
         Top             =   150
         Width           =   795
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcONCity 
         Height          =   315
         Left            =   6570
         TabIndex        =   47
         Top             =   1950
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcCountyLic 
         Height          =   315
         Left            =   4830
         TabIndex        =   58
         Top             =   3015
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcCity 
         Height          =   315
         Left            =   960
         TabIndex        =   35
         Top             =   1950
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcCityLic 
         Height          =   315
         Left            =   1080
         TabIndex        =   56
         Top             =   3000
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   4830
         MaxLength       =   20
         TabIndex        =   64
         Top             =   3480
         Width           =   2595
      End
      Begin VB.TextBox edcP12Plus 
         Height          =   285
         Left            =   9780
         MaxLength       =   13
         TabIndex        =   29
         Top             =   540
         Width           =   1110
      End
      Begin VB.TextBox edcFrequency 
         Height          =   285
         Left            =   3915
         MaxLength       =   6
         TabIndex        =   16
         Top             =   150
         Width           =   795
      End
      Begin VB.TextBox edcPermanentStationID 
         Height          =   285
         Left            =   9780
         MaxLength       =   9
         TabIndex        =   22
         Top             =   150
         Width           =   1110
      End
      Begin VB.PictureBox pbcSTab 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   870
         ScaleHeight     =   150
         ScaleWidth      =   210
         TabIndex        =   12
         Top             =   75
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox pbcTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   0
         Index           =   0
         Left            =   60
         ScaleHeight     =   0
         ScaleWidth      =   60
         TabIndex        =   85
         Top             =   3630
         Width           =   60
      End
      Begin VB.TextBox txtStaPhone 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   62
         Top             =   3480
         Width           =   2595
      End
      Begin VB.TextBox txtCallLetters 
         Height          =   285
         Left            =   1050
         MaxLength       =   40
         TabIndex        =   14
         Top             =   150
         Width           =   2160
      End
      Begin VB.Frame frcMailingAddress 
         Caption         =   "Mailing Address"
         Height          =   1950
         Left            =   120
         TabIndex        =   30
         Top             =   885
         Width           =   5145
         Begin VB.TextBox edcCountry 
            Height          =   285
            Left            =   3000
            MaxLength       =   40
            TabIndex        =   41
            Top             =   1500
            Width           =   1950
         End
         Begin VB.TextBox txtZip 
            Height          =   285
            Left            =   840
            MaxLength       =   20
            TabIndex        =   39
            Top             =   1500
            Width           =   795
         End
         Begin VB.TextBox txtAddr2 
            Height          =   285
            Left            =   840
            MaxLength       =   40
            TabIndex        =   33
            Top             =   645
            Width           =   4110
         End
         Begin VB.TextBox txtAddr1 
            Height          =   285
            Left            =   840
            MaxLength       =   40
            TabIndex        =   32
            Top             =   240
            Width           =   4110
         End
         Begin VB.ComboBox cboState 
            Height          =   315
            ItemData        =   "AffStation.frx":08CA
            Left            =   3390
            List            =   "AffStation.frx":08CC
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   1065
            Width           =   1575
         End
         Begin VB.Label lacCountry 
            Caption         =   "Country:"
            Height          =   255
            Left            =   2205
            TabIndex        =   40
            Top             =   1530
            Width           =   1335
         End
         Begin VB.Label lacAddr1 
            Caption         =   "Address:"
            Height          =   255
            Left            =   30
            TabIndex        =   31
            Top             =   270
            Width           =   1020
         End
         Begin VB.Label lacCity 
            Caption         =   "City:"
            Height          =   255
            Left            =   30
            TabIndex        =   34
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lacState 
            Caption         =   "State:"
            Height          =   255
            Left            =   2835
            TabIndex        =   36
            Top             =   1080
            Width           =   795
         End
         Begin VB.Label lacZip 
            Caption         =   "Zip:"
            Height          =   255
            Left            =   30
            TabIndex        =   38
            Top             =   1530
            Width           =   765
         End
      End
      Begin VB.Frame frcLicense 
         Caption         =   "License"
         Height          =   570
         Left            =   120
         TabIndex        =   54
         Top             =   2835
         Width           =   10875
         Begin VB.ComboBox cboStateLic 
            Height          =   315
            ItemData        =   "AffStation.frx":08CE
            Left            =   8250
            List            =   "AffStation.frx":08D0
            Sorted          =   -1  'True
            TabIndex        =   60
            Top             =   180
            Width           =   2595
         End
         Begin VB.Label lacCityLic 
            Caption         =   "City:"
            Height          =   255
            Left            =   30
            TabIndex        =   55
            Top             =   210
            Width           =   1305
         End
         Begin VB.Label Label30 
            Caption         =   "State:"
            Height          =   255
            Left            =   7470
            TabIndex        =   59
            Top             =   210
            Width           =   1395
         End
         Begin VB.Label lacCountyLic 
            Caption         =   "County:"
            Height          =   255
            Left            =   3705
            TabIndex        =   57
            Top             =   210
            Width           =   675
         End
      End
      Begin VB.Frame frcPhysicalAddress 
         Caption         =   "Physical Address"
         Height          =   1950
         Left            =   5850
         TabIndex        =   42
         Top             =   885
         Width           =   5145
         Begin VB.TextBox txtONAddr2 
            Height          =   285
            Left            =   840
            MaxLength       =   40
            TabIndex        =   45
            Top             =   660
            Width           =   4110
         End
         Begin VB.TextBox txtONZip 
            Height          =   285
            Left            =   840
            MaxLength       =   20
            TabIndex        =   51
            Top             =   1500
            Width           =   795
         End
         Begin VB.TextBox txtONAddr1 
            Height          =   285
            Left            =   840
            MaxLength       =   40
            TabIndex        =   44
            Top             =   240
            Width           =   4110
         End
         Begin VB.TextBox txtONCountry 
            Height          =   285
            Left            =   3000
            MaxLength       =   40
            TabIndex        =   53
            Top             =   1500
            Width           =   1950
         End
         Begin VB.ComboBox cboONState 
            Height          =   315
            ItemData        =   "AffStation.frx":08D2
            Left            =   3390
            List            =   "AffStation.frx":08D4
            Sorted          =   -1  'True
            TabIndex        =   49
            Top             =   1065
            Width           =   1575
         End
         Begin VB.Label lacONAddr1 
            Caption         =   "Address:"
            Height          =   255
            Left            =   30
            TabIndex        =   43
            Top             =   270
            Width           =   1455
         End
         Begin VB.Label lacONCity 
            Caption         =   "City:"
            Height          =   255
            Left            =   30
            TabIndex        =   46
            Top             =   1080
            Width           =   1410
         End
         Begin VB.Label lacONState 
            Caption         =   "State:"
            Height          =   255
            Left            =   2835
            TabIndex        =   48
            Top             =   1080
            Width           =   780
         End
         Begin VB.Label lacONCountry 
            Caption         =   "Country:"
            Height          =   255
            Left            =   2205
            TabIndex        =   52
            Top             =   1530
            Width           =   1335
         End
         Begin VB.Label lacONZip 
            Caption         =   "Zip:"
            Height          =   255
            Left            =   30
            TabIndex        =   50
            Top             =   1530
            Width           =   465
         End
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcOwner 
         Height          =   315
         Left            =   1080
         TabIndex        =   68
         Top             =   3930
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcDMAMarket 
         Height          =   315
         Left            =   4830
         TabIndex        =   70
         Top             =   3930
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcMSAMarket 
         Height          =   315
         Left            =   8355
         TabIndex        =   72
         Top             =   3930
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcOperator 
         Height          =   315
         Left            =   1080
         TabIndex        =   74
         Top             =   4380
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcTerritory 
         Height          =   315
         Left            =   4830
         TabIndex        =   76
         Top             =   4380
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcArea 
         Height          =   315
         Left            =   8355
         TabIndex        =   78
         Top             =   4380
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcMarketRep 
         Height          =   315
         Left            =   1080
         TabIndex        =   80
         Top             =   4830
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcServiceRep 
         Height          =   315
         Left            =   4830
         TabIndex        =   82
         Top             =   4830
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcFormat 
         Height          =   315
         Left            =   8355
         TabIndex        =   66
         Top             =   3480
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcTimeZone 
         Height          =   315
         Left            =   8355
         TabIndex        =   84
         Top             =   4830
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         BackColor       =   -2147483643
         ForeColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin VB.Label lacWatts 
         Caption         =   "Watts:"
         Height          =   255
         Left            =   7995
         TabIndex        =   19
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lacDaylight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Daylight Saving"
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
         Height          =   270
         Left            =   3270
         TabIndex        =   25
         Top             =   555
         Width           =   2340
      End
      Begin VB.Label lacOnAir 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "On Air"
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
         Height          =   270
         Left            =   120
         TabIndex        =   23
         Top             =   555
         Width           =   1035
      End
      Begin VB.Label lacCommercial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Commercial"
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
         Height          =   270
         Left            =   1395
         TabIndex        =   24
         Top             =   555
         Width           =   1650
      End
      Begin VB.Label lacOperator 
         Caption         =   "Operator:"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   4410
         Width           =   1395
      End
      Begin VB.Label lacHistStartDate 
         Caption         =   "Historical Start Date:"
         Height          =   255
         Left            =   6045
         TabIndex        =   26
         Top             =   570
         Width           =   1725
      End
      Begin VB.Label lacServiceRep 
         Caption         =   "Service Rep:"
         Height          =   255
         Left            =   3810
         TabIndex        =   81
         Top             =   4860
         Width           =   1125
      End
      Begin VB.Label lacMktRep 
         Caption         =   "Market Rep:"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   4860
         Width           =   1200
      End
      Begin VB.Label lacArea 
         Caption         =   "Area:"
         Height          =   255
         Left            =   7575
         TabIndex        =   77
         Top             =   4410
         Width           =   900
      End
      Begin VB.Label lacMSAMarket 
         Caption         =   "MSA:"
         Height          =   255
         Left            =   7575
         TabIndex        =   71
         Top             =   3960
         Width           =   1395
      End
      Begin VB.Label lacDMA 
         Caption         =   "DMA:"
         Height          =   255
         Left            =   3810
         TabIndex        =   69
         Top             =   3960
         Width           =   1395
      End
      Begin VB.Label lacOwner 
         Caption         =   "Owner:"
         Height          =   270
         Left            =   120
         TabIndex        =   67
         Top             =   3960
         Width           =   1395
      End
      Begin VB.Label lacFx 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   3810
         TabIndex        =   63
         Top             =   3510
         Width           =   1215
      End
      Begin VB.Label lacMoniker 
         Caption         =   "Moniker:"
         Height          =   255
         Left            =   4890
         TabIndex        =   17
         Top             =   180
         Width           =   900
      End
      Begin VB.Label lacP12Plus 
         Caption         =   "P12+:"
         Height          =   255
         Left            =   9075
         TabIndex        =   28
         Top             =   570
         Width           =   450
      End
      Begin VB.Label lacFrequency 
         Caption         =   "Freq:"
         Height          =   255
         Left            =   3435
         TabIndex        =   15
         Top             =   180
         Width           =   450
      End
      Begin VB.Label lacPermanentStationID 
         Caption         =   "ID:"
         Height          =   255
         Left            =   9480
         TabIndex        =   21
         Top             =   180
         Width           =   300
      End
      Begin VB.Label lacFormat 
         Caption         =   "Format:"
         Height          =   255
         Left            =   7575
         TabIndex        =   65
         Top             =   3510
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "Territory:"
         Height          =   255
         Left            =   3810
         TabIndex        =   75
         Top             =   4410
         Width           =   1395
      End
      Begin VB.Label lacPhone 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   3510
         Width           =   1815
      End
      Begin VB.Label lacZone 
         Caption         =   "Zone:"
         Height          =   255
         Left            =   7575
         TabIndex        =   83
         Top             =   4860
         Width           =   1065
      End
      Begin VB.Label labName 
         Caption         =   "Call Letters:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Multi-cast"
      ForeColor       =   &H80000008&
      Height          =   5325
      Index           =   4
      Left            =   10515
      TabIndex        =   103
      Top             =   11475
      Visible         =   0   'False
      Width           =   9825
      Begin VB.Frame frcMulticastAction 
         Appearance      =   0  'Flat
         Caption         =   "Action"
         ForeColor       =   &H80000008&
         Height          =   2325
         Left            =   225
         TabIndex        =   164
         Top             =   405
         Width           =   9435
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   795
            Left            =   3375
            TabIndex        =   184
            Top             =   1395
            Width           =   2355
            Begin VB.OptionButton rbcMulticastMarket_Add 
               Caption         =   "Stations in any Market"
               Height          =   195
               Index           =   1
               Left            =   45
               TabIndex        =   186
               Top             =   525
               Width           =   1905
            End
            Begin VB.OptionButton rbcMulticastMarket_Add 
               Caption         =   "Stations within Same Market"
               Height          =   195
               Index           =   0
               Left            =   45
               TabIndex        =   185
               Top             =   270
               Width           =   2310
            End
            Begin VB.Label Label2 
               Caption         =   "DMA Market"
               ForeColor       =   &H00000000&
               Height          =   165
               Left            =   60
               TabIndex        =   187
               Top             =   0
               Width           =   1170
            End
         End
         Begin VB.Frame frcMulticastOwner_Add 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   3405
            TabIndex        =   180
            Top             =   540
            Width           =   2355
            Begin VB.OptionButton rbcMulticastOwner_Add 
               Caption         =   "Stations for Same Owner"
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   0
               Left            =   0
               TabIndex        =   181
               Top             =   225
               Width           =   2085
            End
            Begin VB.OptionButton rbcMulticastOwner_Add 
               Caption         =   "Stations for any Owner"
               Height          =   210
               Index           =   1
               Left            =   0
               TabIndex        =   182
               Top             =   495
               Width           =   1935
            End
            Begin VB.Label lacMulticastOwner_Add 
               Caption         =   "Owner"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   0
               TabIndex        =   183
               Top             =   -15
               Width           =   780
            End
         End
         Begin VB.Frame frcMulticastMarket 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   795
            Left            =   390
            TabIndex        =   176
            Top             =   1395
            Width           =   2355
            Begin VB.OptionButton rbcMulticastMarket 
               Caption         =   "Stations within Same Market"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   177
               Top             =   270
               Width           =   2340
            End
            Begin VB.OptionButton rbcMulticastMarket 
               Caption         =   "Stations in any Market"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   178
               Top             =   525
               Width           =   1905
            End
            Begin VB.Label lacMulticastMarket 
               Caption         =   "DMA Market"
               ForeColor       =   &H00000000&
               Height          =   165
               Left            =   60
               TabIndex        =   179
               Top             =   0
               Width           =   1170
            End
         End
         Begin VB.Frame frcMulticastOwner 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   450
            TabIndex        =   172
            Top             =   540
            Width           =   2355
            Begin VB.OptionButton rbcMulticastOwner 
               Caption         =   "Stations for any Owner"
               Height          =   210
               Index           =   1
               Left            =   0
               TabIndex        =   174
               Top             =   495
               Width           =   1935
            End
            Begin VB.OptionButton rbcMulticastOwner 
               Caption         =   "Stations for Same Owner"
               Height          =   210
               Index           =   0
               Left            =   0
               TabIndex        =   173
               Top             =   225
               Width           =   2085
            End
            Begin VB.Label lacMulticastOwner 
               Caption         =   "Owner"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   0
               TabIndex        =   175
               Top             =   -15
               Width           =   780
            End
         End
         Begin VB.OptionButton rbcMulticast 
            Caption         =   "Create Multi-Cast"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   165
            Top             =   210
            Width           =   1875
         End
         Begin VB.OptionButton rbcMulticast 
            Caption         =   "Add to Multi-Cast"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   3180
            TabIndex        =   166
            Top             =   225
            Width           =   1935
         End
         Begin VB.OptionButton rbcMulticast 
            Caption         =   "Remove from Multi-Cast"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   6225
            TabIndex        =   167
            Top             =   240
            Width           =   2370
         End
         Begin VB.Line Line2 
            BorderStyle     =   3  'Dot
            X1              =   5940
            X2              =   5940
            Y1              =   690
            Y2              =   2025
         End
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            X1              =   2955
            X2              =   2955
            Y1              =   690
            Y2              =   2025
         End
      End
      Begin VB.PictureBox pbcPicture4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   30
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   106
         Top             =   180
         Width           =   60
      End
      Begin VB.PictureBox pbcPicture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   90
         Left            =   30
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   105
         Top             =   3090
         Width           =   60
      End
      Begin VB.PictureBox pbcPicture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   315
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   -30
         Width           =   60
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdMulticast 
         Height          =   2520
         Left            =   225
         TabIndex        =   169
         TabStop         =   0   'False
         Top             =   2790
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   4445
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lacMulticastNote 
         Caption         =   "Multicast XXX with:"
         Height          =   240
         Left            =   225
         TabIndex        =   168
         Top             =   2115
         Width           =   3720
      End
      Begin VB.Label lblOwner 
         Caption         =   "Owner:"
         Height          =   195
         Left            =   135
         TabIndex        =   163
         Top             =   165
         Width           =   7095
      End
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2070
      Top             =   6795
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   1725
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   6750
      Width           =   120
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Preferred Method of Contact"
      ForeColor       =   &H80000008&
      Height          =   1125
      Index           =   7
      Left            =   10845
      TabIndex        =   86
      Top             =   5685
      Visible         =   0   'False
      Width           =   8505
      Begin VB.PictureBox pbcSTab 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   615
         ScaleHeight     =   255
         ScaleWidth      =   270
         TabIndex        =   87
         Top             =   15
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox pbcTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   1
         Left            =   90
         ScaleHeight     =   45
         ScaleWidth      =   60
         TabIndex        =   88
         Top             =   3420
         Width           =   60
      End
   End
   Begin VB.ListBox lbcLookup1 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "AffStation.frx":08D6
      Left            =   525
      List            =   "AffStation.frx":08D8
      TabIndex        =   101
      Top             =   6675
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdErase 
      Caption         =   "Erase"
      Height          =   375
      Left            =   5895
      TabIndex        =   99
      Top             =   6540
      Width           =   1335
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   2805
      TabIndex        =   97
      Top             =   6540
      Width           =   1335
   End
   Begin VB.Frame frcSelect 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.Frame frcType 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   195
         Width           =   2130
         Begin VB.OptionButton optSP 
            Caption         =   "Person"
            Height          =   255
            Index           =   1
            Left            =   930
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   30
            Width           =   945
         End
         Begin VB.OptionButton optSP 
            Caption         =   "Station"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   30
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.ComboBox cboStations 
         Height          =   315
         ItemData        =   "AffStation.frx":08DA
         Left            =   6900
         List            =   "AffStation.frx":08DC
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   210
         Width           =   4260
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2385
         TabIndex        =   4
         Top             =   240
         Width           =   3030
         Begin VB.OptionButton optSort 
            Caption         =   "Stations"
            Height          =   255
            Index           =   0
            Left            =   705
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton optSort 
            Caption         =   "DMA"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   810
         End
         Begin VB.Label lblSort 
            Caption         =   "Sort By:"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Call Letters-Band:"
         Height          =   255
         Left            =   5430
         TabIndex        =   8
         Top             =   270
         Width           =   1575
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   15
      Top             =   6495
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7065
      FormDesignWidth =   11790
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7455
      TabIndex        =   100
      Top             =   6540
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4350
      TabIndex        =   98
      Top             =   6540
      Width           =   1335
   End
   Begin VB.Frame frcNotUsed 
      Caption         =   "Not Used- Controls"
      Height          =   3495
      Left            =   9105
      TabIndex        =   137
      Top             =   6810
      Visible         =   0   'False
      Width           =   5925
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   720
         MaxLength       =   1
         TabIndex        =   170
         Top             =   2775
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtMarkRank 
         Enabled         =   0   'False
         Height          =   285
         Left            =   660
         TabIndex        =   147
         Top             =   2295
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.OptionButton optPrefCon 
         Caption         =   "Web Page"
         Height          =   255
         Index           =   4
         Left            =   4575
         TabIndex        =   146
         Top             =   1380
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtWebPage 
         Height          =   285
         Left            =   3135
         MaxLength       =   50
         TabIndex        =   145
         Text            =   "http://www."
         Top             =   1740
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.OptionButton optPrefCon 
         Caption         =   "Post Office"
         Height          =   255
         Index           =   3
         Left            =   4575
         TabIndex        =   144
         Top             =   1140
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optPrefCon 
         Caption         =   "Overnight Mail"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   143
         Top             =   1905
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.OptionButton optPrefCon 
         Caption         =   "Fax"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   142
         Top             =   1920
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.OptionButton optPrefCon 
         Caption         =   "Email"
         Height          =   255
         Index           =   0
         Left            =   3015
         TabIndex        =   141
         Top             =   1380
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.TextBox txtWebEmail 
         Height          =   285
         Left            =   1305
         MaxLength       =   240
         TabIndex        =   139
         Top             =   750
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.TextBox txtMarket 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   60
         TabIndex        =   138
         Top             =   285
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label lacPort 
         Caption         =   "Port:"
         Height          =   240
         Left            =   150
         TabIndex        =   171
         Top             =   2790
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Rank:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   75
         TabIndex        =   148
         Top             =   2325
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblWebEmail 
         Caption         =   "Web Email 1"
         Height          =   255
         Left            =   105
         TabIndex        =   140
         Top             =   780
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Interface"
      ForeColor       =   &H80000008&
      Height          =   4350
      Index           =   1
      Left            =   9465
      TabIndex        =   111
      Top             =   6630
      Visible         =   0   'False
      Width           =   9495
      Begin VB.TextBox edcWebNumber 
         Height          =   285
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   121
         Top             =   1000
         Width           =   495
      End
      Begin VB.TextBox txtWebSpotsPerPage 
         Height          =   285
         Left            =   8205
         MaxLength       =   5
         TabIndex        =   136
         Top             =   3900
         Width           =   735
      End
      Begin VB.TextBox txtIPumpID 
         Height          =   285
         Left            =   2415
         MaxLength       =   10
         TabIndex        =   130
         Top             =   2955
         Width           =   1395
      End
      Begin VB.CheckBox ckcUsedFor 
         Caption         =   "Pledge vs Air (CSV)"
         Height          =   195
         Index           =   4
         Left            =   7290
         TabIndex        =   117
         Top             =   210
         Width           =   2010
      End
      Begin VB.TextBox edcEnterpriseID 
         Height          =   285
         Left            =   1890
         MaxLength       =   5
         TabIndex        =   128
         Top             =   2460
         Width           =   1020
      End
      Begin VB.TextBox txtWebPW 
         Height          =   285
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   125
         Top             =   1980
         Width           =   2730
      End
      Begin VB.CommandButton cmdGenPassword 
         Caption         =   "Generate Password"
         Height          =   375
         Left            =   4335
         TabIndex        =   126
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtWebAddress 
         Height          =   285
         Left            =   1365
         MaxLength       =   90
         TabIndex        =   123
         Top             =   1515
         Width           =   8025
      End
      Begin VB.TextBox txtXDSStationID 
         Height          =   285
         Left            =   1365
         MaxLength       =   40
         TabIndex        =   119
         Top             =   600
         Width           =   1110
      End
      Begin VB.TextBox txtSerialNo1 
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   132
         Top             =   3420
         Width           =   1095
      End
      Begin VB.TextBox txtSerialNo2 
         Height          =   285
         Left            =   4575
         MaxLength       =   10
         TabIndex        =   134
         Top             =   3420
         Width           =   1110
      End
      Begin VB.CheckBox ckcUsedFor 
         Caption         =   "Agreements"
         Height          =   195
         Index           =   0
         Left            =   1365
         TabIndex        =   113
         Top             =   210
         Width           =   1380
      End
      Begin VB.CheckBox ckcUsedFor 
         Caption         =   "X-Digital"
         Height          =   195
         Index           =   1
         Left            =   2790
         TabIndex        =   114
         Top             =   210
         Width           =   1095
      End
      Begin VB.CheckBox ckcUsedFor 
         Caption         =   "Wegener-Compel"
         Height          =   195
         Index           =   2
         Left            =   4095
         TabIndex        =   115
         Top             =   210
         Width           =   1740
      End
      Begin VB.CheckBox ckcUsedFor 
         Caption         =   "OLA"
         Height          =   195
         Index           =   3
         Left            =   6255
         TabIndex        =   116
         Top             =   210
         Width           =   825
      End
      Begin VB.Label lacWebNumber 
         Caption         =   "Web Affiliate Version Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   1080
         Width           =   2220
      End
      Begin VB.Label lacWebSpotsPerPage 
         Caption         =   "Number Spots to Show On Electronic Affidavit System per Page (blank or 0 indicates no limit per page):"
         Height          =   255
         Left            =   120
         TabIndex        =   135
         Top             =   3945
         Width           =   8565
      End
      Begin VB.Label lacIPumpID 
         Caption         =   "Wegener iPump ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   2985
         Width           =   2295
      End
      Begin VB.Label lacEnterpriseID 
         Caption         =   "Transact Enterprise ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   127
         Top             =   2505
         Width           =   2190
      End
      Begin VB.Label lblWebPW 
         Caption         =   "Web Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   124
         Top             =   2010
         Width           =   1335
      End
      Begin VB.Label lacWebAddress 
         Caption         =   "Web Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   1560
         Width           =   1410
      End
      Begin VB.Label lacXDSStationID 
         Caption         =   "XDS Station ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   645
         Width           =   1260
      End
      Begin VB.Label lacSerialNo1 
         Caption         =   "StarGuide- Serial # 1:"
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   3435
         Width           =   1755
      End
      Begin VB.Label lacSerialNo2 
         Caption         =   "Serial # 2:"
         Height          =   255
         Left            =   3705
         TabIndex        =   133
         Top             =   3435
         Width           =   870
      End
      Begin VB.Label lacUsedFor 
         Caption         =   "Used for:"
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   225
         Width           =   1035
      End
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Personnel"
      ForeColor       =   &H80000008&
      Height          =   3735
      Index           =   2
      Left            =   9690
      TabIndex        =   89
      Top             =   6420
      Visible         =   0   'False
      Width           =   10770
      Begin V81Affiliate.AffContactGrid udcContactGrid 
         Height          =   390
         Left            =   30
         TabIndex        =   149
         Top             =   255
         Width           =   9090
         _ExtentX        =   15266
         _ExtentY        =   767
      End
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "History"
      ForeColor       =   &H80000008&
      Height          =   3750
      Index           =   3
      Left            =   9900
      TabIndex        =   90
      Top             =   6210
      Visible         =   0   'False
      Width           =   8715
      Begin VB.PictureBox pbcArrow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   45
         Picture         =   "AffStation.frx":08DE
         ScaleHeight     =   165
         ScaleWidth      =   90
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   420
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox pbcHistTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   210
         Left            =   30
         ScaleHeight     =   210
         ScaleWidth      =   300
         TabIndex        =   95
         Top             =   3510
         Width           =   300
      End
      Begin VB.PictureBox pbcHistSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   45
         ScaleHeight     =   195
         ScaleWidth      =   135
         TabIndex        =   93
         Top             =   720
         Width           =   135
      End
      Begin VB.TextBox txtHistory 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4845
         TabIndex        =   94
         Top             =   600
         Visible         =   0   'False
         Width           =   945
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHistory 
         Height          =   2760
         Left            =   195
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   360
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   4868
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollBars      =   2
         Appearance      =   0
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
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "To Delete Row: 'Select Row'  then Click on Trash"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   795
         TabIndex        =   102
         Top             =   3135
         Width           =   3840
      End
      Begin VB.Image imcTrash 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   255
         Picture         =   "AffStation.frx":0BE8
         Top             =   3120
         Width           =   480
      End
   End
   Begin ComctlLib.TabStrip tscStation 
      Height          =   5670
      Left            =   150
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   660
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   10001
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "&Main"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "&History"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "&Personnel"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "&Sister Stations"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "M&ulti-Cast"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "&Interface"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imcTabColor 
      Left            =   2085
      Top             =   6390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   86
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":0EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":1C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":29FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":3784
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":450A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":5290
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":6016
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":6D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":7B22
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":88A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":962E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AffStation.frx":A3B4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmStation - enters station(affiliate) information
'*  cboStations contains hidden "Index" column
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'*
'******************************************************
Option Explicit
Option Compare Text

Private smOrigWebNumber As String
Private smNewWebNumber As String
Private imShttCode As Integer
Private smShttWebPW As String
Private bmCancelled As Boolean
'Private smShttWebEmail As String
'Private smShttWebEmail2 As String
'Private smShttWebEmail3 As String
Private imFirstTime As Integer
Private imFieldChgd As Integer
Private lmArttCode As Long
Private lmMultiCastGroupID As Long
Private lmClusterGroupID As Long
Private smClusterStations As String
Private imDMAMktCode As Integer
Private imMSAMktCode As Integer
Private imFormatIndex As Integer
Private imTimeZoneIndex As Integer
Private imStateIndex As Integer
Private imOnStateIndex As Integer
Private imStateLicIndex As Integer
Private imInChg As Integer
Private imInSave As Integer
Private imBSMode As Integer
Private imDMAMarketBSMode As Integer
Private imOwnerBSMode As Integer
Private imMSAMarketBSMode As Integer
Private imTabIndex As Integer
Private imIgnoreTabs As Integer
Private smCurCallLetters As String
Private smCurTimeZone As String
Private tmHistoryInfo() As HISTORYINFO
Private lmAttCode() As Long
Private imMntIndex As Integer
Private attrst As ADODB.Recordset
Private DATRST As ADODB.Recordset
Private hmMsg As Integer
Private smOldStaPhoneNum As String
Private IsStatDirty As Boolean
Private smWebExports As String
Private sToFileHeader As String
Private lmAttCodesToUpdateWeb() As Long
Private imVefCode As Integer
Private smCurValue As String
Private imWebPWUpdated As Integer
'Private imMonthlyPostingUpdated As Integer
'Private smExistingMonthlyPosting As String
Private smNewMonthlyPosting As String
Private imWebEmailUpdated As Integer
Private smExistingWebPW As String
Private smExistingWebEmail As String
Private imWebUpdateAll As Integer
Private imScroll As Integer
Private imBaseIdx As Integer
Private bmAdjPledge As Boolean
'Private imIsOwnerMarketDirty As Integer
Private imIsMulticastDirty As Integer
Private tmBaseStaInfo() As OWNEDSTATIONS
Private tmGroup1Sort() As OWNEDSTATIONS
Private tmGroup2Sort() As OWNEDSTATIONS
Private tmGroup3Sort() As OWNEDSTATIONS
Private tmGroup4Sort() As OWNEDSTATIONS
Private bFormWasAlreadyResized As Boolean
Private imIgnoreScroll As Boolean
Private smAttWebInterface As String
Private lmTabColor(0 To 5) As Long
Private bmDoPop As Boolean
Private smWegenerIPump As String
Private smOldMaster As String
Dim tmEmailInfo() As EMAILINFO

'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long


Private imPersFromArrow As Integer
Private bmIgnorePersonnelChange As Boolean

Private imEmailFromArrow As Integer
Private bmIgnoreEmailChange As Boolean

Private bmIgnoreTitleListChanges As Boolean

Private bmIgnoreOwnerChange As Boolean

Private imLastMCColSorted As Integer
Private imLastMCSort As Integer
Private bmIgnoreMulticastChange As Boolean

Private imLastSSColSorted As Integer
Private imLastSSSort As Integer
Private bmIgnoreSisterStationChange As Boolean
'7912
Private smExistingXDSSiteID As String
'10197
Private imExistingHonorDaylight As Integer
'Multi-Cast
Const MCCALLLETTERSINDEX = 0
Const MCMARKETINDEX = 1
Const MCLICCITYINDEX = 2
Const MCMAILSTATEINDEX = 3
Const MCOWNERINDEX = 4
Const MCSELECTEDINDEX = 5
Const MCSHTTCODEINDEX = 6
Const MCDMAMKTCODEINDEX = 7
Const MCSORTINDEX = 8

'Sister Stations
Const SSCALLLETTERSINDEX = 0
Const SSMARKETINDEX = 1
Const SSLICCITYINDEX = 2
Const SSMAILSTATEINDEX = 3
Const SSSELECTEDINDEX = 4
Const SSSHTTCODEINDEX = 5
Const SSSORTINDEX = 6

'Station History
Const SHCALLLETTERSINDEX = 0
Const SHLASTDATEINDEX = 1
Const SHCLTCODEINDEX = 2

Const WEBEMAILINDEX = 0
Const WEBSEQNOINDEX = 1

Const NAMEINDEX = 0
Const PHONEINDEX = 1
Const FAXINDEX = 2
Const EMAILINDEX = 3
Const TITLEINDEX = 4
Const AFCNTINDEX = 5
Const ISCI2INDEX = 6
Const PRCODEINDEX = 7
Const TNTCODEINDEX = 8
Const STATUSINDEX = 9   ' 0=No Change, 1=Changed

Const SCALLLETTERINDEX = 1
Const SDMAMARKETINDEX = 4   'Value must match how it is defined in frmStationSearch.
Const SSHTTCODEINDEX = 19   'Value must match how it is defined in frmStationSearch.  Constant defined in frmContactEMail and here





'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Zone Change Info               *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile(slToFile As String) As Integer
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer
    
    'On Error GoTo mOpenMsgFileErr:
    ilRet = 0
    slToFile = sgExportDirectory & "ZC" & Format$(gNow(), "mm") & Format$(gNow(), "dd") & Format$(gNow(), "yy") & ".txt"
    slNowDate = Format$(gNow(), sgShowDateForm)
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        On Error GoTo 0
        'ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Append As hmMsg
        ilRet = gFileOpen(slToFile, "Append", hmMsg)
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            gMsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    Else
        On Error GoTo 0
        'ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output Lock Write As hmMsg
        ilRet = gFileOpen(slToFile, "Output Lock Write", hmMsg)
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, "** Zone Change Info: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function

Private Sub mClearControls()
    Dim iLoop As Integer
    Dim llCol As Long
    
    On Error GoTo ErrHand
    lmArttCode = 0
    imDMAMktCode = 0
    imMSAMktCode = 0
    imShttCode = 0
    lmMultiCastGroupID = 0
    lmClusterGroupID = 0
    smClusterStations = ""
    imFormatIndex = -1
    imTimeZoneIndex = -1
    imStateIndex = -1
    imOnStateIndex = -1
    imStateLicIndex = -1
    cboStations.Text = "[New]"
    'optSP(0).Value = True
    txtCallLetters.Text = ""
    edcFrequency.Text = ""
    edcPermanentStationID.Text = ""
    smCurCallLetters = ""
    smCurTimeZone = ""
    txtXDSStationID.Text = ""
    cbcMoniker.SetListIndex = -1
    txtStaPhone.Text = ""
    txtAddr1.Text = ""
    txtAddr2.Text = ""
    'txtCity.Text = ""
    cbcCity.SetListIndex = -1
    'txtState.text = ""
    cboState.ListIndex = -1
    edcCountry.Text = ""
    txtONCountry.Text = ""
    txtZip.Text = ""
    'txtTimeZone.text = ""
    cbcTimeZone.SetListIndex = -1
    optPrefCon(1).Value = True
    'txtEmail.text = ""
    txtFax.Text = ""
    txtWebPage.Text = "http://www."
    txtSerialNo1.Text = ""
    txtSerialNo2.Text = ""
    txtPort.Text = ""
    'ckcPC.Value = 0
    'ckcHD.Value = 0
    txtWebEmail.Text = ""
    txtWebPW.Text = ""
    'Select No Rep
    '5/10/07:  Removed Affiliate Rep from Station File
    'For iLoop = 0 To cboAffRep1.ListCount - 1
    '    If cboAffRep1.ItemData(iLoop) = 0 Then
    '        cboAffRep1.ListIndex = 1
    '        Exit For
    '    End If
    'Next iLoop
    'txtARPhone.Text = ""
    'txtARPhone.Enabled = False
    txtWebAddress.Text = ""
    '4/14/21: TTP 9052
    'edcWebNumber.Text = ""
    If Trim$(sgWebNumber) = "" Then
        sgWebNumber = "2"
    End If
    edcWebNumber.Text = sgWebNumber
    'rbcWebSiteVersion(0).Enabled = True
    'rbcWebSiteVersion(1).Enabled = True
    'default value for web number to 1 = original site
    'rbcWebSiteVersion(0).Value = True
    'rbcWebSiteVersion(1).Value = False
        
    'cboFormat.ListIndex = -1
    'cboFormat.Text = ""
    'txtCityLic.Text = ""
    cbcCityLic.SetListIndex = -1
    cbcCountyLic.SetListIndex = -1
    'txtStateLic.text = ""
    cboStateLic.ListIndex = -1
    txtMarket.Text = ""
    'cboDMAMarketCluster.ListIndex = -1
    'cboDMAMarketCluster.Text = ""
    lacDMA.Caption = "DMA:"
    cbcDMAMarket.SetListIndex = -1
    'cboMSAMarketCluster.ListIndex = -1
    'cboMSAMarketCluster.Text = ""
    lacMSAMarket.Caption = "MSA:"
    cbcMSAMarket.SetListIndex = -1
    'cboOwner.ListIndex = -1
    'cboOwner.Text = ""
    cbcOwner.SetListIndex = -1
    lblOwner.Caption = ""
    cbcOperator.SetListIndex = -1
    cbcTerritory.SetListIndex = -1
    cbcFormat.SetListIndex = -1
    cbcArea.SetListIndex = -1
    cbcMarketRep.SetListIndex = -1
    cbcServiceRep.SetListIndex = -1
    txtMarkRank.Text = ""
    txtONAddr1.Text = ""
    txtONAddr2.Text = ""
    'txtONCity.Text = ""
    cbcONCity.SetListIndex = -1
    'txtONState.text = ""
    cboONState.ListIndex = -1
    txtONZip.Text = ""
    'txtSerialNo1.Text = ""
    'txtSerialNo2.Text = ""
    ckcUsedFor(0).Value = vbChecked     'Agreements
    ckcUsedFor(1).Value = vbUnchecked   'X-Digital
    ckcUsedFor(2).Value = vbUnchecked   'Wegener
    ckcUsedFor(3).Value = vbUnchecked   'OLA
    ckcUsedFor(4).Value = vbUnchecked   'Pledge vs Air
    'ckcMonthlyPosting(0).Value = vbUnchecked
    lacOnAir.Caption = "On Air"
    lacOnAir.BackColor = GREEN  '&HC000&
    lacCommercial.Caption = "Commercial"
    lacCommercial.BackColor = GREEN  '&HC000&
    lacDaylight.Caption = "Honor Daylight Savings"
    lacDaylight.BackColor = GREEN  '&HC000&
    edcP12Plus.Text = ""
    edcWatts.Text = ""
    edcHistStartDate.Text = ""
    edcEnterpriseID.Text = ""
    txtIPumpID.Text = ""
    txtWebSpotsPerPage.Text = ""
    'optDaylight(0).Value = True
    imFieldChgd = False
    gGrid_Clear grdHistory, True
    'gGrid_Clear grdEmail, True
    'gGrid_Clear grdPersonnel, True
    'gGrid_Clear grdEmail, True
    udcContactGrid.StationCode = imShttCode
    udcContactGrid.Action 3 'populate
    ReDim tmHistoryInfo(0 To 0) As HISTORYINFO
    ReDim lmAttCodesToUpdateWeb(0 To 0) As Long
    
    'lbcMSAMarketCluster.Clear
    
    ReDim tmGroup1Sort(0 To 0) As OWNEDSTATIONS 'Are Multicast, Same Owner, Same Market
    ReDim tmGroup2Sort(0 To 0) As OWNEDSTATIONS 'Are Multicast, Same Owner, Different Market
    ReDim tmGroup3Sort(0 To 0) As OWNEDSTATIONS 'Not Multicast, Same Owner, Same Market
    ReDim tmGroup4Sort(0 To 0) As OWNEDSTATIONS 'Not Multicast, Same Owner, Different Market
    'lbcMulticast.Clear
    'grdMulticast.Clear
    grdMulticast.Rows = 2
    grdMulticast.TextMatrix(1, MCCALLLETTERSINDEX) = ""
    grdMulticast.TextMatrix(1, MCMARKETINDEX) = ""
    grdMulticast.TextMatrix(1, MCLICCITYINDEX) = ""
    grdMulticast.TextMatrix(1, MCMAILSTATEINDEX) = ""
    grdMulticast.TextMatrix(1, MCOWNERINDEX) = ""
    grdMulticast.Row = grdMulticast.FixedRows
    For llCol = MCCALLLETTERSINDEX To MCOWNERINDEX Step 1
        grdMulticast.Col = llCol
        grdMulticast.CellBackColor = vbWhite
    Next llCol
    bmIgnoreMulticastChange = True
    rbcMulticast(0).Value = False
    rbcMulticast(1).Value = False
    rbcMulticast(2).Value = False
    rbcMulticast(0).Enabled = True
    rbcMulticast(1).Enabled = True
    rbcMulticast(2).Enabled = True
    rbcMulticastOwner(0).Value = True
    rbcMulticastMarket(0).Value = True
    rbcMulticastOwner_Add(0).Value = True
    rbcMulticastMarket_Add(0).Value = True
    
    lacMulticastNote.Visible = False
    bmIgnoreMulticastChange = False
    
    bmIgnoreSisterStationChange = True
    rbcMarketCluster(0).Value = False
    rbcMarketCluster(1).Value = False
    rbcMarketCluster(2).Value = False
    rbcMarketCluster(0).Enabled = True
    rbcMarketCluster(1).Enabled = True
    rbcMarketCluster(2).Enabled = True
    rbcMarketClusterMarket(0).Value = True
    lacMarketClusterNote.Visible = False
    grdSisterStations.Rows = 2
    grdSisterStations.TextMatrix(1, SSCALLLETTERSINDEX) = ""
    grdSisterStations.TextMatrix(1, SSMARKETINDEX) = ""
    grdSisterStations.TextMatrix(1, SSLICCITYINDEX) = ""
    grdSisterStations.TextMatrix(1, SSMAILSTATEINDEX) = ""
    grdSisterStations.Row = grdMulticast.FixedRows
    For llCol = SSCALLLETTERSINDEX To SSMAILSTATEINDEX Step 1
        grdSisterStations.Col = llCol
        grdSisterStations.CellBackColor = vbWhite
    Next llCol
    bmIgnoreSisterStationChange = False
    
    udcContactGrid.StationCode = imShttCode
    udcContactGrid.Action 4 'Clear
    udcContactGrid.Action 3 'populate
    
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-mClearControls"
End Sub

Private Sub mBindControls()
    Dim ilLoop As Integer
    Dim iUpper As Integer
    Dim i As Integer
    Dim ilRet As Integer
    Dim lRow As Long
    Dim temp_rst As ADODB.Recordset
    Dim slTemp As String
    Dim llIdx As Long
    
    On Error GoTo ErrHand
    imShttCode = rst!shttCode    'rst(0).Value
    If rst!shttType = 1 Then    'rst(26).Value = 1 Then
        optSP(1).Value = True
    Else
        optSP(0).Value = True
    End If
    'cboStations.Text = Trim$(rst!shttCallLetters)   'rst(2).Value)
    txtCallLetters.Text = Trim$(rst!shttCallLetters)
    smCurCallLetters = Trim$(rst!shttCallLetters)
    If rst!shttType <> 1 Then
        txtXDSStationID.Text = Trim$(rst!shttStationId)
        '7912
        smExistingXDSSiteID = txtXDSStationID.Text
        txtXDSStationID.Enabled = True
        edcFrequency.Text = Trim$(rst!shttFrequency)
        edcFrequency.Enabled = True
        edcPermanentStationID.Text = Trim$(rst!shttPermStationID)
        edcPermanentStationID.Enabled = True
    Else
        edcFrequency.Text = ""
        edcFrequency.Enabled = False
        edcPermanentStationID.Text = ""
        edcPermanentStationID.Enabled = False
        txtXDSStationID.Text = ""
        '7912
        smExistingXDSSiteID = ""
        txtXDSStationID.Enabled = False
    End If
    cbcMoniker.SetListIndex = -1
    If rst!shttMonikerMntCode > 0 Then
        For ilLoop = 0 To cbcMoniker.ListCount - 1 Step 1
            If cbcMoniker.GetItemData(ilLoop) = rst!shttMonikerMntCode Then
                cbcMoniker.SetListIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    txtAddr1.Text = Trim$(rst!shttAddress1)
    txtAddr2.Text = Trim$(rst!shttAddress2)
    'txtCity.Text = Trim$(rst!shttCity)
    cbcCity.SetListIndex = -1
    If rst!shttCityMntCode > 0 Then
        For ilLoop = 0 To cbcCity.ListCount - 1 Step 1
            If cbcCity.GetItemData(ilLoop) = rst!shttCityMntCode Then
                cbcCity.SetListIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    'txtState.text = Trim$(rst!shttState)
    cboState.ListIndex = -1
    imStateIndex = -1
    If Trim$(rst!shttState) <> "" Then
        For ilLoop = 0 To cboState.ListCount - 1 Step 1
            If StrComp(Trim$(rst!shttState), Trim$(tgStateInfo(cboState.ItemData(ilLoop)).sPostalName), vbTextCompare) = 0 Then
                cboState.ListIndex = ilLoop
                cboState.Text = cboState.List(ilLoop)
                imStateIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    edcCountry.Text = Trim$(rst!shttCountry)
    txtONCountry.Text = Trim$(rst!shttOnCountry)
    txtZip.Text = Trim$(rst!shttZip)
    If rst!shttSelected = -1 Then   'rst(9).Value = -1 Then
        For i = 0 To 2 Step 1
            optPrefCon(i).Value = False
        Next i
    Else
        optPrefCon(rst!shttSelected).Value = True
    End If
    
    'txtEmail.text = Trim$(rst!shttEMail)    'rst(10).Value)
    txtFax.Text = Trim$(rst!shttFax)    'rst(11).Value)
    txtStaPhone.Text = Trim$(rst!shttPhone)    'rst(12).Value)
    'txtTimeZone.text = Trim$(rst!shttTimeZone)
    imTimeZoneIndex = -1
    smCurTimeZone = ""
    cbcTimeZone.SetListIndex = -1
    If Trim$(rst!shttTztCode) > 0 Then
        For ilLoop = 0 To cbcTimeZone.ListCount - 1 Step 1
            If rst!shttTztCode = tgTimeZoneInfo(cbcTimeZone.GetItemData(ilLoop)).iCode Then
                cbcTimeZone.SetListIndex = ilLoop
                'cboTimeZone.Text = cboTimeZone.List(ilLoop)
                smCurTimeZone = tgTimeZoneInfo(cbcTimeZone.GetItemData(ilLoop)).sCSIName
                imTimeZoneIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    txtWebPage.Text = Trim$(rst!shttHomePage)    'rst(14).Value)
    'txtCityLic.Text = Trim$(rst!shttCityLic)
    cbcCityLic.SetListIndex = -1
    If rst!shttCityLicMntCode > 0 Then
        For ilLoop = 0 To cbcCityLic.ListCount - 1 Step 1
            If cbcCityLic.GetItemData(ilLoop) = rst!shttCityLicMntCode Then
                cbcCityLic.SetListIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    cbcCountyLic.SetListIndex = -1
    If rst!shttCountyLicMntCode > 0 Then
        For ilLoop = 0 To cbcCountyLic.ListCount - 1 Step 1
            If cbcCountyLic.GetItemData(ilLoop) = rst!shttCountyLicMntCode Then
                cbcCountyLic.SetListIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    'txtStateLic.text = Trim$(rst!shttStateLic)
    cboStateLic.ListIndex = -1
    imStateLicIndex = -1
    If Trim$(rst!shttStateLic) <> "" Then
        For ilLoop = 0 To cboStateLic.ListCount - 1 Step 1
            If StrComp(Trim$(rst!shttStateLic), Trim$(tgStateInfo(cboStateLic.ItemData(ilLoop)).sPostalName), vbTextCompare) = 0 Then
                cboStateLic.ListIndex = ilLoop
                cboStateLic.Text = cboStateLic.List(ilLoop)
                imStateLicIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    txtSerialNo1.Text = Trim$(rst!shttSerialNo1)
    txtSerialNo2.Text = Trim$(rst!shttSerialNo2)
    'txtPort.text = Trim$(rst!shttPort)
    txtWebAddress.Text = Trim$(rst!shttWebAddress)
    edcWebNumber.Text = Trim$(rst!shttWebNumber)
    If Trim$(rst!shttWebNumber) = "2" Then   '2 = new web site, anything else = 1
'        rbcWebSiteVersion(1).Value = True
        smOrigWebNumber = "2"
        edcWebNumber.Text = "2"
    Else
'        rbcWebSiteVersion(0).Value = True
        smOrigWebNumber = "1"
        edcWebNumber.Text = "1"
    End If
    'txtSerialNo1.Text = ""
    'txtSerialNo2.Text = ""
    txtPort.Text = ""
    
    SQLQuery = "Select mktName, mktRank from Mkt where mktCode = " & rst!shttMktCode
    Set temp_rst = gSQLSelectCall(SQLQuery)
    If Not temp_rst.EOF Then
        txtMarket.Text = Trim$(temp_rst!mktName)  '(24).Value)
        If temp_rst!mktRank = 0 Then
            txtMarkRank.Text = ""
        Else
            txtMarkRank.Text = temp_rst!mktRank
        End If
    Else
        txtMarkRank.Text = ""
        txtMarket.Text = ""
    End If
    '5/10/07:  Removed Affiliate Rep from Station File
    'txtARPhone.Text = ""
    txtONAddr1.Text = Trim$(rst!shttONAddress1) 'rst(27).Value)
    txtONAddr2.Text = Trim$(rst!shttONAddress2) 'rst(28).Value)
    'txtONCity.Text = Trim$(rst!shttOnCity) 'rst(29).Value)
    cbcONCity.SetListIndex = -1
    If rst!shttONCityMntCode > 0 Then
        For ilLoop = 0 To cbcONCity.ListCount - 1 Step 1
            If cbcONCity.GetItemData(ilLoop) = rst!shttONCityMntCode Then
                cbcONCity.SetListIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    'txtONState.text = Trim$(rst!shttONState) 'rst(30).Value)
    cboONState.ListIndex = -1
    imOnStateIndex = -1
    If Trim$(rst!shttONState) <> "" Then
        For ilLoop = 0 To cboONState.ListCount - 1 Step 1
            If StrComp(Trim$(rst!shttONState), Trim$(tgStateInfo(cboONState.ItemData(ilLoop)).sPostalName), vbTextCompare) = 0 Then
                cboONState.ListIndex = ilLoop
                cboONState.Text = cboONState.List(ilLoop)
                imOnStateIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    txtONZip.Text = Trim$(rst!shttOnZip) 'rst(31).Value)
    
'    slTemp = Trim$(rst!shttWebEmail)
'    ilRet = gParseItem(slTemp, 1, ",", smShttWebEmail)
'    txtWebEmail.text = smShttWebEmail
'
'
    txtWebPW.Text = Trim$(rst!shttWebPW)
    'D.S. 07/05/01 moved 5 lines of code from below to here
    'If rst!shttAckDaylight = 0 Then
    '    optDaylight(0).Value = True
    'Else
    '    optDaylight(1).Value = True
    'End If
    If rst!shttAckDaylight = 0 Then
        lacDaylight.Caption = "Honor Daylight Savings"
        lacDaylight.BackColor = GREEN
        '10197
        imExistingHonorDaylight = 0
    Else
        lacDaylight.Caption = "Ignore Daylight Savings"
        lacDaylight.BackColor = vbRed
        '10197
        imExistingHonorDaylight = 1
    End If
    If rst!shttOnAir = "N" Then
        lacOnAir.Caption = "Off Air"
        lacOnAir.BackColor = vbRed
    Else
        lacOnAir.Caption = "On Air"
        lacOnAir.BackColor = GREEN
    End If
    If rst!shttStationType = "N" Then
        lacCommercial.Caption = "Non-Commercial"
        lacCommercial.BackColor = vbRed
    Else
        lacCommercial.Caption = "Commercial"
        lacCommercial.BackColor = GREEN
    End If
    edcP12Plus.Text = Format(rst!shttAudP12Plus, "##,###,###")
    If (DateValue(gAdjYear(rst!shttHistStartDate)) = DateValue("1/1/1970")) Then  'Or (rst!attOnAir = "1/1/70") Then
        edcHistStartDate.Text = ""
    Else
        edcHistStartDate.Text = Format$(rst!shttHistStartDate, sgShowDateForm)
    End If
    edcWatts.Text = Format(rst!shttWatts, "##,###,###")
    
    If rst!shttUsedForAtt = "N" Then
        ckcUsedFor(0).Value = vbUnchecked
    Else
        ckcUsedFor(0).Value = vbChecked
    End If
    If rst!shttUsedForXDigital = "Y" Then
        ckcUsedFor(1).Value = vbChecked
    Else
        ckcUsedFor(1).Value = vbUnchecked
    End If
    If rst!shttUsedForWegener = "Y" Then
        ckcUsedFor(2).Value = vbChecked
    Else
        ckcUsedFor(2).Value = vbUnchecked
    End If
    If rst!shttUsedForOLA = "Y" Then
        ckcUsedFor(3).Value = vbChecked
    Else
        ckcUsedFor(3).Value = vbUnchecked
    End If
    If rst!shttPledgeVsAir = "Y" Then
        ckcUsedFor(4).Value = vbChecked
    Else
        ckcUsedFor(4).Value = vbUnchecked
    End If
    
    txtIPumpID.Text = ""
    If smWegenerIPump = "Y" Then
        txtIPumpID.Text = Trim$(rst!shttIPumpID)
    End If
    If Val(rst!shttSpotsPerWebPage) <= 0 Then
        txtWebSpotsPerPage.Text = ""
    Else
        txtWebSpotsPerPage.Text = rst!shttSpotsPerWebPage
    End If
    
    '5/10/07:  Removed Affiliate Rep from Station File
    'If rst!shttArttCode <= 0 Then   'rst(1).Value <= 0 Then
    '    'Select No Rep
    '    For ilLoop = 0 To cboAffRep1.ListCount - 1
    '        If cboAffRep1.ItemData(ilLoop) = 0 Then
    '            cboAffRep1.ListIndex = 1
    '            txtARPhone.Text = ""
    '            Exit For
    '        End If
    '    Next ilLoop
    'Else
    '    For ilLoop = 0 To cboAffRep1.ListCount - 1
    '        If rst!shttArttCode = CInt(cboAffRep1.ItemData(ilLoop)) Then
    '            SQLQuery = "SELECT arttPhone FROM artt WHERE arttCode = '" & cboAffRep1.ItemData(ilLoop) & "'"
    '            Set rst = gSQLSelectCall(SQLQuery)
    '            txtARPhone.Text = Trim$(rst!arttPhone) 'rst(0).Value
    '            cboAffRep1.ListIndex = ilLoop
    '            Exit For
    '        End If
    '    Next ilLoop
    'End If

    'imFormatIndex = -1
    'cboFormat.ListIndex = -1
    'cboFormat.Text = ""
    'If rst!shttFmtCode > 0 Then
    '    For ilLoop = 0 To cboFormat.ListCount - 1 Step 1
    '        If rst!shttFmtCode = cboFormat.ItemData(ilLoop) Then
    '            cboFormat.ListIndex = ilLoop
    '            cboFormat.Text = cboFormat.List(ilLoop)
    '            imFormatIndex = ilLoop
    '            Exit For
    '        End If
    '    Next ilLoop
    'End If

    'D.S. 07/05/01 moved 5 lines of code from here to above. rst!shttAckDaylight was no longer
    'valid after the call to artt above
    'If rst!shttAckDaylight = 0 Then
    '    optDaylight(0).Value = True
    'Else
    '    optDaylight(1).Value = True
    'End If
    
    'Initialize prior to Owner and DMA avoid populate routines processing as Owner and DMA set
    lmMultiCastGroupID = 0
    rbcMulticast(0).Value = False
    rbcMulticast(1).Value = False
    rbcMulticast(2).Value = False
    rbcMulticast(0).Enabled = True
    rbcMulticast(1).Enabled = True
    rbcMulticast(2).Enabled = True
    
    lmClusterGroupID = 0
    smClusterStations = ""
    rbcMarketCluster(0).Value = False
    rbcMarketCluster(1).Value = False
    rbcMarketCluster(2).Value = False
    rbcMarketCluster(0).Enabled = True
    rbcMarketCluster(1).Enabled = True
    rbcMarketCluster(2).Enabled = True
    
    cbcOwner.SetListIndex = -1
    lmArttCode = rst!shttOwnerArttCode
    For llIdx = 2 To cbcOwner.ListCount - 1 Step 1
        If cbcOwner.GetItemData(CInt(llIdx)) = lmArttCode Then
            cbcOwner.SetListIndex = llIdx
            Exit For
        End If
    Next llIdx
    If cbcOwner.ListIndex > 1 Then
        lblOwner.Caption = "Owner: " & Trim$(cbcOwner.GetName(cbcOwner.ListIndex))
    Else
        lblOwner.Caption = "Owner: "
    End If
    
    lacDMA.Caption = "DMA:"
    cbcDMAMarket.SetListIndex = -1
    imDMAMktCode = rst!shttMktCode
    For llIdx = 2 To cbcDMAMarket.ListCount - 1 Step 1
        If cbcDMAMarket.GetItemData(CInt(llIdx)) = imDMAMktCode Then
            cbcDMAMarket.SetListIndex = llIdx
            Exit For
        End If
    Next llIdx
    For ilLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
        If imDMAMktCode = tgMarketInfo(ilLoop).lCode Then
            If tgMarketInfo(ilLoop).iRank <> 0 Then
                lacDMA.Caption = "DMA:" & tgMarketInfo(ilLoop).iRank
            End If
            Exit For
        End If
    Next ilLoop
    
    lacMSAMarket.Caption = "MSA:"
    cbcMSAMarket.SetListIndex = -1
    imMSAMktCode = rst!shttMetCode
    For llIdx = 2 To cbcMSAMarket.ListCount - 1 Step 1
        If cbcMSAMarket.GetItemData(CInt(llIdx)) = imMSAMktCode Then
            cbcMSAMarket.SetListIndex = llIdx
            Exit For
        End If
    Next llIdx
    For ilLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
        If imMSAMktCode = tgMSAMarketInfo(ilLoop).lCode Then
            If tgMSAMarketInfo(ilLoop).iRank <> 0 Then
                lacMSAMarket.Caption = "MSA:" & tgMSAMarketInfo(ilLoop).iRank
            End If
            Exit For
        End If
    Next ilLoop
    
    cbcOperator.SetListIndex = -1
    For llIdx = 2 To cbcOperator.ListCount - 1 Step 1
        If cbcOperator.GetItemData(CInt(llIdx)) = rst!shttOperatorMntCode Then
            cbcOperator.SetListIndex = llIdx
            Exit For
        End If
    Next llIdx
    
    cbcTerritory.SetListIndex = -1
    For llIdx = 2 To cbcTerritory.ListCount - 1 Step 1
        If cbcTerritory.GetItemData(CInt(llIdx)) = rst!shttMntCode Then
            cbcTerritory.SetListIndex = llIdx
            Exit For
        End If
    Next llIdx
    
    cbcFormat.SetListIndex = -1
    For llIdx = 2 To cbcFormat.ListCount - 1 Step 1
        If cbcFormat.GetItemData(CInt(llIdx)) = rst!shttFmtCode Then
            cbcFormat.SetListIndex = llIdx
            Exit For
        End If
    Next llIdx
    
    
    cbcArea.SetListIndex = -1
    For llIdx = 2 To cbcArea.ListCount - 1 Step 1
        If cbcArea.GetItemData(CInt(llIdx)) = rst!shttAreaMntCode Then
            cbcArea.SetListIndex = llIdx
            Exit For
        End If
    Next llIdx
    
    cbcMarketRep.SetListIndex = -1
    For llIdx = 1 To cbcMarketRep.ListCount - 1 Step 1
        If cbcMarketRep.GetItemData(CInt(llIdx)) = rst!shttMktRepUstCode Then
            cbcMarketRep.SetListIndex = llIdx
            Exit For
        End If
    Next llIdx
    
    cbcServiceRep.SetListIndex = -1
    For llIdx = 1 To cbcServiceRep.ListCount - 1 Step 1
        If cbcServiceRep.GetItemData(CInt(llIdx)) = rst!shttServRepUstCode Then
            cbcServiceRep.SetListIndex = llIdx
            Exit For
        End If
    Next llIdx
    
    edcEnterpriseID.Text = Trim$(rst!shttVieroID)
    
    ''Doug- On 11/17/06 added these three calls
    'ilRet = mDMAMarketFillListBox(0, False, imDMAMktCode)
    'ilRet = mDMAMarketFillStation(lmArttCode, imDMAMktCode)
    'ilRet = mMSAMarketFillListBox(0, False, imMSAMktCode)
    'ilRet = mMSAMarketFillStation(lmArttCode, imMSAMktCode)
    'ilRet = mOwnerFillListBox(0, False, lmArttCode)
    'mMulticastInit
    lblOwner.Caption = ""
    If cbcOwner.ListIndex > 1 Then
        lblOwner.Caption = "Owner: " & Trim$(cbcOwner.GetName(cbcOwner.ListIndex))
    End If
    If cbcDMAMarket.ListIndex > 1 Then
        lblOwner.Caption = lblOwner.Caption & " DMA Market: " & cbcDMAMarket.GetName(cbcDMAMarket.ListIndex)
    End If
    
    'Initialization move prior to Owner and DMA being set
    'Left here just in case values set
    lmMultiCastGroupID = 0
    rbcMulticast(0).Value = False
    rbcMulticast(1).Value = False
    rbcMulticast(2).Value = False
    rbcMulticast(0).Enabled = True
    rbcMulticast(1).Enabled = True
    rbcMulticast(2).Enabled = True
    lacMulticastNote.Visible = False
    If Not IsNull(rst!shttMultiCastGroupID) Then
        If rst!shttMultiCastGroupID <> 0 Then
            lmMultiCastGroupID = rst!shttMultiCastGroupID
            rbcMulticast(2).Value = True
            rbcMulticast(0).Enabled = False
            'Jim 1/22/11: Allow stations not in group to be add to current group
            'rbcMulticast(1).Enabled = False
        Else
            'rbcMulticast(0).Value = True
            rbcMulticast(2).Enabled = False
            rbcMulticastOwner(0).Value = True
            rbcMulticastMarket(0).Value = True
        End If
    Else
        'rbcMulticast(0).Value = True
        rbcMulticast(2).Enabled = False
        rbcMulticastOwner(0).Value = True
        rbcMulticastMarket(0).Value = True
    End If
        
    'Initialization move prior to Owner and DMA being set
    'Left here just in case values set
    lmClusterGroupID = 0
    smClusterStations = ""
    rbcMarketCluster(0).Value = False
    rbcMarketCluster(1).Value = False
    rbcMarketCluster(2).Value = False
    rbcMarketCluster(0).Enabled = True
    rbcMarketCluster(1).Enabled = True
    rbcMarketCluster(2).Enabled = True
    lacMarketClusterNote.Visible = False
    If Not IsNull(rst!shttclustergroupId) Then
        If rst!shttclustergroupId <> 0 Then
            lmClusterGroupID = rst!shttclustergroupId
            rbcMarketCluster(2).Value = True
            rbcMarketCluster(0).Enabled = False
            'Jim 1/22/11: Allow stations not in group to be add to current group
            'rbcMarketCluster(1).Enabled = False
        Else
            'rbcMarketCluster(0).Value = True
            rbcMarketCluster(2).Enabled = False
            rbcMarketClusterMarket(0).Value = True
        End If
    Else
        'rbcMarketCluster(0).Value = True
        rbcMarketCluster(2).Enabled = False
        rbcMarketClusterMarket(0).Value = True
    End If
    mPopMulticast
    mPopSisterStations
    mGetHistory
    'mGetPersonnel
    udcContactGrid.StationCode = imShttCode
    udcContactGrid.Action 3 'populate
    imFieldChgd = False
    'ilRet = gSaveCurrShttState(imShttCode)
    smOldMaster = edcMasterStation.Text
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-mBindControls"
End Sub



Private Sub cbcArea_DblClick()
    Dim ilMnt As Integer
    sgMultiNameType = "A"
    If cbcArea.ListIndex > 1 Then
        sgMultiNameName = cbcArea.GetName(cbcArea.ListIndex)
    Else
        sgMultiNameName = ""
    End If
    frmMultiName.Show vbModal
    mPopArea
    If igMultiNameReturn Then
        For ilMnt = 0 To cbcArea.ListCount - 1 Step 1
            If cbcArea.GetItemData(ilMnt) = lgMultiNameCode Then
                cbcArea.SetListIndex = ilMnt
                Exit Sub
            End If
        Next ilMnt
    End If
    cbcArea.SetListIndex = 1
End Sub

Private Sub cbcArea_GotFocus()
    imIgnoreTabs = False
    cbcArea.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcArea_LostFocus()
    Dim ilMnt As Integer
    
    imMntIndex = cbcArea.ListIndex
    If imMntIndex = 0 Then
        If Not frmMultiName.Visible Then
            sgMultiNameType = "A"
            sgMultiNameName = ""
            frmMultiName.Show vbModal
            mPopArea
            If igMultiNameReturn Then
                For ilMnt = 0 To cbcArea.ListCount - 1 Step 1
                    If cbcArea.GetItemData(ilMnt) = lgMultiNameCode Then
                        cbcArea.SetListIndex = ilMnt
                        Exit Sub
                    End If
                Next ilMnt
            End If
            cbcArea.SetListIndex = 1
        End If
    End If
End Sub

Private Sub cbcArea_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcArea_ReSetLoc()
    cbcArea.Top = lacOperator.Top - 45 + cbcONCity.Height - cbcArea.Height
End Sub

Private Sub cbcCity_DblClick()
    Dim ilMnt As Integer
    sgMultiNameType = "C"
    If cbcCity.ListIndex > 1 Then
        sgMultiNameName = cbcCity.GetName(cbcCity.ListIndex)
    Else
        sgMultiNameName = ""
    End If
    frmMultiName.Show vbModal
    mPopCity
    If igMultiNameReturn Then
        For ilMnt = 0 To cbcCity.ListCount - 1 Step 1
            If cbcCity.GetItemData(ilMnt) = lgMultiNameCode Then
                cbcCity.SetListIndex = ilMnt
                Exit Sub
            End If
        Next ilMnt
    End If
    cbcCity.SetListIndex = 1
End Sub

Private Sub cbcCity_GotFocus()
    imIgnoreTabs = False
    cbcCity.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcCity_LostFocus()
    Dim ilMnt As Integer
    
    imMntIndex = cbcCity.ListIndex
    If imMntIndex = 0 Then
        If Not frmMultiName.Visible Then
            sgMultiNameType = "C"
            sgMultiNameName = ""
            frmMultiName.Show vbModal
            mPopCity
            If igMultiNameReturn Then
                For ilMnt = 0 To cbcCity.ListCount - 1 Step 1
                    If cbcCity.GetItemData(ilMnt) = lgMultiNameCode Then
                        cbcCity.SetListIndex = ilMnt
                        Exit Sub
                    End If
                Next ilMnt
            End If
            cbcCity.SetListIndex = 1
        End If
    End If
End Sub

Private Sub cbcCity_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcCityLic_DblClick()
    Dim ilMnt As Integer
    sgMultiNameType = "C"
    If cbcCityLic.ListIndex > 1 Then
        sgMultiNameName = cbcCityLic.GetName(cbcCityLic.ListIndex)
    Else
        sgMultiNameName = ""
    End If
    frmMultiName.Show vbModal
    mPopCity
    If igMultiNameReturn Then
        For ilMnt = 0 To cbcCityLic.ListCount - 1 Step 1
            If cbcCityLic.GetItemData(ilMnt) = lgMultiNameCode Then
                cbcCityLic.SetListIndex = ilMnt
                Exit Sub
            End If
        Next ilMnt
    End If
    cbcCityLic.SetListIndex = 1
End Sub

Private Sub cbcCityLic_GotFocus()
    imIgnoreTabs = False
    cbcCityLic.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcCityLic_LostFocus()
    Dim ilMnt As Integer
    
    If bmCancelled Then
        cmdCancel_Click
        Exit Sub
    End If
    imMntIndex = cbcCityLic.ListIndex
    If imMntIndex = 0 Then
        If Not frmMultiName.Visible Then
            sgMultiNameType = "C"
            sgMultiNameName = ""
            frmMultiName.Show vbModal
            mPopCity
            If igMultiNameReturn Then
                For ilMnt = 0 To cbcCityLic.ListCount - 1 Step 1
                    If cbcCityLic.GetItemData(ilMnt) = lgMultiNameCode Then
                        cbcCityLic.SetListIndex = ilMnt
                        Exit Sub
                    End If
                Next ilMnt
            End If
            cbcCityLic.SetListIndex = 1
        End If
    End If
End Sub

Private Sub cbcCityLic_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcCountyLic_DblClick()
    Dim ilMnt As Integer
    sgMultiNameType = "Y"
    If cbcCountyLic.ListIndex > 1 Then
        sgMultiNameName = cbcCountyLic.GetName(cbcCountyLic.ListIndex)
    Else
        sgMultiNameName = ""
    End If
    frmMultiName.Show vbModal
    mPopCounty
    If igMultiNameReturn Then
        For ilMnt = 0 To cbcCountyLic.ListCount - 1 Step 1
            If cbcCountyLic.GetItemData(ilMnt) = lgMultiNameCode Then
                cbcCountyLic.SetListIndex = ilMnt
                Exit Sub
            End If
        Next ilMnt
    End If
    cbcCountyLic.SetListIndex = 1
End Sub

Private Sub cbcCountyLic_GotFocus()
    imIgnoreTabs = False
    cbcCountyLic.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcCountyLic_LostFocus()
    Dim ilMnt As Integer
    
    If bmCancelled Then
        cmdCancel_Click
        Exit Sub
    End If
    imMntIndex = cbcCountyLic.ListIndex
    If imMntIndex = 0 Then
        If Not frmMultiName.Visible Then
            sgMultiNameType = "Y"
            sgMultiNameName = ""
            frmMultiName.Show vbModal
            mPopCounty
            If igMultiNameReturn Then
                For ilMnt = 0 To cbcCountyLic.ListCount - 1 Step 1
                    If cbcCountyLic.GetItemData(ilMnt) = lgMultiNameCode Then
                        cbcCountyLic.SetListIndex = ilMnt
                        Exit Sub
                    End If
                Next ilMnt
            End If
            cbcCountyLic.SetListIndex = 1
        End If
    End If
End Sub

Private Sub cbcCountyLic_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcDMAMarket_DblClick()
    Dim ilRet As Integer
    frmStationMktInfo.Show vbModal
    If igMarketReturnCode = -1 Then
        mDMAMarketFillListBox 0, False, imDMAMktCode
    Else
        mDMAMarketFillListBox 0, False, igMarketReturnCode
    End If
End Sub

Private Sub cbcDMAMarket_GotFocus()
    imIgnoreTabs = False
    cbcDMAMarket.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcDMAMarket_LostFocus()
    If cbcDMAMarket.ListIndex = 0 Then
        If Not frmStationMktInfo.Visible Then
            frmStationMktInfo.Show vbModal
            If igMarketReturnCode = -1 Then
                mDMAMarketFillListBox 0, False, imDMAMktCode
            Else
                mDMAMarketFillListBox 0, False, igMarketReturnCode
            End If
        End If
    End If
End Sub

Private Sub cbcDMAMarket_OnChange()
    Dim ilLoop As Integer
    
    lacDMA.Caption = "DMA:"
    If cbcDMAMarket.ListIndex >= 0 Then
        If imDMAMktCode <> cbcDMAMarket.GetItemData(cbcDMAMarket.ListIndex) Then
            imFieldChgd = True
            If ((rbcMulticast(0).Value) Or (rbcMulticast(1).Value)) And (rbcMulticastMarket(0).Value) Then
                mPopMulticast
            End If
            If ((rbcMarketCluster(0).Value) Or (rbcMarketCluster(1).Value)) And (rbcMarketClusterMarket(0).Value) Then
                mPopSisterStations
            End If
            For ilLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
                If cbcDMAMarket.GetItemData(cbcDMAMarket.ListIndex) = tgMarketInfo(ilLoop).lCode Then
                    If tgMarketInfo(ilLoop).iRank <> 0 Then
                        lacDMA.Caption = "DMA:" & tgMarketInfo(ilLoop).iRank
                    End If
                    Exit For
                End If
            Next ilLoop
        End If
    Else
        If imDMAMktCode <> 0 Then
            imFieldChgd = True
            mPopMulticast
            mPopSisterStations
        End If
    End If
End Sub

Private Sub cbcDMAMarket_ReSetLoc()
    cbcDMAMarket.Top = lacOwner.Top - 45 + cbcONCity.Height - cbcDMAMarket.Height
End Sub

Private Sub cbcFormat_DblClick()
    Dim ilFmt As Integer
    sgFormatCall = "S"
    If cbcFormat.ListIndex > 1 Then
        sgFormatName = cbcFormat.GetName(cbcFormat.ListIndex)
    Else
        sgFormatName = ""
    End If
    frmGroupNameFormat.Show vbModal
    mPopFormat
    If igFormatReturn Then
        For ilFmt = 0 To cbcFormat.ListCount - 1 Step 1
            If cbcFormat.GetItemData(ilFmt) = igFormatReturnCode Then
                cbcFormat.SetListIndex = ilFmt
                Exit Sub
            End If
        Next ilFmt
    End If
    cbcFormat.SetListIndex = 1
End Sub

Private Sub cbcFormat_GotFocus()
    imIgnoreTabs = False
    cbcFormat.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcFormat_LostFocus()
    Dim ilFmt As Integer
    
    imFormatIndex = cbcFormat.ListIndex
    If imFormatIndex = 0 Then
        If Not frmGroupNameFormat.Visible Then
            sgFormatCall = "S"
            sgFormatName = ""
            frmGroupNameFormat.Show vbModal
            mPopFormat
            If igFormatReturn Then
                For ilFmt = 0 To cbcFormat.ListCount - 1 Step 1
                    If cbcFormat.GetItemData(ilFmt) = igFormatReturnCode Then
                        cbcFormat.SetListIndex = ilFmt
                        Exit Sub
                    End If
                Next ilFmt
            End If
            cbcFormat.SetListIndex = 1
        End If
    End If
End Sub

Private Sub cbcFormat_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcMarketRep_GotFocus()
    imIgnoreTabs = False
    cbcMarketRep.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcMarketRep_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcMarketRep_ReSetLoc()
    cbcMarketRep.Top = lacMktRep.Top - 45 + cbcONCity.Height - cbcMarketRep.Height
End Sub

Private Sub cbcMoniker_DblClick()
    Dim ilMnt As Integer
    sgMultiNameType = "M"
    If cbcMoniker.ListIndex > 1 Then
        sgMultiNameName = cbcMoniker.GetName(cbcMoniker.ListIndex)
    Else
        sgMultiNameName = ""
    End If
    frmMultiName.Show vbModal
    mPopMoniker
    If igMultiNameReturn Then
        For ilMnt = 0 To cbcMoniker.ListCount - 1 Step 1
            If cbcMoniker.GetItemData(ilMnt) = lgMultiNameCode Then
                cbcMoniker.SetListIndex = ilMnt
                Exit Sub
            End If
        Next ilMnt
    End If
    cbcMoniker.SetListIndex = 1
End Sub

Private Sub cbcMoniker_GotFocus()
    imIgnoreTabs = False
    cbcMoniker.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcMoniker_LostFocus()
    Dim ilMnt As Integer
    
    imMntIndex = cbcMoniker.ListIndex
    If imMntIndex = 0 Then
        If Not frmMultiName.Visible Then
            sgMultiNameType = "M"
            sgMultiNameName = ""
            frmMultiName.Show vbModal
            mPopMoniker
            If igMultiNameReturn Then
                For ilMnt = 0 To cbcMoniker.ListCount - 1 Step 1
                    If cbcMoniker.GetItemData(ilMnt) = lgMultiNameCode Then
                        cbcMoniker.SetListIndex = ilMnt
                        Exit Sub
                    End If
                Next ilMnt
            End If
            cbcMoniker.SetListIndex = 1
        End If
    End If
End Sub

Private Sub cbcMoniker_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcMSAMarket_DblClick()
    Dim ilRet As Integer
    frmStationMSAMktInfo.Show vbModal
    If igMarketReturnCode = -1 Then
        mMSAMarketFillListBox 0, False, imMSAMktCode
    Else
        mMSAMarketFillListBox 0, False, igMarketReturnCode
    End If
End Sub

Private Sub cbcMSAMarket_GotFocus()
    imIgnoreTabs = False
    cbcMSAMarket.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcMSAMarket_LostFocus()
    If cbcMSAMarket.ListIndex = 0 Then
        If Not frmStationMSAMktInfo.Visible Then
            frmStationMSAMktInfo.Show vbModal
            If igMarketReturnCode = -1 Then
                mMSAMarketFillListBox 0, False, imMSAMktCode
            Else
                mMSAMarketFillListBox 0, False, igMarketReturnCode
            End If
        End If
    End If
End Sub

Private Sub cbcMSAMarket_OnChange()
    Dim ilLoop As Integer
    
    imFieldChgd = True
    lacMSAMarket.Caption = "MSA:"
    If cbcMSAMarket.ListIndex >= 0 Then
        For ilLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
            If cbcMSAMarket.GetItemData(cbcMSAMarket.ListIndex) = tgMSAMarketInfo(ilLoop).lCode Then
                If tgMSAMarketInfo(ilLoop).iRank <> 0 Then
                    lacMSAMarket.Caption = "MSA:" & tgMSAMarketInfo(ilLoop).iRank
                End If
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub cbcMSAMarket_ReSetLoc()
    cbcMSAMarket.Top = lacOwner.Top - 45 + cbcONCity.Height - cbcMSAMarket.Height
End Sub

Private Sub cbcOnCity_DblClick()
    Dim ilMnt As Integer
    sgMultiNameType = "C"
    If cbcONCity.ListIndex > 1 Then
        sgMultiNameName = cbcONCity.GetName(cbcONCity.ListIndex)
    Else
        sgMultiNameName = ""
    End If
    frmMultiName.Show vbModal
    mPopCity
    If igMultiNameReturn Then
        For ilMnt = 0 To cbcONCity.ListCount - 1 Step 1
            If cbcONCity.GetItemData(ilMnt) = lgMultiNameCode Then
                cbcONCity.SetListIndex = ilMnt
                Exit Sub
            End If
        Next ilMnt
    End If
    cbcONCity.SetListIndex = 1
End Sub

Private Sub cbcOnCity_GotFocus()
    imIgnoreTabs = False
    cbcONCity.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcOnCity_LostFocus()
    Dim ilMnt As Integer
    
    imMntIndex = cbcONCity.ListIndex
    If imMntIndex = 0 Then
        If Not frmMultiName.Visible Then
            sgMultiNameType = "C"
            sgMultiNameName = ""
            frmMultiName.Show vbModal
            mPopCity
            If igMultiNameReturn Then
                For ilMnt = 0 To cbcONCity.ListCount - 1 Step 1
                    If cbcONCity.GetItemData(ilMnt) = lgMultiNameCode Then
                        cbcONCity.SetListIndex = ilMnt
                        Exit Sub
                    End If
                Next ilMnt
            End If
            cbcONCity.SetListIndex = 1
        End If
    End If
End Sub

Private Sub cbcOnCity_OnChange()
    imFieldChgd = True
End Sub



Private Sub cbcOperator_DblClick()
    Dim ilMnt As Integer
    sgMultiNameType = "O"
    If cbcOperator.ListIndex > 1 Then
        sgMultiNameName = cbcOperator.GetName(cbcOperator.ListIndex)
    Else
        sgMultiNameName = ""
    End If
    frmMultiName.Show vbModal
    mPopOperator
    If igMultiNameReturn Then
        For ilMnt = 0 To cbcOperator.ListCount - 1 Step 1
            If cbcOperator.GetItemData(ilMnt) = lgMultiNameCode Then
                cbcOperator.SetListIndex = ilMnt
                Exit Sub
            End If
        Next ilMnt
    End If
    cbcOperator.SetListIndex = 1
End Sub

Private Sub cbcOperator_GotFocus()
    imIgnoreTabs = False
    cbcOperator.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcOperator_LostFocus()
    Dim ilMnt As Integer
    
    imMntIndex = cbcOperator.ListIndex
    If imMntIndex = 0 Then
        If Not frmMultiName.Visible Then
            sgMultiNameType = "O"
            sgMultiNameName = ""
            frmMultiName.Show vbModal
            mPopOperator
            If igMultiNameReturn Then
                For ilMnt = 0 To cbcOperator.ListCount - 1 Step 1
                    If cbcOperator.GetItemData(ilMnt) = lgMultiNameCode Then
                        cbcOperator.SetListIndex = ilMnt
                        Exit Sub
                    End If
                Next ilMnt
            End If
            cbcOperator.SetListIndex = 1
        End If
    End If
End Sub

Private Sub cbcOperator_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcOperator_ReSetLoc()
    cbcOperator.Top = lacOperator.Top - 45 + cbcONCity.Height - cbcOperator.Height
End Sub

Private Sub cbcOwner_DblClick()
    Dim ilRet As Integer
    frmStationOwnerInfo.Show vbModal
    If lgOwnerReturnCode = -1 Then
        mOwnerFillListBox 0, False, lmArttCode
    Else
        mOwnerFillListBox 0, False, lgOwnerReturnCode
        'ilRet = mDMAMarketFillStation(lmArttCode, imDMAMktCode)
        'ilRet = mMSAMarketFillStation(lmArttCode, imMSAMktCode)
    End If
End Sub

Private Sub cbcOwner_GotFocus()
    imIgnoreTabs = False
    cbcOwner.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcOwner_LostFocus()
    Dim ilArtt As Integer
    Dim llArttCode As Long
    Dim ilRet As Integer
    If bmCancelled Then
        cmdCancel_Click
        Exit Sub
    End If
    If cbcOwner.ListIndex = 0 Then
        If Not frmStationOwnerInfo.Visible Then
            frmStationOwnerInfo.Show vbModal
            If lgOwnerReturnCode = -1 Then
                mOwnerFillListBox 0, False, lmArttCode
            Else
                mOwnerFillListBox 0, False, lgOwnerReturnCode
                'ilRet = mDMAMarketFillStation(lmArttCode, imDMAMktCode)
                'ilRet = mMSAMarketFillStation(lmArttCode, imMSAMktCode)
            End If
        End If
    End If
End Sub

Private Sub cbcOwner_OnChange()
    If bmIgnoreOwnerChange Then
        Exit Sub
    End If
    If cbcOwner.ListIndex >= 0 Then
        If lmArttCode <> cbcOwner.GetItemData(cbcOwner.ListIndex) Then
            imFieldChgd = True
            If ((rbcMulticast(0).Value) Or (rbcMulticast(1).Value)) And (rbcMulticastOwner(0).Value) Then
                mPopMulticast
            End If
            mPopSisterStations
        End If
    Else
        If lmArttCode <> 0 Then
            imFieldChgd = True
            mPopMulticast
            mPopSisterStations
        End If
    End If
    If cbcOwner.ListIndex >= 0 Then
        lmArttCode = cbcOwner.GetItemData(cbcOwner.ListIndex)
    Else
        lmArttCode = 0
    End If
End Sub

Private Sub cbcOwner_ReSetLoc()
    cbcOwner.Top = lacOwner.Top - 45 + cbcONCity.Height - cbcOwner.Height
End Sub

Private Sub cbcServiceRep_GotFocus()
    imIgnoreTabs = False
    cbcServiceRep.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcServiceRep_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcServiceRep_ReSetLoc()
    cbcServiceRep.Top = lacMktRep.Top - 45 + cbcONCity.Height - cbcServiceRep.Height
End Sub

Private Sub cbcTerritory_DblClick()
    Dim ilMnt As Integer
    sgMultiNameType = "T"
    If cbcTerritory.ListIndex > 1 Then
        sgMultiNameName = cbcTerritory.GetName(cbcTerritory.ListIndex)
    Else
        sgMultiNameName = ""
    End If
    frmMultiName.Show vbModal
    mPopTerritory
    If igMultiNameReturn Then
        For ilMnt = 0 To cbcTerritory.ListCount - 1 Step 1
            If cbcTerritory.GetItemData(ilMnt) = lgMultiNameCode Then
                cbcTerritory.SetListIndex = ilMnt
                Exit Sub
            End If
        Next ilMnt
    End If
    cbcTerritory.SetListIndex = 1
End Sub

Private Sub cbcTerritory_GotFocus()
    imIgnoreTabs = False
    cbcTerritory.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcTerritory_LostFocus()
    Dim ilMnt As Integer
    
    imMntIndex = cbcTerritory.ListIndex
    If imMntIndex = 0 Then
        If Not frmMultiName.Visible Then
            sgMultiNameType = "T"
            sgMultiNameName = ""
            frmMultiName.Show vbModal
            mPopTerritory
            If igMultiNameReturn Then
                For ilMnt = 0 To cbcTerritory.ListCount - 1 Step 1
                    If cbcTerritory.GetItemData(ilMnt) = lgMultiNameCode Then
                        cbcTerritory.SetListIndex = ilMnt
                        Exit Sub
                    End If
                Next ilMnt
            End If
            cbcTerritory.SetListIndex = 1
        End If
    End If
End Sub

Private Sub cbcTerritory_OnChange()
    imFieldChgd = True
End Sub

Private Sub cbcTerritory_ReSetLoc()
    'Display list abobe edit box
    cbcTerritory.Top = lacOperator.Top - 45 + cbcONCity.Height - cbcTerritory.Height
End Sub


Private Sub cbcTimeZone_GotFocus()
    imIgnoreTabs = False
    cbcTimeZone.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcTimeZone_OnChange()
    imFieldChgd = True
    imTimeZoneIndex = cbcTimeZone.ListIndex
End Sub

Private Sub cbcTimeZone_ReSetLoc()
    cbcTimeZone.Top = lacMktRep.Top - 45 + cbcONCity.Height - cbcTimeZone.Height
End Sub

Private Sub cboOnState_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long

    If imInSave = True Then
        Exit Sub
    End If
    If imInChg Then
        Exit Sub
    End If
    imInChg = True

    Screen.MousePointer = vbHourglass
    sName = LTrim$(cboONState.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    lRow = SendMessageByString(cboONState.hwnd, CB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        cboONState.ListIndex = lRow
        cboONState.SelStart = iLen
        cboONState.SelLength = Len(cboONState.Text)
        imOnStateIndex = lRow
    Else
        imOnStateIndex = -1
    End If
    imFieldChgd = True
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
End Sub

Private Sub cboOnState_Click()
    cboOnState_Change
End Sub

Private Sub cboOnState_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cboOnState_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboOnState_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboONState.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub



Private Sub cboState_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long

    If imInSave = True Then
        Exit Sub
    End If
    If imInChg Then
        Exit Sub
    End If
    imInChg = True

    Screen.MousePointer = vbHourglass
    sName = LTrim$(cboState.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    lRow = SendMessageByString(cboState.hwnd, CB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        cboState.ListIndex = lRow
        cboState.SelStart = iLen
        cboState.SelLength = Len(cboState.Text)
        imStateIndex = lRow
    Else
        imStateIndex = -1
    End If
    imFieldChgd = True
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
End Sub

Private Sub cboState_Click()
    cboState_Change
End Sub

Private Sub cboState_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cboState_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboState_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboState.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cboStateLic_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long

    If imInSave = True Then
        Exit Sub
    End If
    If imInChg Then
        Exit Sub
    End If
    imInChg = True

    Screen.MousePointer = vbHourglass
    sName = LTrim$(cboStateLic.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    lRow = SendMessageByString(cboStateLic.hwnd, CB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        cboStateLic.ListIndex = lRow
        cboStateLic.SelStart = iLen
        cboStateLic.SelLength = Len(cboStateLic.Text)
        imStateLicIndex = lRow
    Else
        imStateLicIndex = -1
    End If
    imFieldChgd = True
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
End Sub

Private Sub cboStateLic_Click()
    cboStateLic_Change
End Sub

Private Sub cboStateLic_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cboStateLic_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboStateLic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboStateLic.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cboStateLic_LostFocus()
    If bmCancelled Then
        cmdCancel_Click
        Exit Sub
    End If
End Sub

Private Sub cboStations_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long

    If imInSave = True Then
        Exit Sub
    End If
    If imInChg Then
        Exit Sub
    End If
    imInChg = True

    Screen.MousePointer = vbHourglass
    sName = LTrim$(cboStations.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    lRow = SendMessageByString(cboStations.hwnd, CB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        On Error GoTo ErrHand
        cboStations.ListIndex = lRow
        cboStations.SelStart = iLen
        cboStations.SelLength = Len(cboStations.Text)
        If cboStations.ListIndex <= 0 Then
            imShttCode = 0
        Else
            imShttCode = CInt(cboStations.ItemData(cboStations.ListIndex))
        End If
        If imShttCode <= 0 Then
            mClearControls
            IsStatDirty = False
        Else                                                                'Load existing station data
            SQLQuery = "SELECT * "
            SQLQuery = SQLQuery + " FROM shtt"
            SQLQuery = SQLQuery + " WHERE (shttCode = " & imShttCode & ")"
            
            Set rst = gSQLSelectCall(SQLQuery)
            If rst.EOF Then
                gMsgBox "No matching records were found", vbOKOnly
                mClearControls
            Else
                mBindControls
            End If
            mSetTabColors
            IsStatDirty = True
        End If
    End If
    
    'TTP 9056 JJB 2023-05-12   This value was only getting assigned if the value of the textbox was changed OR if it lost focus to the textbox.
    smNewWebNumber = Trim(edcWebNumber.Text)
    
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-cboStations"
    imInChg = False
End Sub

Private Sub cboStations_Click()
    Dim ilFieldChg As Integer
    cboStations_Change
    'ilFieldChg = imFieldChgd
    'imScroll = 1
    'cboOwner_Change
    'imScroll = 1
    'cboDMAMarketCluster_Change
    'Call mDMAMarketFillStation(lmArttCode, imDMAMktCode)
    'imScroll = 1
    'cboMSAMarketCluster_Change
    'Call mMSAMarketFillStation(lmArttCode, imMSAMktCode)
    'imFieldChgd = ilFieldChg
    'mMulticastInit
    'If cboStations.text = "[New]" Then
    '    lbcStations(0).Clear
    '    lbcStations(1).Clear
    '    Screen.MousePointer = vbDefault
    '    Exit Sub
    'End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub cboStations_GotFocus()
    imIgnoreTabs = True
    mHistSetShow
    udcContactGrid.Action 1
End Sub

Private Sub cboStations_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboStations_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboStations.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub





Private Sub ckcUsedFor_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub cmdCancel_Click()
    Dim ilResponse As Integer
    
    bmCancelled = False
    If imFieldChgd Then
        ilResponse = gMsgBox("Changes were made! Are you sure you want to cancel? ", vbYesNo)
        If ilResponse = vbNo Then
            Exit Sub
        End If
    End If
    If imInSave = True Then
        Exit Sub
    End If
    'D.S. 02/04/03
    IsStatDirty = False
    Screen.MousePointer = vbHourglass
    Unload frmStation
End Sub

Private Sub cmdCancel_GotFocus()
    imIgnoreTabs = True
    mHistSetShow
    udcContactGrid.Action 1 'Clear focus
End Sub


Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bmCancelled = True
End Sub

Private Sub cmdDone_Click()
    
    Dim iRet As Integer
    
    If imInSave = True Then
        Exit Sub
    End If
    'D.S. TTP 9746 - 2/25/20 - Checks for any changes in Personnel
    If imFieldChgd = False Then
        imFieldChgd = udcContactGrid.AnyFieldChanged()
    End If
    Screen.MousePointer = vbHourglass
    If imFieldChgd = True Then
        If gMsgBox("Save all changes?", vbYesNo) = vbYes Then
            imInSave = True
            iRet = mSave(False)
            imInSave = False
            If Not iRet Then
                Screen.MousePointer = vbDefault
                Exit Sub    ' Dont exit until user takes care of whatever fields are invalid or missing.
            End If
        End If
    End If
    'D.S. 02/04/03
    IsStatDirty = False
    Screen.MousePointer = vbDefault
    Unload frmStation

    Exit Sub
End Sub

Private Sub cmdDone_GotFocus()

    imIgnoreTabs = True
    mHistSetShow
    udcContactGrid.Action 1 'Clear focus
End Sub


Private Sub cmdErase_Click()
    Dim iRet As Integer
    Dim iLoop As Integer
    Dim iIndex As Integer
    Dim slCallLetters As String
    
    If imInSave = True Then
        Exit Sub
    End If
    On Error GoTo ErrHand
    
    If sgUstWin(1) <> "I" Then
        gMsgBox "Not Allowed to Erase.", vbOKOnly
        Exit Sub
    End If

    If IsStatDirty = True Then
        SQLQuery = "SELECT Count(attCode) FROM att WHERE (attShfCode = " & Trim$(Str$(imShttCode)) & ")"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst(0).Value > 0 Then
            gMsgBox "Agreements must be Erased prior to Erasing Station.", vbOKOnly
            Exit Sub
        End If
        If optSP(1).Value Then
            iRet = gMsgBox("Remove Person?", vbYesNo)
        Else
            iRet = gMsgBox("Remove the Station?", vbYesNo)
        End If
        If iRet = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        
        bmDoPop = False
        If (sgShttTimeStamp <> gFileDateTime(sgDBPath & "Shtt.mkd")) Then
            bmDoPop = True
        End If
        
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If imShttCode = tgStationInfo(iLoop).iCode Then
                For iIndex = iLoop To UBound(tgStationInfo) - 2 Step 1
                    tgStationInfo(iIndex) = tgStationInfo(iIndex + 1)
                Next iIndex
                ReDim Preserve tgStationInfo(0 To UBound(tgStationInfo) - 1) As STATIONINFO
                Exit For
            End If
        Next iLoop
        For iLoop = 0 To UBound(tgStationInfoByCode) - 1 Step 1
            If imShttCode = tgStationInfoByCode(iLoop).iCode Then
                For iIndex = iLoop To UBound(tgStationInfoByCode) - 2 Step 1
                    tgStationInfoByCode(iIndex) = tgStationInfoByCode(iIndex + 1)
                Next iIndex
                ReDim Preserve tgStationInfoByCode(0 To UBound(tgStationInfoByCode) - 1) As STATIONINFO
                Exit For
            End If
        Next iLoop
        'cnn.BeginTrans
        If gGetEMailDistribution Then
            slCallLetters = gGetCallLettersByShttCode(imShttCode)
            iRet = gRemoveStationFromNetwork(slCallLetters)
        End If
         SQLQuery = "DELETE FROM clt WHERE (cltShfCode = " & imShttCode & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
            ''Screen.MousePointer = vbDefault
            ''Exit Sub
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "Station-cmdErase_Click"
            Exit Sub
        End If
        SQLQuery = "DELETE FROM artt WHERE (arttShttCode = " & imShttCode & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
            ''Screen.MousePointer = vbDefault
            ''Exit Sub
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "Station-cmdErase_Click"
            Exit Sub
        End If
        SQLQuery = "DELETE FROM shtt WHERE (shttCode = " & imShttCode & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        'cnn.CommitTrans
        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
            ''Screen.MousePointer = vbDefault
            ''Exit Sub
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "Station-cmdErase_Click"
            Exit Sub
        End If
        imShttCode = 0
        mSort
    End If
    
    Screen.MousePointer = vbDefault
    
    'D.S. 02/04/03
    IsStatDirty = False
    
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-cmdErase"
End Sub

Private Sub cmdErase_GotFocus()
    imIgnoreTabs = True
    mHistSetShow
    udcContactGrid.Action 1 'Clear focus
End Sub

Private Sub cmdGenPassword_Click()
    
    Dim slPassword As String
    
    slPassword = gGeneratePassword(4, 1)
    txtWebPW.Text = slPassword
    
End Sub




Private Sub cmdSave_Click()
    Dim iRet As Integer
    Dim iIndex As Integer
    

    If imInSave = True Then
        Exit Sub
    End If
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    DoEvents
    imInSave = True
    iRet = mSave(False)
    imInSave = False
    If iRet = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    mSort
      
'    SendKeys "%S"
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-cmdSave"
End Sub

Private Sub cmdSave_GotFocus()
    
    'Dim ilRet As Integer
   
    'ilRet = mSave(False)
    imIgnoreTabs = True
    mHistSetShow
    udcContactGrid.Action 1 'Clear focus
End Sub

Private Sub edcEnterpriseID_Change()
    imFieldChgd = True
End Sub

Private Sub edcEnterpriseID_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFrequency_Change()
    imFieldChgd = True
End Sub

Private Sub edcFrequency_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFrequency_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYDECPOINT) And (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub edcHistStartDate_Change()
    imFieldChgd = True
End Sub

Private Sub edcHistStartDate_GotFocus()
    imIgnoreTabs = False
    edcHistStartDate.ZOrder
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcMasterStation_Change()
    imFieldChgd = True
End Sub

Private Sub edcMasterStation_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcWatts_Change()
    imFieldChgd = True
End Sub

Private Sub edcWatts_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcWatts_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYCOMMA) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub edcWebNumber_Change()
    imFieldChgd = True
End Sub

Private Sub edcWebNumber_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcWebNumber_LostFocus()
    smNewWebNumber = Trim(edcWebNumber.Text)
    If (smNewWebNumber <> "1" And smNewWebNumber <> "2") Then
        MsgBox "Please enter a value of 1 or 2."
        edcWebNumber.SetFocus
    End If
End Sub

Private Sub Form_Click()
    mHistSetShow
    udcContactGrid.Action 1 'Clear focus
End Sub

Private Sub grdSisterStations_Click()
    Dim llRow As Long
    If grdSisterStations.MouseRow = 0 Then
        mSSSortCol grdSisterStations.MouseCol
    Else
        If (grdSisterStations.MouseRow >= grdSisterStations.FixedRows) And (grdSisterStations.MouseRow < grdSisterStations.Rows) Then
            If grdSisterStations.TextMatrix(grdSisterStations.MouseRow, SSCALLLETTERSINDEX) <> "" Then
                If grdSisterStations.TextMatrix(grdSisterStations.MouseRow, SSSELECTEDINDEX) <> "1" Then
                    grdSisterStations.TextMatrix(grdSisterStations.MouseRow, SSSELECTEDINDEX) = "1"
                    If rbcMarketCluster(1).Value Then
                        For llRow = grdSisterStations.FixedRows To grdSisterStations.Rows - 1 Step 1
                            If llRow <> grdSisterStations.MouseRow Then
                                If grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "1" Then
                                    grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "0"
                                    mSSPaintRowColor llRow
                                End If
                            End If
                        Next llRow
                    End If
                Else
                    grdSisterStations.TextMatrix(grdSisterStations.MouseRow, SSSELECTEDINDEX) = "0"
                End If
                imFieldChgd = True
                mSSPaintRowColor grdSisterStations.MouseRow
            End If
        End If
    End If
End Sub

Private Sub pbcHistFocus_Click()

End Sub

Private Sub rbcMarketCluster_Click(Index As Integer)
    If rbcMarketCluster(Index).Value Then
        bmIgnoreSisterStationChange = True
        If Index = 0 Then   'Create MarketCluster
            rbcMarketClusterMarket(0).Value = True
            rbcMarketClusterMarket(0).Enabled = True
            rbcMarketClusterMarket(1).Enabled = True
            mPopSisterStations
        ElseIf Index = 1 Then   'Add to Group
            rbcMarketClusterMarket(0).Value = True  'False
            'rbcMarketClusterMarket(1).Value = False
            rbcMarketClusterMarket(0).Enabled = True    'False
            rbcMarketClusterMarket(1).Enabled = True    'False
            mPopSisterStations
        Else    'Remove from Group
            rbcMarketClusterMarket(0).Value = False
            rbcMarketClusterMarket(1).Value = False
            rbcMarketClusterMarket(0).Enabled = False
            rbcMarketClusterMarket(1).Enabled = False
        End If
        bmIgnoreSisterStationChange = False
        mPopSisterStations
    End If
End Sub

Private Sub rbcMarketClusterMarket_Click(Index As Integer)
    If bmIgnoreSisterStationChange Then
        Exit Sub
    End If
    If rbcMarketClusterMarket(Index).Value Then
        mPopSisterStations
    End If
End Sub

Private Sub rbcMulticastMarket_Add_Click(Index As Integer)
    If bmIgnoreMulticastChange Then
        Exit Sub
    End If
    If rbcMulticastMarket_Add(Index).Value Then
        mPopMulticast
    End If

End Sub

Private Sub rbcMulticastOwner_Add_Click(Index As Integer)
    If bmIgnoreMulticastChange Then
        Exit Sub
    End If
    If rbcMulticastOwner_Add(Index).Value Then
        mPopMulticast
    End If
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    lacOnAir.FontBold = True
    lacCommercial.FontBold = True
    lacDaylight.FontBold = True
End Sub

Private Sub txtCallLetters_LostFocus()
    Dim slCallLetters As String
    On Error GoTo ErrHand
    slCallLetters = Trim$(UCase$(txtCallLetters.Text))
    If slCallLetters = "" Then
        '3/10/11:  After save presses we end up here or Done pressed without entering the call letters
        'Removing message
        'gMsgBox "Call Letters not Defined.", vbOKOnly
        Exit Sub
    End If
    SQLQuery = "SELECT shttCode FROM shtt WHERE (shttCallLetters = '" & slCallLetters & "')"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If imShttCode <> rst!shttCode Then       '(0).Value Then
            If optSP(1).Value Then
                gMsgBox "Name Previously Defined.", vbOKOnly
            Else
                gMsgBox "Call Letters Previously Defined.", vbOKOnly
            End If
        End If
    End If
    udcContactGrid.CALLLETTERS = Trim$(txtCallLetters.Text)
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-txtCallLetters"
End Sub

Private Sub txtFax_LostFocus()
    udcContactGrid.FaxNumber = Trim$(txtFax.Text)
End Sub


Private Sub txtIPumpID_Change()
    imFieldChgd = True
End Sub

Private Sub txtIPumpID_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtONCountry_Change()
    imFieldChgd = True
End Sub

Private Sub txtONCountry_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcP12Plus_Change()
    imFieldChgd = True
End Sub

Private Sub edcP12Plus_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcP12Plus_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYCOMMA) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub edcPermanentStationID_Change()
    imFieldChgd = True
End Sub

Private Sub edcPermanentStationID_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPermanentStationID_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    Dim llRow As Long

    If imFirstTime Then
        bgStationVisible = True
    
        pbcClickFocus.Left = -240
        pbcSTab(0).Left = -240
        pbcSTab(1).Left = -240
        pbcTab(0).Left = -240
        pbcTab(1).Left = -240
        
        'grdEmail.ColWidth(WEBSEQNOINDEX) = 0
        'grdEmail.ColWidth(WEBEMAILINDEX) = grdEmail.Width '* 0.6
        ''grdEmail.ColWidth(LASTDATEINDEX) = grdEmail.Width - grdEmail.ColWidth(WEBEMAILINDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        'gGrid_AlignAllColsLeft grdEmail
        'grdEmail.TextMatrix(0, WEBEMAILINDEX) = "Web Email Address "
        ''Debug show the seq. number while developing only
        'grdEmail.TextMatrix(0, LASTDATEINDEX) = "Seq. Num"
        'gGrid_IntegralHeight grdEmail
        'gGrid_Clear grdEmail, True
        
        
        grdMulticast.ColWidth(MCSELECTEDINDEX) = 0
        grdMulticast.ColWidth(MCSHTTCODEINDEX) = 0
        grdMulticast.ColWidth(MCDMAMKTCODEINDEX) = 0
        grdMulticast.ColWidth(MCSORTINDEX) = 0
        grdMulticast.ColWidth(MCCALLLETTERSINDEX) = grdMulticast.Width * 0.15
        grdMulticast.ColWidth(MCMARKETINDEX) = grdMulticast.Width * 0.25
        grdMulticast.ColWidth(MCLICCITYINDEX) = grdMulticast.Width * 0.25
        grdMulticast.ColWidth(MCMAILSTATEINDEX) = grdMulticast.Width * 0.2
        grdMulticast.ColWidth(MCOWNERINDEX) = grdMulticast.Width - grdMulticast.ColWidth(MCCALLLETTERSINDEX) - grdMulticast.ColWidth(MCMARKETINDEX) - grdMulticast.ColWidth(MCLICCITYINDEX) - grdMulticast.ColWidth(MCMAILSTATEINDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        gGrid_AlignAllColsLeft grdMulticast
        grdMulticast.TextMatrix(0, MCCALLLETTERSINDEX) = "Call Letters"
        grdMulticast.TextMatrix(0, MCMARKETINDEX) = "DMA Market"
        grdMulticast.TextMatrix(0, MCLICCITYINDEX) = "License City"
        grdMulticast.TextMatrix(0, MCMAILSTATEINDEX) = "Mailing State"
        grdMulticast.TextMatrix(0, MCOWNERINDEX) = "Owner"
        grdMulticast.Row = 0
        grdMulticast.Col = MCCALLLETTERSINDEX
        grdMulticast.CellBackColor = LIGHTBLUE
        grdMulticast.Col = MCMARKETINDEX
        grdMulticast.CellBackColor = LIGHTBLUE
        grdMulticast.Col = MCLICCITYINDEX
        grdMulticast.CellBackColor = LIGHTBLUE
        grdMulticast.Col = MCMAILSTATEINDEX
        grdMulticast.CellBackColor = LIGHTBLUE
        grdMulticast.Col = MCOWNERINDEX
        grdMulticast.CellBackColor = LIGHTBLUE
        gGrid_IntegralHeight grdMulticast
        grdMulticast.Height = grdMulticast.Height + 30  'To avoid scrolling when click on bottom row
        
        
        grdSisterStations.ColWidth(SSSELECTEDINDEX) = 0
        grdSisterStations.ColWidth(SSSHTTCODEINDEX) = 0
        grdSisterStations.ColWidth(SSSORTINDEX) = 0
        grdSisterStations.ColWidth(SSCALLLETTERSINDEX) = grdSisterStations.Width * 0.2
        grdSisterStations.ColWidth(SSLICCITYINDEX) = grdSisterStations.Width * 0.3
        grdSisterStations.ColWidth(SSMAILSTATEINDEX) = grdSisterStations.Width * 0.2
        grdSisterStations.ColWidth(SSMARKETINDEX) = grdSisterStations.Width - grdSisterStations.ColWidth(SSCALLLETTERSINDEX) - grdSisterStations.ColWidth(SSLICCITYINDEX) - grdSisterStations.ColWidth(SSMAILSTATEINDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        gGrid_AlignAllColsLeft grdSisterStations
        grdSisterStations.TextMatrix(0, SSCALLLETTERSINDEX) = "Call Letters"
        grdSisterStations.TextMatrix(0, SSMARKETINDEX) = "DMA Market"
        grdSisterStations.TextMatrix(0, SSLICCITYINDEX) = "License City"
        grdSisterStations.TextMatrix(0, SSMAILSTATEINDEX) = "Mailing State"
        grdSisterStations.Row = 0
        grdSisterStations.Col = SSCALLLETTERSINDEX
        grdSisterStations.CellBackColor = LIGHTBLUE
        grdSisterStations.Col = SSMARKETINDEX
        grdSisterStations.CellBackColor = LIGHTBLUE
        grdSisterStations.Col = SSLICCITYINDEX
        grdSisterStations.CellBackColor = LIGHTBLUE
        grdSisterStations.Col = SSMAILSTATEINDEX
        grdSisterStations.CellBackColor = LIGHTBLUE
        gGrid_IntegralHeight grdSisterStations
        grdSisterStations.Height = grdSisterStations.Height + 30  'To avoid scrolling when click on bottom row
        
        'Hide column 2
        grdHistory.ColWidth(SHCLTCODEINDEX) = 0
        grdHistory.ColWidth(SHCALLLETTERSINDEX) = grdHistory.Width * 0.6
        grdHistory.ColWidth(SHLASTDATEINDEX) = grdHistory.Width - grdHistory.ColWidth(SHCALLLETTERSINDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        gGrid_AlignAllColsLeft grdHistory
        grdHistory.TextMatrix(0, SHCALLLETTERSINDEX) = "Previous Call Letters"
        grdHistory.TextMatrix(0, SHLASTDATEINDEX) = "Last Active Date"
        gGrid_IntegralHeight grdHistory
        gGrid_Clear grdHistory, True
        pbcHistSTab.Left = -240
        pbcHistTab.Left = -240
        'labSort.Top = optSort(0).Top

        'grdPersonnel.ColWidth(NAMEINDEX) = grdPersonnel.Width * 0.2
        'grdPersonnel.ColWidth(PHONEINDEX) = grdPersonnel.Width * 0.11
        'grdPersonnel.ColWidth(FAXINDEX) = grdPersonnel.Width * 0.11
        'grdPersonnel.ColWidth(EMAILINDEX) = grdPersonnel.Width * 0.24
        'grdPersonnel.ColWidth(TITLEINDEX) = grdPersonnel.Width * 0.13
        'grdPersonnel.ColWidth(AFCNTINDEX) = grdPersonnel.Width * 0.09
        'grdPersonnel.ColWidth(ISCI2INDEX) = grdPersonnel.Width * 0.09
        'grdPersonnel.ColWidth(PRCODEINDEX) = 0
        'grdPersonnel.ColWidth(TNTCODEINDEX) = 0
        'grdPersonnel.ColWidth(STATUSINDEX) = 0

        'gGrid_AlignAllColsLeft grdPersonnel
        'grdPersonnel.TextMatrix(0, NAMEINDEX) = "Name"
        'grdPersonnel.TextMatrix(0, PHONEINDEX) = "Phone #"
        'grdPersonnel.TextMatrix(0, FAXINDEX) = "Fax #"
        'grdPersonnel.TextMatrix(0, EMAILINDEX) = "E-Mail"
        'grdPersonnel.TextMatrix(0, TITLEINDEX) = "Title"
        'grdPersonnel.TextMatrix(0, AFCNTINDEX) = "Aff/E-Mail"
        'grdPersonnel.TextMatrix(1, AFCNTINDEX) = "Contact"
        'grdPersonnel.TextMatrix(0, ISCI2INDEX) = "ISCI Export"
        'grdPersonnel.TextMatrix(1, ISCI2INDEX) = "Contact"
        'gGrid_IntegralHeight grdPersonnel
        'gGrid_Clear grdPersonnel, True

        'Setting the location of the grid must be before mSort or we can't debug the user
        'control during the mSort call to setup the grid
        udcContactGrid.Action 2 'Init
        udcContactGrid.Source = "S"
        udcContactGrid.Move pbcArrow.Width, 0, frcTab(3).Width - pbcArrow.Width, frcTab(3).Height
        
        mSort
        
        
        imFirstTime = False
        imFieldChgd = False
        If sgStationCallSource = "S" Then
            llRow = SendMessageByString(cboStations.hwnd, CB_FINDSTRING, -1, sgTCCallLetters)
            If llRow >= 0 Then
                cboStations.ListIndex = llRow
            End If
        End If
        cboStations.SetFocus
    ElseIf (sgStationCallSource = "S") And (imFieldChgd = False) And (imInSave <> True) Then
        llRow = SendMessageByString(cboStations.hwnd, CB_FINDSTRING, -1, sgTCCallLetters)
        If llRow >= 0 Then
            cboStations.ListIndex = llRow
        End If
        cboStations.SetFocus
    End If
    sgStationCallSource = ""
End Sub

Private Sub Form_Initialize()
    
    Me.Visible = False
    
    Me.Width = (Screen.Width) / 1.1
    Me.Height = (Screen.Height) / 1.2   '1.3
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmStation
    gCenterForm frmStation
    bmCancelled = False
    bmIgnoreMulticastChange = False
    bmIgnoreSisterStationChange = False
    cbcONCity.ReSizeFont = "A"
    cbcONCity.Move frcPhysicalAddress.Left + txtONAddr1.Left, frcPhysicalAddress.Top + lacONCity.Top - 15, cbcONCity.Width, cboONState.Height
    cbcONCity.SetDropDownWidth cbcONCity.Width
    cbcCity.ReSizeFont = "A"
    cbcCity.Move frcMailingAddress.Left + txtAddr1.Left, frcMailingAddress.Top + lacCity.Top - 15, cbcCity.Width, cboState.Height
    cbcCity.SetDropDownWidth cbcCity.Width
    cbcCityLic.ReSizeFont = "A"
    cbcCityLic.Move txtStaPhone.Left, frcLicense.Top + lacCityLic.Top - 45, cbcCityLic.Width, cbcONCity.Height
    cbcCityLic.SetDropDownWidth cbcCityLic.Width
    cbcCountyLic.ReSizeFont = "A"
    cbcCountyLic.Move txtFax.Left, frcLicense.Top + lacCountyLic.Top - 45, cbcCountyLic.Width, cbcONCity.Height
    cbcCountyLic.SetDropDownWidth cbcCountyLic.Width
    cbcMoniker.ReSizeFont = "A"
    cbcMoniker.Height = cbcONCity.Height
    cbcMoniker.SetDropDownWidth cbcMoniker.Width
    cbcOwner.ReSizeFont = "A"
    cbcOwner.Height = cbcONCity.Height
    cbcOwner.SetDropDownWidth cbcOwner.Width
    cbcOwner.PopUpListDirection "A"
    cbcDMAMarket.ReSizeFont = "A"
    cbcDMAMarket.Height = cbcONCity.Height
    cbcDMAMarket.SetDropDownWidth cbcDMAMarket.Width
    cbcDMAMarket.PopUpListDirection "A"
    cbcMSAMarket.ReSizeFont = "A"
    cbcMSAMarket.Height = cbcONCity.Height
    cbcMSAMarket.SetDropDownWidth cbcMSAMarket.Width
    cbcMSAMarket.PopUpListDirection "A"
    cbcOperator.ReSizeFont = "A"
    cbcOperator.Height = cbcONCity.Height
    cbcOperator.SetDropDownWidth cbcOperator.Width
    cbcOperator.PopUpListDirection "A"
    cbcTerritory.ReSizeFont = "A"
    cbcTerritory.Height = cbcONCity.Height
    cbcTerritory.SetDropDownWidth cbcTerritory.Width
    cbcTerritory.PopUpListDirection "A"
    cbcFormat.ReSizeFont = "A"
    cbcFormat.Height = cbcONCity.Height
    cbcFormat.SetDropDownWidth cbcFormat.Width
    cbcArea.ReSizeFont = "A"
    cbcArea.Height = cbcONCity.Height
    cbcArea.SetDropDownWidth cbcArea.Width
    cbcArea.PopUpListDirection "A"
    cbcMarketRep.ReSizeFont = "A"
    cbcMarketRep.Height = cbcONCity.Height
    cbcMarketRep.SetDropDownWidth cbcMarketRep.Width
    cbcMarketRep.PopUpListDirection "A"
    cbcServiceRep.ReSizeFont = "A"
    cbcServiceRep.Height = cbcONCity.Height
    cbcServiceRep.SetDropDownWidth cbcServiceRep.Width
    cbcServiceRep.PopUpListDirection "A"
    cbcTimeZone.ReSizeFont = "A"
    cbcTimeZone.Height = cbcONCity.Height
    cbcTimeZone.SetDropDownWidth cbcTimeZone.Width
    cbcTimeZone.PopUpListDirection "A"
    edcHistStartDate.Height = cbcONCity.Height
    bmIgnoreTitleListChanges = False
    bmIgnoreOwnerChange = False
    lmTabColor(0) = -1
    mSetTabColors
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iIndex As Integer
    Dim ilValue10 As Integer
    
    On Error GoTo ErrHand

    bFormWasAlreadyResized = False
    imIsMulticastDirty = False
    smExistingWebPW = ""
    smExistingWebEmail = ""
    '7912
    smExistingXDSSiteID = ""
    '10192
    imExistingHonorDaylight = -1
    imWebEmailUpdated = False
    imWebPWUpdated = False
    imLastMCColSorted = -1
    imLastMCSort = -1
    imLastSSColSorted = -1
    imLastSSSort = -1
    frmStation.Caption = "Station Information - " & sgClientName
    Screen.MousePointer = vbHourglass
    'Me.Width = (Screen.Width) / 1.1
    'Me.Height = (Screen.Height) / 1.15
    'Me.Top = 850    '(Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    'Fill cboSations
    'mSort
    
    imFirstTime = True
    
    imShowGridBox = False
    imFromArrow = False
    imPersFromArrow = False
    imEmailFromArrow = False
    lmTopRow = -1
    lmEnableRow = -1
    bmAdjPledge = False
    
    imcTrash.Picture = frmDirectory!imcTrashClosed.Picture
    imTabIndex = 1
    imInSave = False
    'Fill cboAffRep1
    '5/10/07:  Removed Affiliate Rep from Station File
    'cboAffRep1.AddItem "[New]", 0
    'cboAffRep1.ItemData(0) = -1
    
    'cboAffRep1.AddItem "[No Rep]", 1
    'cboAffRep1.ItemData(1) = 0
    
    
    'cboDMAMarketCluster.AddItem "[New]", 0
    'cboDMAMarketCluster.ItemData(cboDMAMarketCluster.NewIndex) = -1
    'cboDMAMarketCluster.AddItem "[None]", 1
    'cboDMAMarketCluster.ItemData(cboDMAMarketCluster.NewIndex) = 0
    
    'cboMSAMarketCluster.AddItem "[New]", 0
    'cboMSAMarketCluster.ItemData(cboMSAMarketCluster.NewIndex) = -1
    'cboMSAMarketCluster.AddItem "[None]", 1
    'cboMSAMarketCluster.ItemData(cboMSAMarketCluster.NewIndex) = 0

    '9/17/11: Moved to mPop
    'Call mOwnerFillListBox(0, False, -1)
    'Call mDMAMarketFillListBox(0, False, -1)
    'Call mMSAMarketFillListBox(0, False, -1)

    'mPopTerritory
    'mPopFormat
    'mPopState
    'mPopTimeZone
    'mPopCity
    'mPopCounty
    'mPopMoniker
    'mPopOperator
    'mPopArea
    'mPopMarketRep
    'mPopServiceRep
    
    If gIsUsingNovelty Then
        edcWebNumber.Visible = False
        lacWebNumber.Visible = False
        
        cmdGenPassword.Visible = False
        txtWebPW.Visible = False
        lblWebPW.Visible = False
    End If
    lacWebSpotsPerPage.Visible = False
    txtWebSpotsPerPage.Visible = False
    
    mPopState
    mPopTimeZone
    
    bmDoPop = True
    mPop
    
    '5/10/07:  Removed Affiliate Rep from Station File
    'SQLQuery = "SELECT arttFirstName, arttLastName, arttCode FROM artt Where arttType = 'R' ORDER BY arttLastName"
    'Set rst = gSQLSelectCall(SQLQuery)
    'While Not rst.EOF
    '    cboAffRep1.AddItem Trim$(rst!arttFirstName) & " " & Trim$(rst!arttLastName)
    '    cboAffRep1.ItemData(cboAffRep1.NewIndex) = rst!arttCode
    '    rst.MoveNext
    'Wend
    If sgUstWin(1) <> "I" Then
    '    cboAffRep1.Enabled = False
    '    cmdAffRep.Enabled = False
        cmdSave.Enabled = False
        cmdDone.Enabled = False
        cmdErase.Enabled = False
        frcTab(0).Enabled = False
        frcTab(1).Enabled = False
        frcTab(2).Enabled = False
        imcTrash.Enabled = False
        'Leave frcTab(3) enabled so that user can scroll
    End If
    If sgUsingStationID <> "Y" Then
        lacPermanentStationID.Visible = False
        edcPermanentStationID.Visible = False
    End If
    smWegenerIPump = "N"
    SQLQuery = "Select spfUsingFeatures10 From SPF_Site_Options"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        ilValue10 = Asc(rst!spfUsingFeatures10)
        If (ilValue10 And WEGENERIPUMP) = WEGENERIPUMP Then
            smWegenerIPump = "Y"
        End If
    End If
    If smWegenerIPump = "N" Then
        lacIPumpID.Enabled = False
        txtIPumpID.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-Form Load"
End Sub

Private Sub Form_Resize()
    Dim Ctrl As control
    If bFormWasAlreadyResized Then
        For Each Ctrl In frmStation.Controls
            If TypeOf Ctrl Is TextBox Then
                Ctrl.Height = cboState.Height
            End If
        Next Ctrl
        Exit Sub
    End If
    bFormWasAlreadyResized = True
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    tscStation.Left = frcSelect.Left
    tscStation.Top = frcSelect.Top + frcSelect.Height + 120
    tscStation.Height = cmdCancel.Top - (frcSelect.Top + frcSelect.Height + 240)  'tscStation.ClientTop - tscStation.Top + (10 * frcTab(1).Height) / 9
    tscStation.Width = frcSelect.Width
    frcTab(0).Move tscStation.ClientLeft, tscStation.ClientTop - 45, tscStation.ClientWidth, tscStation.ClientHeight + 90
    frcTab(1).Move tscStation.ClientLeft, tscStation.ClientTop, tscStation.ClientWidth, tscStation.ClientHeight
    frcTab(2).Move tscStation.ClientLeft, tscStation.ClientTop, tscStation.ClientWidth, tscStation.ClientHeight
    frcTab(3).Move tscStation.ClientLeft, tscStation.ClientTop, tscStation.ClientWidth, tscStation.ClientHeight
    frcTab(4).Move tscStation.ClientLeft, tscStation.ClientTop, tscStation.ClientWidth, tscStation.ClientHeight
    frcTab(5).Move tscStation.ClientLeft, tscStation.ClientTop, tscStation.ClientWidth, tscStation.ClientHeight
    'frcTab(6).Move tscStation.ClientLeft, tscStation.ClientTop, tscStation.ClientWidth, tscStation.ClientHeight
    frcTab(0).BorderStyle = 0
    frcTab(1).BorderStyle = 0
    frcTab(2).BorderStyle = 0
    frcTab(3).BorderStyle = 0
    frcTab(4).BorderStyle = 0
    frcTab(5).BorderStyle = 0
    'frcTab(6).BorderStyle = 0
    For Each Ctrl In frmStation.Controls
        If TypeOf Ctrl Is TextBox Then
            Ctrl.Height = cboState.Height
        End If
    Next Ctrl
    tmcStart.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    bgStationVisible = False
    Erase tmHistoryInfo
    Erase lmAttCode
    Erase lmAttCodesToUpdateWeb
    Erase tmBaseStaInfo
    Erase tmGroup1Sort
    Erase tmGroup2Sort
    Erase tmGroup3Sort
    Erase tmGroup4Sort
    Erase tmEmailInfo
    attrst.Close
    DATRST.Close
    If sgStationCallSource = "S" Then
        frmStationSearch.SetFocus
    End If
    Set frmStation = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub frcTab_Click(Index As Integer)
    udcContactGrid.Action 1
End Sub

Private Sub grdHistory_Click()
    Dim llRow As Long
    
    If sgUstWin(1) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdHistory.Col >= grdHistory.Cols - 1 Then
        Exit Sub
    End If
'    lmTopRow = grdHistory.TopRow
'    llRow = grdHistory.Row
'    If grdHistory.TextMatrix(llRow, 0) = "" Then
'        grdHistory.Redraw = False
'        Do
'            llRow = llRow - 1
'        Loop While grdHistory.TextMatrix(llRow, 0) = ""
'        grdHistory.Row = llRow + 1
'        grdHistory.Col = 0
'        grdHistory.Redraw = True
'    End If
'    mHistEnableBox
End Sub

Private Sub grdHistory_GotFocus()
    If grdHistory.Col >= grdHistory.Cols - 1 Then
        Exit Sub
    End If
    'grdHistory_Click
End Sub

Private Sub grdHistory_EnterCell()
    mHistSetShow
    If sgUstWin(1) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
End Sub

Private Sub grdHistory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdHistory.TopRow
    grdHistory.Redraw = False
End Sub

Private Sub grdHistory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim ilFound As Integer
    
    If sgUstWin(1) <> "I" Then
        grdHistory.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdHistory, X, Y)
    If Not ilFound Then
        grdHistory.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdHistory.Col >= grdHistory.Cols - 1 Then
        grdHistory.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdHistory.TopRow
    
    llRow = grdHistory.Row
    If grdHistory.TextMatrix(llRow, SHCALLLETTERSINDEX) = "" Then
        grdHistory.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdHistory.TextMatrix(llRow, SHCALLLETTERSINDEX) = ""
        grdHistory.Row = llRow + 1
        grdHistory.Col = SHCALLLETTERSINDEX
        grdHistory.Redraw = True
    End If
    grdHistory.Redraw = True
    mHistEnableBox
End Sub

Private Sub grdHistory_Scroll()
'    If (lmTopRow <> -1) And (lmTopRow <> grdHistory.TopRow) Then
'        grdHistory.TopRow = lmTopRow
'        lmTopRow = -1
'    End If
    If grdHistory.Redraw = False Then
        grdHistory.Redraw = True
        grdHistory.TopRow = lmTopRow
        grdHistory.Refresh
        grdHistory.Redraw = False
    End If
    If (imShowGridBox) And (grdHistory.Row >= grdHistory.FixedRows) And (grdHistory.Col >= 0) And (grdHistory.Col < grdHistory.Cols - 1) Then
        If grdHistory.RowIsVisible(grdHistory.Row) Then
            txtHistory.Move grdHistory.Left + grdHistory.ColPos(grdHistory.Col) + 30, grdHistory.Top + grdHistory.RowPos(grdHistory.Row) + 30, grdHistory.ColWidth(grdHistory.Col) - 30, grdHistory.RowHeight(grdHistory.Row) - 30
            pbcArrow.Move grdHistory.Left - pbcArrow.Width, grdHistory.Top + grdHistory.RowPos(grdHistory.Row) + (grdHistory.RowHeight(grdHistory.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            txtHistory.Visible = True
            txtHistory.SetFocus
        Else
            pbcClickFocus.SetFocus
            txtHistory.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
        imPersFromArrow = False
        imEmailFromArrow = False
    End If
End Sub

Private Sub grdMulticast_Click()

    Dim llRow As Long
    If grdMulticast.MouseRow = 0 Then
        mMCSortCol grdMulticast.MouseCol
    Else
        If (grdMulticast.MouseRow >= grdMulticast.FixedRows) And (grdMulticast.MouseRow < grdMulticast.Rows) Then
            If grdMulticast.TextMatrix(grdMulticast.MouseRow, MCCALLLETTERSINDEX) <> "" Then
                If grdMulticast.TextMatrix(grdMulticast.MouseRow, MCSELECTEDINDEX) <> "1" Then
                    grdMulticast.TextMatrix(grdMulticast.MouseRow, MCSELECTEDINDEX) = "1"
                    'TTP 10898 JJB 2024-04-26
                    'If rbcMulticast(1).Value Then
                    '    For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
                    '        If llRow <> grdMulticast.MouseRow Then
                    '            If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1" Then
                    '                grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "0"
                    '                mMCPaintRowColor llRow
                    '            End If
                    '        End If
                    '    Next llRow
                    'End If
                Else
                    grdMulticast.TextMatrix(grdMulticast.MouseRow, MCSELECTEDINDEX) = "0"
                End If
                imFieldChgd = True
                mMCPaintRowColor grdMulticast.MouseRow
            End If
        End If
    End If
End Sub

Private Sub imcTrash_Click()
    Dim iLoop As Integer
    Dim iRow As Integer
    Dim iRows As Integer
    
    If imInSave = True Then
        Exit Sub
    End If
    mHistSetShow
    iRow = grdHistory.Row
    iRows = grdHistory.Rows
    If (iRow < 0) Or (iRow > grdHistory.Rows - 1) Then
        Exit Sub
    End If
    If grdHistory.TextMatrix(iRow, SHCLTCODEINDEX) <> "" Then
        imFieldChgd = True
    End If
    grdHistory.RemoveItem iRow
    gGrid_FillWithRows grdHistory

End Sub



Private Sub lacCommercial_Click()
    If lacCommercial.Caption = "Commercial" Then
        lacCommercial.Caption = "Non-Commercial"
        lacCommercial.BackColor = vbRed
    Else
        lacCommercial.Caption = "Commercial"
        lacCommercial.BackColor = GREEN
    End If
    imFieldChgd = True
End Sub

Private Sub lacDaylight_Click()
    If lacDaylight.Caption = "Honor Daylight Savings" Then
        lacDaylight.Caption = "Ignore Daylight Savings"
        lacDaylight.BackColor = vbRed
    Else
        lacDaylight.Caption = "Honor Daylight Savings"
        lacDaylight.BackColor = GREEN
    End If
    imFieldChgd = True
End Sub

Private Sub lacOnAir_Click()
    If lacOnAir.Caption = "On Air" Then
        lacOnAir.Caption = "Off Air"
        lacOnAir.BackColor = vbRed
    Else
        lacOnAir.Caption = "On Air"
        lacOnAir.BackColor = GREEN
    End If
    imFieldChgd = True
End Sub

Private Sub optSort_Click(Index As Integer)
    Dim iLoop As Integer
    Dim iIndex As Integer
    
    If optSort(Index).Value = False Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    mSort
    Screen.MousePointer = vbDefault
End Sub

Private Function mSave(iAsk As Integer) As Integer
    Dim iType As Integer
    Dim iChecked As Integer
    Dim iSelected As Integer
    Dim iIndex As Integer
    Dim iUpper As Integer
    Dim i As Integer
    Dim iPos As Integer
    Dim sCallLetters As String
    Dim sFrequency As String
    Dim lPermanentStationID As Long
    Dim lXDSStationID As Long
    Dim sAddr1 As String
    Dim sAddr2 As String
    Dim sCity As String
    Dim sState As String
    Dim sCountry As String
    Dim slOnCountry As String
    Dim sPDName As String
    Dim sTDName As String
    Dim sMDName As String
    Dim sACName As String
    Dim sACPhone As String
    Dim sMarket As String
    Dim sEmail As String
    Dim sONAddr1 As String
    Dim sONAddr2 As String
    Dim sONCity As String
    Dim sONState As String
    Dim sCityLic As String
    Dim sStateLic As String
    Dim iLoop As Integer
    Dim CurDate As String
    Dim CurTime As String
    Dim iDaylight As Integer
    Dim slOnAir As String
    Dim slStationType As String
    Dim llAudP12Plus As Long
    Dim llWatts As Long
    Dim slHistStartDate As String
    Dim iRetCode As Integer
    Dim slTimeZone As String
    Dim ilTztCode As Integer
    Dim slSerialNo1 As String
    Dim slSerialNo2 As String
    Dim slPort As String
    Dim iVehZonesDefined As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim iRet As Integer
    Dim slWebPage As String
    Dim llMntCode As Long
    Dim ilFmtCode As Integer
    Dim slUsedForAtt As String
    Dim slUsedForXDigital As String
    Dim slUsedForWegener As String
    Dim slUsedForOLA As String
    Dim slUsedForPledgeVsAir As String
    Dim slWebAddress As String
    Dim llOnCityMntCode As Long
    Dim llCityMntCode As Long
    Dim llCityLicMntCode As Long
    Dim llCountyLicMntCode As Long
    Dim llMonikerMntCode As Long
    Dim llOperatorMntCode As Long
    Dim llAreaMntCode As Long
    Dim ilMktRepUstCode As Integer
    Dim ilServRepUstCode As Integer
    Dim slEnterpriseID As String
    Dim llMulticastGroupID As Long
    Dim llMarketClusterGroupID As Long
    Dim shtt_rst As ADODB.Recordset
    Dim ilRet As Integer
    Dim blForcePop As Boolean
    'ttp 5352
    Dim slInvalidEmail As String
    Dim slIPumpID As String
    Dim ilWebSpotsPerPage As Integer
    Dim slOldMaster As String
    Dim ilTotalUpdated As Integer
    Dim slWebNumber As Integer
    Dim blRet As Boolean
    '8418
    Dim ilOldWeb As Integer
    Dim ilNewWeb As Integer
    '8824
    Dim llRow As Long
    Dim ilShttForVatUpdating As Integer
    Dim slShtt As String
    
    On Error GoTo ErrHand
    
    mSave = False
        
    If sgUstWin(1) <> "I" Then
        gMsgBox "Not Allowed to Save.", vbOKOnly
        Exit Function
    End If
    If optSP(1).Value Then
        sCallLetters = Trim$(txtCallLetters.Text)
    Else
        sCallLetters = Trim$(UCase$(txtCallLetters.Text))
    End If
    If sCallLetters = "" Then
        'If Not iAsk Then    '"Not iAsk" is Save button
            gMsgBox "Call Letters must be Defined.", vbOKOnly
        'End If
        Exit Function
    End If
    
    SQLQuery = "SELECT shttCode FROM shtt WHERE (UCase(shttCallLetters) = '" & UCase(sCallLetters) & "')"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If imShttCode <> rst!shttCode Then       '(0).Value Then
            Screen.MousePointer = vbDefault
            If optSP(1).Value Then
                gMsgBox "Name Previously Defined.", vbOKOnly
            Else
                gMsgBox "Call Letters Previously Defined.", vbOKOnly
            End If
            Exit Function
        End If
    End If
    
    iVehZonesDefined = False
    For iVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
        For iZone = LBound(tgVehicleInfo(iVef).sZone) To UBound(tgVehicleInfo(iVef).sZone) Step 1
            If (Trim$(tgVehicleInfo(iVef).sZone(iZone)) <> "") And (Trim$(tgVehicleInfo(iVef).sZone(iZone)) <> "~~~") Then
                iVehZonesDefined = True
                Exit For
            End If
        Next iZone
        If iVehZonesDefined = True Then
            Exit For
        End If
    Next iVef
    If imTimeZoneIndex >= 0 Then
        slTimeZone = tgTimeZoneInfo(cbcTimeZone.GetItemData(imTimeZoneIndex)).sCSIName
        ilTztCode = tgTimeZoneInfo(cbcTimeZone.GetItemData(imTimeZoneIndex)).iCode
    Else
        slTimeZone = ""
        ilTztCode = 0
    End If
    'slTimeZone = Trim$(txtTimeZone.text)
    'If (slTimeZone <> "CST") And (slTimeZone <> "PST") And (slTimeZone <> "MST") And (slTimeZone <> "EST") And (iVehZonesDefined) Then
    If (slTimeZone <> "CST") And (slTimeZone <> "PST") And (slTimeZone <> "MST") And (slTimeZone <> "EST") And (slTimeZone <> "AST") And (slTimeZone <> "HST") And (iVehZonesDefined) Then
        'If Not iAsk Then    '"Not iAsk" is Save button
            gMsgBox "Please select a time zone", vbOKOnly
        'End If
        Exit Function
    End If
    
    imDMAMktCode = 0
    If cbcDMAMarket.ListIndex > 1 Then
        imDMAMktCode = cbcDMAMarket.GetItemData(cbcDMAMarket.ListIndex)
    End If
    imMSAMktCode = 0
    If cbcMSAMarket.ListIndex > 1 Then
        imMSAMktCode = cbcMSAMarket.GetItemData(cbcMSAMarket.ListIndex)
    End If
    If optSP(0).Value Then
        If (imDMAMktCode < 1) And (Trim$(cbcDMAMarket.GetName(cbcDMAMarket.ListIndex)) <> "[None]") Then
            gMsgBox "DMA Market must be Defined.", vbOKOnly
            Exit Function
        End If
        If (imMSAMktCode < 1) And (Trim$(cbcMSAMarket.GetName(cbcMSAMarket.ListIndex)) <> "[None]") Then
            '3/10/11: Remove test completely
            ''gMsgBox "MSA Market must be Defined.", vbOKOnly
            ''Exit Function
            ''2/23/10:  Remove mandatory option from MSA Market if not using MSA split copy
            'If gUsingMSARegions Then
            '    gMsgBox "MSA Market must be Defined.", vbOKOnly
            '    Exit Function
            'Else
            '    imMSAMktCode = 0
            'End If
        End If
        
    End If
    'Check Station ID
    If (sgUsingStationID = "Y") And (optSP(0).Value) Then
        lPermanentStationID = Val(edcPermanentStationID.Text)
        If lPermanentStationID <= 0 Then
            gMsgBox "Station ID must be Defined (Top Right in Main Tab)", vbOKOnly
            Exit Function
        End If
        SQLQuery = "SELECT shttCode, shttCallLetters FROM shtt WHERE shttPermStationID = " & lPermanentStationID
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            If imShttCode <> rst!shttCode Then       '(0).Value Then
                Screen.MousePointer = vbDefault
                gMsgBox "Station ID Previously Defined with " & rst!shttCallLetters, vbOKOnly
                Exit Function
            End If
        End If
    End If
    If (optSP(0).Value) And (StrComp(sCallLetters, smCurCallLetters, 1) <> 0) And (IsStatDirty) And (Trim$(smCurCallLetters) <> "") Then
        'Ask user if this is a chamge in the station call leeter that they want to retain history on
        'and if so, last date the call letters used.
        sgOrigCallLetters = smCurCallLetters
        sgNewCallLetters = sCallLetters
        '10/23/18 D.S. added if statement below to keep the web in sync with the Affiliate
        If sgOrigCallLetters <> sgNewCallLetters Then
            ilTotalUpdated = gExecWebSQLWithRowsEffected("Update WebEMT Set CallLetters = " & "'" & sgNewCallLetters & "'" & " Where CallLetters = " & "'" & sgOrigCallLetters & "'")
            ilTotalUpdated = gExecWebSQLWithRowsEffected("Update Header Set StationName = " & "'" & sgNewCallLetters & "'" & " Where StationName = " & "'" & sgOrigCallLetters & "'")
        End If
        frmHistory.Show vbModal
        If igHistoryStatus = 2 Then
            Exit Function
        End If
    Else
        igHistoryStatus = 0
    End If
    'D.S. 09/14/16 TTP #8171
    If Trim$((sgOrigCallLetters) <> "" And Trim$(sgNewCallLetters) <> "" And (Trim$(sgOrigCallLetters) <> Trim$(sgNewCallLetters))) Then
        If gGetEMailDistribution() Then
            ilRet = gChangeStationName(sgOrigCallLetters, sgNewCallLetters)
        End If
    End If
    If (optSP(0).Value) And (StrComp(slTimeZone, smCurTimeZone, 1) <> 0) And (IsStatDirty) Then
        'Ask user if this is a chamge in the station call leeter that they want to retain history on
        'and if so, last date the call letters used.
        If iVehZonesDefined Then
            If Not mAsk Then
                igTimeZoneStatus = 0
            Else
                sgOrigTimeZone = smCurTimeZone
                sgNewTimeZone = slTimeZone
                'frmStationZone.Show vbModal
                If igTimeZoneStatus = 2 Then
                    Exit Function
                End If
            End If
        Else
            igTimeZoneStatus = 0
        End If
    Else
        igTimeZoneStatus = 0
    End If
     '8350
    If imTabIndex = 3 Then
         udcContactGrid.Action 1
    End If
    '7/28/15: Check Rights
    If Not udcContactGrid.VerifyRights("S") Then
        Exit Function
    End If
    'ttp 5352
    slInvalidEmail = udcContactGrid.InValidEmails()
    If Len(slInvalidEmail) > 0 Then
        If MsgBox("Do you wish to continue to save(OK), or cancel?  The following email(s) are invalid: " & slInvalidEmail, vbOKCancel + vbInformation, "Invalid Email") = vbCancel Then
            Exit Function
        End If
    End If
    Screen.MousePointer = vbHourglass
    DoEvents
    '8418  tests whenever decrementing old to new
    If IsNumeric(smOrigWebNumber) Then
        ilOldWeb = smOrigWebNumber
    Else
        ilOldWeb = 1
    End If
    If IsNumeric(edcWebNumber.Text) Then
        ilNewWeb = edcWebNumber.Text
    Else
        ilNewWeb = 1
    End If
    If ilNewWeb < ilOldWeb Then
        'going from 2 to 1? if agreement tied to web vendor with minimum of 2, returns true
        If mAgreementWithVersion(imShttCode, ilOldWeb) Then
            If MsgBox("Do you wish to continue to save(Ok), or cancel?  You are rolling back to a previous web version, but associated agreements are set up to a later version", vbOKCancel + vbInformation, "Rolling Back Web Number") = vbCancel Then
                edcWebNumber.Text = smOrigWebNumber
                Screen.MousePointer = Default
                Exit Function
            End If
        End If
    End If
    bmDoPop = False
    If (sgShttTimeStamp <> gFileDateTime(sgDBPath & "Shtt.mkd")) Then
        bmDoPop = True
    End If
    
    If optSP(1).Value Then
        sFrequency = ""
        lPermanentStationID = 0
        lXDSStationID = 0
    Else
        sFrequency = edcFrequency.Text
        lPermanentStationID = Val(edcPermanentStationID.Text)
        lXDSStationID = Val(txtXDSStationID.Text)
    End If
    'If IsDirty = False Then
    '    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
    '        If StrComp(sCallLetters, Trim$(tgStationInfo(iLoop).sCallLetters), 1) = 0 Then
    '            gMsgBox "Call Letters Previously Defined.", vbOKOnly
    '            Exit Function
    '        End If
    '    Next iLoop
    'Else
    '    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
    '        If StrComp(sCallLetters, Trim$(tgStationInfo(iLoop).sCallLetters), 1) = 0 Then
    '            If imShttCode <> tgStationInfo(iLoop).iCode Then
    '                gMsgBox "Call Letters Previously Defined.", vbOKOnly
    '                Exit Function
    '            End If
    '        End If
    '    Next iLoop
    'End If
    mSave = True
    CurDate = Format(gNow(), sgShowDateForm)
    CurTime = Format(gNow(), sgShowTimeWSecForm)
        
    If optSP(1).Value = True Then
        iType = 1
    Else
        iType = 0
    End If
    
    If txtMarkRank.Text = "" Then
        txtMarkRank.Text = 0
    End If
    
    'Set arttCode from cboAffRep
    'iARIndex = CInt(cboAffRep1.ListIndex) - 1
    'D.S. 10/09/02 If user chose add new Rep and did not enter a name and clicked ok then
    'save an error was generated
    '5/10/07:  Removed Affiliate Rep from Station File
    'If cboAffRep1.ListIndex < 0 Then
    '    iARIndex = 0
    'Else
    '    iARIndex = cboAffRep1.ItemData(cboAffRep1.ListIndex)
    'End If
        
    'If iARIndex < 0 Then
    '    iARIndex = 0
    'End If
    llMonikerMntCode = 0
    If cbcMoniker.ListIndex > 1 Then
        llMonikerMntCode = cbcMoniker.GetItemData(cbcMoniker.ListIndex)
    End If
    llOperatorMntCode = 0
    If cbcOperator.ListIndex > 1 Then
        llOperatorMntCode = cbcOperator.GetItemData(cbcOperator.ListIndex)
    End If
    llMntCode = 0
    If cbcTerritory.ListIndex > 1 Then
        llMntCode = cbcTerritory.GetItemData(cbcTerritory.ListIndex)
    End If
    ilFmtCode = 0
    If cbcFormat.ListIndex > 1 Then
        ilFmtCode = cbcFormat.GetItemData(cbcFormat.ListIndex)
    End If
    llAreaMntCode = 0
    If cbcArea.ListIndex > 1 Then
        llAreaMntCode = cbcArea.GetItemData(cbcArea.ListIndex)
    End If
    ilMktRepUstCode = 0
    If cbcMarketRep.ListIndex >= 1 Then
        ilMktRepUstCode = cbcMarketRep.GetItemData(cbcMarketRep.ListIndex)
    End If
    If ilMktRepUstCode > 0 Then
        bgMarketRepDefinedByStation = True
    End If
    ilServRepUstCode = 0
    If cbcServiceRep.ListIndex >= 1 Then
        ilServRepUstCode = cbcServiceRep.GetItemData(cbcServiceRep.ListIndex)
    End If
    If ilServRepUstCode > 0 Then
        bgServiceRepDefinedByStation = True
    End If
    
    slWebPage = Trim$(txtWebPage.Text)
    slWebPage = gFixQuote(txtWebPage.Text)
    slWebAddress = Trim$(txtWebAddress.Text)
    slWebAddress = gFixQuote(txtWebAddress.Text)
    'slWebNumber = Trim$(edcWebNumber.Text)
    ilRet = mWebVerChgMultiCastSync()
    smNewWebNumber = Trim$(edcWebNumber.Text)
'    If rbcWebSiteVersion(0).Value = True Or (rbcWebSiteVersion(0).Value = False And rbcWebSiteVersion(1).Value = False) Then   '1 = old web site, 2 = new web site
'        smWebNumber = 1
'    Else
'        smWebNumber = 2
'    End If
    sAddr1 = Trim$(txtAddr1.Text)
    sAddr1 = gFixQuote(sAddr1)
    sAddr2 = Trim$(txtAddr2.Text)
    sAddr2 = gFixQuote(sAddr2)
    'sCity = Trim$(txtCity.Text)
    'sCity = gFixQuote(sCity)
    sCity = ""
    llCityMntCode = 0
    If cbcCity.ListIndex > 1 Then
        llCityMntCode = cbcCity.GetItemData(cbcCity.ListIndex)
        sCity = gFixQuote(cbcCity.Text)
    End If
    'sState = UCase$(Trim$(txtState.text))
    If imStateIndex >= 0 Then
        sState = tgStateInfo(cboState.ItemData(imStateIndex)).sPostalName
    Else
        sState = ""
    End If
    sState = gFixQuote(sState)
    sCountry = Trim$(edcCountry.Text)
    sCountry = gFixQuote(sCountry)
    slOnCountry = Trim$(txtONCountry.Text)
    slOnCountry = gFixQuote(slOnCountry)
    sONAddr1 = Trim$(txtONAddr1.Text)
    sONAddr1 = gFixQuote(sONAddr1)
    sONAddr2 = Trim$(txtONAddr2.Text)
    sONAddr2 = gFixQuote(sONAddr2)
    'sONCity = Trim$(txtONCity.Text)
    sONCity = ""
    llOnCityMntCode = 0
    If cbcONCity.ListIndex > 1 Then
        llOnCityMntCode = cbcONCity.GetItemData(cbcONCity.ListIndex)
        sONCity = gFixQuote(cbcONCity.Text)
    End If
    slSerialNo1 = Trim$(txtSerialNo1.Text)
    slSerialNo2 = Trim$(txtSerialNo2.Text)
    slPort = Trim$(txtPort.Text)
    sONCity = gFixQuote(sONCity)
    'sONState = UCase$(Trim$(txtONState.text))
    If imOnStateIndex >= 0 Then
        sONState = tgStateInfo(cboONState.ItemData(imOnStateIndex)).sPostalName
    Else
        sONState = ""
    End If
    sONState = gFixQuote(sONState)

    sACName = ""
    sACPhone = ""
    ''sMarket = Trim$(txtMarket.Text)
    'sMarket = Trim$(cboDMAMarketCluster.Text)
    'sMarket = gFixQuote(sMarket)
    If cbcDMAMarket.ListIndex > 1 Then
        sMarket = Trim$(cbcDMAMarket.GetName(cbcDMAMarket.ListIndex))
        sMarket = gFixQuote(sMarket)
    Else
        sMarket = ""
    End If
    
    sEmail = "" 'Trim$(txtEmail.Text)
    sEmail = gFixQuote(sEmail)
    'sCityLic = Trim$(txtCityLic.Text)
    'sCityLic = gFixQuote(sCityLic)
    sCityLic = ""
    llCityLicMntCode = 0
    If cbcCityLic.ListIndex > 1 Then
        llCityLicMntCode = cbcCityLic.GetItemData(cbcCityLic.ListIndex)
        sCityLic = gFixQuote(cbcCityLic.GetName(cbcCityLic.ListIndex))
    End If
    llCountyLicMntCode = 0
    If cbcCountyLic.ListIndex > 1 Then
        llCountyLicMntCode = cbcCountyLic.GetItemData(cbcCountyLic.ListIndex)
    End If
    'sStateLic = Trim$(txtStateLic.text)
    If imStateLicIndex >= 0 Then
        sStateLic = tgStateInfo(cboStateLic.ItemData(imStateLicIndex)).sPostalName
    Else
        sStateLic = ""
    End If
    
    'iPos = InStr(sStateLic, "'")
    'If iPos > 0 Then
    '    sStateLic = Left$(sStateLic, iPos) & "'" & Right$(sStateLic, Len(sStateLic) - iPos)
    'End If
    'If optDaylight(0).Value Then
    '    iDaylight = 0
    'Else
    '    iDaylight = 1
    'End If
    If lacDaylight.Caption = "Ignore Daylight Savings" Then
        iDaylight = 1
    Else
        iDaylight = 0
    End If
    If lacOnAir.Caption = "Off Air" Then
        slOnAir = "N"
    Else
        slOnAir = "Y"
    End If
    If lacCommercial.Caption = "Non-Commercial" Then
        slStationType = "N"
    Else
        slStationType = "C"
    End If
    llAudP12Plus = Val(gRemoveChar(edcP12Plus.Text, ","))
    llWatts = Val(gRemoveChar(edcWatts.Text, ","))
    slEnterpriseID = edcEnterpriseID.Text
    'If edcHistoricalDate.Text = "" Then
    '    slHistStartDate = "1/1/1970"
    'Else
    '    slHistStartDate = edcHistoricalDate.Text
    '    If Not gIsDate(slHistStartDate) Then
    '        slHistStartDate = "1/1/1970"
    '    End If
    'End If
    If edcHistStartDate.Text = "" Then
        slHistStartDate = "1/1/1970"
    Else
        slHistStartDate = edcHistStartDate.Text
        If Not gIsDate(slHistStartDate) Then
            slHistStartDate = "1/1/1970"
        End If
    End If
    smShttWebPW = gFixQuote(Trim$(txtWebPW.Text))
'
'
'    smShttWebEmail = gFixQuote(Trim$(txtWebEmail.text))
'
'    If Len(smShttWebEmail) > 0 Then
'        If Len(smShttWebEmail2) > 0 Then
'            smShttWebEmail = smShttWebEmail & "," & smShttWebEmail2
'        End If
'        If Len(smShttWebEmail3) > 0 Then
'            smShttWebEmail = smShttWebEmail & "," & smShttWebEmail3
'        End If
'    End If
    
    slUsedForAtt = "Y"
    If ckcUsedFor(0).Value = vbUnchecked Then
        slUsedForAtt = "N"
    End If
    slUsedForXDigital = "N"
    If ckcUsedFor(1).Value = vbChecked Then
        slUsedForXDigital = "Y"
    End If
    slUsedForWegener = "N"
    If ckcUsedFor(2).Value = vbChecked Then
        slUsedForWegener = "Y"
    End If
    slUsedForOLA = "N"
    If ckcUsedFor(3).Value = vbChecked Then
        slUsedForOLA = "Y"
    End If
    slUsedForPledgeVsAir = "N"
    If ckcUsedFor(4).Value = vbChecked Then
        slUsedForPledgeVsAir = "Y"
    End If
    
    '5/22/07:  setting of imDMAMktCode and lmArttCode moved to BindControl
    'If imDMAMktCode = 0 Then
    '    SQLQuery = "Select shttMktCode from shtt where shttCode = " & imShttCode
    '    Set shtt_rst = gSQLSelectCall(SQLQuery)
    '    If Not shtt_rst.EOF Then
    '        imDMAMktCode = shtt_rst!shttMktCode
    '    Else
    '        imDMAMktCode = 0
    '    End If
    'End If
    'If lmArttCode = 0 Then
    '    SQLQuery = "Select shttOwnerArttCode from shtt where shttCode = " & imShttCode
    '    Set shtt_rst = gSQLSelectCall(SQLQuery)
    '    If Not shtt_rst.EOF Then
    '        lmArttCode = shtt_rst!shttOwnerArttCode
    '    Else
    '        lmArttCode = 0
    '    End If
    'End If
    
    lmArttCode = 0
    If cbcOwner.ListIndex > 1 Then
        lmArttCode = cbcOwner.GetItemData(cbcOwner.ListIndex)
    End If
    
    slIPumpID = ""
    If smWegenerIPump = "Y" Then
        slIPumpID = txtIPumpID.Text
    End If
    
    ilWebSpotsPerPage = Val(txtWebSpotsPerPage.Text)


    If iAsk Then
        If gMsgBox("Save all changes?", vbYesNo) = vbNo Then
            mSave = False
            Exit Function
        End If
    End If

    'ilFmtCode = mSaveFormat()
    'If ilFmtCode < 0 Then
    '    ilFmtCode = 0
    'End If
    
    'If adding a new station...
    If IsStatDirty = False Then
        SQLQuery = "INSERT INTO shtt (shttCallLetters, shttAddress1, "
        SQLQuery = SQLQuery & "shttWebEmail, shttWebPW, "
        SQLQuery = SQLQuery & "shttAddress2,shttCity, shttState, shttCountry, shttZip, shttSelected, "
        SQLQuery = SQLQuery & "shttEmail, shttFax, shttPhone, shttTimeZone, shttHomePage, "
        SQLQuery = SQLQuery & "shttIPumpID, "
        SQLQuery = SQLQuery & "shttOnCityMntCode, "
        SQLQuery = SQLQuery & "shttOnCountry, "
        SQLQuery = SQLQuery & "shttCityMntCode, "
        SQLQuery = SQLQuery & "shttCityLicMntCode, "
        SQLQuery = SQLQuery & "shttCountyLicMntCode, "
        SQLQuery = SQLQuery & "shttAgreementExist, "
        SQLQuery = SQLQuery & "shttCommentExist, "
        SQLQuery = SQLQuery & "shttMktRepUstCode, "
        SQLQuery = SQLQuery & "shttServRepUstCode, "
        SQLQuery = SQLQuery & "shttOperatorMntCode, "
        SQLQuery = SQLQuery & "shttAreaMntCode, "
        SQLQuery = SQLQuery & "shttHistStartDate, "
        SQLQuery = SQLQuery & "shttAudP12Plus, "
        SQLQuery = SQLQuery & "shttWatts, "
        SQLQuery = SQLQuery & "shttMonikerMntCode, "
        SQLQuery = SQLQuery & "shttMultiCastGroupID, "
        SQLQuery = SQLQuery & "shttClusterGroupID, "
        SQLQuery = SQLQuery & "shttVieroID, "
        SQLQuery = SQLQuery & "shttOnAir, "
        SQLQuery = SQLQuery & "shttStationType, "
        SQLQuery = SQLQuery & "shttACName, "
        '5/10/07:  Removed Affiliate Rep from Station File
        'SQLQuery = SQLQuery & "shttACPhone, shttArttCode, shttOwnerArttCode, shttMktCode, shttChecked, shttMarket, shttRank, shttEnterDate, "
        SQLQuery = SQLQuery & "shttACPhone, shttMntCode, shttOwnerArttCode, shttMktCode, shttChecked, shttMarket, shttRank, shttUsfCode, shttEnterDate, "
        SQLQuery = SQLQuery & "shttEnterTime, shttType, shttONAddress1, shttONAddress2, "
        SQLQuery = SQLQuery & "shttONCity, shttONState, shttONZip, "
        SQLQuery = SQLQuery & "shttFrequency, shttPermStationID, "
        SQLQuery = SQLQuery & "shttStationID, shttCityLic, shttStateLic, shttAckDaylight, "
        SQLQuery = SQLQuery & "shttWebAddress, shttFmtCode, shttSerialNo1, shttSerialNo2, shttTztCode, shttWebNumber, "
        SQLQuery = SQLQuery & "shttUsedForAtt, shttUsedForXDigital, shttUsedForWegener, shttUsedForOLA, shttPledgeVsAir, shttSentToXDSStatus, shttPort, shttMetCode, shttSpotsPerWebPage" & ")"
        SQLQuery = SQLQuery & " VALUES ('" & sCallLetters & "', '" & sAddr1 & "', "
        SQLQuery = SQLQuery & "'" & "" & "', '" & smShttWebPW & "', "
        SQLQuery = SQLQuery & "'" & sAddr2 & "', '" & sCity & "', '" & sState & "', "
        SQLQuery = SQLQuery & "'" & sCountry & "', '" & Trim$(txtZip.Text) & "'," & iSelected & ", "
        SQLQuery = SQLQuery & "'" & sEmail & "', '" & Trim$(txtFax.Text) & "', '" & Trim$(txtStaPhone.Text) & "', '" & Trim$(slTimeZone) & "', "
        SQLQuery = SQLQuery & "'" & slWebPage & "', "
        SQLQuery = SQLQuery & "'" & slIPumpID & "', "
        SQLQuery = SQLQuery & llOnCityMntCode & ", "
        SQLQuery = SQLQuery & "'" & slOnCountry & "', "
        SQLQuery = SQLQuery & llCityMntCode & ", "
        SQLQuery = SQLQuery & llCityLicMntCode & ", "
        SQLQuery = SQLQuery & llCountyLicMntCode & ", "
        SQLQuery = SQLQuery & "'" & "N" & "', "
        SQLQuery = SQLQuery & "'" & "N" & "', "
        SQLQuery = SQLQuery & ilMktRepUstCode & ", "
        SQLQuery = SQLQuery & ilServRepUstCode & ", "
        SQLQuery = SQLQuery & llOperatorMntCode & ", "
        SQLQuery = SQLQuery & llAreaMntCode & ", "
        SQLQuery = SQLQuery & "'" & Format$(slHistStartDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & llAudP12Plus & ", "
        SQLQuery = SQLQuery & llWatts & ", "
        SQLQuery = SQLQuery & llMonikerMntCode & ", "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & "'" & slEnterpriseID & "', "
        SQLQuery = SQLQuery & "'" & slOnAir & "', "
        SQLQuery = SQLQuery & "'" & slStationType & "', "
        SQLQuery = SQLQuery & "'" & sACName & "', '" & sACPhone & "', " & llMntCode & ", " & lmArttCode & ", " & imDMAMktCode & ", "
        SQLQuery = SQLQuery & iChecked & ", '" & sMarket & "', " & txtMarkRank.Text & ", " & igUstCode & ",'" & Format$(CurDate, sgSQLDateForm) & "', '" & Format$(CurTime, sgSQLTimeForm) & "', "
        SQLQuery = SQLQuery & iType & ", '" & sONAddr1 & "', '" & sONAddr2 & "', "
        SQLQuery = SQLQuery & "'" & sONCity & "', '" & sONState & "', '" & Trim(txtONZip.Text) & "', "
        SQLQuery = SQLQuery & "'" & sFrequency & "', " & lPermanentStationID & ", "
        SQLQuery = SQLQuery & lXDSStationID & ", '" & sCityLic & "', '" & sStateLic & "', " & iDaylight & ", "
        SQLQuery = SQLQuery & "'" & slWebAddress & "', " & ilFmtCode & ", '" & slSerialNo1 & "', '" & slSerialNo2 & "', " & ilTztCode & ", '" & smNewWebNumber & "', "
        'SQLQuery = SQLQuery & "'" & slUsedForAtt & "', '" & slUsedForXDigital & "', '" & slUsedForWegener & "', '" & slUsedForOLA & "', '" & slPort & "', " & "'" & imMSAMktCode & "', " & "'" & smNewMonthlyPosting & "'" & ")"
        SQLQuery = SQLQuery & "'" & slUsedForAtt & "', '" & slUsedForXDigital & "', '" & slUsedForWegener & "', '" & slUsedForOLA & "', '" & slUsedForPledgeVsAir & "', " & "'N'" & ", '" & slPort & "', " & "'" & imMSAMktCode & "', " & ilWebSpotsPerPage & ")"
        
        'If iAsk Then
        '    If gMsgBox("Save all changes?", vbYesNo) = vbYes Then
        '        'cnn.BeginTrans
        '        'cnn.Execute SQLQuery, rdExecDirect
        '        'cnn.CommitTrans
        '        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
        '            'Screen.MousePointer = vbDefault
        '            'Exit Function
        '            GoSub ErrHand
        '        End If
        '    End If
        'Else
            'cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            'cnn.CommitTrans
            If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                ''Screen.MousePointer = vbDefault
                ''Exit Function
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "Station-mSave"
                mSave = False
                Exit Function
            End If
        'End If
        
        '9/29/11: Replace Max call with matching on Call Letters.
        '         Changed to call letters to avoid the case that two people are adding stations
        '         and the wrong max is returned
        'SQLQuery = "Select MAX(shttCode) from shtt"
        SQLQuery = "Select shttCode from shtt WHERE shttCallLetters = '" & sCallLetters & "'"
        Set rst = gSQLSelectCall(SQLQuery)
        'imShttCode = rst(0).Value
        imShttCode = rst!shttCode
        
        iUpper = UBound(tgStationInfo)
        iRet = mSaveMulticast(llMulticastGroupID)
        iRet = mSaveSisterStations(llMarketClusterGroupID)
        
        tgStationInfo(iUpper).iCode = imShttCode
        tgStationInfo(iUpper).sCallLetters = sCallLetters
        tgStationInfo(iUpper).lOwnerCode = lmArttCode
        tgStationInfo(iUpper).iFormatCode = ilFmtCode
        If cbcDMAMarket.ListIndex > 1 Then
            tgStationInfo(iUpper).sMarket = Trim$(cbcDMAMarket.GetName(cbcDMAMarket.ListIndex))    'Trim$(txtMarket.Text)
        Else
            tgStationInfo(iUpper).sMarket = ""
        End If
        tgStationInfo(iUpper).iMktCode = imDMAMktCode
        '11/21/08:  added setting of StationID
        tgStationInfo(iUpper).lID = lXDSStationID
        tgStationInfo(iUpper).iType = iType
        'D.S. 8/3/01 added below so zone would show up on pledge tab with new stations
        'without having to exit the system
        tgStationInfo(iUpper).sZone = slTimeZone    'Trim$(txtTimeZone.text)
        tgStationInfo(iUpper).iTztCode = ilTztCode
        tgStationInfo(iUpper).lMntCode = llMntCode
        tgStationInfo(iUpper).sTerritory = ""
        If cbcTerritory.ListIndex > 1 Then
            tgStationInfo(iUpper).sTerritory = Trim$(cbcTerritory.GetName(cbcTerritory.ListIndex))    'Trim$(txtMarket.Text)
        Else
            tgStationInfo(iUpper).sTerritory = ""
        End If
        tgStationInfo(iUpper).lAreaMntCode = llAreaMntCode
        tgStationInfo(iUpper).sPostalName = sState
        '11/21/08:  added setting of UsedFor.....
        tgStationInfo(iUpper).sUsedForATT = slUsedForAtt
        tgStationInfo(iUpper).sUsedForXDigital = slUsedForXDigital
        tgStationInfo(iUpper).sUsedForWegener = slUsedForWegener
        tgStationInfo(iUpper).sUsedForOLA = slUsedForOLA
        tgStationInfo(iUpper).sUsedForPledgeVsAir = slUsedForPledgeVsAir
        tgStationInfo(iUpper).sSerialNo1 = slSerialNo1
        tgStationInfo(iUpper).sSerialNo2 = slSerialNo2
        tgStationInfo(iUpper).sPort = slPort
        tgStationInfo(iUpper).lPermStationID = lPermanentStationID
        tgStationInfo(iUpper).iAckDaylight = iDaylight
        tgStationInfo(iUpper).sZip = Trim$(txtZip.Text)
        tgStationInfo(iUpper).sWebAddress = slWebAddress
        tgStationInfo(iUpper).sWebPW = smShttWebPW
        tgStationInfo(iUpper).sFrequency = sFrequency
        tgStationInfo(iUpper).lMonikerMntCode = llMonikerMntCode
        tgStationInfo(iUpper).lMultiCastGroupID = llMulticastGroupID
        tgStationInfo(iUpper).lMarketClusterGroupID = llMarketClusterGroupID
        tgStationInfo(iUpper).sAgreementExist = "N"
        tgStationInfo(iUpper).sCommentExist = "N"
        tgStationInfo(iUpper).iMktRepUstCode = ilMktRepUstCode
        tgStationInfo(iUpper).iServRepUstCode = ilServRepUstCode
        tgStationInfo(iUpper).lCityLicMntCode = llCityLicMntCode
        tgStationInfo(iUpper).lHistStartDate = gDateValue(Format(slHistStartDate, sgShowDateForm))
        tgStationInfo(iUpper).sStationType = slStationType
        tgStationInfo(iUpper).lCountyLicMntCode = llCountyLicMntCode
        tgStationInfo(iUpper).sMailAddress1 = sAddr1
        tgStationInfo(iUpper).sMailAddress2 = sAddr2
        tgStationInfo(iUpper).lMailCityMntCode = llCityMntCode
        tgStationInfo(iUpper).sMailState = sState
        tgStationInfo(iUpper).sOnAir = slOnAir
        tgStationInfo(iUpper).lOperatorMntCode = llOperatorMntCode
        tgStationInfo(iUpper).lAudP12Plus = llAudP12Plus
        tgStationInfo(iUpper).lWatts = llWatts
        tgStationInfo(iUpper).sPhone = txtStaPhone.Text
        tgStationInfo(iUpper).sFax = txtFax.Text
        tgStationInfo(iUpper).sPhyAddress1 = sONAddr1
        tgStationInfo(iUpper).sPhyAddress2 = sONAddr2
        tgStationInfo(iUpper).lPhyCityMntCode = llOnCityMntCode
        tgStationInfo(iUpper).sPhyState = sONState
        tgStationInfo(iUpper).sPhyZip = Trim(txtONZip.Text)
        tgStationInfo(iUpper).sStateLic = sStateLic
        tgStationInfo(iUpper).sEnterpriseID = slEnterpriseID
        tgStationInfo(iUpper).lXDSStationID = lXDSStationID
        '8418
        tgStationInfo(iUpper).sWebNumber = smNewWebNumber
        '11/21/08:  adding tgStationInfo to tgStationInfoByCode
        tgStationInfoByCode(UBound(tgStationInfoByCode)) = tgStationInfo(iUpper)
        ReDim Preserve tgStationInfoByCode(0 To UBound(tgStationInfoByCode) + 1) As STATIONINFO
        '11/21/08: end of addition
        If UBound(tgStationInfoByCode) > 1 Then
            ArraySortTyp fnAV(tgStationInfoByCode(), 0), UBound(tgStationInfoByCode), 0, LenB(tgStationInfoByCode(0)), 0, -1, 0
        End If
        iUpper = iUpper + 1
        ReDim Preserve tgStationInfo(0 To iUpper) As STATIONINFO
        '11/21/08:  Sort Added
        If iUpper > 1 Then
            ArraySortTyp fnAV(tgStationInfo(), 0), UBound(tgStationInfo), 0, LenB(tgStationInfo(0)), 2, LenB(tgStationInfo(0).sCallLetters), 0
        End If
        
        iRet = mSaveHistory()
        'Ret = mSavePersonnel()
        udcContactGrid.StationCode = imShttCode
        udcContactGrid.Action 5
        
        '07-13-15
        'Add station to EDS and link to network - look to see if any vehicle names match the call letters and then test
        'vehicle options on insertions if yes the update the link between network and station
        '08-25-15 Verified
        If gGetEMailDistribution() Then
            blRet = gAddSingleStation(sCallLetters, slOldMaster, edcMasterStation.Text)
        End If
        
'debug 1
        'iRet = bmIgnoreEmailChange
        'iRet = mSaveEmail()
        'iRet = mSaveMulticast(llMultiCastGroupID)
        'iRet = mSaveSisterStations(llMarketClusterGroupID)
        'iRet = mOwnerMarketSave()
    'Or updating an existing station's data
    Else
        SQLQuery = "Update shtt SET "
        SQLQuery = SQLQuery & "shttCallLetters = '" & sCallLetters & "', "
        SQLQuery = SQLQuery & "shttAddress1 = '" & sAddr1 & "', "
        SQLQuery = SQLQuery & "shttAddress2 = '" & sAddr2 & "', "
        SQLQuery = SQLQuery & "shttCity = '" & sCity & "', "
        SQLQuery = SQLQuery & "shttState = '" & sState & "', "
        SQLQuery = SQLQuery & "shttCountry = '" & sCountry & "', "
        SQLQuery = SQLQuery & "shttZip = '" & Trim(txtZip.Text) & "', "
        SQLQuery = SQLQuery & "shttSelected = " & iSelected & ", "
        '7/25/11: shttEMail
        'SQLQuery = SQLQuery & "shttEmail = '" & sEmail & "', "
        SQLQuery = SQLQuery & "shttFax = '" & Trim(txtFax.Text) & "', "
        SQLQuery = SQLQuery & "shttPhone = '" & Trim(txtStaPhone.Text) & "', "
        SQLQuery = SQLQuery & "shttTimeZone = '" & Trim$(slTimeZone) & "', " 'Trim(txtTimeZone.text) & "', "
        SQLQuery = SQLQuery & "shttHomePage = '" & slWebPage & "', "
        SQLQuery = SQLQuery & "shttIPumpID = '" & slIPumpID & "', "
        SQLQuery = SQLQuery & "shttOnCityMntCode = " & llOnCityMntCode & ", "
        SQLQuery = SQLQuery & "shttOnCountry = '" & slOnCountry & "', "
        SQLQuery = SQLQuery & "shttCityMntCode = " & llCityMntCode & ", "
        SQLQuery = SQLQuery & "shttCityLicMntCode = " & llCityLicMntCode & ", "
        SQLQuery = SQLQuery & "shttCountyLicMntCode = " & llCountyLicMntCode & ", "
        SQLQuery = SQLQuery & "shttMktRepUstCode = " & ilMktRepUstCode & ", "
        SQLQuery = SQLQuery & "shttServRepUstCode = " & ilServRepUstCode & ", "
        SQLQuery = SQLQuery & "shttOperatorMntCode = " & llOperatorMntCode & ", "
        SQLQuery = SQLQuery & "shttAreaMntCode = " & llAreaMntCode & ", "
        SQLQuery = SQLQuery & "shttHistStartDate = '" & Format$(slHistStartDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "shttAudP12Plus = " & llAudP12Plus & ", "
        SQLQuery = SQLQuery & "shttWatts = " & llWatts & ", "
        SQLQuery = SQLQuery & "shttMonikerMntCode = " & llMonikerMntCode & ", "
        SQLQuery = SQLQuery & "shttVieroID = '" & slEnterpriseID & "', "
        SQLQuery = SQLQuery & "shttOnAir = '" & slOnAir & "', "
        SQLQuery = SQLQuery & "shttStationType = '" & slStationType & "', "
        '5/10/07:  Removed Affiliate Rep from Station File
        'SQLQuery = SQLQuery & "shttArttCode = " & iARIndex & ", "
        SQLQuery = SQLQuery & "shttMntCode = " & llMntCode & ", "
        SQLQuery = SQLQuery & "shttOwnerArttCode = " & lmArttCode & ", "
        SQLQuery = SQLQuery & "shttMktCode = " & imDMAMktCode & ", "
        SQLQuery = SQLQuery & "shttChecked = " & iChecked & ", "
        SQLQuery = SQLQuery & "shttMarket = '" & sMarket & "', "
        SQLQuery = SQLQuery & "shttRank = " & txtMarkRank.Text & ", "
        SQLQuery = SQLQuery & "shttUsfCode = " & igUstCode & ", "
        SQLQuery = SQLQuery & "shttEnterDate = '" & Format$(CurDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "shttEnterTime = '" & Format$(CurTime, sgSQLTimeForm) & "', "
        SQLQuery = SQLQuery & "shttType = " & iType & ", "
        SQLQuery = SQLQuery & "shttONAddress1 = '" & sONAddr1 & "', "
        SQLQuery = SQLQuery & "shttONAddress2 = '" & sONAddr2 & "', "
        SQLQuery = SQLQuery & "shttONCity = '" & sONCity & "', "
        SQLQuery = SQLQuery & "shttONState = '" & sONState & "', "
        SQLQuery = SQLQuery & "shttONZip = '" & Trim(txtONZip.Text) & "', "
        SQLQuery = SQLQuery & "shttFrequency = '" & sFrequency & "', "
        SQLQuery = SQLQuery & "shttPermStationID = " & lPermanentStationID & ", "
        SQLQuery = SQLQuery & "shttStationID = " & lXDSStationID & ", "
        SQLQuery = SQLQuery & "shttCityLic = '" & sCityLic & "', "
        SQLQuery = SQLQuery & "shttStateLic = '" & sStateLic & "'" & ", "
        SQLQuery = SQLQuery & "shttAckDaylight = " & iDaylight & ", "
        'SQLQuery = SQLQuery & "shttWebEmail = '" & smShttWebEmail & "', "
        SQLQuery = SQLQuery & "shttWebEmail = '" & "" & "', "
        SQLQuery = SQLQuery & "shttWebPW = '" & smShttWebPW & "', "
        SQLQuery = SQLQuery & "shttWebAddress = '" & slWebAddress & "', "
        SQLQuery = SQLQuery & "shttFmtCode = " & ilFmtCode & ", "
        SQLQuery = SQLQuery & "shttSerialNo1 = '" & slSerialNo1 & "', "
        SQLQuery = SQLQuery & "shttSerialNo2 = '" & slSerialNo2 & "', "
        SQLQuery = SQLQuery & "shttTztCode = " & ilTztCode & ", "
        SQLQuery = SQLQuery & "shttWebNumber = '" & smNewWebNumber & "', "
        SQLQuery = SQLQuery & "shttUsedForAtt = '" & slUsedForAtt & "', "
        SQLQuery = SQLQuery & "shttUsedForXDigital = '" & slUsedForXDigital & "', "
        SQLQuery = SQLQuery & "shttUsedForWegener = '" & slUsedForWegener & "', "
        SQLQuery = SQLQuery & "shttUsedForOLA = '" & slUsedForOLA & "', "
        SQLQuery = SQLQuery & "shttPledgeVsAir = '" & slUsedForPledgeVsAir & "', "
        SQLQuery = SQLQuery & "shttSentToXDSStatus = '" & "M" & "', "
        SQLQuery = SQLQuery & "shttPort = '" & slPort & "', "
        SQLQuery = SQLQuery & "shttMetCode = " & imMSAMktCode & ", "
        SQLQuery = SQLQuery & "shttSpotsPerWebPage = " & ilWebSpotsPerPage
        SQLQuery = SQLQuery & " WHERE (shttCode = " & imShttCode & ")"
        
        'If iAsk Then
        '    If gMsgBox("Save all changes?", vbYesNo) = vbYes Then
        '        cnn.BeginTrans
        '        'cnn.Execute SQLQuery, rdExecDirect
        '        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '            GoSub ErrHand:
        '        End If
        '        If Not mSaveHistory() Then
        '            mSave = False
        '            Exit Function
        '        End If
        '        'If Not mSavePersonnel() Then
        '        '    mSave = False
        '        '    Exit Function
        '        'End If
        '        udcContactGrid.StationCode = imShttCode
        '        udcContactGrid.Action 5
'debug 2
        '        'iRet = bmIgnoreEmailChange
        '        'If Not mSaveEmail() Then
        '        '    mSave = False
        '        '    Exit Function
        '        'End If
        '        If Not mSaveMulticast(llMultiCastGroupID) Then
        '            mSave = False
        '            Exit Function
        '        End If
        '        If Not mSaveSisterStations(llMarketClusterGroupID) Then
        '            mSave = False
        '            Exit Function
        '        End If
        '        If Not mOwnerMarketSave() Then
        '            mSave = False
        '            Exit Function
        '        End If
        '        cnn.CommitTrans
        '    End If
        'Else
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "Station-mSave"
                cnn.RollbackTrans
                mSave = False
                Exit Function
            End If
            If Not mSaveHistory() Then
                cnn.RollbackTrans
                mSave = False
                Exit Function
            End If
            'If Not mSavePersonnel() Then
            '    mSave = False
            '    Exit Function
            'End If
            udcContactGrid.StationCode = imShttCode
            udcContactGrid.Action 5
            If Not mSaveMulticast(llMulticastGroupID) Then
                cnn.RollbackTrans
                mSave = False
                Exit Function
            End If
            If Not mSaveSisterStations(llMarketClusterGroupID) Then
                cnn.RollbackTrans
                mSave = False
                Exit Function
            End If
            If Not mOwnerMarketSave() Then
                cnn.RollbackTrans
                mSave = False
                Exit Function
            End If
'debug 3
            'iRet = bmIgnoreEmailChange
            'If bmIgnoreEmailChange Then
            'If Not mSaveEmail() Then
            '    mSave = False
            '    cnn.CommitTrans
            '    Exit Function
            'End If
            'End If
            
            cnn.CommitTrans
            '9/13/11: Place back into affContactGrid
            'ilRet = gWebTestEmailChange()
        'End If
        
        '*** Handle adding and updating passwords ***
        If smExistingWebPW <> "" Then
            If smExistingWebPW <> txtWebPW.Text Then
                Call gWebUpdatePW(imShttCode)
            End If
        End If
        

'        If imWebPWUpdated Then
'            If gMsgBox("Would you like to update all agreements for this station" & Chr(13) & Chr(10) & "that have the same password?", vbYesNo) = vbYes Then
'                imWebUpdateAll = True
'            Else
'                imWebUpdateAll = False
'            End If
'            mAddWebPWToAgrmnt imWebUpdateAll
'        End If
        
        If imWebPWUpdated Then
            imWebUpdateAll = True
            mAddWebPWToAgrmnt imWebUpdateAll
        End If
        '10/3/18: Dan- I changed this code: Removed if statement.  Also in the routine gVatSetToGoToWebByShttCode I will ignore the VendorID. 10/10 Dan re-added
        'TTP 8824 reopened
        ''7912
        'Dan 8824 master sister station changed?  update all.
        If smOldMaster <> edcMasterStation.Text Then
            gVatSetToGoToWebByShttCode imShttCode, 0
            For llRow = 1 To grdMulticast.Rows - 1 Step 1
                'row could be blank
                slShtt = grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX)
                If IsNumeric(slShtt) Then
                    ilShttForVatUpdating = slShtt
                    If ilShttForVatUpdating > 0 And ilShttForVatUpdating <> imShttCode Then
                        gVatSetToGoToWebByShttCode ilShttForVatUpdating, 0
                    End If
                End If
           Next llRow
        End If
        If smExistingXDSSiteID <> txtXDSStationID.Text Then
            '7941 moved this to global
           ' mUpdateAgreementXDSSiteID
           gVatSetToGoToWebByShttCode imShttCode, Vendors.XDS_Break
        End If
        '10192
        If imExistingHonorDaylight > -1 Then
            If (lacDaylight.BackColor = GREEN & imExistingHonorDaylight = 1) Or (lacDaylight.BackColor = vbRed & imExistingHonorDaylight = 0) Then
                gVatSetToGoToWebByShttCode imShttCode, 0
            End If
        End If
        'Set it back to nothing so we don't do it again
        smExistingWebPW = ""
        
        '*** Handle adding and updating email addresses ***
'        If smExistingWebEmail <> "" Then
'            If smExistingWebEmail <> txtWebEmail.text Then
'                imWebEmailUpdated = True
'            End If
'        End If
        'D.S. 12/08/08
'        If imWebEmailUpdated Then
'            If gMsgBox("Would you like to update all agreements for this station" & Chr(13) & Chr(10) & "that have the same email address?", vbYesNo) = vbYes Then
'                imWebUpdateAll = True
'            Else
'                imWebUpdateAll = False
'            End If
'            mAddWebEmailToAgrmnt imWebUpdateAll
'        End If
'        smExistingWebEmail = ""
        
        If UBound(lmAttCodesToUpdateWeb) > 0 Then
            'D.S. 06/01/11 I dont' think that this call is useful any longer
            'Call mUpdateWebSite
        End If
        
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).iCode = imShttCode Then
                If Len(sCallLetters) > 0 Then
                    tgStationInfo(iLoop).sCallLetters = sCallLetters
                End If
                tgStationInfo(iLoop).lOwnerCode = lmArttCode
                tgStationInfo(iLoop).iFormatCode = ilFmtCode
                'If (Len(cboDMAMarketCluster.Text) > 0) And (cboDMAMarketCluster.ListIndex > 1) Then
                '    'Doug- On 11/17/06 changed which name is used
                '    tgStationInfo(iLoop).sMarket = Trim$(cboDMAMarketCluster.Text) 'sMarket
                'End If
                If cbcDMAMarket.ListIndex > 1 Then
                    tgStationInfo(iLoop).sMarket = Trim$(cbcDMAMarket.GetName(cbcDMAMarket.ListIndex))    'Trim$(txtMarket.Text)
                Else
                    tgStationInfo(iLoop).sMarket = ""
                End If
                tgStationInfo(iLoop).iMktCode = imDMAMktCode
                tgStationInfo(iLoop).lID = lXDSStationID
                tgStationInfo(iLoop).iType = iType
                tgStationInfo(iLoop).sZone = Trim$(slTimeZone)  'Trim$(txtTimeZone.text)
                tgStationInfo(iLoop).iTztCode = ilTztCode
                tgStationInfo(iLoop).lMntCode = llMntCode
                tgStationInfo(iLoop).sTerritory = ""
                If cbcTerritory.ListIndex > 1 Then
                    tgStationInfo(iLoop).sTerritory = Trim$(cbcTerritory.GetName(cbcTerritory.ListIndex))    'Trim$(txtMarket.Text)
                Else
                    tgStationInfo(iLoop).sTerritory = ""
                End If
                tgStationInfo(iLoop).lAreaMntCode = llAreaMntCode
                tgStationInfo(iLoop).sPostalName = sState
                '11/21/08:  Added setting of UsedFor.....
                tgStationInfo(iLoop).sUsedForATT = slUsedForAtt
                tgStationInfo(iLoop).sUsedForXDigital = slUsedForXDigital
                tgStationInfo(iLoop).sUsedForWegener = slUsedForWegener
                tgStationInfo(iLoop).sUsedForOLA = slUsedForOLA
                tgStationInfo(iLoop).sUsedForPledgeVsAir = slUsedForPledgeVsAir
                tgStationInfo(iLoop).sSerialNo1 = slSerialNo1
                tgStationInfo(iLoop).sSerialNo2 = slSerialNo2
                tgStationInfo(iLoop).sPort = slPort
                tgStationInfo(iLoop).lPermStationID = lPermanentStationID
                tgStationInfo(iLoop).iAckDaylight = iDaylight
                tgStationInfo(iLoop).sZip = Trim$(txtZip.Text)
                tgStationInfo(iLoop).sWebAddress = slWebAddress
                tgStationInfo(iLoop).sWebPW = smShttWebPW
                tgStationInfo(iLoop).sFrequency = sFrequency
                tgStationInfo(iLoop).lMonikerMntCode = llMonikerMntCode
                tgStationInfo(iLoop).lMultiCastGroupID = llMulticastGroupID
                tgStationInfo(iLoop).lMarketClusterGroupID = llMarketClusterGroupID
                tgStationInfo(iLoop).iMktRepUstCode = ilMktRepUstCode
                tgStationInfo(iLoop).iServRepUstCode = ilServRepUstCode
                tgStationInfo(iLoop).lCityLicMntCode = llCityLicMntCode
                tgStationInfo(iLoop).lHistStartDate = gDateValue(Format(slHistStartDate, sgShowDateForm))
                tgStationInfo(iLoop).sStationType = slStationType
                tgStationInfo(iLoop).lCountyLicMntCode = llCountyLicMntCode
                tgStationInfo(iLoop).sMailAddress1 = sAddr1
                tgStationInfo(iLoop).sMailAddress2 = sAddr2
                tgStationInfo(iLoop).lMailCityMntCode = llCityMntCode
                tgStationInfo(iLoop).sMailState = sState
                tgStationInfo(iLoop).sOnAir = slOnAir
                tgStationInfo(iLoop).lOperatorMntCode = llOperatorMntCode
                tgStationInfo(iLoop).lAudP12Plus = llAudP12Plus
                tgStationInfo(iLoop).lWatts = llWatts
                tgStationInfo(iLoop).sPhone = txtStaPhone.Text
                tgStationInfo(iLoop).sFax = txtFax.Text
                tgStationInfo(iLoop).sPhyAddress1 = sONAddr1
                tgStationInfo(iLoop).sPhyAddress2 = sONAddr2
                tgStationInfo(iLoop).lPhyCityMntCode = llOnCityMntCode
                tgStationInfo(iLoop).sPhyState = sONState
                tgStationInfo(iLoop).sPhyZip = Trim(txtONZip.Text)
                tgStationInfo(iLoop).sStateLic = sStateLic
                tgStationInfo(iLoop).sEnterpriseID = slEnterpriseID
                tgStationInfo(iLoop).lXDSStationID = lXDSStationID
                '8418
                tgStationInfo(iLoop).sWebNumber = smNewWebNumber
                '11/21/08: Sort added
                If iUpper > 1 Then
                    ArraySortTyp fnAV(tgStationInfo(), 0), UBound(tgStationInfo), 0, LenB(tgStationInfo(0)), 2, LenB(tgStationInfo(0).sCallLetters), 0
                End If
                '11/21/08:  End of addition
                Exit For
            End If
        Next iLoop
        '11/21/08:  Setting tgStationInfoByCode
        iLoop = gBinarySearchStationInfoByCode(imShttCode)
        If iLoop <> -1 Then
            If Len(sCallLetters) > 0 Then
                tgStationInfoByCode(iLoop).sCallLetters = sCallLetters
            End If
            tgStationInfoByCode(iLoop).lOwnerCode = lmArttCode
            tgStationInfoByCode(iLoop).iFormatCode = ilFmtCode
            'If (Len(cboDMAMarketCluster.Text) > 0) And (cboDMAMarketCluster.ListIndex > 1) Then
            '    tgStationInfoByCode(iLoop).sMarket = Trim$(cboDMAMarketCluster.Text) 'sMarket
            'End If
            If cbcDMAMarket.ListIndex > 1 Then
                tgStationInfoByCode(iLoop).sMarket = Trim$(cbcDMAMarket.GetName(cbcDMAMarket.ListIndex))    'Trim$(txtMarket.Text)
            Else
                tgStationInfoByCode(iLoop).sMarket = ""
            End If
            tgStationInfoByCode(iLoop).iMktCode = imDMAMktCode
            tgStationInfoByCode(iLoop).lID = lXDSStationID
            tgStationInfoByCode(iLoop).iType = iType
            tgStationInfoByCode(iLoop).sZone = Trim$(slTimeZone)  'Trim$(txtTimeZone.text)
            tgStationInfoByCode(iLoop).iTztCode = ilTztCode
            tgStationInfoByCode(iLoop).lMntCode = llMntCode
            If cbcTerritory.ListIndex > 1 Then
                tgStationInfoByCode(iLoop).sTerritory = Trim$(cbcTerritory.GetName(cbcTerritory.ListIndex))    'Trim$(txtMarket.Text)
            Else
                tgStationInfoByCode(iLoop).sTerritory = ""
            End If
            tgStationInfoByCode(iLoop).lAreaMntCode = llAreaMntCode
            tgStationInfoByCode(iLoop).sPostalName = sState
            tgStationInfoByCode(iLoop).sUsedForATT = slUsedForAtt
            tgStationInfoByCode(iLoop).sUsedForXDigital = slUsedForXDigital
            tgStationInfoByCode(iLoop).sUsedForWegener = slUsedForWegener
            tgStationInfoByCode(iLoop).sUsedForOLA = slUsedForOLA
            tgStationInfoByCode(iLoop).sUsedForPledgeVsAir = slUsedForPledgeVsAir
            tgStationInfoByCode(iLoop).sSerialNo1 = slSerialNo1
            tgStationInfoByCode(iLoop).sSerialNo2 = slSerialNo2
            tgStationInfoByCode(iLoop).sPort = slPort
            tgStationInfoByCode(iLoop).lPermStationID = lPermanentStationID
            tgStationInfoByCode(iLoop).iAckDaylight = iDaylight
            tgStationInfoByCode(iLoop).sZip = Trim$(txtZip.Text)
            tgStationInfoByCode(iLoop).sWebAddress = slWebAddress
            tgStationInfoByCode(iLoop).sWebPW = smShttWebPW
            tgStationInfoByCode(iLoop).sFrequency = sFrequency
            tgStationInfoByCode(iLoop).lMonikerMntCode = llMonikerMntCode
            tgStationInfoByCode(iLoop).lMultiCastGroupID = llMulticastGroupID
            tgStationInfoByCode(iLoop).lMarketClusterGroupID = llMarketClusterGroupID
            tgStationInfoByCode(iLoop).iMktRepUstCode = ilMktRepUstCode
            tgStationInfoByCode(iLoop).iServRepUstCode = ilServRepUstCode
            tgStationInfoByCode(iLoop).lCityLicMntCode = llCityLicMntCode
            tgStationInfoByCode(iLoop).lHistStartDate = gDateValue(Format(slHistStartDate, sgShowDateForm))
            tgStationInfoByCode(iLoop).sStationType = slStationType
            tgStationInfoByCode(iLoop).lCountyLicMntCode = llCountyLicMntCode
            tgStationInfoByCode(iLoop).sMailAddress1 = sAddr1
            tgStationInfoByCode(iLoop).sMailAddress2 = sAddr2
            tgStationInfoByCode(iLoop).lMailCityMntCode = llCityMntCode
            tgStationInfoByCode(iLoop).sOnAir = slOnAir
            tgStationInfoByCode(iLoop).lOperatorMntCode = llOperatorMntCode
            tgStationInfoByCode(iLoop).lAudP12Plus = llAudP12Plus
            tgStationInfoByCode(iLoop).lWatts = llWatts
            tgStationInfoByCode(iLoop).sPhone = txtStaPhone.Text
            tgStationInfoByCode(iLoop).sFax = txtFax.Text
            tgStationInfoByCode(iLoop).sPhyAddress1 = sONAddr1
            tgStationInfoByCode(iLoop).sPhyAddress2 = sONAddr2
            tgStationInfoByCode(iLoop).lPhyCityMntCode = llOnCityMntCode
            tgStationInfoByCode(iLoop).sPhyState = sONState
            tgStationInfoByCode(iLoop).sPhyZip = Trim(txtONZip.Text)
            tgStationInfoByCode(iLoop).sStateLic = sStateLic
            tgStationInfoByCode(iLoop).sEnterpriseID = slEnterpriseID
            tgStationInfoByCode(iLoop).lXDSStationID = lXDSStationID
            '8418
            tgStationInfoByCode(iLoop).sWebNumber = smNewWebNumber
        End If
        '11/21/08:  End of addition
        mGetHistory
        'mGetPersonnel
        'mGetEmail
        'udcContactGrid.StationCode = imShttCode
        'udcContactGrid.Action 5
        'mSort
        If StrComp(slTimeZone, smCurTimeZone, 1) <> 0 Then
            mZoneChange slTimeZone
        End If
        'D.S. 5/3/16 TTP 7994
        If IsNumeric(edcWebNumber.Text) Then
            smNewWebNumber = Trim(edcWebNumber.Text)
        Else
            smNewWebNumber = "1"
        End If
        'smNewWebNumber = Trim(edcWebNumber.Text)
        'If rbcWebSiteVersion(0).Value = True Then
        If gTestAccessToWebServer Then
            ilTotalUpdated = gExecWebSQLWithRowsEffected("Update Header Set WebSiteVersion = " & CInt(smNewWebNumber) & " Where StationName = " & "'" & smCurCallLetters & "'")
            If ilTotalUpdated = -1 Then
                gLogMsg "Error: Failed to update the web's Web Number - frmStations mSave", "AffErrorLog.Txt", False
                gMsgBox "Error: Failed to update the web's Web Number - frmStations mSave", vbCritical
            End If
        Else
            MsgBox "No Access to the web server."
            gLogMsg "Error: Failed to get access to the web. - frmStations mSave", "AffErrorLog.Txt", False
        End If
    End If
    '08-25-15 Verified
    If gGetEMailDistribution Then
        blRet = gAddSingleStation(sCallLetters, smOldMaster, edcMasterStation.Text)
    End If
    '11/26/17: Set Changed date/time
    gFileChgdUpdate "shtt.mkd", False
    cboStations.SetFocus
    imFieldChgd = False
    mSort
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mSave"
End Function

Private Sub optSort_GotFocus(Index As Integer)
    imIgnoreTabs = True
    mHistSetShow
    udcContactGrid.Action 1 'Clear focus
End Sub

Private Sub optSP_Click(Index As Integer)
    If optSP(Index).Value = False Then
        Exit Sub
    End If
    If Index = 0 Then
        labName.Caption = "Call Letters:"
        txtXDSStationID.Enabled = True
        edcFrequency.Enabled = True
        edcPermanentStationID.Enabled = True
    Else
        labName.Caption = "Name:"
        edcFrequency.Text = ""
        edcFrequency.Enabled = False
        edcPermanentStationID.Text = ""
        edcPermanentStationID.Enabled = False
        txtXDSStationID.Enabled = False
        txtXDSStationID.Text = ""
    End If
    Screen.MousePointer = vbHourglass
    mSort
    Screen.MousePointer = vbDefault
End Sub

Private Sub optSP_GotFocus(Index As Integer)
    imIgnoreTabs = False
End Sub

Private Sub pbcClickFocus_GotFocus()
    mHistSetShow
    udcContactGrid.Action 1 'Clear focus
End Sub

Private Sub pbcHistSTab_GotFocus()
    If GetFocus() <> pbcHistSTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(1) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mHistEnableBox
        Exit Sub
    End If
    If txtHistory.Visible Then
        mHistSetShow
        If grdHistory.Col = SHCALLLETTERSINDEX Then
            If grdHistory.Row > grdHistory.FixedRows Then
                lmTopRow = -1
                grdHistory.Row = grdHistory.Row - 1
                If Not grdHistory.RowIsVisible(grdHistory.Row) Then
                    grdHistory.TopRow = grdHistory.TopRow - 1
                End If
                grdHistory.Col = SHLASTDATEINDEX
                mHistEnableBox
            Else
'                'cmdDone.SetFocus
'                SendKeys "%P", True
'                ' txtPD.SetFocus
            End If
        Else
            grdHistory.Col = grdHistory.Col - 1
            mHistEnableBox
        End If
    Else
        lmTopRow = -1
        grdHistory.TopRow = grdHistory.FixedRows
        grdHistory.Col = SHCALLLETTERSINDEX
        grdHistory.Row = grdHistory.FixedRows
        mHistEnableBox
    End If
End Sub

Private Sub pbcHistTab_GotFocus()
    Dim llRow As Long
    
    If GetFocus() <> pbcHistTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(1) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If txtHistory.Visible Then
        mHistSetShow
        If grdHistory.Col = grdHistory.Cols - 2 Then
            llRow = grdHistory.Rows
            Do
                llRow = llRow - 1
            Loop While grdHistory.TextMatrix(llRow, SHCALLLETTERSINDEX) = ""
            llRow = llRow + 1
            If (grdHistory.Row + 1 < llRow) Then
                lmTopRow = -1
                grdHistory.Row = grdHistory.Row + 1
                If Not grdHistory.RowIsVisible(grdHistory.Row) Then
                    grdHistory.TopRow = grdHistory.TopRow + 1
                End If
                grdHistory.Col = SHCALLLETTERSINDEX
                If Trim$(grdHistory.TextMatrix(grdHistory.Row, SHCALLLETTERSINDEX)) <> "" Then
                    mHistEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdHistory.Left - pbcArrow.Width, grdHistory.Top + grdHistory.RowPos(grdHistory.Row) + (grdHistory.RowHeight(grdHistory.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If txtHistory.Text <> "" Then
                    lmTopRow = -1
                    If grdHistory.Row + 1 >= grdHistory.Rows Then
                        grdHistory.AddItem ""
                    End If
                    grdHistory.Row = grdHistory.Row + 1
                    If Not grdHistory.RowIsVisible(grdHistory.Row) Then
                        grdHistory.TopRow = grdHistory.TopRow + 1
                    End If
                    grdHistory.Col = SHCALLLETTERSINDEX
                    'mHistEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdHistory.Left - pbcArrow.Width, grdHistory.Top + grdHistory.RowPos(grdHistory.Row) + (grdHistory.RowHeight(grdHistory.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdHistory.Col = grdHistory.Col + 1
            mHistEnableBox
        End If
    Else
        lmTopRow = -1
        grdHistory.TopRow = grdHistory.FixedRows
        grdHistory.Col = SHCALLLETTERSINDEX
        grdHistory.Row = grdHistory.FixedRows
        mHistEnableBox
    End If
End Sub


Private Sub pbcSTab_GotFocus(Index As Integer)
    If imIgnoreTabs Then
        imIgnoreTabs = False
'        If frcTab(0).Visible Then
'            'If optSP(0).Value Then
'            '    optSP(0).SetFocus
'            'ElseIf optSP(1).Value Then
'            '    optSP(1).SetFocus
'            'Else
'                txtCallLetters.SetFocus
'            'End If
'        ElseIf frcTab(1).Visible Then
'            'If optPrefCon(0).Value Then
'            '    optPrefCon(0).SetFocus
'            'ElseIf optPrefCon(1).Value Then
'            '    optPrefCon(1).SetFocus
'            'ElseIf optPrefCon(2).Value Then
'            '    optPrefCon(2).SetFocus
'            'ElseIf optPrefCon(3).Value Then
'            '    optPrefCon(3).SetFocus
'            'ElseIf optPrefCon(4).Value Then
'            '    optPrefCon(4).SetFocus
'            'ElseIf optPrefCon(5).Value Then
'            '    optPrefCon(5).SetFocus
'            'Else
'            '    txtEmail.SetFocus
'            'End If
'            If txtFax.Visible Then
'                txtFax.SetFocus
'            End If
'        Else
'            'txtPD.SetFocus
'        End If
        Exit Sub
    End If
    mHistSetShow
'    If Index = 0 Then
        'SendKeys "%P", True
        'cmdAffRep.SetFocus
        cmdDone.SetFocus
'    ElseIf Index = 1 Then
'        SendKeys "%L", True
'        'txtTimeZone.SetFocus
'        cboTimeZone.SetFocus
'    ElseIf Index = 2 Then
'        SendKeys "%C", True
'        'txtWebPage.SetFocus
'        If txtONZip.Visible Then
'            txtONZip.SetFocus
'        End If
'    End If
End Sub

Private Sub pbcTab_GotFocus(Index As Integer)
    If imIgnoreTabs Then
        imIgnoreTabs = False
'        If frcTab(0).Visible Then
'            'txtTimeZone.SetFocus
'            cboTimeZone.SetFocus
'        ElseIf frcTab(1).Visible Then
'            'txtWebPage.SetFocus
'            If txtONZip.Visible Then
'                txtONZip.SetFocus
'            End If
'        ElseIf frcTab(2).Visible Then
'            '5/10/07:  Removed Affiliate Rep from Station File
'            'cmdAffRep.SetFocus
'        ElseIf frcTab(3).Visible Then
'            cmdDone.SetFocus
'        End If
        Exit Sub
    End If
    mHistSetShow
'    If Index = 0 Then
'        SendKeys "%C", True
'        'If optPrefCon(0).Value Then
'        '    optPrefCon(0).SetFocus
'        'ElseIf optPrefCon(1).Value Then
'        '    optPrefCon(1).SetFocus
'        'ElseIf optPrefCon(2).Value Then
'        '    optPrefCon(2).SetFocus
'        'ElseIf optPrefCon(3).Value Then
'        '    optPrefCon(3).SetFocus
'        'ElseIf optPrefCon(4).Value Then
'        '    optPrefCon(4).SetFocus
'        'Else
'        '    txtEmail.SetFocus
'        'End If
'        If txtFax.Visible Then
'            txtFax.SetFocus
'        End If
'    ElseIf Index = 1 Then
'        SendKeys "%P", True
'        'txtPD.SetFocus
'    ElseIf Index = 2 Then
''        SendKeys "%L", True
''        'If optSP(0).Value Then
''        '    optSP(0).SetFocus
''        'ElseIf optSP(1).Value Then
''        '    optSP(1).SetFocus
''        'Else
''        '    txtCallLetters.SetFocus
''        'End If
''        cmdDone.SetFocus
'        SendKeys "%H", True
'        pbcHistSTab.SetFocus
'    End If
End Sub

Private Sub rbcMulticast_Click(Index As Integer)
    If rbcMulticast(Index).Value Then
        bmIgnoreMulticastChange = True
        If Index = 0 Then   'Create Multicast Group
            rbcMulticastOwner(0).Value = True
            rbcMulticastOwner(0).Enabled = True
            rbcMulticastOwner(1).Enabled = True
            rbcMulticastMarket(0).Value = True
            rbcMulticastMarket(0).Enabled = True
            rbcMulticastMarket(1).Enabled = True
            
            rbcMulticastOwner_Add(0).Enabled = False
            rbcMulticastOwner_Add(1).Enabled = False
            rbcMulticastMarket_Add(0).Enabled = False
            rbcMulticastMarket_Add(1).Enabled = False
            
            mPopMulticast
        ElseIf Index = 1 Then   'Add to Multicast Group
            rbcMulticastOwner(0).Value = False
            rbcMulticastOwner(1).Value = False
            rbcMulticastOwner(0).Enabled = False
            rbcMulticastOwner(1).Enabled = False
            rbcMulticastMarket(0).Value = False
            rbcMulticastMarket(1).Value = False
            rbcMulticastMarket(0).Enabled = False
            rbcMulticastMarket(1).Enabled = False
            
            rbcMulticastOwner_Add(0).Value = True
            rbcMulticastOwner_Add(0).Enabled = True
            rbcMulticastOwner_Add(1).Enabled = True
            rbcMulticastMarket_Add(0).Value = True
            rbcMulticastMarket_Add(0).Enabled = True
            rbcMulticastMarket_Add(1).Enabled = True
            
            mPopMulticast
        Else    'Remove from Multicast Group
            rbcMulticastOwner(0).Value = False
            rbcMulticastOwner(1).Value = False
            rbcMulticastOwner(0).Enabled = False
            rbcMulticastOwner(1).Enabled = False
            rbcMulticastMarket(0).Value = False
            rbcMulticastMarket(1).Value = False
            rbcMulticastMarket(0).Enabled = False
            rbcMulticastMarket(1).Enabled = False
            
            rbcMulticastOwner_Add(0).Enabled = False
            rbcMulticastOwner_Add(1).Enabled = False
            rbcMulticastMarket_Add(0).Enabled = False
            rbcMulticastMarket_Add(1).Enabled = False
        End If
        bmIgnoreMulticastChange = False
        mPopMulticast
    End If
End Sub

Private Sub rbcMulticastMarket_Click(Index As Integer)
    If bmIgnoreMulticastChange Then
        Exit Sub
    End If
    If rbcMulticastMarket(Index).Value Then
        mPopMulticast
    End If
End Sub

Private Sub rbcMulticastOwner_Click(Index As Integer)
    If bmIgnoreMulticastChange Then
        Exit Sub
    End If
    If rbcMulticastOwner(Index).Value Then
        mPopMulticast
    End If
End Sub

Private Sub tscStation_BeforeClick(Cancel As Integer)

    Dim ilRet As Integer
    
    If tscStation.Tabs(1).Selected Then
        If cboStations.Text = "[New]" And txtCallLetters.Text = "" Then
            gMsgBox "Please Select a Station Before Continuing", vbOKOnly
            Cancel = True
            Exit Sub
        End If
        Cancel = False
    End If

    Exit Sub

End Sub

Private Sub tscStation_Click()

    Dim ilRow As Integer
    Dim ilLen As Integer
    Dim ilRet As Integer
    
    If imTabIndex = tscStation.SelectedItem.Index Then
        Exit Sub
    End If
    mHistSetShow
    udcContactGrid.Action 1 'Clear focus
    'frcTab(tscStation.SelectedItem.Index - 1).Visible = True
    'frcTab(imTabIndex - 1).Visible = False
    frcTab(0).Visible = False
    frcTab(1).Visible = False
    frcTab(2).Visible = False
    frcTab(3).Visible = False
    frcTab(4).Visible = False
    frcTab(5).Visible = False
    udcContactGrid.Action 1
    imTabIndex = tscStation.SelectedItem.Index
    Select Case imTabIndex
        Case 1  'Main
            frcTab(0).Visible = True
        Case 2  'History
            frcTab(3).Visible = True
            grdHistory.Redraw = True
        Case 3  'Personnel
            frcTab(2).Visible = True
        Case 5  'Multi-Cast
            lblOwner.Caption = ""
            If cbcOwner.ListIndex > 1 Then
                lblOwner.Caption = "Owner: " & Trim$(cbcOwner.GetName(cbcOwner.ListIndex))
            End If
            If cbcDMAMarket.ListIndex > 1 Then
                lblOwner.Caption = lblOwner.Caption & " DMA Market: " & cbcDMAMarket.GetName(cbcDMAMarket.ListIndex)
            End If
            lblOwner.Caption = Trim$(lblOwner.Caption)
            'lacMulticastNote.Caption = "Multicast " & Trim$(txtCallLetters.Text) & " with:"
            frcTab(4).Visible = True
        Case 4  'Sister Stations
            lacMarketCluster.Caption = ""
            If cbcOwner.ListIndex > 1 Then
                lacMarketCluster.Caption = "Owner: " & Trim$(cbcOwner.GetName(cbcOwner.ListIndex))
            End If
            If cbcDMAMarket.ListIndex > 1 Then
                lacMarketCluster.Caption = lacMarketCluster.Caption & " DMA Market: " & cbcDMAMarket.GetName(cbcDMAMarket.ListIndex)
            End If
            lacMarketCluster.Caption = Trim$(lacMarketCluster.Caption)
            frcTab(5).Visible = True
        Case 6  'Interface
            frcTab(1).Visible = True
    End Select
    'If imTabIndex = 5 Then
    '    lblOwner.Caption = "Owner: " & Trim$(cboOwner.Text)
    '    mMulticastInit
    'End If
    
    
    'If imTabIndex = 6 Then
    '    lbcDMAMarketCluster.Height = frcMarketCluster.Height - lbcDMAMarketCluster.Top
    '    lbcMSAMarketCluster.Height = lbcDMAMarketCluster.Height
    '    'ilRet = mDMAMarketFillListBox(0, False, imDMAMktCode)
    '    'ilRet = mDMAMarketFillStation(lmArttCode, imDMAMktCode)
    '    'ilRet = mMSAMarketFillListBox(0, False, imMSAMktCode)
    '    'ilRet = mMSAMarketFillStation(lmArttCode, imMSAMktCode)
    '    'ilRet = mOwnerFillListBox(0, True, lmArttCode)
    'End If
    imIgnoreTabs = True
    
End Sub

Private Sub tscStation_GotFocus()
    mHistSetShow
    udcContactGrid.Action 1 'Clear focus
End Sub

Private Sub txtAddr1_Change()
    imFieldChgd = True
End Sub

Private Sub txtAddr1_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtAddr2_Change()
    imFieldChgd = True
End Sub

Private Sub txtAddr2_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub


Private Sub txtPort_Change()
    imFieldChgd = True
End Sub

Private Sub txtPort_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtSerialNo2_Change()
    imFieldChgd = True
End Sub

Private Sub txtSerialNo2_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtCallLetters_Change()
    imFieldChgd = True
End Sub

Private Sub txtCallLetters_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtCallLetters_KeyPress(KeyAscii As Integer)
    If optSP(0).Value Then
        If KeyAscii >= 97 And KeyAscii <= 122 Then
            KeyAscii = KeyAscii - 32
        End If
    End If
End Sub


Private Sub edcCountry_Change()
    imFieldChgd = True
End Sub

Private Sub edcCountry_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtFax_Change()
    imFieldChgd = True
End Sub

Private Sub txtFax_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtHistory_Change()
    Dim slStr As String
    
    Select Case grdHistory.Col
        Case SHCALLLETTERSINDEX
            If grdHistory.Text <> txtHistory.Text Then
                imFieldChgd = True
            End If
            grdHistory.Text = txtHistory.Text
        Case SHLASTDATEINDEX
            slStr = txtHistory.Text
            If gIsDate(slStr) Then
                If grdHistory.Text <> txtHistory.Text Then
                    imFieldChgd = True
                End If
                grdHistory.Text = txtHistory.Text
            End If
    End Select
End Sub

Private Sub txtHistory_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtMarket_Change()
    imFieldChgd = True
End Sub

Private Sub txtMarket_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtMarkRank_Change()
    imFieldChgd = True
End Sub

Private Sub txtMarkRank_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtONAddr1_Change()
    imFieldChgd = True
End Sub

Private Sub txtONAddr1_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtONAddr2_Change()
    imFieldChgd = True
End Sub

Private Sub txtONAddr2_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtONZip_Change()
    imFieldChgd = True
End Sub

Private Sub txtONZip_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtSerialNo1_Change()
    imFieldChgd = True
End Sub

Private Sub txtSerialNo1_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtStaPhone_Change()
    imFieldChgd = True
End Sub

Private Sub txtStaPhone_GotFocus()
    imIgnoreTabs = False
    smOldStaPhoneNum = txtStaPhone.Text
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtStaPhone_LostFocus()
    
    udcContactGrid.PhoneNumber = Trim$(txtStaPhone.Text)
    'D.S. 10/1/02
    'Rule: If the user changes the phone number in the Stations screen on the Locations
    'tab and it mathes any of the other phone numbers on the Personnel tab or on the
    'Agreements screen contact then change the number to the new number.  This is only true
    'if changed on the Locations tab in Stations.  We do not change all matching numbers if
    'the number is updated anywhere else. Per conversation w/Charles S. at Global 10/1/02.
    If Trim$(smOldStaPhoneNum) <> Trim$(txtStaPhone.Text) Then

' JD 10/24/05 Commented out for now until determined if this will still be used.
'        If Trim$(txtPDPhone.Text) = Trim$(smOldStaPhoneNum) Then
'            txtPDPhone.Text = Trim$(txtStaPhone.Text)
'        End If
'        If Trim$(txtACPhone.Text) = Trim$(smOldStaPhoneNum) Then
'            txtACPhone.Text = Trim$(txtStaPhone.Text)
'        End If
'        If Trim$(txtTDPhone.Text) = Trim$(smOldStaPhoneNum) Then
'            txtTDPhone.Text = Trim$(txtStaPhone.Text)
'        End If
'        If Trim$(txtMDPhone.Text) = Trim$(smOldStaPhoneNum) Then
'            txtMDPhone.Text = Trim$(txtStaPhone.Text)
'        End If
    End If
    'End 10/1/02

End Sub

Private Sub txtXDSStationID_Change()
    imFieldChgd = True
End Sub

Private Sub txtXDSStationID_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtXDSStationID_KeyPress(KeyAscii As Integer)
    'If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub


Private Sub txtWebAddress_Change()
    imFieldChgd = True
End Sub

Private Sub txtWebAddress_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtWebEmail_Change()
    imFieldChgd = True
    imWebEmailUpdated = True
End Sub

Private Sub txtWebEmail_GotFocus()
    imIgnoreTabs = False
    smExistingWebEmail = txtWebEmail.Text
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtWebEmail_LostFocus()

    Dim ilRet As Integer
    
    ilRet = gTestForMultipleEmail(txtWebEmail.Text, "Reg")
    If ilRet = False Then
        gMsgBox sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Web Email Address Before Continuing", vbExclamation
        gLogMsg sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Web Email Address Before Continuing", "WebEmailLog.Txt", False
        'If txtWebEmail.Visible Then
        '    txtWebEmail.SetFocus
        'End If
        Screen.MousePointer = vbDefault
    End If
    
End Sub



Private Sub txtWebPage_Change()
    imFieldChgd = True
End Sub

Private Sub txtWebPage_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtWebPW_Change()
    imFieldChgd = True
End Sub

Private Sub txtWebPW_GotFocus()
    smExistingWebPW = txtWebPW.Text
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtZip_Change()
    imFieldChgd = True
End Sub

Private Sub txtZip_GotFocus()
    imIgnoreTabs = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub mSort()
    Dim iLoop As Integer
    Dim iIndex As Integer
    Dim mkt_rst As ADODB.Recordset
    Dim shtt_rst As ADODB.Recordset
    Dim slMktName As String
    Dim slSaveStationName As String
    Dim lRow As Long
    Dim ilPos As Integer
    
    cboStations.Visible = False
    Screen.MousePointer = vbHourglass
    slSaveStationName = Trim(cboStations.Text)
    ilPos = InStr(slSaveStationName, " ")
    If ilPos > 0 Then
        slSaveStationName = Left(slSaveStationName, ilPos - 1)
    End If
    mPop
    mClearControls
    imShttCode = -1
    cboStations.Text = ""
    cboStations.Clear
    If sgStationCallSource = "S" Then
        For lRow = frmStationSearch!grdStations.FixedRows To frmStationSearch!grdStations.Rows - 1 Step 1
            If Trim$(frmStationSearch!grdStations.TextMatrix(lRow, SCALLLETTERINDEX)) <> "" Then
                cboStations.AddItem Trim$(frmStationSearch!grdStations.TextMatrix(lRow, SCALLLETTERINDEX)) & ", " & Trim$(frmStationSearch!grdStations.TextMatrix(lRow, SDMAMARKETINDEX))
                cboStations.ItemData(cboStations.NewIndex) = Val(frmStationSearch!grdStations.TextMatrix(lRow, SSHTTCODEINDEX))
            End If
        Next lRow
        cboStations.AddItem "[New]", 0
        cboStations.ItemData(0) = 0
    Else
        If optSort(0).Value = True Then
            For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                If ((optSP(0).Value) And (tgStationInfo(iLoop).iType = 0)) Or ((optSP(1).Value) And (tgStationInfo(iLoop).iType = 1)) Then
                    slMktName = Trim$(tgStationInfo(iLoop).sMarket)
                    cboStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & slMktName
                    cboStations.ItemData(cboStations.NewIndex) = tgStationInfo(iLoop).iCode
                End If
            Next iLoop
            cboStations.AddItem "[New]", 0
            cboStations.ItemData(0) = 0
        Else
            For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                If ((optSP(0).Value) And (tgStationInfo(iLoop).iType = 0)) Or ((optSP(1).Value) And (tgStationInfo(iLoop).iType = 1)) Then
                    slMktName = Trim$(tgStationInfo(iLoop).sMarket)
                    If StrComp(slMktName, "", 1) = 0 Then
                        slMktName = "No Market"
                    End If
                    cboStations.AddItem slMktName & ", " & Trim$(tgStationInfo(iLoop).sCallLetters)
                    cboStations.ItemData(cboStations.NewIndex) = tgStationInfo(iLoop).iCode
                End If
            Next iLoop
            cboStations.AddItem "[New]", 0
            cboStations.ItemData(0) = 0
        End If
    End If
    cboStations.ListIndex = 0
    
    If Len(slSaveStationName) > 0 Then
        lRow = SendMessageByString(cboStations.hwnd, CB_FINDSTRING, -1, slSaveStationName)
        If lRow >= 0 Then
            cboStations.ListIndex = lRow
        End If
    End If
    cboStations.Visible = True
    
    cboStations.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Function mSaveHistory() As Integer
    Dim iLoop As Integer
    Dim iRow As Integer
    Dim iFound As Integer
    Dim lCltCode As Long
    Dim sCallLetters As String
    Dim sEndDate As String
    
    On Error GoTo ErrHand
    
    'Delete all records
    SQLQuery = "DELETE FROM clt WHERE (cltShfCode = " & imShttCode & ")"
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "Station-mSaveHistory"
        mSaveHistory = False
        Exit Function
    End If
    If igHistoryStatus = 1 Then
        SQLQuery = "INSERT INTO clt (cltShfCode, cltCallLetters, cltEndDate) "
        SQLQuery = SQLQuery & " VALUES ( " & imShttCode & ", '" & sgOrigCallLetters & "', '" & Format$(sgLastAirDate, sgSQLDateForm) & "')"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "Station-mSaveHistory"
            mSaveHistory = False
            Exit Function
        End If
    End If
    For iRow = grdHistory.FixedRows To grdHistory.Rows - 1 Step 1
        grdHistory.Row = iRow
        grdHistory.Col = SHCALLLETTERSINDEX
        sCallLetters = UCase$(Trim$(grdHistory.Text))
        If sCallLetters <> "" Then
            grdHistory.Col = SHLASTDATEINDEX
            If gIsDate(grdHistory.Text) Then
                sEndDate = Format$(grdHistory.Text, sgShowDateForm)
                lCltCode = 0
                SQLQuery = "INSERT INTO clt (cltShfCode, cltCallLetters, cltEndDate) "
                SQLQuery = SQLQuery & " VALUES ( " & imShttCode & ", '" & sCallLetters & "', '" & Format$(sEndDate, sgSQLDateForm) & "')"
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Station-mSaveHistory"
                    mSaveHistory = False
                    Exit Function
                End If
            End If
        End If
    Next iRow
    mSaveHistory = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mSaveHistory"
    mSaveHistory = False
End Function

Private Sub mGetHistory()
    Dim iUpper As Integer
    Dim iRows As Integer
    Dim llRow As Long
    
    On Error GoTo ErrHand
    
    gGrid_Clear grdHistory, True
    llRow = grdHistory.FixedRows
    ReDim tmHistoryInfo(0 To 0) As HISTORYINFO
    SQLQuery = "SELECT cltCallLetters, cltEndDate, cltCode FROM clt WHERE cltShfCode = " & imShttCode & " ORDER BY cltEndDate desc"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        If llRow + 1 > grdHistory.Rows Then
            grdHistory.AddItem ""
        End If
        grdHistory.Row = llRow
        grdHistory.TextMatrix(llRow, SHCALLLETTERSINDEX) = Trim$(rst!cltCallLetters)
        grdHistory.TextMatrix(llRow, SHLASTDATEINDEX) = Format$(rst!cltEndDate, sgShowDateForm)
        grdHistory.TextMatrix(llRow, SHCLTCODEINDEX) = rst!cltCode
        llRow = llRow + 1
        iUpper = UBound(tmHistoryInfo)
        tmHistoryInfo(iUpper).lCode = rst!cltCode
        tmHistoryInfo(iUpper).sCallLetters = rst!cltCallLetters
        tmHistoryInfo(iUpper).sEndDate = Format$(rst!cltEndDate, sgShowDateForm)
        tmHistoryInfo(iUpper).sDelete = "N"
        ReDim Preserve tmHistoryInfo(0 To iUpper + 1) As HISTORYINFO
        rst.MoveNext
    Wend
    If llRow >= grdHistory.Rows Then
        grdHistory.AddItem ""
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-mGetHistory"
End Sub

Private Sub mRemapTimes(slTimeZone As String)


'************************* This code is a relic!  It's place was taken by mZoneChange  ******************************
    '7701-not changed since it isn't used!
    Dim llAtt As Long
    Dim ilRet As Integer
    Dim llOldAttCode As Long
    Dim ilShttCode As Integer
    Dim ilLoop As Integer
    Dim ilOldTimeAdj As Integer
    Dim ilNewTimeAdj As Integer
    Dim ilVef As Integer
    Dim ilZone As Integer
    Dim ilVefCode As Integer
    Dim slEndDate As String
    Dim slSDate As String
    Dim llAttCode As Long
    Dim llFdTime As Long
    Dim llPdTime As Long
    Dim slFdStTime As String
    Dim slFdEdTime As String
    Dim slPdStTime As String
    Dim slPdEdTime As String
    Dim ilLDay As Integer
    Dim ilTDay As Integer
    Dim CurDate As String
    Dim CurTime As String
    Dim slVefName As String
    Dim slDelay As String
    Dim slFileName As String
    Dim llTemp As Long
    Dim slPledgeType As String
    
    ReDim ilFdDay(0 To 6) As Integer
    ReDim ilPdDay(0 To 6) As Integer
    
    '************************
    Exit Sub
    '************************
       
    If Not IsStatDirty Then
        Exit Sub
    End If
    
    ilRet = mOpenMsgFile(slFileName)
    If ilRet = False Then
        Exit Sub
    End If
    On Error GoTo ErrHand
    Print #hmMsg, "  " & smCurCallLetters
    Print #hmMsg, "  " & "Start Date of Zone Change " & Format$(sgTimeZoneChangeDate, sgShowDateForm)
    
    ReDim lmAttCode(0 To 0) As Long
    CurDate = Format(gNow(), sgShowDateForm)
    CurTime = Format(gNow(), sgShowTimeWSecForm)
    slSDate = sgTimeZoneChangeDate
    slEndDate = gAdjYear(Format$(DateValue(slSDate) - 1, sgSQLDateForm))
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery + " FROM att"
    SQLQuery = SQLQuery + " WHERE (attShfCode = " & imShttCode
    SQLQuery = SQLQuery + " AND attOffAir >= '" & Format$(sgTimeZoneChangeDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery + " AND attDropDate >= '" & Format$(sgTimeZoneChangeDate, sgSQLDateForm) & "'" & ")"
    Set attrst = gSQLSelectCall(SQLQuery)
    While Not attrst.EOF
        lmAttCode(UBound(lmAttCode)) = attrst!attCode
        ReDim Preserve lmAttCode(0 To UBound(lmAttCode) + 1) As Long
        attrst.MoveNext
    Wend
    For llAtt = 0 To UBound(lmAttCode) - 1 Step 1
        DoEvents
        SQLQuery = "SELECT *"
        SQLQuery = SQLQuery + " FROM att"
        SQLQuery = SQLQuery + " WHERE (attCode = " & lmAttCode(llAtt) & ")"
        Set attrst = gSQLSelectCall(SQLQuery)
        If Not attrst.EOF Then
            llOldAttCode = attrst!attCode
            ilShttCode = attrst!attshfcode
            ilVefCode = attrst!attvefCode
            slPledgeType = attrst!attPledgeType
            If Trim$(slPledgeType) = "" Then
                If gIsPledgeByAvails(llOldAttCode) Then
                    slPledgeType = "A"
                Else
                    slPledgeType = "D"
                End If
            End If
            'Determine station zone
            For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                If tgStationInfo(ilLoop).iCode = ilShttCode Then
                    ilOldTimeAdj = 0
                    ilNewTimeAdj = 0
                    slDelay = ""
                    slVefName = ""
                    For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                        If tgVehicleInfo(ilVef).iCode = ilVefCode Then
                            slVefName = tgVehicleInfo(ilVef).sVehicle
                            For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                                If StrComp(smCurTimeZone, tgVehicleInfo(ilVef).sZone(ilZone), 1) = 0 Then
                                    ilOldTimeAdj = tgVehicleInfo(ilVef).iVehLocalAdj(ilZone)
                                    Exit For
                                End If
                            Next ilZone
                            Exit For
                        End If
                    Next ilVef
                    For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                        If tgVehicleInfo(ilVef).iCode = ilVefCode Then
                            For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                                If StrComp(slTimeZone, tgVehicleInfo(ilVef).sZone(ilZone), 1) = 0 Then
                                    ilNewTimeAdj = tgVehicleInfo(ilVef).iVehLocalAdj(ilZone)
                                    Exit For
                                End If
                            Next ilZone
                            Exit For
                        End If
                    Next ilVef
                    Exit For
                End If
            Next ilLoop
            'Insert Pledges with remap
            SQLQuery = "SELECT * "
            SQLQuery = SQLQuery + " FROM dat"
            SQLQuery = SQLQuery + " WHERE (datAtfCode= " & llOldAttCode
            SQLQuery = SQLQuery + " AND datShfCode= " & ilShttCode
            SQLQuery = SQLQuery + " AND datVefCode = " & ilVefCode & ")"
            SQLQuery = SQLQuery & " ORDER BY datFdStTime"
            Set DATRST = gSQLSelectCall(SQLQuery)
            DoEvents
            If Not DATRST.EOF Then
                'If DATRST!datDACode = 1 Or DATRST!datDACode = 0 Then
                If (slPledgeType = "D") Or (slPledgeType = "A") Then
                    'Terminate current agreement
                    SQLQuery = "UPDATE att SET "
                    SQLQuery = SQLQuery & "attOffAir = '" & Format$(slEndDate, sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery & "attSentToXDSStatus = '" & "M" & "'"
                    'SQLQuery = SQLQuery & "attDropDate = '" & slEDate & "'"
                    SQLQuery = SQLQuery & " WHERE attCode = " & llOldAttCode
                    cnn.BeginTrans
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "Station-mSaveHistory"
                        Print #hmMsg, ""
                        Print #hmMsg, gMsg
                        Print #hmMsg, "** Zone Change Incomplete: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
                        Print #hmMsg, ""
                        Close #hmMsg
                        gMsgBox "See Status File " & slFileName, vbOKOnly
                        Exit Sub
                    End If
                    'Insert new agreement
                    'D.S. 8/2/05
                    llTemp = gFindAttHole()
                    If llTemp = -1 Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    'SQLQuery = "INSERT INTO att(attCode, attShfCode, attVefCode, attAgreeStart, "
                    'SQLQuery = SQLQuery & "attAgreeEnd, attOnAir, attOffAir, attSigned, attSignDate, "
                    'SQLQuery = SQLQuery & "attLoad, attTimeType, attComp, attBarCode, attDropDate, "
                    'SQLQuery = SQLQuery & "attUsfCode, attEnterDate, attEnterTime, attNotice, "
                    'SQLQuery = SQLQuery & "attCarryCmml, attNoCDs, attSendTape, attACName, "
                    'SQLQuery = SQLQuery & "attACPhone, attGenLog, attGenCP, attPostingType, attPrintCP, "
                    'SQLQuery = SQLQuery & "attComments, attGenOther, attStartTime, attMulticast, attWebInterface, "
                    'SQLQuery = SQLQuery & "attContractPrinted, "
                    'SQLQuery = SQLQuery & "attMktRepUstCode, "
                    'SQLQuery = SQLQuery & "attServRepUstCode, "
                    'SQLQuery = SQLQuery & "attVehProgStartTime, "
                    'SQLQuery = SQLQuery & "attVehProgEndTime, "
                    'SQLQuery = SQLQuery & "attExportToWeb, "
                    'SQLQuery = SQLQuery & "attExportToUnivision, "
                    'SQLQuery = SQLQuery & "attExportToMarketron, "
                    'SQLQuery = SQLQuery & "attExportToCBS, "
                    'SQLQuery = SQLQuery & "attExportToClearCh, "
                    'SQLQuery = SQLQuery & "attUnused "
                    'SQLQuery = SQLQuery & ")"
                    'SQLQuery = SQLQuery & " VALUES"
                    'SQLQuery = SQLQuery & "(" & llTemp & ", " & ilShttCode & ", " & ilVefCode & ", '" & Format$(attrst!attAgreeStart, sgSQLDateForm) & "', '" & Format$(attrst!attAgreeEnd, sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & "'" & Format$(slSDate, sgSQLDateForm) & "', '" & Format$(attrst!attOffAir, sgSQLDateForm) & "', " & attrst!attSigned & ", "
                    'SQLQuery = SQLQuery & "'" & Format$(attrst!attSignDate, sgSQLDateForm) & "', " & attrst!attLoad & ", " & attrst!attTimeType & ", "
                    'SQLQuery = SQLQuery & attrst!attComp & ", " & attrst!attBarCode & ", '" & Format$(attrst!attDropDate, sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & igUstCode & ", '" & Format$(CurDate, sgSQLDateForm) & "', '" & Format$(CurTime, sgSQLTimeForm) & "', '" & attrst!attNotice & "', "
                    'SQLQuery = SQLQuery & attrst!attCarryCmml & ", " & attrst!attNoCDs & ", " & attrst!attSendTape & ", '" & attrst!attACName & "', "
                    'SQLQuery = SQLQuery & "'" & attrst!attACPhone & "', '" & attrst!attGenLog & "', '" & attrst!attGenCP & "', " & attrst!attPostingType & ", " & attrst!attPrintCP & ", "
                    'SQLQuery = SQLQuery & "'" & attrst!attComments & "', '" & attrst!attGenOther & "', '" & Format$(attrst!attStartTime, sgSQLTimeForm) & "', '" & attrst!attMulticast & "', '" & attrst!attWebInterface & "', "
                    'SQLQuery = SQLQuery & "'" & attrst!attContractPrinted & "', "
                    'SQLQuery = SQLQuery & attrst!attMktRepUstCode & ", "
                    'SQLQuery = SQLQuery & attrst!attServRepUstCode & ", "
                    'SQLQuery = SQLQuery & "'" & Format$(attrst!attVehProgStartTime, sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & "'" & Format$(attrst!attVehProgEndTime, sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & "'" & attrst!attExportToWeb & "', "
                    'SQLQuery = SQLQuery & "'" & attrst!attExportToUnivision & "', "
                    'SQLQuery = SQLQuery & "'" & attrst!attExportToMarketron & "', "
                    'SQLQuery = SQLQuery & "'" & attrst!attExportToCBS & "', "
                    'SQLQuery = SQLQuery & "'" & attrst!attExportToClearCh & "', "
                    'SQLQuery = SQLQuery & "'" & "" & "'"
                    'SQLQuery = SQLQuery & ")"
                    
                    SQLQuery = "INSERT INTO att(attCode, attShfCode, attVefCode, attAgreeStart, attAgreeEnd, "
                    SQLQuery = SQLQuery & "attOnAir, attOffAir, attSigned, attSignDate, "
                    SQLQuery = SQLQuery & "attLoad, attTimeType, attComp, attBarCode, attDropDate, "
                    SQLQuery = SQLQuery & "attUsfCode, attEnterDate, attEnterTime, attNotice, "
                    SQLQuery = SQLQuery & "attCarryCmml, attNoCDs, attSendTape, attACName, "
                    SQLQuery = SQLQuery & "attACPhone, attGenLog, attGenCP, attPostingType, attPrintCP, "
                    SQLQuery = SQLQuery & "attExportType, attLogType, attPostType, attWebPW, attWebEmail, "
                    SQLQuery = SQLQuery & "attSendLogEMail, attSuppressNotice, attLabelID, attLabelShipInfo, "
                    SQLQuery = SQLQuery & "attComments, attGenOther, attStartTime, attMulticast, "
                    SQLQuery = SQLQuery & "attRadarClearType, attArttCode, attStatus, attNCR, ,attFormerNCR, attForbidSplitLive, "
                    SQLQuery = SQLQuery & "attXDReceiverID, attVoiceTracked, attWebInterface, "
                    SQLQuery = SQLQuery & "attContractPrinted, "
                    SQLQuery = SQLQuery & "attMktRepUstCode, "
                    SQLQuery = SQLQuery & "attServRepUstCode, "
                    SQLQuery = SQLQuery & "attVehProgStartTime, "
                    SQLQuery = SQLQuery & "attVehProgEndTime, "
                    SQLQuery = SQLQuery & "attExportToWeb, "
                    SQLQuery = SQLQuery & "attExportToUnivision, "
                    SQLQuery = SQLQuery & "attExportToMarketron, "
                    SQLQuery = SQLQuery & "attExportToCBS, "
                    SQLQuery = SQLQuery & "attExportToClearCh, "
                    SQLQuery = SQLQuery & "attPledgeType, "
                    SQLQuery = SQLQuery & "attNoAirPlays, "
                    SQLQuery = SQLQuery & "attDesignVersion, "
                    SQLQuery = SQLQuery & "attIDCReceiverID, "
                    SQLQuery = SQLQuery & "attSentToXDSStatus, "
                    SQLQuery = SQLQuery & "attAudioDelivery, "
                    SQLQuery = SQLQuery & "attExportToJelli, "
                    '3/23/15: Add Send Delays to XDS
                    SQLQuery = SQLQuery & "attSendDelayToXDS, "
                    SQLQuery = SQLQuery & "attServiceAgreement, "
                    '4/3/19
                    SQLQuery = SQLQuery & "attExcludeFillSpot, "
                    SQLQuery = SQLQuery & "attExcludeCntrTypeQ, "
                    SQLQuery = SQLQuery & "attExcludeCntrTypeR, "
                    SQLQuery = SQLQuery & "attExcludeCntrTypeT, "
                    SQLQuery = SQLQuery & "attExcludeCntrTypeM, "
                    SQLQuery = SQLQuery & "attExcludeCntrTypeS, "
                    SQLQuery = SQLQuery & "attExcludeCntrTypeV, "
                    
                    SQLQuery = SQLQuery & "attUnused "
                    SQLQuery = SQLQuery & ")"
                    SQLQuery = SQLQuery & " VALUES"
                    SQLQuery = SQLQuery & "(" & llTemp & ", " & ilShttCode & ", " & imVefCode & ", '" & Format$(attrst!attAgreeStart, sgSQLDateForm) & "', '" & Format$(attrst!attAgreeEnd, sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery & "'" & Format$(slSDate, sgSQLDateForm) & "', '" & Format$(attrst!attOffAir, sgSQLDateForm) & "', " & attrst!attSigned & ", "
                    SQLQuery = SQLQuery & "'" & Format$(attrst!attSignDate, sgSQLDateForm) & "', " & attrst!attLoad & ", " & attrst!attTimeType & ", "
                    SQLQuery = SQLQuery & attrst!attComp & ", " & attrst!attBarCode & ", '" & Format$(attrst!attDropDate, sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery & igUstCode & ", '" & Format$(CurDate, sgSQLDateForm) & "', '" & Format$(CurTime, sgSQLTimeForm) & "', '" & attrst!attNotice & "', "
                    SQLQuery = SQLQuery & attrst!attCarryCmml & ", " & attrst!attNoCDs & ", " & attrst!attSendTape & ", '" & attrst!attACName & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attACPhone & "', '" & attrst!attGenLog & "', '" & attrst!attGenCP & "', " & attrst!attPostingType & ", " & attrst!attPrintCP & ", "
                    SQLQuery = SQLQuery & attrst!attExportType & ", " & attrst!attLogType & ", " & attrst!attPostType & ", '" & attrst!attWebPW & "', '" & attrst!attWebEmail & "', "
                    SQLQuery = SQLQuery & attrst!attSendLogEmail & ", '" & attrst!attSuppressNotice & "', '" & attrst!attLabelID & "', '" & attrst!attLabelShipInfo & "', '" & attrst!attComments & "', '" & attrst!attGenOther & "', '" & Format$(attrst!attStartTime, sgSQLTimeForm) & "', '" & attrst!attMulticast & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attRadarClearType & "', attrst!attArttCode, '" & attrst!attStatus & "', '" & attrst!attNCR & "' , '" & attrst!attFormerNCR & "',  '" & attrst!attForbidSplitLive & "', "
                    SQLQuery = SQLQuery & "attrst!attXDReceiverID, '" & attrst!attVoiceTracked & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attWebInterface & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attContractPrinted & "', "
                    SQLQuery = SQLQuery & attrst!attMktRepUstCode & ", "
                    SQLQuery = SQLQuery & attrst!attServRepUstCode & ", "
                    SQLQuery = SQLQuery & "'" & Format$(attrst!attVehProgStartTime, sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery & "'" & Format$(attrst!attVehProgEndTime, sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExportToWeb & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExportToUnivision & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExportToMarketron & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExportToCBS & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExportToClearCh & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attPledgeType & "', "
                    SQLQuery = SQLQuery & attrst!attNoAirPlays & ", "
                    SQLQuery = SQLQuery & attrst!attDesignVersion & ", "
                    SQLQuery = SQLQuery & "'" & attrst!attIDCReceiverID & "', "
                    SQLQuery = SQLQuery & "'" & "M" & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attAudioDelivery & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExportToJelli & "', "
                    '3/23/15: Add Send Delays to XDS
                    SQLQuery = SQLQuery & "'" & attrst!attSendDelayToXDS & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attServiceAgreement & "', "
                    '4-3-19
                    SQLQuery = SQLQuery & "'" & attrst!attExcludeFillSpot & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeQ & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeR & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeT & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeM & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeS & "', "
                    SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeV & "', "
                    SQLQuery = SQLQuery & "'" & "" & "'"
                    SQLQuery = SQLQuery & ")"
                    
                    
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "Station-mSaveHistory"
                        Print #hmMsg, ""
                        Print #hmMsg, gMsg
                        Print #hmMsg, "** Zone Change Incomplete: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
                        Print #hmMsg, ""
                        Close #hmMsg
                        gMsgBox "See Status File " & slFileName, vbOKOnly
                        Exit Sub
                    End If
                    'SQLQuery = "Select MAX(attCode) from att"
                    'Set rst = gSQLSelectCall(SQLQuery)
                    'llAttCode = rst(0).Value
                    
                    If llTemp = 0 Then
                        SQLQuery = "SELECT MAX(attCode) from att"
                        Set rst = gSQLSelectCall(SQLQuery)
                        llAttCode = rst(0).Value
                    Else
                        llAttCode = llTemp
                    End If
                    
                    While Not DATRST.EOF
                        For ilLDay = 0 To 6 Step 1
                            ilFdDay(ilLDay) = 0
                        Next ilLDay
                        If (DATRST!datFdMon = 1) Then
                            ilFdDay(0) = 1
                        End If
                        If (DATRST!datFdTue = 1) Then
                            ilFdDay(1) = 1
                        End If
                        If (DATRST!datFdWed = 1) Then
                            ilFdDay(2) = 1
                        End If
                        If (DATRST!datFdThu = 1) Then
                            ilFdDay(3) = 1
                        End If
                        If (DATRST!datFdFri = 1) Then
                            ilFdDay(4) = 1
                        End If
                        If (DATRST!datFdSat = 1) Then
                            ilFdDay(5) = 1
                        End If
                        If (DATRST!datFdSun = 1) Then
                            ilFdDay(6) = 1
                        End If
                    
                        llFdTime = gTimeToLong(DATRST!datFdStTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                        If llFdTime < 0 Then
                            llFdTime = llFdTime + 86400
                            ilTDay = ilFdDay(0)
                            For ilLDay = 0 To 5 Step 1
                                ilFdDay(ilLDay) = ilFdDay(ilLDay + 1)
                            Next ilLDay
                            ilFdDay(6) = ilTDay
                        ElseIf llFdTime > 86400 Then
                            llFdTime = llFdTime - 86400
                            ilTDay = ilFdDay(6)
                            For ilLDay = 5 To 0 Step -1
                                ilFdDay(ilLDay + 1) = ilFdDay(ilLDay)
                            Next ilLDay
                            ilFdDay(0) = ilTDay
                        End If
                        slFdStTime = Format$(gLongToTime(llFdTime), "hh:mm:ss")
                        llFdTime = gTimeToLong(DATRST!datFdEdTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                        If llFdTime < 0 Then
                            llFdTime = llFdTime + 86400
                        ElseIf llFdTime > 86400 Then
                            llFdTime = llFdTime - 86400
                        End If
                        slFdEdTime = Format$(gLongToTime(llFdTime), "hh:mm:ss")
                        If DATRST!datFdStatus = 0 Then
                            slPdStTime = slFdStTime
                            slPdEdTime = slFdEdTime
                        Else
                            If ilNewTimeAdj <> ilOldTimeAdj Then
                                slDelay = "Check Delays"
                            End If
                            For ilLDay = 0 To 6 Step 1
                                ilPdDay(ilLDay) = 0
                            Next ilLDay
                            If (DATRST!datPdMon = 1) Then
                                ilPdDay(0) = 1
                            End If
                            If (DATRST!datPdTue = 1) Then
                                ilPdDay(1) = 1
                            End If
                            If (DATRST!datPdWed = 1) Then
                                ilPdDay(2) = 1
                            End If
                            If (DATRST!datPdThu = 1) Then
                                ilPdDay(3) = 1
                            End If
                            If (DATRST!datPdFri = 1) Then
                                ilPdDay(4) = 1
                            End If
                            If (DATRST!datPdSat = 1) Then
                                ilPdDay(5) = 1
                            End If
                            If (DATRST!datPdSun = 1) Then
                                ilPdDay(6) = 1
                            End If
                            llPdTime = gTimeToLong(DATRST!datPdStTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                            If llPdTime < 0 Then
                                llPdTime = llPdTime + 86400
                                ilTDay = ilPdDay(0)
                                For ilLDay = 0 To 5 Step 1
                                    ilPdDay(ilLDay) = ilPdDay(ilLDay + 1)
                                Next ilLDay
                                ilPdDay(6) = ilTDay
                            ElseIf llPdTime > 86400 Then
                                llPdTime = llPdTime - 86400
                                ilTDay = ilPdDay(6)
                                For ilLDay = 5 To 0 Step -1
                                    ilPdDay(ilLDay + 1) = ilPdDay(ilLDay)
                                Next ilLDay
                                ilPdDay(0) = ilTDay
                            End If
                            slPdStTime = Format$(gLongToTime(llPdTime), "hh:mm:ss")
                            llPdTime = gTimeToLong(DATRST!datPdEdTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                            If llPdTime < 0 Then
                                llPdTime = llPdTime + 86400
                            ElseIf llPdTime > 86400 Then
                                llPdTime = llPdTime - 86400
                            End If
                            slPdEdTime = Format$(gLongToTime(llPdTime), "hh:mm:ss")
                        End If
                        If DATRST!datFdStatus = 0 Then
                            'SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, datDACode, "
                            SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, "
                            SQLQuery = SQLQuery & "datFdMon, datFdTue, datFdWed, datFdThu, "
                            SQLQuery = SQLQuery & "datFdFri, datFdSat, datFdSun, datFdStTime, datFdEdTime, datFdStatus, "
                            SQLQuery = SQLQuery & "datPdMon, datPdTue, datPdWed, datPdThu, datPdFri, "
                            SQLQuery = SQLQuery & "datPdSat, datPdSun, datPdDayFed, datPdStTime, datPdEdTime)"
                            SQLQuery = SQLQuery & " VALUES (" & 0 & ", " & llAttCode & ", " & ilShttCode & ", " & ilVefCode
                            SQLQuery = SQLQuery & "," '& DATRST!datDACode & ","
                            SQLQuery = SQLQuery & ilFdDay(0) & ", " & ilFdDay(1) & ","
                            SQLQuery = SQLQuery & ilFdDay(2) & ", " & ilFdDay(3) & ","
                            SQLQuery = SQLQuery & ilFdDay(4) & ", " & ilFdDay(5) & ","
                            SQLQuery = SQLQuery & ilFdDay(6) & ", "
                            SQLQuery = SQLQuery & "'" & Format$(slFdStTime, sgSQLTimeForm) & "','" & Format$(slFdEdTime, sgSQLTimeForm) & "',"
                            SQLQuery = SQLQuery & DATRST!datFdStatus & ","
                            SQLQuery = SQLQuery & ilFdDay(0) & ", " & ilFdDay(1) & ","
                            SQLQuery = SQLQuery & ilFdDay(2) & ", " & ilFdDay(3) & ","
                            SQLQuery = SQLQuery & ilFdDay(4) & ", " & ilFdDay(5) & ","
                            SQLQuery = SQLQuery & ilFdDay(6) & ", "
                            SQLQuery = SQLQuery & "'" & DATRST!datPdDayFed & "', "
                            SQLQuery = SQLQuery & "'" & Format$(slPdStTime, sgSQLTimeForm) & "','" & Format$(slPdEdTime, sgSQLTimeForm) & "')"
                        Else
                            'SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, datDACode, "
                            SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, "
                            SQLQuery = SQLQuery & "datFdMon, datFdTue, datFdWed, datFdThu, "
                            SQLQuery = SQLQuery & "datFdFri, datFdSat, datFdSun, datFdStTime, datFdEdTime, datFdStatus, "
                            SQLQuery = SQLQuery & "datPdMon, datPdTue, datPdWed, datPdThu, datPdFri, "
                            SQLQuery = SQLQuery & "datPdSat, datPdSun, datPdDayFed, datPdStTime, datPdEdTime)"
                            SQLQuery = SQLQuery & " VALUES (" & 0 & ", " & llAttCode & ", " & ilShttCode & ", " & ilVefCode
                            SQLQuery = SQLQuery & "," '& DATRST!datDACode & ","
                            'SQLQuery = SQLQuery & datrst!datFdMon & ", " & datrst!datFdTue & ","
                            'SQLQuery = SQLQuery & datrst!datFdWed & ", " & datrst!datFdThu & ","
                            'SQLQuery = SQLQuery & datrst!datFdFri & ", " & datrst!datFdSat & ","
                            'SQLQuery = SQLQuery & datrst!datFdSun & ", "
                            SQLQuery = SQLQuery & ilFdDay(0) & ", " & ilFdDay(1) & ","
                            SQLQuery = SQLQuery & ilFdDay(2) & ", " & ilFdDay(3) & ","
                            SQLQuery = SQLQuery & ilFdDay(4) & ", " & ilFdDay(5) & ","
                            SQLQuery = SQLQuery & ilFdDay(6) & ", "
                            SQLQuery = SQLQuery & "'" & Format$(slFdStTime, sgSQLTimeForm) & "','" & Format$(slFdEdTime, sgSQLTimeForm) & "',"
                            SQLQuery = SQLQuery & DATRST!datFdStatus & ","
                            SQLQuery = SQLQuery & ilPdDay(0) & ", " & ilPdDay(1) & ","
                            SQLQuery = SQLQuery & ilPdDay(2) & ", " & ilPdDay(3) & ","
                            SQLQuery = SQLQuery & ilPdDay(4) & ", " & ilPdDay(5) & ","
                            SQLQuery = SQLQuery & ilPdDay(6) & ", "
                            SQLQuery = SQLQuery & "'" & DATRST!datPdDayFed & "', "
                            SQLQuery = SQLQuery & "'" & Format$(slPdStTime, sgSQLTimeForm) & "','" & Format$(slPdEdTime, sgSQLTimeForm) & "')"
                        End If
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/12/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "Station-mSaveHistory"
                            Print #hmMsg, ""
                            Print #hmMsg, gMsg
                            Print #hmMsg, "** Zone Change Incomplete: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
                            Print #hmMsg, ""
                            Close #hmMsg
                            gMsgBox "See Status File " & slFileName, vbOKOnly
                            Exit Sub
                        End If
                        DATRST.MoveNext
                    Wend
                    cnn.CommitTrans
                    Print #hmMsg, "     Process Complete For " & slVefName & " " & slDelay
                End If
            End If
        End If
    Next llAtt
    Print #hmMsg, ""
    Print #hmMsg, "** Zone Change Completed: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Close #hmMsg
    gCleanUpAtt
    gMsgBox "See Status File " & slFileName, vbOKOnly
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-mRemapTimes"
    Print #hmMsg, ""
    Print #hmMsg, gMsg
    Print #hmMsg, "** Zone Change Incomplete: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Close #hmMsg
    gMsgBox "See Status File " & slFileName, vbOKOnly
End Sub

Public Sub mAddWebPWToAgrmnt(iUpdateAll As Integer)

    Dim rst_pw As ADODB.Recordset
    Dim slTemp As String

    On Error GoTo ErrHand
    
    If imShttCode = 0 Then
        Exit Sub
    End If
    
    SQLQuery = "SELECT attCode, attShfCode, attWebPW "
    SQLQuery = SQLQuery & " FROM att"
    SQLQuery = SQLQuery + " WHERE (attShfCode = " & imShttCode
    SQLQuery = SQLQuery + " AND attExportType = " & 1 & ")"
    Set rst_pw = gSQLSelectCall(SQLQuery)
    
    While Not rst_pw.EOF
        If iUpdateAll Then
            SQLQuery = "UPDATE att SET attWebPW = '" & Trim$(smShttWebPW) & "'"
            SQLQuery = SQLQuery + " WHERE (attCode= " & rst_pw!attCode
            SQLQuery = SQLQuery + " AND attWebPw = '" & smExistingWebPW & "'" & ")"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "Station-mAddWebPWToAgrmnt"
                Exit Sub
            End If
        End If
        AddattCodeForWebUpdates (rst_pw!attCode)
        rst_pw.MoveNext
    Wend

    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mAddWebPWToAgrmnt"
End Sub
Public Sub mUpdateAgreementXDSSiteID()

    Dim rst_pw As ADODB.Recordset
    Dim slTemp As String

    On Error GoTo ErrHand
    
    If imShttCode = 0 Then
        Exit Sub
    End If
    
    SQLQuery = "SELECT attCode From att where attshfcode = " & imShttCode
    Set rst_pw = gSQLSelectCall(SQLQuery)
    While Not rst_pw.EOF
        SQLQuery = "UPDATE VAT_Vendor_Agreement SET vatSentToWeb = '' WHERE vatWvtVendorId = " & Vendors.XDS_Break & " AND vatattcode = " & rst_pw!attCode
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "Station-mUpdateAgreemenrXDSSiteID"
            Exit Sub
        End If
        rst_pw.MoveNext
    Wend
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "mUpdateAgreementXDSSiteID"
End Sub
Public Sub mAddWebEmailToAgrmnt(iUpdateAll As Integer)

'D.S. 12/11/08  Relic Code

'    Dim rst_email As ADODB.Recordset
'    Dim slTemp As String
'    Dim i As Integer
'    Dim NeedToAddAttCode As Boolean
'
'    On Error GoTo ErrHand
'
'    If imShttCode = 0 Then
'        Exit Sub
'    End If
'
'    SQLQuery = "SELECT attCode, attShfCode, attWebEmail "
'    SQLQuery = SQLQuery & " FROM att"
'    SQLQuery = SQLQuery + " WHERE (attShfCode = " & imShttCode
'    SQLQuery = SQLQuery + " AND attExportType = " & 1 & ")"
'    Set rst_email = gSQLSelectCall(SQLQuery)
'
'    While Not rst_email.EOF
'        If iUpdateAll Then
'            SQLQuery = "UPDATE att SET attWebEmail = '" & Trim$(smShttWebEmail) & "'"
'            SQLQuery = SQLQuery + " WHERE (attCode= " & rst_email!attCode & ")"
'            SQLQuery = SQLQuery + " AND attWebEmail = '" & Trim$(smExistingWebEmail) & "'"
'            'cnn.Execute SQLQuery, rdExecDirect
'            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                GoSub ErrHand:
'            End If
'        End If
'            AddattCodeForWebUpdates (rst_email!attCode)
'        rst_email.MoveNext
'    Wend
'
'    Exit Sub
'

End Sub

Private Sub AddattCodeForWebUpdates(attCode As Long)
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    ' Check to see if we need to add this attCode to the array
    For i = 0 To UBound(lmAttCodesToUpdateWeb) - 1
        If lmAttCodesToUpdateWeb(i) = attCode Then
            Exit Sub
        End If
    Next i
    ReDim Preserve lmAttCodesToUpdateWeb(0 To UBound(lmAttCodesToUpdateWeb) + 1) As Long
    lmAttCodesToUpdateWeb(UBound(lmAttCodesToUpdateWeb) - 1) = attCode
    Exit Sub
    
ErrHand:
    gMsgBox "AddattCodeForWebUpdates, " & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    Exit Sub
End Sub

Private Function mUpdateWebSite() As Boolean
    On Error GoTo ErrHand
    Dim SQLQuery As String
    Dim slStr As String
    Dim slVefName As String
    Dim hmToHeader As Integer
    Dim iRet As Integer
    Dim i As Integer
    Dim RS_Veh As ADODB.Recordset
    Dim cprst As ADODB.Recordset
    Dim FTPAddress As String
    Dim rst_Temp As ADODB.Recordset
    Dim slTemp As String
    Dim ilVefCode As Integer
    'Dim slFTP As String
    Dim slFileName As String
    Dim slTemp1 As String
    Dim llTotalSpotRecords As Long

    slTemp1 = gGetComputerName()
    If slTemp1 = "N/A" Then
        slTemp1 = "Unknown"
    End If
    slTemp1 = slTemp1 & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & ".txt"
    slFileName = "WebHeaders_" & slTemp1

    Call gLoadOption(sgWebServerSection, "FTPAddress", FTPAddress)

    frmProgressMsg.Show
    'frmProgressMsg.SetMessage 0, "Updating Web Site..." & vbCrLf & vbCrLf & "[" & FTPAddress & "]"
    DoEvents
    Screen.MousePointer = vbHourglass
    
    mUpdateWebSite = False

    Call gLoadOption(sgWebServerSection, "WebExports", smWebExports)
    smWebExports = gSetPathEndSlash(smWebExports, True)
    sToFileHeader = smWebExports & slFileName
    'hmToHeader = FreeFile
    'iRet = 0
    'Open sToFileHeader For Output Lock Write As hmToHeader
    iRet = gFileOpen(sToFileHeader, "Output Lock Write", hmToHeader)
    If iRet <> 0 Then
        Screen.MousePointer = vbDefault
        frmProgressMsg.SetMessage 1, "Unable to open file " & sToFileHeader & vbCrLf & "Web site not updated"
        Exit Function
    End If
    
    'Print #hmToHeader, "attCode , NetworkSWProvider, WebsiteProvider, StationProvider, NetworkName, VehicleName, StationName, LogType, PostType, StartTime, StationEmail, StationPW, AggreementEmail, AggreementPW, SendLogEmail, VehicleFTPSite, TimeZone, ShowAvailNames, Multicast, WebLogSummary, WebLogFeedTime"
    Print #hmToHeader, gBuildWebHeaderDetail()

    For i = 0 To UBound(lmAttCodesToUpdateWeb) - 1
        llTotalSpotRecords = gExecWebSQLWithRowsEffected("Select Count(*) from Spots where attCode = " & lmAttCodesToUpdateWeb(i))
        If llTotalSpotRecords = -1 Then
            gLogMsg "ERROR: Station - Unable to obtain Web Spot record count.", "AffErrorLog.Txt", False
            Screen.MousePointer = vbDefault
            Exit Function
        Else
            If llTotalSpotRecords > 0 Then
                SQLQuery = "SELECT shttCallLetters, shttMarket, shttTimeZone, shttCode, attVefCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, attPrintCP, attCode, attTimeType, attGenCP, attStartTime, attLogType, attPostType, shttWebEmail, shttWebPW, attWebEmail, attSendLogEmail, attWebPW, attMulticast, attWebInterface"
        SQLQuery = SQLQuery + " FROM shtt, cptt, att"
        SQLQuery = SQLQuery + " WHERE (ShttCode = " & imShttCode & ""
        SQLQuery = SQLQuery + " AND attCode = " & lmAttCodesToUpdateWeb(i) & ")"
        Set cprst = gSQLSelectCall(SQLQuery)
        If cprst.EOF Then
            Screen.MousePointer = vbDefault
            frmProgressMsg.SetMessage 1, "Unable to find the att record to update" & vbCrLf & "Web site not updated"
            Exit Function
        End If
        
        ilVefCode = cprst!attvefCode
        SQLQuery = "Select vefName From VEF_Vehicles Where vefCode = " & ilVefCode & ""
        Set RS_Veh = gSQLSelectCall(SQLQuery)
        If RS_Veh.EOF Then
            Screen.MousePointer = vbDefault
            frmProgressMsg.SetMessage 1, "Unable to find the vehicle name from att attVehCode." & vbCrLf & "Web site not updated"
            Exit Function
        End If
        slVefName = RS_Veh!vefName
        
        
        smAttWebInterface = Trim$(gGetWebInterface(cprst!attCode))
        slStr = gBuildWebHeaders(cprst, ilVefCode, slVefName, imShttCode, smAttWebInterface, False, "A", "", "", "", "")
        Print #hmToHeader, slStr
            End If
        End If
    Next i
    Close #hmToHeader
    
    If Not gFTPFileToWebServer(sToFileHeader, slFileName) Then
        Screen.MousePointer = vbDefault
        frmProgressMsg.SetMessage 1, "Unable to update the Web Server." & vbCrLf & "Web site not updated"
        Exit Function
    End If
    If Not gSendCmdToWebServer("ImportHeaders.dll", slFileName) Then
        Screen.MousePointer = vbDefault
        frmProgressMsg.SetMessage 1, "FAIL: Unable to instruct Web Server to Import..."
        Exit Function
    End If
    Unload frmProgressMsg
    ReDim lmAttCodesToUpdateWeb(0 To 0)
    imWebPWUpdated = False
    mUpdateWebSite = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-UpdateWebSite"
End Function

Private Sub mHistSetShow()
    If (lmEnableRow >= grdHistory.FixedRows) And (lmEnableRow < grdHistory.Rows) Then
        'Set any field that that should only be set after user leaves the cell
    End If
    imShowGridBox = False
    pbcArrow.Visible = False
    txtHistory.Visible = False
End Sub

Private Sub mHistEnableBox()
    If (grdHistory.Row >= grdHistory.FixedRows) And (grdHistory.Row < grdHistory.Rows) And (grdHistory.Col >= 0) And (grdHistory.Col < grdHistory.Cols - 1) Then
        lmEnableRow = grdHistory.Row
        imShowGridBox = True
        pbcArrow.Move grdHistory.Left - pbcArrow.Width, grdHistory.Top + grdHistory.RowPos(grdHistory.Row) + (grdHistory.RowHeight(grdHistory.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        Select Case grdHistory.Col
            Case SHCALLLETTERSINDEX  'Call Letters
                txtHistory.Move grdHistory.Left + grdHistory.ColPos(grdHistory.Col) + 30, grdHistory.Top + grdHistory.RowPos(grdHistory.Row) + 15, grdHistory.ColWidth(grdHistory.Col) - 30, grdHistory.RowHeight(grdHistory.Row) - 15
                txtHistory.Text = grdHistory.Text
                If txtHistory.Height > grdHistory.RowHeight(grdHistory.Row) - 15 Then
                    txtHistory.FontName = "Arial"
                    txtHistory.Height = grdHistory.RowHeight(grdHistory.Row) - 15
                End If
                txtHistory.Visible = True
                txtHistory.SetFocus
            Case SHLASTDATEINDEX  'Date
                txtHistory.Move grdHistory.Left + grdHistory.ColPos(grdHistory.Col) + 30, grdHistory.Top + grdHistory.RowPos(grdHistory.Row) + 15, grdHistory.ColWidth(grdHistory.Col) - 30, grdHistory.RowHeight(grdHistory.Row) - 15
                txtHistory.Text = grdHistory.Text
                If txtHistory.Height > grdHistory.RowHeight(grdHistory.Row) - 15 Then
                    txtHistory.FontName = "Arial"
                    txtHistory.Height = grdHistory.RowHeight(grdHistory.Row) - 15
                End If
                txtHistory.Visible = True
                txtHistory.SetFocus
        End Select
    End If
End Sub

Private Function mAsk() As Integer

    Dim att_rst As ADODB.Recordset
    Dim dat_rst As ADODB.Recordset

    'D.S. 4/04 Test is the station has any agreements.  If it does test to see if it has any
    'agreements that are not CD/Tape daypart agreements.  If daypart or avails agreemnts are found
    'then return True which forces the question to Time Remap or not
    'datDACode = 0 => Daypart
    '            1 => Avails
    '            2 => CD/Tape
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    mAsk = False
    'Find out if there are any agreements
    SQLQuery = "SELECT attCode, attPledgeType FROM att WHERE attShfCode = " & imShttCode
    Set att_rst = gSQLSelectCall(SQLQuery)
    While Not att_rst.EOF
        'SQLQuery = "SELECT datDaCode FROM dat WHERE datAtfCode = " & att_rst!attCode
        'Set dat_rst = gSQLSelectCall(SQLQuery)
        'While Not dat_rst.EOF
            'If dat_rst!datDACode <> 2 Then
            If att_rst!attPledgeType <> "C" Then
                mAsk = True
                Exit Function
            End If
         '   dat_rst.MoveNext
        'Wend
        att_rst.MoveNext
    Wend
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mAsk"
End Function

Private Function mOwnerFillListBox(iSelect As Integer, iSetFocus As Integer, llArttCode As Long) As Integer
    '5/22/07:  Pass ilArttCode
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim temp_rst As ADODB.Recordset
    Dim slName As String
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    bmIgnoreOwnerChange = True
    
    mOwnerFillListBox = False
    cbcOwner.Clear
    
    
    cbcOwner.AddItem "[New]"
    cbcOwner.SetItemData = -1
    cbcOwner.AddItem "[None]"
    cbcOwner.SetItemData = 0
    
    SQLQuery = "SELECT * FROM artt where arttType = " & "'O'" & " ORDER BY arttLastName"
    Set temp_rst = gSQLSelectCall(SQLQuery)
    
    While Not temp_rst.EOF
        cbcOwner.AddItem Trim$(temp_rst!arttLastName)
        cbcOwner.SetItemData = temp_rst!arttCode
        temp_rst.MoveNext
    Wend
    
    If iSelect = 0 Then
        'SQLQuery = "SELECT shttOwnerArttCode FROM shtt where shttCode = " & imShttCode 'CInt(cboStations.ItemData(cboStations.ListIndex))
        'Set temp_rst = gSQLSelectCall(SQLQuery)
        'If Not temp_rst.EOF Then
        '    For ilIdx = 0 To cboOwner.ListCount - 1
        '        If cboOwner.ItemData(ilIdx) = CInt(temp_rst!shttOwnerArttCode) Then
        '            cboOwner.ListIndex = ilIdx
        '            lmArttCode = CInt(temp_rst!shttOwnerArttCode)
        '            Exit For
        '        End If
        '    Next ilIdx
        'End If
        If llArttCode > 0 Then
            For ilIdx = 0 To cbcOwner.ListCount - 1 Step 1
                If cbcOwner.GetItemData(ilIdx) = llArttCode Then
                    cbcOwner.SetListIndex = ilIdx
                    Exit For
                End If
            Next ilIdx
        End If
    End If
    If iSelect = 1 Then
        cbcOwner.SetListIndex = igIndex + 1
    End If
    
    'slName = cboOwner.Text
    'ilRow = SendMessageByString(cboOwner.hwnd, CB_FINDSTRING, -1, slName)
    
    'Doug- On 11/17/06 added iSetFocus and test for which tab is visible.  I did this because of call added in mBindControls
    If iSetFocus And tscStation.Tabs(1).Selected Then
        cbcOwner.SetFocus
    End If
    Screen.MousePointer = vbDefault
    mOwnerFillListBox = True
    bmIgnoreOwnerChange = False
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mOwnerFillListBox"
End Function

Private Function mDMAMarketFillListBox(iSelect As Integer, iSetFocus As Integer, ilMktCode As Integer) As Integer
    '5/22/07: Pass Market code
    'D.S. 11/11/05
    
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim temp_rst As ADODB.Recordset
    Dim slName As String
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    

    mDMAMarketFillListBox = False
    cbcDMAMarket.Clear
    
    cbcDMAMarket.AddItem "[New]"
    cbcDMAMarket.SetItemData = -1
    cbcDMAMarket.AddItem "[None]"
    cbcDMAMarket.SetItemData = 0
    
    SQLQuery = "SELECT * FROM mkt ORDER BY mktName"
    Set temp_rst = gSQLSelectCall(SQLQuery)
    
    While Not temp_rst.EOF
        cbcDMAMarket.AddItem Trim$(temp_rst!mktName)
        cbcDMAMarket.SetItemData = temp_rst!mktCode
        temp_rst.MoveNext
    Wend
    
    If iSelect = 0 Then
        'SQLQuery = "SELECT shttMktCode FROM shtt where shttCode = " & imShttCode 'CInt(cboStations.ItemData(cboStations.ListIndex))
        'Set temp_rst = gSQLSelectCall(SQLQuery)
        'If Not temp_rst.EOF Then
        '    For ilIdx = 0 To cboDMAMarketCluster.ListCount - 1
        '        If cboDMAMarketCluster.ItemData(ilIdx) = CInt(temp_rst!shttMktCode) Then
        '            cboDMAMarketCluster.ListIndex = ilIdx
        '            imDMAMktCode = CInt(temp_rst!shttMktCode)
        '            Exit For
        '        End If
        '    Next ilIdx
        'End If
        If ilMktCode > 0 Then
            For ilIdx = 0 To cbcDMAMarket.ListCount - 1 Step 1
                If cbcDMAMarket.GetItemData(ilIdx) = ilMktCode Then
                    cbcDMAMarket.SetListIndex = ilIdx
                    Exit For
                End If
            Next ilIdx
        End If
    End If
    If iSelect = 1 Then
        cbcDMAMarket.SetListIndex = igIndex + 1
    End If

    'slName = cboDMAMarketCluster.Text
    'ilRow = SendMessageByString(cboDMAMarketCluster.hwnd, CB_FINDSTRING, -1, slName)
    
    'Doug- On 11/17/06 added iSetFocus and test for which tab is visible.  I did this because of call added in mBindControls
    If iSetFocus And tscStation.Tabs(1).Selected Then
        cbcDMAMarket.SetFocus
    End If
    Screen.MousePointer = vbDefault
    mDMAMarketFillListBox = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mDMAMarketFillListBox"
End Function



Private Function mMSAMarketFillListBox(iSelect As Integer, iSetFocus As Integer, ilMetCode As Integer) As Integer
    '5/22/07: Pass Market code
    'D.S. 11/11/05
    
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim temp_rst As ADODB.Recordset
    Dim slName As String
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    

    mMSAMarketFillListBox = False
    cbcMSAMarket.Clear
    
    
    cbcMSAMarket.AddItem "[New]"
    cbcMSAMarket.SetItemData = -1
    cbcMSAMarket.AddItem "[None]"
    cbcMSAMarket.SetItemData = 0
    
    SQLQuery = "SELECT * FROM met ORDER BY metName"
    Set temp_rst = gSQLSelectCall(SQLQuery)
    
    While Not temp_rst.EOF
        cbcMSAMarket.AddItem Trim$(temp_rst!metName)
        cbcMSAMarket.SetItemData = temp_rst!metCode
        temp_rst.MoveNext
    Wend
    
    If iSelect = 0 Then
        'SQLQuery = "SELECT shttmetCode FROM shtt where shttCode = " & imShttCode 'CInt(cboStations.ItemData(cboStations.ListIndex))
        'Set temp_rst = gSQLSelectCall(SQLQuery)
        'If Not temp_rst.EOF Then
        '    For ilIdx = 0 To cboMSAMarketCluster.ListCount - 1
        '        If cboMSAMarketCluster.ItemData(ilIdx) = CInt(temp_rst!shttmetCode) Then
        '            cboMSAMarketCluster.ListIndex = ilIdx
        '            immetCode = CInt(temp_rst!shttmetCode)
        '            Exit For
        '        End If
        '    Next ilIdx
        'End If
        If ilMetCode > 0 Then
            For ilIdx = 0 To cbcMSAMarket.ListCount - 1 Step 1
                If cbcMSAMarket.GetItemData(ilIdx) = ilMetCode Then
                    cbcMSAMarket.SetListIndex = ilIdx
                    Exit For
                End If
            Next ilIdx
        End If
    End If
    If iSelect = 1 Then
        cbcMSAMarket.SetListIndex = igIndex + 1
    End If

    'slName = cboMSAMarketCluster.Text
    'ilRow = SendMessageByString(cboMSAMarketCluster.hwnd, CB_FINDSTRING, -1, slName)
    
    'Doug- On 11/17/06 added iSetFocus and test for which tab is visible.  I did this because of call added in mBindControls
    If iSetFocus And tscStation.Tabs(1).Selected Then
        cbcMSAMarket.SetFocus
    End If
    Screen.MousePointer = vbDefault
    mMSAMarketFillListBox = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mMSAMarketFillListBox"
End Function



'**************************************************************************
' Personnel grid stuff
'
'**************************************************************************

Private Function mOwnerMarketSave()

    mOwnerMarketSave = True
    'gMsgBox "Owner/Market Save Entered"

End Function


Private Function mZoneChange(sNewTimeZone As String) As Integer

    'D.S. 1/7/07
    'This function replaces the mRemapTimes function
    'Purpose: When the time zone is changed for a station it needs to update the AST, DAT and the web site with the new
    'time and date values for all of the agreements that are for that station and are current.
    
    Dim ilRet As Integer
    Dim ilDay As Integer
    Dim ilLDay As Integer
    Dim ilTDay As Integer
    Dim ilZone As Integer
    Dim ilOldTimeAdj As Integer
    Dim ilNewTimeAdj As Integer
    Dim ilFinalTimeAdj As Integer
    Dim llVefCode As Long
    Dim llAtt As Long
    Dim llFdTime As Long
    Dim llPdTime As Long
    Dim slFileName As String
    Dim slCurDtTime As String
    Dim slCurDate As String
    Dim slCurTime As String
    Dim slVefName As String
    Dim slFdStTime As String
    Dim slFdEdTime As String
    Dim tmpStr As String
    
    Dim slFdTime As String
    Dim slFdDate As String
    Dim slPdStTime As String
    Dim slPdEdTime As String
    Dim slPdDate As String
    Dim slAirTime As String
    Dim slAirDate As String
    Dim slLastPostedDate As String
    Dim slDelay As String
    Dim slToFile As String
    
    Dim attrst As ADODB.Recordset
    Dim astrst As ADODB.Recordset
    Dim DATRST As ADODB.Recordset

    ReDim ilFdDay(0 To 6) As Integer
    ReDim ilPdDay(0 To 6) As Integer
    '7701
    Dim slAttExportToUnivision As String
    Dim slattExportToMarketron As String
    Dim slattExportToCBS As String
    Dim slattExportToClearCh As String
    'problem with date/time format Dan 10/15/15
    Dim slDate As String
    Dim slTime As String
    Dim slSQLQuery As String
    Dim llLoop As Long
    Dim slCDStartTime As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llSpotCount As Long
    Dim llRowsEffected As Long

    On Error GoTo ErrHand
    
    mZoneChange = False
    ilRet = gMsgBox("Zone changed from " & smCurTimeZone & " to " & sNewTimeZone & "." & Chr(13) & Chr(10) & "This will cause agreements defined as Avail Posting to be altered." & Chr(13) & Chr(10) & "Agreements defined as Daypart and CD/Tape will be left unchanged." & Chr(13) & Chr(10) & "Proceed with change?", vbYesNo)
    If ilRet = vbNo Then
        mZoneChange = True
        Exit Function
    End If
    
    If Not IsStatDirty Then
        Exit Function
    End If
    '10/3/18: Dan- In the routine gVatSetToGoToWebByShttCode I will ignore the VendorID.  Dan no longer ignored
    'TTP 8824 reopened
    ''7941 time zone change? Update web on next export
    gVatSetToGoToWebByShttCode imShttCode, Vendors.XDS_Break
    slToFile = "ZC" & Format$(gNow(), "mm") & Format$(gNow(), "dd") & Format$(gNow(), "yy") & ".txt"
    gLogMsg smCurCallLetters, slToFile, False
    gLogMsg "Zone Change on " & Format$(gNow, sgShowDateForm), slToFile, False
    
    ReDim lmAttCode(0 To 0) As Long
    
    slCurDtTime = Format(Now(), "ddddd ttttt")
    slCurDate = Format(slCurDtTime, sgShowDateForm)
    slCurTime = Format(slCurDtTime, "hh:mm:ss")
    
    ilDay = Weekday(slCurDate, vbMonday) - 1
    
    'D.S. Gather the Avail agreements that are current as of today. Do not get the CD/Tape or Daypart agreements
    slSQLQuery = "SELECT *"
    slSQLQuery = slSQLQuery + " FROM att"
    slSQLQuery = slSQLQuery + " WHERE (attShfCode = " & imShttCode
    slSQLQuery = slSQLQuery + " AND attOffAir >= '" & Format$(slCurDtTime, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery + " AND attDropDate >= '" & Format$(slCurDtTime, sgSQLDateForm) & "'" & ")"
    slSQLQuery = slSQLQuery + " AND (attTimeType = 1 " & ")"
    
    Set attrst = gSQLSelectCall(slSQLQuery)
    'D.S. Build an array of the attcodes we got from above call
    While Not attrst.EOF
        lmAttCode(UBound(lmAttCode)) = attrst!attCode
        ReDim Preserve lmAttCode(0 To UBound(lmAttCode) + 1) As Long
        attrst.MoveNext
    Wend

    For llAtt = 0 To UBound(lmAttCode) - 1 Step 1
        DoEvents
        slSQLQuery = "SELECT *"
        slSQLQuery = slSQLQuery + " FROM att "
        slSQLQuery = slSQLQuery + " WHERE (attCode = " & lmAttCode(llAtt) & ")"
        Set attrst = gSQLSelectCall(slSQLQuery)
        If Not attrst.EOF Then
            ilOldTimeAdj = 0
            ilNewTimeAdj = 0
            llVefCode = gBinarySearchVef(CLng(attrst!attvefCode))
            
            'Get the Zone offsets
            If llVefCode <> -1 Then
                slVefName = tgVehicleInfo(llVefCode).sVehicle
                For ilZone = LBound(tgVehicleInfo(llVefCode).sZone) To UBound(tgVehicleInfo(llVefCode).sZone) Step 1
                    If StrComp(smCurTimeZone, tgVehicleInfo(llVefCode).sZone(ilZone), 1) = 0 Then
                        ilOldTimeAdj = tgVehicleInfo(llVefCode).iVehLocalAdj(ilZone)
                        Exit For
                    End If
                Next ilZone
                
                For ilZone = LBound(tgVehicleInfo(llVefCode).sZone) To UBound(tgVehicleInfo(llVefCode).sZone) Step 1
                    If StrComp(sNewTimeZone, tgVehicleInfo(llVefCode).sZone(ilZone), 1) = 0 Then
                        ilNewTimeAdj = tgVehicleInfo(llVefCode).iVehLocalAdj(ilZone)
                        Exit For
                    End If
                Next ilZone
            End If
           
            ilFinalTimeAdj = ilNewTimeAdj - ilOldTimeAdj
           
        '********************************* Update DAT *********************************
           
            slSQLQuery = "SELECT * "
            slSQLQuery = slSQLQuery + " FROM dat"
            slSQLQuery = slSQLQuery + " WHERE datAtfCode= " & lmAttCode(llAtt)
            slSQLQuery = slSQLQuery & " ORDER BY datFdStTime"
            Set DATRST = gSQLSelectCall(slSQLQuery)
            DoEvents
            While Not DATRST.EOF
                For ilLDay = 0 To 6 Step 1
                    ilFdDay(ilLDay) = 0
                Next ilLDay
                If (DATRST!datFdMon = 1) Then
                    ilFdDay(0) = 1
                End If
                If (DATRST!datFdTue = 1) Then
                    ilFdDay(1) = 1
                End If
                If (DATRST!datFdWed = 1) Then
                    ilFdDay(2) = 1
                End If
                If (DATRST!datFdThu = 1) Then
                    ilFdDay(3) = 1
                End If
                If (DATRST!datFdFri = 1) Then
                    ilFdDay(4) = 1
                End If
                If (DATRST!datFdSat = 1) Then
                    ilFdDay(5) = 1
                End If
                If (DATRST!datFdSun = 1) Then
                    ilFdDay(6) = 1
                End If
            
                llFdTime = gTimeToLong(DATRST!datFdStTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                If llFdTime < 0 Then
                    llFdTime = llFdTime + 86400
                    ilTDay = ilFdDay(0)
                    For ilLDay = 0 To 5 Step 1
                        ilFdDay(ilLDay) = ilFdDay(ilLDay + 1)
                    Next ilLDay
                    ilFdDay(6) = ilTDay
                ElseIf llFdTime > 86400 Then
                    llFdTime = llFdTime - 86400
                    ilTDay = ilFdDay(6)
                    For ilLDay = 5 To 0 Step -1
                        ilFdDay(ilLDay + 1) = ilFdDay(ilLDay)
                    Next ilLDay
                    ilFdDay(0) = ilTDay
                End If
                slFdStTime = Format$(gLongToTime(llFdTime), "hh:mm:ss")
                llFdTime = gTimeToLong(DATRST!datFdEdTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                If llFdTime < 0 Then
                    llFdTime = llFdTime + 86400
                ElseIf llFdTime > 86400 Then
                    llFdTime = llFdTime - 86400
                End If
                slFdEdTime = Format$(gLongToTime(llFdTime), "hh:mm:ss")
                If DATRST!datFdStatus = 0 Then
                    slPdStTime = slFdStTime
                    slPdEdTime = slFdEdTime
                'Else
                    If ilNewTimeAdj <> ilOldTimeAdj Then
                        slDelay = "Check Delays"
                    End If
                    For ilLDay = 0 To 6 Step 1
                        ilPdDay(ilLDay) = 0
                    Next ilLDay
                    If (DATRST!datPdMon = 1) Then
                        ilPdDay(0) = 1
                    End If
                    If (DATRST!datPdTue = 1) Then
                        ilPdDay(1) = 1
                    End If
                    If (DATRST!datPdWed = 1) Then
                        ilPdDay(2) = 1
                    End If
                    If (DATRST!datPdThu = 1) Then
                        ilPdDay(3) = 1
                    End If
                    If (DATRST!datPdFri = 1) Then
                        ilPdDay(4) = 1
                    End If
                    If (DATRST!datPdSat = 1) Then
                        ilPdDay(5) = 1
                    End If
                    If (DATRST!datPdSun = 1) Then
                        ilPdDay(6) = 1
                    End If
                    llPdTime = gTimeToLong(DATRST!datPdStTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                    If llPdTime < 0 Then
                        llPdTime = llPdTime + 86400
                        ilTDay = ilPdDay(0)
                        For ilLDay = 0 To 5 Step 1
                            ilPdDay(ilLDay) = ilPdDay(ilLDay + 1)
                        Next ilLDay
                        ilPdDay(6) = ilTDay
                    ElseIf llPdTime > 86400 Then
                        llPdTime = llPdTime - 86400
                        ilTDay = ilPdDay(6)
                        For ilLDay = 5 To 0 Step -1
                            ilPdDay(ilLDay + 1) = ilPdDay(ilLDay)
                        Next ilLDay
                        ilPdDay(0) = ilTDay
                    End If
                    slPdStTime = Format$(gLongToTime(llPdTime), "hh:mm:ss")
                    llPdTime = gTimeToLong(DATRST!datPdEdTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                    If llPdTime < 0 Then
                        llPdTime = llPdTime + 86400
                    ElseIf llPdTime > 86400 Then
                        llPdTime = llPdTime - 86400
                    End If
                    slPdEdTime = Format$(gLongToTime(llPdTime), "hh:mm:ss")
                End If
                
                slSQLQuery = "UPDATE dat"
                slSQLQuery = slSQLQuery & " SET datFdMon = " & ilFdDay(0) & ","
                slSQLQuery = slSQLQuery & "datFdTue = " & ilFdDay(1) & ","
                slSQLQuery = slSQLQuery & "datFdWed = " & ilFdDay(2) & ","
                slSQLQuery = slSQLQuery & "datFdThu = " & ilFdDay(3) & ","
                slSQLQuery = slSQLQuery & "datFdFri = " & ilFdDay(4) & ","
                slSQLQuery = slSQLQuery & "datFdSat = " & ilFdDay(5) & ","
                slSQLQuery = slSQLQuery & "datFdSun = " & ilFdDay(6) & ","
                
                slSQLQuery = slSQLQuery & "datPdMon = " & ilPdDay(0) & ","
                slSQLQuery = slSQLQuery & "datPdTue = " & ilPdDay(1) & ","
                slSQLQuery = slSQLQuery & "datPdWed = " & ilPdDay(2) & ","
                slSQLQuery = slSQLQuery & "datPdThu = " & ilPdDay(3) & ","
                slSQLQuery = slSQLQuery & "datPdFri = " & ilPdDay(4) & ","
                slSQLQuery = slSQLQuery & "datPdSat = " & ilPdDay(5) & ","
                slSQLQuery = slSQLQuery & "datPdSun = " & ilPdDay(6) & ","
                
                If bmAdjPledge Then
                    slSQLQuery = slSQLQuery & "datFdStTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datFdStTime), sgSQLTimeForm) & "', "
                    slSQLQuery = slSQLQuery & "datFdEdTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datFdEdTime), sgSQLTimeForm) & "', "
                    slSQLQuery = slSQLQuery & "datPdStTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datPdStTime), sgSQLTimeForm) & "', "
                    slSQLQuery = slSQLQuery & "datPdEdTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datPdEdTime), sgSQLTimeForm) & "' "
                Else
                    slSQLQuery = slSQLQuery & "datFdStTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datFdStTime), sgSQLTimeForm) & "', "
                    slSQLQuery = slSQLQuery & "datFdEdTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datFdEdTime), sgSQLTimeForm) & "' "
                End If
                slSQLQuery = slSQLQuery & " WHERE (datCode = " & DATRST!datCode & ")"
                'cnn.Execute slSQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Station-mZoneChange"
                    mZoneChange = False
                    Exit Function
                End If
                DATRST.MoveNext
            Wend
            
            '********************************* Update AST *********************************
            '7701
            slAttExportToUnivision = ""
            slattExportToMarketron = ""
            slattExportToCBS = ""
            slattExportToClearCh = ""
            If gIsVendorWithAgreement(lmAttCode(llAtt), Vendors.cBs) Then
                slattExportToCBS = "Y"
            End If
            If gIsVendorWithAgreement(lmAttCode(llAtt), Vendors.iheart) Then
                slattExportToClearCh = "Y"
            End If
            If gIsVendorWithAgreement(lmAttCode(llAtt), Vendors.NetworkConnect) Then
                slattExportToMarketron = "Y"
            End If
            '7701
            slLastPostedDate = gGetLastPostedDate(lmAttCode(llAtt), attrst!attExportType, attrst!attExportToWeb, slAttExportToUnivision, slattExportToMarketron, slattExportToCBS, slattExportToClearCh)
            Screen.MousePointer = vbHourglass
            'slLastPostedDate = gGetLastPostedDate(lmAttCode(llAtt), attrst!attExportType, attrst!attExportToWeb, attrst!attExportToUnivision, attrst!attExportToMarketron, attrst!attExportToCBS, attrst!attExportToClearCh)
           
            slSQLQuery = "Select * FROM ast WHERE "
            slSQLQuery = slSQLQuery + " astAtfCode = " & lmAttCode(llAtt)
            Set astrst = gSQLSelectCall(slSQLQuery)
            While Not astrst.EOF
                                 
                'Feed Date and Time
                'Dan 10/15/15 issue with format
                slDate = Format(astrst!astFeedDate, sgShowDateForm)
                slTime = Format(astrst!astFeedTime, sgShowTimeWSecForm)
                tmpStr = slDate & " " & slTime
                'tmpStr = Trim$(astrst!astFeedDate) & " " & Trim$(astrst!astFeedTime)
                tmpStr = DateAdd("h", ilFinalTimeAdj, tmpStr)
                slFdDate = Format(tmpStr, sgSQLDateForm)
                slFdTime = Format(tmpStr, sgSQLTimeForm)
                
                '12/13/13: Pledge information removed (DAT used instead)
                'Pledge Date & Start and End Time
                'tmpStr = Trim$(astrst!astPledgeDate) & " " & Trim$(astrst!astPledgeStartTime)
                'tmpStr = DateAdd("h", ilFinalTimeAdj, tmpStr)
                'slPdDate = Format(tmpStr, sgSQLDateForm)
                'slPdStTime = Format(tmpStr, sgSQLTimeForm)
                'tmpStr = Trim$(astrst!astPledgeDate) & " " & Trim$(astrst!astPledgeEndTime)
                'tmpStr = DateAdd("h", ilFinalTimeAdj, tmpStr)
                'slPdEdTime = Format(tmpStr, sgSQLTimeForm)
                
                'Air Date & Time
                'Dan 10/15/15 issue with format
                slDate = Format(astrst!astAirDate, sgShowDateForm)
                slTime = Format(astrst!astAirTime, sgShowTimeWSecForm)
                tmpStr = slDate & " " & slTime
               ' tmpStr = Trim$(astrst!astAirDate) & " " & Trim$(astrst!astAirTime)
                If bmAdjPledge Then
                    tmpStr = DateAdd("h", ilFinalTimeAdj, tmpStr)
                Else
                    tmpStr = DateAdd("h", 0, tmpStr)
                End If
                slAirDate = Format(tmpStr, sgSQLDateForm)
                slAirTime = Format(tmpStr, sgSQLTimeForm)
                
                'see if the feed changed due to the zone changed
                '0 = no date changed, 1 = date moved forward a day, -1 date moved back a day
                'dan m 10/15/15 more time format issues
                slTime = Format(astrst!astFeedTime, sgShowTimeWSecForm)
                ilRet = mZoneChangesDate(ilFinalTimeAdj, slTime)
               ' ilRet = mZoneChangesDate(ilFinalTimeAdj, astrst!astFeedTime)
                If ilRet = 0 Then
                    'No feed date change so we can update the record
                    slSQLQuery = "UPDATE ast SET"
                    slSQLQuery = slSQLQuery & " astFeedTime = " & "'" & slFdTime & "', "
                    slSQLQuery = slSQLQuery & " astFeedDate = " & "'" & slFdDate & "', "
                    
                    If astrst!astCPStatus <> 1 Then
                        'If the ast has been posted then don't update the airdate or airtime
                        slSQLQuery = slSQLQuery & " astAirTime = " & "'" & slAirTime & "', "
                        slSQLQuery = slSQLQuery & " astAirDate = " & "'" & slAirDate & "', "
                    End If
                    
                    '12/13/13: Pledge information removed (DAT used instead)
                    'slSQLQuery = slSQLQuery & " astPledgeStartTime = " & "'" & slPdStTime & "', "
                    'slSQLQuery = slSQLQuery & " astPledgeEndTime = " & "'" & slPdEdTime & "', "
                    'slSQLQuery = slSQLQuery & " astPledgeDate = " & "'" & slPdDate & "'"
                    slSQLQuery = slSQLQuery & " astUstCode = " & igUstCode
                    slSQLQuery = slSQLQuery & " WHERE astCode = " & astrst!astCode
                    'cnn.Execute slSQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "Station-mZoneChange"
                        mZoneChange = False
                        Exit Function
                    End If
                Else
                    'Feed date changed so, we must delete the old record and insert a now one with the same astCode
                    slSQLQuery = "DELETE FROM Ast WHERE astCode = " & astrst!astCode
                    'cnn.Execute slSQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "Station-mZoneChange"
                        mZoneChange = False
                        Exit Function
                    End If

                    slSQLQuery = "INSERT INTO ast"
                    slSQLQuery = slSQLQuery + "(astCode, astAtfCode, astShfCode, astVefCode, "
                    slSQLQuery = slSQLQuery + "astSdfCode, astLsfCode, astAirDate, astAirTime, "
                    '12/13/13: Support New AST layout
                    'slSQLQuery = slSQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, "
                    'slSQLQuery = slSQLQuery + "astPledgeStartTime, astPledgeEndTime, astPledgeStatus)"
                    slSQLQuery = slSQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, "
                    slSQLQuery = slSQLQuery + "astAdfCode, astDatCode, astCpfCode, astRsfCode, astStationCompliant, astAgencyCompliant, astAffidavitSource, astCntrNo, astLen, astLkAstCode, astMissedMnfCode, astUstCode)"
                    slSQLQuery = slSQLQuery + " VALUES "
                    slSQLQuery = slSQLQuery + "(" & astrst!astCode & ", " & astrst!astAtfCode & ", " & astrst!astShfCode & ", "
                    slSQLQuery = slSQLQuery & astrst!astVefCode & ", " & astrst!astSdfCode & ", " & astrst!astLsfCode & ", "
                    
                    If astrst!astCPStatus <> 1 Then
                        'If the ast has been posted then don't update the airdate or airtime
                        slSQLQuery = slSQLQuery + "'" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "', "
                    Else
                        'Use the old posted date and time
                        slSQLQuery = slSQLQuery + "'" & Format$(astrst!astAirDate, sgSQLDateForm) & "', '" & Format$(astrst!astAirTime, sgSQLTimeForm) & "', "
                    End If
                    
                    slSQLQuery = slSQLQuery & astrst!astStatus & ", " & astrst!astCPStatus & ", '" & Format$(slFdDate, sgSQLDateForm) & "', "
                    '12/13/13: Support New AST layout
                    'slSQLQuery = slSQLQuery & "'" & Format$(slFdTime, sgSQLTimeForm) & "', '" & Format$(slPdDate, sgSQLDateForm) & "', "
                    'slSQLQuery = slSQLQuery & "'" & Format$(slPdStTime, sgSQLTimeForm) & "', '" & Format$(slPdEdTime, sgSQLTimeForm) & "', " & astrst!astPledgeStatus & ")"
                    slSQLQuery = slSQLQuery & "'" & Format$(slFdTime, sgSQLTimeForm) & "', "
                    slSQLQuery = slSQLQuery & astrst!astAdfCode & ", " & astrst!astDatCode & ", " & astrst!astCpfCode & ", " & astrst!astRsfCode & ", "
                    slSQLQuery = slSQLQuery & "'" & astrst!astStationCompliant & "', '" & astrst!astAgencyCompliant & "', '" & gRemoveIllegalChars(astrst!astAffidavitSource) & "', " & astrst!astCntrNo & ", " & astrst!astLen & ", " & astrst!astLkAstCode & ", " & astrst!astMissedMnfCode & ", " & igUstCode & ")"
                    
                    'cnn.Execute slSQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "Station-mZoneChange"
                        mZoneChange = False
                        Exit Function
                    End If
                End If

                astrst.MoveNext
            Wend
            
            '4/2/16: Update program times
            'For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            '    If tgStationInfo(llLoop).iCode = imShttCode Then
            '        If Len(sNewTimeZone) = 1 Then
            '            tgStationInfo(llLoop).sZone = sNewTimeZone & "ST"
            '        Else
            '            tgStationInfo(llLoop).sZone = sNewTimeZone
            '        End If
            '        slCDStartTime = ""
            '        ilRet = gDetermineAgreementTimes(imShttCode, attrst!attvefCode, Format$(attrst!attOnAir, "m/d/yy"), Format$(attrst!attOffAir, "m/d/yy"), Format$(attrst!attDropDate, "m/d/yy"), slCDStartTime, slStartTime, slEndTime)
            '        slSQLQuery = "Update att Set "
            '        slSQLQuery = slSQLQuery & "attVehProgStartTime = '" & Format$(slStartTime, sgSQLTimeForm) & "', "
            '        slSQLQuery = slSQLQuery & "attVehProgEndTime = '" & Format$(slEndTime, sgSQLTimeForm) & "'"
            '        slSQLQuery = slSQLQuery & " Where attCode = " & attrst!attCode
            '        'cnn.Execute SQLQuery, rdExecDirect
            '        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '            GoSub ErrHand:
            '        End If
            '        Exit For
            '    End If
            'Next llLoop
            slStartTime = "1/1/2000 " & Format(attrst!attVehProgStartTime, sgShowTimeWSecForm)
            slStartTime = DateAdd("h", ilFinalTimeAdj, slStartTime)
            slEndTime = "1/1/2000 " & Format(attrst!attVehProgEndTime, sgShowTimeWSecForm)
            slEndTime = DateAdd("h", ilFinalTimeAdj, slEndTime)
            slSQLQuery = "Update att Set "
            slSQLQuery = slSQLQuery & "attVehProgStartTime = '" & Format$(slStartTime, sgSQLTimeForm) & "', "
            slSQLQuery = slSQLQuery & "attVehProgEndTime = '" & Format$(slEndTime, sgSQLTimeForm) & "'"
            slSQLQuery = slSQLQuery & " Where attCode = " & attrst!attCode
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "Station-mZoneChange"
                mZoneChange = False
                Exit Function
            End If
            
            '********************************* Update Web Site *********************************
            'ilRet = gAdjustWebTimeZone(lmAttCode(llAtt), 3600 * ilFinalTimeAdj)

            If (slattExportToMarketron <> "Y") Then
                slSQLQuery = "Select Count(*) from Spots Where attCode = " & lmAttCode(llAtt)
                llSpotCount = gExecWebSQLWithRowsEffected(slSQLQuery)
                If llSpotCount > 0 Then
                    slSQLQuery = "Update Spots Set FeedTime = DateAdd(HOUR, " & ilFinalTimeAdj & ", FeedTime) Where attCode = " & lmAttCode(llAtt)
                    llRowsEffected = gExecWebSQLWithRowsEffected(slSQLQuery)
                    If llRowsEffected = -1 Then
                        mZoneChange = False
                        gLogMsg "Error: Failed to Update Web: " & SQLQuery, slToFile, False
                        Exit Function
                    End If
                End If
                slSQLQuery = "Select Count(*) from Spot_History Where attCode = " & lmAttCode(llAtt)
                llSpotCount = gExecWebSQLWithRowsEffected(slSQLQuery)
                If llSpotCount > 0 Then
                    slSQLQuery = "Update Spot_History Set FeedTime = DateAdd(HOUR, " & ilFinalTimeAdj & ", FeedTime) Where attCode = " & lmAttCode(llAtt)
                    llRowsEffected = gExecWebSQLWithRowsEffected(slSQLQuery)
                    If llRowsEffected = -1 Then
                        mZoneChange = False
                        gLogMsg "Error: Failed to Update Web: " & slSQLQuery, slToFile, False
                        Exit Function
                    End If
                End If
            End If
            
            gLogMsg "Process Complete For " & slVefName & " " & slDelay, slToFile, False
            'If attrst!attExportType = 2 Then
            '7701
'            If attrst!attExportToUnivision = "Y" Then
'                gLogMsg "      This was a Univision agreement.  You need to re-export the spots.", slToFile, False
'            End If
'            If attrst!attExportToMarketron = "Y" Then
            If slattExportToMarketron = "Y" Then
                gLogMsg "      This was a Marketron agreement.  You need to re-export the spots.", slToFile, False
            End If

        End If
    Next llAtt
    
    gLogMsg "", slToFile, False
    gLogMsg "** Zone Change Completed: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **", slToFile, False
    gLogMsg "", slToFile, False
    mZoneChange = True
    gMsgBox "Please refer to file " & slToFile & " in the messages folder for results", vbOKOnly
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-mZoneChange"
End Function

Public Function mZoneChangesDate(iOffSet As Integer, sDatTime As String) As Integer

    Dim slOrigDtTime As String
    Dim slNewDtTime As String
    
    slOrigDtTime = "1/1/2000 " & sDatTime
    slNewDtTime = DateAdd("h", iOffSet, slOrigDtTime)
    
    'did we back up one day?
    If DateValue(slOrigDtTime) < DateValue(slNewDtTime) Then
        mZoneChangesDate = -1
        Exit Function
    End If
    
    'did we go forward one day?
    If DateValue(slOrigDtTime) > DateValue(slNewDtTime) Then
        mZoneChangesDate = 1
        Exit Function
    End If
    
    'No change in date
    mZoneChangesDate = 0
    
End Function


Private Sub mPopTerritory()
    Dim slTerritoryName As String
    Dim llTerritoryMntCode As Long
    Dim ilMnt As Integer
    
    On Error GoTo ErrHand
    
    If cbcTerritory.ListIndex > 1 Then
        slTerritoryName = Trim$(cbcTerritory.Text)
        llTerritoryMntCode = cbcTerritory.GetItemData(cbcTerritory.ListIndex)
    Else
        slTerritoryName = ""
        llTerritoryMntCode = -2
    End If


    cbcTerritory.Clear
    cbcTerritory.AddItem ("[New]")
    cbcTerritory.SetItemData = -1
    cbcTerritory.AddItem ("[None]")
    cbcTerritory.SetItemData = 0
    SQLQuery = "SELECT * FROM Mnt WHERE mntType = 'T' ORDER BY mntName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        cbcTerritory.AddItem Trim$(rst!mntName)
        cbcTerritory.SetItemData = rst!mntCode
        rst.MoveNext
    Loop
    'If slTerritoryName <> "" Then
    If llTerritoryMntCode > 0 Then
        For ilMnt = 0 To cbcTerritory.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcTerritory.GetName(ilMnt)), slTerritoryName, vbTextCompare) = 0 Then
            If cbcTerritory.GetItemData(ilMnt) = llTerritoryMntCode Then
                cbcTerritory.SetListIndex = ilMnt
                Exit For
            End If
        Next ilMnt
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mPopTerritory"
End Sub

Private Sub mPopMulticast()
    Dim ilShtt As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim ilLoop As Integer
    Dim llLoop As Long
    Dim blIncludeStation As Boolean
    Dim temp2_rst As ADODB.Recordset
    Dim temp_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gSetMousePointer grdMulticast, grdMulticast, vbHourglass
    
    With grdMulticast
        '.Redraw = False
        .Rows = 2
        .TextMatrix(grdMulticast.FixedRows, MCCALLLETTERSINDEX) = ""
        .TextMatrix(grdMulticast.FixedRows, MCMARKETINDEX) = ""
        .TextMatrix(grdMulticast.FixedRows, MCLICCITYINDEX) = ""
        .TextMatrix(grdMulticast.FixedRows, MCMAILSTATEINDEX) = ""
        .TextMatrix(grdMulticast.FixedRows, MCOWNERINDEX) = ""
        .Row = grdMulticast.FixedRows
        
        For llCol = MCCALLLETTERSINDEX To MCOWNERINDEX Step 1
            .Col = llCol
            .CellBackColor = vbWhite
        Next llCol
    
        llRow = .FixedRows
    End With
    
    If (rbcMulticast(1).Value) And (lmMultiCastGroupID <= 0) Then   'Add to Group
        lacMulticastNote.Caption = "Add " & Trim$(txtCallLetters.Text) & " to Multi-Cast:"
        lacMulticastNote.Visible = True
        
        With grdMulticast
            .ColWidth(MCOWNERINDEX) = 0
            .ColWidth(MCMARKETINDEX) = 0
            .ColWidth(MCLICCITYINDEX) = 0
            .ColWidth(MCMAILSTATEINDEX) = 0
            .ColWidth(MCCALLLETTERSINDEX) = grdMulticast.Width - GRIDSCROLLWIDTH  '(5 * grdStation.Columns(6).Width) / 6
        End With
        
        SQLQuery = "SELECT DISTINCT shttMultiCastGroupID FROM shtt ORDER BY shttMultiCastGroupID"
        Set temp_rst = gSQLSelectCall(SQLQuery)
        
        Do While Not temp_rst.EOF
            If temp_rst!shttMultiCastGroupID > 0 Then
            
                SQLQuery = ""
                SQLQuery = SQLQuery & " SELECT "
                SQLQuery = SQLQuery & "     shttCallLetters "
                SQLQuery = SQLQuery & " FROM "
                SQLQuery = SQLQuery & "     shtt "
                SQLQuery = SQLQuery & " WHERE "
                SQLQuery = SQLQuery & "     shttMultiCastGroupID = " & temp_rst!shttMultiCastGroupID
                
                Set temp2_rst = gSQLSelectCall(SQLQuery)
                
                slStr = ""
                Do While Not temp2_rst.EOF
                    If slStr = "" Then
                        slStr = Trim$(temp2_rst!shttCallLetters)
                    Else
                        slStr = slStr & "," & Trim$(temp2_rst!shttCallLetters)
                    End If
                    temp2_rst.MoveNext
                Loop
                
                With grdMulticast
                    If llRow + 1 > .Rows Then
                        .AddItem ""
                    End If
                    .Row = llRow
                    .TextMatrix(llRow, MCCALLLETTERSINDEX) = slStr
                    .TextMatrix(llRow, MCMARKETINDEX) = ""
                    .TextMatrix(llRow, MCOWNERINDEX) = ""
                    .TextMatrix(llRow, MCSHTTCODEINDEX) = temp_rst!shttMultiCastGroupID
                    .TextMatrix(llRow, MCDMAMKTCODEINDEX) = ""
                    llRow = llRow + 1
                End With
                
            End If
            temp_rst.MoveNext
        Loop
        
    ElseIf rbcMulticast(2).Value Then   'Remove from Group
        
        lacMulticastNote.Caption = ""   'Trim$(txtCallLetters.Text) & " Multicast with:"
        lacMulticastNote.Visible = False
        
        With grdMulticast
            .ColWidth(MCCALLLETTERSINDEX) = grdMulticast.Width * 0.15
            .ColWidth(MCMARKETINDEX) = grdMulticast.Width * 0.25
            .ColWidth(MCLICCITYINDEX) = grdMulticast.Width * 0.25
            .ColWidth(MCMAILSTATEINDEX) = grdMulticast.Width * 0.2
            .ColWidth(MCOWNERINDEX) = .Width - .ColWidth(MCCALLLETTERSINDEX) - .ColWidth(MCMARKETINDEX) - .ColWidth(MCLICCITYINDEX) - .ColWidth(MCMAILSTATEINDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        End With
        
        SQLQuery = ""
        SQLQuery = SQLQuery & vbCrLf & " SELECT "
        SQLQuery = SQLQuery & vbCrLf & "       shttMultiCastGroupID "
        SQLQuery = SQLQuery & vbCrLf & " FROM "
        SQLQuery = SQLQuery & vbCrLf & "       shtt "
        SQLQuery = SQLQuery & vbCrLf & " WHERE "
        SQLQuery = SQLQuery & vbCrLf & "       shttCode = " & imShttCode
        
        Set temp_rst = gSQLSelectCall(SQLQuery)
        
        If Not temp_rst.EOF Then
            lmMultiCastGroupID = temp_rst!shttMultiCastGroupID
            If lmMultiCastGroupID > 0 Then
            
                SQLQuery = ""
                SQLQuery = SQLQuery & " SELECT "
                SQLQuery = SQLQuery & "     shttCallLetters, "
                SQLQuery = SQLQuery & "     shttCode, "
                SQLQuery = SQLQuery & "     shttMktCode, "
                SQLQuery = SQLQuery & "     shttCityLicMntCode, "
                SQLQuery = SQLQuery & "     shttState,  "
                SQLQuery = SQLQuery & "     mktName, "
                SQLQuery = SQLQuery & "     arttLastName "
                SQLQuery = SQLQuery & " FROM "
                SQLQuery = SQLQuery & "     shtt "
                SQLQuery = SQLQuery & "         LEFT JOIN artt ON shttOwnerArttCode = arttCode "
                SQLQuery = SQLQuery & "         LEFT JOIN mkt ON shttMktCode = mktCode"
                SQLQuery = SQLQuery & " WHERE "
                SQLQuery = SQLQuery & "     shttMultiCastGroupID = " & lmMultiCastGroupID
                
                Set temp2_rst = gSQLSelectCall(SQLQuery)
                
                Do While Not temp2_rst.EOF
                    If llRow + 1 > grdMulticast.Rows Then
                        grdMulticast.AddItem ""
                    End If
                    grdMulticast.Row = llRow
                    grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX) = Trim$(temp2_rst!shttCallLetters)
                    
                    If IsNull(temp2_rst!mktName) Then
                        grdMulticast.TextMatrix(llRow, MCMARKETINDEX) = ""
                    Else
                        grdMulticast.TextMatrix(llRow, MCMARKETINDEX) = Trim$(temp2_rst!mktName)
                    End If
                    
                    grdMulticast.TextMatrix(llRow, SSLICCITYINDEX) = ""
                    If temp2_rst!shttCityLicMntCode > 0 Then
                        For ilLoop = 0 To cbcCityLic.ListCount - 1 Step 1
                            If cbcCityLic.GetItemData(ilLoop) = temp2_rst!shttCityLicMntCode Then
                                grdMulticast.TextMatrix(llRow, SSLICCITYINDEX) = Trim$(cbcCityLic.GetName(ilLoop))
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    grdMulticast.TextMatrix(llRow, SSMAILSTATEINDEX) = ""
                    If Trim$(temp2_rst!shttState) <> "" Then
                        For ilLoop = 0 To cboState.ListCount - 1 Step 1
                            If StrComp(Trim$(temp2_rst!shttState), Trim$(tgStateInfo(cboState.ItemData(ilLoop)).sPostalName), vbTextCompare) = 0 Then
                                grdMulticast.TextMatrix(llRow, SSMAILSTATEINDEX) = Trim$(cboState.List(ilLoop))
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If IsNull(temp2_rst!arttLastName) Then
                        grdMulticast.TextMatrix(llRow, MCOWNERINDEX) = ""
                    Else
                        grdMulticast.TextMatrix(llRow, MCOWNERINDEX) = Trim$(temp2_rst!arttLastName)
                    End If
                    grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "0"
                    grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX) = temp2_rst!shttCode
                    grdMulticast.TextMatrix(llRow, MCDMAMKTCODEINDEX) = temp2_rst!shttMktCode
                    mMCPaintRowColor llRow
                    llRow = llRow + 1
                    temp2_rst.MoveNext
                Loop
            End If
        End If
        
    ElseIf (rbcMulticast(0).Value) Or ((rbcMulticast(1).Value) And (lmMultiCastGroupID > 0)) Then    'Create Group
        
        If (rbcMulticast(0).Value) Then
            lacMulticastNote.Caption = "Create Multi-Cast with " & Trim$(txtCallLetters.Text) & " and:"
        Else
            lacMulticastNote.Caption = "Add to Multi-Cast with " & Trim$(txtCallLetters.Text)
        End If
        
        lacMulticastNote.Visible = True
        grdMulticast.ColWidth(MCCALLLETTERSINDEX) = grdMulticast.Width * 0.15
        grdMulticast.ColWidth(MCMARKETINDEX) = grdMulticast.Width * 0.25
        grdMulticast.ColWidth(MCLICCITYINDEX) = grdMulticast.Width * 0.25
        grdMulticast.ColWidth(MCMAILSTATEINDEX) = grdMulticast.Width * 0.2
        grdMulticast.ColWidth(MCOWNERINDEX) = grdMulticast.Width - grdMulticast.ColWidth(MCCALLLETTERSINDEX) - grdMulticast.ColWidth(MCMARKETINDEX) - grdMulticast.ColWidth(MCLICCITYINDEX) - grdMulticast.ColWidth(MCMAILSTATEINDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        
        lmArttCode = 0
        If cbcOwner.ListIndex > 1 Then
            lmArttCode = cbcOwner.GetItemData(cbcOwner.ListIndex)
        End If
        
        imDMAMktCode = 0
        If cbcDMAMarket.ListIndex > 1 Then
            imDMAMktCode = cbcDMAMarket.GetItemData(cbcDMAMarket.ListIndex)
        End If
        
' *************************** TTP 10688 JJB 2023-05-16
        SQLQuery = ""
        SQLQuery = SQLQuery & " SELECT "
        SQLQuery = SQLQuery & "     shttCallLetters, "
        SQLQuery = SQLQuery & "     shttCode, "
        SQLQuery = SQLQuery & "     shttMktCode, "
        SQLQuery = SQLQuery & "     shttCityLicMntCode, "
        SQLQuery = SQLQuery & "     shttState, "
        SQLQuery = SQLQuery & "     mktName, "
        SQLQuery = SQLQuery & "     arttLastName "
        SQLQuery = SQLQuery & " FROM "
        SQLQuery = SQLQuery & "     shtt "
        SQLQuery = SQLQuery & "         LEFT JOIN artt ON shttOwnerArttCode = arttCode "
        SQLQuery = SQLQuery & "         LEFT JOIN mkt ON shttMktCode = mktCode "
        
        If rbcMulticast(0).Value Then 'New Group
            If rbcMulticastOwner(0).Value Then 'Same Owner
                SQLQuery = SQLQuery & " WHERE shttOwnerArttCode = " & lmArttCode
                If rbcMulticastMarket(0).Value Then
                    SQLQuery = SQLQuery & " AND shttMktCode = " & imDMAMktCode
                End If
            Else
                If rbcMulticastMarket(0).Value Then
                    SQLQuery = SQLQuery & " WHERE shttMktCode = " & imDMAMktCode
                End If
            End If
        End If
        
        If rbcMulticast(1).Value Then 'Add to group
            If rbcMulticastOwner_Add(0).Value Then 'Same Owner
                SQLQuery = SQLQuery & " WHERE shttOwnerArttCode = " & lmArttCode
                If rbcMulticastMarket_Add(0).Value Then
                    SQLQuery = SQLQuery & " AND shttMktCode = " & imDMAMktCode
                End If
            Else
                If rbcMulticastMarket_Add(0).Value Then
                    SQLQuery = SQLQuery & " WHERE shttMktCode = " & imDMAMktCode
                End If
            End If
        End If
        
        SQLQuery = SQLQuery & " ORDER BY "
        SQLQuery = SQLQuery & "     shttCallLetters"
        
        Set temp2_rst = gSQLSelectCall(SQLQuery)
        
        Dim i As Integer
        i = 1
        Do While Not temp2_rst.EOF
            If temp2_rst!shttCode <> imShttCode Then
                If llRow + 1 > grdMulticast.Rows Then
                    grdMulticast.AddItem ""
                End If
                grdMulticast.Row = llRow
                grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX) = Trim$(temp2_rst!shttCallLetters)
                
                If IsNull(temp2_rst!mktName) Then
                    grdMulticast.TextMatrix(llRow, MCMARKETINDEX) = ""
                Else
                    grdMulticast.TextMatrix(llRow, MCMARKETINDEX) = Trim$(temp2_rst!mktName)
                End If
                
                grdMulticast.TextMatrix(llRow, SSLICCITYINDEX) = ""
                If temp2_rst!shttCityLicMntCode > 0 Then
                    For ilLoop = 0 To cbcCityLic.ListCount - 1 Step 1
                        If cbcCityLic.GetItemData(ilLoop) = temp2_rst!shttCityLicMntCode Then
                            grdMulticast.TextMatrix(llRow, SSLICCITYINDEX) = Trim$(cbcCityLic.GetName(ilLoop))
                            Exit For
                        End If
                    Next ilLoop
                End If
                
                grdMulticast.TextMatrix(llRow, SSMAILSTATEINDEX) = ""
                If Trim$(temp2_rst!shttState) <> "" Then
                    For ilLoop = 0 To cboState.ListCount - 1 Step 1
                        If StrComp(Trim$(temp2_rst!shttState), Trim$(tgStateInfo(cboState.ItemData(ilLoop)).sPostalName), vbTextCompare) = 0 Then
                            grdMulticast.TextMatrix(llRow, SSMAILSTATEINDEX) = Trim$(cboState.List(ilLoop))
                            Exit For
                        End If
                    Next ilLoop
                End If
                
                If IsNull(temp2_rst!arttLastName) Then
                    grdMulticast.TextMatrix(llRow, MCOWNERINDEX) = ""
                Else
                    grdMulticast.TextMatrix(llRow, MCOWNERINDEX) = Trim$(temp2_rst!arttLastName)
                End If
                
                grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX) = temp2_rst!shttCode
                grdMulticast.TextMatrix(llRow, MCDMAMKTCODEINDEX) = temp2_rst!shttMktCode
                llRow = llRow + 1
            End If
            temp2_rst.MoveNext
            i = i + 1
        Loop

'************************

'        For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
'            blIncludeStation = True
'            If tgStationInfo(llLoop).lMultiCastGroupID > 0 Then
'                blIncludeStation = False
'            End If
'            If rbcMulticastOwner(0).Value Then
'                If tgStationInfo(llLoop).lOwnerCode <> lmArttCode Then
'                    blIncludeStation = False
'                End If
'            End If
'            If rbcMulticastMarket(0).Value Then
'                If tgStationInfo(llLoop).iMktCode <> imDMAMktCode Then
'                    blIncludeStation = False
'                End If
'            End If
'            If blIncludeStation Then
'                If tgStationInfo(llLoop).iCode = imShttCode Then
'                    blIncludeStation = False
'                End If
'            End If
'            If tgStationInfo(llLoop).iType = 1 Then
'                blIncludeStation = False
'            End If
'            If blIncludeStation Then
'                If llRow + 1 > grdMulticast.Rows Then
'                    grdMulticast.AddItem ""
'                End If
'                grdMulticast.Row = llRow
'                grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX) = Trim$(tgStationInfo(llLoop).sCallLetters)
'                grdMulticast.TextMatrix(llRow, MCMARKETINDEX) = ""
'                If tgStationInfo(llLoop).iMktCode > 0 Then
'                    For ilLoop = 0 To cbcDMAMarket.ListCount - 1 Step 1
'                        If cbcDMAMarket.GetItemData(ilLoop) = tgStationInfo(llLoop).iMktCode Then
'                            grdMulticast.TextMatrix(llRow, MCMARKETINDEX) = Trim$(cbcDMAMarket.GetName(ilLoop))
'                            Exit For
'                        End If
'                    Next ilLoop
'                End If
'                grdMulticast.TextMatrix(llRow, MCLICCITYINDEX) = ""
'                If tgStationInfo(llLoop).lCityLicMntCode > 0 Then
'                    For ilLoop = 0 To cbcCityLic.ListCount - 1 Step 1
'                        If cbcCityLic.GetItemData(ilLoop) = tgStationInfo(llLoop).lCityLicMntCode Then
'                            grdMulticast.TextMatrix(llRow, MCLICCITYINDEX) = Trim$(cbcCityLic.GetName(ilLoop))
'                            Exit For
'                        End If
'                    Next ilLoop
'                End If
'                grdMulticast.TextMatrix(llRow, MCMAILSTATEINDEX) = ""
'                If Trim$(tgStationInfo(llLoop).sPostalName) <> "" Then
'                    For ilLoop = 0 To cboState.ListCount - 1 Step 1
'                        If StrComp(Trim$(tgStationInfo(llLoop).sPostalName), Trim$(tgStateInfo(cboState.ItemData(ilLoop)).sPostalName), vbTextCompare) = 0 Then
'                            grdMulticast.TextMatrix(llRow, MCMAILSTATEINDEX) = Trim$(cboState.List(ilLoop))
'                            Exit For
'                        End If
'                    Next ilLoop
'                End If
'                grdMulticast.TextMatrix(llRow, MCOWNERINDEX) = ""
'                If tgStationInfo(llLoop).lOwnerCode > 0 Then
'                    For ilLoop = 0 To cbcOwner.ListCount - 1 Step 1
'                        If cbcOwner.GetItemData(ilLoop) = tgStationInfo(llLoop).lOwnerCode Then
'                            grdMulticast.TextMatrix(llRow, MCOWNERINDEX) = Trim$(cbcOwner.GetName(ilLoop))
'                            Exit For
'                        End If
'                    Next ilLoop
'                End If
'                grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX) = tgStationInfo(llLoop).iCode
'                grdMulticast.TextMatrix(llRow, MCDMAMKTCODEINDEX) = tgStationInfo(llLoop).iMktCode
'                llRow = llRow + 1
'            End If
'        Next llLoop
    End If
    
    gSetMousePointer grdMulticast, grdMulticast, vbDefault
    grdMulticast.Redraw = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mPopMulticast"
End Sub

Private Sub mPopSisterStations()
    Dim ilShtt As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim ilLoop As Integer
    Dim llLoop As Long
    Dim blIncludeStation As Boolean
    Dim temp_rst As ADODB.Recordset
    Dim temp2_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    edcMasterStation.Text = ""
    gSetMousePointer grdSisterStations, grdSisterStations, vbHourglass
    grdSisterStations.Redraw = False
    grdSisterStations.Rows = 2
    grdSisterStations.TextMatrix(grdSisterStations.FixedRows, SSCALLLETTERSINDEX) = ""
    grdSisterStations.TextMatrix(grdSisterStations.FixedRows, SSMARKETINDEX) = ""
    grdSisterStations.TextMatrix(grdSisterStations.FixedRows, SSLICCITYINDEX) = ""
    grdSisterStations.TextMatrix(grdSisterStations.FixedRows, SSMAILSTATEINDEX) = ""
    grdSisterStations.Row = grdSisterStations.FixedRows
    For llCol = SSCALLLETTERSINDEX To SSMAILSTATEINDEX Step 1
        grdSisterStations.Col = llCol
        grdSisterStations.CellBackColor = vbWhite
    Next llCol
    llRow = grdSisterStations.FixedRows
    If (rbcMarketCluster(1).Value) And (lmClusterGroupID <= 0) Then   'Add to Group
        lacMarketClusterNote.Caption = "Add " & Trim$(txtCallLetters.Text) & " to Sister Stations:"
        lacMarketClusterNote.Visible = True
        grdSisterStations.ColWidth(SSMARKETINDEX) = 0
        grdSisterStations.ColWidth(SSLICCITYINDEX) = 0
        grdSisterStations.ColWidth(SSMAILSTATEINDEX) = 0
        grdSisterStations.ColWidth(SSCALLLETTERSINDEX) = grdSisterStations.Width - GRIDSCROLLWIDTH  '(5 * grdStation.Columns(6).Width) / 6
        SQLQuery = "SELECT DISTINCT shttClusterGroupID FROM shtt ORDER BY shttClusterGroupID"
        Set temp_rst = gSQLSelectCall(SQLQuery)
        Do While Not temp_rst.EOF
            If temp_rst!shttclustergroupId > 0 Then
                '9/25/15
                slStr = ""
                SQLQuery = "SELECT shttCallLetters, shttMasterCluster FROM shtt WHERE shttClusterGroupID = " & temp_rst!shttclustergroupId
                Set temp2_rst = gSQLSelectCall(SQLQuery)
                Do While Not temp2_rst.EOF
                    If slStr = "" Then
                        slStr = Trim$(temp2_rst!shttCallLetters)
                    Else
                        slStr = slStr & "," & Trim$(temp2_rst!shttCallLetters)
                    End If
                    temp2_rst.MoveNext
                Loop
                If llRow + 1 > grdSisterStations.Rows Then
                    grdSisterStations.AddItem ""
                End If
                grdSisterStations.Row = llRow
                grdSisterStations.TextMatrix(llRow, SSCALLLETTERSINDEX) = slStr
                grdSisterStations.TextMatrix(llRow, SSMARKETINDEX) = ""
                grdSisterStations.TextMatrix(llRow, SSLICCITYINDEX) = ""
                grdSisterStations.TextMatrix(llRow, SSMAILSTATEINDEX) = ""
                grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX) = temp_rst!shttclustergroupId
                'If temp_rst!shttMasterCluster = "Y" Then
                '    edcMasterStation.Text = slStr
                'End If
                llRow = llRow + 1
            End If
            temp_rst.MoveNext
        Loop
    ElseIf rbcMarketCluster(2).Value Then   'Remove from Group
        lacMarketClusterNote.Caption = ""   'Trim$(txtCallLetters.Text) & " MarketCluster with:"
        lacMarketClusterNote.Visible = False
        smClusterStations = ""
        grdSisterStations.ColWidth(SSCALLLETTERSINDEX) = grdSisterStations.Width * 0.2
        grdSisterStations.ColWidth(SSLICCITYINDEX) = grdSisterStations.Width * 0.3
        grdSisterStations.ColWidth(SSMAILSTATEINDEX) = grdSisterStations.Width * 0.2
        grdSisterStations.ColWidth(SSMARKETINDEX) = grdSisterStations.Width - grdSisterStations.ColWidth(SSCALLLETTERSINDEX) - grdSisterStations.ColWidth(SSLICCITYINDEX) - grdSisterStations.ColWidth(SSMAILSTATEINDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        SQLQuery = "SELECT shttClusterGroupID FROM shtt WHERE shttCode = " & imShttCode
        Set temp_rst = gSQLSelectCall(SQLQuery)
        If Not temp_rst.EOF Then
            lmClusterGroupID = temp_rst!shttclustergroupId
            If lmClusterGroupID > 0 Then
                SQLQuery = "SELECT shttCallLetters, shttCode, shttMktCode, shttMasterCluster, shttCityLicMntCode, shttState, mktName FROM shtt LEFT JOIN mkt ON shttMktCode = mktCode"
                SQLQuery = SQLQuery & " WHERE shttClusterGroupID = " & lmClusterGroupID
                Set temp2_rst = gSQLSelectCall(SQLQuery)
                Do While Not temp2_rst.EOF
                    If llRow + 1 > grdSisterStations.Rows Then
                        grdSisterStations.AddItem ""
                    End If
                    grdSisterStations.Row = llRow
                    smClusterStations = smClusterStations & " " & Trim$(temp2_rst!shttCallLetters)
                    grdSisterStations.TextMatrix(llRow, SSCALLLETTERSINDEX) = Trim$(temp2_rst!shttCallLetters)
                    If IsNull(temp2_rst!mktName) Then
                        grdSisterStations.TextMatrix(llRow, SSMARKETINDEX) = ""
                    Else
                        grdSisterStations.TextMatrix(llRow, SSMARKETINDEX) = Trim$(temp2_rst!mktName)
                    End If
                    grdSisterStations.TextMatrix(llRow, SSLICCITYINDEX) = ""
                    If temp2_rst!shttCityLicMntCode > 0 Then
                        For ilLoop = 0 To cbcCityLic.ListCount - 1 Step 1
                            If cbcCityLic.GetItemData(ilLoop) = temp2_rst!shttCityLicMntCode Then
                                grdSisterStations.TextMatrix(llRow, SSLICCITYINDEX) = Trim$(cbcCityLic.GetName(ilLoop))
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    grdSisterStations.TextMatrix(llRow, SSMAILSTATEINDEX) = ""
                    If Trim$(temp2_rst!shttState) <> "" Then
                        For ilLoop = 0 To cboState.ListCount - 1 Step 1
                            If StrComp(Trim$(temp2_rst!shttState), Trim$(tgStateInfo(cboState.ItemData(ilLoop)).sPostalName), vbTextCompare) = 0 Then
                                grdSisterStations.TextMatrix(llRow, SSMAILSTATEINDEX) = Trim$(cboState.List(ilLoop))
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "0"
                    grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX) = temp2_rst!shttCode
                    mSSPaintRowColor llRow
                    If temp2_rst!shttMasterCluster = "Y" Then
                        edcMasterStation.Text = Trim$(temp2_rst!shttCallLetters)
                    End If
                    llRow = llRow + 1
                    temp2_rst.MoveNext
                Loop
                smClusterStations = Trim$(smClusterStations)
            End If
            'SQLQuery = "SELECT shttCallLetters FROM shtt "
            'SQLQuery = SQLQuery & " WHERE shttClusterGroupID = " & lmClusterGroupID
            'SQLQuery = SQLQuery & " AND sgtMasterStation = '" & "Y" & "'"
            'Set temp2_rst = gSQLSelectCall(SQLQuery)
            'If Not temp2_rst.EOF Then
            '    edcMasterStation.Text = Trim$(temp2_rst!shttCallLetters)
            'End If
        End If
    ElseIf (rbcMarketCluster(0).Value) Or ((rbcMarketCluster(1).Value) And (lmClusterGroupID > 0)) Then    'Create Group
        If (rbcMarketCluster(0).Value) Then
            lacMarketClusterNote.Caption = "Create Sister Stations with " & Trim$(txtCallLetters.Text) & " and:"
        Else
            lacMarketClusterNote.Caption = "Add to Sister Stations " & smClusterStations
        End If
        lacMarketClusterNote.Visible = True
        edcMasterStation.Text = Trim$(txtCallLetters.Text)
        grdSisterStations.ColWidth(SSCALLLETTERSINDEX) = grdSisterStations.Width * 0.2
        grdSisterStations.ColWidth(SSLICCITYINDEX) = grdSisterStations.Width * 0.3
        grdSisterStations.ColWidth(SSMAILSTATEINDEX) = grdSisterStations.Width * 0.2
        grdSisterStations.ColWidth(SSMARKETINDEX) = grdSisterStations.Width - grdSisterStations.ColWidth(SSCALLLETTERSINDEX) - grdSisterStations.ColWidth(SSLICCITYINDEX) - grdSisterStations.ColWidth(SSMAILSTATEINDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        lmArttCode = 0
        If cbcOwner.ListIndex > 1 Then
            lmArttCode = cbcOwner.GetItemData(cbcOwner.ListIndex)
        End If
        imDMAMktCode = 0
        If cbcDMAMarket.ListIndex > 1 Then
            imDMAMktCode = cbcDMAMarket.GetItemData(cbcDMAMarket.ListIndex)
        End If
'        SQLQuery = "SELECT shttCallLetters, shttCode, shttMktCode, shttCityLicMntCode, shttState, mktName FROM shtt LEFT JOIN mkt ON shttMktCode = mktCode"
'        'Same Owner
'        SQLQuery = SQLQuery & " WHERE shttOwnerArttCode = " & lmArttCode
'        If rbcMarketClusterMarket(0).Value Then
'            SQLQuery = SQLQuery & " AND shttMktCode = " & imDMAMktCode
'        End If
'        SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
'        Set temp2_rst = gSQLSelectCall(SQLQuery)
'        Do While Not temp2_rst.EOF
'            If temp2_rst!shttCode <> imShttCode Then
'                If llRow + 1 > grdSisterStations.Rows Then
'                    grdSisterStations.AddItem ""
'                End If
'                grdSisterStations.Row = llRow
'                grdSisterStations.TextMatrix(llRow, SSCALLLETTERSINDEX) = Trim$(temp2_rst!shttCallLetters)
'                If IsNull(temp2_rst!mktName) Then
'                    grdSisterStations.TextMatrix(llRow, SSMARKETINDEX) = ""
'                Else
'                    grdSisterStations.TextMatrix(llRow, SSMARKETINDEX) = Trim$(temp2_rst!mktName)
'                End If
'                grdSisterStations.TextMatrix(llRow, SSLICCITYINDEX) = ""
'                If temp2_rst!shttCityLicMntCode > 0 Then
'                    For ilLoop = 0 To cbcCityLic.ListCount - 1 Step 1
'                        If cbcCityLic.GetItemData(ilLoop) = temp2_rst!shttCityLicMntCode Then
'                            grdSisterStations.TextMatrix(llRow, SSLICCITYINDEX) = Trim$(cbcCityLic.GetName(ilLoop))
'                            Exit For
'                        End If
'                    Next ilLoop
'                End If
'                grdSisterStations.TextMatrix(llRow, SSMAILSTATEINDEX) = ""
'                If Trim$(temp2_rst!shttState) <> "" Then
'                    For ilLoop = 0 To cboState.ListCount - 1 Step 1
'                        If StrComp(Trim$(temp2_rst!shttState), Trim$(tgStateInfo(cboState.ItemData(ilLoop)).sPostalName), vbTextCompare) = 0 Then
'                            grdSisterStations.TextMatrix(llRow, SSMAILSTATEINDEX) = Trim$(cboState.List(ilLoop))
'                            Exit For
'                        End If
'                    Next ilLoop
'                End If
'                grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX) = temp2_rst!shttCode
'                llRow = llRow + 1
'            End If
'            temp2_rst.MoveNext
'        Loop
        For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            blIncludeStation = True
            If tgStationInfo(llLoop).lMarketClusterGroupID > 0 Then
                blIncludeStation = False
            End If
            If tgStationInfo(llLoop).lOwnerCode <> lmArttCode Then
                blIncludeStation = False
            End If
            If rbcMarketClusterMarket(0).Value Then
                If tgStationInfo(llLoop).iMktCode <> imDMAMktCode Then
                    blIncludeStation = False
                End If
            End If
            If blIncludeStation Then
                If tgStationInfo(llLoop).iCode = imShttCode Then
                    blIncludeStation = False
                End If
            End If
            If blIncludeStation Then
                If llRow + 1 > grdSisterStations.Rows Then
                    grdSisterStations.AddItem ""
                End If
                grdSisterStations.Row = llRow
                grdSisterStations.TextMatrix(llRow, SSCALLLETTERSINDEX) = Trim$(tgStationInfo(llLoop).sCallLetters)
                grdSisterStations.TextMatrix(llRow, SSMARKETINDEX) = ""
                If tgStationInfo(llLoop).iMktCode > 0 Then
                    For ilLoop = 0 To cbcDMAMarket.ListCount - 1 Step 1
                        If cbcDMAMarket.GetItemData(ilLoop) = tgStationInfo(llLoop).iMktCode Then
                            grdSisterStations.TextMatrix(llRow, SSMARKETINDEX) = Trim$(cbcDMAMarket.GetName(ilLoop))
                            Exit For
                        End If
                    Next ilLoop
                End If
                grdSisterStations.TextMatrix(llRow, SSLICCITYINDEX) = ""
                If tgStationInfo(llLoop).lCityLicMntCode > 0 Then
                    For ilLoop = 0 To cbcCityLic.ListCount - 1 Step 1
                        If cbcCityLic.GetItemData(ilLoop) = tgStationInfo(llLoop).lCityLicMntCode Then
                            grdSisterStations.TextMatrix(llRow, SSLICCITYINDEX) = Trim$(cbcCityLic.GetName(ilLoop))
                            Exit For
                        End If
                    Next ilLoop
                End If
                grdSisterStations.TextMatrix(llRow, SSMAILSTATEINDEX) = ""
                If Trim$(tgStationInfo(llLoop).sPostalName) <> "" Then
                    For ilLoop = 0 To cboState.ListCount - 1 Step 1
                        If StrComp(Trim$(tgStationInfo(llLoop).sPostalName), Trim$(tgStateInfo(cboState.ItemData(ilLoop)).sPostalName), vbTextCompare) = 0 Then
                            grdSisterStations.TextMatrix(llRow, SSMAILSTATEINDEX) = Trim$(cboState.List(ilLoop))
                            Exit For
                        End If
                    Next ilLoop
                End If
                grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX) = tgStationInfo(llLoop).iCode
                llRow = llRow + 1
            End If
        Next llLoop
    End If
    gSetMousePointer grdSisterStations, grdSisterStations, vbDefault
    grdSisterStations.Redraw = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mPopSisterStations"
End Sub
Private Sub mPopArea()
    Dim slAreaName As String
    Dim llAreaMntCode As Long
    Dim ilMnt As Integer
    
    On Error GoTo ErrHand
    
    If cbcArea.ListIndex > 1 Then
        slAreaName = Trim$(cbcArea.Text)
        llAreaMntCode = cbcArea.GetItemData(cbcArea.ListIndex)
    Else
        slAreaName = ""
        llAreaMntCode = -2
    End If


    cbcArea.Clear
    cbcArea.AddItem ("[New]")
    cbcArea.SetItemData = -1
    cbcArea.AddItem ("[None]")
    cbcArea.SetItemData = 0
    SQLQuery = "SELECT * FROM Mnt WHERE mntType = 'A' ORDER BY mntName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        cbcArea.AddItem Trim$(rst!mntName)
        cbcArea.SetItemData = rst!mntCode
        rst.MoveNext
    Loop
    'If slAreaName <> "" Then
    If llAreaMntCode > 0 Then
        For ilMnt = 0 To cbcArea.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcArea.GetName(ilMnt)), slAreaName, vbTextCompare) = 0 Then
            If cbcArea.GetItemData(ilMnt) = llAreaMntCode Then
                cbcArea.SetListIndex = ilMnt
                Exit For
            End If
        Next ilMnt
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mPopArea"
End Sub

Private Sub mPopMarketRep()
    Dim slMarketRepName As String
    Dim ilMktRepUstCode As Integer
    Dim ilUst As Integer
    
    On Error GoTo ErrHand
    
    If cbcMarketRep.ListIndex > 1 Then
        slMarketRepName = Trim$(cbcMarketRep.Text)
        ilMktRepUstCode = cbcMarketRep.GetItemData(cbcMarketRep.ListIndex)
    Else
        slMarketRepName = ""
        ilMktRepUstCode = -2
    End If


    cbcMarketRep.Clear
    cbcMarketRep.AddItem ("[None]")
    cbcMarketRep.SetItemData = 0
    SQLQuery = "SELECT UstName, ustReportName, ustCode FROM Ust INNER JOIN Dnt ON ustDntCode = dntCode WHERE dntType = " & "'M'" & " ORDER BY UstName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        If Trim$(rst!ustReportName) <> "" Then
            cbcMarketRep.AddItem Trim$(rst!ustReportName)
        Else
            cbcMarketRep.AddItem Trim$(rst!ustname)
        End If
        cbcMarketRep.SetItemData = rst!ustCode
        rst.MoveNext
    Loop
    'If slMarketRepName <> "" Then
    If ilMktRepUstCode > 0 Then
        For ilUst = 0 To cbcMarketRep.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcMarketRep.GetName(ilUst)), slMarketRepName, vbTextCompare) = 0 Then
            If cbcMarketRep.GetItemData(ilUst) = ilMktRepUstCode Then
                cbcMarketRep.SetListIndex = ilUst
                Exit For
            End If
        Next ilUst
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mPopMarketRep"
End Sub
Private Sub mPopServiceRep()
    Dim slServiceRepName As String
    Dim ilServRepUstCode As Integer
    Dim ilUst As Integer
    
    On Error GoTo ErrHand
    If cbcServiceRep.ListIndex > 1 Then
        slServiceRepName = Trim$(cbcServiceRep.Text)
        ilServRepUstCode = cbcServiceRep.GetItemData(cbcServiceRep.ListIndex)
    Else
        slServiceRepName = ""
        ilServRepUstCode = -2
    End If
    cbcServiceRep.Clear
    cbcServiceRep.AddItem ("[None]")
    cbcServiceRep.SetItemData = 0
    SQLQuery = "SELECT UstName, ustReportName, ustCode FROM Ust INNER JOIN Dnt ON ustDntCode = dntCode WHERE dntType = " & "'S'" & " ORDER BY UstName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        If Trim$(rst!ustReportName) <> "" Then
            cbcServiceRep.AddItem Trim$(rst!ustReportName)
        Else
            cbcServiceRep.AddItem Trim$(rst!ustname)
        End If
        cbcServiceRep.SetItemData = rst!ustCode
        rst.MoveNext
    Loop
    'If slServiceRepName <> "" Then
    If ilServRepUstCode > 0 Then
        For ilUst = 0 To cbcServiceRep.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcServiceRep.GetName(ilUst)), slServiceRepName, vbTextCompare) = 0 Then
            If cbcServiceRep.GetItemData(ilUst) = ilServRepUstCode Then
                cbcServiceRep.SetListIndex = ilUst
                Exit For
            End If
        Next ilUst
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mPopServiceRep"
End Sub
Private Sub mPopFormat()
    Dim slFormatName As String
    Dim llFormatFmtCode As Long
    Dim ilFmt As Integer
    Dim ilRet As Integer
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    If cbcFormat.ListIndex > 1 Then
        slFormatName = Trim$(cbcFormat.Text)
        llFormatFmtCode = cbcFormat.GetItemData(cbcFormat.ListIndex)
    Else
        slFormatName = ""
        llFormatFmtCode = -2
    End If


    cbcFormat.Clear
    cbcFormat.AddItem ("[New]")
    cbcFormat.SetItemData = -1
    cbcFormat.AddItem ("[None]")
    cbcFormat.SetItemData = 0
    ilRet = gPopFormats()
    For ilRow = 0 To UBound(tgFormatInfo) - 1 Step 1
        cbcFormat.AddItem Trim$(tgFormatInfo(ilRow).sName)
        cbcFormat.SetItemData = tgFormatInfo(ilRow).lCode
    Next ilRow
    'If slFormatName <> "" Then
    If llFormatFmtCode > 0 Then
        For ilFmt = 0 To cbcFormat.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcFormat.GetName(ilFmt)), slFormatName, vbTextCompare) = 0 Then
            If cbcFormat.GetItemData(ilFmt) = llFormatFmtCode Then
                cbcFormat.SetListIndex = ilFmt
                Exit For
            End If
        Next ilFmt
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-mPopFormat"
End Sub

Private Function mSaveFormat() As Integer
'    Dim slName As String
'    Dim llRow As Long
'
'    On Error GoTo ErrHand
'
'    'If cboFormat.ListIndex >= 0 Then
'    '    mSaveFormat = cboFormat.ItemData(cboFormat.ListIndex)
'    If imFormatIndex >= 0 Then
'        mSaveFormat = cboFormat.ItemData(imFormatIndex)
'        Exit Function
'    End If
'    slName = Trim$(cboFormat.Text)
'    If slName = "" Then
'        mSaveFormat = 0
'        Exit Function
'    End If
'    SQLQuery = "INSERT INTO FMT_Station_Format (fmtName, fmtUstCode, fmtGroupName, fmtDftCode, fmtUnused) "
'    SQLQuery = SQLQuery & " VALUES ( " & "'" & slName & "', " & igUstCode & ", " & "''" & ", " & 0 & ",''" & ")"
'    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'        GoSub ErrHand:
'    End If
'    mPopFormat
'    llRow = SendMessageByString(cboFormat.hwnd, CB_FINDSTRING, -1, slName)
'    If llRow >= 0 Then
'        cboFormat.ListIndex = llRow
'        mSaveFormat = cboFormat.ItemData(llRow)
'        Exit Function
'    End If
'    mSaveFormat = -1
'    Exit Function
'
End Function

Private Sub mPopState()
    Dim ilRet As Integer
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    cboState.Clear
    cboStateLic.Clear
    cboONState.Clear
    ilRet = gPopStates()
    For ilRow = 0 To UBound(tgStateInfo) - 1 Step 1
        cboState.AddItem Trim$(tgStateInfo(ilRow).sPostalName) & " (" & Trim$(tgStateInfo(ilRow).sName) & ")"
        cboState.ItemData(cboState.NewIndex) = ilRow    'tgStateInfo(ilRow).iCode
        cboStateLic.AddItem Trim$(tgStateInfo(ilRow).sPostalName) & " (" & Trim$(tgStateInfo(ilRow).sName) & ")"
        cboStateLic.ItemData(cboStateLic.NewIndex) = ilRow    'tgStateInfo(ilRow).iCode
        cboONState.AddItem Trim$(tgStateInfo(ilRow).sPostalName) & " (" & Trim$(tgStateInfo(ilRow).sName) & ")"
        cboONState.ItemData(cboONState.NewIndex) = ilRow    'tgStateInfo(ilRow).iCode
    Next ilRow
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-mPopState"
End Sub

Private Sub mPopTimeZone()
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim slStr As String
    
    On Error GoTo ErrHand
    
    cbcTimeZone.Clear
    ilRet = gPopTimeZones()
    For ilRow = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
        cbcTimeZone.AddItem Trim$(tgTimeZoneInfo(ilRow).sName) & " (" & Left$(Trim$(tgTimeZoneInfo(ilRow).sCSIName), 1) & "TZ)"
        cbcTimeZone.SetItemData = ilRow 'tgTimeZoneInfo(ilRow).iCode
    Next ilRow
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-mPopTimeZone"
End Sub


Private Sub mPopCity()
    Dim slCityName As String
    Dim slOnCityName As String
    Dim slCityLicName As String
    Dim llCityMntCode As Long
    Dim llOnCityMntCode As Long
    Dim llCityLicMntCode As Long
    Dim ilMnt As Integer
    
    On Error GoTo ErrHand
    
    If cbcCity.ListIndex > 1 Then
        slCityName = Trim$(cbcCity.Text)
        llCityMntCode = cbcCity.GetItemData(cbcCity.ListIndex)
    Else
        slCityName = ""
        llCityMntCode = -2
    End If
    If cbcONCity.ListIndex > 1 Then
        slOnCityName = Trim$(cbcONCity.Text)
        llOnCityMntCode = cbcONCity.GetItemData(cbcONCity.ListIndex)
    Else
        slOnCityName = ""
        llOnCityMntCode = -2
    End If

    If cbcCityLic.ListIndex > 1 Then
        slCityLicName = Trim$(cbcCityLic.Text)
        llCityLicMntCode = cbcCityLic.GetItemData(cbcCityLic.ListIndex)
    Else
        slCityLicName = ""
        llCityLicMntCode = -2
    End If

    cbcCity.Clear
    cbcCity.AddItem ("[New]")
    cbcCity.SetItemData = -1
    cbcCity.AddItem ("[None]")
    cbcCity.SetItemData = 0
    SQLQuery = "SELECT * FROM Mnt WHERE mntType = 'C' ORDER BY mntName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        cbcCity.AddItem Trim$(rst!mntName)
        cbcCity.SetItemData = rst!mntCode
        rst.MoveNext
    Loop
    'If slCityName <> "" Then
    If llCityMntCode > 0 Then
        For ilMnt = 0 To cbcCity.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcCity.GetName(ilMnt)), slCityName, vbTextCompare) = 0 Then
            If cbcCity.GetItemData(ilMnt) = llCityMntCode Then
                cbcCity.SetListIndex = ilMnt
                Exit For
            End If
        Next ilMnt
    End If

    
    cbcONCity.Clear
    cbcONCity.AddItem ("[New]")
    cbcONCity.SetItemData = -1
    cbcONCity.AddItem ("[None]")
    cbcONCity.SetItemData = 0
    SQLQuery = "SELECT * FROM Mnt WHERE mntType = 'C' ORDER BY mntName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        cbcONCity.AddItem Trim$(rst!mntName)
        cbcONCity.SetItemData = rst!mntCode
        rst.MoveNext
    Loop
    'If slOnCityName <> "" Then
    If llOnCityMntCode > 0 Then
        For ilMnt = 0 To cbcONCity.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcOnCity.GetName(ilMnt)), slOnCityName, vbTextCompare) = 0 Then
            If cbcONCity.GetItemData(ilMnt) = llOnCityMntCode Then
                cbcONCity.SetListIndex = ilMnt
                Exit For
            End If
        Next ilMnt
    End If
    

    cbcCityLic.Clear
    cbcCityLic.AddItem ("[New]")
    cbcCityLic.SetItemData = -1
    cbcCityLic.AddItem ("[None]")
    cbcCityLic.SetItemData = 0
    SQLQuery = "SELECT * FROM Mnt WHERE mntType = 'C' ORDER BY mntName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        cbcCityLic.AddItem Trim$(rst!mntName)
        cbcCityLic.SetItemData = rst!mntCode
        rst.MoveNext
    Loop
    'If slCityName <> "" Then
    If llCityLicMntCode > 0 Then
        For ilMnt = 0 To cbcCityLic.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcCityLic.GetName(ilMnt)), slCityName, vbTextCompare) = 0 Then
            If cbcCityLic.GetItemData(ilMnt) = llCityLicMntCode Then
                cbcCityLic.SetListIndex = ilMnt
                Exit For
            End If
        Next ilMnt
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mPopCity"
End Sub

Private Sub mPopCounty()
    Dim slCountyLicName As String
    Dim llCountyLicMntCode As Long
    Dim ilMnt As Integer
    
    On Error GoTo ErrHand
    
    If cbcCountyLic.ListIndex > 1 Then
        slCountyLicName = Trim$(cbcCountyLic.Text)
        llCountyLicMntCode = cbcCountyLic.GetItemData(cbcCountyLic.ListIndex)
    Else
        slCountyLicName = ""
        llCountyLicMntCode = -2
    End If


    cbcCountyLic.Clear
    cbcCountyLic.AddItem ("[New]")
    cbcCountyLic.SetItemData = -1
    cbcCountyLic.AddItem ("[None]")
    cbcCountyLic.SetItemData = 0
    SQLQuery = "SELECT * FROM Mnt WHERE mntType = 'Y' ORDER BY mntName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        cbcCountyLic.AddItem Trim$(rst!mntName)
        cbcCountyLic.SetItemData = rst!mntCode
        rst.MoveNext
    Loop
    'If slCountyName <> "" Then
    If llCountyLicMntCode > 0 Then
        For ilMnt = 0 To cbcCountyLic.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcCountyLic.GetName(ilMnt)), slCountyName, vbTextCompare) = 0 Then
            If cbcCountyLic.GetItemData(ilMnt) = llCountyLicMntCode Then
                cbcCountyLic.SetListIndex = ilMnt
                Exit For
            End If
        Next ilMnt
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mPopCounty"
End Sub

Private Sub mPopMoniker()
    Dim slMonikerName As String
    Dim llMonikerMntCode As Long
    Dim ilMnt As Integer
    
    On Error GoTo ErrHand
    If cbcMoniker.ListIndex > 1 Then
        slMonikerName = Trim$(cbcMoniker.Text)
        llMonikerMntCode = cbcMoniker.GetItemData(cbcMoniker.ListIndex)
    Else
        slMonikerName = ""
        llMonikerMntCode = -2
    End If
    cbcMoniker.Clear
    cbcMoniker.AddItem ("[New]")
    cbcMoniker.SetItemData = -1
    cbcMoniker.AddItem ("[None]")
    cbcMoniker.SetItemData = 0
    SQLQuery = "SELECT * FROM Mnt WHERE mntType = 'M' ORDER BY mntName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        cbcMoniker.AddItem Trim$(rst!mntName)
        cbcMoniker.SetItemData = rst!mntCode
        rst.MoveNext
    Loop
    'If slMonikerName <> "" Then
    If llMonikerMntCode > 0 Then
        For ilMnt = 0 To cbcMoniker.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcMoniker.GetName(ilMnt)), slMonikerName, vbTextCompare) = 0 Then
            If cbcMoniker.GetItemData(ilMnt) = llMonikerMntCode Then
                cbcMoniker.SetListIndex = ilMnt
                Exit For
            End If
        Next ilMnt
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mPopMoniker"
End Sub

Private Sub mPopOperator()
    Dim slOperatorName As String
    Dim llOperatorMntCode As Long
    Dim ilMnt As Integer
    
    On Error GoTo ErrHand
    If cbcOperator.ListIndex > 1 Then
        slOperatorName = Trim$(cbcOperator.Text)
        llOperatorMntCode = cbcOperator.GetItemData(cbcOperator.ListIndex)
    Else
        slOperatorName = ""
        llOperatorMntCode = -2
    End If
    cbcOperator.Clear
    cbcOperator.AddItem ("[New]")
    cbcOperator.SetItemData = -1
    cbcOperator.AddItem ("[Same as Owner]")
    cbcOperator.SetItemData = 0
    SQLQuery = "SELECT * FROM Mnt WHERE mntType = 'O' ORDER BY mntName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        cbcOperator.AddItem Trim$(rst!mntName)
        cbcOperator.SetItemData = rst!mntCode
        rst.MoveNext
    Loop
    'If slOperatorName <> "" Then
    If llOperatorMntCode > 0 Then
        For ilMnt = 0 To cbcOperator.ListCount - 1 Step 1
            'If StrComp(Trim$(cbcOperator.GetName(ilMnt)), slOperatorName, vbTextCompare) = 0 Then
            If cbcOperator.GetItemData(ilMnt) = llOperatorMntCode Then
                cbcOperator.SetListIndex = ilMnt
                Exit For
            End If
        Next ilMnt
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mPopOperator"
End Sub

Private Sub mMCPaintRowColor(llRow As Long)
    Dim llCol As Long

    grdMulticast.Row = llRow
    For llCol = MCCALLLETTERSINDEX To MCOWNERINDEX Step 1
        grdMulticast.Col = llCol
        If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) <> "1" Then
            grdMulticast.CellBackColor = vbWhite
        Else
            grdMulticast.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub

Private Sub mSSPaintRowColor(llRow As Long)
    Dim llCol As Long

    grdSisterStations.Row = llRow
    For llCol = SSCALLLETTERSINDEX To SSMAILSTATEINDEX Step 1
        grdSisterStations.Col = llCol
        If grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) <> "1" Then
            grdSisterStations.CellBackColor = vbWhite
        Else
            grdSisterStations.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub

Private Sub mMCSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
        slStr = Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX))
        If slStr <> "" Then
            slSort = UCase$(Trim$(grdMulticast.TextMatrix(llRow, ilCol)))
            If slSort = "" Then
                slSort = "!"
            End If
            slStr = grdMulticast.TextMatrix(llRow, MCSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastMCColSorted) Or ((ilCol = imLastMCColSorted) And (imLastMCSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdMulticast.TextMatrix(llRow, MCSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdMulticast.TextMatrix(llRow, MCSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastMCColSorted Then
        imLastMCColSorted = MCSORTINDEX
    Else
        imLastMCColSorted = -1
        imLastMCSort = -1
    End If
    gGrid_SortByCol grdMulticast, MCCALLLETTERSINDEX, MCSORTINDEX, imLastMCColSorted, imLastMCSort
    imLastMCColSorted = ilCol
End Sub

Private Sub mSSSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
        slStr = Trim$(grdMulticast.TextMatrix(llRow, SSCALLLETTERSINDEX))
        If slStr <> "" Then
            slSort = UCase$(Trim$(grdMulticast.TextMatrix(llRow, ilCol)))
            If slSort = "" Then
                slSort = "!"
            End If
            slStr = grdMulticast.TextMatrix(llRow, SSSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastSSColSorted) Or ((ilCol = imLastSSColSorted) And (imLastSSSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdMulticast.TextMatrix(llRow, SSSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdMulticast.TextMatrix(llRow, SSSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastSSColSorted Then
        imLastSSColSorted = SSSORTINDEX
    Else
        imLastSSColSorted = -1
        imLastSSSort = -1
    End If
    gGrid_SortByCol grdMulticast, SSCALLLETTERSINDEX, SSSORTINDEX, imLastSSColSorted, imLastSSSort
    imLastSSColSorted = ilCol
End Sub
Private Function mSaveMulticast(llMulticastGroupID As Long) As Integer
    Dim llRow As Long
    Dim ilSelCount As Integer
    Dim ilNotSelCount As Integer
    Dim llNewGrpID As Long
    Dim ilMktCode As Integer
    Dim ilRet As Integer
    On Error GoTo ErrHand

    mSaveMulticast = False
    llMulticastGroupID = 0
    If rbcMulticast(0).Value Then   'Create Multicast group
        ilSelCount = 0
        For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
            If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1" Then
                ilSelCount = ilSelCount + 1
            End If
        Next llRow
        If ilSelCount <= 0 Then
            mSaveMulticast = True
            Exit Function
        Else
            llNewGrpID = gMulticastMaxGroupID() + 1
            'imDMAMktCode = 0
            'If cbcDMAMarket.ListIndex > 1 Then
            '    imDMAMktCode = cbcDMAMarket.GetItemData(cbcDMAMarket.ListIndex)
            'End If
            'SQLQuery = "Insert Into mgt ( "
            'SQLQuery = SQLQuery & "mgtGroupID, "
            'SQLQuery = SQLQuery & "mgtShfCode, "
            'SQLQuery = SQLQuery & "mgtMktCode, "
            'SQLQuery = SQLQuery & "mgtEnteredDate, "
            'SQLQuery = SQLQuery & "mgtRemovedDate, "
            'SQLQuery = SQLQuery & "mgtUsfCode, "
            'SQLQuery = SQLQuery & "mgtUnused "
            'SQLQuery = SQLQuery & ") "
            'SQLQuery = SQLQuery & "Values ( "
            'SQLQuery = SQLQuery & llNewGrpID & ", "
            'SQLQuery = SQLQuery & imShttCode & ", "
            'SQLQuery = SQLQuery & imDMAMktCode & ", "
            'SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
            'SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
            'SQLQuery = SQLQuery & igUstCode & ", "
            'SQLQuery = SQLQuery & "'" & "" & "' "
            'SQLQuery = SQLQuery & ") "
            'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '    GoSub ErrHand:
            'End If
            SQLQuery = "UPDATE shtt SET shttMultiCastGroupID = " & llNewGrpID & " WHERE shttCode = " & imShttCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "Station-mSaveMulticast"
                mSaveMulticast = False
                Exit Function
            End If
            For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
                If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1" Then
                    'SQLQuery = "Insert Into mgt ( "
                    'SQLQuery = SQLQuery & "mgtGroupID, "
                    'SQLQuery = SQLQuery & "mgtShfCode, "
                    'SQLQuery = SQLQuery & "mgtMktCode, "
                    'SQLQuery = SQLQuery & "mgtEnteredDate, "
                    'SQLQuery = SQLQuery & "mgtRemovedDate, "
                    'SQLQuery = SQLQuery & "mgtUsfCode, "
                    'SQLQuery = SQLQuery & "mgtUnused "
                    'SQLQuery = SQLQuery & ") "
                    'SQLQuery = SQLQuery & "Values ( "
                    'SQLQuery = SQLQuery & llNewGrpID & ", "
                    'SQLQuery = SQLQuery & grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX) & ", "
                    'SQLQuery = SQLQuery & grdMulticast.TextMatrix(llRow, MCDMAMKTCODEINDEX) & ", "
                    'SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & igUstCode & ", "
                    'SQLQuery = SQLQuery & "'" & "" & "' "
                    'SQLQuery = SQLQuery & ") "
                    'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '    GoSub ErrHand:
                    'End If
                    SQLQuery = "UPDATE shtt SET shttMultiCastGroupID = " & llNewGrpID & " WHERE shttCode = " & grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX)
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "Station-mSaveMulticast"
                        mSaveMulticast = False
                        Exit Function
                    End If
                    ilRet = gBinarySearchStationInfoByCode(grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX))
                    If ilRet <> -1 Then
                        tgStationInfoByCode(ilRet).lMultiCastGroupID = llNewGrpID
                        ilRet = gBinarySearchStation(Trim$(tgStationInfoByCode(ilRet).sCallLetters))
                        If ilRet <> -1 Then
                            tgStationInfo(ilRet).lMultiCastGroupID = llNewGrpID
                        End If
                    End If
                End If
            Next llRow
            llMulticastGroupID = llNewGrpID
        End If
    ElseIf (rbcMulticast(1).Value) And (lmMultiCastGroupID <= 0) Then   'Add to Group
        For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
            If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1" Then
                llNewGrpID = grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX)
                'imDMAMktCode = 0
                'If cbcDMAMarket.ListIndex > 1 Then
                '    imDMAMktCode = cbcDMAMarket.GetItemData(cbcDMAMarket.ListIndex)
                'End If
                'SQLQuery = "Insert Into mgt ( "
                'SQLQuery = SQLQuery & "mgtGroupID, "
                'SQLQuery = SQLQuery & "mgtShfCode, "
                'SQLQuery = SQLQuery & "mgtMktCode, "
                'SQLQuery = SQLQuery & "mgtEnteredDate, "
                'SQLQuery = SQLQuery & "mgtRemovedDate, "
                'SQLQuery = SQLQuery & "mgtUsfCode, "
                'SQLQuery = SQLQuery & "mgtUnused "
                'SQLQuery = SQLQuery & ") "
                'SQLQuery = SQLQuery & "Values ( "
                'SQLQuery = SQLQuery & llNewGrpID & ", "
                'SQLQuery = SQLQuery & imShttCode & ", "
                'SQLQuery = SQLQuery & imDMAMktCode & ", "
                'SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
                'SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
                'SQLQuery = SQLQuery & igUstCode & ", "
                'SQLQuery = SQLQuery & "'" & "" & "' "
                'SQLQuery = SQLQuery & ") "
                'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '    GoSub ErrHand:
                'End If
                SQLQuery = "UPDATE shtt SET shttMultiCastGroupID = " & llNewGrpID & " WHERE shttCode = " & imShttCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Station-mSaveMulticast"
                    mSaveMulticast = False
                    Exit Function
                End If
                llMulticastGroupID = llNewGrpID
                Exit For
            End If
        Next llRow
    ElseIf (rbcMulticast(1).Value) And (lmMultiCastGroupID > 0) Then   'Add to Group
        For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
            If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1" Then
                SQLQuery = "UPDATE shtt SET shttMultiCastGroupID = " & lmMultiCastGroupID & " WHERE shttCode = " & grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX)
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Station-mSaveMulticast"
                    mSaveMulticast = False
                    Exit Function
                End If
                ilRet = gBinarySearchStationInfoByCode(grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX))
                If ilRet <> -1 Then
                    tgStationInfoByCode(ilRet).lMultiCastGroupID = lmMultiCastGroupID
                    ilRet = gBinarySearchStation(Trim$(tgStationInfoByCode(ilRet).sCallLetters))
                    If ilRet <> -1 Then
                        tgStationInfo(ilRet).lMultiCastGroupID = lmMultiCastGroupID
                    End If
                End If
            End If
        Next llRow
    ElseIf rbcMulticast(2).Value Then   'Remove from group
        ilSelCount = 0
        For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
            If Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX)) <> "" Then
                If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1" Then
                    ilSelCount = ilSelCount + 1
                End If
            End If
        Next llRow
        If ilSelCount = 0 Then
            llMulticastGroupID = lmMultiCastGroupID
            mSaveMulticast = True
            Exit Function
        End If
        ilNotSelCount = 0
        For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
            If Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX)) <> "" Then
                If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "0" Then
                    ilNotSelCount = ilNotSelCount + 1
                End If
            End If
        Next llRow
        If ilNotSelCount = 1 Then
            For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
                If Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX)) <> "" Then
                    If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "0" Then
                        grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1"
                    End If
                End If
            Next llRow
        End If
        For llRow = grdMulticast.FixedRows To grdMulticast.Rows - 1 Step 1
            If Trim$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX)) <> "" Then
                If grdMulticast.TextMatrix(llRow, MCSELECTEDINDEX) = "1" Then
                    'SQLQuery = "DELETE FROM mgt WHERE (mgtShfCode = " & grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX) & ")"
                    'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '    GoSub ErrHand:
                    'End If
                    SQLQuery = "UPDATE shtt SET shttMultiCastGroupID = " & 0 & " WHERE shttCode = " & grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX)
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "Station-mSaveMulticast"
                        mSaveMulticast = False
                        Exit Function
                    '8824 handle removing multicast
                    Else
                        gVatSetToGoToWebByShttCode grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX), 0
                    End If
                    ilRet = gBinarySearchStationInfoByCode(grdMulticast.TextMatrix(llRow, MCSHTTCODEINDEX))
                    If ilRet <> -1 Then
                        tgStationInfoByCode(ilRet).lMultiCastGroupID = 0
                        ilRet = gBinarySearchStation(Trim$(tgStationInfoByCode(ilRet).sCallLetters))
                        If ilRet <> -1 Then
                            tgStationInfo(ilRet).lMultiCastGroupID = 0
                        End If
                    End If
                Else
                    If Trim$(UCase$(grdMulticast.TextMatrix(llRow, MCCALLLETTERSINDEX))) = Trim$(UCase$(txtCallLetters.Text)) Then
                        llMulticastGroupID = lmMultiCastGroupID
                    End If
                End If
            End If
        Next llRow
    End If
    mSaveMulticast = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mSaveMulticast"
    mSaveMulticast = False
End Function

Private Function mSaveSisterStations(llMarketClusterGroupID As Long) As Integer
    Dim llRow As Long
    Dim ilSelCount As Integer
    Dim ilNotSelCount As Integer
    Dim llNewGrpID As Long
    Dim ilMktCode As Integer
    Dim ilRet As Integer
    Dim temp_rst As ADODB.Recordset
    On Error GoTo ErrHand

    mSaveSisterStations = False
    llMarketClusterGroupID = 0
    If rbcMarketCluster(0).Value Then   'Create MarketCluster group
        ilSelCount = 0
        For llRow = grdSisterStations.FixedRows To grdSisterStations.Rows - 1 Step 1
            If grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "1" Then
                ilSelCount = ilSelCount + 1
            End If
        Next llRow
        If ilSelCount <= 0 Then
            mSaveSisterStations = True
            Exit Function
        Else
            llNewGrpID = gMarketClusterMaxGroupID() + 1
            'imDMAMktCode = 0
            'SQLQuery = "Insert Into sgt ( "
            'SQLQuery = SQLQuery & "sgtGroupID, "
            'SQLQuery = SQLQuery & "sgtShttCode, "
            'SQLQuery = SQLQuery & "sgtMasterStation, "
            'SQLQuery = SQLQuery & "sgtEnteredDate, "
            'SQLQuery = SQLQuery & "sgtRemovedDate, "
            'SQLQuery = SQLQuery & "sgtUstCode, "
            'SQLQuery = SQLQuery & "sgtUnused "
            'SQLQuery = SQLQuery & ") "
            'SQLQuery = SQLQuery & "Values ( "
            'SQLQuery = SQLQuery & llNewGrpID & ", "
            'SQLQuery = SQLQuery & imShttCode & ", "
            'If UCase$(Trim$(edcMasterStation.Text)) = UCase$(Trim$(txtCallLetters.Text)) Then
            '    SQLQuery = SQLQuery & "'" & "Y" & "', "
            'Else
            '    SQLQuery = SQLQuery & "'" & "N" & "', "
            'End If
            'SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
            'SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
            'SQLQuery = SQLQuery & igUstCode & ", "
            'SQLQuery = SQLQuery & "'" & "" & "' "
            'SQLQuery = SQLQuery & ") "
            'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '    GoSub ErrHand:
            'End If
            SQLQuery = "UPDATE shtt SET shttClusterGroupID = " & llNewGrpID & " WHERE shttCode = " & imShttCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "Station-mSaveSisterStations"
                mSaveSisterStations = False
                Exit Function
            End If
            For llRow = grdSisterStations.FixedRows To grdSisterStations.Rows - 1 Step 1
                If grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "1" Then
                    'SQLQuery = "Insert Into sgt ( "
                    'SQLQuery = SQLQuery & "sgtGroupID, "
                    'SQLQuery = SQLQuery & "sgtShttCode, "
                    'SQLQuery = SQLQuery & "sgtMasterStation, "
                    'SQLQuery = SQLQuery & "sgtEnteredDate, "
                    'SQLQuery = SQLQuery & "sgtRemovedDate, "
                    'SQLQuery = SQLQuery & "sgtUstCode, "
                    'SQLQuery = SQLQuery & "sgtUnused "
                    'SQLQuery = SQLQuery & ") "
                    'SQLQuery = SQLQuery & "Values ( "
                    'SQLQuery = SQLQuery & llNewGrpID & ", "
                    'SQLQuery = SQLQuery & grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX) & ", "
                    'If UCase$(Trim$(edcMasterStation.Text)) = UCase$(Trim$(grdSisterStations.TextMatrix(llRow, SSCALLLETTERSINDEX))) Then
                    '    SQLQuery = SQLQuery & "'" & "Y" & "', "
                    'Else
                    '    SQLQuery = SQLQuery & "'" & "N" & "', "
                    'End If
                    'SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & igUstCode & ", "
                    'SQLQuery = SQLQuery & "'" & "" & "' "
                    'SQLQuery = SQLQuery & ") "
                    'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '    GoSub ErrHand:
                    'End If
                    SQLQuery = "UPDATE shtt SET shttClusterGroupID = " & llNewGrpID & " WHERE shttCode = " & grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX)
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "Station-mSaveSisterStations"
                        mSaveSisterStations = False
                        Exit Function
                    End If
                    ilRet = gBinarySearchStationInfoByCode(grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX))
                    If ilRet <> -1 Then
                        tgStationInfoByCode(ilRet).lMarketClusterGroupID = llNewGrpID
                        ilRet = gBinarySearchStation(Trim$(tgStationInfoByCode(ilRet).sCallLetters))
                        If ilRet <> -1 Then
                            tgStationInfo(ilRet).lMarketClusterGroupID = llNewGrpID
                        End If
                    End If
                End If
            Next llRow
            llMarketClusterGroupID = llNewGrpID
        End If
    ElseIf (rbcMarketCluster(1).Value) And (lmClusterGroupID <= 0) Then   'Add to Group
        For llRow = grdSisterStations.FixedRows To grdSisterStations.Rows - 1 Step 1
            If grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "1" Then
                '8824
                smOldMaster = "DAN8824"
                llNewGrpID = grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX)
                'SQLQuery = "Insert Into sgt ( "
                'SQLQuery = SQLQuery & "sgtGroupID, "
                'SQLQuery = SQLQuery & "sgtShttCode, "
                'SQLQuery = SQLQuery & "sgtMasterStation, "
                'SQLQuery = SQLQuery & "sgtEnteredDate, "
                'SQLQuery = SQLQuery & "sgtRemovedDate, "
                'SQLQuery = SQLQuery & "sgtUstCode, "
                'SQLQuery = SQLQuery & "sgtUnused "
                'SQLQuery = SQLQuery & ") "
                'SQLQuery = SQLQuery & "Values ( "
                'SQLQuery = SQLQuery & llNewGrpID & ", "
                'SQLQuery = SQLQuery & imShttCode & ", "
                'If UCase$(Trim$(edcMasterStation.Text)) = UCase$(Trim$(txtCallLetters.Text)) Then
                '    SQLQuery = SQLQuery & "'" & "Y" & "', "
                'Else
                '    SQLQuery = SQLQuery & "'" & "N" & "', "
                'End If
                'SQLQuery = SQLQuery & imDMAMktCode & ", "
                'SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
                'SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
                'SQLQuery = SQLQuery & igUstCode & ", "
                'SQLQuery = SQLQuery & "'" & "" & "' "
                'SQLQuery = SQLQuery & ") "
                'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '    GoSub ErrHand:
                'End If
                SQLQuery = "UPDATE shtt SET shttClusterGroupID = " & llNewGrpID & " WHERE shttCode = " & imShttCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Station-mSaveSisterStations"
                    mSaveSisterStations = False
                    Exit Function
                End If
                llMarketClusterGroupID = llNewGrpID
                Exit For
            End If
        Next llRow
    ElseIf (rbcMarketCluster(1).Value) And (lmClusterGroupID > 0) Then   'Add to Group
        For llRow = grdSisterStations.FixedRows To grdSisterStations.Rows - 1 Step 1
            If grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "1" Then
                '8824
                smOldMaster = "DAN8824"
                llNewGrpID = lmClusterGroupID
                SQLQuery = "UPDATE shtt SET shttClusterGroupID = " & lmClusterGroupID & " WHERE shttCode = " & grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX)
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Station-mSaveSisterStations"
                    mSaveSisterStations = False
                    Exit Function
                End If
                ilRet = gBinarySearchStationInfoByCode(grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX))
                If ilRet <> -1 Then
                    tgStationInfoByCode(ilRet).lMarketClusterGroupID = lmClusterGroupID
                    ilRet = gBinarySearchStation(Trim$(tgStationInfoByCode(ilRet).sCallLetters))
                    If ilRet <> -1 Then
                        tgStationInfo(ilRet).lMarketClusterGroupID = lmClusterGroupID
                    End If
                End If
            End If
        Next llRow
    ElseIf rbcMarketCluster(2).Value Then   'Remove from group
        ilSelCount = 0
        For llRow = grdSisterStations.FixedRows To grdSisterStations.Rows - 1 Step 1
            If Trim$(grdSisterStations.TextMatrix(llRow, SSCALLLETTERSINDEX)) <> "" Then
                If grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "1" Then
                    ilSelCount = ilSelCount + 1
                End If
            End If
        Next llRow
        If ilSelCount = 0 Then
            llMarketClusterGroupID = lmClusterGroupID
            llNewGrpID = lmClusterGroupID
        Else
            '8824
            smOldMaster = "DAN8824"
            ilNotSelCount = 0
            For llRow = grdSisterStations.FixedRows To grdSisterStations.Rows - 1 Step 1
                If Trim$(grdSisterStations.TextMatrix(llRow, SSCALLLETTERSINDEX)) <> "" Then
                    If grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "0" Then
                        ilNotSelCount = ilNotSelCount + 1
                    End If
                End If
            Next llRow
            If ilNotSelCount = 1 Then
                For llRow = grdSisterStations.FixedRows To grdSisterStations.Rows - 1 Step 1
                    If Trim$(grdSisterStations.TextMatrix(llRow, SSCALLLETTERSINDEX)) <> "" Then
                        If grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "0" Then
                            grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "1"
                        End If
                    End If
                Next llRow
            End If
            'llNewGrpID = 0
            'SQLQuery = "SELECT shttClusterGroupID FROM shtt WHERE ShttCode = " & imShttCode
            'Set rst = gSQLSelectCall(SQLQuery)
            'If Not rst.EOF Then
            '    llNewGrpID = rst!shttClusterGroupID
            'End If
            llNewGrpID = lmClusterGroupID
            For llRow = grdSisterStations.FixedRows To grdSisterStations.Rows - 1 Step 1
                If Trim$(grdSisterStations.TextMatrix(llRow, SSCALLLETTERSINDEX)) <> "" Then
                    If grdSisterStations.TextMatrix(llRow, SSSELECTEDINDEX) = "1" Then
                        'SQLQuery = "DELETE FROM sgt WHERE (sgtShttCode = " & grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX) & ")"
                        'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '    GoSub ErrHand:
                        'End If
                        SQLQuery = "UPDATE shtt SET shttClusterGroupID = " & 0 & ", "
                        SQLQuery = SQLQuery & " shttMasterCluster = " & "'N'"
                        SQLQuery = SQLQuery & " WHERE shttCode = " & grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX)
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/12/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "Station-mSaveSisterStations"
                            mSaveSisterStations = False
                            Exit Function
                        End If
                        ilRet = gBinarySearchStationInfoByCode(grdSisterStations.TextMatrix(llRow, SSSHTTCODEINDEX))
                        If ilRet <> -1 Then
                            tgStationInfoByCode(ilRet).lMarketClusterGroupID = 0
                            ilRet = gBinarySearchStation(Trim$(tgStationInfoByCode(ilRet).sCallLetters))
                            If ilRet <> -1 Then
                                tgStationInfo(ilRet).lMarketClusterGroupID = 0
                            End If
                        End If
                    Else
                        If Trim$(UCase$(grdSisterStations.TextMatrix(llRow, SSCALLLETTERSINDEX))) = Trim$(UCase$(txtCallLetters.Text)) Then
                            llMarketClusterGroupID = lmClusterGroupID
                        End If
                    End If
                End If
            Next llRow
        End If
    End If
    If rbcMarketCluster(0).Value Or rbcMarketCluster(1).Value Or ((rbcMarketCluster(2).Value) And (llNewGrpID <> 0)) Then
        SQLQuery = "SELECT shttCallLetters, shttCode FROM shtt WHERE shttClusterGroupID = " & llNewGrpID
        Set rst = gSQLSelectCall(SQLQuery)
        Do While Not rst.EOF
            If UCase$(Trim$(edcMasterStation.Text)) = UCase$(Trim$(rst!shttCallLetters)) Then
                SQLQuery = "UPDATE shtt SET shttMasterCluster = '" & "Y" & "' WHERE shttCode = " & rst!shttCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Station-mSaveSisterStations"
                    mSaveSisterStations = False
                    Exit Function
                End If
            Else
                SQLQuery = "UPDATE shtt SET shttMasterCluster = '" & "N" & "' WHERE shttCode = " & rst!shttCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Station-mSaveSisterStations"
                    mSaveSisterStations = False
                    Exit Function
                End If
            End If
            rst.MoveNext
        Loop
    End If
    mSaveSisterStations = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mSaveSisterStations"
    mSaveSisterStations = False
End Function

Private Sub udcContactGrid_CallLettersChanged(slCallLetters As String)
    txtCallLetters.Text = slCallLetters
    imFieldChgd = True
End Sub

Private Sub udcContactGrid_FaxChanged(slFax As String)
    txtFax.Text = slFax
    imFieldChgd = True
End Sub

Private Sub udcContactGrid_PhoneChanged(slPhone As String)
    txtStaPhone.Text = slPhone
    imFieldChgd = True
End Sub

Private Sub mSetTabColors()
    'Program: Paint and from menu select Image->Attribute, then width 86 by height 13
    '    Set custom color to Grey R=236; G=233; B=216
    '    Create Title in Black (back to Color and click on back) and center
    '    Zoom
    '    Back to color and click on custom grey color
    '    Fill in
    '    Save bmp
    '    Repeat for Green lettering
    'Add image control
    '    Add bmp to image control (black, then green)
    '
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilRepaintTabs As Integer
    Dim llColor As Long
    'Dim ilcListImage As ListImage
    
    On Error GoTo mSetTabColorErr
    ilRepaintTabs = False
    For ilLoop = 0 To 5 Step 1
        llColor = vbBlack   'vbBlue
        Select Case ilLoop
            Case 0  'Main
            Case 1  'History
            Case 2  'Personnel
            Case 3  'Sister Stations
                If lmClusterGroupID > 0 Then
                    llColor = vbGreen
                End If
            Case 4  'Multicast
                If lmMultiCastGroupID > 0 Then
                    llColor = vbGreen
                End If
            Case 5  'Interface
        End Select
        If lmTabColor(ilLoop) <> llColor Then
            ilRepaintTabs = True
            lmTabColor(ilLoop) = llColor
        End If
    Next ilLoop
    If Not ilRepaintTabs Then
        Exit Sub
    End If

    Set tscStation.ImageList = imcTabColor
    For ilLoop = 0 To 5 Step 1
        If lmTabColor(ilLoop) = vbBlack Then
            tscStation.Tabs(ilLoop + 1).Image = 2 * (ilLoop + 1) - 1
        Else
            tscStation.Tabs(ilLoop + 1).Image = 2 * (ilLoop + 1)
        End If
    Next ilLoop
    Exit Sub
mSetTabColorErr:
    Resume Next
        
End Sub

Private Sub mPop()
    Dim slShttTimeStamp As String
    
    'Reload stations if file date/time changed
    If bmDoPop Then
        sgShttTimeStamp = ""
        gPopStations
        gPopVff
        
        Call mOwnerFillListBox(0, False, -1)
        Call mDMAMarketFillListBox(0, False, -1)
        Call mMSAMarketFillListBox(0, False, -1)
        
        mPopTerritory
        mPopFormat
        mPopCity
        mPopCounty
        mPopMoniker
        mPopOperator
        mPopArea
        mPopMarketRep
        mPopServiceRep
    Else
        sgShttTimeStamp = gFileDateTime(sgDBPath & "Shtt.mkd")
    End If
    bmDoPop = False
End Sub

Private Function mWebVerChgMultiCastSync() As Boolean

    Dim blWebVersionChanged As Boolean
    Dim ilGroupID As Integer
    Dim ilShttCode As Integer
    Dim slCallLetters As String
    Dim ilRet As String
    Dim shtt_rst As ADODB.Recordset
    Dim slHdr As String
    Dim slFtr As String
    Dim slStr As String
    Dim ilOrigWebVer As Integer
    Dim ilNewWebVer As Integer
    Dim slNewWebVer As String
    Dim blDiff As Boolean
    Dim slCurrSta As String
    Dim ilLoop As Integer
    Dim ilRowsEffected As Integer
    Dim slSQLQuery As String
    Dim slWebStaToUpdate() As String
    Dim ilIdx As Integer
    
    On Error GoTo ErrHand
    mWebVerChgMultiCastSync = False
    'Is it a multicast station?
    ilRet = gIsMulticast(imShttCode)
    If ilRet Then


'        If rbcWebSiteVersion(1).Value = True Then
        If Not gIsUsingNovelty Then
            If smNewWebNumber = "2" Then
                ilNewWebVer = 2
                slNewWebVer = "2"
            Else
                ilNewWebVer = 1
                slNewWebVer = "1"
            End If
    
            If smOrigWebNumber <> smNewWebNumber Then
                blWebVersionChanged = True
            End If
        End If


        'Determine if the web site version number has been changed. All stations multicast in the same group must use the same web version
'        If (rbcWebSiteVersion(1).Value = True And smWebNumber = "1") Or (rbcWebSiteVersion(0).Value = True And smWebNumber = "2") Then
'            blWebVersionChanged = True
'        End If
        
        If Not gIsUsingNovelty And blWebVersionChanged Then
            slCurrSta = gGetCallLettersByShttCode(imShttCode)
            ilGroupID = gGetStaMulticastGroupID(imShttCode)
            SQLQuery = "Select shttCode, shttCallLetters, shttWebNumber from shtt where shttMultiCastGroupID = " & ilGroupID
            Set shtt_rst = gSQLSelectCall(SQLQuery)
            slHdr = "You have chosen to change " & slCurrSta & " web site version to site " & slNewWebVer & "." & "  It is a multicast station.  "
            slHdr = slHdr & "All stations being multicast in the same group must use the same web site version.  "
            slHdr = slHdr & "Currently, there is a conflict with one or more of the stations in the multicast group." & sgCRLF & sgCRLF
            slFtr = "Would you like to change all stations in this group to the " & slNewWebVer & "?"
            blDiff = False
            ReDim slWebStaToUpdate(0 To 0)
            While Not shtt_rst.EOF
                slCallLetters = shtt_rst!shttCallLetters
                slWebStaToUpdate(UBound(slWebStaToUpdate)) = Trim(slCallLetters)
                ReDim Preserve slWebStaToUpdate(0 To UBound(slWebStaToUpdate) + 1)
                ilShttCode = shtt_rst!shttCode
                If ilShttCode <> imShttCode Then
                    If shtt_rst!shttWebNumber <> "1" And shtt_rst!shttWebNumber <> "2" Then
                        ilOrigWebVer = 1   'if undefined then default to the old original web site = 1
                    Else
                        ilOrigWebVer = Trim(shtt_rst!shttWebNumber)
                    End If
                    If ilNewWebVer <> ilOrigWebVer Then
                        blDiff = True
                        If ilOrigWebVer = 1 Then
                            slStr = slStr & Trim(slCallLetters) & " - " & "1" & sgCRLF
                        Else
                            slStr = slStr & Trim(slCallLetters) & " - " & "2" & sgCRLF
                        End If
                    End If
                End If
                shtt_rst.MoveNext
            Wend
            If blDiff Then
                ilRet = MsgBox(slHdr & slStr & sgCRLF & slFtr, vbYesNo)
            End If
            If ilRet = vbYes Then
                'Update the multicast stations with the new web version
                SQLQuery = "Update shtt SET shttWebNumber = " & ilNewWebVer & " WHERE shttMultiCastGroupID = " & ilGroupID
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "frmStation-mWebVerChgMultiCastSync"
                    mWebVerChgMultiCastSync = False
                    Exit Function
                End If
                'D.S. 03/28/18 TTP:8856 added for loop to send update by call letters rather than groupID
                For ilIdx = 0 To UBound(slWebStaToUpdate) - 1 Step 1
                SQLQuery = "Update Header Set WebSiteVersion = " & ilNewWebVer
                    SQLQuery = SQLQuery & " Where StationName = " & "'" & slWebStaToUpdate(ilIdx) & "'"
                For ilLoop = 0 To 5 Step 1
                    ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
                    If ilRowsEffected <> -1 Then
                        Exit For
                    End If
                    Sleep (1000)
                Next ilLoop
                Next ilIdx
                shtt_rst.Close
            Else
                edcWebNumber.Text = smOrigWebNumber
                edcWebNumber.Refresh
'                If ilOrigWebVer = 1 Then
'                    rbcWebSiteVersion(0).Value = True
'                Else
'                    rbcWebSiteVersion(1).Value = True
'                End If
        End If
    End If
    End If
    Erase slWebStaToUpdate
    mWebVerChgMultiCastSync = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStation-mWebVerChgMultiCastSync"
    mWebVerChgMultiCastSync = False
End Function

Public Function gSaveCurrShttState(iShttCode As Integer)

    Dim ilLoop As Integer
    Dim ilFound As Boolean
    
    ilFound = False
    gPopStations
    For ilLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(ilLoop).iCode = iShttCode Then
            tgShttSavedInfo(0) = tgStationInfo(ilLoop)
            ilFound = True
            Exit For
        End If
    Next ilLoop
End Function
'8418
Private Function mAgreementWithVersion(ilShttCode As Integer, ilMinVersion As Integer) As Boolean
    Dim blRet As Boolean
    Dim rstShtt As ADODB.Recordset
    Dim tlVendors() As VendorInfo
On Error GoTo ErrHand
    blRet = False
    tlVendors = gGetAvailableVendors()
    SQLQuery = "select distinct vatWvtVendorId as Vendor from Vat_Vendor_Agreement inner join att on vatAttCode = AttCode where attShfCode = " & ilShttCode
    Set rstShtt = cnn.Execute(SQLQuery)
    Do While Not rstShtt.EOF
        If gVendorMinVersion(rstShtt!Vendor, tlVendors()) >= ilMinVersion Then
            blRet = True
            Exit Do
        End If
        rstShtt.MoveNext
    Loop
    mAgreementWithVersion = blRet
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-mAgreementWithVersion"
    mAgreementWithVersion = True
End Function
''8824
'Private Function mGetShttCode(slCall As String) As Integer
'    Dim ilRet As Integer
'    Dim rstShtt As ADODB.Recordset
'On Error GoTo ErrHand
'    ilRet = 0
'    SQLQuery = "select shttcode from shtt where shttcallletters = '" & slCall & "'"
'    Set rstShtt = cnn.Execute(SQLQuery)
'    If Not rstShtt.EOF Then
'        ilRet = rstShtt!shttCode
'    End If
'    mGetShttCode = ilRet
'    Exit Function
'ErrHand:
'    gHandleError "AffErrorLog.txt", "frmStation-mGetShttCode"
'    mGetShttCode = True
'End Function

